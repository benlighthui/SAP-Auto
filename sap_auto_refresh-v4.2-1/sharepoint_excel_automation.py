#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SharePoint Excel 自动化脚本
=============================================================
功能模式：
  【测试模式】写入 A1 时间戳 — 验证流程可行性
  【正式模式】SAP Analysis for Office 数据刷新（含刷新完成检测）

操作流程：
  1. 遍历 SharePoint 主文件夹下所有子文件夹（按名称字母顺序）
  2. 每个子文件夹筛选文件名含 OR Qty / SR Qty / SR Margin 的最新 Excel 文件
  3. 通过 ms-excel:ofe|u| 协议以 SharePoint 在线模式打开（Open in App）
  4. 根据当前模式执行操作（写入时间戳 / SAP 刷新）
  5. 保存回 SharePoint
  6. 遵循先开后关规则：始终保持至少一个 Excel 处于打开状态

SAP 刷新完成检测机制：
  - Application.Ready == True（Excel 可交互）
  - Application.CalculationState == xlDone（计算完成）
  - 状态栏文字检测（pywinauto 读取底部状态栏）
  - UsedRange.Rows.Count 前后对比（记录行数变化，仅日志用途）

命令行菜单：
  [1] 设置 SharePoint 文件夹地址
  [2] 切换运行模式（测试 / 正式）
  [3] 预览处理计划并执行
  [4] 退出

运行前准备：
  1. 执行 launch_edge_debug.bat 以调试模式启动 Edge
  2. 在打开的 Edge 中登录 SharePoint 账号
  3. 运行本脚本：python sharepoint_excel_automation.py
"""

import os
import re
import sys
import time
import json
import pythoncom
import win32com.client
from datetime import datetime
from playwright.sync_api import sync_playwright

# ============================================================
# 配置参数
# ============================================================
EDGE_DEBUG_PORT      = 9222
EDGE_USER_DATA_DIR   = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "edge_sp_automation")
EXCEL_LOAD_TIMEOUT   = 90
TARGET_SUFFIXES      = ["OR Qty", "SR Qty", "SR Margin"]
REFRESH_WAIT_SECONDS = 60          # 固定等待（仅作为兜底超时参考）
REFRESH_TIMEOUT      = 60          # SAP 刷新最大等待时间（秒）
REFRESH_POLL_INTERVAL = 3          # 轮询间隔（秒）
TIME_FORMAT          = "%Y-%m-%d %H:%M:%S"

# 运行时变量（由 config.json 加载或菜单设置）
SITE_URL                    = ""
MAIN_FOLDER_SERVER_RELATIVE = ""
MAIN_FOLDER_WEB_URL         = ""
TEST_MODE                   = True   # True=测试模式, False=正式模式（SAP刷新）
ANCHOR_WORKBOOK_NAME        = ""     # 用户预先打开的 Excel 工作簿名称（锚定工作簿）
# ============================================================


# ────────────────────────────────────────────────────────────
#  配置管理（config.json 持久化）
# ────────────────────────────────────────────────────────────

def _get_config_path():
    """获取 config.json 路径（兼容 EXE 打包模式）"""
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "config.json")


CONFIG_FILE = _get_config_path()


def _save_config():
    """将当前配置保存到 config.json"""
    cfg_data = {
        "sharepoint_folder_url":       MAIN_FOLDER_WEB_URL,
        "site_url":                    SITE_URL,
        "main_folder_server_relative": MAIN_FOLDER_SERVER_RELATIVE,
        "main_folder_web_url":         MAIN_FOLDER_WEB_URL,
        "test_mode":                   TEST_MODE,
        "last_updated":                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"配置保存失败: {e}", "WARN")


def load_config_from_file():
    """启动时从 config.json 加载已保存的配置"""
    global SITE_URL, MAIN_FOLDER_SERVER_RELATIVE, MAIN_FOLDER_WEB_URL, TEST_MODE

    if not os.path.exists(CONFIG_FILE):
        return False

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)

        site_url    = cfg.get("site_url", "")
        folder_path = cfg.get("main_folder_server_relative", "")
        web_url     = cfg.get("main_folder_web_url", "")
        mode        = cfg.get("test_mode", True)

        if not site_url or not folder_path:
            return False

        SITE_URL                    = site_url
        MAIN_FOLDER_SERVER_RELATIVE = folder_path
        MAIN_FOLDER_WEB_URL         = web_url
        TEST_MODE                   = mode
        return True

    except Exception:
        return False


def parse_sharepoint_url(url):
    """
    从用户输入的 SharePoint 文件夹 URL 自动解析出三个配置值。
    支持两种 URL 格式：
      1. 路径型: https://xxx.sharepoint.com/teams/site/DocLib/FolderA/Forms/AllItems.aspx
      2. 参数型: https://xxx.sharepoint.com/teams/site/DocLib/Forms/AllItems.aspx?id=%2Fteams%2Fsite%2FDocLib%2FFolderA%2FFolderB&viewid=xxx
    当 URL 中包含 id 查询参数时，优先使用 id 参数作为实际文件夹路径。
    """
    from urllib.parse import urlparse, parse_qs, unquote

    url = url.strip().rstrip("/")
    parsed = urlparse(url)
    host   = f"{parsed.scheme}://{parsed.netloc}"
    path   = parsed.path

    # ── 去掉 /Forms/AllItems.aspx 等后缀，得到文档库路径 ──
    for suffix in ["/Forms/AllItems.aspx", "/Forms/AllItems", "/AllItems.aspx"]:
        if path.lower().endswith(suffix.lower()):
            path = path[: -len(suffix)]
            break

    # ── 检查 id 查询参数（SharePoint 深层文件夹导航时使用）──
    query_params = parse_qs(parsed.query)
    id_param = query_params.get("id", query_params.get("Id", query_params.get("ID", [None])))[0]

    if id_param:
        # id 参数存在，使用它作为实际文件夹的 ServerRelativeUrl
        folder_path = unquote(id_param).rstrip("/")
        log(f"  检测到 id 参数，实际文件夹路径: {folder_path}", "INFO")
    else:
        # 无 id 参数，直接使用 path
        folder_path = path

    # ── 解析 site_url ──
    parts = [p for p in path.split("/") if p]
    if len(parts) >= 2 and parts[0].lower() in ("teams", "sites"):
        site_path = "/" + "/".join(parts[:2])
    else:
        site_path = "/" + parts[0] if parts else ""

    site_url = host + site_path

    # ── 构造 web_url（用于浏览器导航）──
    # 使用文档库路径（path）+ id 参数指向实际文件夹
    if id_param:
        from urllib.parse import quote
        web_url = host + path + "/Forms/AllItems.aspx?id=" + quote(folder_path, safe="/")
    else:
        web_url = host + folder_path + "/Forms/AllItems.aspx"

    return site_url, folder_path, web_url


# ────────────────────────────────────────────────────────────
#  命令行菜单系统
# ────────────────────────────────────────────────────────────

def show_main_menu():
    """显示主菜单"""
    mode_label = "测试模式（写入 A1 时间戳）" if TEST_MODE else "正式模式（SAP Analysis 刷新）"
    url_label  = MAIN_FOLDER_WEB_URL if MAIN_FOLDER_WEB_URL else "（未设置）"

    print()
    print("=" * 60)
    print("  SharePoint Excel 自动化工具")
    print("=" * 60)
    print()
    print(f"  当前模式  : {mode_label}")
    print(f"  目标文件夹: {url_label}")
    print()
    print("  ┌─────────────────────────────────────┐")
    print("  │  [1] 设置 SharePoint 文件夹地址      │")
    print("  │  [2] 切换运行模式                    │")
    print("  │  [3] 预览处理计划并执行              │")
    print("  │  [4] 退出                           │")
    print("  └─────────────────────────────────────┘")
    print()


def menu_set_url():
    """菜单选项1：设置 SharePoint 文件夹地址"""
    global SITE_URL, MAIN_FOLDER_SERVER_RELATIVE, MAIN_FOLDER_WEB_URL

    print()
    print("  ── 设置 SharePoint 文件夹地址 ──")
    print()
    if MAIN_FOLDER_WEB_URL:
        print(f"  当前地址: {MAIN_FOLDER_WEB_URL}")
        print()
    print("  请输入 SharePoint 主文件夹的网址")
    print("  示例: https://abb.sharepoint.com/teams/iaservice/Test0727/Forms/AllItems.aspx")
    print()

    input_url = input("  地址: ").strip()

    if not input_url:
        print("  [取消] 未输入内容，返回主菜单")
        return

    if not input_url.startswith("https://"):
        print("  [错误] 地址必须以 https:// 开头")
        return

    try:
        site_url, folder_path, web_url = parse_sharepoint_url(input_url)
    except Exception as e:
        print(f"  [错误] URL 解析失败: {e}")
        return

    SITE_URL                    = site_url
    MAIN_FOLDER_SERVER_RELATIVE = folder_path
    MAIN_FOLDER_WEB_URL         = web_url

    print()
    print("  ✓ 解析成功：")
    print(f"    站点地址  : {SITE_URL}")
    print(f"    文件夹路径: {MAIN_FOLDER_SERVER_RELATIVE}")
    print(f"    访问地址  : {MAIN_FOLDER_WEB_URL}")

    _save_config()
    print("  ✓ 配置已保存至 config.json")


def menu_toggle_mode():
    """菜单选项2：切换运行模式"""
    global TEST_MODE

    print()
    print("  ── 切换运行模式 ──")
    print()
    print(f"  当前模式: {'测试模式（写入 A1 时间戳）' if TEST_MODE else '正式模式（SAP Analysis 刷新）'}")
    print()
    print("  [1] 测试模式 — 在 A1 写入时间戳验证流程")
    print("  [2] 正式模式 — 执行 SAP Analysis 数据刷新")
    print()

    choice = input("  请选择 [1/2]: ").strip()

    if choice == "1":
        TEST_MODE = True
        print("  ✓ 已切换为：测试模式（写入 A1 时间戳）")
        _save_config()
    elif choice == "2":
        TEST_MODE = False
        print("  ✓ 已切换为：正式模式（SAP Analysis 刷新）")
        print("  ⚠ 注意：正式模式需要 SAP Analysis for Office 已安装并可用")
        _save_config()
    else:
        print("  [取消] 无效选择，返回主菜单")


def menu_run():
    """菜单选项3：预览处理计划并执行"""
    if not SITE_URL or not MAIN_FOLDER_SERVER_RELATIVE:
        print()
        print("  [错误] 尚未设置 SharePoint 文件夹地址，请先选择 [1] 进行设置")
        return

    mode_label = "测试模式（写入 A1 时间戳）" if TEST_MODE else "正式模式（SAP Analysis 刷新）"
    print()
    print(f"  运行模式: {mode_label}")
    print(f"  目标地址: {MAIN_FOLDER_WEB_URL}")
    print()

    confirm_start = input("  输入 Y 连接 Edge 并构建处理计划：").strip().upper()
    if confirm_start != "Y":
        print("  [取消] 返回主菜单")
        return

    pythoncom.CoInitialize()

    with sync_playwright() as p:
        browser = connect_enterprise_edge(p)
        if browser is None:
            return

        try:
            context = browser.contexts[0]
            page    = context.pages[0] if context.pages else context.new_page()
            log("已成功连接到企业版 Edge 浏览器", "OK")
        except Exception as e:
            log(f"获取浏览器页面失败: {e}", "ERROR")
            return

        file_queue = build_file_queue(page)
        if not file_queue:
            log("未找到任何待处理文件，返回主菜单", "WARN")
            return

        # 打印处理计划预览
        print()
        print("─" * 55)
        print("处理计划预览：")
        print("─" * 55)
        for idx, f in enumerate(file_queue, 1):
            suffix_tag = f.get("suffix_tag", "")
            print(f"  {idx:>2}. [{f['folder']}] [{suffix_tag}]  {f['name']}")
        print("─" * 55)
        print(f"  共 {len(file_queue)} 个文件")
        print(f"  运行模式: {mode_label}")
        if not TEST_MODE:
            print(f"  刷新超时: {REFRESH_TIMEOUT} 秒")
        print("─" * 55)

        confirm_exec = input("\n  输入 Y 开始执行，其他键取消：").strip().upper()
        if confirm_exec != "Y":
            log("用户取消执行，返回主菜单", "INFO")
            return

        process_file_queue(file_queue)

    print()


# ────────────────────────────────────────────────────────────
#  工具函数
# ────────────────────────────────────────────────────────────

def log(msg, level="INFO"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    icons = {"INFO": "ℹ ", "OK": "✓ ", "ERROR": "✗ ",
             "WAIT": "⏳", "STEP": "▶ ", "WARN": "⚠ "}
    print(f"[{timestamp}] {icons.get(level, '  ')} {msg}")


# ────────────────────────────────────────────────────────────
#  SharePoint 数据获取（REST API + 已登录浏览器 Session）
# ────────────────────────────────────────────────────────────

def _get_direct_subfolders(page, folder_server_relative):
    """获取指定文件夹下的直接子文件夹列表（单层）"""
    api_url = (
        f"{SITE_URL}/_api/web"
        f"/GetFolderByServerRelativeUrl(\'{folder_server_relative}\')"
        f"/Folders?$select=Name,ServerRelativeUrl&$orderby=Name"
    )
    try:
        result = page.evaluate(f"""
            async () => {{
                const resp = await fetch("{api_url}", {{
                    headers: {{"Accept": "application/json;odata=verbose"}}
                }});
                if (!resp.ok) throw new Error("HTTP " + resp.status);
                const data = await resp.json();
                return data.d.results;
            }}
        """)
    except Exception as e:
        log(f"  REST API 调用失败 [{folder_server_relative}]: {e}", "ERROR")
        return []

    if not result:
        return []

    folders = [
        {"name": item["Name"], "server_url": item["ServerRelativeUrl"]}
        for item in result
        if not item["Name"].startswith("_")
    ]
    return folders


def get_subfolders_via_api(page):
    """
    递归获取主文件夹（用户输入 URL 对应的文件夹）下所有层级的子文件夹。
    使用广度优先遍历（BFS），确保所有深层子文件夹都能被发现。
    返回包含直接子文件夹及其所有后代文件夹的列表，按路径字母顺序排序。
    """
    log(f"正在递归读取子文件夹列表...", "WAIT")
    log(f"  起始文件夹: {MAIN_FOLDER_SERVER_RELATIVE}", "INFO")

    all_folders = []
    queue = [MAIN_FOLDER_SERVER_RELATIVE]  # BFS 队列
    visited = set()

    while queue:
        current_path = queue.pop(0)
        if current_path in visited:
            continue
        visited.add(current_path)

        children = _get_direct_subfolders(page, current_path)
        for child in children:
            child_path = child["server_url"]
            if child_path not in visited:
                all_folders.append(child)
                queue.append(child_path)

    all_folders.sort(key=lambda x: x["server_url"].lower())

    if not all_folders:
        log("主文件夹下未找到任何子文件夹", "WARN")
        # 如果没有子文件夹，把主文件夹自身也作为扫描对象
        main_folder_name = MAIN_FOLDER_SERVER_RELATIVE.rstrip("/").split("/")[-1]
        all_folders = [{"name": main_folder_name, "server_url": MAIN_FOLDER_SERVER_RELATIVE}]
        log(f"  将主文件夹自身作为扫描对象: {main_folder_name}", "INFO")
    else:
        log(f"共找到 {len(all_folders)} 个子文件夹（含所有层级，已按路径排序）", "OK")
        for f in all_folders:
            # 显示相对于主文件夹的路径，更直观
            rel_path = f["server_url"].replace(MAIN_FOLDER_SERVER_RELATIVE, "").lstrip("/")
            log(f'    {rel_path if rel_path else f["name"]}')

    return all_folders
def get_top_excel_files_via_api(page, folder_server_url, folder_name):
    """获取指定文件夹中修改日期最新的 Excel 文件"""
    api_url = (
        f"{SITE_URL}/_api/web"
        f"/GetFolderByServerRelativeUrl(\'{folder_server_url}\')"
        f"/Files?$select=Name,ServerRelativeUrl,TimeLastModified"
        f"&$orderby=TimeLastModified desc&$top=20"
    )
    try:
        result = page.evaluate(f"""
            async () => {{
                const resp = await fetch("{api_url}", {{
                    headers: {{"Accept": "application/json;odata=verbose"}}
                }});
                if (!resp.ok) throw new Error("HTTP " + resp.status);
                const data = await resp.json();
                return data.d.results;
            }}
        """)
    except Exception as e:
        log(f"  获取 [{folder_name}] 文件列表失败: {e}", "ERROR")
        return []

    if not result:
        return []

    excel_exts = (".xlsx", ".xls", ".xlsm", ".xlsb")
    all_excel = [
        {
            "name":     f["Name"],
            "modified": f["TimeLastModified"],
            "full_url": f"https://abb.sharepoint.com{f['ServerRelativeUrl']}"
        }
        for f in result
        if f["Name"].lower().endswith(excel_exts)
    ]

    selected = []
    for suffix in TARGET_SUFFIXES:
        matched = [
            f for f in all_excel
            if suffix.lower() in os.path.splitext(f["name"])[0].lower()
        ]
        if matched:
            matched.sort(key=lambda x: x["modified"], reverse=True)
            best = matched[0]
            best["suffix_tag"] = suffix
            selected.append(best)
        else:
            log(f"  [WARN] [{folder_name}] 未找到后缀为 [{suffix}] 的文件", "WARN")

    return selected


def build_file_queue(page):
    """构建完整有序文件处理队列"""
    log("=" * 55)
    log("第一阶段：构建文件处理队列", "STEP")
    log("=" * 55)

    log(f"导航至主文件夹...", "WAIT")
    page.goto(MAIN_FOLDER_WEB_URL, wait_until="domcontentloaded", timeout=60000)
    time.sleep(5)

    subfolders = get_subfolders_via_api(page)
    if not subfolders:
        log("未找到任何子文件夹，脚本退出", "ERROR")
        return []

    file_queue = []
    for folder in subfolders:
        log(f"\n  扫描: {folder['name']}", "STEP")
        files = get_top_excel_files_via_api(page, folder["server_url"], folder["name"])
        if not files:
            log("  → 未找到 Excel 文件，跳过", "WARN")
            continue
        for f in files:
            try:
                dt = datetime.fromisoformat(f["modified"].replace("Z", "+00:00"))
                f["modified_display"] = dt.strftime("%Y-%m-%d %H:%M")
            except Exception:
                f["modified_display"] = f["modified"]
            suffix_tag = f.get("suffix_tag", "")
            log(f"  + [{suffix_tag}]  {f['name']}  (修改: {f['modified_display']})")
            f["folder"] = folder["name"]
            file_queue.append(f)

    log(f"\n队列构建完成，共 {len(file_queue)} 个文件待处理", "OK")
    return file_queue


# ────────────────────────────────────────────────────────────
#  Excel 窗口管理（win32com）
# ────────────────────────────────────────────────────────────

def open_file_via_ms_protocol(file_url):
    """通过 ms-excel:ofe|u| 协议触发本地 Excel 以 SharePoint 在线模式打开文件"""
    ms_url = f"ms-excel:ofe|u|{file_url}"
    log(f"  ms-excel 协议触发: {file_url.split('/')[-1]}", "WAIT")
    os.startfile(ms_url)


def get_current_workbook_names():
    """获取当前 Excel 中所有已打开的工作簿名称集合"""
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetActiveObject("Excel.Application")
        return {wb.Name for wb in excel.Workbooks}
    except Exception:
        return set()


def detect_anchor_workbook():
    """
    检测用户预先打开的 Excel 工作簿（锚定工作簿）。
    用户在运行程序前已打开一个 Excel 并登录 SAP 账户，
    该工作簿在整个运行期间始终保持打开，不会被程序关闭。

    作用：
      1. 确保 Excel 进程和 COM 连接始终稳定
      2. 先开后关机制不再需要复杂处理
      3. SAP 插件已完成初始化，首文件无 COM 断连风险

    返回：
      str — 锚定工作簿名称，未检测到返回空字符串
    """
    global ANCHOR_WORKBOOK_NAME
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb_names = [wb.Name for wb in excel.Workbooks]

        if wb_names:
            ANCHOR_WORKBOOK_NAME = wb_names[0]
            log(f"  检测到锚定工作簿: [{ANCHOR_WORKBOOK_NAME}]", "OK")
            log(f"  当前已打开 {len(wb_names)} 个工作簿: {wb_names}", "INFO")
            return ANCHOR_WORKBOOK_NAME
        else:
            log("  ⚠ 未检测到已打开的 Excel 工作簿", "WARN")
            log("  请先打开一个 Excel 文件并登录 SAP 账户", "WARN")
            return ""
    except Exception as e:
        log(f"  ⚠ 无法连接 Excel 进程: {e}", "WARN")
        log("  请先打开一个 Excel 文件并登录 SAP 账户", "WARN")
        return ""


def open_and_wait_for_new_workbook(file_url, filename):
    """打开文件并等待新工作簿出现，验证路径为 SharePoint 在线地址。
    增强：深层 COM 稳定性验证，确保 SAP Analysis 插件初始化完成。"""
    log(f"  [Open in App] {filename}", "STEP")
    names_before = get_current_workbook_names()
    open_file_via_ms_protocol(file_url)

    log(f"  等待 Excel 加载（最长 {EXCEL_LOAD_TIMEOUT}s）...", "WAIT")
    deadline = time.time() + EXCEL_LOAD_TIMEOUT
    while time.time() < deadline:
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.GetActiveObject("Excel.Application")
            new_names = {wb.Name for wb in excel.Workbooks} - names_before
            if new_names:
                new_wb_name = list(new_names)[0]
                # ── 等待文件加载和 SAP Analysis 插件初始化 ────
                init_wait = 10
                log(f"  工作簿已检测到，等待 {init_wait}s 确保加载及 SAP 插件初始化完成...", "WAIT")
                time.sleep(init_wait)

                # ── 深层 COM 连接稳定性验证 ───────────────────
                # 不仅检查 Application.Ready 和 Name，还验证
                # Sheets.Count / UsedRange 等深层 COM 操作，
                # 确保 SAP 插件初始化完成后 COM 连接真正可用。
                wb_obj = None
                for wb in excel.Workbooks:
                    if wb.Name == new_wb_name:
                        wb_obj = wb
                        break

                if wb_obj is not None:
                    com_stable = False
                    max_checks = 10          # 最多检查 10 次（共 ~30s）
                    check_wait = 3           # 每次失败后等待 3s
                    for check_i in range(max_checks):
                        try:
                            # 基础检查
                            _ = wb_obj.Application.Ready
                            _ = wb_obj.Name
                            # 深层检查：验证 COM 可进行实际工作簿操作
                            _ = wb_obj.Sheets.Count
                            _ = wb_obj.Application.Version
                            try:
                                _ = wb_obj.ReadOnly
                            except Exception:
                                pass  # ReadOnly 在某些情况下可能暂时不可用，不阻塞
                            com_stable = True
                            log(f"  COM 连接稳定性验证通过（第 {check_i+1} 次）", "OK")
                            break
                        except Exception as ce:
                            log(f"  COM 连接验证第 {check_i+1}/{max_checks} 次未通过: {ce}", "WARN")
                            time.sleep(check_wait)
                            # 重新获取 COM 对象（刷新连接）
                            try:
                                pythoncom.CoInitialize()
                                excel = win32com.client.GetActiveObject("Excel.Application")
                                for wb in excel.Workbooks:
                                    if wb.Name == new_wb_name:
                                        wb_obj = wb
                                        break
                            except Exception:
                                pass

                    if com_stable:
                        wb_path = wb_obj.Path
                        if wb_path.lower().startswith("https://"):
                            log(f"  SharePoint 在线模式确认", "OK")
                        else:
                            log(f"  [WARN] 文件路径为本地路径: {wb_path}", "WARN")

                        # ── 延迟二次验证：捕捉 SAP 插件延迟断连 ──
                        log(f"  等待 5s 后进行二次 COM 稳定性验证...", "WAIT")
                        time.sleep(5)
                        try:
                            pythoncom.CoInitialize()
                            excel_r = win32com.client.GetActiveObject("Excel.Application")
                            for wb_r in excel_r.Workbooks:
                                if wb_r.Name == new_wb_name:
                                    wb_obj = wb_r
                                    break
                            _ = wb_obj.Application.Ready
                            _ = wb_obj.Sheets.Count
                            _ = wb_obj.Name
                            log(f"  二次 COM 稳定性验证通过", "OK")
                        except Exception as e2:
                            log(f"  二次验证未通过: {e2}，等待 15s 后重试...", "WARN")
                            time.sleep(15)
                            try:
                                pythoncom.CoInitialize()
                                excel_r2 = win32com.client.GetActiveObject("Excel.Application")
                                for wb_r2 in excel_r2.Workbooks:
                                    if wb_r2.Name == new_wb_name:
                                        wb_obj = wb_r2
                                        break
                                _ = wb_obj.Application.Ready
                                _ = wb_obj.Sheets.Count
                                log(f"  二次验证重试后通过", "OK")
                            except Exception as e3:
                                log(f"  二次验证重试仍未通过: {e3}，继续（由重试机制兜底）", "WARN")

                        return wb_obj
                    else:
                        log(f"  COM 连接 {max_checks} 次验证后仍不稳定，尝试最后刷新...", "WARN")
                        # 最后一次尝试：重新获取并等待
                        time.sleep(5)
                        try:
                            pythoncom.CoInitialize()
                            excel = win32com.client.GetActiveObject("Excel.Application")
                            for wb in excel.Workbooks:
                                if wb.Name == new_wb_name:
                                    wb_obj = wb
                                    break
                            _ = wb_obj.Application.Ready
                            log(f"  最终刷新后 COM 连接恢复", "OK")
                            return wb_obj
                        except Exception:
                            log(f"  COM 连接不稳定，将继续等待...", "WARN")
        except Exception:
            pass
        time.sleep(1.5)

    log(f"  超时：未检测到 [{filename}] 的 Excel 窗口", "ERROR")
    return None
# ────────────────────────────────────────────────────────────
#  SAP Analysis for Office 操作（pywinauto）
# ────────────────────────────────────────────────────────────

def focus_excel_window(wb_name):
    """将指定工作簿的 Excel 窗口置于前台"""
    try:
        from pywinauto import Application
        title_keyword = os.path.splitext(wb_name)[0]
        app = Application(backend="uia").connect(
            title_re=f".*{re.escape(title_keyword)}.*",
            class_name="XLMAIN",
            timeout=10
        )
        win = app.top_window()
        win.set_focus()
        time.sleep(0.8)
        return app, win
    except Exception as e:
        log(f"  定位 Excel 窗口失败: {e}", "ERROR")
        return None, None


def click_analysis_workbook_prompt(workbook):
    """
    点击 Excel 功能区：
      Analysis（选项卡）→ 数据分析（组）→ 提示▼（下拉）→ 工作簿提示
    """
    from pywinauto import Application
    import pywinauto.mouse as pw_mouse

    wb_name = workbook.Name
    log(f"  点击 Analysis → 提示▼ → 工作簿提示...", "STEP")

    try:
        title_keyword = os.path.splitext(wb_name)[0]
        app = Application(backend="uia").connect(
            title_re=f".*{re.escape(title_keyword)}.*",
            class_name="XLMAIN",
            timeout=10
        )
        win = app.top_window()
        win.set_focus()
        time.sleep(0.8)

        # 第一步：点击 Analysis 选项卡
        try:
            analysis_tab = win.child_window(title="Analysis", control_type="TabItem")
            analysis_tab.click_input()
            log("  [1/3] Analysis 选项卡 已点击", "INFO")
            time.sleep(1)
        except Exception as e:
            log(f"  Analysis 选项卡未找到: {e}", "ERROR")
            return False

        # 第二步：点击"提示"下拉按钮
        try:
            prompt_btn = win.child_window(title="提示", control_type="SplitButton")
            rect = prompt_btn.rectangle()
            arrow_x = rect.right - 8
            arrow_y = rect.top + rect.height() // 2
            pw_mouse.click(button="left", coords=(arrow_x, arrow_y))
            log("  [2/3] 提示▼ 下拉箭头 已点击", "INFO")
            time.sleep(0.8)
        except Exception as e:
            log(f"  SplitButton 未找到，尝试 Button 类型: {e}", "WARN")
            try:
                prompt_btn = win.child_window(title="提示", control_type="Button")
                prompt_btn.click_input()
                log("  [2/3] 提示 Button 已点击", "INFO")
                time.sleep(0.8)
            except Exception as e2:
                log(f"  提示按钮未找到: {e2}", "ERROR")
                return False

        # 第三步：点击"工作簿提示"菜单项
        try:
            wb_prompt = win.child_window(title="工作簿提示", control_type="MenuItem")
            wb_prompt.click_input()
            log("  [3/3] 工作簿提示 已点击", "OK")
            time.sleep(0.5)
            return True
        except Exception as e:
            log(f"  工作簿提示菜单项未找到: {e}", "ERROR")
            return False

    except Exception as e:
        log(f"  点击 Analysis 操作失败: {e}", "ERROR")
        return False






def wait_for_server_data_fetch(wb_name, timeout=180, poll_interval=1):
    """
    等待 SAP "正在从服务器获取数据" 弹窗消失。

    逻辑：
      - 同时检测"获取数据"弹窗和"SAP 确认对话框"
      - 如果 SAP 确认对话框已出现 → 立即返回（数据获取已完成，可以点确认了）
      - 10s 内未检测到获取数据弹窗 → 直接返回
      - 检测到获取数据弹窗 → 等待消失（最长 timeout 秒），期间若确认对话框出现也立即返回

    参数：
      wb_name      : 工作簿名称
      timeout      : 弹窗出现后最大等待时间（默认 180s）
      poll_interval: 轮询间隔（默认 1s）

    返回：
      dict — {"detected": bool, "wait_seconds": float, "result": str}
    """
    from pywinauto import Desktop

    DETECT_WINDOW = 10  # 未检测到弹窗的最大等待秒数

    fetch_keywords = [
        "正在从服务器获取数据", "从服务器获取数据", "获取数据",
        "Retrieving data", "Fetching data", "Loading data",
        "retrieving data from server", "Reading data",
        "正在读取", "正在加载", "正在获取",
        "Please wait", "请稍候", "请等待",
    ]

    # SAP 确认对话框的按钮关键词（与 auto_click_sap_confirmation_dialog 一致）
    confirm_btn_titles = ["确定", "OK", "确认", "Yes", "是", "Apply", "应用",
                          "ok", "yes", "Ok", "YES", "确 定", "OK "]

    result_info = {"detected": False, "wait_seconds": 0, "result": "NOT_DETECTED"}
    start_time = time.time()
    detected_once = False
    detect_time = None

    log(f"  监测数据获取弹窗（{DETECT_WINDOW}s 内无弹窗则跳过，发现确认对话框则立即继续）...", "WAIT")

    while True:
        elapsed = time.time() - start_time

        # ── 超时判断 ──────────────────────────────────────
        if detected_once:
            if elapsed > timeout:
                result_info["result"] = "TIMEOUT"
                result_info["wait_seconds"] = round(elapsed, 1)
                log(f"  [超时] 服务器数据获取超过 {timeout}s 未完成", "ERROR")
                return result_info
        else:
            if elapsed > DETECT_WINDOW:
                log(f"  {DETECT_WINDOW}s 内未检测到数据获取弹窗，跳过", "INFO")
                return result_info

        # ── 优先检测：SAP 确认对话框是否已出现 ──────────────
        # 如果确认对话框已弹出，说明数据获取已完成，应立即进入点击确认步骤
        if _detect_confirm_dialog_exists(wb_name, confirm_btn_titles):
            wait_sec = round(time.time() - start_time, 1)
            result_info["wait_seconds"] = wait_sec
            if detected_once:
                result_info["detected"] = True
                result_info["result"] = "COMPLETED"
                log(f"  ✓ 检测到确认对话框，数据获取已完成（耗时 {wait_sec}s）", "OK")
            else:
                result_info["result"] = "CONFIRM_DIALOG_READY"
                log(f"  ✓ 确认对话框已就绪，无需等待数据获取（{wait_sec}s）", "OK")
            return result_info

        found_fetch_dialog = False

        # ── 检测: 全局搜索获取数据弹窗 ──────────────────────
        try:
            desktop = Desktop(backend="uia")
            all_wins = desktop.windows()
            for w in all_wins:
                try:
                    w_title = w.window_text()
                    w_title_lower = w_title.lower()

                    if any(kw.lower() in w_title_lower for kw in fetch_keywords):
                        found_fetch_dialog = True
                        if not detected_once:
                            detected_once = True
                            detect_time = time.time()
                            log(f"  检测到数据获取弹窗: [{w_title}]", "INFO")
                        break

                    try:
                        static_texts = w.descendants(control_type="Text")
                        for st in static_texts:
                            st_text = st.window_text()
                            if any(kw.lower() in st_text.lower() for kw in fetch_keywords):
                                found_fetch_dialog = True
                                if not detected_once:
                                    detected_once = True
                                    detect_time = time.time()
                                    log(f"  检测到数据获取弹窗（内容匹配）: [{st_text[:50]}]", "INFO")
                                break
                        if found_fetch_dialog:
                            break
                    except Exception:
                        pass
                except Exception:
                    continue
        except Exception:
            pass

        # ── 检测: 状态栏 ────────────────────────────────────
        if not found_fetch_dialog:
            try:
                from pywinauto import Application as _App
                title_keyword = os.path.splitext(wb_name)[0]
                _app = _App(backend="uia").connect(
                    title_re=f".*{re.escape(title_keyword)}.*",
                    class_name="XLMAIN", timeout=3
                )
                _win = _app.top_window()
                try:
                    status_bar = _win.child_window(control_type="StatusBar")
                    children = status_bar.children()
                    status_text = " ".join([c.window_text() for c in children if c.window_text()])
                    if any(kw.lower() in status_text.lower() for kw in fetch_keywords):
                        found_fetch_dialog = True
                        if not detected_once:
                            detected_once = True
                            detect_time = time.time()
                            log(f"  状态栏显示数据获取中: [{status_text[:60]}]", "INFO")
                except Exception:
                    pass
            except Exception:
                pass

        # ── 弹窗消失 = 数据获取完成 ────────────────────────
        if detected_once and not found_fetch_dialog:
            wait_sec = time.time() - detect_time
            result_info["detected"] = True
            result_info["wait_seconds"] = round(wait_sec, 1)
            result_info["result"] = "COMPLETED"
            log(f'  ✓ 服务器数据获取完成（耗时 {result_info["wait_seconds"]}s）', "OK")
            return result_info

        # 进度打印
        if detected_once and int(elapsed) > 0 and int(elapsed) % 20 == 0:
            log(f"  [{int(elapsed)}s] 服务器数据获取中，继续等待...", "WAIT")

        time.sleep(poll_interval)


def _detect_confirm_dialog_exists(wb_name, confirm_btn_titles):
    """
    快速检测 SAP 确认对话框是否已出现。
    通过搜索桌面上所有窗口中是否存在匹配的确认按钮来判断。
    轻量级检测，不执行点击。

    返回 True = 确认对话框已出现
    """
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import find_windows

    # ── 检测 Excel 窗口内的对话框 ──────────────────────────
    try:
        title_keyword = os.path.splitext(wb_name)[0]
        app = Application(backend="uia").connect(
            title_re=f".*{re.escape(title_keyword)}.*",
            class_name="XLMAIN", timeout=2
        )
        win = app.top_window()

        # 搜索对话框子窗口中的确认按钮
        for ctrl_type in ["Window", "Dialog", "Pane"]:
            try:
                dialogs = win.children(control_type=ctrl_type)
                for dlg in dialogs:
                    try:
                        buttons = dlg.descendants(control_type="Button")
                        for btn in buttons:
                            if btn.window_text().strip() in confirm_btn_titles:
                                return True
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass

    # ── 检测独立 SAP 弹窗 ──────────────────────────────────
    sap_keywords = ["SAP", "Analysis", "Prompt", "提示", "确认",
                    "Workbook", "工作簿", "Refresh", "刷新"]
    try:
        desktop = Desktop(backend="uia")
        for w in desktop.windows():
            try:
                w_title = w.window_text()
                if any(kw.lower() in w_title.lower() for kw in sap_keywords):
                    buttons = w.descendants(control_type="Button")
                    for btn in buttons:
                        if btn.window_text().strip() in confirm_btn_titles:
                            return True
            except Exception:
                continue
    except Exception:
        pass

    # ── 检测 #32770 系统对话框 ─────────────────────────────
    try:
        for handle in find_windows(class_name="#32770"):
            try:
                dlg_app = Application(backend="uia").connect(handle=handle)
                dlg_win = dlg_app.top_window()
                buttons = dlg_win.descendants(control_type="Button")
                for btn in buttons:
                    if btn.window_text().strip() in confirm_btn_titles:
                        return True
            except Exception:
                continue
    except Exception:
        pass

    return False




def _click_cancel_on_fetch_dialog(fetch_keywords):
    """
    在"正在从服务器获取数据"弹窗上点击"取消"按钮。
    搜索策略：
      1. 通过窗口标题匹配获取数据弹窗，在其中搜索取消按钮
      2. 全局搜索所有窗口中的取消按钮
      3. 搜索 #32770 系统对话框中的取消按钮
    """
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import find_windows

    cancel_btn_titles = [
        "取消", "Cancel", "cancel", "CANCEL",
        "取 消", "Abort", "abort", "中止", "停止",
    ]

    log(f"  正在搜索取消按钮...", "WAIT")

    # ── 策略 1：在标题匹配的获取数据弹窗中搜索 ──────────
    try:
        desktop = Desktop(backend="uia")
        for w in desktop.windows():
            try:
                w_title = w.window_text().lower()
                if any(kw.lower() in w_title for kw in fetch_keywords):
                    # 找到获取数据弹窗，搜索取消按钮
                    buttons = w.descendants(control_type="Button")
                    for btn in buttons:
                        btn_text = btn.window_text().strip()
                        if btn_text in cancel_btn_titles:
                            btn.click_input()
                            log(f"  ✓ 已点击取消按钮 [{btn_text}]（获取数据弹窗）", "OK")
                            time.sleep(1)
                            return True
            except Exception:
                continue
    except Exception:
        pass

    # ── 策略 2：全局搜索含取消按钮的 SAP 相关窗口 ────────
    sap_keywords = ["SAP", "Analysis", "BEx", "请稍候", "Please wait",
                    "获取", "Retriev", "Loading", "读取"]
    try:
        desktop = Desktop(backend="uia")
        for w in desktop.windows():
            try:
                w_title = w.window_text()
                if any(kw.lower() in w_title.lower() for kw in sap_keywords + fetch_keywords):
                    buttons = w.descendants(control_type="Button")
                    for btn in buttons:
                        btn_text = btn.window_text().strip()
                        if btn_text in cancel_btn_titles:
                            btn.click_input()
                            log(f"  ✓ 已点击取消按钮 [{btn_text}]（窗口: {w_title[:30]}）", "OK")
                            time.sleep(1)
                            return True
            except Exception:
                continue
    except Exception:
        pass

    # ── 策略 3：搜索 #32770 系统对话框 ───────────────────
    try:
        for handle in find_windows(class_name="#32770"):
            try:
                dlg_app = Application(backend="uia").connect(handle=handle)
                dlg_win = dlg_app.top_window()
                buttons = dlg_win.descendants(control_type="Button")
                for btn in buttons:
                    btn_text = btn.window_text().strip()
                    if btn_text in cancel_btn_titles:
                        btn.click_input()
                        log(f"  ✓ 已点击取消按钮 [{btn_text}]（系统对话框）", "OK")
                        time.sleep(1)
                        return True
            except Exception:
                continue
    except Exception:
        pass

    # ── 策略 4：键盘 Escape 兜底 ─────────────────────────
    try:
        import pywinauto.keyboard as kb
        kb.send_keys("{ESC}")
        log(f"  已发送 ESC 键（兜底取消）", "INFO")
        time.sleep(1)
        return True
    except Exception:
        pass

    log(f"  ⚠ 未找到取消按钮", "WARN")
    return False

def wait_and_handle_sap_popups(wb_name, fetch_timeout=180, confirm_timeout=30):
    """
    点击"工作簿提示"后的统一弹窗处理函数。
    并行检测两种弹窗，根据实际情况自动响应：

    场景 A（快速）：直接弹出 SAP 确认对话框 → 立即点击确认
    场景 B（等待）：先弹出"正在从服务器获取数据" → 等待消失 → 再点击确认

    每轮轮询同时检测：
      1. SAP 确认对话框（含确定/OK 等按钮）→ 发现即点击
      2. "正在从服务器获取数据"弹窗 → 发现则标记为等待中

    参数：
      wb_name         : 工作簿名称
      fetch_timeout   : 获取数据弹窗最大等待时间（默认 180s）
      confirm_timeout : 无任何弹窗时的超时退出时间（默认 30s）

    返回：
      dict — {
        "fetch_info":       {"detected": bool, "wait_seconds": float, "result": str},
        "confirm_clicked":  bool,
        "scenario":         "A_DIRECT_CONFIRM" / "B_FETCH_THEN_CONFIRM" / "TIMEOUT"
      }
    """
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import find_windows

    # ── 关键词定义 ────────────────────────────────────────
    fetch_keywords = [
        "正在从服务器获取数据", "从服务器获取数据", "获取数据",
        "Retrieving data", "Fetching data", "Loading data",
        "retrieving data from server", "Reading data",
        "正在读取", "正在加载", "正在获取",
        "Please wait", "请稍候", "请等待",
    ]

    confirm_btn_titles = ["确定", "OK", "确认", "Yes", "是", "Apply", "应用",
                          "ok", "yes", "Ok", "YES", "确 定", "OK "]

    sap_dialog_keywords = ["SAP", "Analysis", "Prompt", "提示", "确认",
                           "Workbook", "工作簿", "Refresh", "刷新"]

    # ── 结果结构 ──────────────────────────────────────────
    result = {
        "fetch_info":      {"detected": False, "wait_seconds": 0, "result": "NOT_DETECTED"},
        "confirm_clicked": False,
        "scenario":        "TIMEOUT",
    }

    start_time = time.time()
    fetch_detected = False
    fetch_detect_time = None
    fetch_disappeared = False
    fetch_disappear_time = None
    poll_interval = 1

    log(f"  ── 统一弹窗监测（并行检测获取数据/确认对话框）──", "STEP")

    while True:
        elapsed = time.time() - start_time

        # ── 超时判断 ──────────────────────────────────────
        if fetch_detected and not fetch_disappeared:
            # 正在等待获取数据弹窗消失
            if elapsed > fetch_timeout:
                log(f"  [超时] 服务器数据获取超过 {fetch_timeout}s，尝试点击取消按钮...", "ERROR")

                # ── 超时后自动点击"取消"按钮释放 Excel ──
                cancel_clicked = _click_cancel_on_fetch_dialog(fetch_keywords)
                result["fetch_info"]["result"] = "TIMEOUT_CANCELLED" if cancel_clicked else "TIMEOUT"
                result["fetch_info"]["wait_seconds"] = round(elapsed, 1)
                return result
        elif fetch_detected and fetch_disappeared:
            # 获取数据弹窗已消失，最多再等 5s 确认按钮
            since_disappeared = time.time() - fetch_disappear_time
            if since_disappeared > 5:
                log(f"  获取数据弹窗消失后 5s 内未检测到确认按钮，跳过确认步骤", "INFO")
                result["scenario"] = "B_FETCH_NO_CONFIRM"
                return result
        else:
            # 无获取数据弹窗的情况下，confirm_timeout 秒后退出
            if elapsed > confirm_timeout:
                log(f"  {int(elapsed)}s 内未检测到任何弹窗，超时退出", "WARN")
                return result

        # ── 扫描桌面所有窗口（一次扫描，多重判断）─────────
        found_fetch = False
        confirm_btn_ref = None  # 找到的确认按钮引用

        try:
            desktop = Desktop(backend="uia")
            all_wins = desktop.windows()

            for w in all_wins:
                try:
                    w_title = w.window_text()
                    w_title_lower = w_title.lower()

                    # ── 判断是否为"获取数据"弹窗 ──────────
                    if any(kw.lower() in w_title_lower for kw in fetch_keywords):
                        found_fetch = True
                        if not fetch_detected:
                            fetch_detected = True
                            fetch_detect_time = time.time()
                            log(f"  检测到数据获取弹窗: [{w_title}]", "INFO")
                        continue  # 获取数据弹窗本身不含确认按钮，跳过

                    # ── 判断是否为 SAP 确认对话框 ─────────
                    is_sap_dialog = any(kw.lower() in w_title_lower for kw in sap_dialog_keywords)
                    if is_sap_dialog:
                        btn = _find_confirm_button_in_window(w, confirm_btn_titles)
                        if btn is not None:
                            confirm_btn_ref = btn
                            break

                except Exception:
                    continue

            # 补充：检查窗口内静态文本匹配获取数据关键词
            if not found_fetch and not fetch_detected:
                for w in all_wins:
                    try:
                        texts = w.descendants(control_type="Text")
                        for t in texts:
                            if any(kw.lower() in t.window_text().lower() for kw in fetch_keywords):
                                found_fetch = True
                                fetch_detected = True
                                fetch_detect_time = time.time()
                                log(f"  检测到数据获取弹窗（内容匹配）: [{t.window_text()[:50]}]", "INFO")
                                break
                        if found_fetch:
                            break
                    except Exception:
                        continue

        except Exception:
            pass

        # 补充：Excel 窗口内搜索确认按钮
        if confirm_btn_ref is None:
            try:
                title_keyword = os.path.splitext(wb_name)[0]
                app = Application(backend="uia").connect(
                    title_re=f".*{re.escape(title_keyword)}.*",
                    class_name="XLMAIN", timeout=2
                )
                win = app.top_window()
                for ctrl_type in ["Window", "Dialog", "Pane"]:
                    try:
                        for dlg in win.children(control_type=ctrl_type):
                            btn = _find_confirm_button_in_window(dlg, confirm_btn_titles)
                            if btn is not None:
                                confirm_btn_ref = btn
                                break
                    except Exception:
                        pass
                    if confirm_btn_ref:
                        break
            except Exception:
                pass

        # 补充：#32770 系统对话框
        if confirm_btn_ref is None:
            try:
                for handle in find_windows(class_name="#32770"):
                    try:
                        dlg_app = Application(backend="uia").connect(handle=handle)
                        dlg_win = dlg_app.top_window()
                        btn = _find_confirm_button_in_window(dlg_win, confirm_btn_titles)
                        if btn is not None:
                            confirm_btn_ref = btn
                            break
                    except Exception:
                        continue
            except Exception:
                pass

        # ── 处理逻辑 ─────────────────────────────────────

        # 情况 1: 找到确认按钮 → 点击并返回
        if confirm_btn_ref is not None:
            try:
                btn_text = confirm_btn_ref.window_text().strip()
                confirm_btn_ref.click_input()
                time.sleep(0.5)
                result["confirm_clicked"] = True

                if fetch_detected:
                    fetch_wait = round(time.time() - fetch_detect_time, 1) if fetch_detect_time else 0
                    result["fetch_info"] = {"detected": True, "wait_seconds": fetch_wait, "result": "COMPLETED"}
                    result["scenario"] = "B_FETCH_THEN_CONFIRM"
                    log(f"  ✓ 场景B: 数据获取({fetch_wait}s) → 点击确认 [{btn_text}]", "OK")
                else:
                    result["scenario"] = "A_DIRECT_CONFIRM"
                    log(f"  ✓ 场景A: 直接点击确认 [{btn_text}]（耗时 {round(elapsed,1)}s）", "OK")

                return result

            except Exception as e:
                log(f"  点击确认按钮失败: {e}", "WARN")

        # 情况 2: 获取数据弹窗消失，但确认按钮还没出现 → 标记消失，限时 5s 等确认
        if fetch_detected and not found_fetch and not fetch_disappeared:
            fetch_disappeared = True
            fetch_disappear_time = time.time()
            fetch_wait = round(fetch_disappear_time - fetch_detect_time, 1)
            log(f"  数据获取弹窗已消失（耗时 {fetch_wait}s），等待确认对话框（最多 5s）...", "INFO")
            result["fetch_info"] = {"detected": True, "wait_seconds": fetch_wait, "result": "COMPLETED"}

        # 进度打印
        if int(elapsed) > 0 and int(elapsed) % 15 == 0:
            status = "数据获取中" if (fetch_detected and not fetch_disappeared) else "等待弹窗"
            log(f"  [{int(elapsed)}s] {status}...", "WAIT")

        time.sleep(poll_interval)


def _find_confirm_button_in_window(window, confirm_titles):
    """
    在指定窗口中查找确认按钮，返回按钮对象（不点击）。
    先搜索直接子控件，再搜索所有后代控件。
    返回 None 表示未找到。
    """
    try:
        # 直接子按钮
        for btn in window.children(control_type="Button"):
            if btn.window_text().strip() in confirm_titles:
                return btn
        # 深层搜索
        for btn in window.descendants(control_type="Button"):
            if btn.window_text().strip() in confirm_titles:
                return btn
    except Exception:
        pass
    return None



def auto_click_sap_confirmation_dialog(wb_name, timeout=15, poll_interval=1):
    """
    自动检测并点击 SAP Analysis 弹出的确认对话框。
    点击"工作簿提示"后，SAP 会弹出一个窗口要求确认刷新，
    本函数自动查找该对话框并点击"确定/OK"按钮。

    查找策略（按优先级）：
      1. 在 Excel 主窗口下查找 Dialog/Window 类型的子窗口
      2. 查找独立的 SAP 弹窗进程
      3. 兜底：用 pywinauto find_windows 全局搜索

    按钮匹配关键词（中英文兼容）：
      确定 / OK / 确认 / Yes / 是 / Apply / 应用

    参数：
      wb_name      : 当前工作簿名称（用于定位 Excel 窗口）
      timeout      : 最长等待对话框出现的秒数（默认 15s）
      poll_interval: 轮询间隔秒数（默认 1s）

    返回：
      True  — 成功点击了确认按钮
      False — 超时未发现对话框或点击失败
    """
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import find_windows

    log(f"  等待 SAP 确认对话框弹出（最长 {timeout}s）...", "WAIT")

    # 常见确认按钮的标题关键词
    confirm_titles = ["确定", "OK", "确认", "Yes", "是", "Apply", "应用",
                      "ok", "yes", "Ok", "YES", "确 定", "OK "]

    start_time = time.time()

    while time.time() - start_time < timeout:
        try:
            # ── 策略 1: 在 Excel 主窗口中查找对话框子窗口 ──────
            title_keyword = os.path.splitext(wb_name)[0]
            try:
                app = Application(backend="uia").connect(
                    title_re=f".*{re.escape(title_keyword)}.*",
                    class_name="XLMAIN",
                    timeout=3
                )
                win = app.top_window()

                # 查找 Dialog / Window / Pane 类型的子窗口
                for ctrl_type in ["Window", "Dialog", "Pane"]:
                    try:
                        dialogs = win.children(control_type=ctrl_type)
                        for dlg in dialogs:
                            clicked = _try_click_confirm_button(dlg, confirm_titles)
                            if clicked:
                                return True
                    except Exception:
                        pass

                # 查找 Excel 窗口内所有 Button，直接匹配确认按钮
                try:
                    all_buttons = win.descendants(control_type="Button")
                    for btn in all_buttons:
                        btn_title = btn.window_text().strip()
                        if btn_title in confirm_titles:
                            btn.click_input()
                            log(f"  ✓ 已点击确认按钮 [{btn_title}]（Excel 窗口内）", "OK")
                            time.sleep(0.5)
                            return True
                except Exception:
                    pass

            except Exception:
                pass

            # ── 策略 2: 查找独立的 SAP 弹窗 ────────────────────
            sap_keywords = ["SAP", "Analysis", "Prompt", "提示", "确认",
                            "Workbook", "工作簿", "Refresh", "刷新"]
            try:
                desktop = Desktop(backend="uia")
                all_wins = desktop.windows()
                for w in all_wins:
                    try:
                        w_title = w.window_text()
                        if any(kw.lower() in w_title.lower() for kw in sap_keywords):
                            log(f"  发现疑似 SAP 对话框: [{w_title}]", "INFO")
                            clicked = _try_click_confirm_button(w, confirm_titles)
                            if clicked:
                                return True
                    except Exception:
                        continue
            except Exception:
                pass

            # ── 策略 3: 全局查找 #32770 类型对话框（Windows 通用对话框）──
            try:
                dialog_handles = find_windows(class_name="#32770")
                for handle in dialog_handles:
                    try:
                        dlg_app = Application(backend="uia").connect(handle=handle)
                        dlg_win = dlg_app.top_window()
                        dlg_title = dlg_win.window_text()
                        log(f"  发现系统对话框: [{dlg_title}]", "INFO")
                        clicked = _try_click_confirm_button(dlg_win, confirm_titles)
                        if clicked:
                            return True
                    except Exception:
                        continue
            except Exception:
                pass

        except Exception as e:
            log(f"  对话框检测异常: {e}", "WARN")

        time.sleep(poll_interval)

    log(f"  未检测到 SAP 确认对话框（等待 {timeout}s 后超时）", "WARN")
    return False


def _try_click_confirm_button(window, confirm_titles):
    """
    在指定窗口中查找并点击确认按钮。
    遍历所有 Button 控件，匹配 confirm_titles 中的任一标题。

    返回 True 表示成功点击，False 表示未找到。
    """
    try:
        buttons = window.children(control_type="Button")
        for btn in buttons:
            btn_title = btn.window_text().strip()
            if btn_title in confirm_titles:
                btn.click_input()
                log(f"  ✓ 已点击确认按钮 [{btn_title}]", "OK")
                time.sleep(0.5)
                return True

        # 如果直接子控件没找到，尝试 descendants（深层搜索）
        buttons_deep = window.descendants(control_type="Button")
        for btn in buttons_deep:
            btn_title = btn.window_text().strip()
            if btn_title in confirm_titles:
                btn.click_input()
                log(f"  ✓ 已点击确认按钮 [{btn_title}]（深层搜索）", "OK")
                time.sleep(0.5)
                return True

    except Exception:
        pass
    return False


# ────────────────────────────────────────────────────────────
#  SAP 刷新完成检测
# ────────────────────────────────────────────────────────────

def _wait_for_new_data_workbook(known_wb_names, timeout=120):
    """
    确认点击后，等待新的数据表格出现。
    排除锚定工作簿和确认前已知的工作簿。
    期间同时处理可能出现的"获取数据"弹窗。

    参数:
      known_wb_names : 确认前已打开的工作簿名称集合
      timeout        : 最长等待时间（秒）

    返回:
      workbook COM 对象（新数据表格），未检测到返回 None
    """
    log(f"  ── 等待新数据表格出现（最长 {timeout}s）──", "STEP")

    start_time = time.time()
    fetch_seen = False

    while time.time() - start_time < timeout:
        elapsed = time.time() - start_time

        # ── 检测"获取数据"弹窗（确认后可能再次出现）──
        if _detect_server_fetch_dialog(""):
            if not fetch_seen:
                fetch_seen = True
                log(f"  确认后检测到获取数据弹窗，等待消失...", "INFO")
            # 弹窗存在时继续等待
            if int(elapsed) > 0 and int(elapsed) % 15 == 0:
                log(f"  [{int(elapsed)}s] 获取数据弹窗仍在...", "WAIT")
            time.sleep(1)
            continue
        elif fetch_seen:
            log(f"  确认后获取数据弹窗已消失（耗时 {round(elapsed, 1)}s）", "INFO")
            fetch_seen = False  # 重置，继续检测新表格

        # ── 检测新工作簿 ──────────────────────────────────
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.GetActiveObject("Excel.Application")
            current_names = {wb.Name for wb in excel.Workbooks}
            new_names = current_names - known_wb_names

            # 排除锚定工作簿
            if ANCHOR_WORKBOOK_NAME:
                new_names.discard(ANCHOR_WORKBOOK_NAME)

            if new_names:
                new_wb_name = list(new_names)[0]
                log(f"  检测到新工作簿: [{new_wb_name}]（等待 {round(elapsed,1)}s）", "INFO")

                # 等待 3s 确保新表格加载完成
                time.sleep(3)

                # 获取 COM 对象
                for wb in excel.Workbooks:
                    if wb.Name == new_wb_name:
                        # 激活新工作簿
                        try:
                            wb.Activate()
                            time.sleep(1)
                        except Exception:
                            pass
                        return wb

        except Exception:
            pass

        # 进度打印
        if int(elapsed) > 0 and int(elapsed) % 10 == 0:
            log(f"  [{int(elapsed)}s] 等待新数据表格...", "WAIT")

        time.sleep(1)

    log(f"  [超时] {timeout}s 内未检测到新数据表格", "WARN")
    return None


def _get_target_workbook_exclude_anchor():
    """
    获取除锚定工作簿外的唯一目标工作簿 COM 对象。
    预期 Excel 中只有 2 个工作簿：锚定 + 目标。
    使用 win32com.client.Dispatch 确保返回可调用方法的 COM 对象。
    """
    try:
        pythoncom.CoInitialize()
        # 使用 Dispatch 而非 GetActiveObject，确保 COM 方法可正常调用
        excel = win32com.client.Dispatch("Excel.Application")
        target_name = None
        for wb in excel.Workbooks:
            if wb.Name != ANCHOR_WORKBOOK_NAME:
                if target_name is None:
                    target_name = wb.Name
                else:
                    log(f"  发现多个非锚定工作簿，使用第一个: [{target_name}]", "WARN")
                    break

        if target_name is None:
            return None

        # 通过名称重新获取，确保 COM 绑定正确
        try:
            wb_obj = excel.Workbooks(target_name)
            return wb_obj
        except Exception:
            # 备选：遍历获取
            for wb in excel.Workbooks:
                if wb.Name == target_name:
                    return wb
            return None
    except Exception as e:
        log(f"  获取目标工作簿失败: {e}", "WARN")
        return None


def _reactivate_target_workbook(wb_name):
    """
    确认点击后 Excel 可能跳转到其他工作簿（如 SAP 数据库表格），
    此函数将焦点切回目标工作簿，确保后续检测针对正确的文件。
    排除锚定工作簿。
    """
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetActiveObject("Excel.Application")

        # 检查当前活动工作簿是否是目标
        try:
            active_name = excel.ActiveWorkbook.Name
            if active_name == wb_name:
                log(f"  当前活动工作簿已是目标: [{wb_name}]", "OK")
                return True
            else:
                log(f"  当前活动工作簿为 [{active_name}]，需切回 [{wb_name}]", "INFO")
        except Exception:
            pass

        # 切换到目标工作簿
        for wb in excel.Workbooks:
            if wb.Name == wb_name:
                wb.Activate()
                time.sleep(1)
                log(f"  已切换活动工作簿至 [{wb_name}]", "OK")
                return True

        log(f"  未找到工作簿 [{wb_name}]", "WARN")
        return False
    except Exception as e:
        log(f"  切换活动工作簿失败: {e}", "WARN")
        return False


def _get_non_anchor_workbook_name():
    """
    获取当前 Excel 中除锚定工作簿外的工作簿名称。
    用于确认点击后 Excel 跳转时，重新定位目标工作簿。
    """
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetActiveObject("Excel.Application")
        for wb in excel.Workbooks:
            if wb.Name != ANCHOR_WORKBOOK_NAME:
                return wb.Name
    except Exception:
        pass
    return None


def wait_for_refresh_complete(workbook):
    """
    轮询检测 SAP Analysis 数据刷新是否完成。
    检测维度（四重确认）：
      1. Application.Ready == True（Excel 可交互）
      2. Application.CalculationState == 0（xlDone，计算完成）
      3. 状态栏检测（pywinauto 读取 Excel 底部状态栏文字）
      4. 无"正在从服务器获取数据"弹窗存在

    智能启动等待：
      - 不再使用固定 5s 等待
      - 改为轮询检测 Excel 是否进入忙碌状态（Ready=False 或有数据获取弹窗）
      - 检测到忙碌后开始计时；若 10s 内未检测到忙碌则继续（兼容无延迟场景）

    返回值:
      True  — 刷新完成
      False — 超时未完成
    """
    excel = workbook.Application
    wb_name = workbook.Name
    timeout = REFRESH_TIMEOUT
    interval = REFRESH_POLL_INTERVAL

    log(f"  ── 开始检测刷新状态（超时 {timeout}s）──", "STEP")

    # ── 智能等待刷新动作启动（替代固定 5s）──────────────────
    log(f"  智能等待刷新动作启动...", "WAIT")
    launch_wait_start = time.time()
    refresh_started = False
    while time.time() - launch_wait_start < 15:
        try:
            if not excel.Ready:
                refresh_started = True
                log(f"  检测到 Excel 进入忙碌状态（Application.Ready=False）", "INFO")
                break
        except Exception:
            refresh_started = True
            log(f"  Excel COM 无响应，刷新可能已启动", "INFO")
            break

        # 也检查是否有服务器获取数据弹窗
        if _detect_server_fetch_dialog(wb_name):
            refresh_started = True
            log(f"  检测到服务器数据获取弹窗，刷新已启动", "INFO")
            break

        time.sleep(0.5)

    if not refresh_started:
        log(f"  15s 内未检测到明确的刷新启动信号，继续监测...", "INFO")

    start_time = time.time()
    last_status_log = 0
    consecutive_ready_count = 0   # 连续就绪计数（防止误判）
    CONSECUTIVE_READY_THRESHOLD = 3  # 连续 3 次检测到就绪才确认完成

    while True:
        elapsed = time.time() - start_time

        # 超时检查
        if elapsed > timeout:
            log(f"  [超时] 已等待 {int(elapsed)}s，刷新未在限定时间内完成", "ERROR")
            return False

        # ── 检测 1: Application.Ready ──────────────────────────
        try:
            app_ready = excel.Ready
        except Exception:
            app_ready = False

        # ── 检测 2: CalculationState ──────────────────────────
        # xlDone = 0, xlCalculating = 1, xlPending = 2
        try:
            calc_state = excel.CalculationState
            calc_done = (calc_state == 0)
        except Exception:
            calc_done = False

        # ── 检测 3: 状态栏文字（pywinauto）─────────────────────
        status_bar_ready = _check_status_bar_ready(wb_name)

        # ── 检测 4: 无服务器获取数据弹窗 ───────────────────────
        server_fetch_active = _detect_server_fetch_dialog(wb_name)
        if server_fetch_active:
            # 有获取数据弹窗，重置连续就绪计数
            consecutive_ready_count = 0

        # ── 检测 5: 确认活动工作簿是目标（非锚定）──────────────
        try:
            active_wb_name = excel.ActiveWorkbook.Name
            if active_wb_name == ANCHOR_WORKBOOK_NAME and active_wb_name != wb_name:
                # Excel 焦点跳到了锚定工作簿，切回目标
                _reactivate_target_workbook(wb_name)
                # 切换后需重新获取状态，本轮不计入就绪
                consecutive_ready_count = 0
                time.sleep(interval)
                continue
        except Exception:
            pass

        # ── 综合判断 ──────────────────────────────────────────
        all_ready = app_ready and calc_done and status_bar_ready and not server_fetch_active
        if all_ready:
            consecutive_ready_count += 1
        else:
            consecutive_ready_count = 0

        if consecutive_ready_count >= CONSECUTIVE_READY_THRESHOLD:
            log(f"  ✓ 刷新完成！（耗时 {int(elapsed)}s，连续 {CONSECUTIVE_READY_THRESHOLD} 次确认就绪）", "OK")
            log(f"    Application.Ready   = True", "INFO")
            log(f"    CalculationState    = xlDone", "INFO")
            log(f"    状态栏              = 就绪", "INFO")
            log(f"    服务器获取数据弹窗  = 无", "INFO")
            return True

        # 每 15 秒打印一次进度
        if elapsed - last_status_log >= 15:
            fetch_str = "活跃" if server_fetch_active else "无"
            log(f"  [{int(elapsed)}s] Ready={app_ready} | CalcDone={calc_done} | StatusBar={status_bar_ready} | ServerFetch={fetch_str}", "WAIT")
            last_status_log = elapsed

        # ── 轮询期间也检测弹窗（处理刷新过程中的二次确认）────
        try:
            auto_click_sap_confirmation_dialog(wb_name, timeout=1, poll_interval=0.5)
        except Exception:
            pass

        time.sleep(interval)



def _detect_server_fetch_dialog(wb_name):
    """
    快速检测是否存在"正在从服务器获取数据"类弹窗。
    轻量级检测（不等待），用于在刷新轮询中快速判断。

    返回 True = 弹窗存在（数据获取中），False = 未检测到
    """
    from pywinauto import Desktop

    fetch_keywords = [
        "正在从服务器获取数据", "从服务器获取数据", "获取数据",
        "Retrieving data", "Fetching data", "Loading data",
        "正在读取", "正在加载", "正在获取",
        "Please wait", "请稍候", "请等待",
    ]

    try:
        desktop = Desktop(backend="uia")
        all_wins = desktop.windows()
        for w in all_wins:
            try:
                w_title = w.window_text()
                if any(kw.lower() in w_title.lower() for kw in fetch_keywords):
                    return True
                # 也检查窗口内的静态文本
                try:
                    texts = w.descendants(control_type="Text")
                    for t in texts:
                        t_text = t.window_text()
                        if any(kw.lower() in t_text.lower() for kw in fetch_keywords):
                            return True
                except Exception:
                    pass
            except Exception:
                continue
    except Exception:
        pass
    return False


def _check_status_bar_ready(wb_name):
    """
    通过 pywinauto 读取 Excel 底部状态栏文字。
    检查是否包含"就绪"/"Ready" 关键字。
    若无法读取则返回 True（不阻塞流程）。
    排除锚定工作簿窗口，确保检测的是目标工作簿。
    """
    try:
        from pywinauto import Application

        # 排除锚定工作簿：确保匹配的是目标文件而非锚定文件
        title_keyword = os.path.splitext(wb_name)[0]
        if ANCHOR_WORKBOOK_NAME:
            anchor_keyword = os.path.splitext(ANCHOR_WORKBOOK_NAME)[0]
            if title_keyword == anchor_keyword:
                return True  # 不检测锚定工作簿的状态栏

        app = Application(backend="uia").connect(
            title_re=f".*{re.escape(title_keyword)}.*",
            class_name="XLMAIN",
            timeout=5
        )
        win = app.top_window()
        
        # 二次确认窗口标题不是锚定工作簿
        win_title = win.window_text()
        if ANCHOR_WORKBOOK_NAME:
            anchor_kw = os.path.splitext(ANCHOR_WORKBOOK_NAME)[0]
            if anchor_kw.lower() in win_title.lower() and title_keyword.lower() not in win_title.lower():
                return True  # 误匹配到锚定窗口，不阻塞

        # 尝试读取状态栏
        try:
            status_bar = win.child_window(auto_id="StatusBar", control_type="StatusBar")
            status_text = status_bar.window_text()
        except Exception:
            # 备用方式：尝试通过 control_type 查找
            try:
                status_bar = win.child_window(control_type="StatusBar")
                # 获取所有子元素的文本
                children = status_bar.children()
                status_text = " ".join([c.window_text() for c in children if c.window_text()])
            except Exception:
                # 无法读取状态栏，默认不阻塞
                return True

        if not status_text:
            return True

        # 判断是否就绪
        ready_keywords = ["就绪", "Ready", "ready"]
        busy_keywords  = ["正在计算", "Calculating", "正在刷新", "Refreshing",
                          "正在连接", "Connecting", "请等待", "Please wait",
                          "正在从服务器获取数据", "获取数据", "Retrieving",
                          "Fetching", "Loading", "正在读取", "正在加载",
                          "正在获取", "请稍候", "数据传输"]

        # 如果包含忙碌关键字，明确返回 False
        for kw in busy_keywords:
            if kw.lower() in status_text.lower():
                return False

        # 如果包含就绪关键字，返回 True
        for kw in ready_keywords:
            if kw.lower() in status_text.lower():
                return True

        # 状态栏文字不明确，默认返回 True（不阻塞）
        return True

    except Exception:
        # 连接失败，不阻塞
        return True


def get_used_row_count(workbook):
    """获取第一个可见 Sheet 的已使用行数"""
    try:
        sheet = get_first_visible_sheet(workbook)
        row_count = sheet.UsedRange.Rows.Count
        return row_count
    except Exception:
        return -1


# ────────────────────────────────────────────────────────────
#  Excel 写入 / 诊断 / 关闭
# ────────────────────────────────────────────────────────────

def get_first_visible_sheet(workbook):
    """返回工作簿中第一个可见的 Sheet，跳过 SAP 隐藏 Sheet"""
    try:
        for i in range(1, workbook.Sheets.Count + 1):
            sh = workbook.Sheets(i)
            if sh.Visible == -1:
                log(f"  目标Sheet: [{sh.Name}]（第{i}个，可见）", "INFO")
                return sh
        log(f"  未找到可见Sheet，回退到 Sheets(1): [{workbook.Sheets(1).Name}]", "WARN")
        return workbook.Sheets(1)
    except Exception as e:
        log(f"  get_first_visible_sheet 失败，回退到 Sheets(1): {e}", "WARN")
        return workbook.Sheets(1)


def get_fresh_workbook(wb_name):
    """重新获取最新的工作簿 COM 对象"""
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.GetActiveObject("Excel.Application")
        for wb in excel.Workbooks:
            if wb.Name == wb_name:
                return wb
        log(f"  get_fresh_workbook: 未找到工作簿 [{wb_name}]", "WARN")
        return None
    except Exception as e:
        log(f"  get_fresh_workbook 失败: {e}", "WARN")
        return None


def write_a1_via_com(workbook, value, diag):
    """方法1：通过 win32com 直接写入 A1 单元格"""
    try:
        sheet = get_first_visible_sheet(workbook)
        cell  = sheet.Range("A1")
        cell.NumberFormat = "@"
        cell.Value = value
        log(f"  [COM] A1 写入: {value}  Sheet: [{sheet.Name}]", "INFO")
        diag["a1_write_result"] = f"COM_SUCCESS: {value}"
        diag["target_sheet"]    = sheet.Name
        return True
    except Exception as e:
        log(f"  [COM] 写入失败: {e}", "WARN")
        diag["a1_write_result"] = f"COM_FAILED: {e}"
        return False


def write_a1_via_sendkeys(wb_name, value, diag):
    """方法2（兜底）：通过 pywinauto SendKeys 模拟键盘写入 A1"""
    try:
        from pywinauto import Application
        from pywinauto.keyboard import send_keys

        title_keyword = os.path.splitext(wb_name)[0]
        app = Application(backend="uia").connect(
            title_re=f".*{re.escape(title_keyword)}.*",
            class_name="XLMAIN",
            timeout=10
        )
        win = app.top_window()
        win.set_focus()
        time.sleep(0.5)

        try:
            name_box = win.child_window(auto_id="Box", control_type="Edit")
            name_box.click_input()
        except Exception:
            send_keys("^{HOME}")
            time.sleep(0.3)
            send_keys(value)
            send_keys("{ENTER}")
            log(f"  [SendKeys-备用] A1 写入完成: {value}", "INFO")
            diag["a1_write_result"] = f"SENDKEYS_FALLBACK_SUCCESS: {value}"
            return True

        time.sleep(0.3)
        send_keys("A1{ENTER}")
        time.sleep(0.4)
        send_keys(value)
        send_keys("{ENTER}")
        time.sleep(0.3)

        log(f"  [SendKeys] A1 写入完成: {value}", "INFO")
        diag["a1_write_result"] = f"SENDKEYS_SUCCESS: {value}"
        return True

    except Exception as e:
        log(f"  [SendKeys] 写入失败: {e}", "ERROR")
        diag["a1_write_result"] = f"SENDKEYS_FAILED: {e}"
        return False


def run_sap_analysis_refresh(workbook, is_first_file):
    """
    根据 TEST_MODE 选择执行模式。
    增强重试机制：进入时预存 wb_name，防止 COM 断连后无法获取名称。
    """

    MAX_RETRIES = 5
    RETRY_WAIT  = 20

    # ── 关键：COM 可用时立即缓存工作簿名称 ────────────────
    saved_wb_name = None
    try:
        saved_wb_name = workbook.Name
    except Exception:
        pass

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            if attempt > 1:
                if saved_wb_name is None:
                    log(f"  [重试 {attempt}/{MAX_RETRIES}] 无法获取工作簿名称，终止重试", "ERROR")
                    break

                log(f"  [重试 {attempt}/{MAX_RETRIES}] 等待 {RETRY_WAIT}s 后重新获取 COM 对象...", "WARN")
                time.sleep(RETRY_WAIT)
                pythoncom.CoInitialize()

                fresh_wb = None
                for stab_check in range(8):
                    fresh_wb = get_fresh_workbook(saved_wb_name)
                    if fresh_wb is not None:
                        try:
                            _ = fresh_wb.Application.Ready
                            _ = fresh_wb.Sheets.Count
                            _ = fresh_wb.Name
                            log(f"  [重试 {attempt}/{MAX_RETRIES}] COM 稳定性验证通过（第 {stab_check+1}/8 次）", "OK")
                            break
                        except Exception:
                            log(f"  [重试 {attempt}/{MAX_RETRIES}] COM 验证第 {stab_check+1}/8 次未通过，等待 5s...", "WARN")
                            fresh_wb = None
                            time.sleep(5)
                    else:
                        log(f"  [重试 {attempt}/{MAX_RETRIES}] 获取工作簿失败（第 {stab_check+1}/8 次），等待 5s...", "WARN")
                        time.sleep(5)

                if fresh_wb is not None:
                    workbook = fresh_wb
                else:
                    log(f"  [重试 {attempt}/{MAX_RETRIES}] 无法获取稳定的工作簿对象，继续...", "WARN")
                    continue

            result = _do_sap_refresh(workbook, is_first_file)
            return result

        except Exception as e:
            err_msg = str(e).lower()
            is_com_error = any(kw in err_msg for kw in [
                "没有连接到服务器", "not connected", "disconnected",
                "rpc 服务器不可用", "rpc server", "调用被拒绝",
                "call was rejected", "对象没有连接",
                "server not available", "服务器不可用",
                "应用程序正在忙", "application is busy",
                "对象已与其客户端断开连接", "has disconnected",
            ])

            if is_com_error and attempt < MAX_RETRIES:
                log(f"  [COM 错误] {e}", "WARN")
                log(f"  COM 连接不稳定，将进行第 {attempt+1} 次重试...", "WARN")
                continue
            else:
                log(f"  处理流程出错: {e}", "ERROR")
                return False

    log(f"  重试 {MAX_RETRIES} 次后仍失败", "ERROR")
    return False
def _do_sap_refresh(workbook, is_first_file):
    """
    实际执行 SAP 刷新/测试写入的内部函数。
    由 run_sap_analysis_refresh 调用，支持重试机制。
    """
    try:
        excel = workbook.Application

        # ── 退出受保护视图 ───────────────────────────────────
        for pv in excel.ProtectedViewWindows:
            try:
                if pv.Workbook.Name == workbook.Name:
                    pv.Edit()
                    time.sleep(1)
                    log("  已退出受保护视图", "INFO")
                    break
            except Exception:
                break

        # ── 检查只读状态 ─────────────────────────────────────
        if workbook.ReadOnly:
            log("  [WARN] 文件以只读模式打开，尝试启用编辑...", "WARN")
            try:
                workbook.LockServerFile()
            except Exception:
                pass
            time.sleep(1)

        # ════════════════════════════════════════════════════════
        # 测试模式：写入 A1 时间戳
        # ════════════════════════════════════════════════════════
        if TEST_MODE:
            now_str = datetime.now().strftime(TIME_FORMAT)

            diag = diagnose_workbook(workbook)
            save_diagnosis_log(diag)

            if diag["is_readonly"]:
                log("  [诊断] 文件只读，等待 6s 后重新获取工作簿对象...", "WARN")
                time.sleep(6)
                workbook = get_fresh_workbook(workbook.Name)
                if workbook is None:
                    log("  重新获取工作簿失败，跳过写入", "ERROR")
                    diag["a1_write_result"] = "FAILED: 只读且重新获取工作簿失败"
                    save_diagnosis_log(diag)
                    return False

            write_ok = write_a1_via_com(workbook, now_str, diag)

            if not write_ok:
                log("  win32com 写入失败，切换 pywinauto SendKeys 兜底写入...", "WARN")
                write_ok = write_a1_via_sendkeys(workbook.Name, now_str, diag)

            if not write_ok:
                log("  两种写入方式均失败，请查看 diagnosis_log.txt", "ERROR")
                save_diagnosis_log(diag)
                return False

            # 二次确认
            try:
                workbook = get_fresh_workbook(workbook.Name)
                if workbook:
                    visible_sh = get_first_visible_sheet(workbook)
                    actual     = visible_sh.Range("A1").Value
                    actual_str = str(actual).strip().replace("+00:00", "").strip()
                    if actual_str == now_str.strip():
                        log(f"  二次确认 A1 = [{actual}]  写入成功", "OK")
                        diag["a1_write_result"] = f"VERIFIED: {actual}"
                    else:
                        log(f"  二次确认 A1 = [{actual}]  与写入值不符", "WARN")
                        diag["a1_write_result"] = f"MISMATCH: expected={now_str}, actual={actual}"
            except Exception as verify_err:
                log(f"  二次确认读取失败: {verify_err}", "WARN")

            try:
                workbook.Save()
                log("  已保存至 SharePoint", "OK")
            except Exception as save_err:
                log(f"  保存失败: {save_err}", "ERROR")
                diag["a1_write_result"] += f" | SAVE_FAILED: {save_err}"

            save_diagnosis_log(diag)
            return True

        # ════════════════════════════════════════════════════════
        # 正式模式：SAP Analysis for Office 数据刷新
        # ════════════════════════════════════════════════════════
        else:
            wb_name = workbook.Name

            # ── 记录刷新前行数 ────────────────────────────────
            rows_before = get_used_row_count(workbook)
            log(f"  刷新前行数: {rows_before}", "INFO")

            # ── SAP 登录：用户已预先登录，无需等待 ─────────
            # （用户在运行程序前已打开 Excel 并登录 SAP 账户）

            # ── 操作结果跟踪（用于 diagnosis_log 详细记录）────
            op_record = {
                "timestamp":          datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "wb_name":            wb_name,
                "mode":               "FORMAL_SAP_REFRESH",
                "folder":             "",
                "rows_before":        rows_before,
                "rows_after":         -1,
                "row_diff":           0,
                "click_analysis_ok":  False,
                "server_fetch":       {"detected": False, "wait_seconds": 0, "result": "N/A"},
                "confirm_dialog_ok":  False,
                "refresh_ok":         False,
                "refresh_elapsed_s":  0,
                "save_ok":            False,
                "error_detail":       "",
                "is_readonly":        False,
                "protected_view":     False,
                "excel_version":      "N/A",
            }

            try:
                op_record["is_readonly"]    = workbook.ReadOnly
                op_record["excel_version"]  = workbook.Application.Version
                op_record["protected_view"] = False
                try:
                    for pv in workbook.Application.ProtectedViewWindows:
                        if pv.Workbook.Name == workbook.Name:
                            op_record["protected_view"] = True
                            break
                except Exception:
                    pass
            except Exception:
                pass

            refresh_start_time = time.time()

            # ── 操作前 COM 预检查（防止 SAP 插件初始化导致断连）───
            log(f"  操作前 COM 预检查...", "WAIT")
            for pre_check in range(3):
                try:
                    _ = workbook.Application.Ready
                    _ = workbook.Sheets.Count
                    _ = workbook.Name
                    log(f"  COM 预检查通过", "OK")
                    break
                except Exception as pre_e:
                    log(f"  COM 预检查第 {pre_check+1} 次失败: {pre_e}", "WARN")
                    if pre_check < 2:
                        time.sleep(5)
                        # 尝试刷新 COM 对象
                        try:
                            pythoncom.CoInitialize()
                            fresh = get_fresh_workbook(wb_name)
                            if fresh is not None:
                                workbook = fresh
                                log(f"  COM 对象已刷新", "OK")
                        except Exception:
                            pass
                    else:
                        # 三次都失败，抛出异常触发重试机制
                        raise pre_e

            # ── 点击 Analysis → 提示▼ → 工作簿提示 ───────────
            success = click_analysis_workbook_prompt(workbook)
            op_record["click_analysis_ok"] = success
            if not success:
                log("  [WARN] 点击操作失败，跳过刷新", "WARN")
                op_record["error_detail"] = "click_analysis_workbook_prompt failed"
                _save_formal_diagnosis_log(op_record)
                return False

            # ── 阶段一：等待并处理弹窗（获取数据 + 确认对话框并行检测）──
            popup_result = wait_and_handle_sap_popups(wb_name, fetch_timeout=180, confirm_timeout=30)
            op_record["server_fetch"]      = popup_result["fetch_info"]
            op_record["confirm_dialog_ok"] = popup_result["confirm_clicked"]

            # ── 阶段二：确认后循环等待 20s 并尝试保存目标工作簿 ──
            # 目标工作簿 = 除锚定表格外的唯一打开表格
            MAX_SAVE_ATTEMPTS = 10   # 最多尝试 10 次（共 ~200s）
            SAVE_WAIT         = 20   # 每次等待秒数

            save_ok = False
            for save_attempt in range(1, MAX_SAVE_ATTEMPTS + 1):
                log(f"  [{save_attempt}/{MAX_SAVE_ATTEMPTS}] 等待 {SAVE_WAIT}s 后尝试保存...", "WAIT")
                time.sleep(SAVE_WAIT)

                # 定位目标工作簿（除锚定外的唯一表格）
                target_wb = _get_target_workbook_exclude_anchor()
                if target_wb is None:
                    log(f"  [{save_attempt}/{MAX_SAVE_ATTEMPTS}] 未找到目标工作簿，继续等待...", "WARN")
                    continue

                target_name = target_wb.Name
                log(f"  [{save_attempt}/{MAX_SAVE_ATTEMPTS}] 目标工作簿: [{target_name}]", "INFO")

                # 尝试保存（兼容 COM 绑定异常）
                try:
                    if callable(getattr(target_wb, 'Save', None)):
                        target_wb.Save()
                    else:
                        # COM 晚绑定时 Save 可能不是 callable，通过 Dispatch 重新获取
                        pythoncom.CoInitialize()
                        excel_re = win32com.client.Dispatch("Excel.Application")
                        excel_re.Workbooks(target_name).Save()
                    log(f"  ✓ 已保存: [{target_name}]", "OK")
                    op_record["save_ok"] = True
                    save_ok = True

                    # 记录行数
                    rows_after = get_used_row_count(target_wb)
                    row_diff = rows_after - rows_before if (rows_before >= 0 and rows_after >= 0) else 0
                    op_record["rows_after"] = rows_after
                    op_record["row_diff"]   = row_diff
                    log(f"  保存后行数: {rows_after}", "INFO")
                    _save_refresh_log(target_name, rows_before, rows_after, row_diff, True)
                    break

                except Exception as save_err:
                    log(f"  [{save_attempt}/{MAX_SAVE_ATTEMPTS}] 保存失败: {save_err}，继续等待...", "WARN")
                    continue

            if not save_ok:
                log(f"  {MAX_SAVE_ATTEMPTS} 次尝试后仍无法保存", "ERROR")
                op_record["error_detail"] += "save_failed_all_attempts; "

            op_record["refresh_ok"]        = save_ok
            op_record["refresh_elapsed_s"] = round(time.time() - refresh_start_time, 1)

            # ── 保存详细诊断日志 ──────────────────────────────
            _save_formal_diagnosis_log(op_record)

            return save_ok

    except Exception as e:
        err_msg = str(e).lower()
        is_com_error = any(kw in err_msg for kw in [
            "没有连接到服务器", "not connected", "disconnected",
            "rpc 服务器不可用", "rpc server", "调用被拒绝",
            "call was rejected", "对象没有连接",
            "server not available", "服务器不可用",
            "应用程序正在忙", "application is busy",
            "对象已与其客户端断开连接", "has disconnected",
        ])
        if is_com_error:
            # COM 连接错误向上抛出，由 run_sap_analysis_refresh 重试
            log(f"  [COM 连接错误] {e}，将尝试重试...", "WARN")
            raise
        else:
            log(f"  处理流程出错: {e}", "ERROR")
            return False


def _save_refresh_log(wb_name, rows_before, rows_after, row_diff, refresh_ok):
    """保存 SAP 刷新结果到 refresh_log.txt"""
    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(script_dir, "refresh_log.txt")

    status = "SUCCESS" if refresh_ok else "TIMEOUT"
    if row_diff > 0:
        change_desc = f"+{row_diff} rows (new data)"
    elif row_diff == 0:
        change_desc = "No change (data identical)"
    else:
        change_desc = f"{row_diff} rows (data reduced)"

    lines = [
        "─" * 50,
        f"Time        : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"File        : {wb_name}",
        f"Status      : {status}",
        f"Rows Before : {rows_before}",
        f"Rows After  : {rows_after}",
        f"Change      : {change_desc}",
        "",
    ]

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    except Exception:
        pass


def diagnose_workbook(workbook):
    """诊断工作簿的当前状态"""
    diag = {
        "timestamp":        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "wb_name":          "N/A",
        "wb_path":          "N/A",
        "path_is_online":   False,
        "is_readonly":      None,
        "sheets_count":     None,
        "sheet1_name":      "N/A",
        "a1_current_value": "N/A",
        "a1_write_result":  "NOT ATTEMPTED",
        "protected_view":   False,
        "excel_version":    "N/A",
        "error_detail":     "",
        "target_sheet":     "N/A",
    }

    try:
        diag["wb_name"]        = workbook.Name
        diag["wb_path"]        = workbook.Path
        diag["path_is_online"] = workbook.Path.lower().startswith("https://")
        diag["is_readonly"]    = workbook.ReadOnly
        diag["excel_version"]  = workbook.Application.Version
        diag["sheets_count"]   = workbook.Sheets.Count
        diag["sheet1_name"]    = workbook.Sheets(1).Name

        visible_sheet = get_first_visible_sheet(workbook)
        diag["target_sheet"] = visible_sheet.Name
        try:
            a1_val = visible_sheet.Range("A1").Value
            diag["a1_current_value"] = str(a1_val) if a1_val is not None else "(空)"
        except Exception as e:
            diag["a1_current_value"] = f"READ_ERROR: {e}"

        try:
            for pv in workbook.Application.ProtectedViewWindows:
                if pv.Workbook.Name == workbook.Name:
                    diag["protected_view"] = True
                    break
        except Exception:
            pass

    except Exception as e:
        diag["error_detail"] = str(e)

    # 输出诊断摘要
    log(f"  ── 工作簿诊断报告 ──────────────────────────", "INFO")
    log(f"  文件名      : {diag['wb_name']}", "INFO")
    log(f"  SharePoint路径: {'是' if diag['path_is_online'] else '否（本地: ' + diag['wb_path'] + '）'}", "INFO")
    log(f"  只读模式    : {'是 ⚠' if diag['is_readonly'] else '否'}", "INFO")
    log(f"  受保护视图  : {'是 ⚠' if diag['protected_view'] else '否'}", "INFO")
    log(f"  Sheet数量   : {diag['sheets_count']}  |  Sheet1名称: {diag['sheet1_name']}", "INFO")
    log(f"  A1当前值    : {diag['a1_current_value']}", "INFO")
    log(f"  Excel版本   : {diag['excel_version']}", "INFO")
    if diag["error_detail"]:
        log(f"  诊断异常    : {diag['error_detail']}", "WARN")
    log(f"  ─────────────────────────────────────────", "INFO")

    return diag


def save_diagnosis_log(diag):
    """将诊断结果追加保存到 diagnosis_log.txt"""
    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(script_dir, "diagnosis_log.txt")

    lines = [
        "=" * 60,
        f"诊断时间        : {diag['timestamp']}",
        f"文件名          : {diag['wb_name']}",
        f"文件路径        : {diag['wb_path']}",
        f"是否SharePoint路径: {'是' if diag['path_is_online'] else '否'}",
        f"只读模式        : {'是' if diag['is_readonly'] else '否'}",
        f"受保护视图      : {'是' if diag['protected_view'] else '否'}",
        f"Sheet数量       : {diag['sheets_count']}",
        f"Sheet1名称      : {diag['sheet1_name']}",
        f"写入目标Sheet   : {diag['target_sheet']}",
        f"A1当前值        : {diag['a1_current_value']}",
        f"A1写入结果      : {diag['a1_write_result']}",
        f"Excel版本       : {diag['excel_version']}",
        f"异常详情        : {diag['error_detail'] if diag['error_detail'] else '无'}",
        "",
    ]

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        log(f"  诊断日志已保存: {log_path}", "OK")
    except Exception as e:
        log(f"  诊断日志保存失败: {e}", "WARN")



def _save_formal_diagnosis_log(op_record):
    """
    将正式模式下每次操作的详细结果保存到 diagnosis_log.txt。
    记录内容包括：
      - 文件信息（名称、文件夹、只读状态、受保护视图）
      - 操作步骤结果（Analysis 点击、服务器获取、确认弹窗、刷新状态）
      - 数据变化（行数对比）
      - 时间信息（总耗时、服务器获取耗时）
      - 错误详情
    """
    if getattr(sys, "frozen", False):
        script_dir = os.path.dirname(sys.executable)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(script_dir, "diagnosis_log.txt")

    fetch_info = op_record.get("server_fetch", {})
    fetch_detected = fetch_info.get("detected", False)
    fetch_wait     = fetch_info.get("wait_seconds", 0)
    fetch_result   = fetch_info.get("result", "N/A")

    # 计算行数变化描述
    row_diff = op_record.get("row_diff", 0)
    if row_diff > 0:
        row_change_desc = f"+{row_diff} 行（新增数据）"
    elif row_diff == 0:
        row_change_desc = "无变化"
    else:
        row_change_desc = f"{row_diff} 行（数据减少，需检查）"

    # 总体结果判定
    if op_record.get("refresh_ok") and op_record.get("save_ok"):
        overall = "SUCCESS ✓"
    elif op_record.get("refresh_ok") and not op_record.get("save_ok"):
        overall = "REFRESH_OK_BUT_SAVE_FAILED ⚠"
    elif not op_record.get("click_analysis_ok"):
        overall = "CLICK_FAILED ✗"
    else:
        overall = "REFRESH_TIMEOUT ⚠"

    lines = [
        "",
        "=" * 70,
        f"  正式模式操作记录 — {op_record.get('timestamp', 'N/A')}",
        "=" * 70,
        f"  总体结果          : {overall}",
        "-" * 70,
        f"  【文件信息】",
        f"    文件名           : {op_record.get('wb_name', 'N/A')}",
        f"    所属文件夹       : {op_record.get('folder', 'N/A')}",
        f"    运行模式         : {op_record.get('mode', 'N/A')}",
        f"    只读模式         : {'是 ⚠' if op_record.get('is_readonly') else '否'}",
        f"    受保护视图       : {'是 ⚠' if op_record.get('protected_view') else '否'}",
        f"    Excel 版本       : {op_record.get('excel_version', 'N/A')}",
        "",
        f"  【操作步骤结果】",
        f"    1. Analysis 点击  : {'成功 ✓' if op_record.get('click_analysis_ok') else '失败 ✗'}",
        f"    2. 服务器获取数据 : {'检测到' if fetch_detected else '未检测到'}",
        f"       - 获取等待时间 : {fetch_wait}s",
        f"       - 获取结果     : {fetch_result}",
        f"    3. 确认对话框点击 : {'成功 ✓' if op_record.get('confirm_dialog_ok') else '未检测到/无需确认'}",
        f"    4. SAP 数据刷新   : {'完成 ✓' if op_record.get('refresh_ok') else '超时/失败 ✗'}",
        f"    5. 保存回SharePoint: {'成功 ✓' if op_record.get('save_ok') else '失败/未执行 ✗'}",
        "",
        f"  【数据变化】",
        f"    刷新前行数       : {op_record.get('rows_before', 'N/A')}",
        f"    刷新后行数       : {op_record.get('rows_after', 'N/A')}",
        f"    行数变化         : {row_change_desc}",
        "",
        f"  【时间统计】",
        f"    总操作耗时       : {op_record.get('refresh_elapsed_s', 0)}s",
        f"    服务器获取耗时   : {fetch_wait}s",
        "",
        f"  【错误详情】",
        f"    {op_record.get('error_detail', '无') if op_record.get('error_detail') else '无'}",
        "=" * 70,
        "",
    ]

    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        log(f"  诊断日志已保存: {log_path}", "OK")
    except Exception as e:
        log(f"  诊断日志保存失败: {e}", "WARN")


def close_workbook_safe(workbook, wb_name_hint=None):
    """关闭工作簿。COM 断连时通过 wb_name_hint 重新获取再关闭。"""
    name = wb_name_hint

    try:
        name = workbook.Name
    except Exception:
        if name:
            log(f"  COM 对象不可用，使用备选名称: {name}", "WARN")
        else:
            log(f"  COM 对象不可用且无备选名称，跳过关闭", "WARN")
            return

    if name == ANCHOR_WORKBOOK_NAME:
        log(f"  [锚定工作簿] 跳过关闭: {name}", "INFO")
        return

    # 尝试原始 COM 对象
    try:
        workbook.Close(SaveChanges=False)
        log(f"  已关闭: {name}", "OK")
        time.sleep(1)
        return
    except Exception:
        pass

    # 重新获取再关闭
    try:
        pythoncom.CoInitialize()
        fresh_wb = get_fresh_workbook(name)
        if fresh_wb is not None:
            fresh_wb.Close(SaveChanges=False)
            log(f"  已关闭（重新获取COM）: {name}", "OK")
            time.sleep(1)
            return
    except Exception:
        pass

    log(f"  无法关闭 [{name}]，COM 连接已断开", "WARN")


# ────────────────────────────────────────────────────────────
#  核心调度逻辑：先开后关
# ────────────────────────────────────────────────────────────

def process_file_queue(file_queue):
    """
    逐个处理规则（锚定工作簿保持打开）：
      Open(File_1) → 处理 → Save → Close(File_1)
      Open(File_2) → 处理 → Save → Close(File_2)
      ...
    锚定工作簿全程不关闭，保证 COM 连接和 SAP 插件稳定。
    """
    if not file_queue:
        log("文件队列为空，跳过处理", "WARN")
        return

    total = len(file_queue)
    log("\n" + "=" * 55)
    log("第二阶段：依次处理 Excel 文件", "STEP")
    log("=" * 55)

    # ── 检测锚定工作簿（用户预先打开的 Excel）──────────────
    anchor = detect_anchor_workbook()
    if not anchor:
        log("  ⚠ 未检测到预先打开的 Excel 工作簿！", "ERROR")
        log("  请先打开一个 Excel 文件并登录 SAP 账户后重试", "ERROR")
        return

    is_first      = True
    success_count = 0
    fail_count    = 0

    for i, file_info in enumerate(file_queue):
        filename = file_info["name"]
        file_url = file_info["full_url"]
        folder   = file_info.get("folder", "")

        log(f"\n[{i+1}/{total}] 文件夹: {folder}", "INFO")

        # 1. 打开文件
        new_wb = open_and_wait_for_new_workbook(file_url, filename)
        if new_wb is None:
            log(f"  [{filename}] 加载失败，跳过", "WARN")
            fail_count += 1
            continue

        # 2. 执行处理
        result = run_sap_analysis_refresh(new_wb, is_first_file=is_first)
        is_first = False

        if result:
            success_count += 1
        else:
            fail_count += 1

        # 3. 处理完毕，立即关闭当前文件（锚定工作簿不会被关闭）
        log("  关闭当前文件...", "INFO")
        close_workbook_safe(new_wb, wb_name_hint=filename)

    # 打印汇总
    log("\n" + "=" * 55)
    log("执行汇总", "STEP")
    log("=" * 55)
    log(f"  总文件数: {total}", "INFO")
    log(f"  成功: {success_count}", "OK")
    log(f"  失败: {fail_count}", "ERROR" if fail_count > 0 else "INFO")
    log("\n全部处理完成！锚定工作簿仍保持打开。", "OK")
# ────────────────────────────────────────────────────────────
#  Edge 连接（Playwright）
# ────────────────────────────────────────────────────────────

def connect_enterprise_edge(playwright_instance):
    """连接到以调试模式启动的企业版 Edge，支持多端口重试"""
    ports_to_try = [EDGE_DEBUG_PORT, 9223, 9224]

    for port in ports_to_try:
        try:
            log(f"尝试连接 Edge 调试端口 {port}...", "WAIT")
            browser = playwright_instance.chromium.connect_over_cdp(
                f"http://localhost:{port}",
                timeout=8000
            )

            contexts = browser.contexts
            if not contexts:
                log(f"  端口 {port} 连接成功但无浏览器上下文，跳过", "WARN")
                continue

            pages = contexts[0].pages
            if not pages:
                log(f"  端口 {port} 无打开页面，尝试新建页面...", "WARN")
                pages = [contexts[0].new_page()]

            current_url = pages[0].url
            log(f"  当前页面: {current_url[:80]}", "INFO")

            if "login.microsoftonline.com" in current_url or "login.live.com" in current_url:
                log(f"  检测到登录页面，请先在 Edge 中完成 SharePoint 账号登录", "WARN")
                input("  完成登录后按 Enter 继续 → ")

            log(f"  ✓ Edge 连接成功（端口 {port}）", "OK")
            return browser

        except Exception as e:
            err_msg = str(e)
            if "Connection refused" in err_msg or "No connection" in err_msg:
                log(f"  端口 {port} 未响应（Edge 未以调试模式启动）", "WARN")
            elif "Timeout" in err_msg:
                log(f"  端口 {port} 连接超时", "WARN")
            else:
                log(f"  端口 {port} 连接失败: {err_msg[:80]}", "WARN")
            continue

    log("", "ERROR")
    log("=" * 55, "ERROR")
    log("无法连接到企业版 Edge，请按以下步骤排查：", "ERROR")
    log("", "INFO")
    log("  1. 确认已运行 launch_edge_debug.bat", "INFO")
    log("  2. 确认 launch_edge_debug.bat 窗口保持打开", "INFO")
    log("  3. 确认企业 IT 策略未禁用远程调试端口", "INFO")
    log("  4. 确认没有其他程序占用端口 9222", "INFO")
    log("     （命令行运行: netstat -ano | findstr 9222）", "INFO")
    log("=" * 55, "ERROR")
    return None


# ────────────────────────────────────────────────────────────
#  主入口
# ────────────────────────────────────────────────────────────

def main():
    """主入口：加载配置 → 菜单循环"""
    load_config_from_file()

    while True:
        show_main_menu()
        choice = input("  请选择 [1-4]: ").strip()

        if choice == "1":
            menu_set_url()
        elif choice == "2":
            menu_toggle_mode()
        elif choice == "3":
            menu_run()
        elif choice == "4":
            print("\n  再见！")
            break
        else:
            print("  [提示] 无效选择，请输入 1-4")


if __name__ == "__main__":
    main()
