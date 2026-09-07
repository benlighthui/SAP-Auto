# SharePoint Excel 自动化工具 — 使用说明

## 收到的文件

| 文件名 | 说明 |
|--------|------|
| `SharePoint_Excel_Tool.exe` | 主程序（双击运行，无需安装 Python） |
| `launch_edge_debug.bat` | Edge 启动脚本（必须先运行） |

---

## 使用步骤（3步即可）

### 第一步：启动 Edge（每次使用前执行）

1. 双击 `launch_edge_debug.bat`
2. 等待 Edge 浏览器弹出
3. **在 Edge 中登录你的 SharePoint 账号**（访问目标文件夹确认能正常打开）
4. **保持 bat 窗口不要关闭**

### 第二步：运行主程序

1. 双击 `SharePoint_Excel_Tool.exe`
2. 看到命令行菜单：

```
============================================================
  SharePoint Excel 自动化工具
============================================================

  当前模式  : 测试模式（写入 A1 时间戳）
  目标文件夹: （未设置）

  ┌─────────────────────────────────────┐
  │  [1] 设置 SharePoint 文件夹地址     │
  │  [2] 切换运行模式                   │
  │  [3] 预览处理计划并执行             │
  │  [4] 退出                           │
  └─────────────────────────────────────┘
```

### 第三步：操作菜单

1. **输入 `1`** → 粘贴 SharePoint 文件夹地址
   - 直接从浏览器地址栏复制，例如：
   - `https://abb.sharepoint.com/teams/iaservice/Test0727/Forms/AllItems.aspx`

2. **输入 `2`** → 选择运行模式
   - `1` = 测试模式（仅在 A1 写入时间戳，用于验证流程）
   - `2` = 正式模式（执行 SAP Analysis 数据刷新）

3. **输入 `3`** → 预览待处理文件列表 → 输入 `Y` 开始执行

---

## 注意事项

| 项目 | 说明 |
|------|------|
| Edge 版本 | 所有成员需使用相同大版本的 Edge（如 130.x） |
| SAP Analysis | 正式模式需要电脑已安装 SAP Analysis for Office 插件 |
| 网络要求 | 需能访问 SharePoint（公司内网或 VPN） |
| 首次 SAP 登录 | 正式模式下第一个文件会暂停，提示手动完成 SAP 账号登录 |
| config.json | 自动生成在 EXE 同目录，保存上次设置，下次启动自动加载 |
| 文件不要重命名 | `launch_edge_debug.bat` 不要改名 |

---

## 常见问题

### Q: 提示"无法连接到企业版 Edge"

**原因**：Edge 未以调试模式启动  
**解决**：
1. 关闭所有 Edge 窗口（包括后台进程）
2. 重新双击 `launch_edge_debug.bat`
3. 再运行 EXE

### Q: 提示"端口 9222 已被占用"

**原因**：上次的 Edge 调试进程未完全关闭  
**解决**：
1. 按 `Ctrl+Shift+Esc` 打开任务管理器
2. 结束所有 `msedge.exe` 进程
3. 重新运行 `launch_edge_debug.bat`

### Q: Excel 打开文件超时

**原因**：SharePoint 连接慢或文件较大  
**解决**：等待网络稳定后重试，或检查 SharePoint 是否可正常访问

### Q: 文件以只读模式打开

**原因**：其他人正在编辑该文件  
**解决**：程序会自动尝试解锁，若失败会记录到 `diagnosis_log.txt`

---

## 文件夹结构（部署后）

```
📁 任意文件夹/
├── SharePoint_Excel_Tool.exe    ← 主程序
├── launch_edge_debug.bat        ← Edge 启动脚本
├── config.json                  ← 自动生成（保存配置）
└── diagnosis_log.txt            ← 自动生成（运行日志）
```
