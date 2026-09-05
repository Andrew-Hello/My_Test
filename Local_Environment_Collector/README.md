# Local Environment Collector

这个目录用于把本机开发/运行环境一次性整理成可提交到 GitHub 的诊断快照，方便 ChatGPT/Codex 在无法直接访问本机时进行适配。

## 使用方法

双击：

```text
Collect_Local_Environment.cmd
```

脚本会在：

```text
Local_Environment_Collector/reports/
```

生成同一时间戳的：

```text
*_environment.json
*_environment.txt
```

把这两个文件 Commit + Push 到 `dev`，然后告诉 ChatGPT：

> 读取最新 environment report，按我的本机环境适配。

## 当前采集内容

- Windows 版本、Build、系统架构、时区、语言
- CPU、内存、GPU、显示器布局、基础 DPI 信息
- PowerShell 版本与执行策略
- `py` launcher、Python 版本/路径、pip 配置与镜像
- PyMuPDF / Pillow / ImageHash / NumPy / SciPy 的安装与 import 状态
- Word / Excel / PowerPoint COM 可用性、版本与 Build
- Acrobat/PDFMaker 注册状态与 Acrobat/Reader 版本
- Git / Node.js / npm / .NET / Java
- 当前 Git 分支、commit、origin 和已跟踪文件的本地修改状态

## 隐私设计

默认不会主动记录：

- Windows 用户名
- 电脑名
- 密码、Token、API Key
- 任意环境变量全集

用户目录会替换为 `%USERPROFILE%`，URL 中若存在认证信息会进行遮罩。

这个采集器是通用工具，不只服务于 Office PDF 项目。以后本机环境、Python/npm/.NET、Office COM、显示器/DPI 等问题，都可以先生成一次环境快照再分析。
