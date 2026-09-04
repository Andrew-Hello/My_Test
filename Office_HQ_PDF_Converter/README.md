# Office HQ PDF Converter

这是一个面向 Windows + Microsoft Office 桌面版的高保真 PDF 转换脚本。

支持：

- Word：`.docx`、`.doc`
- Excel：`.xlsx`、`.xls`
- PowerPoint：`.pptx`、`.ppt`

## 最简单的用法

把一个或多个 Office 文档直接拖到：

```text
DragDrop_Office_to_PDF.cmd
```

脚本会在 **原文档同一目录** 生成 PDF。

例如：

```text
D:\Data\Report.docx
```

生成：

```text
D:\Data\Report.pdf
```

如果同名 PDF 已经存在，为避免覆盖，会依次生成：

```text
Report_HQ.pdf
Report_HQ_2.pdf
Report_HQ_3.pdf
...
```

## 设计目标

这个脚本不是通过“虚拟 PDF 打印机”转换，也不使用 LibreOffice 或第三方 Office 解析器，而是直接调用本机安装的 Microsoft Word / Excel / PowerPoint 原生固定格式 PDF 导出接口。

这样做的目的，是尽量保留：

- 原始 Office 排版
- 矢量文字和矢量图形
- Office 图表
- 高分辨率图片
- 字体显示
- 页面尺寸与幻灯片尺寸
- 原有对象位置

脚本始终以只读方式打开源文档，不会保存或修改原文件。

## 高质量策略

### Word

优先尝试新版 Word 的 `ExportAsFixedFormat3`，并显式启用：

```text
OptimizeFor = Print
OptimizeForImageQuality = True
BitmapMissingFonts = True
UseISO19005_1 = False
```

如果当前 Word 版本不支持，再自动降级到 `ExportAsFixedFormat2`，最后才使用经典 `ExportAsFixedFormat`。

其中 `OptimizeForImageQuality=True` 的目的就是避免在 PDF 导出阶段继续降低图片质量。

### Excel

使用 Excel 原生：

```text
ExportAsFixedFormat
Quality = xlQualityStandard
```

`xlQualityStandard` 是 Excel 固定格式导出接口中的最高质量档。

同时脚本设置：

```text
IgnorePrintAreas = True
```

这样可以避免因为工作簿中历史遗留的“打印区域”而把区域外的图片、图表或单元格内容完全排除在 PDF 外。

注意：这意味着某些原本故意设置了打印区域的 Excel 文件，生成的 PDF 页数可能会增加。

### PowerPoint

使用 PowerPoint 原生 `ExportAsFixedFormat`，设置：

```text
FixedFormatType = PDF
Intent = Print
OutputType = Slides
BitmapMissingFonts = True
UseISO19005_1 = False
```

`Print` 意图比 `Screen` 更适合高质量输出。

## 自动诊断模式

现在每次转换都会自动在：

```text
Office_HQ_PDF_Converter/diagnostics/
```

生成 JSON 诊断文件，不需要额外操作。

单个文件诊断示例：

```text
20260904_112233_456_Report.json
```

批量转换还会生成：

```text
20260904_112240_123_batch_summary.json
```

诊断内容包括：

- Windows / PowerShell / .NET 环境
- Office 应用与版本
- 实际调用的 PDF 导出接口
- 源文件与 PDF 的大小和 SHA256
- 转换耗时
- Word：页数、InlineShape、浮动 Shape、表格数量
- Excel：工作表、Shape、图片、图表、已保存打印区域统计
- PowerPoint：幻灯片、Shape、图片、图表、幻灯片尺寸
- PDF 页数启发式检测
- 转换失败时的错误信息

对 `.docx/.xlsx/.pptx` 还会直接读取 Office Open XML 压缩包内部的 `media` 目录，统计：

- 嵌入媒体数量
- 媒体未压缩总容量
- PNG/JPG/EMF/SVG 等扩展名分布

这对于判断“源文档明明包含大量高清图片，但生成 PDF 异常偏小”非常有帮助。

### 隐私设计

诊断 JSON 默认**不记录**：

- Windows 用户名
- 电脑名
- 源文档完整本地路径
- PDF 完整本地路径

只记录文件名和技术环境，更适合提交到 GitHub 后供 GPT / Codex 分析。

注意：文件名本身仍会记录，如果文件名敏感，请不要 Push 对应 JSON。

## GPT 诊断闭环

推荐工作流：

```text
Office 文件
    ↓
拖到 DragDrop_Office_to_PDF.cmd
    ↓
生成高质量 PDF
    ↓
自动生成 diagnostics/*.json
    ↓
GitHub Desktop
Commit to dev
    ↓
Push origin
    ↓
ChatGPT 读取真实运行结果
    ↓
继续分析 / 修改脚本
```

转换完成后可以直接在 ChatGPT 中说：

```text
看看 Office PDF diagnostics
```

## 为什么不用 PDF 打印机

虚拟打印机通常会把 Office 页面经过打印管线重新生成，有时会带来：

- 图片重采样
- 透明效果变化
- 字体替换
- 矢量对象退化
- 页面尺寸偏差
- Excel / PowerPoint 对象裁切

本工具优先使用 Office 自身的 PDF 引擎，通常能获得更高的版式一致性。

## 运行要求

必须满足：

1. Windows
2. 已安装 Microsoft Office 桌面版
3. Word / Excel / PowerPoint 能正常打开对应文件
4. Windows PowerShell 可用（Windows 10/11 默认具备）

不需要安装 Python、Node.js、Ghostscript 或其他 PDF 工具。

## 安全设计

Office 文档通过 COM 自动化打开时，脚本会尽量设置：

```text
AutomationSecurity = ForceDisable
```

以阻止文档中的 VBA 宏在自动转换过程中运行。

源文件以 ReadOnly 模式打开，转换后直接关闭，不执行保存。

## 一个必须说明的质量边界

“最高质量导出”只能保留 **源文件中目前仍然存在的质量**。

例如：

- 一张原本 6000×4000 的图片如果早已被 Word 压缩成 1200×800，PDF 无法恢复丢失的像素。
- 如果 Office 文件内部已经把图片压缩或裁剪并删除了编辑数据，转换脚本无法把被删除的原图重新找回来。
- 某些字体的许可证禁止嵌入 PDF 时，Office 可能使用位图方式保存文字外观；脚本启用了 `BitmapMissingFonts=True`，优先保证视觉一致性。

因此，这个脚本的目标是：

> **转换阶段不主动牺牲质量，并尽可能让 Office 使用其最高质量的原生 PDF 输出路径。**

## 文件说明

```text
Office_HQ_PDF_Converter/
├─ DragDrop_Office_to_PDF.cmd
├─ Convert-OfficeToPDF.ps1
├─ diagnostics/
│  └─ README.md
└─ README.md
```

## 批量转换

可以同时选择多个 Office 文件，一起拖到：

```text
DragDrop_Office_to_PDF.cmd
```

脚本会逐个转换，并在窗口最后显示：

```text
Success: N
Failed: N
```

## GitHub / 本地测试建议

这个工具必须在安装了 Microsoft Office 的真实 Windows 电脑上测试，普通 GitHub Actions runner 并没有完整 Microsoft Office，因此不适合用云端 CI 验证最终 PDF 视觉效果。

但现在本地测试后会自动生成 `diagnostics/*.json`，这些文件非常适合 Push 到 GitHub，让 GPT 基于真实 Office 环境继续诊断。
