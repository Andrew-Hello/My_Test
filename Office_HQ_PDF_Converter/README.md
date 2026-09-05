# Office HQ PDF Converter

面向 Windows + Microsoft Office 桌面版的高保真 Office -> PDF 转换工具。

支持：

- Word：`.docx`、`.doc`
- Excel：`.xlsx`、`.xls`
- PowerPoint：`.pptx`、`.ppt`

## 最简单的用法

把一个或多个 Office 文档直接拖到：

```text
DragDrop_Office_to_PDF.cmd
```

PDF 会生成在原文档同目录。

如果同名 PDF 已存在，不覆盖原文件，而是自动使用：

```text
Report_HQ.pdf
Report_HQ_2.pdf
...
```

## v2 的质量策略

从样例 `demo_Report.docx` 得到的实测结果表明，Microsoft Word 2024 经典 PDF 导出会把绝大多数图片压到约 200 PPI，而同一文档通过 Acrobat PDFMaker 高质量设置导出时，大部分图片可保留约 600-880 PPI。

因此 v2 已改为：

```text
Adobe Acrobat PDFMaker
        ↓ 失败/不存在
Word / Excel / PowerPoint 原生高质量接口
        ↓ 再失败
经典 Office PDF fallback
```

拖拽启动器现在默认调用：

```text
Convert-OfficeToPDF-v2.ps1
```

旧版 `Convert-OfficeToPDF.ps1` 暂时保留，方便回退和对照。

## Adobe PDFMaker 路径

脚本会在 Word / Excel / PowerPoint 中寻找：

```text
PDFMaker.OfficeAddin
```

找到后会读取 PDFMaker 当前的转换设置，并保留当前 `JobOptions`，然后只修改自动化所需参数，例如：

- 输出 PDF 文件名
- 不弹出保存路径窗口
- 不自动打开 PDF
- 静默/自动化模式
- 转换全部页面

因此，如果你已经在 Acrobat PDFMaker 中把 Adobe PDF 设置调成最高质量，脚本会优先复用这套高质量设置。

这条路径的目的，是尽量复现手工点击 Acrobat 功能区“创建 PDF”得到的效果。

## Office 原生 fallback

### Word

依次尝试：

```text
ExportAsFixedFormat3 + OptimizeForImageQuality=True
ExportAsFixedFormat2 + OptimizeForImageQuality=True
ExportAsFixedFormat classic
```

Microsoft 文档明确说明 `OptimizeForImageQuality=True` 会阻止图片下采样、保留更好的原始质量。但并不是所有 Word 安装版本都会暴露这些较新的 COM 方法，因此日志会记录每次失败的 HRESULT 和异常。

### Excel

使用 `Workbook.ExportAsFixedFormat` + `xlQualityStandard`。

### PowerPoint

使用 `ExportAsFixedFormat` + Print intent；失败时再退回 SaveAs PDF。

## 日志

每次运行都会生成文本日志：

```text
Office_HQ_PDF_Converter/logs/
```

日志包括：

- Office 版本与 Build
- 是否找到 PDFMaker
- PDFMaker ProgID / Description
- 当前 `JobOptions`
- PDFMaker 自动化参数
- `CreatePDFEx` 调用结果
- Word 新/旧导出接口的异常和 HRESULT
- 最终使用的 PDF 引擎
- 输出 PDF 大小
- 成功/失败统计

日志默认不记录完整本地路径，便于 Push 到 GitHub 后让 GPT 诊断。

## JSON diagnostics

转换完成后还会在：

```text
Office_HQ_PDF_Converter/diagnostics/
```

生成结构化 JSON，记录转换引擎、文件大小、Office 版本、PDFMaker 信息和对应日志文件名。

## PDF 质量分析器

仓库中另有：

```text
Analyze-PDF-Quality.py
```

用于比较不同 PDF 内部图片的：

- 像素尺寸
- 有效 PPI
- PDF 图片压缩 Filter
- 压缩后 image stream 大小
- 图片对象数量
- PDF Producer / Creator

样例自动比较报告：

```text
diagnostics/quality_comparison_latest.md
diagnostics/quality_comparison_latest.json
```

## 当前样例结论

`demo_Report.docx`：

- DOCX 约 5.26 MB
- 23 个媒体文件
- 约 5.20 MB 嵌入媒体

Word 原生 `demo_Report.pdf`：

- PDF 约 1.69 MB
- 图片有效 PPI 中位数约 199.7
- >=300 PPI 图片放置数量：0

Acrobat PDFMaker `demo_Report_HD.pdf`：

- PDF 约 7.36 MB
- 图片有效 PPI 中位数约 760.3
- 40 个图片放置对象中 38 个 >=600 PPI

因此当前项目的默认技术目标已经从“尽量调高 Office ExportAsFixedFormat”调整为：

> **优先自动调用 Acrobat PDFMaker，尽量复现 Acrobat 手工最高质量导出的结果；Office 原生接口仅作为兼容 fallback。**

## 运行要求

基础模式：

1. Windows
2. Microsoft Office 桌面版
3. Windows PowerShell 5.1+

最高质量推荐模式：

1. 上述条件
2. 安装 Adobe Acrobat（非仅 Reader）
3. Office 中可看到 Acrobat / PDFMaker COM Add-in
4. 在 Acrobat PDFMaker 首选项里选择你的高质量 Adobe PDF 设置

## 测试流程

Pull `dev` 后：

1. 把 `demo_Report.docx` 拖到 `DragDrop_Office_to_PDF.cmd`。
2. 检查新 PDF 文件大小和视觉质量。
3. 检查 `logs/*.log`。
4. 检查 `diagnostics/*_v2.json`。
5. Commit + Push 这些日志/诊断文件。
6. 告诉 GPT 查看 Office PDF logs。

下一轮开发会根据真实 PDFMaker 自动化日志继续修正，目标是让拖拽脚本生成的 PDF 尽可能接近 `demo_Report_HD.pdf`。
