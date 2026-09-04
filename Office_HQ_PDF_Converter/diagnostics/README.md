# Diagnostics

此目录用于保存 `Office_HQ_PDF_Converter` 在真实 Windows + Microsoft Office 环境中运行后生成的诊断 JSON。

运行 `DragDrop_Office_to_PDF.cmd` 后，脚本会自动在这里生成：

```text
YYYYMMDD_HHMMSS_mmm_文档名.json
YYYYMMDD_HHMMSS_mmm_batch_summary.json
```

诊断数据默认不会记录：

- Windows 用户名
- 电脑名
- 源文档完整本地路径
- PDF 完整本地路径

因此更适合提交到 GitHub 后供 GPT / Codex 分析。

典型字段包括：

- Windows / PowerShell / .NET 环境
- Office 应用与版本
- 源文件大小、SHA256
- PDF 大小、SHA256
- 转换耗时
- Word 页数、图片/形状/表格数
- Excel 工作表、形状、图片、图表、打印区域统计
- PowerPoint 幻灯片、形状、图片、图表、页面尺寸
- DOCX/XLSX/PPTX 内部嵌入媒体数量与总容量
- 使用的实际 Office PDF 导出接口
- PDF 页数启发式检测结果
- 转换错误信息（若失败）

## 提交给 GPT 分析

转换完成后，如果 GitHub Desktop 的 Changes 中出现本目录下的新 JSON：

1. Commit to `dev`
2. Push origin
3. 在 ChatGPT 中告诉我：`看看 Office PDF diagnostics`

即可直接基于真实本机转换结果继续诊断。

注意：诊断 JSON 会包含源文件的**文件名**。如果文件名本身敏感，请在 Push 前重命名源文件或不要上传对应诊断记录。
