# Office HQ PDF Converter

面向 Windows + Microsoft Office 桌面版的高保真 Office -> PDF 转换工具。

支持：

- Word：`.docx`、`.doc`
- Excel：`.xlsx`、`.xls`
- PowerPoint：`.pptx`、`.ppt`

## 当前正式方案：v3

默认拖拽入口：

```text
DragDrop_Office_to_PDF.cmd
```

它现在调用：

```text
Convert-OfficeToPDF-v3.ps1
```

v3 不再使用 Acrobat PDFMaker，也不再使用会在部分 Word 版本中阻塞的 `ExportAsFixedFormat2`。

稳定路径是：

```text
Office 原生排版/分页
        ↓
生成临时结构 PDF
        ↓
读取 DOCX/XLSX/PPTX 内部原始 media
        ↓
与 PDF 中被压缩的图片对象自动匹配
        ↓
用原始高分辨率像素无损回填 PDF image stream
        ↓
最终 HQ PDF
```

Office 在这里主要负责它最擅长的版式、分页、字体、表格和矢量内容；最终图片质量由我们自己的重建器接管。

## 样例验证结果

`demo_Report.docx` 的实测对比：

| 指标 | Word 原生 PDF | v3 重建原型 | Acrobat HD |
|---|---:|---:|---:|
| PDF 大小 | 1.693 MB | 5.228 MB | 7.356 MB |
| 页面数 | 19 | 19 | 19 |
| 图片放置对象 | 40 | 40 | 40 |
| PPI 中位数 | 199.7 | 760.3 | 760.3 |
| P90 PPI | 199.8 | 882.6 | 882.6 |
| >=300 PPI | 0/40 | 38/40 | 38/40 |
| >=600 PPI | 0/40 | 38/40 | 38/40 |
| PDF 图片编码 | Flate + JPEG | 全部 Flate | 全部 Flate |

因此 `demo_Report_Rebuilt.pdf` 已经在有效图片分辨率分布上达到 Acrobat HD 样例水平。

## 输出状态

v3 日志会明确标记结果，避免“文件名叫 HQ、实际却走低清 fallback”的情况。

### HQ_REBUILT

表示现代 OOXML 文件（DOCX/XLSX/PPTX）已经完成原始图片回填：

```text
HQ_REBUILT
```

这是当前认可的高质量结果。

### NATIVE_ONLY

表示只能得到 Office 原生 PDF：

```text
NATIVE_ONLY
```

这类 PDF 可能仍只有约 200 PPI，因此文件名会明确带：

```text
_NATIVE_ONLY.pdf
```

目前 `.doc/.xls/.ppt` 旧二进制格式先按这一模式处理。

### FAILED

表示转换失败，查看 `logs/`。

## Python 高清重建器

核心脚本：

```text
Rebuild-HighRes-PDF-Images.py
```

支持直接读取 OOXML 包中的：

```text
DOCX -> word/media/
XLSX -> xl/media/
PPTX -> ppt/media/
```

需要 Python 3 以及：

```text
pymupdf
pillow
ImageHash
```

v3 会自动检测。如果 Python 已存在但模块缺失，会尝试一次：

```text
python -m pip install --user --upgrade pymupdf pillow ImageHash
```

如果电脑没有 Python，Office 原生结构 PDF 仍可生成，但 v3 会明确标记为 `NATIVE_ONLY`，不会伪装成 HQ 结果。

后续计划是把这个 Python 重建器打包成独立 EXE，使最终用户不需要单独安装 Python。

## FixedFormat2 为什么停用

在当前实测机器：

```text
Word Version = 16.0
Build = 16.0.17932
```

`Document.ExportAsFixedFormat2(... OptimizeForImageQuality=True)` 通过 PowerShell COM 调用后会出现无限阻塞，因此不再用于正式转换流程。

`Test_Word_FixedFormat2_HQ.cmd` 已标记为 retired test，不应再用于日常转换。

## 日志与诊断

每次运行生成：

```text
Office_HQ_PDF_Converter/logs/
Office_HQ_PDF_Converter/diagnostics/
```

日志会记录：

- Office 版本和 Build
- 原生结构 PDF 大小
- Python 路径
- HQ 依赖是否可用
- 图片重建是否成功
- `HQ_REBUILT / NATIVE_ONLY / FAILED`
- 最终 PDF 大小
- 转换耗时

图片匹配诊断 JSON 还会记录：

- OOXML 原始媒体数量
- PDF 唯一图片对象数量
- 每个 source image 与 PDF xref 的对应关系
- 原始/低清像素尺寸
- 像素面积倍率
- 匹配模式
- 实际替换数量
- 替换错误

## 使用

1. GitHub Desktop 切换到 `dev`。
2. Fetch / Pull。
3. 将 `.docx/.xlsx/.pptx/.doc/.xls/.ppt` 拖到：

```text
DragDrop_Office_to_PDF.cmd
```

4. 查看终端最后是否出现：

```text
HQ_REBUILT
```

5. 检查生成 PDF。
6. 如需诊断，把 `logs/` 与 `diagnostics/` 的新文件 Commit + Push 到 `dev`。

## 关于“不依赖 Office / Acrobat”

当前 v3 已经不依赖 Acrobat。

仍使用 Office 作为版式引擎，因为准确复刻 Word/Excel/PowerPoint 的分页、字体度量、表格、浮动对象和图表布局是整个问题中最复杂的部分。

现代 `.docx/.xlsx/.pptx` 本身是 OOXML，可以完全独立解析。长期可以进一步开发“不需要安装 Office”的独立渲染器，但这属于第二阶段；当前优先目标是先获得稳定、可批量使用、图片质量达到 Acrobat HD 水平的实际工具。
