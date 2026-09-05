# Office HQ PDF Converter

Windows + Microsoft Office 桌面版的高保真 Office → PDF 转换工具。

支持：

- Word：`.docx`、`.doc`
- Excel：`.xlsx`、`.xls`
- PowerPoint：`.pptx`、`.ppt`

## 使用方法

把一个或多个 Office 文档直接拖到：

```text
DragDrop_Office_to_PDF.cmd
```

成功结果会生成在原文档同目录，并命名为：

```text
Report_HQ.pdf
Report_HQ_2.pdf
...
```

如果高清恢复阶段失败，不会冒充 HQ 结果，而会明确生成：

```text
Report_NATIVE_ONLY.pdf
```

并在日志中标记 `NATIVE_ONLY`。

## 当前正式质量路线

```text
Office 原生排版引擎
        ↓
生成临时结构 PDF（temp/）
        ↓
src/HQImageRebuilder.py
        ↓
读取 OOXML 原始 raster media
        ↓
关系 / 裁剪 / PDF SMask 透明度语义匹配
        ↓
使用原始像素无损回填 PDF 图片 XObject
        ↓
*_HQ.pdf
```

Office 负责分页、字体、表格、图表和对象布局；我们自己的重建器负责恢复原始栅格图像质量。

正式流程不使用：

- Adobe Acrobat PDFMaker
- Word `ExportAsFixedFormat2`
- ImageHash

## 高清匹配引擎

当前匹配器保持严格的视觉阈值，不通过简单放宽 pHash 来追求更高的 replaced 数量。核心策略是先恢复文档和 PDF 的真实语义，再进行匹配。

当前引擎标识：

```text
builtin_phash_dct16+pdf_smask+pptx_srcRect+docx_srcRect
```

### PDF SMask 透明度重建

Office 常把透明 PNG 拆成 PDF 基础图像 XObject + Soft Mask。旧算法只读取基础图像，会让肉眼相同的透明 Logo、条带等得到很高的 pHash 差异。

现在匹配器会：

```text
Base XObject + SMask
        ↓
重新合成为 RGBA 图像
        ↓
白底视觉归一化
        ↓
pHash / 宽高比匹配
```

这项改进已经真实解决 DOCX 和 XLSX 中的透明图片误判。

### DOCX 结构感知

DOCX 会读取：

```text
word/document.xml
word/header*.xml
word/footer*.xml
word/footnotes.xml
word/endnotes.xml
word/comments*.xml
对应的 *_rels/*.rels
```

解析：

```text
pic:pic → a:blip → r:embed → word/media/*
```

如果图片存在 `a:srcRect`，会先在 DOCX 原始高清图片上复现同样裁剪，再参与匹配。因此 Word 中裁剪过的图片不再只能依赖模糊视觉猜测。

### PPTX 结构感知

PPTX 读取 slide relationships 和 DrawingML `a:srcRect`。原始高清图片先复现 PowerPoint 裁剪，再匹配 PDF 对象。

真实大型 PPTX 验证文件中：

- 65 个原始 raster media；
- 115 个 PDF 图片对象；
- 18 个 srcRect 裁剪候选；
- 16 个语义裁剪恢复；
- 56 个图片对象安全回填；
- replacement errors = 0。

## Excel 输出规则

Excel 固定执行：

```text
PaperSize = A4
Zoom = False
FitToPagesWide = 1
FitToPagesTall = False
```

也就是：

- 每个工作表强制使用 A4；
- 保留原有横向/纵向方向；
- 所有列在水平方向强制压缩到 1 个 A4 页面宽；
- 纵向不强制压成一页，长表允许自然分页；
- PDF 导出后再执行 OOXML 原图高清恢复。

## PowerPoint 输出规则

PowerPoint 直接使用稳定的 `SaveAs PDF` 结构导出：

- 保持源 SlideWidth / SlideHeight；
- 保持原始宽高比；
- 不强制 A4；
- 不拉伸；
- 之后执行 PPTX relationship/srcRect 高清恢复。

## 旧二进制格式

旧格式先在 `temp/` 中升级为现代 OOXML，再沿用同一套 HQ 流程：

```text
.doc → 临时 .docx → HQ
.xls → 临时 .xlsx → HQ
.ppt → 临时 .pptx → HQ
```

原文件从不修改。若 Office 无法完成旧格式升级，则明确降级为 `NATIVE_ONLY`。

此兼容代码已经实现，但仍需要更多真实 `.doc/.xls/.ppt` 文件做 Windows 实机验证。

## Python 依赖

高清恢复器只依赖：

```text
PyMuPDF
Pillow
```

脚本使用同一个 Python 解释器的 pip metadata 检查依赖。如果已经安装不会重复安装；缺失时先尝试当前 pip 源，再尝试官方 PyPI。

## 目录结构

```text
Office_HQ_PDF_Converter/
├─ DragDrop_Office_to_PDF.cmd
├─ Convert-OfficeToPDF.ps1
├─ src/
│  └─ HQImageRebuilder.py
├─ test/
│  ├─ Analyze-PDF-Quality.py
│  └─ samples/
│     ├─ demo_Report.docx
│     ├─ demo_Report_lowres.pdf
│     ├─ current/
│     └─ validation_20260905/
│        ├─ docx/
│        ├─ pptx/
│        └─ xlsx/
├─ temp/
├─ logs/
└─ diagnostics/
```

`temp/` 是转换中间目录并由 Git 忽略；真实测试文档统一放在 `test/samples/` 下，不再堆在正式项目根目录。

## 日志与诊断

每次运行会生成：

```text
logs/*_Office_HQ_PDF.log
diagnostics/*_conversion.json
diagnostics/*_image_rebuild.json
```

图像重建诊断会记录：

- 原始 media 数量；
- semantic candidate 数量；
- DOCX/PPTX `srcRect` 裁剪候选；
- PDF soft mask 重建数量；
- pHash / 宽高比 / 像素面积倍率；
- 实际替换数量；
- replacement errors。

## 当前验证基线

### DOCX 高清基线：demo_Report

新版语义匹配 + SMask 重建后的结果：

```text
源 raster media              22
DOCX relationship picture    22
DOCX srcRect crop candidate   1
PDF soft mask                 4
PDF raster image object      22
安全回填                     22 / 22
```

质量保持：

- Word 原生 PDF median PPI ≈ 199.7；
- 高清重建 PDF median PPI ≈ 760.3；
- `>=600 PPI` = 38/40；
- 与 Acrobat HD 样例的有效 PPI 分布一致。

旧算法曾只能恢复 21/22；新的 DOCX relationship + PDF SMask 引擎达到 22/22，并且不需要降低 pHash 安全阈值。

### DOCX 真实验证：Test_docx (2)

该文件在旧算法中只有 1/3 图片能够安全回填。新版发现 PDF 中有 2 个 Soft Mask，将透明度重新合成后得到：

```text
DOCX semantic picture    3
PDF soft mask             2
PDF raster image          3
安全回填                  3 / 3
replacement errors        0
```

因此这次提升是修复 PDF 透明图像解释方式，而不是放宽匹配条件。

### XLSX 真实验证：Test_xlsx (1)

旧算法中该图片虽然源文件和 PDF 都是 `518×118`，但 pHash 差异很高，因此保守地 `replaced=0`。

新版识别到 1 个 PDF Soft Mask 后：

```text
源 raster media       1
PDF soft mask          1
PDF raster image       1
安全回填               1 / 1
replacement errors     0
```

说明 Excel 透明 PNG 同样受益于 SMask 重建。

### PPTX 验证基线

当前小型 PPTX 回归样例：

```text
原始 raster media       14
semantic candidates     16
srcRect crop candidates  6
semantic crop replaced   5
PDF soft masks            2
PDF raster image         15
安全回填                 14 / 15
replacement errors        0
```

大型实机验证 PPTX 则达到 56 个安全回填，其中 16 个为 srcRect 语义裁剪恢复。

### 2026-09-05 多文档实机验证

真实批量测试：

```text
3 × DOCX
4 × XLSX
1 × PPTX
合计 8 份
```

同批运行结果：

```text
Success=8
FailedOrDegraded=0
```

## CI

GitHub Actions 当前强制守住以下质量门槛：

- PowerShell / Python 语法通过；
- DOCX `demo_Report` 必须 22/22 唯一 raster 图片全部恢复；
- DOCX median PPI 不得低于 600，且 `>=600 PPI` 放置数量不得低于 38；
- DOCX `Test_docx (2)` 必须维持 3/3 回填并重建至少 2 个 PDF Soft Mask；
- PPTX 小型基线不得低于 14 个安全回填，并必须实际执行 srcRect 语义恢复；
- XLSX `Test_xlsx (1)` 必须维持 1 个 Soft Mask + 1/1 图片安全回填；
- 所有上述回归必须保持 replacement errors = 0。

## 通用本机环境快照

仓库根目录另有：

```text
Local_Environment_Collector/
```

出现本机环境相关问题时，双击：

```text
Local_Environment_Collector/Collect_Local_Environment.cmd
```

然后把 `reports/` 生成的 JSON/TXT Commit + Push 即可继续诊断。
