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

对于现代 OOXML 格式（DOCX/XLSX/PPTX），成功结果会生成在原文档同目录，并命名为：

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

正式流程固定为：

```text
Office 原生排版引擎
        ↓
生成临时结构 PDF（temp/）
        ↓
src/HQImageRebuilder.py
        ↓
从 DOCX/XLSX/PPTX 的 OOXML media 中读取原始图片
        ↓
匹配 PDF 中被压缩的图片对象
        ↓
使用原始像素无损回填
        ↓
*_HQ.pdf
```

Office 负责分页、字体、表格、图表和对象布局；我们自己的重建器负责恢复原始栅格图像质量。

正式流程不再使用：

- Adobe Acrobat PDFMaker
- Word `ExportAsFixedFormat2`
- ImageHash

## Excel 输出规则

Excel 现在有固定的 PDF 页面策略，避免不同工作簿自身的打印缩放设置造成列被裁切或横向分页：

```text
PaperSize = A4
Zoom = False
FitToPagesWide = 1
FitToPagesTall = False
```

也就是：

- 每个工作表强制使用 A4 纸张；
- 保留工作表原来的横向/纵向方向；
- 所有列在水平方向强制压缩到 **1 个 A4 页面宽**；
- 纵向不强制压缩为一页，内容较长时允许自然分页；
- PDF 导出后再执行 OOXML 原图高清恢复。

## PowerPoint 输出规则

PowerPoint 使用自身的 `SaveAs PDF` 作为稳定结构导出路径，不再先调用当前环境中会报 COM warning 的 `ExportAsFixedFormat`。

PDF 页面保持源 PPT/PPTX 的幻灯片画布尺寸，因此：

- 不强制改成 A4；
- 不拉伸；
- 不改变宽高比；
- 之后再执行 OOXML 原图高清恢复。

日志会记录幻灯片宽度、高度和宽高比。

## Python 依赖

高清恢复器只依赖：

```text
PyMuPDF
Pillow
```

脚本使用同一个 Python 解释器的 pip metadata 检查依赖。如果已经安装，不会再次安装；如果缺失，会先尝试当前 pip 镜像，再尝试一次官方 PyPI。

`ImageHash` 已完全移除。

## 目录结构

```text
Office_HQ_PDF_Converter/
├─ DragDrop_Office_to_PDF.cmd      # 用户入口
├─ Convert-OfficeToPDF.ps1         # 唯一正式主脚本
├─ src/
│  └─ HQImageRebuilder.py          # 正式高清图片恢复核心
├─ test/
│  ├─ Analyze-PDF-Quality.py       # PDF 质量分析测试工具
│  └─ samples/                     # 测试文档/基准 PDF
├─ temp/                           # 转换中间文件；Git 忽略
├─ logs/                           # 可读运行日志
└─ diagnostics/                    # JSON 诊断结果
```

## 日志与诊断

每次运行会生成：

```text
logs/*_Office_HQ_PDF.log
diagnostics/*_conversion.json
diagnostics/*_image_rebuild.json
```

日志会记录 Office 版本、格式专用布局策略、Python 环境、HQ 重建统计、最终结果、文件大小与耗时。

## 当前样例基线

DOCX 样例：

- Word 原生 PDF：有效 PPI 中位数约 **199.7**，`>=600 PPI` 为 **0/40**
- 高清重建 PDF：有效 PPI 中位数约 **760.3**，`>=600 PPI` 为 **38/40**
- Acrobat HD 基准：有效 PPI 中位数约 **760.3**，`>=600 PPI` 为 **38/40**

2026-09-05 的三格式实机测试：

- `demo_docx.docx` → `demo_docx_HQ.pdf`：约 **5.228 MB**，HQ_REBUILT
- `demo_pptx.pptx` → `demo_pptx_HQ.pdf`：约 **2.016 MB**，HQ_REBUILT
- `demo_xlsx.xlsx` → `demo_xlsx_HQ.pdf`：约 **0.112 MB**，HQ_REBUILT

三者同批运行结果：`Success=3; FailedOrDegraded=0`。

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

## 旧二进制格式

`.doc/.xls/.ppt` 可以转换，但因为它们不是 OOXML ZIP 包，当前版本无法直接从 `media/` 恢复原始图片，因此结果会标记为 `NATIVE_ONLY`。后续可以单独开发“旧格式 → 临时 OOXML → HQ 恢复”的兼容阶段。
