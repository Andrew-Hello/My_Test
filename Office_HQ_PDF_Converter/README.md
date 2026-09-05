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

成功结果生成在原文档同目录：

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

现代 OOXML 文件的正式流程：

```text
Office 原生排版引擎
        ↓
生成临时结构 PDF（temp/）
        ↓
src/HQImageRebuilder.py
        ↓
读取 DOCX/XLSX/PPTX 内部原始栅格媒体
        ↓
结合 PDF 图像对象进行安全匹配
        ↓
原始像素无损回填
        ↓
*_HQ.pdf
```

Office 负责分页、字体、表格、图表和对象布局；我们自己的重建器负责恢复原始栅格图像质量。

正式流程不再使用：

- Adobe Acrobat PDFMaker
- Word `ExportAsFixedFormat2`
- ImageHash

## Excel 输出规则

Excel 使用固定的 PDF 页面策略，避免工作簿原有打印设置造成列被裁切或横向分页：

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

PowerPoint 使用自身的 `SaveAs PDF` 作为稳定结构导出路径。

PDF 页面保持源 PPT/PPTX 的幻灯片画布尺寸，因此：

- 不强制改成 A4；
- 不拉伸；
- 不改变宽高比；
- 日志记录幻灯片宽度、高度和宽高比。

### PPTX 裁剪图片高清恢复

对于 PPTX，高清恢复器不再只比较 `ppt/media/` 原图和 PDF 图片。

新版会额外解析：

```text
ppt/slides/slide*.xml
ppt/slides/_rels/slide*.xml.rels
```

识别每张图片的：

- `r:embed`：实际引用的 `ppt/media/image*.png`；
- `a:srcRect`：PowerPoint 对图片执行的左/上/右/下裁剪比例。

然后先在原始高清图片上复现同样的裁剪，再进行 pHash + 宽高比匹配。只有语义裁剪和视觉内容同时吻合才回填，因此无需通过放宽阈值来冒险替换。

当前 PPTX 样例实测：

- 原始栅格媒体：14 个；
- PDF 唯一栅格对象：15 个；
- 识别到 `srcRect` 裁剪候选：6 个；
- 其中安全语义裁剪回填：5 个；
- 总回填对象：**由上一版 9/15 提升到 14/15**；
- replacement errors：0。

## 旧二进制格式升级

`.doc/.xls/.ppt` 不再默认停留在 `NATIVE_ONLY`。

新版会先在 `temp/` 中建立一次临时现代 OOXML 文件：

```text
.doc → .docx   Word SaveAs2 / wdFormatXMLDocument = 12
.xls → .xlsx   Excel SaveAs / xlOpenXMLWorkbook = 51
.ppt → .pptx   PowerPoint SaveAs / ppSaveAsOpenXMLPresentation = 24
```

随后按现代格式继续执行：

```text
临时 OOXML
   ↓
Office 结构 PDF
   ↓
OOXML 原始媒体高清恢复
   ↓
*_HQ.pdf
```

原始 `.doc/.xls/.ppt` 文件不会被修改；临时 OOXML 用完后立即删除。

成功时日志/诊断会标记：

```text
HQ_REBUILT_FROM_LEGACY
```

如果旧文件无法可靠转换成临时 OOXML，才退回 Office 原生 PDF，并明确标记 `NATIVE_ONLY`。

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

日志会记录：

- Office 版本；
- Excel A4 / 1 页宽策略；
- PowerPoint 画布尺寸和宽高比；
- PPTX `srcRect` 语义裁剪恢复统计；
- 老格式临时 OOXML 升级是否成功；
- Python 环境；
- HQ 重建统计；
- 最终结果、文件大小和耗时。

## 当前样例基线

DOCX 样例：

- Word 原生 PDF：有效 PPI 中位数约 **199.7**，`>=600 PPI` 为 **0/40**；
- 高清重建 PDF：有效 PPI 中位数约 **760.3**，`>=600 PPI` 为 **38/40**；
- Acrobat HD 基准：有效 PPI 中位数约 **760.3**，`>=600 PPI` 为 **38/40**。

2026-09-05 三格式实机测试：

- `demo_docx.docx` → `demo_docx_HQ.pdf`：约 **5.228 MB**，HQ_REBUILT；
- `demo_pptx.pptx` → `demo_pptx_HQ.pdf`：约 **2.016 MB**，HQ_REBUILT；
- `demo_xlsx.xlsx` → `demo_xlsx_HQ.pdf`：约 **0.112 MB**，HQ_REBUILT；
- 同批运行：`Success=3; FailedOrDegraded=0`。

自动回归测试还会验证：

- DOCX 恢复对象不少于 20；
- DOCX 重建 PDF median PPI 不低于 600；
- PPTX 能解析真实 `srcRect`；
- PPTX 总回填不少于既有基线；
- 不允许出现 replacement errors。

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
