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

样例测试已经证明，Word 2024 经典 PDF 导出会把大量嵌入图片降到约 200 PPI；而我们的高清恢复方案可以把同一样例恢复到与 Acrobat 高质量 PDF 相同的有效 PPI 分布。

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

Office 负责最擅长的分页、字体、表格、图表和对象布局；我们自己的重建器负责夺回图片质量。

正式流程**不再使用**：

- Adobe Acrobat PDFMaker
- Word `ExportAsFixedFormat2`
- ImageHash

这些路线已经通过实机测试证明不够稳定或没有必要。

## Python 依赖

高清恢复器现在只依赖：

```text
PyMuPDF
Pillow
```

脚本首先检查当前 Python 是否能：

```python
import pymupdf
import PIL
```

如果已经安装，不会运行 pip。

如果缺失，会先尝试当前 pip 镜像；若镜像缺包，再尝试一次官方 PyPI。

`ImageHash` 已完全移除，所以即使当前镜像没有 ImageHash，也不会影响正式转换。

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

- Office 版本与 Build
- 结构 PDF 是否成功
- Python runner
- PyMuPDF/Pillow 检查状态
- HQ 重建是否成功
- 最终结果是 `HQ_REBUILT` / `NATIVE_ONLY` / `FAILED`
- 文件大小与耗时

## 通用本机环境快照

仓库根目录另有：

```text
Local_Environment_Collector/
```

当转换器或其他项目出现本机环境相关问题时，双击：

```text
Local_Environment_Collector/Collect_Local_Environment.cmd
```

然后把 `Local_Environment_Collector/reports/` 生成的 JSON/TXT Commit + Push，ChatGPT 就可以基于真实环境继续适配，而不需要反复手工询问版本、路径和依赖。

## 当前样例基线

历史样例 `demo_Report.docx` 的测试结果：

- Word 原生 PDF：有效 PPI 中位数约 **199.7**，`>=600 PPI` 为 **0/40**
- 高清重建 PDF：有效 PPI 中位数约 **760.3**，`>=600 PPI` 为 **38/40**
- Acrobat HD 基准：有效 PPI 中位数约 **760.3**，`>=600 PPI` 为 **38/40**

因此当前生产路线的图片分辨率已经达到样例 Acrobat HD 的水平。

## 旧二进制格式

`.doc/.xls/.ppt` 目前仍然可以转换，但因为它们不是 OOXML ZIP 包，当前版本无法直接从 `media/` 恢复原始图片，所以结果会标记为 `NATIVE_ONLY`。后续可单独开发“旧格式 → 临时 OOXML → HQ 恢复”的兼容阶段。
