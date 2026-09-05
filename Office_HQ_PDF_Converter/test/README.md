# Test assets

这个目录只存放测试脚本与测试样例，不参与正式拖拽转换入口。

结构约定：

```text
test/
├─ Analyze-PDF-Quality.py
└─ samples/
```

- `Analyze-PDF-Quality.py`：用于检查 PDF 内图片的像素尺寸、有效 PPI、压缩 Filter 与图片流大小。
- `samples/`：测试 DOCX/PDF 等样例文件。

正式运行代码位于项目根目录与 `src/`；临时转换中间文件统一写入 `temp/`。
