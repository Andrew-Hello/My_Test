# Conversion logs

`DragDrop_Office_to_PDF.cmd` now launches `Convert-OfficeToPDF-v2.ps1`.

Every run creates a timestamped text log in this folder, for example:

```text
20260905_112233_456_Office_HQ_PDF.log
```

The log records:

- detected Office version/build
- whether `PDFMaker.OfficeAddin` was found and connected
- PDFMaker ProgID / description
- current PDFMaker `JobOptions`
- PDFMaker settings that were applied
- `CreatePDFEx` call outcome
- Word `ExportAsFixedFormat3` / `ExportAsFixedFormat2` / classic fallback errors including HRESULT
- final conversion engine and output size
- total success/failure count

The log intentionally records file names rather than full local paths so it can be pushed to GitHub for diagnosis with less privacy risk.

After a local test, commit and push the new `.log` file together with the corresponding diagnostics JSON so GPT can inspect the exact conversion path.
