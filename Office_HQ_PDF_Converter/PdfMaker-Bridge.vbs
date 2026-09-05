' Acrobat PDFMaker automation bridge for Word documents.
' VBScript is used because COM ByRef Variant marshaling is closer to VBA than PowerShell.
Option Explicit

Dim sourcePath, outputPath
If WScript.Arguments.Count < 2 Then
    WScript.Echo "ERROR|Usage: PdfMaker-Bridge.vbs <source> <output>"
    WScript.Quit 2
End If

sourcePath = WScript.Arguments(0)
outputPath = WScript.Arguments(1)

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(sourcePath) Then
    WScript.Echo "ERROR|Source file not found"
    WScript.Quit 3
End If
If fso.FileExists(outputPath) Then
    On Error Resume Next
    fso.DeleteFile outputPath, True
    On Error GoTo 0
End If

Dim word, doc, addin, pmkr, stng, retVal
On Error Resume Next
Set word = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR|Create Word failed|" & Hex(Err.Number) & "|" & Err.Description
    WScript.Quit 10
End If
word.Visible = True
word.DisplayAlerts = 0
word.WindowState = 2 ' wdWindowStateMinimize
Err.Clear

Set doc = word.Documents.Open(sourcePath, False, True, False)
If Err.Number <> 0 Then
    WScript.Echo "ERROR|Open document failed|" & Hex(Err.Number) & "|" & Err.Description
    word.Quit
    WScript.Quit 11
End If
doc.Activate
WScript.Sleep 1200

Set addin = word.COMAddIns.Item("PDFMaker.OfficeAddin")
If Err.Number <> 0 Or addin Is Nothing Then
    WScript.Echo "ERROR|PDFMaker.OfficeAddin not found|" & Hex(Err.Number) & "|" & Err.Description
    doc.Close False
    word.Quit
    WScript.Quit 12
End If

Err.Clear
addin.Connect = True
WScript.Sleep 1500
Set pmkr = addin.Object
If Err.Number <> 0 Or pmkr Is Nothing Then
    WScript.Echo "ERROR|PDFMaker automation object unavailable|" & Hex(Err.Number) & "|" & Err.Description
    doc.Close False
    word.Quit
    WScript.Quit 13
End If

' Important: stng is an untyped Variant/Object receiving a COM ByRef result, like late-bound VBA.
Err.Clear
pmkr.GetCurrentConversionSettings stng
If Err.Number <> 0 Or Not IsObject(stng) Then
    WScript.Echo "WARN|GetCurrentConversionSettings failed|" & Hex(Err.Number) & "|" & Err.Description
    Err.Clear
    pmkr.GetDefaultConversionSettings stng
End If
If Err.Number <> 0 Or Not IsObject(stng) Then
    WScript.Echo "ERROR|GetDefaultConversionSettings failed|" & Hex(Err.Number) & "|" & Err.Description
    doc.Close False
    word.Quit
    WScript.Quit 14
End If

WScript.Echo "INFO|PDFMaker settings object acquired"
Err.Clear
stng.OutputPDFFileName = outputPath
stng.PromptForPDFFilename = False
stng.ShouldShowProgressDialog = False
stng.ViewPDFFile = False
stng.ConvertAllPages = True
stng.AddLinks = True
stng.AddBookmarks = True
If Err.Number <> 0 Then
    WScript.Echo "WARN|Some PDFMaker settings were not accepted|" & Hex(Err.Number) & "|" & Err.Description
    Err.Clear
End If

retVal = 0
pmkr.CreatePDFEx stng, retVal
If Err.Number <> 0 Then
    WScript.Echo "ERROR|CreatePDFEx failed|" & Hex(Err.Number) & "|" & Err.Description
    doc.Close False
    word.Quit
    WScript.Quit 15
End If
WScript.Echo "INFO|CreatePDFEx returned|retval=" & CStr(retVal)

Dim i
For i = 1 To 300
    If fso.FileExists(outputPath) Then
        If fso.GetFile(outputPath).Size > 0 Then Exit For
    End If
    WScript.Sleep 400
Next

If Not fso.FileExists(outputPath) Then
    WScript.Echo "ERROR|PDF file did not appear after CreatePDFEx"
    doc.Close False
    word.Quit
    WScript.Quit 16
End If

WScript.Echo "OK|PDF created|bytes=" & CStr(fso.GetFile(outputPath).Size)
doc.Close False
word.Quit
WScript.Quit 0
