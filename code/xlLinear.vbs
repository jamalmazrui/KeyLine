Option Explicit

Const cDelimiter = " = "
Const cExtensionInix = ".inix"
Const cMultilineDelimiter = "`"

Dim bFileExists
Dim iCol, iRow
Dim oCell, oExcelApp, oSheet, oUsedRange, oWorkbook
Dim sCellValue, sFileContent, sFilePath, sOutputFile

' Validate input parameters
If WScript.Arguments.Count <> 1 Then
WScript.Echo "Usage: cscript xlLinear.vbs <path-to-xlsx-file>"
WScript.Quit 1
End If

sFilePath = Trim(WScript.Arguments(0))
If LCase(Right(sFilePath, 5)) <> ".xlsx" Then
WScript.Echo "Error: Input file must have .xlsx extension."
WScript.Quit 1
End If

sOutputFile = Replace(sFilePath, ".xlsx", cExtensionInix)

' Open Excel application
Set oExcelApp = CreateObject("Excel.Application")
On Error Resume Next
Set oWorkbook = oExcelApp.Workbooks.Open(sFilePath)
If Err.Number <> 0 Then
WScript.Echo "Error: Unable to open Excel file."
oExcelApp.Quit
WScript.Quit 1
End If
On Error GoTo 0

oExcelApp.DisplayAlerts = False
sFileContent = ""

' Iterate through sheets
For Each oSheet In oWorkbook.Sheets
sFileContent = sFileContent & "[" & Trim(oSheet.Name) & "]" & vbCrLf

Set oUsedRange = oSheet.UsedRange
For iRow = 1 To oUsedRange.Rows.Count
For iCol = 1 To oUsedRange.Columns.Count
Set oCell = oUsedRange.Cells(iRow, iCol)
sCellValue = Trim(CStr(oCell.Value))
If sCellValue <> "" Then
sCellValue = RemoveBrackets(sCellValue)
sFileContent = sFileContent & LCase(oCell.Address(False, False)) & cDelimiter & FormatValue(sCellValue) & vbCrLf
End If
Next
Next
sFileContent = sFileContent & vbCrLf
Next

' Save to file as UTF-8 with BOM
If FileSave(sOutputFile, sFileContent) Then
WScript.Echo "Conversion complete: " & sOutputFile
Else
WScript.Echo "Error: Unable to save file."
End If

' Cleanup
oWorkbook.Close False
oExcelApp.Quit
WScript.Quit 0

' Function to format cell values for .inix file
Function FormatValue(sValue)
Dim sFormatted, aLines, i
If InStr(sValue, vbCrLf) > 0 Or InStr(sValue, vbLf) > 0 Then
aLines = Split(sValue, vbLf)
For i = LBound(aLines) To UBound(aLines)
aLines(i) = RemoveBrackets(Trim(aLines(i)))
Next
sFormatted = cMultilineDelimiter & vbLf & Join(aLines, vbLf) & vbLf & cMultilineDelimiter
Else
sFormatted = RemoveBrackets(sValue)
End If
FormatValue = sFormatted
End Function

' Function to remove brackets from text
Function RemoveBrackets(sValue)
RemoveBrackets = Replace(Replace(sValue, "[", ""), "]", "")
End Function

' Function to save a string to a file as UTF-8 with BOM
Function FileSave(sPath, sContent)
Dim oStream, sBom
sBom = Chr(239) & Chr(187) & Chr(191) ' UTF-8 BOM
On Error Resume Next
Set oStream = CreateObject("ADODB.Stream")
oStream.Type = 2 ' Text
oStream.Charset = "utf-8"
oStream.Open
oStream.WriteText sBom & sContent
oStream.SaveToFile sPath, 2 ' Overwrite
oStream.Close
Set oStream = Nothing
FileSave = (Err.Number = 0)
On Error GoTo 0
End Function
