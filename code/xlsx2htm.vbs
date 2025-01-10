Option Explicit

' Constants
Const c_sHtmlFormat = 44

' Integer variables
Dim iSheetIndex, iRow, iCol

' Object variables
Dim oExcel, oFSO, oOutputFile, oSheet, oWorkbook, oCell

' String variables
Dim sCombinedHtml, sExcelFile, sOutputHtmlFile, sValue, sSheetName, sTableContent

' Check arguments
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript excelToHtml.vbs <ExcelFileName>"
    WScript.Quit 1
End If

' Get input file name
sExcelFile = WScript.Arguments(0)

' Validate file
Set oFSO = CreateObject("Scripting.FileSystemObject")
If Not oFSO.FileExists(sExcelFile) Then
    WScript.Echo "Error: File not found: " & sExcelFile
    WScript.Quit 1
End If

' Generate output file name
sOutputHtmlFile = oFSO.GetParentFolderName(sExcelFile) & "\" & oFSO.GetBaseName(sExcelFile) & ".htm"

' Initialize Excel application
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
oExcel.DisplayAlerts = False

' Open workbook
wscript.echo "Opening " & oFso.getBaseName(sExcelFile) & ".xlsx"
Set oWorkbook = oExcel.Workbooks.Open(sExcelFile)

' Initialize combined HTML content with proper structure
sCombinedHtml = "<!DOCTYPE html>" & vbCrLf & _
                "<html>" & vbCrLf & _
                "<head><title>" & oFSO.GetBaseName(sExcelFile) & "</title></head>" & vbCrLf & _
                "<body>" & vbCrLf

' Process each sheet
For iSheetIndex = 1 To oWorkbook.Sheets.Count
    Set oSheet = oWorkbook.Sheets(iSheetIndex)
    If oSheet.Name = "" Then
        sSheetName = "Sheet" & iSheetIndex
    Else
        sSheetName = oSheet.Name
    End If
wscript.echo "Sheet " & sSheetName

    ' Initialize table content
    sTableContent = "<table border='1'>" & vbCrLf

    ' Iterate through UsedRange
    For iRow = 1 To oSheet.UsedRange.Rows.Count
if iRow = 1 then sTableContent = sTableContent & "<thead>" & vbCrLf
if iRow = 2 then sTableContent = sTableContent & "<tbody>" & vbCrLf
sTableContent = sTableContent & "<tr>" & vbCrLf
        For iCol = 1 To oSheet.UsedRange.Columns.Count
            Set oCell = oSheet.Cells(iRow, iCol)
sValue = ""
on error resume next
sValue = cStr(oCell.Value)
on error goto 0
if iRow = 1 then
            sTableContent = sTableContent & "<th>" & sValue & "</th>" & vbCrLf
else
sTableContent = sTableContent & "<td>" & sValue & "</td>" & vbCrLf
end if
        Next
         sTableContent = sTableContent & "</tr>" & vbCrLf

if iRow = 1 then sTableContent = sTableContent & "</thead>" & vbCrLf
if iRow = oSheet.UsedRange.Rows.Count then sTableContent = sTableContent & "</tbody>" & vbCrLf
next
    sTableContent = sTableContent & "</table>" & vbCrLf

    ' Append sheet content with heading
    sCombinedHtml = sCombinedHtml & "<h2>" & sSheetName & "</h2>" & vbCrLf & sTableContent
Next

' Close HTML structure
sCombinedHtml = sCombinedHtml & "</body>" & vbCrLf & "</html>"

wscript.echo "Saving " & oFso.getBasename(sOutputHtmlFile) & ".htm"
' Write combined HTML to output file
Set oOutputFile = oFSO.CreateTextFile(sOutputHtmlFile, True)
oOutputFile.Write sCombinedHtml
oOutputFile.Close

' Cleanup
oWorkbook.Close False
oExcel.Quit

Set oExcel = Nothing
Set oFSO = Nothing
Set oWorkbook = Nothing

