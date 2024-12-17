Option Explicit

Dim iArgCount, iNewCol, iCol, iColCount, iRow, iRowCount, iWidth, iMaxWidth
Dim oRange, oApp, oBooks, oBook, oSheet, oColumn, oCell, oShell, oSystem
Dim sSourceFile, sTargetFile, sHeader

Const msoAutomationSecurityLow = 1
Const xlWorkbookDefault = 51
Const xlCSV = 6

' xl vertical alignment enumeration
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

Set oShell =CreateObject("Wscript.Shell")
Set oSystem =CreateObject("Scripting.FileSystemObject")

iArgCount = WScript.Arguments.Count
If iArgCount < 1 Then
WScript.Echo "Specify a source file as the single parameter."
WScript.Echo "Or specify a source file as the first parameter and a target .xlsx file as the second parameter."
WScript.Quit
End If

sSourceFile = WScript.Arguments(0)
If InStr(sSourceFile, "\") = 0 Then sSourceFile = oSystem.BuildPath(oShell.CurrentDirectory, sSourceFile)
' If InStr(sSourceFile, "\") = 0 Then sSourceFile = oShell.CurrentDirectory & "\" & sSourceFile
If not oSystem.FileExists(sSourceFile) Then
WScript.Echo "Cannot find " & sSourceFile
WScript.Quit
End If

sTargetFile = oSystem.GetBaseName(sSourceFile) & ".xlsx"
If iArgCount > 1 Then sTargetFile = WScript.Arguments(1)
If InStr(sTargetFile, "\") = 0 Then sTargetFile = oSystem.BuildPath(oShell.CurrentDirectory, sTargetFile)

Set oApp = CreateObject("Excel.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
oApp.Visible = False

Set oBooks = oApp.WorkBooks
' expression.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
WScript.Echo "Opening " & oSystem.GetFileName(sSourceFile)
Set oBook = oBooks.Open(sSourceFile)

For Each oSheet in oBook.Sheets
Set oRange = oSheet.UsedRange
iRowCount = oRange.Rows.Count
iColCount = oRange.Columns.Count
For iCol = 1 to iColCount
iNewCol = iColCount - iCol + 1
Set oCell = oSheet.Cells(1, iNewCol)
sHeader = oCell.Text
If sHeader = "" Then 
wscript.echo "Deleting column" & iNewCol
oCell.EntireColumn.Delete
End If
Next

Set oRange = oSheet.UsedRange
wscript.echo "listobject count " & oSheet.ListObjects.count
on error resume next
oRange.Select()
wscript.echo "areas count " & oRange.areas.count
on error goto 0
iRowCount = oRange.Rows.Count
iColCount = oRange.Columns.Count
wscript.echo iRowCount & " rows"
wscript.echo iColCount & " columns"
For iCol = 1 to iColCount
' wscript.echo "Column " & iCol
iMaxWidth = 0
For iRow = 1 to iRowCount
' wscript.echo "Row " & iRow
Set oCell = oRange.Cells(iRow, iCol)
If iRow = 1 Then sHeader = oCell.Text
iWidth = Len(oCell.Text)
If iMaxWidth < iWidth Then iMaxWidth = iWidth
Next
iWidth = 40
If iMaxWidth > iWidth Then
iMaxWidth = iWidth
oSheet.Columns(iCol).EntireColumn.WrapText = true
End If
 wscript.echo "Setting column " & sHeader & " to width " & iMaxWidth
' oCell.Column.EntireColumn.Width = iWidth
' oRange.Columns(iCol).EntireColumn.Width = iMaxWidth
' oSheet.Columns(iCol).EntireColumn.Width = iMaxWidth
oSheet.Columns(iCol).EntireColumn.ColumnWidth = iMaxWidth
' oSheet.Columns(iCol).Width = iMaxWidth
Next
oSheet.Rows.VerticalAlignment = xlVAlignTop    

Set oRange = oSheet.Range("A1")
oSheet.Names.Add "ColumnTitle", oRange

oSheet.Rows(1).Font.Bold = True


' Format code is above now
if false then
oSheet.Columns.AutoFit
oSheet.Cells.WrapText = True
oSheet.Rows.AutoFit
oSheet.Rows.VerticalAlignment = xlVAlignTop  
End If 
Next

' expression.SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
WScript.Echo "Saving " & oSystem.GetFileName(sTargetFile)
oBook.SaveAs sTargetFile, xlWorkbookDefault                

oBook.Close(0)
oApp.Quit

