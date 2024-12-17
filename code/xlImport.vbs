Option Explicit

Dim bCreateWorkbook
Dim iSpec, iArgCount
Dim oTables, oTable, oApp, oBooks, oBook, oSheet, oShell, oSystem, oFile
Dim aFiles, aImportFiles
Dim sRoot, sName, sDir, sSpec, sImportFile, sFile, sSourceFile, sTargetFile, sScriptVbs, sHomerLibVbs

Const msoAutomationSecurityLow = 1
Const xlWorkbookDefault = 51
Const xlCSV = 6

' xl vertical alignment enumeration
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

Function FileInclude(sFile)
With CreateObject("Scripting.FileSystemObject")
ExecuteGlobal .openTextFile(sFile).readAll()
End With
FileInclude = True
End Function

' Main
sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile) 
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs

iArgCount = WScript.Arguments.Count
If iArgCount < 2 Then
WScript.Echo "Specify a source .xlsx file as the first parameter."
WScript.Echo "specify .csv or other files to import as subsequent parameters."
WScript.Quit
End If

sSourceFile = WScript.Arguments(0)
sSourceFile = PathGetCurrentDirectory() & "\" & sSourceFile
' sSourceFile = "c:\azn\class\0\Audit.xlsx"
If InStr(sSourceFile, "\") = 0 Then sSourceFile = oSystem.BuildPath(oShell.CurrentDirectory, sSourceFile)
' If InStr(sSourceFile, "\") = 0 Then sSourceFile = oShell.CurrentDirectory & "\" & sSourceFile
If not oSystem.FileExists(sSourceFile) Then
' WScript.Echo "Cannot find " & sSourceFile
' W Script.Quit
bCreateWorkbook = True
Else bCreateWorkbook = False
End If

sName = PathGetName(sSourceFile)
sTargetFile = sSourceFile

sDir = PathGetCurrentDirectory()
aImportFiles = Array()
For iSpec = 1 to WScript.Arguments.Count - 1
sSpec = WSH.Arguments(iSpec)
aFiles = PathGetSpec(sDir, sSpec, "")
For Each sFile in aFiles
' print sFile
If Not ArrayContains(aImportFiles, sFile, True) Then ArrayAdd aImportFiles, sFile
Next
Next

print StringPlural("file", ArrayCount(aImportFiles)) & " to import"
Set oApp = CreateObject("Excel.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
oApp.Visible = False

Set oBooks = oApp.WorkBooks
If bCreateWorkbook Then
print "Creating " & sName
Set oBook = oBooks.Add
For Each oSheet in oBook.Sheets
If oBook.Sheets.Count > 1 Then oSheet.Delete
Next

Else
WScript.Echo "Opening " & sName
' expression.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
Set oBook = oBooks.Open(sSourceFile)
End If

const xlDelimited = 1
const xlTextQualifierDoubleQuote = 1

Const xlInsertDeleteCells = 1
Const xlInsertEntireRows = 2
Const xlOverwriteCells = 0

For each sImportFile in aImportFiles
sRoot = PathGetRoot(sImportFile)
print sRoot
On Error Resume Next
oBook.Sheets(sRoot).Delete
On Error GoTo 0

Set oSheet = oBook.Sheets.Add
oSheet.Name = sRoot
Set oTables = oSheet.QueryTables
Set oTable = oTables.Add("TEXT;" & sImportFile, oSheet.Range("A1"))
oTable.Name = sRoot
oTable.FieldNames = True
' oTable.RowNumbers = False
' oTable.FillAdjacentFormulas = False
' oTable.PreserveFormatting = True
' oTable.RefreshOnFileOpen = False
' oTable.RefreshStyle = xlInsertDeleteCells
oTable.RefreshStyle = xlInsertEntireRows
' oTable.SavePassword = False
oTable.SaveData = True
oTable.AdjustColumnWidth = False
' oTable.RefreshPeriod = 0
' oTable.TextFilePromptOnRefresh = False
' oTable.TextFilePlatform = 437
' oTable.TextFileStartRow = 1
oTable.TextFileParseType = xlDelimited
oTable.TextFileTextQualifier = xlTextQualifierDoubleQuote
oTable.TextFileConsecutiveDelimiter = True
oTable.TextFileTabDelimiter = False
' oTable.TextFileSemicolonDelimiter = False
oTable.TextFileCommaDelimiter = True
' oTable.TextFileSpaceDelimiter = False
' oTable.TextFileOtherDelimiter = "|"
' oTable.TextFileColumnDataTypes = Array(1, 1, 1)
' oTable.TextFileTrailingMinusNumbers = True
' oTable.Refresh BackgroundQuery = False
oTable.Refresh  
Next

WScript.Echo "Saving " & sName
If bCreateWorkbook Then
oBook.Sheets("Sheet1").Delete
' expression.SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
oBook.SaveAs sTargetFile, xlWorkbookDefault                
Else
oBook.Save
End If

oBook.Close(0)
oApp.Quit

