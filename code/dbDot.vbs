' oConn.Open "Driver={Microsoft Visual FoxPro Driver};" & _
' "SourceType=DBF;" & _
' "SourceDB=c:\somepath\mySourceDbFolder;" & _
' "Exclusive=No"
' For more information, see:  Visual FoxPro ODBC Driver and Q165492
' To view Microsoft KB articles related to ODBC Driver for Visual FoxPro, click here
' 
' 
' oConn.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
' "Dbq=c:\somepath\mydb.mdb;" & _            "Uid=admin;" & _
' "Pwd="
' 
' 
' oConn.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
' "Dbq=c:\somepath\mydb.mdb;" & _
' "Exclusive=1;" & _
' "Uid=admin;" & _
' "Pwd="
' 
' 
' OLE DB Provider for Microsoft Jet  For standard security
' oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
' "Data Source=c:\somepath\myDb.mdb;" & _
' "User Id=admin;" & _
' "Password="
' 
' 
' oConn.Mode = adModeShareExclusive
' oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _            "Data Source=c:\somepath\myDb.mdb;" & _
' "User Id=admin;" & _
' "Password="

 ' DSN={SQLite3 Datasource};Database=full-path-to-db;...
' RUNDLL32 sqlite3odbc.dll,install
' "PRAGMA table_info(...)" SQLite statement
' DSN={SQLite3 Datasource};Database=full-path-to-db;...
' Driver={SQLite3 ODBC Driver};Database=full-path-to-db;...
' Timeout (integer)	lock time out in milliseconds; default 100000
' SyncPragma (string)	value for PRAGMA SYNCHRONOUS; default empty
' C:\> RUNDLL32 [path]sqliteodbc.dll,install [quiet]
' C:\> RUNDLL32 [path]sqlite3odbc.dll,uninstall [quiet]

Option Explicit
WScript.Echo"Starting Dot"

' adodb getRows enumeration
const adGetRowsRest = -1

Const WindowStyle = 0 'hidden
Const HIDDEN = 0 ' window style
Const NORMAL = 1 ' window style
Const MAX = 3
Const Wait = True

Function FileInclude(sFile)
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

function getTables()
dim aReturn
dim oCatalog, oTables, oTable
dim sTable

arrayClear aReturn
Set oCatalog = CreateObject("ADOX.Catalog")
Set oCatalog.ActiveConnection = oConnect
print 11
Set oTables = oCatalog.Tables
print 12
print oTables.count
For Each oTable In oTables
sTable = oTable.Name
ArrayAdd aReturn, sTable
Next
getTables = aReturn
end function

function clipGetText()
dim oHtml

set oHtml = createObject("htmlfile")
' clipGetText = oHtml.ParentWindow.ClipboardData.getData("text")
call oHtml.ParentWindow.ClipboardData.setData("text", "test")
Set oHtml = Nothing
exit function

Dim oApp
Set oApp = CreateObject("InternetExplorer.Application")
oApp.Navigate("about:blank")
clipGetText = oApp.document.parentwindow.clipboardData.GetData("text")
oApp.Quit
set oApp = nothing
End Function

Function multiArrayToTables(aTable, sRoot, aFields, sExtensions)
Dim a, aRepeatFields, aSections, aKeys
Dim dExtensions, dFields, dRepeatFields, dIni, dSection
Dim iIndex, iRowCount, iSectionField, iSection, iField, iRow, iCol, iFieldCount, iSectionCount
Dim oWordTable, oRange, oDoc, oDocApp, oApp, oBook, oSheet
Dim sTargetExt, sCurDir, s, sTargetHtm, sTargetHtml, sTargetDocx, sRepeatField, sField, sSection, sKey, sValue, sTargetCsv, sTargetXlsx

' wdAutoFit enumeration
Const wdAutoFitContent = 1 'The table is automatically sized to fit the content contained in the table.
Const wdAutoFitFixed = 0 'The table is set to a fixed size, regardless of the content, and is not automatically sized.
Const wdAutoFitWindow = 2 'The table is automatically sized to the width of the active window.

Const wdFormatFilteredHtmlL = 10
Const wdFormatDocumentDefault = 16

' xl vertical alignment enumeration
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

Const xlWorkbookDefault = 51
Const xlCSV = 6
Const msoAutomationSecurityLow = 1

Set dExtensions = CreateDictionary
sExtensions = Replace(sExtensions, ".", " ")
sExtensions = StringReplaceAll(sExtensions, "  ", " ")
sExtensions = LCase(Trim(sExtensions))
a = Split(sExtensions, " ")
For Each s in A
' print s
dExtensions.Add s, ""
Next

' sTargetExt = PathGetUnique(sCurDir, sRoot, sExtensions)
sTargetCsv = PathGetUnique(sCurDir, sRoot, "csv")
sTargetDocx = PathGetUnique(sCurDir, sRoot, "docx")
sTargetHtm = PathGetUnique(sCurDir, sRoot, "Htm")
sTargetHtml = PathGetUnique(sCurDir, sRoot, "Html")
sTargetXlsx = PathGetUnique(sCurDir, sRoot, "xlsx")

Set oDocApp = CreateObject("Word.Application")
oDocApp.Visible = False
oDocApp.DisplayAlerts = False
oDocApp.ScreenUpdating = False

Set oDoc = oDocApp.Documents.Add
Set oRange = oDoc.Content
iFieldCount = arrayCount(aFields)
iRowCount = uBound(aTable, 2) + 1
' Start at 1 because of row of column headers
Set oWordTable = oDoc.Tables.Add(oRange, iRowCount + 1, iFieldCount)

Set oApp = CreateObject("Excel.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
oApp.Visible = False

Set oBook = oApp.Workbooks.Add
Set oSheet = oBook.Sheets(1)

' Set column headers
For iField = 0 to iFieldCount - 1
sField = aFields(iField)
oWordTable.Rows(1).Cells(iField + 1).Range.Text = sField
oSheet.Cells(1, iField + 1) = sField
Next

Set oRange = oSheet.Range("A1")
oSheet.Names.Add "ColumnTitle", oRange

set dFields = createDictionary
for each sField in aFields
' iIndex = arrayIndex(aAllFields, sField, true)
iIndex = arrayIndex(aFields, sField, true)
dFields.add sField, iIndex
' print sField & "=" & iIndex
next

' Populate rows
for iRow = 0 to iRowCount - 1
For iField = 0 to iFieldCount - 1
' sField = aFields(iField)
' iIndex = dFields(sField)
' print "field=" & sField
' print "index=" & iIndex
' print "row=" & iRow
' print "col=" & iField 

' vValue = aTable(iRow, iIndex)
' vValue = aTable(iIndex, iRow)
VValue = aTable(iField, iRow)
sValue = "" & vValue
' oWordTable.Rows(iRow + 2).Cells(iField + 1).Range.Text = vValue
if dExtensions.exists("docx") or dExtensions.exists("Html") or dExtensions.exists("Htmll") then oWordTable.Rows(iRow + 2).Cells(iField + 1).Range.Text = sValue
if dExtensions.exists("csv") or dExtensions.exists("xlsx") then oSheet.Cells(iRow + 2, iField + 1) = vValue
Next
if (iRow > 0) and (iRow mod 100) = 0 then print "row " & iRow
Next

oWordTable.AllowAutoFit = True
oWordTable.AutoFitBehavior wdAutoFitContent
' oWordTable.Columns.AutoFit
oWordTable.Rows(1).HeadingFormat = True
oWordTable.ApplyStyleHeadingRows = True
oDoc.Bookmarks.Add "ColumnTitle", oWordTable.Rows(1).Cells(2).Range
if dExtensions.Exists("docx") Then oDoc.SaveAs sTargetDocx, wdFormatDocumentDefault
if dExtensions.Exists("Html") Then oDoc.SaveAs sTargetHtml, wdFormatFilteredHtmlL  
if dExtensions.Exists("Htmll") Then oDoc.SaveAs sTargetHtmll, wdFormatFilteredHtmlL  
oDoc.Close 0
If Not oDocApp.NormalTemplate.Saved Then oDocApp.NormalTemplate.Saved = True
oDocApp.Quit

oSheet.Columns.AutoFit
oSheet.Cells.WrapText = True
oSheet.Rows.AutoFit
oSheet.Rows.VerticalAlignment = xlVAlignTop  

' print sTargetXlsx
' print dExtensions.count
' print dExtensions.exists("xlsx")
if dExtensions.Exists("csv") Then oSheet.SaveAs sTargetCsv, xlCsv
if dExtensions.Exists("xlsx") Then oSheet.SaveAs sTargetXlsx, xlWorkbookDefault  

oBook.Close 0
oApp.Quit
if fileExists(sTargetCsv) then shellOpen(sTargetCsv)
if fileExists(sTargetDocx) then shellOpen(sTargetDocx)
if fileExists(sTargetHtm) then shellOpen(sTargetHtm)
if fileExists(sTargetHtml) then shellOpen(sTargetHtml)
if fileExists(sTargetXlsx) then shellOpen(sTargetXlsx)
End Function

function cmdChoose(sPrompt, sDefaultChar)
' Choices separated by /
' Trigger letter is cap in choice
' Parenthetical default in lower case could precede ? at end
' Example: Continue/Stop (c)?
dim iPart
dim sChoice, sChar

aParts = split(sPrompt, "/")
set dChoices = createDictionary
for each sPart in aParts
aChars = array()
for iChar = 1 to len(sPart)
sChar = mid(sPrompt, iChar, 1)
if stringIsUpper(sChar) then
dChoices.add sChar, sPart
exit for
end if
next
next

sChar = cmdPrompt(sPrompt)
if not dChoices.key(sChar) then sChar = sDefaultChar
sChoice = dChoices(sChar)
cmdChoices = sChoice
end function

function inputForm(sTitle, aFields)
dim dOutput
dim sControl, sCodeDir, sCommand, sCurDir, sSettingsDir, sIniFormExe, sInputBase, sInputIni, sInputTxt, sName, sOutputBase, sOutputIni, sOutputTxt, sTempDir, sTempTmp, sText, sValue

' sTempDir = PathGetSpecialFolder("TEMP")
sTempDir = ShellExpandEnvironmentVariables("%TEMP%")
sTempTmp = PathCombine(sTempDir, "temp.tmp")
sCodeDir = PathGetFolder(WScript.ScriptFullName)
sSettingsDir = StringChopRight(sCodeDir, 4) + "settings"

sIniFormExe = PathCombine(sCodeDir, "IniForm.exe")

sInputBase = "input.ini"
sInputIni = PathCombine(sTempDir, sInputBase)
sInputTxt = PathCombine(sTempDir, "input.txt")
sOutputTxt = PathCombine(sTempDir, "Output.txt")
sOutputBase = "Output.ini"
sOutputIni = PathCombine(sTempDir, sOutputBase)

sText = "[" & sTitle & "]" & vbCrLf
sText = sText & "control=form" & vbCrLf
for each sName in aFields
sValue = "" & oTable.fields(sName).value
sText = sText & "[" & sName & "]" & vbCrLf
sControl = "edit"
if stringContains(sValue, vbLf, true) or stringContains(sValue, vbCr, true) then sControl = "memo"
sText = sText & "control=" & sControl & vbCrLf
sText = sText & "value=" & sValue & vbCrLf
next
sText = sText & "[OK]" & vbCrLf
sText = sText & "control=button" & vbCrLf
sText = sText & "[Cancel]" & vbCrLf
sText = sText & "control=button" & vbCrLf
stringToFile sText, sInputIni

StringToFile sText, sInputIni
PathSetCurrentDirectory(sTempDir)
sCommand = StringQuote(sIniFormExe) & " " & StringQuote(sTempDir)
' ShellRun sCommand, NORMAL, Wait
ShellRun sCommand, MAX, Wait
' sText = FileToString(sOutputTxt)
' sText = StringConvertToUnixLineBreak(sText)
' a = RegExpExtract(sText, "(^|\n)\[\[Pick\]\](.|\n)*?($|\n\[\[)", False)
sText = ""
' sText = StringTrimWhiteSpace(sText)
PathSetCurrentDirectory(sCurDir)
FileDelete sInputIni
FileDelete sInputTxt
FileDelete sOutputTxt

set dOutput = createDictionary
If FileExists(sOutputIni) Then 
Set dOutput = IniToDictionary(sOutputIni)
FileDelete sOutputIni
If dOutput("Results")("Cancel") <> "1" Then
for each sName in dOutput("Results").keys
if arrayContains(aAllFields, sName, false) then 
sValue = dOutput("Results")(sName)
' dim d
' set d = dOutput("Results")
' sValue = d(sName)
sValue = Replace(sValue, "|", vblf)
oTable.fields(sName).value = sValue
end if
next
oTable.update
end if
end if
set inputForm = dOutput
end function

function endProgram()
print "Ending DbDot"
oTable.close
oConnect.close
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
wscript.quit
end function

function stringToFields(sText)
dim i
dim s
dim a, aReturn

aReturn = array()
a = stringToTrimArray(sText, ",")

for each s in a
if arrayContains(aAllFields, s, true) then arrayAdd aReturn, s
next
stringToFields = aReturn
end function

function fillFieldArrays(aAllFields, aAddFields, aEditFields, aExportFields, aListFields, aNextFields, aShowFields, aViewFields)
dim oField
dim sName

arrayClear aAllFields
arrayClear aAddFields
for each oField in oTable.fields
sName = oField.name
arrayAdd aAllFields, sName
select case sName
case sIdField, "added", "updated", "marked", "look", "unq"
case else
arrayAdd aAddFields, sName
end select
next
aEditFields = arrayCopy(aAddFields)
aExportFields = arrayCopy(aExportFields)
aListFields = arrayCopy(aAddFields)
aNextFields = arrayCopy(aAllFields)
aShowFields = arrayCopy(aAddFields)
' aShowFields = array("look")
aViewFields = arrayCopy(aAddFields)
end function

function getRowString(aFields)
dim a1, a2
dim iField, iFieldBound, iCount, i, iBookmark
dim oField
Dim sName, sValue, sReturn

iFieldBound = arrayBound(aFields)
for iField = 0 to iFieldBound
sName = aFields(iField)
sValue = "" & oTable.fields(sName).value
sReturn = sReturn & sValue
if iField < iFieldBound then sReturn = sReturn & ", "
next
getRowString = sReturn
exit function

iBookmark = oTable.Bookmark
a2 = oTable.getRows(1, iBookmark)
' print arrayBound(a2, 1)
' print arrayBound(a2, 2)
oTable.bookmark = iBookmark
iCount = oTable.fields.count
' print iCount
' a1 = array(iCount)
redim a1(iCount)
' a1 = array()
for i = 0 to iCount - 1
a1(i) = a2(i, 0)
' arrayAdd a1,  a2(i, 0)
sReturn = sReturn & a2(i, 0) & vbTab
next
' sReturn = join(a1, vbTab)
getRowString = sReturn
end function

function ToString(x)
ToString = "null"
on error resume next
ToString = cstr(x)
on error goto 0
end function

function printRow()
print "row " & oTable.AbsolutePosition
end function

function arrayCommand(aItems)
dim aReturn, aParts
dim iPart, iItem
dim sPart, sItem

aReturn = array()
aParts = split(aItems(0), " ")
for iPart = 0 to arrayBound(aParts)
arrayAdd aReturn, aParts(iPart)
next

for iItem = 1 to arrayBound(aItems)
arrayAdd aReturn, aItems(iItem)
next
arrayCommand = aReturn
end function

function stringToTrimArray(sItems, sDelim)
dim aParams, aItems, aReturn
dim iItem
dim sItem

aItems = split(sItems, sDelim)
aReturn = array()
for iItem = 0 to arrayBound(aItems)
sItem = trim(aItems(iItem))
if len(sItem) > 0 then arrayAdd aReturn, sItem
next
stringToTrimArray = aReturn
end function

Function DBCreateDatabase(sDb, sConnectString)
Dim oCatalog

dbCreateDatabase = False
If Not FileDelete(sDb) Then Exit Function
Set oCatalog = CreateObject("ADOX.Catalog")
oCatalog.Create sConnectString
Set oCatalog = Nothing
DBCreateDatabase = FileExists(sDb)
End Function

Function DBEval(oRS, sCode)
Dim aParts
Dim s, sPart
Dim v

With oRS
aParts = Split(sCode, ";")
v = vbNull
For Each sPart In aParts
s = RegExpReplace(sPart, "\$([_a-zA-Z0-9]+)", ".Fields(" & Chr(34) & "$1" & Chr(34) & ").Value", True)
' s = Replace(s, "'", Chr(34))
sHomer = sHomer & s & vbCrLf
On Error Resume Next
v = Eval(s)
On Error GoTo 0
If Err.Number Then v = vbNull
Next
End With
DBEval = v
End Function

Function DBExecuteCommand (oCon, sCommand, bShowDone)
Dim iResult
Dim sResult, sText, sTitle
Dim vResult

if oCon is nothing then
' speak "empty connection found in DBExecuteCommand"
set oCon = DBOpenConnection("DBDialog") ' added by Chip
end if
sText = ""
Set vResult = Nothing
On Error Resume Next
Set vResult = oCon.Execute(sCommand, iResult, adCmdText)
On Error GoTo 0
If Err.Number Or vResult Is Nothing Then
sTitle = "Error"
sText = sText & DBGetErrorText(oCon) & vbCrLf
DBExecuteCommand = False
Else
sTitle = "Done"
sResult = ""
On Error Resume Next
' sResult = vResult.GetString(adClipString, vResult.RecordCount, vbTab, vbCrLf, "")
' sResult = vResult.GetString(adClipString, vbNull, vbTab, vbCrLf, "")
sResult = vResult.GetString
sResult = StringConvertToWinLineBreak(sResult)
vResult.Close
On Error GoTo 0
If IsNumeric(iResult) And Num(iResult) <> 0 Then
' sText = sText & "Affected " & StringPlural("record", iResult) & vbCrLf
sText = sText & "Records Affected: " & iResult & vbCrLf & vbCrLf & sResult
End If
If Len(sResult) = 0 Then
' sText = sText & "Done!" & vbCrLf
Else
' sText = sText & "Output:" & vbCrLf
sText = sText & sResult & vbCrLf
End If
DBExecuteCommand = True
End If
sText = sText & vbCrLf & "SQL Command: " & vbCrLf & sCommand & vbCrLf

' DialogShow "Results", sText
' If sTitle = "Error" Or bShowDone Then DialogShow sTitle, sText
If sTitle = "Error" Or bShowDone Then DialogMemo sTitle, "", sText, True
End Function

Function DBOpenRecordSet(oConnection, sTable)
Dim oRS

Set oRS = CreateObject("AdoDb.RecordSet")
oRS.CursorLocation = adUseServer
oRS.Open sTable, oConnection, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
Set DBOpenRecordSet = oRS
End Function

Function DBIsRecord(iPosition)
DBIsRecord = False
If iPosition = AdPosBof Then Exit Function
If iPosition = AdPosEof Then Exit Function
If iPosition = AdPosUnknown Then Exit Function
DBIsRecord = True
End Function

Function DBOpenConnection(sMdb)
Dim oConnection
Dim s, sConnection

' sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
Set oConnection = CreateObject("AdoDB.Connection")
' oConnection.Open sConnection
' oConnection.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=C:\Documents and Settings\Owner.DWEEB1\Application Data\GW Micro\Window-Eyes\users\default\DbDialog.mdb"
s = PathCombine(ClientInformation.ScriptPath, "DbDialog.mdb")
oConnection.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & s
Set DBOpenConnection = oConnection
End Function


Function DBGetCondition(sField, sType, sValue)
Dim b
Dim s, sPrefix, sRest

sPrefix = ""
sRest = sValue
If StringLead(sRest, "<>", False) Or StringLead(sRest, ">=", False) Or StringLead(sRest, "<=", False) Then
sPrefix = Left(sRest, 2)
sRest = StringChopLeft(sRest, 2)
ElseIf StringLead(sRest, "=", False) Or StringLead(sRest, ">", False) Or StringLead(sRest, "<", False) Then
sPrefix = Left(sRest, 1)
sRest = StringChopLeft(sRest, 1)
End If

Select Case sType
Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Counter", "Identity", "AutoIncrement"
s = sField & " " & sPrefix & " " & sRest
Case "Date", "DateTime"
s = sField & " " & sPrefix & " " & "#" & sRest & "#"
Case "String", "Text", "Memo"
If Len(sPrefix) = 0 Then
s = sField & " like " & "'*" & sRest & "*'"
Else
s = sField & " " & sPrefix & " " & "'" & sRest & "'"
End If
Case "Boolean", "YesNo"
b = CBool(sValue)
If b Then
s = "True"
Else
s = "False"
End If
s = sField
' Case "Empty", "Null", "Object", "Unknown", "Nothing", "Error"
Case Else
' s = "vbNull"
s = ""
End Select

' DbGetCondition = sField & " = " & s
DbGetCondition = s
End Function

Function DBGetErrorText(oConnection)
Dim oErrors, oError
Dim sText
if oConnection is nothing then
' speak "empty connection found in DBGetErrorText"
set oConnection = DBOpenConnection("DBDialog") ' added by Chip
end if

' If Err.Number = -2147467259 Then sText = "ADO Provider error" & vbCrLf
If Err.Number Then
sText = "Runtime Error Number: " & Err.Number & vbCrLf
sText = "Description: " & Err.Description & vbCrLf
End If

Set oErrors = oConnection.Errors
For Each oError In oErrors
sText = sText & "ADO Provider Error Number: " & oError.Number & vbCrLf
sText = sText & "Description: " & oError.Description & vbCrLf
sText = sText & "Native Error: " & oError.NativeError & vbCrLf
sText = sText & "SQL State: " & oError.SQLState & vbCrLf
sText = sText & "Source: " & oError.Source & vbCrLf
Next
Err.Clear
oErrors.Clear
DBGetErrorText = sText
End Function

Function DBGetFieldNamesAndTypes(sCreateTable, aNames, aTypes)
Dim a, aParts
Dim iIndex
Dim sType, sParts, sPart

ArrayClear aNames
ArrayClear aTypes
aParts = DBArray(sCreateTable)
For Each sPart In aParts
a = Split(sPart, " ")
If UBound(a) >= 1 Then
ArrayAdd aNames, a(0)
sType = a(1)
iIndex = InStr(sType, "(")
If iIndex Then sType = Left(sType, iIndex - 1)
ArrayAdd aTypes, sType
End If
Next
End Function

Function DBGetIndexNames(sMdb, sTable)
Dim aIndexes
Dim oCatalog, oTables, oTable, oIndexes, oIndex, oConnection
Dim sConnection, sIndex

DBGetIndexNames = Array()
If Not FileExists (sMdb) Then Exit Function

ArrayClear aIndexes
sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
Set oCatalog = CreateObject("ADOX.Catalog")
Set oConnection = DBOpenConnection(sMdb)
Set oCatalog.ActiveConnection = oConnection
Set oTables = oCatalog.Tables
For Each oTable in oTables
If StrComp(sTable, oTable.Name) = 0 Then
Set oIndexes = oTable.Indexes
For Each oIndex In oIndexes
sIndex = oIndex.Name
ArrayAdd aIndexes, sIndex
Next
Exit For
End If
Next
oConnection.Close
Set oConnection = Nothing
Set oCatalog = Nothing
Set oIndexes = Nothing
Set oIndex = Nothing
Set oTables = Nothing
Set oTable = Nothing
DBGetIndexNames = aIndexes
End Function

Function DBGetFieldLabels(aFields)
Dim aLabels
Dim i
Dim sField, sLabel

ArrayClear aLabels
' For i = 0 To oRS.Fields.Count - 1
' sField = oRS.Fields(i)
For Each sField in aFields
sLabel = Replace(sField, "_", " ")
If InStr(sLabel, "&") = 0 Then sLabel = "&" & sLabel
ArrayAdd aLabels, sLabel
Next
DBGetFieldLabels = aLabels
End Function

Function DBGetLookText(oDB, oRS, sTable, sIDField, iID)
Dim aFields
Dim oConnection, oIni
Dim s, sFields, SSQL

Set oConnection = oRS.ActiveConnection
if oConnection is nothing then
' speak "empty connection found in DBGetLookText"
set oConnection = DBOpenConnection("DBDialog") ' added by Chip
end if
Set oIni = IniFile(oDB("Cfg"))
sFields = oIni.Text(sTable, "ResultFields", "")
aFields = DBArray(sFields)
ArrayRemove aFields, 0
sFields = DBString(aFields)
If Len(sFields) = 0 Then Exit Function

sSQL = "select " & sFields & " from " & sTable & " Where " & sIDField & " = " & iID
' MsgBox SSQL
Set oRS = oConnection.Execute(sSQL)
s = oRS.GetString
s = StringConvertToWinLineBreak(s)
DBGetLookText = s
oRS.Close
Set oRS = Nothing
Set oIni = Nothing
End Function

Function DBGetString(sConnection, sTable, sFields, sFilter, sIndex)
Dim oConnection, oRS
Dim SSQL

Set oConnection = CreateObject("AdoDB.Connection")
oConnection.Open sConnection
Set oRS = CreateObject("AdoDB.RecordSet")
sSQL = "select " & sFields & " from " & sTable
sFilter = CSTR(sFilter)
If sFilter <> "0" And Len(sFilter) > 0 Then sSQL = sSQL & " where " & sFilter
sIndex = CSTR(sIndex)
If sIndex <> "0" And Len(sIndex) > 0 Then sSQL = sSQL & " order by " & sIndex
' MsgBox SSQL
oRS.CursorLocation = adUseServer
oRS.Open sSQL, oConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
' If DBIsRecord(oRS.AbsolutePosition) Then DBGetString = oRS.GetString(adClipString, oRS.RecordCount, vbTab, vbCrLf, "")
On Error Resume Next
DBGetString = oRS.GetString(adClipString, oRS.RecordCount, vbTab, vbCrLf, "")
On Error GoTo 0
oRS.Close
Set oRS = Nothing
oConnection.Close
Set oConnection = Nothing
End Function

Function DBGetTableNames(sMdb)
Dim aTables
Dim oCatalog, oTables, oTable, oConnection
Dim sConnection, sTable

DBGetTableNames = Array()
If Not FileExists (sMdb) Then Exit Function

ArrayClear aTables
' sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
sConnection = sConnectString
Set oCatalog = CreateObject("ADOX.Catalog")
Set oConnection = DBOpenConnection(sMdb)
Set oCatalog.ActiveConnection = oConnection
Set oTables = oCatalog.Tables
For Each oTable In oTables
sTable = oTable.Name
ArrayAdd aTables, sTable
Next
oConnection.Close
Set oConnection = Nothing
Set oCatalog = Nothing
Set oTables = Nothing
Set oTable = Nothing
DBGetTableNames = aTables
End Function

Function DBGetRecordSet(oDB)
Dim oConnection, oRS
Dim sTable

Set oDialogConnection = DBOpenConnection(oDB)
sTable = oDB("Table")
Set oRS = CreateObject("Adodb.RecordSet")
oRS.CursorLocation = adUseClient
oRS.Open sTable, oDialogConnection, adOpenKeySet, adLockOptimistic, adCmdTableDirect
Set DBGetRecordSet = oRS
End Function

Function DBStringToVariant(sValue, sType)
Dim vResult

Select Case sType
Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Counter", "Identity", "AutoIncrement"
If IsNumeric(sValue) Then
vResult = Num(sValue)
Else
vResult = vbNull
End If
Case "Date", "DateTime"
If IsDate(sValue) Then
vResult = CDate(sValue)
Else
vResult = vbNull
End If
Case "String", "Text", "Memo"
vResult = vbNull
On Error Resume Next
vResult = CStr(sValue)
On Error GoTo 0
Case "Boolean", "YesNo"
vResult = CBool(sValue)
' Case "Empty", "Null", "Object", "Unknown", "Nothing", "Error"
Case Else
vResult = vbNull
End Select
DBStringToVariant = vResult
End Function

Function DBVariantToString(vValue)
Dim sResult

DBVariantToString = ""
sResult = ""
If TypeName(vValue) = "DateTime" Or TypeName(vValue) = "Date" Then
If vValue = DateValue("December 31, 1899") Then Exit Function
End If

On Error Resume Next
sResult = CStr(vValue)
On Error GoTo 0
DBVariantToString = sResult
End Function

' Main
' ADO constants
Const adAffectAll = 3
Const adAffectCurrent = 1
Const adAffectGroup = 2
Const adBatchOptimistic = 4
Const adBookmarkCurrent = 0
Const adBookmarkFirst = 1
Const adBookmarkLast = 2
Const adClipString = 2
Const adCmdTableDirect = 512
Const adCmdText = 1
Const adExecuteNoRecords = 128
Const adFilterAffectedRecords = 2
Const adFilterConflictingRecords = 5
Const adFilterFetchedRecords = 3
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adLockOptimistic = 3
Const adLockReadOnly = 1
Const adOpenDynamic = 2
Const adOpenForwardOnly = 0
Const adOpenKeySet = 1
Const adOpenStatic = 3
Const adPosBOF = -2
Const adPosEOF = -3
Const adPosUnknown = -1
Const adSearchBackward = -1
Const adSearchForward	= 1
Const adUseClient = 3
Const adUseServer = 2

dim aDbTables, a, aFields, aTable, aArgs, aAllFields, aAddFields, aEditFields, aExportFields, aListFields, aNextFields, aShowFields, aViewFields, aParams
dim dResults, dTables, dTable
dim iShortest, iLongest, iLength, iChoice, iLoop, iArgCount, iArg, iName, iPosition, iBookmark, iRecordsAffected, iParam, iJump
Dim oField, oFields, oRs, oFile, oTable, oConnect, oSystem
dim sTargetLog, sPaxDbRoot, sTempFile, sExtensions, sShortest, sLongest, sChoice, sShow, sUpdate, sInitial, sList, sIdField, sName, sView, sSeek, sRow, sSetting, sPosition, sSort, sFilter, sParam, sInputCmdRest, sInputParamRest, sValue, sFind, sField, sCommand, sInput, sSql, sTable, sConnectString, sCodeDir, sCurDir, sDir, sFile, sHomerLibVbs, sSettingsDir, sPaxDb, sPaxDbBase, sScriptVbs, sSQLite3Exe, sTempDir, sTempTmp, sWildcards
dim vMin, vMax, vValue

sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile)
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs

iArgCount = WScript.Arguments.Count
sWildcards = "*.*"
If iArgCount > 0 Then sWildcards = WScript.Arguments(0)

sCurDir = PathGetCurrentDirectory()
sTempDir = PathGetSpecialFolder("TEMP")
sTempDir = ShellExpandEnvironmentVariables("%TEMP%")
sTempTmp = PathCombine(sTempDir, "temp.tmp")
sCodeDir = PathGetFolder(WScript.ScriptFullName)
sSettingsDir = StringChopRight(sCodeDir, 4) + "settings"

sPaxDbRoot = "Pax"
sPaxDbBase = sPaxDbRoot & ".db"
sPaxDb = PathCombine(sSettingsDir, sPaxDbBase)
sTargetLog = pathCombine(sSettingsDir, sPaxDbRoot & ".log")

sCurDir = pathGetCurrentDirectory
sTable = ""
iArgCount = wscript.arguments.count
if iArgCount > 0 then 
sPaxDb = wscript.arguments(0)
if inStr(sPaxDb, "\") = 0 then sPaxDb = pathCombine(sCurDir, sPaxDb)
if iArgCount > 1 then sTable = wscript.arguments(1)
end if

sConnectString = "DRIVER=SQLite3 ODBC Driver;Database=" & sPaxDb & ";"
' sConnectString = "DSN={SQLite3 Datasource};Database=" & sPaxDb &";"
sSQLite3Exe = PathCombine(sTempDir, "SQLite3.exe")

Set oConnect = CreateObject("ADODB.Connection")
wscript.echo sConnectString
oConnect.open sConnectString
aDbTables = getTables()
print arrayCount(aDbTables)
if len(sTable) = 0 then sTable = aDbTables(0)
sSql = ""
set dTables = createDictionary
do while true
print "Table " & sTable
sIdField = mid(sTable, 1, len(sTable) - 1) & "_id"
if dTables.exists(sTable) then
set oTable = dTables(sTable)("object")
sSql = dTables(sTable)("sql")
sIdField = dTables(sTable)("idField")
sUpdate = dTables(sTable)("update")
aAllFields = dTables(sTable)("allFields")
aAddFields = dTables(sTable)("addFields")
aEditFields = dTables(sTable)("editFields")
aExportFields = dTables(sTable)("exportFields")
aListFields = dTables(sTable)("listFields")
aNextFields = dTables(sTable)("nextFields")
aShowFields = dTables(sTable)("showFields")
aViewFields = dTables(sTable)("viewFields")
else
set oTable = CreateObject("ADODB.Recordset")
oTable.CursorLocation = adUseClient
' if stringLead(sTable, "select", true) then
if len(sSql) > 0 then
' oTable.Open sSql, oConnect, adOpenKeySet, adLockOptimistic, adCmdText
' oTable.Open sSql, oConnect, adOpenStatic, adLockOptimistic, adCmdText
oTable.Open sSql, oConnect, adOpenDynamic, adLockOptimistic, adCmdText
else
sSql = ""
oTable.Open sTable, oConnect, adOpenKeySet, adLockOptimistic, adCmdTableDirect
end if
print stringPlural("field", oTable.fields.count)
print stringPlural("record", oTable.recordCount)
fillFieldArrays aAllFields, aAddFields, aEditFields, aExportFields, aListFields, aNextFields, aShowFields, aViewFields
set dTable = createDictionary
dTable.add "sql", sSql
dTable.add "idField", sIdField
dTable.add "update", sUpdate
dTable.add "allFields", aAllFields
dTable.add "addFields", aAddFields
dTable.add "editFields", aEditFields
dTable.add "exportFields", aExportFields
dTable.add "listFields", aListFields
dTable.add "nextFields", aNextFields
dTable.add "showFields", aShowFields
dTable.add "viewFields", aViewFields
dTable.add "object", oTable

dTables.add sTable, dTable
end if

' printRow
do while true
sInput = cmdPrompt(". ")
sInitial = ""
if len(sInput) > 0 then sInitial = left(sInput, 1)
if stringContains(";/?+-", sInitial, true) then sInput = trim(sInitial & " " & mid(sInput, 2))
aArgs = stringToTrimArray(sInput, " ")
sCommand = ""
sInputCmdRest = ""
sParam = ""
sInputParamRest = ""
if arrayCount(aArgs) > 0 then
sCommand = lCase(aArgs(0))
aArgs = arrayRemove(aArgs, 0)
sInputCmdRest = join(aArgs, " ")
end if

if arrayCount(aArgs) > 0 then
sParam = lCase(aArgs(0))
aArgs = arrayRemove(aArgs, 0)
sInputParamRest = join(aArgs, " ")
end if

select case sCommand
case ""
printRow
case "add", "&"
oConnect.beginTrans
oTable.addNew
a = aAddFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
print sUpdate
if sUpdate = "gui" then
set dResults = inPutForm(sCommand, a)
else
for each sName in a
sValue = "" & oTable.fields(sName)
' print sName & " = " & sValue
sValue = cmdPrompt(sName & ": ")
if len(sValue) > 0 then oTable.fields(sName) = sValue
next
end if
sChoice = cmdPrompt("Save +y/n?")
if sChoice = "n" then 
print 9
oConnect.rollbackTrans
else
oConnect.commitTrans
end if
case "edit", "^"
a = aEditFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
if sUpdate = "gui" then
set dResults = inPutForm(sCommand, a)
else
for each sName in a
sValue = "" & oTable.fields(sName)
print sName & " = " & sValue

sValue = cmdPrompt(sName & ": ")
' sValue = dialogInput("Input", sName & ": ", sValue)
if len(sValue) > 0 then oTable.fields(sName) = sValue
next
end if
case "list"
a = aListFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
iLoop = 1
do while not oTable.eof
sList = ""
for iName = 0 to arrayBound(a)
sName = a(iName)
sValue = "" & oTable.fields(sName)
sList = sList & sValue
if iName < arrayBound(a) then sList = sList & ", "
next
print iLoop & ". " & sList

if (iLoop mod 10) = 0 then
printBlank
sChoice = cmdPrompt("Choose number or just Enter to continue:")
if sChoice = "" then sChoice = "0"
iChoice = cInt(sChoice)
if (iChoice >=1) and (iChoice <= iLoop) then
oTable.move iLoop - iChoice - 2
printRow
exit do
end if
end if
oTable.moveNext
if oTable.eof then
print "End"
oTable.movePrevious
printRow
exit do
end if
iLoop = iLoop + 1
loop
case "log"
if sParam = "delete" then
fileDelete sTargetLog
else
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
shellOpen sTargetLog
end if
case "view", "="
a = aViewFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
for each sName in a
sValue = "" & oTable.fields(sName).value
sView = sView & sValue
if iName < arrayBound(aViewFields) then sView = sView & " | "
print sName & " = " & sValue
next
case "count"
print oTable.recordCount
case "mark"
if sParam = "all" then
iBookmark = oTable.bookmark
oTable.moveFirst
do while not oTable.eof
oTable.fields("marked").value = true
oTable.moveNext
loop
oTable.bookmark = iBookmark
else
oTable.fields("marked") = true
end if
printRow
case "unmark"
if sParam = "all" then
iBookmark = oTable.bookmark
oTable.moveFirst
do while not oTable.eof
oTable.fields("marked").value = false
oTable.moveNext
loop
oTable.bookmark = iBookmark
else
oTable.fields("marked") = false
end if
printRow
case "last"
oTable.moveLast
printRow
case "export"
sExtensions = sParam
if len(sExtensions) = 0 then sExtensions = "xlsx"
if sExtensions = "c" then sExtensions = "csv"
if sExtensions = "d" then sExtensions = "docx"
if sExtensions = "h" then sExtensions = "Html"
if sExtensions = "x" then sExtensions = "xlsx"

aFields = aExportFields
if len(sInputParamRest) > 0 then aFields = stringToFields(sInputParamRest)
' aTable = oTable.getRows(oTable.recordCount, 1)
aTable = oTable.getRows(adGetRowsRest  , 1, aFields)
call multiArrayToTables(aTable, sTable, aFields, sExtensions)
case "fields"
for each sField in aAllFields
print sField
next
case "filter"
sFilter = sInputCmdRest
if sParam = "marked" then sFilter = "marked = true"
if sParam = "unmarked" then sFilter = "marked = false"
oTable.Filter = sFilter
print stringPlural("record", oTable.RecordCount)
case "find"
sFind = sInputCmdRest
oTable.find sFind
case "goto", "go", "#"
sPosition = sParam
if sPosition = "" then
' do nothing
elseIf sPosition = "bookmark" then
oTable.bookmark = iBookmark
else
iPosition = cInt(sPosition)
oTable.absolutePosition  = iPosition
end if
printRow
case "jump"
iJump = cInt(sParam)
oTable.move iJump
printRow
case "longest"
sField = sParam
iLongest = 0
oTable.moveFirst
do while not oTable.eof
sValue = "" & oTable.fields(sField).value
iLength = len(sValue)
if iLength > iLongest then
iLongest = iLength
sLongest = sValue
iPosition = oTable.absolutePosition
end if
oTable.moveNext
loop
print stringPlural("character", iLongest)
print sLongest
oTable.absolutePosition = iPosition
case "max"
sField = sParam
oTable.moveFirst
vMax = oTable.fields(sField).value
iPosition = 1
do while not oTable.eof
vValue = oTable.fields(sField).value
if vValue > vMax then
vMax = vValue
iPosition = oTable.absolutePosition
end if
oTable.moveNext
loop
print vMax
oTable.absolutePosition = iPosition
case "min"
sField = sParam
oTable.moveFirst
vMin = oTable.fields(sField).value
iPosition = 1
do while not oTable.eof
vValue = oTable.fields(sField).value
if vValue < vMin then
vMin = vValue
iPosition = oTable.absolutePosition
end if
oTable.moveNext
loop
print vMin
oTable.absolutePosition = iPosition
case "shortest"
sField = sParam
iShortest = 2000000000
oTable.moveFirst
do while not oTable.eof
sValue = "" & oTable.fields(sField).value
iLength = len(sValue)
if iLength < iShortest then
iShortest = iLength
sShortest = sValue
iPosition = oTable.absolutePosition
end if
oTable.moveNext
loop
print stringPlural("character", iShortest)
print sShortest
oTable.absolutePosition = iPosition
case "next", "+"
iBookmark = oTable.bookMark
if len(sInputCmdRest) = 0 then
oTable.moveNext
else
do while true
oTable.moveNext
if oTable.eof then 
print "Not found"
oTable.bookmark = iBookmark
exit do
end if
sRow = getRowString(aNextFields)
if stringContains(sRow, sInputCmdRest, true) then exit do
loop
end if
printRow
case "previous", "prev", "-"
iBookmark = oTable.bookMark
if len(sInputCmdRest) = 0 then
oTable.movePrevious
else
do while not oTable.bof
oTable.movePrevious
sRow = getRowString(aNextFields)
if stringContains(sRow, sInputCmdRest, true) then exit do
loop
end if
if oTable.bof then
print "Not found"
oTable.bookmark = iBookmark
end if
printRow
case "quit", "exit", "x"
endProgram
case "exec", ";"
' on error resume next
set oRs = oConnect.Execute(sInputCmdRest, iRecordsAffected, adCmdText)
' on error goto 0
print stringPlural("record", iRecordsAffected) & " affected"
case "remove"
oTable.delete
oTable.moveNext
printRow
case "requery"
oTable.requery
printRow
case "seek"
sSeek = sParam
sRow = getRowString
print sRow
case "select", "*"
sIdField = ""
sSql = sInputCmdRest
sSql = "select " & sSql
sTable = "select"
exit do
case "set"
sSetting = sParam
select case sSetting
case "bookmark"
iBookmark = oTable.bookmark
printRow
case "add"
aAddFields = stringToFields(sInputParamRest)
case "edit"
aEditFields = stringToFields(sInputParamRest)
case "list"
aListFields = stringToFields(sInputParamRest)
case "next"
aNextFields = stringToFields(sInputParamRest)
case "form"
sUpdate = sInputParamRest
' if len(sUpdate) = 0 then sUpdate = "cmd"
if sUpdate <> "gui" then sUpdate = "cmd"
case "show"
aShowFields = stringToFields(sInputParamRest)
case "view"
aViewFields = arrayCopy(aAllFields)
if len(sInputParamRest) > 0 then aViewFields = stringToFields(sInputParamRest)
case else
print "unknown setting"
end select
case "show", "?"
a = aShowFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
sShow = getRowString(a)
print sShow
case "sort"
sSort = sInputCmdRest
oTable.Sort = sSort
' oTable.Requery
' oTable.resync
printRow
case "table", "@"
if len(sParam) = 0 then
print sTable
else
sTable = sParam
exit do
end if

sTable = sParam
case "first"
oTable.moveFirst
' printRow
case else
on error resume next
execute sInput
on error goto 0
end select
loop
loop
endProgram
