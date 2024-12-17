Option Explicit
WScript.Echo"Starting Dot"

Function FileInclude(sFile)
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

Function DBAddItemFromRecord(oDB, oRS, oDlg, sButton)
Dim aResults
Dim IID, iIndex, iBookmark
Dim oLst
Dim sText, sIDField

DBAddItemFromRecord = Array(0, 0)
If Not oHomer.DBIsRecord(oRS.AbsolutePosition) Then Exit Function

Set oLst = oDlg.Control("lst")
sIDField = oDB("IDField")
iID = oRS.Fields(sIDField).Value
sText = oHomer.DBGetSelectText(oDB, oRS, oDlg, sButton)
' iIndex = oLst.Add(sText, iID)
iIndex = oLst.Insert(oRS.AbsolutePosition, sText, iID)
oLst.FocusedIndex = iIndex
oHomer.DBSetStatus oDB, oRS, oDlg, sButton
DBAddItemFromRecord = Array(iID, iIndex)
End Function

Function DBArray(s)
DBArray = Split(s, ", ")
End Function

Function DBString(a)
DBString = Join(a, ", ")
End Function

Function DBClose(oDB, oRS, oDlg, sButton, sMdb, sCfg, sTable)
Dim aFields, aResults
Dim oIni, oConnection
Dim s, sIDField, sFields, sField, sValue

Set oIni = IniFile(oDB("Cfg"))
sIDField = oDB("IDField")
aResults = Array(sMdb, sCfg, sTable)

oIni.Text(oDB("Table"), "ParentTable") = ""
oIni.Text(oDB("Table"), "ParentIDField") = ""
oIni.Text(oDB("Table"), "ParentIDValue") = ""
oIni.Text(oDB("Table"), "ParentFilter") = ""

If sButton <> "Cancel" Then
s = ""
If oHomer.DBIsRecord(oRS.AbsolutePosition) Then s = CStr(oRS.Fields(sIDField))
oIni.Text(oDB("Table"), "IDValue") = s

If oHomer.DBIsRecord(oRS.AbsolutePosition) Then
aFields = oDB("aResultFields")
For Each sField In aFields
sValue = oRS.Fields(sField).Value
oHomer.ArrayAdd aResults, sValue
Next
End If
End If
DBClose = aResults

On Error Resume Next
Set oConnection= oRS.ActiveConnection
oRS.Close
Set oRS = Nothing
oConnection.Close
Set oConnection = Nothing
On Error GoTo 0
oDlg.Close
End Function

Function DBCreateTable(sMdb, sCfg, sTable)
Dim aTables
Dim bResult
Dim oConnection, oDB
Dim sSql, sCreateTable

Set oDB = oHomer.DBGetConfiguration(sMdb, sCfg, sTable)
Set oConnection = oHomer.DBOpenConnection(sMdb)
sCreateTable = oDB("CreateTable")
sSQL = "create table " & sTable & " (" & sCreateTable & ")"
bResult = DBExecuteCommand(oConnection, sSQL, False)
oConnection.Close
Set oConnection = Nothing
Set oDB = Nothing

aTables = oHomer.DBGetTableNames(sMdb)
' msgbox join(atables, vblf), 0, "tables"
DBCreateTable = oHomer.ArrayContains(aTables, sTable, True)
End Function

Function DBCreateMDB(sMdb)
Dim oCatalog
Dim sConnection

dbCreateMDB = False
If Not oHomer.FileDelete(sMdb) Then Exit Function
sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
Set oCatalog = CreateObject("ADOX.Catalog")
oCatalog.Create sConnection
Set oCatalog = Nothing
DBCreateMDB = oHomer.FileExists(sMdb)
End Function

Function DBEval(oRS, sCode)
Dim aParts
Dim s, sPart
Dim v

With oRS
aParts = Split(sCode, ";")
v = vbNull
For Each sPart In aParts
s = oHomer.RegExpReplace(sPart, "\$([_a-zA-Z0-9]+)", ".Fields(" & Chr(34) & "$1" & Chr(34) & ").Value", True)
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
set oCon = oHomer.DBOpenConnection("DBDialog") ' added by Chip
end if
sText = ""
Set vResult = Nothing
On Error Resume Next
Set vResult = oCon.Execute(sCommand, iResult, adCmdText)
On Error GoTo 0
If Err.Number Or vResult Is Nothing Then
sTitle = "Error"
sText = sText & oHomer.DBGetErrorText(oCon) & vbCrLf
DBExecuteCommand = False
Else
sTitle = "Done"
sResult = ""
On Error Resume Next
' sResult = vResult.GetString(adClipString, vResult.RecordCount, vbTab, vbCrLf, "")
' sResult = vResult.GetString(adClipString, vbNull, vbTab, vbCrLf, "")
sResult = vResult.GetString
sResult = oHomer.StringConvertToWinLineBreak(sResult)
vResult.Close
On Error GoTo 0
If IsNumeric(iResult) And oHomer.Num(iResult) <> 0 Then
' sText = sText & "Affected " & oHomer.StringPlural("record", iResult) & vbCrLf
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

' oHomer.DialogShow "Results", sText
' If sTitle = "Error" Or bShowDone Then oHomer.DialogShow sTitle, sText
If sTitle = "Error" Or bShowDone Then oHomer.DialogMemo sTitle, "", sText, True
End Function

Function DBGetConfiguration(sMdb, sCfg, sTable)
Dim a, aKeys, aNames, aParts, aLabels, aTypes
Dim i, iIndex
Dim oDB, oIni, oField2Label, oField2Type
Dim s, sValue, sCreateTable, sPart, sKeys, sKey, sName, sType, sAlias

Set oDB = oHomer.CreateDictionary
Set DBGetConfiguration = oDB
If Not oHomer.FileExists(sCfg) Then Exit Function

oDB("MDB") = sMDB
oDB("File") = oHomer.PathGetName(sMdb)
oDB("Folder") = oHomer.PathGetFolder(sMdb)
oDB("Cfg") = sCfg
Set oIni = IniFile(sCfg)
oDB("Table") = sTable

sAlias = oIni.Text(sTable, "Alias", sTable)
oDB("Alias") = sAlias
sKeys = oIni.GetSectionKeys(sAlias)
aKeys = Split(sKeys, vbNullChar)
For Each sKey in aKeys
sValue = oIni.Text(sAlias, sKey, "")
oDB(sKey) = sValue
If oHomer.StringTrail(sKey, "Fields", True) Then oDB("a" & sKey) = oHomer.DBArray(sValue)
Next

If Len(oDB("Title")) = 0 Then oDB("Title") = oDB("Table") & " Records"
If IsEmpty(oDB("Filter")) Or CSTR(oDB("Filter")) = "0" Then oDB("Filter") = ""
If IsEmpty(oDB("IndexFields")) Or CSTR(oDB("IndexFields")) = "0" Then
oDB("IndexFields") = ""
oDB("aIndexFields") = Array()
End If

If Len(oDB("Provider")) = 0 Then oDB("Provider") = "Microsoft.JET.OLEDB.4.0"
oDB("ConnectionString") = "Provider=" & oDB("Provider") & ";Data Source=" & oDB("MDB")

sCreateTable = oDB("CreateTable")
oHomer.DBGetFieldNamesAndTypes sCreateTable, aNames, aTypes
oDb("aFieldNames") = aNames
oDb("FieldNames") = oHomer.DBString(oDB("aFieldNames"))
If Len(oDB("FieldLabels")) = 0 Then
oDB("aFieldLabels") = oHomer.DBGetFieldLabels(aNames)
oDB("FieldLabels") = oHomer.DBString(oDB("AFieldLabels"))
Else
oDB("aFieldLabels") = oHomer.DBArray(oDB("FieldLabels"))
End If

oDb("aFieldTypes") = aTypes
oDb("FieldTypes") = oHomer.DBString(oDB("aFieldTypes"))

aLabels = oHomer.DBArray(oDB("FieldLabels"))
Set oField2Label = oHomer.CreateDictionary
For i = 0 To UBound(aNames)
oField2Label.Add aNames(i), aLabels(i)
Next
oDB.Add "Field2Label", oField2Label

oHomer.ArrayClear aLabels
For Each sName In oDB("AInputFields")
oHomer.ArrayAdd aLabels, oField2Label(sName)
Next
oDB("aInputLabels") = aLabels
oDB("InputLabels") = oHomer.DBString(aLabels)

Set oField2Type = oHomer.CreateDictionary
aNames = oHomer.DBArray(oDB("FieldNames"))
For i = 0 To UBound(aNames)
oField2Type.Add aNames(i), aTypes(i)
Next
oDB.Add "Field2Type", oField2Type

oHomer.ArrayClear aTypes
For Each sName In oDB("aInputFields")
oHomer.ArrayAdd aTypes, oField2Type(sName)
Next
oDB("aInputTypes") = aTypes
oDB("InputTypes") = oHomer.DBString(aTypes)

Set DBGetConfiguration = oDB
End Function

' Public Function DBGetConfiguredTables(oDB, bIncludeCurrent)
Function DBGetConfiguredTables(sCfg, sExcludeTable)
Dim aTables, aNames
Dim iIndex
Dim oIni
Dim sTables, sTable, sNames, sName

' sTable = oDB("Table")
' sCfg = oDB("Cfg")
Set oIni = IniFile(sCfg)
sNames = oIni.GetSectionNames
aNames = Split(sNames, vbNullChar)
oHomer.ArrayClear aTables
For Each sName In aNames
If oHomer.StringEquiv(sName, "Configuration") Then
' Do Nothing
ElseIf oHomer.StringEquiv(sName, "Hotkeys") Then
' Do Nothing
ElseIf oHomer.StringLead(sName, "Transfer ", True) Then
' Do nothing
ElseIf oHomer.StringLead(sName, "Report ", True) Then
' Do nothing
' ElseIf Not bIncludeCurrent And oHomer.StringEquiv(sName, sTable) Then
ElseIf oHomer.StringEquiv(sName, sExcludeTable) Then
' Do nothing
Else
oHomer.ArrayAdd aTables, sName
End If
Next
DBGetConfiguredTables = aTables
End Function

Function DBOpenRecordSet(oConnection, sTable)
Dim oRS

Set oRS = CreateObject("AdoDb.RecordSet")
oRS.CursorLocation = adUseServer
oRS.Open sTable, oConnection, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
Set DBOpenRecordSet = oRS
End Function

Function DBTransfer(oDB, oConnection, oSourceRS, oTargetRS, bExport, bCurrentRecord)
Dim i, j, iOldBookmark
Dim aSections, aParts, aTransfers, aTables
Dim oIni, source, target
Dim s, sTable, sIDField, sField, sCfg, sAlias, sSections, sParts, sPart, sSourceTable, sTargetTable, sSection
Dim v

Set oIni = IniFile(oDB("Cfg"))
sSections = oIni.GetSectionNames
aSections = Split(sSections, vbNullChar)
oHomer.ArrayClear aTables
oHomer.ArrayClear aTransfers
For Each sSection In aSections
sAlias = oIni.Text(sSection, "Alias", "")
aParts = Split(sSection, " ")
If aParts(0) = "Transfer" And Not oHomer.StringEquiv(sSection, "Transfer") Then
sSourceTable = aParts(1)
sTargetTable = aParts(2)
If bExport And (oHomer.StringEquiv(sSourceTable, oDB("Table")) Or oHomer.StringEquiv(sSourceTable, oDB("Alias"))) Then
oHomer.ArrayAdd aTables, sTargetTable
oHomer.ArrayAdd aTransfers, sSection
ElseIf Not bExport And (oHomer.StringEquiv(sTargetTable, oDB("Table")) Or oHomer.StringEquiv(sTargetTable, oDB("Alias"))) Then
oHomer.ArrayAdd aTables, sSourceTable
oHomer.ArrayAdd aTransfers, sSection
End If
ElseIf Not oHomer.StringEquiv(oDB("Table"), sSection) And (oHomer.StringEquiv(oDB("Table"), sAlias) Or oHomer.StringEquiv(oDB("Alias"), sAlias) Or oHomer.StringEquiv(oDB("Alias"), sSection)) Then
oHomer.ArrayAdd aTables, sSection
oHomer.ArrayAdd aTransfers, sSection
End If
Next
If oHomer.ArrayBound(aTables) < 0 Then
Speak "No transfer settings found!"
Exit Function
End If

sSection = oHomer.DialogPick("Pick Table", aTables, aTransfers, True)
If Len(sSection) = 0 Then Exit Function

aParts = Split(sSection, " ")
If aParts(0) = "Transfer" Then
If bExport Then
sTable = aParts(2)
Else
sTable = aParts(1)
End If
Else
sTable = sSection
sSection = ""
End If

sAlias = oIni.Text(sTable, "Alias", sTable)
sIDField = oIni.Text(sAlias, "IDField", sTable & "_ID")

If bExport Then
Set oTargetRS = oHomer.DBOpenRecordSet(oConnection, sTable)
Else
Set oSourceRS = oHomer.DBOpenRecordSet(oConnection, sTable)
End If
Set source = oSourceRS
Set target = oTargetRS

Speak "Please wait"
j = 0
iOldBookmark = 0
If oHomer.DBIsRecord(oSourceRS.AbsolutePosition) Then iOldBookmark = oSourceRS.Bookmark
If not bCurrentRecord Then oSourceRS.MoveFirst
Do While Not oSourceRS.Eof
j = j + 1
If not bCurrentRecord Then Speak j
oTargetRS.AddNew
For i = 0 To oTargetRS.Fields.Count - 1
sField = oTargetRS.Fields(i).Name
If Len(sSection) = 0 Then
If oHomer.StringEquiv(sField, sIDField) Then
v = vbNull
Else
v = oSourceRS.Fields(sField).Value
End If
Else
sParts = oIni.Text(sSection, sField, "")
If Len(sParts) > 0 Then
v = oHomer.DBEval(oSourceRS, sParts)
End If
End If
On Error Resume Next
If v <> vbNull Then oTargetRS.Fields(sField).Value = v
On Error GoTo 0
Next
oTargetRS.Update
If bCurrentRecord Then Exit Do
oSourceRS.MoveNext
Loop
Speak "Done!"
If iOldBookmark <> 0 Then oSourceRS.Bookmark = iOldBookmark

If bExport Then
oTargetRS.Close
Set oTargetRS = Nothing
Else
oSourceRS.Close
Set oSourceRS = Nothing
End If
End Function

Function DBReport(oDB, oSourceRS)
Dim i, j, iOldBookmark
Dim aLines, aSections, aParts, aReports, aTables
Dim oIni, source, target
Dim s, sValue, sTable, sIDField, sField, sCfg, sAlias, sSections, sParts, sPart, sSourceTable, sTargetTable, sSection, sText, sLine, sLines, sFile, sExe, sCommand
Dim v

Set oIni = IniFile(oDB("Cfg"))
sSections = oIni.GetSectionNames
aSections = Split(sSections, vbNullChar)
oHomer.ArrayClear aTables
oHomer.ArrayClear aReports
For Each sSection In aSections
sAlias = oIni.Text(sSection, "Alias", "")
aParts = Split(sSection, " ")
If aParts(0) = "Report" Then
sSourceTable = aParts(1)
sTargetTable = aParts(2)
If oHomer.StringEquiv(sSourceTable, oDB("Table")) Or oHomer.StringEquiv(sSourceTable, oDB("Alias")) Then
oHomer.ArrayAdd aTables, sTargetTable
oHomer.ArrayAdd aReports, sSection
End If
End If
Next
If oHomer.ArrayBound(aTables) < 0 Then
Speak "No report settings found!"
Exit Function
End If

sSection = oHomer.DialogPick("Pick Report", aTables, aReports, True)
If Len(sSection) = 0 Then Exit Function

aParts = Split(sSection, " ")
sTable = aParts(2)

sAlias = oIni.Text(sTable, "Alias", sTable)
sIDField = oIni.Text(sAlias, "IDField", sTable & "_ID")

Set source = oSourceRS

Speak "Please wait"
j = 0
iOldBookmark = 0
If oHomer.DBIsRecord(oSourceRS.AbsolutePosition) Then iOldBookmark = oSourceRS.Bookmark
sText = ""
oSourceRS.MoveFirst
sValue = oIni.Text(sSection, "Top")
v = oHomer.DBEval(oSourceRS, sValue)
If sValue = "Blank" Or Len(v) > 0 Then sText = sText & v & vbCrLf
Do While Not oSourceRS.Eof
j = j + 1
Speak j
sLines = oIni.GetSectionKeys(sSection)
aLines = Split(sLines, vbNullChar)
For Each sLine In aLines
If oHomer.StringLead(sLine, "Line", True) Then
sValue = oIni.Text(sSection, sLine, "")
v = oHomer.DBEval(oSourceRS, sValue)
If sValue = "Blank" Or Len(v) > 0 Then sText = sText & v & vbCrLf
End If
Next
sValue = oIni.Text(sSection, "Record")
v = oHomer.DBEval(oSourceRS, sValue)
If sValue = "Blank" Or Len(v) > 0 Then sText = sText & v & vbCrLf
oSourceRS.MoveNext
Loop
sValue = oIni.Text(sSection, "Bottom")
v = oHomer.DBEval(oSourceRS, sValue)
If sValue = "Blank" Or Len(v) > 0 Then sText = sText & v & vbCrLf
Speak "Done!"
If iOldBookmark <> 0 Then oSourceRS.Bookmark = iOldBookmark
sFile = oHomer.PathCombine(ClientInformation.ScriptPath, "DbDialog.tmp")
oHomer.StringToFile sHomer, sFile
sFile = oHomer.PathCombine(ClientInformation.ScriptPath, "DbDialog.txt")
oHomer.StringToFile sText, sFile
sExe = oHomer.WinEyesGetScriptEditor
oHomer.ShellOpenWith sExe, sFile
End Function

Function DBIsRecord(iPosition)
DBIsRecord = False
If iPosition = AdPosBof Then Exit Function
If iPosition = AdPosEof Then Exit Function
If iPosition = AdPosUnknown Then Exit Function
DBIsRecord = True
End Function

Function DBIsValid(aResults)
DBIsValid = False
If Not IsArray(aResults) Then Exit Function
If UBound(aResults) <> 1 Then Exit Function
DBIsValid = aResults(0) <> 0 And aResults(1) <> 0
End Function

Function DBOpenConnection(sMdb)
Dim oConnection
Dim s, sConnection

' sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
Set oConnection = CreateObject("AdoDB.Connection")
' oConnection.Open sConnection
' oConnection.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=C:\Documents and Settings\Owner.DWEEB1\Application Data\GW Micro\Window-Eyes\users\default\DbDialog.mdb"
s = oHomer.PathCombine(ClientInformation.ScriptPath, "DbDialog.mdb")
oConnection.Open "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & s
Set DBOpenConnection = oConnection
End Function


Function DBGetCondition(sField, sType, sValue)
Dim b
Dim s, sPrefix, sRest

sPrefix = ""
sRest = sValue
If oHomer.StringLead(sRest, "<>", False) Or oHomer.StringLead(sRest, ">=", False) Or oHomer.StringLead(sRest, "<=", False) Then
sPrefix = Left(sRest, 2)
sRest = StringChopLeft(sRest, 2)
ElseIf oHomer.StringLead(sRest, "=", False) Or oHomer.StringLead(sRest, ">", False) Or oHomer.StringLead(sRest, "<", False) Then
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

oHomer.ArrayClear aNames
oHomer.ArrayClear aTypes
aParts = oHomer.DBArray(sCreateTable)
For Each sPart In aParts
a = Split(sPart, " ")
If UBound(a) >= 1 Then
oHomer.ArrayAdd aNames, a(0)
sType = a(1)
iIndex = InStr(sType, "(")
If iIndex Then sType = Left(sType, iIndex - 1)
oHomer.ArrayAdd aTypes, sType
End If
Next
End Function

Function DBGetIndexNames(sMdb, sTable)
Dim aIndexes
Dim oCatalog, oTables, oTable, oIndexes, oIndex, oConnection
Dim sConnection, sIndex

DBGetIndexNames = Array()
If Not oHomer.FileExists (sMdb) Then Exit Function

oHomer.ArrayClear aIndexes
sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
Set oCatalog = CreateObject("ADOX.Catalog")
Set oConnection = oHomer.DBOpenConnection(sMdb)
Set oCatalog.ActiveConnection = oConnection
Set oTables = oCatalog.Tables
For Each oTable in oTables
If StrComp(sTable, oTable.Name) = 0 Then
Set oIndexes = oTable.Indexes
For Each oIndex In oIndexes
sIndex = oIndex.Name
oHomer.ArrayAdd aIndexes, sIndex
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

oHomer.ArrayClear aLabels
' For i = 0 To oRS.Fields.Count - 1
' sField = oRS.Fields(i)
For Each sField in aFields
sLabel = Replace(sField, "_", " ")
If InStr(sLabel, "&") = 0 Then sLabel = "&" & sLabel
oHomer.ArrayAdd aLabels, sLabel
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
set oConnection = oHomer.DBOpenConnection("DBDialog") ' added by Chip
end if
Set oIni = IniFile(oDB("Cfg"))
sFields = oIni.Text(sTable, "ResultFields", "")
aFields = oHomer.DBArray(sFields)
oHomer.ArrayRemove aFields, 0
sFields = oHomer.DBString(aFields)
If Len(sFields) = 0 Then Exit Function

sSQL = "select " & sFields & " from " & sTable & " Where " & sIDField & " = " & iID
' MsgBox SSQL
Set oRS = oConnection.Execute(sSQL)
s = oRS.GetString
s = oHomer.StringConvertToWinLineBreak(s)
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
' If oHomer.DBIsRecord(oRS.AbsolutePosition) Then DBGetString = oRS.GetString(adClipString, oRS.RecordCount, vbTab, vbCrLf, "")
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
If Not oHomer.FileExists (sMdb) Then Exit Function

oHomer.ArrayClear aTables
sConnection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & sMdb
Set oCatalog = CreateObject("ADOX.Catalog")
Set oConnection = oHomer.DBOpenConnection(sMdb)
Set oCatalog.ActiveConnection = oConnection
Set oTables = oCatalog.Tables
For Each oTable In oTables
sTable = oTable.Name
oHomer.ArrayAdd aTables, sTable
Next
oConnection.Close
Set oConnection = Nothing
Set oCatalog = Nothing
Set oTables = Nothing
Set oTable = Nothing
DBGetTableNames = aTables
End Function

Function DBPopulateControls(oDB, oRS, oDlg, sButton)
Dim aFields, aSelectFields, aIndexes
Dim iIndex, iID
Dim oCon, oIni, oLst, oLbl
Dim s, sSQL, sField, sIDField, sText, sIndex, sFilter

Set oIni = IniFile(oDB("Cfg"))
Set oLst = oDlg.Control("lst")
Set oLbl = oDlg.Control("lbl")
aSelectFields = oDB("aSelectFields")
sIDField = oDB("IDField")
Set oCon = oRS.ActiveConnection
if oCon is nothing then
' speak "empty connection found in DBPopulateControls"
set oCon = oHomer.DBOpenConnection("DBDialog") ' added by Chip
end if


oLst.Clear
aIndexes = oHomer.DBGetIndexNames(oDB("Mdb"), oDB("Table"))
aFields = oDB("aIndexFields")
oHomer.ArrayInsert aFields, sIDField, 0
For Each sField In aFields
If oHomer.ArrayContains(aIndexes, sField, True) Then
' Do nothing
Else
Speak "Creating index " & sField
sSQL = "Create Index " & sField & " On " & oDB("Table") & " (" & sField & ")"
oHomer.DBExecuteCommand oCon, sSQL, False
End If
Next

oRS.Sort = oDB("IndexFields")
' s = oHomer.Str(oDB("IndexFields"))
' oHomer.DialogShow "indexfields", s
' oRS.Sort = s
sFilter = oDB("ParentFilter")
If IsEmpty(sFilter) Or CSTR(sFilter) = "0" Then sFilter = ""
s = oDB("Filter")
If CStr(s) = "0" Then s = ""
If Len(sFilter) > 0 And Len(s) > 0 Then
sFilter = sFilter & " And " & s
ElseIf Len(s) > 0 Then
sFilter = s
End If

oRS.Filter = sFilter
If Len(sFilter) > 0 Then Speak "With Filter"
If Not oHomer.DBIsRecord(oRS.AbsolutePosition) Then
Speak "No items!"
Else
oRS.MoveFirst
Do While Not oRS.Eof
speak "found one" ' temporary
sText = oHomer.DBGetSelectText(oDB, oRS, oDlg, sButton)
iID = oRS.Fields(sIDField).Value
oLst.Add sText, iID
oRS.MoveNext
Loop

s = oIni.Text(oDB("Table"), "IDValue", "")
If Len(s) = 0 Then
oRS.MoveFirst
Else
s = sIDField & " = " & s
oRS.Find s
If oRS.Eof Then oRS.MoveFirst
End If
oHomer.DBScrollListToTable oDB, oRS, oDlg, sButton
End If
oHomer.DBSetStatus oDB, oRS, oDlg, sButton
End Function

Function DBGetInputFields(oDB, oRS, oDlg, sButton)
Dim aInputLabels, aInputValues, aInputTypes, aInputFields, aResults, aFields
Dim i, iIndex, iID, iBookmark, iParentID
Dim oConnection, oIni
Dim s, sLook, sField, sType, sValue, sIDField, sParentIDField, sTable, sFilter, sConnection, sIndex, sFields
Dim vValue

DBGetInputFields = Array()
' Set oConnection = oRS.ActiveConnection
aInputLabels = oDB("aInputLabels")
aInputFields = oDB("aInputFields")
aInputtypes = oDB("aInputTypes")
sParentIDField = oDB("ParentIDField")
sLook = ""

oHomer.ArrayClear aInputValues
Select Case sButton
Case "Add"
For Each sField in aInputFields
If Len(sParentIDField) = 0 Then
sValue = ""
ElseIf oHomer.StringEquiv(sField, sParentIDField) Then
iParentID = 0
s = oDB("ParentIDValue")
If Len(s) > 0 Then iParentID = oHomer.Num(s)
sValue = CStr(iParentID)
sTable = oDB("ParentTable")
Set oIni = IniFile(oDB("Cfg"))
sFields = oIni.Text(sTable, "ResultFields", "")
aFields = oHomer.DBArray(sFields)
oHomer.ArrayRemove aFields, 0
sFields = oHomer.DBString(aFields)
sFilter = oDB("ParentFilter")
sConnection = oDB("ConnectionString")
sIndex = ""
' sLook = oHomer.DBGetLookText(oDB, oRS, sTable, sParentIDField, iParentID)
sLook = DBGetString(sConnection, sTable, sFields, sFilter, sIndex)
ElseIf Len(sLook) > 0 Then
sValue = sLook
sLook = ""
Else
sValue = ""
End If
oHomer.ArrayAdd aInputValues, sValue
Next
Case "Filter", "Modify Batch"
For Each sField in aInputFields
oHomer.ArrayAdd aInputValues, ""
Next
Case "Copy", "Modify", "View"
For i = 0 To UBound(aInputFields)
sField = aInputFields(i)
vValue = oRS.Fields(sField).Value
sType = aInputTypes(i)
sValue = oHomer.DBVariantToString(vValue)
oHomer.ArrayAdd aInputValues, sValue
Next
End Select
' DBGetInputFields = oHomer.DialogMultiInput(sButton, aInputLabels, aInputValues)
DBGetInputFields = oHomer.DialogDBInput(sButton, aInputLabels, aInputValues, oDB)
End Function

Function DBGetRecordSet(oDB)
Dim oConnection, oRS
Dim sTable

Set oDialogConnection = oHomer.DBOpenConnection(oDB)
sTable = oDB("Table")
Set oRS = CreateObject("Adodb.RecordSet")
oRS.CursorLocation = adUseClient
oRS.Open sTable, oDialogConnection, adOpenKeySet, adLockOptimistic, adCmdTableDirect
Set DBGetRecordSet = oRS
End Function

Function DBGetSelectText(oDB, oRS, oDlg, sButton)
Dim aSelectFields, aResults
Dim iID, iIndex, iBookmark
Dim sText, sField, sIDField

DBGetSelectText = ""
If Not oHomer.DBIsRecord(oRS.AbsolutePosition) Then Exit Function

aSelectFields = oDB("aSelectFields")
sText = ""
For Each sField In aSelectFields
If Len(sText) > 0 Then sText = sText & vbTab
sText = sText & oRS.Fields(sField).Value
Next
DBGetSelectText = sText
End Function

Function DBSaveInputFields(oDB, oRS, oDlg, sButton, aResults)
Dim aInputFields, aInputTypes
Dim i, iIndex, iID, iBookmark
Dim oCon
Dim sField, sResult, sType, sIDField, sText, sTitle
Dim vResult

DBSaveInputFields = Array(0,0)
Set oCon = oRS.ActiveConnection
if oCon is nothing then
' speak "empty connection found in DBSaveInputFields"
set oCon = oHomer.DBOpenConnection("DBDialog") ' added by Chip
end if
aInputFields = oDB("AInputFields")
aInputtypes = oDB("aInputtypes")

Select Case sButton
Case "Add", "Copy"
oRS.AddNew
oRS.Fields("Added").Value = Date
oRS.Fields("Tagged").Value = " "
Case "Modify"
If Not oHomer.DBIsRecord(oRS.AbsolutePosition) Then Exit Function
oRS.Fields("Modified").Value = Date
End Select

For i = 0 To UBound(aResults)
sResult = aResults(i)
sType = aInputTypes(i)
If sButton = "Modify Batch" And Len(sResult) = 0 Then
' Do nothing
ElseIf sType <> "Null" Then
vResult = oHomer.DBStringToVariant(sResult, sType)
sField = aInputFields(i)
oRS.Fields(sField).Value = vResult
End If
Next
On Error Resume Next
oRS.Update
On Error GoTo 0
If Err.Number Then
sTitle = "Error"
sText = Err.Description & vbCrLf
sText = sText & oHomer.DBGetErrorText(oCon) & vbCrLf
oHomer.DialogShow sTitle, sText
End If

Select Case sButton
Case "Add", "Copy"
DBSaveInputFields = oHomer.DBAddItemFromRecord(oDB, oRS, oDlg, sButton)
Case "Modify"
DBSaveInputFields = oHomer.DBSetItemFromRecord(oDB, oRS, oDlg, sButton)
End Select
End Function

Function DBSaveOption(oDB, sKey, vValue)
Dim oIni
Dim sSection, sCfg, sValue

If IsArray(vValue) Then
oDB("a" & sKey) = vValue
sValue = oHomer.DBString(vValue)
Else
sValue = vValue
End If
oDB(sKey) = sValue
sCfg = oDB("Cfg")
sSection = oDB("Table")
Set oIni = IniFile(sCfg)
' sValue = oHomer.StringQuote(sValue)
oIni.Text(sSection, sKey) = sValue
End Function

Function DBSetItemFromRecord(oDB, oRS, oDlg, sButton)
Dim aResults
Dim IID, iIndex, iBookmark
Dim oLst
Dim sText, sIDField

DBSetItemFromRecord = Array(0,0)
If Not oHomer.DBIsRecord(oRS.AbsolutePosition) Then Exit Function

Set oLst = oDlg.Control("lst")
sIDField = oDB("IDField")
iID = oRS.Fields(sIDField).Value
aResults = oHomer.DBScrollListToTable(oDB, oRS, oDlg, sButton)
If Not oHomer.DBIsValid(aResults) Then Exit Function

sText = oHomer.DBGetSelectText(oDB, oRS, oDlg, sButton)
oLst.Text(iIndex) = sText
iIndex = aResults(1)
oLst.FocusedIndex = iIndex
oHomer.DBSetStatus oDB, oRS, oDlg, sButton
DBSetItemFromRecord = Array(iID, iIndex)
End Function

Function DBScrollListToTable(oDB, oRS, oDlg, sButton)
Dim aResults
Dim iIndex, iID, iBookmark
Dim oLst
Dim sIDField

DBScrollListToTable = Array(0,0)
If Not oHomer.DBIsRecord(oRS.AbsolutePosition) Then Exit Function

Set oLst = oDlg.Control("lst")
sIDField = oDB("IDField")
iID = oRS.Fields(sIDField).Value
iIndex = oLst.FindData(iID)
If iIndex > 0 Then
oLst.FocusedIndex = iIndex
DBScrollListToTable = Array(iID, iIndex)
End If
' speak "id " & iID
' speak "index " & iindex
End Function

Function DBScrollTableToList(oDB, oRS, oDlg, sButton)
Dim aResults
Dim iIndex, iID, iBookmark
Dim oLst
Dim s, sIDField

DBScrollTableToList = Array(0,0)
Set oLst = oDlg.Control("lst")
iIndex = oLst.FocusedIndex
If iIndex = 0 Then Exit Function
If oRS.RecordCount <= 0 Then Exit Function

iID = oLst.Data(iIndex)
sIDField = oDB("IDField")
s = sIDField & " = " & IID
iBookmark = 0
If oHomer.DBIsRecord(oRS.AbsolutePosition) Then iBookmark = oRS.Bookmark
oRS.MoveFirst
oRS.Find s, 0, adSearchForward, adBookmarkFirst
If oRS.Eof Then
If iBookmark <> 0 Then oRS.Bookmark = iBookmark
Exit Function
End If
DBScrollTableToList = Array(iID, iIndex)
End Function

Function DBSetStatus(oDB, oRS, oDlg, sButton)
Dim aStatusFields, aResults
Dim i, iID, iIndex, iBookmark
Dim oLbl, oField2Label
Dim sField, sText, sValue, sIDField

DBSetStatus = ""
Set oLbl = oDlg.Control("lbl")
If oRS.Eof Then
sText = "End of File"
ElseIf oRS.Bof Then
sText = "Beginning of File"
ElseIf oRS.AbsolutePosition = adPosUnknown Then
sText = "Unknown record"
Else
aStatusFields = oDB("AStatusFields")
Set oField2Label = oDB("Field2Label")
sText = ""
For Each sField In aStatusFields
sValue = "" & oRS.Fields(sField).Value & vbTab
If sValue <> vbTab Then sText = sText & oField2Label(sField) & ": " & sValue
Next
End If
oLbl.Text = sText
DBSetStatus = sText
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

dim iArgCount, iArg
Dim oRs, oFile, oTable, oConnect, oSystem
dim sSql, sTable, sConnectString, sBinDir, sCurDir, sDir, sFile, sHomerLibVbs, sIniDir, sPaxDb, sPaxDbBase, sScriptVbs, sSQLite3Exe, sTempDir, sTempTmp, sWildcards

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
sBinDir = PathGetFolder(WScript.ScriptFullName)
sIniDir = StringChopRight(sBinDir, 3) + "ini"
sPaxDbBase = "Pax.db"
sPaxDb = PathCombine(sIniDir, sPaxDbBase)
sConnectString = "DRIVER=SQLite3 ODBC Driver;Database=" & sPaxDb & ";"
sSQLite3Exe = PathCombine(sTempDir, "SQLite3.exe")

Set oConnect = CreateObject("ADODB.Connection")
sTable = "rules"
sSql = "select * from rules"

set oTable = CreateObject("ADODB.Recordset")
' oRS.CursorLocation = adUseServer
' oRS.Open sTable, oConnection, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
' oTable.CursorLocation = adUseClient
print 10
' oTable.Open sTable, oConnect, adOpenKeySet, adLockOptimistic, adCmdTableDirect
oTable.Open sSql, oConnect
print 11
Do Until oTable.EOF
' Do While Not(oTable.EOF)
wscript.echo oTable.Fields(1)
WScript.Echo oTable("name").Value
oTable.MoveNext
Loop
oTable.Close
