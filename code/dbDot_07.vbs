 ' DSN={SQLite3 Datasource};Database=full-path-to-db;...
' RUNDLL32 sqlite3odbc.dll,install
' "PRAGMA table_info(...)" SQLite statement. If SELECTs are used which
' DSN={SQLite3 Datasource};Database=full-path-to-db;...
' Driver={SQLite3 ODBC Driver};Database=full-path-to-db;...
' Timeout (integer)	lock time out in milliseconds; default 100000
' SyncPragma (string)	value for PRAGMA SYNCHRONOUS; default empty
' C:\> RUNDLL32 [path]sqliteodbc.dll,install [quiet]
' C:\> RUNDLL32 [path]sqlite3odbc.dll,uninstall [quiet]

Option Explicit
WScript.Echo"Starting Dot"

Function FileInclude(sFile)
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

function endProgram()
print "Ending DbDot"
oTable.close
oConnect.close
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

function fillFieldArrays(aAllFields, aAddFields, aEditFields, aListFields, aNextFields, aViewFields)
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
aListFields = arrayCopy(aAddFields)
aNextFields = arrayCopy(aAllFields)
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

dim a, aArgs, aAllFields, aAddFields, aEditFields, aListFields, aNextFields, aViewFields, aParams
dim dTables, dTable
dim iLoop, iArgCount, iArg, iName, iPosition, iBookmark, iRecordsAffected, iParam, iJump
Dim oField, oFields, oRs, oFile, oTable, oConnect, oSystem
dim sInitial, sList, sIdField, sName, sView, sSeek, sRow, sSetting, sPosition, sSort, sFilter, sParam, sInputCmdRest, sInputParamRest, sValue, sFind, sField, sCommand, sInput, sSql, sTable, sConnectString, sBinDir, sCurDir, sDir, sFile, sHomerLibVbs, sIniDir, sPaxDb, sPaxDbBase, sScriptVbs, sSQLite3Exe, sTempDir, sTempTmp, sWildcards

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
' sConnectString = "DSN={SQLite3 Datasource};Database=" & sPaxDb &";"
sSQLite3Exe = PathCombine(sTempDir, "SQLite3.exe")

Set oConnect = CreateObject("ADODB.Connection")
oConnect.Open sConnectString
sTable = "rules"
set dTables = createDictionary
do while true
print "Table " & sTable
sIdField = mid(sTable, 1, len(sTable) - 1) & "_id"
set oTable = CreateObject("ADODB.Recordset")
oTable.CursorLocation = adUseClient
oTable.Open sTable, oConnect, adOpenKeySet, adLockOptimistic, adCmdTableDirect
fillFieldArrays aAllFields, aAddFields, aEditFields, aListFields, aNextFields, aViewFields
if dTables.exists(sTable) then
set oTable = dTables(sTable)("object")
else
set dTable = createDictionary
dTables.add sTable, dTable
dTable.add "idField", sIdField
dTable.add "allFields", aAllFields
dTable.add "addFields", aAddFields
dTable.add "editFields", aEditFields
dTable.add "listFields", aListFields
dTable.add "nextFields", aNextFields
dTable.add "viewFields", aViewFields
 
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
case "add"
oTable.addNew
a = aAddFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
for each sName in a
sValue = cmdPrompt(sName & ": ")
if len(sValue) > 0 then oTable.fields(sName) = sValue
next
case "edit", "/"
a = aEditFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
for each sName in a
sValue = "" & oTable.fields(sName)
' sValue = cmdPrompt(sName & ": ")
sValue = dialogInput("Input", sName & ": ", sValue)
if len(sValue) > 0 then oTable.fields(sName) = sValue
next
case "list"
a = aListFields
if len(sInputCmdRest) > 0 then a = stringToFields(sInputCmdRest)
iLoop = 0
do while not oTable.eof
sList = ""
for iName = 0 to arrayBound(a)
sName = a(iName)
sValue = "" & oTable.fields(sName)
sList = sList & sValue
if iName < arrayBound(a) then sList = sList & ", "
next
print sList
oTable.moveNext
iLoop = iLoop + 1
if iLoop = 10 then exit do
loop
case "view"
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
case "filter"
sFilter = sInputCmdRest
if sParam = "marked" then sFilter = "marked = true"
if sParam = "unmarked" then sFilter = "marked = false"
oTable.Filter = sFilter
print stringPlural("record", oTable.RecordCount)
case "find"
sFind = sInputCmdRest
oTable.find sFind
case "goto", "go"
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
case "next", "+"
iBookmark = oTable.bookMark
if len(sInputCmdRest) = 0 then
oTable.moveNext
else
do while not oTable.eof
oTable.moveNext
sRow = getRowString(aNextFields)
if stringContains(sRow, sInputCmdRest, true) then exit do
loop
end if
if oTable.eof then
print "Not found"
oTable.bookmark = iBookmark
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
case "view"
aViewFields = stringToFields(sInputParamRest)
case else
print "unknown setting"
end select
case "show", "?"
sName = sParam
sValue = "" & oTable.fields(sName)
print sValue
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
printRow
case else
on error resume next
execute sInput
on error goto 0
end select
loop
loop
endProgram
