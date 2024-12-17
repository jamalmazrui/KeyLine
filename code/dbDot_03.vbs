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

function ToString(x)
ToString = "null"
on error resume next
ToString = cstr(x)
on error goto 0
end function

function printPosition()
print "row " & oTable.AbsolutePosition
end function

function arrayCommand(aItems)
dim aReturn, aParts
dim iPart, iItem
dim sPart, sItem

aReturn = array()
aParts = split(aItems(0), " ")
for iPart = 0 to ubound(aParts)
arrayAdd aReturn, aParts(iPart)
next

for iItem = 1 to ubound(aItems)
arrayAdd aReturn, aItems(iItem)
next
arrayCommand = aReturn
end function

function arrayFromString(sItems)
dim aParams, aItems, aReturn
dim iItem
dim sItem

arrayFromString = aReturn
aItems = split(sItems, ",")
' for each sItem in aItems
aReturn = array()
for iItem = 0 to ubound(aItems)
sItem = trim(aItems(iItem))
if iItem = 0 then sItem = lcase(sItem)
if len(sItem) > 0 then arrayAdd aReturn, sItem
next
arrayFromString = aReturn
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

dim aParams
dim iParam, iJump, iArgCount, iArg
Dim oField, oFields, oRs, oFile, oTable, oConnect, oSystem
dim sParam, sInputRest, sValue, sFind, sField, sCommand, sInput, sSql, sTable, sConnectString, sBinDir, sCurDir, sDir, sFile, sHomerLibVbs, sIniDir, sPaxDb, sPaxDbBase, sScriptVbs, sSQLite3Exe, sTempDir, sTempTmp, sWildcards

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
sTable = "rules"
sSql = "select * from rules"

set oTable = CreateObject("ADODB.Recordset")
oConnect.Open sConnectString
oTable.CursorLocation = adUseClient
oTable.Open sTable, oConnect, adOpenKeySet, adLockOptimistic, adCmdTableDirect
' oTable.Open sSql, oConnect

for each oField in oTable.Fields
print oField.name & ", " & oField.value
Next

' Do While Not(oTable.EOF)
' Do Until oTable.EOF
' print oTable("look").Value
' oTable.MoveNext
' Loop

do while true
' sInput = cmdPrompt("Dot: ")
sInput = cmdPrompt(". ")

aParams = arrayFromString(sInput)
sCommand = ""
sInputRest = ""
if arrayCount(aParams) > 0 then  
aParams = arrayCommand(aParams)
sCommand = aParams(0)
for iParam = 1 to ubound(aParams)
sParam = aParams(iParam)
sInputRest = sInputRest & sParam
if iParam < ubound(aParams) then sInputRest = sInputRest & " "
next
end if
' print sCommand
select case sCommand
case ""
printPosition
case "count"
print oTable.recordCount
case "end"
oTable.moveLast
printPosition
case "find"
sFind = aParams(1)
print sFind
oTable.moveFirst
oTable.find sFind
case "goto", "go"
oTable.absolutePosition = cInt(aParams(1))
printPosition
case "jump"
iJump = cInt(aParams(1))
oTable.move iJump
printPosition
case "next"
oTable.moveNext
printPosition
case "previous", "prev"
oTable.movePrevious
printPosition
case "quit", "exit", "x"
print "Closing dbDot"
wscript.quit
case "exec", ";"
print sInputRest
on error resume next
set oRs = oConnect.Execute(sInputRest)
on error goto 0
print stringPlural("record", oRs.RecordCount)
case "show", "?"
sField = aParams(1)
' print oTable.Fields(sField).Value
sValue = ToString(oTable.Fields(sField).Value)
print sValue
case "start"
oTable.moveFirst
printPosition
case else
on error resume next
execute sInput
on error goto 0
end select
loop
oTable.Close
