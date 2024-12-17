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
Dim oField, oFields, oRs, oFile, oTable, oConnect, oSystem
dim sInput, sSql, sTable, sConnectString, sBinDir, sCurDir, sDir, sFile, sHomerLibVbs, sIniDir, sPaxDb, sPaxDbBase, sScriptVbs, sSQLite3Exe, sTempDir, sTempTmp, sWildcards

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
' print sConnectString
oConnect.Open sConnectString
' oTable.Open sTable, oConnect
oTable.CursorLocation = adUseClient
' oTable.Open sTable, oConnect, adOpenKeySet, adLockOptimistic, adCmdTableDirect
oTable.Open sSql, oConnect

for each oField in oTable.Fields
print oField.Name
Next

' Do While Not(oTable.EOF)
' Do Until oTable.EOF
' print oTable("look").Value
' oTable.MoveNext
' Loop

do while true
sInput = cmdPrompt("Dot: ")
' executeGlobal sInput
on error resume next
execute sInput
on error goto 0
print "row " & oTable.AbsolutePosition
loop
oTable.Close
