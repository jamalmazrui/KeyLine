Option Explicit
WScript.Echo"Starting Dot"

Function FileInclude(sFile)
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

' Main
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
' print sConnectString
oConnect.Open sConnectString
' oTable.Open sTable, oConnect
oTable.Open sSql, oConnect
Do Until oTable.EOF
' Do While Not(oTable.EOF)
wscript.echo oTable.Fields(1)
WScript.Echo oTable("name").Value
oTable.MoveNext
Loop
oTable.Close
