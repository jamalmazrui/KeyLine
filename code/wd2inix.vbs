Option Explicit
WScript.Echo"Starting ini2tables"

Dim aIni
Dim bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dTable, dProperties, d, dStyle, dIni, dSourceIni
Dim iPad, iTable, iTableCount, iStartRow, iError, iSuggestionCount, iErrorCount, iArg, iArgCount, iCount
Dim oTable, oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sCommand, sBinDir, sTableIni, sExtensions, sFields, sRepeatFields, sScriptVbs, sHomerLibVbs, sDir, sFile, sTargetTxt, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

Dim WdDoNotSaveChanges: WdDoNotSaveChanges = 0

Dim msoPropertyTypeBoolean : msoPropertyTypeBoolean = 2
Dim msoPropertyTypeDate : msoPropertyTypeDate = 3
Dim msoPropertyTypeFloat : msoPropertyTypeFloat = 5
Dim msoPropertyTypeNumber : msoPropertyTypeNumber = 1 ' Integer
Dim msoPropertyTypeString : msoPropertyTypeString = 4

Function FileInclude(sFile)
' With CreateObject("Scripting.FileSystemObject")
' ExecuteGlobal .openTextFile(sFile).readAll()
' End With

executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
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
If iArgCount < 1 Then
Print "Specify a source .docx file as a parameter."
Quit
End If

sSourceDocx = WScript.Arguments(0)
If InStr(sSourceDocx, "\") = 0 Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If not FileExists(sSourceDocx) Then Quit "Cannot find " & sSourceDocx

bReadOnly = True
sTargetIni = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceDocx) + ".inix")

ProcessTerminateAllModule "WinWord"
Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = 0
oApp.ScreenUpdating = False
' Set oDoc = oApp.Documents.Open(sSourceDocx)
Set oDoc = oApp.Documents.Open(sSourceDocx, False, True, False)

iTableCount = oDoc.Tables.Count
iPad = Len(Trim(CStr(iTableCount)))
iTable = 0
For Each oTable in oDoc.Tables
Set dTable = Nothing
On Error Resume Next
Set dTable = TableToDictionary(oTable)
On Error goTo 0
If not dTable Is Nothing Then
iTable = iTable + 1
' DictionaryToIni dTable, sTargetIni
sText = DictionaryToString(dTable)
' sText = RegExpReplace(sText, "^\[.*\]$", "[]", False)
sText = RegExpReplace(sText, "(^|\n)\[.*\]", "$1[]", False)
sTableIni = sTargetIni
s = StringPadLeft(CStr(iTable), iPad, "0")
' If iTableCount > 1 Then sTableIni = PathCombine(PathGetFolder(sTargetIni), PathGetRoot(sTargetIni) & "-" & iTable & ".inix")
If iTableCount > 1 Then sTableIni = PathCombine(PathGetFolder(sTargetIni), PathGetRoot(sTargetIni) & "-" & s & ".inix")
StringToFile sText, sTableIni
' StringToFileUTF8 sText, sTableIni
sBinDir = PathGetFolder(WScript.ScriptFullName)
sCommand = StringQuote(PathCombine(sBinDir, "utf8b64.exe")) & " " & StringQuote(sTableIni)
' DialogShow sCommand, ""
' ShellWait sCommand
ShellRun sCommand, 0, True
' Exit For
End If
Next ' oTable

oDoc.Close 0
if not oApp.NormalTemplate.Saved Then oApp.NormalTemplate.Saved = True
oApp.Quit
rem echo "Done"
