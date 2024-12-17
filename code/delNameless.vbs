Option Explicit
WScript.Echo"Starting delNameless"

Dim aDelFiles, a, aSelectFiles, aFiles
Dim bIgnoreCase, bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dDelFiles, dOutput, dProperties, d, dStyle, dIni, dSourceIni
Dim iMatchCount, iNewFolderCount, iNewFileCount, iStem, iDelFile, iBound, iFile, iFileCount, iError, iSuggestionCount, iErrorCount, iArg, iArgCount, iCount
Dim oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sNewRoot, sNewName, sNewFile, sDelFile, sStem, sDelStem, sDelRoot, sDelExt, sReplace, sExt, sRoot, sFlags, sWildcards, sChoice, sOutputTxt, sInputTxt, sSourceExe, sFiles, sSourcePart, sCommand, sSourceBase, sIniDir, sBinDir, sTempTmp, sSfkExe, sIniFormExe, sInputIni, sInputPart, sInputBase, sOutputIni, sOutputPart, sOutputBase, sTempDir, sCurDir, sScriptVbs, sHomerLibVbs, sDir, sFile, sTargetTxt, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

Const WindowStyle = 0 'hidden
Const HIDDEN = 0 ' window style
Const NORMAL = 1 ' window style
Const MAX = 3
Const Wait = True

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

sWildcards = "*.*"
If iArgCount > 0 Then sWildcards = WScript.Arguments(0)
sCurDir = PathGetCurrentDirectory()
sTempDir = PathGetSpecialFolder("TEMP")
sTempDir = ShellExpandEnvironmentVariables("%TEMP%")
sTempTmp = PathCombine(sTempDir, "temp.tmp")
sBinDir = PathGetFolder(WScript.ScriptFullName)
sIniDir = StringChopRight(sBinDir, 3) + "ini"
sSfkExe = PathCombine(sBinDir, "sfk.exe")
sSourceExe = PathCombine(sBinDir, "IniForm.exe")
sIniFormExe = PathCombine(sTempDir, "IniForm.exe")

sSourcePart = "PickFiles"
sSourceBase = sSourcePart + "_input.ini"
sSourceIni = PathCombine(sIniDir, sSourceBase)

sInputBase = "input.ini"
sInputIni = PathCombine(sTempDir, sInputBase)
sInputTxt = PathCombine(sTempDir, "input.txt")
sOutputTxt = PathCombine(sTempDir, "Output.txt")
sOutputBase = "Output.ini"
sOutputIni = PathCombine(sTempDir, sOutputBase)

sCommand = StringQuote(sSfkExe) + " dupfind . +delete"
sCommand = "cmd.exe /c " & sCommand
sCommand = sCommand & " >" & sTempTmp
' sText = ShellExec(sCommand)
sDir = PathGetCurrentDirectory()
sFlags = "/b"
aFiles = PathGetSpec(sDir, sWildcards, sFlags)
' print Join(aFiles, vbCrLf)
aDelFiles = Array()
Set dDelFiles = CreateDictionary
iBound = ArrayBound(aFiles)
iCount = iBound + 1
s = iCount & " matches"
If iCount = 1 Then s = "1 match"
print s
iNewFileCount = 0
iNewFolderCount = 0
For iFile = 0 to iBound
sFile = aFiles(iFile)
sDir = PathGetFolder(sFile)
sName = PathGetName(sFile)
sRoot = PathGetRoot(sFile)
sExt = PathGetExtension(sFile)
sMatch = "[a-zA-Z]"
sReplace = "0$1$2$3"
bIgnoreCase = True
iMatchCount = RegExpCount(sRoot, sMatch, bIgnoreCase)
' print sRoot & " = " & iMatchCount
If iMatchCount = 0 and FileExists(sFile) Then dDelFiles.Add sFile, ""
Next ' sFile in aFiles
If dDelFiles.Count > 0 Then print "Deleting"
For Each sFile in dDelFiles.Keys
sName = PathGetName(sFile)
print sName
FileDelete(sFile)
Next
print "Deleted " & StringPlural("file", dDelFiles.Count)
