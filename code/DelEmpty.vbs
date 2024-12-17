Option Explicit
WScript.Echo"Starting delEmpty"

Dim aDelFiles, a, aSelectFiles, aFiles
Dim bIgnoreCase, bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dDelFiles, dOutput, dProperties, d, dStyle, dIni, dSourceIni
Dim iSize, iStem, iDelFile, iBound, iFile, iFileCount, iError, iSuggestionCount, iErrorCount, iArg, iArgCount, iCount
Dim oFiles, oSystem, oDir, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sDelFile, sStem, sDelStem, sDelRoot, sDelExt, sReplace, sExt, sRoot, sFlags, sWildcards, sChoice, sOutputTxt, sInputTxt, sSourceExe, sFiles, sSourcePart, sCommand, sSourceBase, sIniDir, sBinDir, sTempTmp, sSfkExe, sIniFormExe, sInputIni, sInputPart, sInputBase, sOutputIni, sOutputPart, sOutputBase, sTempDir, sCurDir, sScriptVbs, sHomerLibVbs, sDir, sFile, sTargetTxt, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

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
Print "Finding empty files"
' sText = ShellExec(sCommand)
sDir = PathGetCurrentDirectory()
sFlags = "/a:-d /b /o:-s"
aFiles = PathGetSpec(sDir, sWildcards, sFlags)
Set oDir = oSystem.GetFolder(sDir)
Set oFiles = oDir.Files
' print Join(aFiles, vbCrLf)
aDelFiles = Array()
Set dDelFiles = CreateDictionary
' iBound = ArrayBound(aFiles)
iCount = oFiles.Count
' print StringPlural("total file", iBound - 1)
print StringPlural("total file", iCount)
' For iFile = 0 to iBound
For each oFile in oFiles
' print "iFile=" & iFile
' sFile = aFiles(iFile)
sFile = oFile.Path
' iSize = FileGetSize(sFile)
iSize = oFile.Size
' print iSize & ", " & sFile
If iSize = 0 Then dDelFiles.Add sFile, ""
Next ' sFile in aFiles

aFiles = dDelFiles.Keys
iFileCount = ArrayCount(aFiles)
Print StringPlural("empty file", iFileCount)
If iFileCount = 0 Then Quit ""

For Each sFile in aFiles
printBlank
Print PathGetName(sFile)
' If not (FileExists(sFile) and FileDelete(sFile)) Then print "Error"
' On error resume next
Set oFile = oSystem.GetFile(sFile)
oFile.Delete
' Set oFile = oSystem.GetFile(PathGetShort(sFile))
' FileDelete(PathGetShort(sFile))
FileDelete(sFile)
' on error goto 0
if FileExists(sFile) then print "Error"
Next ' aFiles
printBlank
Print "Done"

