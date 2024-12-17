Option Explicit
WScript.Echo"Starting delSimilar"

Dim aDelFiles, a, aSelectFiles, aFiles
Dim bIgnoreCase, bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dDelFiles, dOutput, dProperties, d, dStyle, dIni, dSourceIni
Dim iStem, iDelFile, iBound, iFile, iFileCount, iError, iSuggestionCount, iErrorCount, iArg, iArgCount, iCount
Dim oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
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
Print "Finding similar files"
' sText = ShellExec(sCommand)
sDir = PathGetCurrentDirectory()
sFlags = "/a:-d /b /o:-s"
aFiles = PathGetSpec(sDir, sWildcards, sFlags)
' print Join(aFiles, vbCrLf)
aDelFiles = Array()
Set dDelFiles = CreateDictionary
iBound = ArrayBound(aFiles)
print StringPlural("total file", iBound - 1)
For iFile = 0 to iBound
' print "iFile=" & iFile
sFile = aFiles(iFile)
' if True Then
If not dDelFiles.Exists(sFile) Then
sRoot = PathGetRoot(sFile)
sExt = PathGetExtension(sFile)
' Do not match chapter1 or part2
' sMatch = "^(.*?)(\(|\)|\[|\]|[-_ 0-9])+$"
' try non-alpha suffix
' sMatch = "(\(|\[|-|_)+[0-9]+(\)|\])*"
' sMatch = "^.*?[a-zA-Z].*?[^a-zA-Z]+"
' sMatch = "^(.*?[a-zA-Z].*?)[^0-9a-zA-Z]+[0-9][^a-zA-Z]*"
sMatch = "^(.*?[a-zA-Z].*?)[^0-9a-zA-Z]+[0-9][^a-zA-Z]*"
sMatch = "^(.+) *[_-]\d+"
sReplace = "$1"
bIgnoreCase = True
sStem = RegExpReplace(sRoot, sMatch, sReplace, bIgnoreCase)
iStem = Len(sStem)
' print "sRoot=" & sRoot
' print "sStem=" & sStem
For iDelFile = (iFile + 1) to iBound
' print "iDelFile=" & iDelFile
sDelFile = aFiles(iDelFile)
sDelRoot = PathGetRoot(sDelFile)
sDelExt = PathGetExtension(sDelFile)
' If Len(sDelRoot) >= iStem Then
If Len(sDelRoot) >= iStem and iStem > 0 Then
If LCase(Left(sDelRoot, 1)) = LCase(Left(sRoot, 1)) Then
If LCase(sDelExt) = LCase(sExt) Then
sDelStem = RegExpReplace(sDelRoot, sMatch, sReplace, bIgnoreCase)
' If sDelStem = sStem then ArrayAdd aDelFiles, sDelFile
' If sDelStem = sStem and not dDelFiles.Exists(sDelFile) Then dDelFiles.Add sDelFile, ""
' If lCase(sDelStem) = LCase(sStem) Then print "sDelFile=" & sDelFile
' If lCase(sDelStem) = LCase(sStem) Then dDelFiles.Add sDelFile, ""
If lCase(sDelStem) = LCase(sStem) Then dDelFiles(sDelFile) = ""
' If sDelStem = sStem then print"Delete " & sDelRoot
End If
End If
End If ' lower(sExt) == Lower(sDelExt)
Next ' iDelFile
End If ' if not dDelFiles.Exists(sFile)
Next ' sFile in aFiles

' print StringPlural("similar file", ArrayCount(aDelFiles))
' aFiles = aDelFiles
aFiles = dDelFiles.Keys
iFileCount = ArrayCount(aFiles)
Print StringPlural("similar file", iFileCount)
If iFileCount = 0 Then Quit ""

' Use Input.txt
If False Then
sFiles = Join(aFiles, "|")

sText = FileToString(sSourceIni)
sText = StringConvertToUnixLineBreak(sText)
sText = Replace(sText, vbLf & "[Pick]" & vbLf, vbLf & "[Pick]" & vbLf & "range=" & sFiles & vbLf)

' print sText
StringToFile sText, sInputIni
Else
FileCopy sSourceIni, sInputIni
sFiles = Join(aFiles, vbCrLf)
sText = "[[Pick]]" & vbCrLf & sFiles

StringToFile sText, sInputTxt
End If

PathSetCurrentDirectory(sTempDir)
' FileCopy sSourceExe, sIniFormExe
' sCommand = StringQuote(sIniFormExe)
sCommand = StringQuote(sSourceExe) & " " & StringQuote(sTempDir)
' sCommand = "cmd.exe /c " & sCommand
' ShellExec sCommand
' ShellRun sCommand, MAX, Wait
ShellRun sCommand, NORMAL, Wait

' Use output.txt
If False Then
sText = FileToString(sOutputIni)
sText = StringConvertToUnixLineBreak(sText)
aFiles = RegExpExtract(sText, "\nPick=.*\n", False)
sFiles = aFiles(0)
sText = Replace(sText, vbLf & "Pick=", "")
sText = StringTrimWhiteSpace(sText)
aFiles = Split(sText, "|")
sText = Join(aFiles, vbCrLf)
Else
sText = FileToString(sOutputTxt)
sText = StringConvertToUnixLineBreak(sText)
' sText = RegExpReplace(sText, "^.*?\n", "", False)
a = RegExpExtract(sText, "(^|\n)\[\[Pick\]\](.|\n)*?($|\n\[\[)", False)
sText = ""
If ArrayCount(a) > 0 Then sText = a(0)
sText = sText & vbLf
sText = RegExpReplace(sText, "^.*?\n", "", False)
sText = RegExpReplace(sText, "\n.*?$", "", False)
sText = StringTrimWhiteSpace(sText)
aSelectFiles = Split(sText, vbLf)
End If

PathSetCurrentDirectory(sCurDir)
FileDelete sInputIni
FileDelete sInputTxt
FileDelete sOutputTxt
FileDelete sIniFormExe

If FileExists(sOutputIni) Then
Set dOutput = IniToDictionary(sOutputIni)
FileDelete sOutputIni
If dOutput("Results")("Cancel") <> "1" Then
If dOutput("Results")("&Selected") = "1" Then aFiles = aSelectFiles
' Don't confirm after escaping IniForm
iFileCount = ArrayCount(aFiles)
' s = "Delete " & StringPlural("file", ArrayCount(aFiles)) & "?"
s = "Delete " & StringPlural("file", iFileCount) & "?"
if iFileCount = 0 Then
sChoice = "N"
else
sChoice = DialogConfirm(s, sText, "N")
End If
if sChoice = "Y" Then
print "Deleting"
For Each sFile in aFiles
printBlank
' If not (FileExists(sFile) and FileDelete(sFile)) Then print "Error"
Print PathGetName(sFile)
' If not (FileExists(sFile) and FileDelete(sFile)) Then print "Error"
On error resume next
FileDelete(sFile)
FileDelete(PathGetShort(sFile))
on error goto 0
if FileExists(sFile) then print "Error"
Next ' aFiles
printBlank
Print "Done"
End If ' sChoice
End If
End If

