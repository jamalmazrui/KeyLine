Option Explicit
WScript.Echo"Starting DelDupes"

Dim a, aSelectFiles, aFiles
Dim bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dOutput, dProperties, d, dStyle, dIni, dSourceIni
Dim iFile, iFileCount, iError, iSuggestionCount, iErrorCount, iArg, iArgCount, iCount
Dim oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sChoice, sOutputTxt, sInputTxt, sSourceExe, sFiles, sSourcePart, sCommand, sSourceBase, sIniDir, sBinDir, sTempTmp, sSfkExe, sIniFormExe, sInputIni, sInputPart, sInputBase, sOutputIni, sOutputPart, sOutputBase, sTempDir, sCurDir, sScriptVbs, sHomerLibVbs, sDir, sFile, sTargetTxt, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

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
Print "Finding duplicates"
' sText = ShellExec(sCommand)
ShellRun sCommand, HIDDEN, Wait
sText = FileToString(sTempTmp)
sText = StringConvertToUnixLineBreak(sText)
FileDelete(sTempTmp)

aFiles = RegExpExtract(sText, "DEL :.*?\n", False)
iFileCount = ArrayCount(aFiles)
Print StringPlural("file", iFileCount)
If iFileCount = 0 Then Quit ""

For iFile = 0 to iFileCount - 1
aFiles(iFile) = StringTrimWhiteSpace(Replace(aFiles(iFile), "DEL :", ""))
Next ' aFiles


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
Print sFile
If not FileDelete(sFile) Then print "Error"
Next ' aFiles
printBlank
Print "Done"
End If ' sChoice
End If
End If

