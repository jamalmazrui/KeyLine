Option Explicit
WScript.Echo"Starting DocxSpelling"

Dim aIni
Dim bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dProperties, d, dStyle, dIni, dSourceIni
Dim iError, iSuggestionCount, iErrorCount, iArgCount, iCount
Dim oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sScriptVbs, sHomerLibVbs, sDir, sFile, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetTxt, sText, sConfigFile, sSourceIni

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
AppendEcho "Specify a source .docx file as a parameter."
AppendEcho "Optionally specify a configuration .ini file as a second parameter"
Quit
End If

sSourceDocx = WScript.Arguments(0)
' If Not InStr(sSourceDocx, "\") Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If InStr(sSourceDocx, "\") = 0 Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If not FileExists(sSourceDocx) Then Quit "Cannot find " & sSourceDocx

If iArgCount > 1 Then
sSourceIni = WScript.Arguments(1)
If InStr(sSourceIni, "\") = 0 Then sSourceIni = PathCombine(PathGetCurrentDirectory(), sSourceIni)
If not FileExists(sSourceIni) Then Quit "Cannot find " & sSourceIni
Set dSourceIni = IniToDictionary(sSourceIni)
bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)
bLogActions = False
' bReadOnly = False
bReadOnly = True
Else
sSourceIni = ""
Set dSourceIni = CreateDictionary()
bLogActions = False
bReadOnly = True
End If

sTargetTxt = PathCombine(PathGetCurrentDirectory(), "SPELLING-" & PathGetRoot(sSourceDocx) & ".txt")
sTargetLog = PathCombine(PathGetCurrentDirectory(), "SPELLING-" & PathGetRoot(sSourceDocx) & ".log")

Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
Set oDocs = oApp.Documents
bAddToRecentFiles = False
bConfirmConversions = False
print "Opening " & PathGetName(sSourceDocx)
Set oDoc = oDocs.Open(sSourceDocx, bAddToRecentFiles, bReadOnly, bConfirmConversions)

If not bReadOnly Then
 AppendEcho "Applying " & PathGetName(sSourceIni)

If Not oDoc.Saved Then
sBackupDocx = FileBackup(sSourceDocx)
If Len(sBackupDocx) = 0 Then
AppendEcho "Error creating backup "
Else
AppendEcho "Creating backup " & PathGetName(sBackupDocx)
AppendEcho "Saving " & PathGetName(sSourceDocx)
oDoc.Save
If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed
End If
End If
End If
End if ' Not bReadOnly

Set oContent = oDoc.Content
oApp.ResetIgnoreAll
oDoc.SpellingChecked = False
If oApp.CheckSpelling(oContent.Text) Then
AppendEcho "No spelling errors"
Else
Set oErrors = oDoc.SpellingErrors
iErrorCount = oErrors.Count
AppendEcho StringPlural("Spelling error", iErrorCount)
For iError = 1 to iErrorCount
PrintBlank
AppendBlank
Set oError = oErrors(iError)
sWord = oError.Text
AppendEcho iError & ". " & sWord
Set oSuggestions = oApp.GetSpellingSuggestions(sWord)
iSuggestionCount = oSuggestions.Count
AppendEcho StringPlural("suggestion", iSuggestionCount)
For Each oSuggestion in oSuggestions
AppendEcho oSuggestion.Name
Next
Next
End If

oApp.NormalTemplate.Saved = True
' oDoc.Close(wdDoNotSaveChanges)
oApp.Quit

' Create target txt
print "Saving " & PathGetName(sTargetTxt)
StringToFile sHomerText, sTargetTxt

If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
End If
echo "Done"
