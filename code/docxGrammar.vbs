Option Explicit
WScript.Echo"Starting DocxGrammar"

Dim aIni
Dim bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dProperties, d, dStyle, dIni, dSourceIni
Dim iError, iSuggestionCount, iErrorCount, iArgCount, iCount
Dim oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sScriptVbs, sHomerLibVbs, sDir, sFile, sTargetTxt, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

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
Print "Optionally specify a configuration .ini file as a second parameter"
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
' bReadOnly = False
bReadOnly = True
Else
sSourceIni = ""
Set dSourceIni = CreateDictionary()
bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)
bLogActions = False
bReadOnly = True
End If

sTargetTxt = PathCombine(PathGetCurrentDirectory(), "GRAMMAR-" & PathGetRoot(sSourceDocx) & ".txt")
sTargetLog = PathCombine(PathGetCurrentDirectory(), "GRAMMAR-" & PathGetRoot(sSourceDocx) & ".log")

Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
Set oDocs = oApp.Documents
bAddToRecentFiles = False
bConfirmConversions = False
Print "Opening " & PathGetName(sSourceDocx)
Set oDoc = oDocs.Open(sSourceDocx, bAddToRecentFiles, bReadOnly, bConfirmConversions)

If not bReadOnly Then
 Print "Applying " & PathGetName(sSourceIni)

If Not oDoc.Saved Then
sBackupDocx = FileBackup(sSourceDocx)
If Len(sBackupDocx) = 0 Then
Print "Error creating backup "
Else
Print "Creating backup " & PathGetName(sBackupDocx)
print "Saving " & PathGetName(sSourceDocx)
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
oDoc.GrammarChecked = False
' If oDoc.CheckGrammar() Then
' If oApp.CheckGrammar(oContent.Text) Then
If False Then
print "No Grammar errors"
Else
Set oErrors = oDoc.GrammaticalErrors
iErrorCount = oErrors.Count
print StringPlural("Grammar error", iErrorCount)
For iError = 1 to iErrorCount
PrintBlank
Set oError = oErrors(iError)
sWord = oError.Text
' print iError & ". " & sWord
AppendEcho iError & ". " & sWord
If iError < iErrorCount Then AppendBlank
if false then
Set oSuggestions = oApp.GetGrammarSuggestions(sWord)
iSuggestionCount = oSuggestions.Count
print StringPlural("suggestion", iSuggestionCount)
For Each oSuggestion in oSuggestions
print oSuggestion.Name
Next
end if
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
