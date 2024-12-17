Option Explicit
WScript.Echo"Starting ini2tables"

Dim aIni
Dim bBackupDocx, bLogActions, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dProperties, d, dStyle, dIni, dSourceIni
Dim iError, iSuggestionCount, iErrorCount, iArg, iArgCount, iCount
Dim oSystem, oFile, oContent, oSuggestions, oSuggestion, oErrors, oError, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sExtensions, sFields, sRepeatFields, sScriptVbs, sHomerLibVbs, sDir, sFile, sTargetTxt, sWord, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

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
Print "Specify a source .ini file as a parameter."
Quit
End If

sSourceIni = WScript.Arguments(0)
If InStr(sSourceIni, "\") = 0 Then sSourceIni = PathCombine(PathGetCurrentDirectory(), sSourceIni)
If not FileExists(sSourceIni) Then Quit "Cannot find " & sSourceIni

Set dSourceIni = IniToDictionary(sSourceIni)
bReadOnly = True

sFields = ""
If iArgCount > 1 Then sFields = WScript.Arguments(1)
sRepeatFields = ""
If iArgCount > 2 Then sRepeatFields = WScript.Arguments(2)
If iArgCount > 3 Then
sExtensions = ""
For iArg = 3 to iArgCount - 1
sExtensions = sExtensions + WScript.Arguments(iArg)
If iArg < iArgCount - 1 Then sExtensions = sExtensions + " "
Next
 sRepeatFields = WScript.Arguments(2)
Else
sExtensions = "csv docx htm xlsx"
End If

IniToTables sSourceIni, sFields, sRepeatFields, sExtensions


rem echo "Done"
