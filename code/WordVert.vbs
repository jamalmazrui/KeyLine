Option Explicit
WScript.Echo"Starting wdVert"

Dim a, aStyles, aIni
Dim bValue, bFound, bAddToRecentFiles, bConfirmConversions, bIncludePageNumbers, bHidePageNumbersInWeb, bRightAlignPageNumbers, bUseFields, bUseHeadingStyles, bUseHyperlinks, bUseOutlineLevels, bReadOnly
Dim bFormat, bForward, bMatchAlefHamza, bMatchAllWordForms, bMatchCase, bMatchControl, bMatchDiacritics, bMatchKashida, bMatchSoundsLike, bMatchWholeWord, bMatchWildcards
Dim d, dHeadingStyles, dStyle, dIni, dSourceIni, dSection
Dim i, iLevel, iReplaceCount, iTableId, iReplace, iWrap, iForward, iArgCount, iCount, iLowerHeadingLevel, iUpperHeadingLevel
Dim oSystem, oFile, oParagraph, oField, oAddedStyles, oApp, oData, oDoc, oDocs, oFind, oFont, oFormat, oProperty, oRange, oReplace, oStyle, oStyles, oToc, oTocs
Dim sBackupDocx, sTargetLog, sScriptVbs, sHomerLibVbs, sDir, sCode, sFindStyle, sReplaceStyle, sKey, sFind, sFindText, sReplaceText, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni, sSection

Const WdDoNotSaveChanges = 0

Function FileInclude(sFile)
With CreateObject("Scripting.FileSystemObject")
ExecuteGlobal .openTextFile(sFile).readAll()
End With

' executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

' Main
sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile) 
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs

Dim aSourceFiles, aExtensions, aFormats
Dim bWordExisted, bErrorEvent
Dim iFormat, iSourceFile, iTargetFormat, iConvertCount, iSourceCount
Dim sProcess, sSource, sSourceDir, sWildcards, sSourceFile, sSourceName, sTargetFile, sTargetFormat, sTargetExtension, sTargetDir, sTarget
Dim oErrorEvent, oExtensions

Const msoEncodingUTF8 = 65001
Const wdFormatDocument = 0
Const wdFormatDocument97 = 0
Const wdFormatTemplate = 1
Const wdFormatTemplate97 = 1
Const wdFormatText = 2
Const wdFormatTextLineBreaks = 3
Const wdFormatDOSText = 4
Const wdFormatDOSTextLineBreaks = 5
Const wdFormatRTF = 6
Const wdFormatEncodedText = 7
Const wdFormatUnicodeText = 7
Const wdFormatHTML = 8
Const wdFormatWebArchive = 9
Const wdFormatFilteredHTML = 10
Const wdFormatXML = 11
Const wdFormatXMLDocument = 12
Const wdFormatXMLDocumentMacroEnabled = 13
Const wdFormatXMLTemplate = 14
Const wdFormatXMLTemplateMacroEnabled = 15
Const wdFormatOriginalFormatting = 16
Const wdFormatDocumentDefault = 16
Const wdFormatPDF = 17
Const wdFormatXPS = 18
Const wdFormatFlatXML = 19
Const wdFormatFlatXMLMacroEnabled = 20
Const wdFormatFlatXMLTemplate = 21
Const wdFormatFlatXMLTemplateMacroEnabled = 22
Const wdFormatPlainText = 22
Const wdFormatOpenDocumentText = 23

Function HelpAndExit()
Dim s
s = "Help for WdVert.exe -- Convert files using the API of Microsoft Word"
s = s & vbCrLf & "Syntax:"
s = s & vbCrLf & "WdVert Source Target TargetType"
s = s & vbCrLf & "where Source is the path to a file, directory, or wildcard specification"
s = s & vbCrLf & "optional Target is the path to either a file or directory, defaulting to the source directory"
s = s & vbCrLf & "optional TargetType is the target file type, as indicated by a common extension, integer constant, or string constant, defaulting to the txt extension"
Echo(s)
wscript.Quit
End Function

' Main program
sProcess = "WinWord.exe"
Set oExtensions = CreateDictionary()
oExtensions("doc") = wdFormatDocument
oExtensions("htm") = wdFormatFilteredHTML
oExtensions("html") = wdFormatFilteredHTML
oExtensions("pdf") = wdFormatPDF
oExtensions("rtf") = wdFormatRTF
oExtensions("txt") = wdFormatText
oExtensions("odt") = wdFormatOpenDocumentText  
oExtensions("xps") = wdFormatXPS    
oExtensions("mht") = wdFormatWebArchive      
oExtensions("mhtm") = wdFormatWebArchive      
oExtensions("xml") = wdFormatXML        
oExtensions("docx") = wdFormatXMLDocument        

iArgCount = WScript.Arguments.Count
If iArgCount = 0 Then HelpAndExit()

sSource = WScript.Arguments(0)

sSourceDir = sSource
If Not FolderExists(sSourceDir) Then sSourceDir = PathGetFolder(sSource)
sSource = PathGetFull(sSource)
If Len(sSourceDir) = 0 Then sSourceDir = PathGetCurrentDirectory()

sTarget = sSourceDir
If iArgCount > 1 Then sTarget = WScript.Args(1)
sTarget = PathGetFull(sTarget)
sTargetDir = sTarget
If Not FolderExists(sTargetDir) Then sTargetDir = PathGetFolder(sTargetDir)

sTargetExtension = "txt"
If Not FolderExists(sTarget) Then sTargetExtension = PathGetExtension(sTarget)
If iArgCount > 2 Then sTargetExtension = WScript.Args(2)
iTargetFormat = -1
If oExtensions.Exists(sTargetExtension) Then
iTargetFormat = oExtensions(sTargetExtension)
ElseIf LCase(Digit(sTargetExtension) Then
iTargetFormat = Number(sTargetExtension)
sTargetExtension = ""
Else
iTargetFormat = Eval(sTargetExtension)
If IsString(iTargetFormat) And Not StringLen(iTargetFormat) Then iTargetFormat = Eval("wdFormat" & sTargetExtension)
sTargetExtension = ""
End If

If Len(sTargetExtension) = 0 Then
aExtensions = oExtensions.Keys
aFormats = oExtensions.Items
For i = 0 To UBound(aFormats) - 1
iFormat = aFormats(i)
If iFormat = iTargetFormat Then
sTargetExtension = aExtensions(i)
Exit For
End If
 Next
End If

sWildCards = "*.*"
If Not FolderExists(sSource) Then sWildcards = PathGetName(sSource)
aSourceFiles = PathGetSpec(sSourceDir, sWildcards, "")
iSourceCount = 0
If IsArray(aSourceFiles) Then iSourceCount  = ArrayCount(aSourceFiles)

If ProcessIsModuleActive(sProcess) Then bWordexisted = True
Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
Set oDocs = oApp.Documents

iConvertCount = 0
For iSourceFile = 0 To iSourceCount - 1
If iSourceFile = 1 Then Echo("Converting")
sSourceFile = aSourceFiles(iSourceFile)
' sSourceFile = PathCombine(sSourceDir, sSourceFile)
s = sTargetExtension
If Len(s) > 0 Then s = "." & s
sTargetFile = sTarget
If FolderExists(sTargetFile) Then sTargetFile = PathCombine(sTargetDir, PathGetRoot(sSourceFile) & s)
If LCase(sSourceFile) = LCase(sTargetFile) Then ContinueLoop

sSourceName = PathGetName(sSourceFile)
Echo(sSourceName)

' Set oDoc = oDocs.Open(sSourceFile, AddToRecentFiles = False, ReadOnly = True, ConfirmConversions = False)
print "sSourceFile " & sSourceFile
Set oDoc = oDocs.Open(sSourceFile, False, True, False)

If FileExists(sTargetFile) Then FileDelete(sTargetFile)
' oDoc.SaveAs(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks)
oDoc.SaveAs sTargetFile, iTargetFormat, False, "", False, "", False, False, False, False, False, msoEncodingUTF8  
if bErrorEvent then bErrorEvent = 0
If FileExists(sTargetFile) Then iConvertCount = iConvertCount + 1
'' oDoc.Close()
Set oDoc = Nothing
Next
Set oDocs = Nothing

oApp.Quit()
Set oApp = Nothing
If Not bWordExisted And ProcessIsModuleActive(sProcess) Then ProcessClose(sProcess)


Echo("Converted " & iConvertCount & " out of " & StringPlural("file", iSourceCount))
