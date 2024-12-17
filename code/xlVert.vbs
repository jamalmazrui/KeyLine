Option Explicit
WScript.Echo"Starting xlVert"

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

Dim bExcelExisted, bErrorEvent, bNoteLabel, bOutlineLabel, bCommentLabel, bHyperlinkLabel
Dim iSlideCount, iShapeCount, iSlide, iShape, iCommentCount, iComment, iNoteCount, iNote, iHyperlinkCount, iHyperlink
Dim oErrorEvent, oXlss, oXls, oTextFrame, oTextRange, oShapes, oShape, oSlides, oSlide, oNotes, oNote, oComments, oComment, oHyperlinks, oHyperlink
Dim sContents, sBody, sSourceFile, sTargetFile, sXls
Dim aSourceFiles, aExtensions, aFormats
Dim iFormat, iSourceFile, iTargetFormat, iConvertCount, iSourceCount, iSheetCount, iSheet
Dim sProcess, sSource, sSourceDir, sWildcards, sSourceName, sTargetFormat, sTargetExtension, sTargetDir, sTarget
Dim oExtensions, oLabels, oSheets, oSheet

Const msoFalse = 0
Const msoTrue = -1
Dim sDivider : sDivider = vbCrLf & "----------" & vbCrLf & vbFormFeed & vbCrLf
Const msoEncodingUTF8 = 65001

Const msoAutomationSecurityLow = 1
Const msoAutomationSecurityByUI = 2
Const msoAutomationSecurityForceDisable = 3

Const xlSYLK = 2
Const xlWKS = 4
Const xlWK1 = 5
Const xlCSV = 6
Const xlDBF2 = 7
Const xlDBF3 = 8
Const xlDIF = 9
Const xlDBF4 = 11
Const xlWJ2WD1 = 14
Const xlWK3 = 15
Const xlExcel2 = 16
Const xlTemplate = 17
Const xlAddIn = 18
Const xlTextMac = 19
Const xlTextWindows = 20
Const xlTextMSDOS = 21
Const xlCSVMac = 22
Const xlCSVWindows = 23
Const xlCSVMSDOS = 24
Const xlIntlMacro = 25
Const xlIntlAddIn = 26
Const xlExcel2FarEast = 27
Const xlWorks2FarEast = 28
Const xlExcel3 = 29
Const xlWK1FMT = 30
Const xlWK1ALL = 31
Const xlWK3FM3 = 32
Const xlExcel4 = 33
Const xlWQ1 = 34
Const xlExcel4Workbook = 35
Const xlTextPrinter = 36
Const xlWK4 = 38
Const xlExcel7 = 39
Const xlWJ3 = 40
Const xlWJ3FJ3 = 41
Const xlUnicodeText = 42
Const xlExcel9795 = 43
Const xlHtml = 44
Const xlWebArchive = 45
Const xlXMLSpreadsheet = 46
Const xlExcel12 = 50
Const xlWorkbookDefault = 51
Const xlOpenXMLWorkbook = 51
Const xlOpenXMLWorkbookMacroEnabled = 52
Const xlOpenXMLTemplateMacroEnabled = 53
Const xlOpenXMLTemplate = 54
Const xlOpenXMLAddIn = 55
Const xlExcel8 = 56
Const xlOpenDocumentSpreadsheet = 60
Const xlWorkbookNormal = -4143
Const xlCurrentPlatformText = -4158

Function filewriteUtf8b(sFile, sBody)
StringToFile sBody, sFile
End Function

Function HelpAndExit()
Dim s
s = "Help for XlVert.exe -- Convert files using the API of Microsoft Excel"
s = s & vbCrLf & "Syntax:"
s = s & vbCrLf & "XlVert Source Target TargetType"
s = s & vbCrLf & "where Source is the path to a file, directory, or wildcard specification"
s = s & vbCrLf & "optional Target is the path to either a file or directory, defaulting to the source directory"
s = s & vbCrLf & "optional TargetType is the target file type, as indicated by a common extension, integer constant, or string constant, defaulting to the txt extension"
Echo(s)
WScript.Quit
End Function

' Main program
sProcess = "Excel.exe"
'' Set oLabels = CreateDictionary()

Set oExtensions = CreateDictionary()
oExtensions("xls") = xlWorkbookDefault
oExtensions("csv") = xlCSV
' oExtensions("dbf") = xlDBF3
oExtensions("dif") = xlDIF
oExtensions("htm") = xlHtml
oExtensions("html") = xlHtml
oExtensions("sylk") = xlSYLK
oExtensions("txt") = xlTextWindows
oExtensions("ods") = xlOpenDocumentSpreadsheet
oExtensions("mht") = xlWebArchive
oExtensions("mhtm") = xlWebArchive
oExtensions("xml") = xlXMLSpreadsheet
' oExtensions("xml") = xlOpenXMLWorkbook
oExtensions("wks") = xlWKS

iArgCount = WScript.Arguments.Count
If iArgCount = 0 Then HelpAndExit()

sSource = WScript.Arguments(0)
' sSource = "*.xls"

sSourceDir = sSource
If Not FolderExists(sSourceDir) Then sSourceDir = PathGetFolder(sSource)
sSource = PathGetFull(sSource)
If IsBlank(sSourceDir) Then sSourceDir = PathGetCurrentDirectory

sTarget = sSourceDir
If iArgCount > 1 Then sTarget = WScript.Arguments(1)
sTarget = PathGetFull(sTarget)
sTargetDir = sTarget
If Not FolderExists(sTargetDir) Then sTargetDir = PathGetFolder(sTargetDir)

sTargetExtension = "txt"
If Not FolderExists(sTarget) Then sTargetExtension = PathGetExtension(sTarget)
If iArgCount > 2 Then sTargetExtension = WScript.Arguments(2)
iTargetFormat = -1
If oExtensions.Exists(sTargetExtension) Then
iTargetFormat = oExtensions(sTargetExtension)
ElseIf StringIsDigit(sTargetExtension) Then
iTargetFormat = Number(sTargetExtension)
sTargetExtension = ""
Else
iTargetFormat = Eval(sTargetExtension)
If IsString(iTargetFormat) And Not StringLen(iTargetFormat) Then iTargetFormat = Eval("wdFormat" & sTargetExtension)
sTargetExtension = ""
End If

If IsBlank(sTargetExtension) Then
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
If IsArray(aSourceFiles) Then iSourceCount = ArrayCount(aSourceFiles)

If ProcessIsModuleActive(sProcess) Then bExcelexisted = True
Set oApp = CreateObject("Excel.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
' oApp.Visible = True ; Needed for automation to work
oApp.Visible = False

Set oXlss = oApp.WorkBooks
' DialogShow("oXlss", IsObj(oXlss))

iConvertCount = 0
For iSourceFile = 0 To iSourceCount - 1
If iSourceFile = 1 Then Echo("Converting")
sSourceFile = aSourceFiles(iSourceFile)
' sSourceFile = PathCombine(sSourceDir, sSourceFile)
s = sTargetExtension
If IsNonBlank(s) Then s = "." & s
sTargetFile = sTarget
If FolderExists(sTargetFile) Then sTargetFile = PathCombine(sTargetDir, PathGetRoot(sSourceFile) & s)
If LCase(sSourceFile) = LCase(sTargetFile) Then ContinueLoop

sSourceName = PathGetName(sSourceFile)
Echo(sSourceName)

' expression.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)
Set oXls = oXlss.Open(sSourceFile)
' DialogShow("oXls", IsObj(oXls))

If FileExists(sTargetFile) Then FileDelete(sTargetFile)
' DialogShow(sTargetExtension)
If sTargetExtension = "txt" Then
Set oSheets = oXls.sheets
iSheetCount =oSheets.Count
sText = ""
sContents = "Contents" & vbCrLf
For iSheet = 1 To iSheetCount
Set oSheet = oSheets.Item(iSheet)
If FileExists(sTargetFile) Then FileDelete(sTargetFile)
oSheet.SaveAs sTargetFile, xlTextWindows
sName = StringTrimWhiteSpace(oSheet.Name)
If iSheetCount > 1 Then sName = "Sheet " & iSheet & ": " & sName
sBody = FileToString(sTargetFile)
sBody = sName & vbCrLf & vbCrLf & sBody
If IsNonBlank(sText) Then sText = sText & sDivider
sText = sText & sBody
sContents = sContents & vbCrLf & sName
Set oSheet = Nothing
oXls.Close()
Set oXls = Nothing
Set oXls = oXlss.Open(sSourceFile)
Set oSheets = oXls.Sheets
Next

If iSheetCount > 1 Then sText = iSheetCount & "Sheets" & sContents & vbCrLf & sText
' sText = StringRegExpReplace(sText, "\r\n\s*?\r\n\s*?(\r\n\s*?)+", "\r\n\r\n")
sText = StringTrimWhiteSpace(sText) & vbCrLf
filewriteUtf8b sTargetFile, sText
Else
' DialogShow("IsObject", IsObj(oXls))
' expression.SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
oXls.SaveAs sTargetFile, iTargetFormat
' DialogShow(bErrorEvent, sTargetFile)
if bErrorEvent then bErrorEvent = 0
End If

If FileExists(sTargetFile) Then iConvertCount = iConvertCount + 1
oXls.Close()
Set oXls = Nothing
Next
Set oXlss = Nothing

oApp.Quit()
Set oApp = Nothing
If Not bExcelExisted And ProcessIsModuleActive(sProcess) Then ProcessClose(sProcess)

Echo("Converted " & iConvertCount & " out of " & StringPlural("file", iSourceCount))
