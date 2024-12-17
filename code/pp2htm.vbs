' Option Explicit
WScript.Echo "Starting pp2htm"

Dim a, aStyles, aIni
Dim bValue, bFound, bAddToRecentFiles, bConfirmConversions, bIncludePageNumbers, bHidePageNumbersInWeb, bRightAlignPageNumbers, bUseFields, bUseHeadingStyles, bUseHyperlinks, bUseOutlineLevels, bReadOnly
Dim bFormat, bForward, bMatchAlefHamza, bMatchAllWordForms, bMatchCase, bMatchControl, bMatchDiacritics, bMatchKashida, bMatchSoundsLike, bMatchWholeWord, bMatchWildcards
Dim d, dHeadingStyles, dStyle, dIni, dSourceIni, dSection
Dim iRowCount, iColCount, iRow, iCol, i, iLevel, iReplaceCount, iTableId, iReplace, iWrap, iForward, iArgCount, iCount, iLowerHeadingLevel, iUpperHeadingLevel
Dim oTable, oRow, oCell, oSystem, oFile, oParagraph, oField, oAddedStyles, oApp, oData, oDoc, oDocs, oFind, oFont, oFormat, oProperty, oRange, oReplace, oStyle, oStyles, oToc, oTocs
Dim oTag, oTags
Dim sAlign, sCell, sSourceName, sTargetRoot, sTargetName, sTargetExtension, sBackupDocx, sTargetLog, sScriptVbs, sHomerLibVbs, sDir, sCode, sFindStyle, sReplaceStyle, sKey, sFind, sFindText, sReplaceText, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni, sSection

Const WdDoNotSaveChanges = 0

Function FileInclude(sFile)
if true then
executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
else
wscript.echo sFile
With CreateObject("Scripting.FileSystemObject")
ExecuteGlobal .openTextFile(sFile).readAll()
End With
end if

FileInclude = True
End Function

' Main
sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile)
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs

Dim bPowerPointExisted, bErrorEvent, bNoteLabel, bOutlineLabel, bCommentLabel, bHyperlinkLabel
Dim iSlideCount, iShapeCount, iSlide, iShape, iCommentCount, iComment, iNoteCount, iNote, iHyperlinkCount, iHyperlink
Dim oErrorEvent, oPpts, oPpt, oTextFrame, oTextRange, oShapes, oShape, oSlides, oSlide, oNotes, oNote, oComments, oComment, oHyperlinks, oHyperlink
Dim sSourceFile, sTargetFile, sTargetPdf, sPpt, oLabels

Dim aSourceFiles, aExtensions, aFormats
Dim iFormat, iSourceFile, iTargetFormat, iConvertCount, iSourceCount
Dim sProcess, sSource, sSourceDir, sWildcards, sTargetDir, sTarget
Dim oExtensions

Const msoFalse = 0
Const msoTrue = -1
Dim sDivider : sDivider = vbCrLf & "----------" & vbCrLf & vbFormFeed & vbCrLf
Const msoEncodingUTF8 = 65001

Const msoAutomationSecurityLow = 1
Const msoAutomationSecurityByUI = 2
Const msoAutomationSecurityForceDisable = 3

Const ppSaveAsPresentation = 1
Const ppSaveAsText = 2
Const ppSaveAsTemplate = 5
Const ppSaveAsRTF = 6
Const ppSaveAsShow = 7
Const ppSaveAsAddIn = 8
Const ppSaveAsDefault = 11
Const ppSaveAsHTML = 12
Const ppSaveAsHTMLv3 = 13
Const ppSaveAsHTMLDual = 14
Const ppSaveAsMetaFile = 15
Const ppSaveAsGIF = 16
Const ppSaveAsJPG = 17
Const ppSaveAsPNG = 18
Const ppSaveAsBMP = 19
Const ppSaveAsWebArchive = 20
Const ppSaveAsTIF = 21
Const ppSaveAsEMF = 23
Const ppSaveAsOpenXMLPresentation = 24
Const ppSaveAsOpenXMLPresentationMacroEnabled = 25
Const ppSaveAsOpenXMLTemplate = 26
Const ppSaveAsOpenXMLTemplateMacroEnabled = 27
Const ppSaveAsOpenXMLShow = 28
Const ppSaveAsOpenXMLShowMacroEnabled = 29
Const ppSaveAsOpenXMLAddin = 30
Const ppSaveAsOpenXMLTheme = 31
Const ppSaveAsPDF = 32
Const ppSaveAsXPS = 33
Const ppSaveAsXMLPresentation = 34
Const ppSaveAsOpenDocumentPresentation = 35
Const ppSaveAsExternalConverter = 36

Dim dBulletType : Set dBulletType = CreateDictionary()
dim ppBulletMixed : ppBulletMixed = -2
dim ppBulletNone : ppBulletNone = 0
dim ppBulletNumbered : ppBulletNumbered = 2
dim ppBulletPicture : ppBulletPicture = 3
dim ppBulletUnnumbered : ppBulletUnnumbered = 1

Dim dPlaceHolderType : Set dPlaceHolderType = CreateDictionary()
dPlaceHolderType.Add 9, "Bitmap"
dPlaceHolderType.Add 2, "Body"
dPlaceHolderType.Add 3, "Center Title"
dPlaceHolderType.Add 8, "Chart"
dPlaceHolderType.Add 16, "Date"
dPlaceHolderType.Add 15, "Footer"
dPlaceHolderType.Add 14, "Header"
dPlaceHolderType.Add 10, "Media Clip"
dPlaceHolderType.Add -2, "Mixed"
dPlaceHolderType.Add 7, "Object"
dPlaceHolderType.Add 11, "Organization Chart"
dPlaceHolderType.Add 18, "Picture"
dPlaceHolderType.Add 13, "Slide Number"
dPlaceHolderType.Add 4, "Subtitle"
dPlaceHolderType.Add 12, "Table"
dPlaceHolderType.Add 1, "Title"
dPlaceHolderType.Add 6, "Vertical Body"
dPlaceHolderType.Add 17, "Vertical Object"
dPlaceHolderType.Add 5, "Vertical Title"

Dim dLayoutType : Set dLayoutType = CreateDictionary()
dLayoutType.Add 12, "Blank"
dLayoutType.Add 8, "Chart"
dLayoutType.Add 6, "Chart and text"
dLayoutType.Add 10, "ClipArt and text"
dLayoutType.Add 26, "ClipArt and vertical text"
dLayoutType.Add 34, "Comparison"
dLayoutType.Add 35, "Content with caption"
dLayoutType.Add 32, "Custom"
dLayoutType.Add 24, "Four objects"
dLayoutType.Add 15, "Large object"
dLayoutType.Add 18, "MediaClip and text"
dLayoutType.Add -2, "Mixed"
dLayoutType.Add 16, "Object"
dLayoutType.Add 14, "Object and text"
dLayoutType.Add 30, "Object and two objects"
dLayoutType.Add 19, "Object over text"
dLayoutType.Add 7, "Organization chart"
dLayoutType.Add 36, "Picture with caption"
dLayoutType.Add 33, "Section header"
dLayoutType.Add 4, "Table"
dLayoutType.Add 2, "Text"
dLayoutType.Add 5, "Text and chart"
dLayoutType.Add 9, "Text and ClipArt"
dLayoutType.Add 17, "Text and MediaClip"
dLayoutType.Add 13, "Text and object"
dLayoutType.Add 21, "Text and two objects"
dLayoutType.Add 20, "Text over object"
dLayoutType.Add 1, "Title"
dLayoutType.Add 11, "Title only"
dLayoutType.Add 3, "Two-column text"
dLayoutType.Add 29, "Two objects"
dLayoutType.Add 31, "Two objects and object"
dLayoutType.Add 22, "Two objects and text"
dLayoutType.Add 23, "Two objects over text"
dLayoutType.Add 25, "Vertical text"
dLayoutType.Add 27, "Vertical title and text"
dLayoutType.Add 28, "Vertical title and text over chart"

Dim dShapeType : Set dShapeType = CreateDictionary()
dShapeType.Add 30, "3D model"
dShapeType.Add 1, "AutoShape"
dShapeType.Add 2, "Callout"
dShapeType.Add 20, "Canvas"
dShapeType.Add 3, "Chart"
dShapeType.Add 4, "Comment"
dShapeType.Add 27, "Content Office Add-in"
dShapeType.Add 21, "Diagram"
dShapeType.Add 7, "Embedded OLE object"
dShapeType.Add 8, "Form control"
dShapeType.Add 5, "Freeform"
dShapeType.Add 28, "Graphic"
dShapeType.Add 6, "Group"
dShapeType.Add 24, "SmartArt graphic"
dShapeType.Add 22, "Ink"
dShapeType.Add 23, "Ink comment"
dShapeType.Add 9, "Line"
dShapeType.Add 31, "Linked 3D model"
dShapeType.Add 29, "Linked graphic"
dShapeType.Add 10, "Linked OLE object"
dShapeType.Add 11, "Linked picture"
dShapeType.Add 16, "Media"
dShapeType.Add 12, "OLE control object"
dShapeType.Add 13, "Picture"
dShapeType.Add 14, "Placeholder"
dShapeType.Add 18, "Script anchor"
dShapeType.Add -2, "Mixed shape type"
dShapeType.Add 25, "Slicer"
dShapeType.Add 19, "Table"
dShapeType.Add 17, "Text box"
dShapeType.Add 15, "Text effect"
dShapeType.Add 26, "Web video"

Function FileSaveUTF16(sText, sFile)
' Saves string to text file, replacing if it exists

Dim bReplace
Dim oSystem, oFile

FileSaveUTF16 = False
If Not FileDelete(sFile) Then Exit Function

bReplace = True
Set oSystem =CreateObject("Scripting.FilesystemObject")
' DialogShow StringIsUnicode(sText), sText
if true Then
' If StringIsUnicode(sText) Then
Set oFile =oSystem.CreateTextFile(sFile, bReplace, True)
Else
Set oFile =oSystem.CreateTextFile(sFile, bReplace, False)
End If
On Error Resume Next
oFile.Write sText
On Error GoTo 0
oFile.Close
FileSaveUTF16 = FileExists(sFile)

Set oFile = Nothing
Set oSystem = Nothing
End Function

Function ProcessComment(oComment, sText)
Dim sLabel : sLabel = "Comment"
Dim sReturn : sReturn = oComment.Text
If LCase(sReturn) = LCase(sLabel) Then sReturn = ""
Dim sAuthor : sAuthor = StringTrimWhiteSpace(oComment.Author)
if sReturn and sAuthor Then sReturn = "By " & sAuthor & vbCrLf & sReturn

If Not sReturn Then
sReturn = sText
ElseIf oLabels(sLabel) Then
' sReturn = sText & vbCrLf & vbCrLf & sLabel & ":" & vbCrLf & sReturn
oLabels(sLabel) = False
Else
sReturn = sText & vbCrLf & vbCrLf & sReturn
End If
Return sReturn
End Function

Function ProcessHyperlink(oHyperlink, sText)
Dim sLabel : sLabel = "Hyperlink"
Dim sAddress : sAddress = StringTrimWhiteSpace(oHyperlink.Address)
Dim sTextToDisplay 
sTextToDisplay = ""
on error resume next
sTextToDisplay = StringTrimWhiteSpace(oHyperlink.TextToDisplay)
on error goto 0

Dim sScreenTip : sScreenTip = StringTrimWhiteSpace(oHyperlink.ScreenTip)
if len(sTextToDisplay) = 0 then sTextToDisplay = sScreenTip
Dim sReturn: sReturn = sText
sReturn = sReturn & "- [" & sTextToDisplay & "](" & sAddress & ")" & vbCrLf
ProcessHyperlink = sReturn
End Function

Function ProcessShape(oShape, sText, sLabel, iShape)
Dim oTextFrame, oTextRange, oTextEffect
Dim sReturn : sReturn = ""
Dim sTextRange : sTextRange = ""
Dim sChartTitle : sChartTitle = ""
Dim sTextEffect : sTextEffect = ""
Dim sAlternativeText : sAlternativeText = ""
dim sFont: sFont = ""
dim iFontSize: iFontSize = 0
dim iBulletType : iBulletType = ppBulletNone

If oShape.HasChart Then
sChartTitle = ""
on error resume next
sChartTitle = oShape.Chart.ChartTitle.Text
on error goto 0

end if

If oShape.HasTextFrame Then
Set oTextFrame = oShape.TextFrame
Set oTextRange = oTextFrame.TextRange
sTextRange = StringTrimWhiteSpace(oTextRange.Text)
iBulletType = oTextRange.ParagraphFormat.Bullet.Type
if iBulletType = ppBulletNumbered then
sTextRange = "1. " & Replace(sTextRange, vbCr, vbCr & "1. ")
sTextRange = Replace(sTextRange, vbCr & "1. " & vbCr, vbCr & vbCr)
end if
if iBulletType = ppBulletUnnumbered then
' sTextRange = StringConvertToUnixLineBreak(sTextRange)
' wscript.echo "count " & regexpcount(sTextRange, vbCr, true)
' wscript.echo len(sTextRange)
' sTextRange = RegexReplace(sTextRange, "^(.+)$", "- $1")
sTextRange = "- " & Replace(sTextRange, vbCr, vbCr & "- ")
sTextRange = Replace(sTextRange, vbCr & "- " & vbCr, vbCr & vbCr)
' wscript.echo len(sTextRange)
end if

sFont = oTextRange.Font.Name
iFontSize = oTextRange.Font.Size
else
on error resume next
sTextRange = oShape.Text
on error goto 0
End If
sTextRange = replace(sTextRange, vbCr, vbCrLf)

Dim sName : sName = oShape.Name
Dim iType : iType = oShape.type
Dim sType : sType = dShapeType(iType)
Dim sFormat : sFormat = ""
on error resume next
sFormat = dPlaceHolderType(oShape.PlaceholderFormat.Type)
on error goto 0

sAlternativeText = StringTrimWhiteSpace(oShape.AlternativeText)
' If iType = 15 Then ' TextEffect
on error resume next
Set oTextEffect = oShape.TextEffect
sTextEffect = StrimTrimWhiteSpace(oTextEffect.Text)
on error goto 0

s = "####"
if sLabel = "Notes" then s = s & "#"
' sReturn = sReturn & s & " Shape " & iShape & ": " & sType
' if len(sFormat) > 0 then sReturn = sReturn & ", " & sFormat
sReturn = sReturn & vbCrLf & vbCrLf

' sReturn = AddField(sReturn, "Name", sName)
sReturn = AddField(sReturn, "Chart Title", sChartTitle)
' if oShape.Decorative then sReturn = AddField(sReturn, "Decorative", "True")
sAlternativeText = StringTrimWhiteSpace(oShape.AlternativeText)
if sAlternativeText <> "" then sReturn = sReturn & "![" & sAlternativeText & "]()" & vbCrLf
sReturn = AddField(sReturn, "Effect", sTextEffect)
if oShape.HasTable then
set oTable = oShape.Table
sReturn = AddField(sReturn, "Title", oTable.Title)
iRowCount = oTable.Rows.Count
iColCount = oTable.Columns.Count
sReturn = sReturn & vbCrLf
iRow = 1
iCol = 1
sAlign = ""
For each oRow in oTable.Rows
for each oCell in oRow.Cells
sCell = oCell.Shape.TextFrame.TextRange.Text
if iCol > 1 then
sReturn = sReturn & "|"
sAlign = sAlign & "|"
end if
sReturn = sReturn & sCell
sAlign = sAlign & ":--"
iCol = iCol + 1
Next
sReturn = sReturn & vbCrLf
if iRow = 1 then sReturn = sReturn & sAlign & vbCrLf
iRow = iRow + 1
next
End If
if len(sTextRange) > 0 then
if sFormat = "Subtitle" then sTextRange = vbCrLf & "### " & sTextRange
if sFormat <> "Title" then sReturn = sReturn & vbCrLf & vbCrLf & sTextRange & vbCrLf & vbCrLf
end if
on error resume next
sReturn = AddField(sReturn, "Footer", oSlide.HeadersFooters.FooterText)
on error goto 0

' sReturn = StringTrimWhiteSpace(sReturn)
If IsNumeric(sReturn) Then sReturn = ""

If IsBlank(sReturn) Then
sReturn = sText
ElseIf oLabels(sLabel) Then
sReturn = sText & vbCrLf & vbCrLf & sLabel & ":" & vbCrLf & sReturn
oLabels(sLabel) = False
Else
sReturn = sText & vbCrLf & vbCrLf & sReturn
End If

Set oTextEffect = Nothing
Set oTextRange = Nothing
Set oTextFrame = Nothing

ProcessShape = sReturn
End Function

Function HelpAndExit()
Dim s
s = "Help for PpVert.exe -- Convert files using the API of Microsoft PowerPoint"
s = s & vbCrLf & "Syntax:"
s = s & vbCrLf & "PpVert Source Target TargetType"
s = s & vbCrLf & "where Source is the path to a file, directory, or wildcard specification"
s = s & vbCrLf & "optional Target is the path to either a file or directory, defaulting to the source directory"
s = s & vbCrLf & "optional TargetType is the target file type, as indicated by a common extension, integer constant, or string constant, defaulting to the txt extension"
Echo(s)
WScript.Quit
End Function

function AddField(sBodyText, sFieldName, sFieldValue)
AddField = sBodyText
if Len("" & sFieldName) = 0 then exit function
if Len("" & sFieldValue) = 0 then exit function
AddField = sBodyText & sFieldValue & " \" & vbCrLf
End Function

Function ExportPdf(oPresentation, sFilePdf)
' expression.PrintOut (From, To, PrintToFile, Copies, Collate)
Const ppFixedFormatTypePDF = 2
Const ppFixedFormatIntentScreen = 1
const ppPrintAll = 1
const ppPrintHandoutVerticalFirst = 1
Const ppPrintOutputNotesPages = 5
const msoCTrue = 1
const msoFalse = 0
Dim vFrameSlides, vHandoutOrder, vPrintHiddenSlides, vPrintRange, vRangeType, vSlideShowName, vIncludeDocProperties, vKeepIRMSettings, vDocStructureTags, vBitmapMissingFonts, vUseISO19005_1, vExternalExporter
vFrameSlides = false
vHandoutOrder = 1
vPrintHiddenSlides = false
set vPrintRange = oPresentation.PrintOptions.Ranges.Add(1, oPresentation.Slides.Count)
vRangeType = 1
vSlideShowName = null
vIncludeDocProperties = false
vKeepIRMSettings = false
vDocStructureTags = true
vBitmapMissingFonts = false
vUseISO19005_1 = false
vExternalExporter = null

If FileExists(sFilePdf) then FileDelete(sFilePdf)
' oPresentation.ExportAsFixedFormat sFilePdf, ppFixedFormatTypePDF, ppFixedFormatIntentScreen  , vFrameSlides, vHandoutOrder, ppPrintOutputNotesPages  , vPrintHiddenSlides, vPrintRange, vRangeType, vSlideShowName, vIncludeDocProperties, vKeepIRMSettings, vDocStructureTags, vBitmapMissingFonts, vUseISO19005_1, vExternalExporter
oPresentation.ExportAsFixedFormat sFilePdf, ppFixedFormatTypePDF, ppFixedFormatIntentScreen  , vFrameSlides, vHandoutOrder, ppPrintOutputNotesPages  , vPrintHiddenSlides, vPrintRange, vRangeType, , vIncludeDocProperties, vKeepIRMSettings, vDocStructureTags, vBitmapMissingFonts, vUseISO19005_1
' oPresentation.ExportAsFixedFormat sFilePdf, ppFixedFormatTypePDF, ppFixedFormatIntentScreen, msoCTrue, ppPrintHandoutVerticalFirst, ppPrintOutputNotesPages, msoFalse, null, null, null, False, False, False, False, False

ExportPdf = FileExists(sFilePdf)

End Function
' Main program
sProcess = "PowerPnt.exe"
Set oLabels = CreateDictionary()

Set oExtensions = CreateDictionary()
oExtensions("ppt") = ppSaveAsPresentation
oExtensions("htm") = ppSaveAsHTML
oExtensions("html") = ppSaveAsHTML
oExtensions("pdf") = ppSaveAsPDF
oExtensions("rtf") = ppSaveAsRTF
oExtensions("txt") = ppSaveAsText
oExtensions("odp") = ppSaveAsOpenDocumentPresentation
oExtensions("xps") = ppSaveAsXPS
oExtensions("mht") = ppSaveAsWebArchive
oExtensions("mhtm") = ppSaveAsWebArchive
oExtensions("xml") = ppSaveAsOpenXMLPresentation
oExtensions("pptx") = ppSaveAsXMLPresentation

iArgCount = WScript.Arguments.Count
If iArgCount = 0 Then HelpAndExit()

sSource = WScript.Arguments(0)
' sSource = "c:\broadband\*.pp*"

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
If IsString(iTargetFormat) And Not StringLen(iTargetFormat) Then iTargetFormat = Eval("ppSaveAs" & sTargetExtension)
sTargetExtension = ""
End If

If IsBlank(sTargetExtension) Then
aExtensions = oExtensions.Keys
aFormats = oExtensions.Items
For i = 0 To UBound(aFormats)
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

If ProcessIsModuleActive(sProcess) Then bPowerPointexisted = True
Set oApp = CreateObject("PowerPoint.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.Visible = True ' Needed for automation to work

Set oPpts = oApp.Presentations

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
sTargetPdf = PathGetFolder(sSourceFile) & "\" & PathGetRoot(sSourceFile) & ".pdf"
sTargetRoot = PathGetRoot(sSourceName)
sTargetName = sTargetRoot & ".md"
sTargetFile = sTargetName

' expression.Open(FileName, ReadOnly, Untitled, WithWindow)
Set oPpt = oPpts.Open(sSourceFile, msoTrue, msoTrue, msoFalse)

If FileExists(sTargetFile) Then FileDelete(sTargetFile)
If sTargetExtension = "txt" or sTargetExtension = "md" Then
sPpt = oPpt.name
sText = "---" & vbCrLf & "title: " & oPpt.Name & vbCrlf & "---" & vbCrlf
Set oSlides = oPpt.slides
iSlideCount = oSlides.Count
sText = sText & vbCrLf & StringPlural("Slide", iSlideCount) & vbCrLf & vbCrLf
For iSlide = 1 To iSlideCount
Set oSlide = oSlides.Item(iSlide)
Set oShapes = oSlide.Shapes
sText = sText & vbcrlf
sText = sText & "## Slide " & iSlide
on error resume next
sText = sText & "## Slide " & iSlide & ": " & oShapes.Title.TextFrame.TextRange.Text
on error goto 0
iShapeCount = oShapes.Count
sText = sText & vbCrLf
if oSlide.HasNotesPage then s = s & " with Notes"
Set oHyperlinks = oSlide.Hyperlinks
iHyperlinkCount = oHyperlinks.Count
if iHyperLinkCount > 0 then s = s & " and Links"
' sText = sText & s & vbCrLf & vbCrLf
for iShape = 1 to iShapeCount
Set oShape = oShapes.Item(iShape)
sText = ProcessShape(oShape, sText, "Outline", iShape)
Set oShape = Nothing
Next
Set oShapes = Nothing

if (oSlide.HasNotesPage) then
sText = sText & vbCrLf
Set oNotes = oSlide.NotesPage
iNoteCount = oNotes.Count
sText = sText & vbCrLf
' sText = sText & "### Notes" & vbCrLf & vbCrLf
for iNote = 1 to iNoteCount
Set oNote = oNotes.Item(iNote)
Set oShapes = oNote.Shapes
iShapeCount = oShapes.Count
For iShape = 1 To iShapeCount
Set oShape = oShapes.Item(iShape)
sText = ProcessShape(oShape, sText, "Notes", iShape)
Set oShape = Nothing
Next
Set oShapes = Nothing
Set oNote = Nothing
Next
Set oNotes = Nothing
end if

if false then
Set oComments = oSlide.Comments
iCommentCount = oComments.Count
bCommentLabel = True
for iComment = 1 to iCommentCount
Set oComment = oComments.Item(iComment)
sText = ProcessComment(oComment, sText)
Set oComment = Nothing
Next
Set oComments = Nothing
end if

if iHyperlinkCount > 0 then
sText = sText & vbCrLf
sText = sText & "### Links" & vbCrLf & vbCrLf
bHyperlinkLabel = True
for iHyperlink = 1 to iHyperlinkCount
on error resume next
Set oHyperlink = oHyperlinks.Item(iHyperlink)
sText = ProcessHyperlink(oHyperlink, sText)
on error goto 0
Set oHyperlink = Nothing
Next
Set oHyperlinks = Nothing
end if

Set oSlide = Nothing
Next
Set oSlides = Nothing

' sText = RegexpReplace(sText, vbCrLf & vbCrLf & vbCrLf & "+", vbCrLf& vbCrLf, true)
sText = RegexpReplace(sText, "([^(])(https?:\S+)([^)])", "$1<$2>$3", true)
sText = Replace(sText, vbCrLf & "#", vbCrLf & vbCrLf & "#")
sText = StringReplaceAll(sText, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
' FileSaveUTF16 sText, sTargetFile
' ExportPdf oPpt, sTargetPdf
FileWriteUtf8 sTargetFile, sText
Else
' expression.SaveAs(Filename, FileFormat, EmbedFonts)
oPpt.SaveAs sTargetFile, iTargetFormat
if bErrorEvent then bErrorEvent = 0
End If

If FileExists(sTargetFile) Then iConvertCount = iConvertCount + 1
oPpt.Close()
Set oPpt = Nothing
Next
Set oPpts = Nothing

oApp.Quit()
Set oApp = Nothing
If Not bPowerPointExisted And ProcessIsModuleActive(sProcess) Then ProcessClose(sProcess)


Echo("Converted " & iConvertCount & " out of " & StringPlural("file", iSourceCount))
