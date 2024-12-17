Option Explicit
WScript.Echo"Starting DocxProperties"

Dim aRevisions, aIni
Dim bBackupDocx, bLogActions, bValue, bChangeAttachedTemplate, bLinkToContent, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dRevisionTypes, dProperties, d, dStyle, dIni, dSourceIni
Dim iRevision, iRevisionCount, iValue, iLanguageID, iArgCount, iCount
Dim nValue
Dim oFormattedText, oTempDoc, oFile, oSystem, oProperties, oFormat, oFont, oStyle, oApp, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sRevisionText, sType, sOld, sFix, sNew, sHomerLibVbs, sDir, sScriptVbs, sKey, sTargetLog, sBackupDocx, sProperty, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

Const WdDoNotSaveChanges = 0

Const wdEnglishUS = 1033
Const wdMixedLanguage = 9999999

Const msoPropertyTypeBoolean = 2
Const msoPropertyTypeDate = 3
Const msoPropertyTypeFloat = 5
Const msoPropertyTypeNumber = 1 ' Integer
Const msoPropertyTypeString = 4

' View enumeration
Const wdMasterView = 5
Const wdNormalView = 1
Const wdOutlineView = 2
Const wdPrintPreview = 4
Const wdPrintView = 3
Const wdReadingView = 7
Const wdWebView = 6

Function FileInclude(sFile)
' With CreateObject("Scripting.FileSystemObject")
' ExecuteGlobal .openTextFile(sFile).readAll()
' End With

executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

Function CreateRevisionTypesDictionary()
Dim d

Set d = CreateDictionary
d.Add 0, "No revision"
d.Add 7, "Revision marked as a conflict"
d.Add 2, "Deletion"
d.Add 5, "Field display changed"
d.Add 1, "Insertion"
d.Add 4, "Paragraph number changed"
d.Add 10, "Paragraph property changed"
d.Add 3, "Property changed"
d.Add 6, "Revision marked as reconciled conflict"
d.Add 9, "Replaced"
d.Add 12, "Section property changed"
d.Add 8, "Style changed"
d.Add 13, "Style definition changed"
d.Add 11, "Table property changed"
d.Add 17, "Table cell deleted"
d.Add 16, "Table cell inserted"
d.Add 18, "Table cells merged"
d.Add 14, "Content moved from"
d.Add 15, "Content moved to"
Set CreateRevisionTypesDictionary = d
End Function

Function CreateReadabilityDictionary()
Dim dReadability

Set dReadability = CreateDictionary
dReadability.Add "Words", 1
dReadability.Add "Characters", 2
dReadability.Add "Paragraphs", 3
dReadability.Add "Sentences", 4
dReadability.Add "SentencesPerParagraph", 5
dReadability.Add "WordsPerSentence", 6
dReadability.Add "CharactersPerWord", 7
dReadability.Add "PassiveSentences", 8
dReadability.Add "FleschReadingEase", 9
dReadability.Add "Flesch-KincaidGradeLevel", 10
Set CreateReadabilityDictionary = dReadability
End Function

Function GetLanguageID(oDoc)
Dim iLanguageID

oDoc.LanguageDetected = False
iLanguageID = 0
On Error Resume Next
oDoc.Content.DetectLanguage
iLanguageID = oDoc.Content.LanguageID
On Error GoTo 0
If iLanguageID = 0 or iLanguageID = wdMixedLanguage Then
iLanguageID = wdEnglishUS
oDoc.Content.LanguageDetected = True
oDoc.Content.LanguageID = iLanguageID
End If
GetLanguageID = iLanguageID
End Function

Function ShowOther(oDoc)
AppendBlank
AppendLine "[Other]"
AppendLine "AttachedTemplate = " & PathCombine(oDoc.AttachedTemplate.Path, oDoc.AttachedTemplate.Name)
AppendLine "ActiveWritingStyle = " & oDoc.ActiveWritingStyle(iLanguageID)
AppendLine "ActiveTheme = " & oDoc.ActiveTheme
AppendLine "ActiveThemeDisplayName = " & oDoc.ActiveThemeDisplayName
' AppendLine "ApplyQuickStyleSet = " & CBool(oDoc.ApplyQuickStyleSet)
' AppendLine "Compatibility = " & CBool(oDoc.Compatibility)
AppendLine "DefaultTableStyle = " & oDoc.DefaultTableStyle
AppendLine "DefaultTabStop = " & CSng(oDoc.DefaultTabStop)
AppendLine "DoNotEmbedSystemFonts = " & CBool(oDoc.DoNotEmbedSystemFonts)
AppendLine "EmbedTrueTypeFonts = " & CBool(oDoc.EmbedTrueTypeFonts)
AppendLine "Final = " & CBool(oDoc.Final)
AppendLine "JustificationMode = " & CInt(oDoc.JustificationMode)
' PrintPreview is an Application property and a Document method
AppendLine "PrintPreview = " & CBool(oApp.PrintPreview)
AppendLine "ReadOnlyRecommended = " & CBool(oDoc.ReadOnlyRecommended)
AppendLine "RemoveDateAndTime = " & CBool(oDoc.RemoveDateAndTime)
AppendLine "RemoveNumbers = " & CBool(oDoc.RemoveNumbers)
AppendLine "RemovePersonalInformation = " & CBool(oDoc.RemovePersonalInformation)
AppendLine "SaveEncoding = " & CInt(oDoc.SaveEncoding)
AppendLine "SaveSubsetFonts = " & CBool(oDoc.SaveSubsetFonts)
' AppendLine "SetCompatibilityMode = " & CBool(oDoc.SetCompatibilityMode)
AppendLine "ShowGrammaticalErrors = " & CBool(oDoc.ShowGrammaticalErrors)
AppendLine "ShowSpellingErrors = " & CBool(oDoc.ShowSpellingErrors)
AppendLine "StyleSortMethod = " & CInt(oDoc.StyleSortMethod)
AppendLine "TextEncoding = " & CInt(oDoc.TextEncoding)
AppendLine "TextLineEnding = " & CInt(oDoc.TextLineEnding)
AppendLine "TrackFormatting = " & CBool(oDoc.TrackFormatting)
AppendLine "TrackMoves = " & CBool(oDoc.TrackMoves)
AppendLine "TrackRevisions = " & CBool(oDoc.TrackRevisions)
AppendLine "UpdateStylesOnOpen = " & CBool(oDoc.UpdateStylesOnOpen)
AppendLine "View = " & CInt(oDoc.ActiveWindow.View)
' WebPagePreview is a method not property
' AppendLine "WebPagePreview = " & CBool(oDoc.WebPagePreview)
End Function

Function ShowReadabilityStatistics(oDoc)
Dim aStats
Dim dReadability
Dim iStat, iStatistic, iStatisticCount
Dim oStat, oRange, oStatistic
Dim sStat

iStatisticCount = oDoc.ReadabilityStatistics.Count
If iStatisticCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("ReadabilityStatistic", iStatisticCount) & "]"
ArrayClear aStats
ArrayAdd aStats, "Characters"
ArrayAdd aStats, "Words"
ArrayAdd aStats, "Sentences"
ArrayAdd aStats, "Paragraphs"
ArrayAdd aStats, "CharactersPerWord"
ArrayAdd aStats, "WordsPerSentence"
ArrayAdd aStats, "SentencesPerParagraph"
ArrayAdd aStats, "PassiveSentences"
ArrayAdd aStats, "FleschReadingEase"
ArrayAdd aStats, "Flesch-KincaidGradeLevel"

Set dReadability = CreateReadabilityDictionary
For Each sStat in aStats
iStat = dReadability(sStat)
Set oStat = oDoc.ReadabilityStatistics(iStat)
AppendLine oStat.Name & " = " & oStat.Value
Next
End Function

Function ShowMisc(oDoc)
AppendBlank
AppendLine "[; Miscellaneous]"
' AppendLine StringPlural("character", oDoc.Characters.Count)
' AppendLine StringPlural("word", oDoc.words.Count)
if oDoc.Sections.Count > 1 then AppendLine StringPlural("section", oDoc.sections.Count)
AppendLine "LanguageID = " & iLanguageID
End Function

Function ShowTables(oDoc)
Dim iTableCount
Dim oTable

iTableCount = oDoc.Tables.Count
if iTableCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("table", iTableCount) & "]"
For Each oTable in oDoc.Tables
AppendLine "(" & oTable.Title & ")"
Next
End Function

Function ShowShapes(oDoc)
Dim iShapeCount
Dim oShape
iShapeCount = oDoc.Shapes.Count
if iShapeCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("Shape", iShapeCount) & "]"
For Each oShape in oDoc.Shapes
' AppendLine oShape.Title
' AppendLine oShape.Name
' AppendLine oShape.AlternativeText
AppendLine Trim(oShape.Title & " (" & oShape.AlternativeText) & ")"
Next
End Function

Function ShowInlineShapes(oDoc)
Dim iInlineShapeCount
Dim oInlineShape
iInlineShapeCount = oDoc.InlineShapes.Count
if iInlineShapeCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("InlineShape", iInlineShapeCount) & "]"
For Each oInlineShape in oDoc.InlineShapes
AppendLine Trim(oInlineShape.Title & " (" & oInlineShape.AlternativeText) & ")"
' AppendLine oInlineShape.Name
' AppendLine oInlineShape.AlternativeText
Next
End Function

Function ShowRevisions(oDoc)
Dim iRevision, iRevisionCount
Dim dRevisionTypes
Dim oRevision
Dim sDescription

iRevisionCount = oDoc.Revisions.Count
if iRevisionCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("Revision", iRevisionCount) & "]"

aRevisions = Array()
Set dRevisionTypes = CreateRevisionTypesDictionary
Set oFormattedText = oDoc.Content.FormattedText
Set oTempDoc = oApp.Documents.Add
oTempDoc.Content.FormattedText = oFormattedText
iRevisionCount = oTempDoc.Revisions.Count
For iRevision = oTempDoc.Revisions.Count to 1 Step -1
Set oRevision = oDoc.Revisions(iRevision)
sType = Trim(oRevision.FormatDescription)
If Len(sType) = 0 Then sType = dRevisionTypes(oRevision.Type)
sFix = oRevision.Range.Text
sNew = StringTrimWhiteSpace(oRevision.Range.Sentences(1).Text)
Set oRange = oRevision.Range
oRevision.Reject
sOld = StringTrimWhiteSpace(oRange.Sentences(1).Text)

sRevisionText = sType & vbCrLf
sRevisionText = sRevisionText & "Old: " & sOld & vbCrLf
sRevisionText = sRevisionText & "Fix: " & sFix & vbCrLf
sRevisionText = sRevisionText & "New: " & sNew
ArrayAdd aRevisions, sRevisionText
Next

oTempDoc.Close 0
oDoc.Activate
For iRevision = (iRevisionCount - 1) to 0 Step -1
AppendLine (iRevisionCount - iRevision) & ". " & aRevisions(iRevision)
If iRevision > 0 Then AppendBlank
Next
End Function

Function ShowBookmarks(oDoc)
Dim iBookmarkCount
Dim oBookmark
iBookmarkCount = oDoc.Bookmarks.Count
if iBookmarkCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("Bookmark", iBookmarkCount) & "]"
For Each oBookmark in oDoc.Bookmarks
AppendLine oBookmark.Name
Next
End Function

Function ShowIndexes(oDoc)
Dim iIndexCount
Dim oIndex
iIndexCount = oDoc.Indexes.Count
if iIndexCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("Index", iIndexCount) & "]"
For Each oIndex in oDoc.Indexes
AppendLine oIndex.Range.Text
Next
End Function


Function ShowTablesOfContents(oDoc)
Dim iTableOfContentsCount
Dim oTableOfContents
Dim s

iTableOfContentsCount = oDoc.TablesOfContents.Count
if iTableOfContentsCount = 0 Then Exit Function
AppendBlank
s = "table of contents"
if iTableOfContentsCount <> 1 then s = "tables of contents"
AppendLine "[; " & iTableOfContentsCount & " " & s & "]"
' AppendLine "[; " & StringPlural("TablesOfContents", iTableOfContentsCount) & "]"
For Each oTableOfContents in oDoc.TablesOfContents
AppendLine oTableOfContents.Range.Text
' AppendLine StringConvertToWinLineBreak(oTableOfContents.Range.Text)
Next
End Function

Function ShowTablesOfFigures(oDoc)
Dim iTableOfFiguresCount
Dim oTableOfFigures
Dim s

iTableOfFiguresCount = oDoc.TablesOfFigures.Count
if iTableOfFiguresCount = 0 Then Exit Function
AppendBlank
s = "table of Figures"
if iTableOfFiguresCount <> 1 then s = "tables of Figures"
AppendLine "[; " & iTableOfFiguresCount & " " & s & "]"
For Each oTableOfFigures in oDoc.TablesOfFigures
AppendLine oTableOfFigures.Range.Text
Next
End Function

Function ShowTablesOfAuthorities(oDoc)
Dim iTableOfAuthoritiesCount
Dim oTableOfAuthorities
Dim s

iTableOfAuthoritiesCount = oDoc.TablesOfAuthorities.Count
if iTableOfAuthoritiesCount = 0 Then Exit Function
AppendBlank
s = "table of Authorities"
if iTableOfAuthoritiesCount <> 1 then s = "tables of Authorities"
AppendLine "[; " & iTableOfAuthoritiesCount & " " & s & "]"
For Each oTableOfAuthorities in oDoc.TablesOfAuthorities
AppendLine oTableOfAuthorities.Range.Text
Next
End Function

Function ShowHeadings(oDoc)
Dim aHeadings
Dim dHeadings
Dim iHeadingLevel, iHeadingCount
Dim oParagraph
Dim sHeading

' Set dHeadings = CreateDictionary
ArrayClear aHeadings
For Each oParagraph in oDoc.Paragraphs
' If StringLead(oParagraph.Style, "Heading ", False) Then dHeadings.Add oParagraph.Range.Text, oParagraph.OutlineLevel
If StringLead(oParagraph.Style, "Heading ", False) Then ArrayAdd aHeadings, oParagraph.OutlineLevel & ". " & StringTrimWhiteSpace(oParagraph.Range.Text)
Next

' iHeadingCount = dHeadings.Count
iHeadingCount = ArrayCount(aHeadings)
If iHeadingCount = 0 Then Exit Function

AppendBlank
AppendLine "[; " & StringPlural("heading", iHeadingCount) & "]"
' For Each sHeading in dHeadings.Keys
For Each sHeading in aHeadings
' iHeadingLevel = dHeadings(sHeading)
' AppendLine iHeadingLevel & ". " & sHeading
AppendLine sHeading
Next
End Function

Function ShowFootNotes(oDoc)
Dim iFootNoteCount
Dim oFootNote
iFootNoteCount = oDoc.FootNotes.Count
if iFootNoteCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("FootNote", iFootNoteCount) & "]"
For Each oFootNote in oDoc.FootNotes
' AppendLine oFootNote.Reference.Text
AppendLine oFootNote.Range.Text
Next
End Function

Function ShowEndNotes(oDoc)
Dim iEndNoteCount
Dim oEndNote
iEndNoteCount = oDoc.EndNotes.Count
if iEndNoteCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("EndNote", iEndNoteCount) & "]"
For Each oEndNote in oDoc.EndNotes
AppendLine oEndNote.Reference.Text
AppendLine oEndNote.Range.Text
Next
End Function

Function ShowFields(oDoc)
Dim iFieldCount
Dim oField
iFieldCount = oDoc.Fields.Count
if iFieldCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("Field", iFieldCount) & "]"
For Each oField in oDoc.Fields
AppendLine oField.Code.Text
AppendLine oField.Result.Text
Next
End Function

Function ShowLinks(oDoc)
Dim iLinkCount
Dim oLink
iLinkCount = oDoc.Hyperlinks.Count
if iLinkCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("link", iLinkCount) & "]"
For Each oLink in oDoc.HyperLinks
On Error Resume Next
AppendLine Trim(oLink.TextToDisplay & "(" & oLink.Address) & ")"
On Error GoTo 0
Next
End Function

Function ShowComments(oDoc)
Dim iComment, iCommentCount
Dim oComment
iCommentCount = oDoc.Comments.Count
if iCommentCount = 0 Then Exit Function
AppendBlank
AppendLine "[; " & StringPlural("Comment", iCommentCount) & "]"
iComment = 0
For Each oComment in oDoc.Comments
iComment = iComment + 1
' On Error Resume Next
AppendLine iComment & ". " & StringTrimWhiteSpace(oComment.Scope.Text)
AppendLine Trim(oComment.Range.Text)
' AppendLine iComment & ". " & "Scope: " & Trim(oComment.Scope.Text)
' AppendLine "Context: " & StringTrimWhiteSpace(oComment.Scope.Sentences(1).Text)
' AppendLine "Comment: " & Trim(oComment.Range.Text)
' On Error GoTo 0
If iComment < iCommentCount Then AppendLine ""
Next
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
bBackupDocx = GetGlobalValue(dSourceIni, "BackupDocx", True)
bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)

bReadOnly = False
ProcessTerminateAllModule "WinWord"
Else
sSourceIni = ""
Set dSourceIni = CreateDictionary()
bReadOnly = True
End If

sTargetIni = PathCombine(PathGetCurrentDirectory(), "PROPERTIES-" & PathGetRoot(sSourceDocx) & ".ini")
sTargetLog = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceDocx) & "-DocxProperties.log")

Set oApp = CreateObject("Word.Application")
' oApp.Visible = False
oApp.Visible = True
' oApp.DisplayAlerts = False
oApp.DisplayAlerts = 0
oApp.ScreenUpdating = False
Set oDocs = oApp.Documents
bAddToRecentFiles = False
bConfirmConversions = False
Print "Opening " & PathGetName(sSourceDocx)
Set oDoc = oDocs.Open(sSourceDocx, bAddToRecentFiles, bReadOnly, bConfirmConversions)
iLanguageID = GetLanguageID(oDoc)

If not bReadOnly Then
 Print "Applying " & PathGetName(sSourceIni)
If dSourceIni.Exists("BuiltIn") Then
Set dProperties = dSourceIni("BuiltIn")
Set oProperties = oDoc.BuiltInDocumentProperties
For Each oProperty In oProperties
sProperty = oProperty.Name
If dProperties.Exists(sProperty) Then
sValue = dProperties(sProperty)
Print "Setting property " & sProperty & " = " & sValue
oProperty.Value = sValue
End If 'dProperties.Exists(sProperty
Next
End If ' dSourceIni.Exists("BuiltIn")  

If dSourceIni.Exists("Custom") Then
Set dProperties = dSourceIni("Custom")
' Update existing custom properties
For Each sProperty in dProperties.Keys
sValue = dProperties(sProperty)
Set oProperties = oDoc.CustomDocumentProperties
For Each oProperty in oProperties
If oProperty.Name = sProperty Then
If Len(sValue) = 0 Then
Print "Deleting property " & sProperty
oProperty.Delete
Else
Print "Setting property " & sProperty & " = " & sValue
oProperty.Value = sValue
End If ' Len(sValue) = 0
dProperties.Remove sProperty
End If ' oProperty.Name = sProperty
Next
Next

' Add custom properties
For Each sProperty in dProperties.Keys
sValue = dProperties(sProperty)
Print "Adding property " & sProperty & "= " & sValue
oProperties.Add sProperty, bLinkToContent, msoPropertyTypeString, sValue
Next
End If ' dSourceIni.Exists("Custom")  

If dSourceIni.Exists("Other") Then
Set dProperties = dSourceIni("Other")
For Each sProperty in dProperties.Keys
sKey = sProperty
sValue = dProperties(sProperty)
bValue = ForceBool(sValue)
iValue = ForceInt(sValue)
nValue = ForceSng(sValue)
Select Case sProperty
Case "ActiveWritingStyle"
' Returns or sets the writing style for a specified language in the specified document. Read/write String
print "Setting " & sKey & " = " & sValue
oDoc.ActiveWritingStyle(iLanguageID) = sValue
Case "ActiveTheme"
' Applies or removes a theme
if Len(sValue) = 0 or LCase(sValue) = "none" Then
print "RemoveTheme"
oDoc.RemoveTheme
Else
print "Setting ActiveTheme = " & sValue
oDoc.ApplyTheme sValue
End If
Case "ApplyQuickStyleSet"
' Applies the specified StyleSet to the document
print "Setting QuickStyleSet = " & sValue
oDoc.ApplyQuickStyleSet2 sValue
Case "AttachedTemplate"
' Returns or sets the attached template
print "Setting AttachedTemplate = " & sValue
oDoc.AttachedTemplate = sValue
Case "Compatibility"
' True if the compatibility option specified by the Type argument is enabled. Compatibility options affect how a document is displayed in Microsoft Word. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.Compatibility = bValue
Case "DefaultTableStyle"
' Returns a Variant that represents the table style that is applied to all newly created tables in a document. Read-only
print "Setting " & sKey & " = " & iValue
bChangeAttachedTemplate = False
oDoc.SetDefaultTableStyle sValue, bChangeAttachedTemplate
Case "DefaultTabStop"
' Returns or sets the interval (in points) between the default tab stops in the specified document. Read/write Single
print "Setting " & sKey & " = " & nValue
oDoc.DefaultTabStop = nValue
Case "DoNotEmbedSystemFonts"
' True for Microsoft Word to not embed common system fonts. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.DoNotEmbedSystemFonts = bValue
Case "EmbedTrueTypeFonts"
' True if Microsoft Word embeds TrueType fonts in a document when it is saved. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.EmbedTrueTypeFonts = bValue
Case "Final"
' Returns or sets a Boolean that indicates whether a document is final. Read/write
print "Setting " & sKey & " = " & bValue
oDoc.Final = bValue
Case "JustificationMode"
' Returns or sets the character spacing adjustment for the specified document. Read/write WdJustificationMode
print "Setting " & sKey & " = " & iValue
oDoc.JustificationMode = iValue
Case "Password"
' Sets a password that must be supplied to open the specified document. Write-only String
Print "Setting Password = " & sValue
oDoc.Password = sValue
' PrintPreview is an Application property and a Document method
Case "PrintPreview"
' Switches the view to print preview
' print "Setting " & sKey & " = " & bValue
oApp.PrintPreview = bValue
Case "ReadOnlyRecommended"
' True if Microsoft Word displays a message box whenever a user opens the document, suggesting that it be opened as read-only. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.ReadOnlyRecommended = bValue
Case "RemoveDateAndTime"
' Sets or returns a Boolean indicating whether a document stores the date and time metadata for tracked changes.
print "Setting " & sKey & " = " & bValue
oDoc.RemoveDateAndTime = bValue
Case "RemoveNumbers"
' Removes numbers or bullets from the specified document
if bValue then
print "RemoveNumbers"
oDoc.RemoveNumbers
end if
Case "RemovePersonalInformation"
' True if Microsoft Word removes all user information from comments, revisions, and the Properties dialog box upon saving a document. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.RemovePersonalInformation = bValue
Case "SaveEncoding"
' Returns or sets the encoding to use when saving a document. Read/write MsoEncoding
print "Setting " & sKey & " = " & iValue
oDoc.SaveEncoding = iValue
Case "SaveSubsetFonts"
' True if Microsoft Word saves a subset of the embedded TrueType fonts with the document. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.SaveSubsetFonts = bValue
Case "SetCompatibilityMode"
' Sets the compatibility mode for the document
print "Setting " & sKey & " = " & iValue
oDoc.CompatibilityMode = iValue
Case "ShowGrammaticalErrors"
' True if grammatical errors are marked by a wavy green line in the specified document. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.ShowGrammaticalErrors = bValue
Case "ShowSpellingErrors"
' True if Microsoft Word underlines spelling errors in the document. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.ShowSpellingErrors = bValue
Case "StyleSortMethod"
' Returns or sets aWdStyleSort constant that represents the sort method to use when sorting styles in the Styles task pane. Read/write
print "Setting " & sKey & " = " & iValue
oDoc.StyleSortMethod = iValue
Case "TextEncoding"
' Returns or sets the code page, or character set, that Microsoft Word uses for a document saved as an encoded text file. Read/write MsoEncoding
print "Setting " & sKey & " = " & iValue
oDoc.TextEncoding = iValue
Case "TextLineEnding"
' Returns or sets a WdLineEndingType constant indicating how Microsoft Word marks the line and paragraph breaks in documents saved as text files. Read/write
print "Setting " & sKey & " = " & iValue
oDoc.TextLineEnding = iValue
Case "TrackFormatting"
' Returns or sets a Boolean that represents whether to track formatting changes when change tracking is turned on. Read/write
print "Setting " & sKey & " = " & bValue
oDoc.TrackFormatting = bValue
Case "TrackMoves"
' Returns or sets a Boolean that represents whether to mark moved text when Track Changes is turned on. Read/write
print "Setting " & sKey & " = " & bValue
oDoc.TrackMoves = bValue
Case "TrackRevisions"
' True if changes are tracked in the specified document. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.TrackRevisions = bValue
Case "UpdateStylesOnOpen"
' True if the styles in the specified document are updated to match the styles in the attached template each time the document is opened. Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.UpdateStylesOnOpen = bValue
Case "View"
Print "Setting View = " & iValue
oDoc.ActiveWindow.View = iValue
Case "WebPagePreview"
' WebPagePreview is a method not property
' Displays a preview of the current document as it would look if saved as a Web page
print "Setting " & sKey & " = " & bValue
' oDoc.WebPagePreview = bValue
If bValue Then oDoc.WebPagePreview
End Select ' sProperty
Next ' sProperty
End If ' dSourceIni.Exists("Other")  

If Not oDoc.Saved Then
If bBackupDocx Then
sBackupDocx = FileBackup(sSourceDocx)
If Len(sBackupDocx) = 0 Then
Print "Error creating backup "
Else
Print "Creating backup " & PathGetName(sBackupDocx)
End If
End If

If not bBackupDocx or Len(sBackupDocx) > 0 Then
print "Saving " & PathGetName(sSourceDocx)
oDoc.Save
End If
End If
End If ' Not bReadOnly

' Create target ini
Print "Creating " & PathGetName(sTargetIni)
AppendLine "[BuiltIn]"
Set oProperties = oDoc.BuiltInDocumentProperties
For Each oProperty In oProperties
' Avoid error on LastPrintDate
On Error Resume Next
' print oProperty.Name
AppendLine oProperty.Name & "= " & oProperty.Value
On Error GoTo 0
Next
AppendBlank
AppendLine "[Custom]"
For Each oProperty In oDoc.CustomDocumentProperties
On Error Resume Next
AppendLine oProperty.Name & "= " & oProperty.Value
On Error GoTo 0
Next

ShowOther(oDoc)
ShowBookmarks(oDoc)
ShowComments(oDoc)
ShowEndNotes(oDoc)
ShowFields(oDoc)
ShowFootNotes(oDoc)
ShowHeadings(oDoc)
ShowIndexes(oDoc)
ShowInlineShapes(oDoc)
ShowLinks(oDoc)
ShowReadabilityStatistics(oDoc)
ShowRevisions(oDoc)
ShowShapes(oDoc)
ShowTables(oDoc)
ShowTablesOfAuthorities(oDoc)
ShowTablesOfContents(oDoc)
ShowTablesOfFigures(oDoc)
ShowMisc(oDoc)

' On Error Resume Next
' oApp.PrintPreview = False
If oDoc.PrintPreview Then oDoc.ClosePrintPreview
oDoc.ActiveWindow.View.Type = wdNormalView
if oDocs.Count > 0 Then oDocs.Close wdDoNotSaveChanges
' On Error GoTo 0
If Not oApp.NormalTemplate.Saved Then oApp.NormalTemplate.Save
oApp.Quit

StringToFile sHomerText, sTargetIni

If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
End If

echo "Done"
