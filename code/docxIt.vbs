Option Explicit
WScript.Echo"Starting DocxIt"

Dim a, aStyles, aIni
Dim bLoop, bBackupDocx, bLogActions, bValue, bFound, bAddToRecentFiles, bConfirmConversions, bIncludePageNumbers, bHidePageNumbersInWeb, bRightAlignPageNumbers, bUseFields, bUseHeadingStyles, bUseHyperlinks, bUseOutlineLevels, bReadOnly
Dim bFormat, bForward, bMatchAlefHamza, bMatchAllWordForms, bMatchCase, bMatchControl, bMatchDiacritics, bMatchKashida, bMatchSoundsLike, bMatchWholeWord, bMatchWildcards
Dim d, dHeadingStyles, dStyle, dIni, dSourceIni, dSection
Dim iValue, i, iLevel, iReplaceCount, iTableId, iReplace, iWrap, iForward, iArgCount, iCount, iLowerHeadingLevel, iUpperHeadingLevel
Dim oTables, oTable, oFindFormat, oFindFont, oReplaceFormat, oReplaceFont, oSystem, oFile, oParagraph, oField, oAddedStyles, oApp, oData, oDoc, oDocs, oFind, oFont, oFormat, oProperty, oRange, oReplace, oStyle, oStyles, oToc, oTocs
Dim nValue
Dim sBackupDocx, sTargetLog, sScriptVbs, sHomerLibVbs, sDir, sCode, sFindStyle, sReplaceStyle, sKey, sFind, sFindText, sReplaceText, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni, sSection

' wdAutoFit enumeration
Const wdAutoFitContent = 1 'The table is automatically sized to fit the content contained in the table.
Const wdAutoFitFixed = 0 'The table is set to a fixed size, regardless of the content, and is not automatically sized.
Const wdAutoFitWindow = 2 'The table is automatically sized to the width of the active window.
Const wdFormatDocumentDefault = 16

' wdStyleType enumeration
Const wdStyleTypeParagraph = 1
Const wdStyleTypeCharacter = 2
Const wdStyleTypeTable = 3
Const wdStyleTypeList = 4

' wdOrganizerObject enumeration
Const wdOrganizerObjectStyles = 0
Const wdOrganizerObjectAutoText = 1
Const wdOrganizerObjectCommandBars = 2
Const wdOrganizerObjectProjectItems = 3

Const wdOutlineLevel1 = 1 ' Outline level 1
Const wdOutlineLevelBodyText = 10 'No outline level

Const wdRDIAll = 99 
' Removes all document information.
Const wdRDIComments = 1 
' Removes document comments.
Const wdRDIContentType = 16 
' Removes content type information.
Const wdRDIDocumentManagementPolicy = 15 
' Removes document management policy information.
Const wdRDIDocumentProperties = 8 
' Removes document properties.
Const wdRDIDocumentServerProperties = 14 
' Removes document server properties.
Const wdRDIDocumentWorkspace = 10 
' Removes document workspace information.
Const wdRDIEmailHeader = 5 
' Removes e-mail header information.
Const wdRDIInkAnnotations = 11 
' Removes ink annotations.
Const wdRDIRemovePersonalInformation = 4 
' Removes personal information.
Const wdRDIRevisions = 2 
' Removes revision marks.
Const wdRDIRoutingSlip = 6 
' Removes routing slip information.
Const wdRDISendForReview = 7 
' Removes information stored when sending a document for review.
Const wdRDITaskpaneWebExtensions = 17 
' Removes taskpane web extensions information.
Const wdRDITemplate = 9 
' Removes template information.
Const wdRDIVersions = 3 
' Removes document version information.

Const wdNoProtection = -1
 
Const WdCollapseEnd = 0
Const WdDoNotSaveChanges = 0

Const wdFindContinue = 1

Const wdReplaceOne = 1
Const wdReplaceAll = 2

Function FileInclude(sFile)
With CreateObject("Scripting.FileSystemObject")
ExecuteGlobal .openTextFile(sFile).readAll()
End With

' executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

Function CopyStyles(oStylesLocal, sSourceFile, sTargetFile)
Dim oStyleLocal

print "Source " & PathGetName(sSourceFile)
print "Target " & PathGetName(sTargetFile)
For Each oStyleLocal in oStylesLocal
' if oStyle.InUse And oStyle.Type = wdStyleTypeParagraph Then
if oStyleLocal.InUse Then
' print oStyleLocal.NameLocal
oApp.OrganizerCopy sSourceFile, sTargetFile, oStyleLocal.NameLocal, wdOrganizerObjectStyles 
End If
Next ' For Each oStyle in oDoc.Styles
End Function

Function DeleteUnusedStyles(oDoc)
Dim bDelete
Dim oFind, oStyleLocal

For Each oStyleLocal In oDoc.Styles
bDelete = False
If not oStyleLocal.BuiltIn Then
If not oStyle.InUse Then
bDelete = True
Else
oApp.Selection.HomeKey wdStory
Set oFind = oApp.Selection.Find
oFind.ClearFormatting
oFind.Style = oStyleLocal
oFind.Text = ""
oFind.Replacement.Text = ""
oFind.Forward = True
oFind.Wrap = wdFindStop
oFind.Format = True
oFind.MatchCase = False
oFind.MatchWholeWord = False
oFind.MatchWildcards = False
oFind.MatchSoundsLike = False
oFind.MatchAllWordForms = False

oFind.Execute
If not oFind.Found Then bDelete = True
End If
If bDelete Then
print oStyleLocal.NameLocal
 oStyle.Delete
End If
End If
Next ' oStyle
End Function

Function FixOutline()
' Fix paragraph outline levels when not heading styles
For Each oParagraph In oDoc.Content.Paragraphs
sStyle = oParagraph.Style.NameLocal
If Not StringLead(sStyle, "Heading", False) and oParagraph.OutlineLevel <> wdOutlineLevelBodyText Then
Print "Setting style " & sStyle & ", OutlineLevel to Body Text"
oParagraph.OutlineLevel = wdOutlineLevelBodyText  
End If
Next

End Function

Function FixToc(oDoc, oToc, dHeadingStyles, bUseHyperLinks, bUseOutlineLevels)
' Remove any additional styles so they are only added explicitly
For Each oStyle in oToc.HeadingStyles
oStyle.Delete
Next

' Add explicit styles
For Each sStyle in dHeadingStyles.Keys
iLevel = dHeadingStyles(sStyle)
oToc.HeadingStyles.Add sStyle, iLevel
Next

' Fix the TOC field code if UseHyperlinks and UseOutlineLevels are misconfigured
For Each oField in oDoc.Fields
sCode = oField.Code.Text
If InStr(sCode, "TOC ") Then
If bUseHyperLinks and Not InStr(sCode, " \h") Then sCode = sCode & " \h"
If Not bUseOutlineLevels and InStr(sCode, " \u") Then sCode = Replace(sCode, " \u", "")
oField.Code.Text = sCode
Exit For
End If
Next
' FixOutline
oToc.Update
End Function

' Main
sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile) 
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs
' FileInclude "HomerLib.vbs"

iArgCount = WScript.Arguments.Count

If iArgCount < 2 Then Quit "Specify a source .docx file as the first parameter an a configuration .ini file as the second parameter."

sSourceDocx = WScript.Arguments(0)
' If Not InStr(sSourceDocx, "\") Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If InStr(sSourceDocx, "\") = 0 Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If not FileExists(sSourceDocx) Then Quit "Cannot find " & sSourceDocx

sSourceIni = WScript.Arguments(1)
sSourceIni = GetIniFile(sSourceIni)
If InStr(sSourceIni, "\") = 0 Then sSourceIni = PathCombine(PathGetCurrentDirectory(), sSourceIni)
If not FileExists(sSourceIni) Then Quit "Cannot find " & sSourceIni
Set dSourceIni = IniToDictionary(sSourceIni)
bBackupDocx = GetGlobalValue(dSourceIni, "BackupDocx", True)
bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)

sTargetLog = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceDocx) & "-DocxIt.log")

ProcessTerminateAllModule "WinWord"
Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False

Set oDocs = oApp.Documents
bAddToRecentFiles = False
bReadOnly = False
bConfirmConversions = False
Print "Opening " & PathGetName(sSourceDocx)
Set oDoc = oDocs.Open(sSourceDocx, bAddToRecentFiles, bReadOnly, bConfirmConversions)
' Make this optional
' If oDoc.ProtectionType <> wdNoProtection Then
' Print "Unprotecting document"
' oDoc.Unprotect
' End If
' Make this optional
' On Error Resume Next
' oDoc.Convert
' On Error Goto 0

If Not bReadOnly Then
Print "Applying " & PathGetName(sSourceIni)
' Apply configuration changes
For Each sSection in dSourceIni.Keys
Set dSection = dSourceIni(sSection)
sValue = dSection(sSection)
Select Case sSection
Case "Tasks"
Set dSection = dSourceIni("Tasks")
For Each sKey in dSection.Keys
sValue = dSection(sKey)
bValue = ForceBool(sValue)
iValue = ForceInt(sValue)
nValue = ForceSng(sValue)
oDoc.Content.Select
Set oRange = oApp.ActiveWindow.Selection
Select Case sKey
Case "SwapWithEndnotes"
If bValue Then
print "SwapWithEndnotes"
oDoc.Footnotes.SwapWithEndnotes
End If
Case "SwapWithFootnotes"
If bValue Then
print "SwapWithFootnotes"
oDoc.Endnotes.SwapWithFootnotes
End If
Case "UnprotectDocument"
If bValue Then
print "UnprotectDocument"
If oDoc.ProtectionType <> wdNoProtection Then
' Print "Unprotecting document"
oDoc.Unprotect
End If
End If
Case "UpgradeFormat"
If bValue Then
print "UpgradeFormat"
On Error Resume Next
oDoc.Convert
On Error GoTo 0
End If
Case "AcceptAllRevisions"
' Accepts all tracked changes in the specified document
If bValue Then
Print "AcceptAllRevisions"
If oDoc.Revisions.Count > 0 Then oRange.AcceptAllRevisions
End If
Case "AddToFavorites"
' Creates a shortcut to the document or hyperlink and adds it to the Favorites folder
If bValue Then
print "AddToFavorites"
oDoc.AddToFavorites
End If
Case "AttachedTemplate"
' Returns a Template object that represents the template attached to the specified document. Read/write Variant
If Len(sValue) > 0 Then
print "Setting " & sKey & " = " & sValue
oDoc.AttachedTemplate = sValue
End If
Case "AutoFormatDocument"
' Automatically formats a document
If bValue Then
print "AutoFormatDocument"
oDoc.AutoFormat
End If
Case "ClearCharacterDirectFormatting"
If bValue Then
Print "ClearCharacterDirectFormatting"
oRange.ClearCharacterDirectFormatting
End If
Case "ClearCharacterStyle"
If bValue Then
Print "ClearCharacterStyle"
oRange.ClearCharacterStyle
End If
Case "ClearParagraphDirectFormatting"
If bValue Then
Print "ClearParagraphDirectFormatting"
oRange.ClearParagraphDirectFormatting
End If
Case "ClearParagraphStyle"
If bValue Then
Print "ClearParagraphStyle"
oRange.ClearParagraphStyle
End If
Case "ConvertNumbersToText"
' Changes the list numbers and LISTNUM fields in the specified Document to text
If bValue Then
print "ConvertNumbersToText"
oDoc.ConvertNumbersToText
End If
Case "DeleteAllComments"
' Deletes all comments from the Comments collection in a document
If bValue Then
Print "DeleteAllComments"
If oDoc.Comments.Count > 0 Then oDoc.DeleteAllComments
End If
Case "DetectLanguage"
' Analyzes the specified text to determine the language that it is written in
If bValue Then
print "DetectLanguage"
oDoc.DetectLanguage
End If
Case "FitToPages"
' Decreases the font size of text just enough so that the document will fit on one fewer pages
if bValue Then
print "FitToPages"
oDoc.FitToPages
End If
Case "DocumentKind"
' Returns or sets the format type that Microsoft Word uses when automatically formatting the specified document. Read/write WdDocumentKind
If iValue <> 0 Then
print "Setting " & sKey & " = " & iValue
oDoc.Kind = iValue
End If
Case "ListCommands"
if bValue then
' print "ListCommands"
print "Saving WordCommands.docx"
oDoc.ListCommands
oApp.ActiveDocument.SaveAs "WordCommands.docx"
oApp.ActiveDocument.Close wdDoNotSaveChanges
oDoc.Activate
End if
Case "PrintOut"
' Prints all or part of the specified document
if bValue then
print "Printing"
oDoc.Printout
end if
Case "PrintRevisions"
' True if revision marks are printed with the document. False if revision marks aren't printed (that is, tracked changes are printed as if they'd been accepted). Read/write Boolean
print "Setting " & sKey & " = " & bValue
oDoc.PrintRevisions = bValue
Case "ProtectDocument"
' Protects the specified document from unauthorized changes
if bValue then
print "ProtectDocument"
oDoc.Protect
End if
Case "RejectAllRevisions"
' Rejects all tracked changes in the specified document
If bValue Then
Print "RejectAllRevisions"
If oDoc.Revisions.Count > 0 Then oRange.RejectAllRevisions
End If
Case "RemoveDateAndTime"
' Sets or returns a Boolean indicating whether a document stores the date and time metadata for tracked changes.
print "Setting " & sKey & " = " & bValue
oDoc.RemoveDateAndTime = bValue
Case "RemoveDocumentInformation"
' Removes sensitive information, properties, comments, and other metadata from a document
If bValue Then
Print "RemoveDocumentInformation"
oDoc.RemoveDocumentInformation wdRDIAll
End If
Case "RemoveNumbers"
' Removes numbers or bullets from the specified document
if bValue then
print "RemoveNumbers"
oDoc.RemoveNumbers
end if
Case "Repaginate"
' Repaginates the entire document
if bValue then
print "Repaginate"
oDoc.Repaginate
end if
Case "CopyStyles"
print "CopyStyles"
s = sSourceDocx
If InStr(sValue, ";") Then
a = Split(sValue, ";")
s = a(0)
sValue = a(1)
End If
CopyStyles oDoc.Styles, s, sValue
Case "ApplyQuickStyleSet"
If Len(sValue) > 0 Then
print "ApplyQuickStyleSet"
print "Source " & PathGetName(sValue)
oDoc.ApplyQuickStyleSet2 sValue
End If
Case "SaveAsQuickStyleSet"
If Len(sValue) > 0 Then
print "SaveAsQuickStyleSet"
print "Target " & PathGetName(sValue)
oDoc.SaveAsQuickStyleSet sValue
End If
Case "FormatTables"
' ApplyStyleHeadingRows, True for Microsoft Word to apply heading-row formatting to the first row of the selected table. Read/write Boolean.
if bValue then
Print "FormatTables"
For Each oTable in oDoc.Tables
oTable.AllowAutoFit = True
oTable.AutoFitBehavior wdAutoFitContent
' oTable.Columns.AutoFit
oTable.Rows(1).HeadingFormat = True
oTable.ApplyStyleHeadingRows = True
oDoc.Bookmarks.Add "ColumnTitle", oTable.Rows(1).Cells(2).Range
Next
end if

Case "UnprotectDocument"
' Removes protection from the specified document.
if bValue then
print "UnprotectDocument"
oDoc.Unprotect
end if
Case "RemoveLockedStyles"
If bValue Then
print "RemoveLockedStyles"
oDoc.RemoveLockedStyles
End If
Case "LockQuickStyleSet"
' Returns or sets a Boolean that represents whether users can change which set of Quick Styles is being used.
print "Setting LockQuickStyleSet = " & bValue
oDoc.LockQuickStyleSet = bValue
Case "DeleteUnusedStyles"
If bValue Then
print "DeleteUnusedStyles"
DeleteUnusedStyles oDoc
End If
Case "UpdateStyles"
' Copies all styles from the attached template into the document, overwriting any existing styles in the document that have the same name
If bValue Then
print "UpdateStyles"
print "Source " & PathGetName(oDoc.AttachedTemplate.FullName)
oDoc.UpdateStyles
end if
Case "CopyStylesFromTemplate"
' Copies all styles from the attached template into the document, overwriting like styles and adding unique template styles
If Len(sValue) > 0 Then
print "CopyStylesFromTemplate"
If sValue = "NormalTemplate" Then sValue = oApp.NormalTemplate.FullName
If sValue = "AttachedTemplate" Then sValue = oDoc.AttachedTemplate.FullName
print "Source " & PathGetName(sValue)
oDoc.CopyStylesFromTemplate sValue
end if
Case "LeftMargin"
If iValue <> 0 Then
print "Setting " & sKey & " = " & iValue
oDoc.PageSetup.LeftMargin = oApp.InchesToPoints(iValue)
end if
Case "RightMargin"
If iValue <> 0 Then
print "Setting " & sKey & " = " & iValue
oDoc.PageSetup.RightMargin = oApp.InchesToPoints(iValue)
end if
Case "TopMargin"
If iValue <> 0 Then
print "Setting " & sKey & " = " & iValue
oDoc.PageSetup.TopMargin = oApp.InchesToPoints(iValue)
end if
Case "BottomMargin"
If iValue <> 0 Then
print "Setting " & sKey & " = " & iValue
oDoc.PageSetup.BottomMargin = oApp.InchesToPoints(iValue)
end if
End Select ' sKey
Next ' sKey
Case "TOC"
Set oRange = oDoc.Content
Set oFind = oRange.Find
oFind.ClearFormatting
sFind = "Table of Contents"
If dSection.Exists("Title") Then sFind = dSection("Title")
sFindText = sFind & "^p"
If Not oFind.Execute(sFindText) Then
Print "Cannot find TOC title " & sFind
Else
Print "Generating table of contents"
DeleteObject oDoc.Bookmarks, "toc"
oDoc.Bookmarks.Add "toc", oRange
oRange.Collapse WdCollapseEnd
Set oTocs = oDoc.TablesOfContents
Set dHeadingStyles = CreateDictionary
iLowerHeadingLevel = 2
iUpperHeadingLevel = 1
For Each sKey in dSection.Keys
sValue = dSection(sKey)
Select Case sKey
Case "HeadingStyles"
aStyles = Split(sValue, ";")
For Each sStyle in aStyles
sStyle = Trim(sStyle)
a = Split(sStyle, ",")
s = a(0)
i = CInt(a(1))
Print "Adding heading style " & sStyle
dHeadingStyles.Add s, i
Next
Case "HidePageNumbersInWeb"
Print "Setting " & sKey & " = " & sValue
bHidePageNumbersInWeb = CBool(sValue)
Case "IncludePageNumbers"
Print "Setting " & sKey & " = " & sValue
bIncludePageNumbers = CBool(sValue)
Case "LowerHeadingLevel"
Print "Setting " & sKey & " = " & sValue
iLowerHeadingLevel = CInt(sValue)
Case "RightAlignPageNumbers"
Print "Setting " & sKey & " = " & sValue
bRightAlignPageNumbers = CBool(sValue)
Case "UpperHeadingLevel"
Print "Setting " & sKey & " = " & sValue
iUpperHeadingLevel = CInt(sValue)
Case "UseFields"
Print "Setting " & sKey & " = " & sValue
bUseFields = CBool(sValue)
Case "UseHeadingStyles"
Print "Setting " & sKey & " = " & sValue
bUseHeadingStyles = CBool(sValue)
Case "UseHyperlinks"
Print "Setting " & sKey & " = " & sValue
bUseHyperlinks = CBool(sValue)
Case "UseOutlineLevels"
Print "Setting " & sKey & " = " & sValue
bUseOutlineLevels = CBool(sValue)
End Select ' sKey
Next ' sKey

' If oTocs.Count > 0 Then oTocs(1).Delete
For Each oToc In oTocs
If oToc.Range.Text = oRange.Text Then oToc.Delete
Next

Set oToc = oTocs.Add(oRange, bUseHeadingStyles, iUpperHeadingLevel, iLowerHeadingLevel, bUseFields, iTableID, bRightAlignPageNumbers, bIncludePageNumbers, oAddedStyles, bUseHyperlinks, bHidePageNumbersInWeb, bUseOutlineLevels)
FixToc oDoc, oToc, dHeadingStyles, bUseHyperLinks, bUseOutlineLevels
End If ' oFind.Execute

Case "Global"
' Do Nothing

Case Else
' If dSection.Count > 0 Then Print "Replacing"
print sSection
Set oRange = oDoc.Content
Set oFind = oRange.Find
sFindText = ""
oFind.ClearFormatting
Set oReplace = oFind.Replacement
sReplaceText = ""
oReplace.ClearFormatting
iWrap = wdFindContinue
iReplace = wdReplaceOne
For Each sKey in dSection.Keys
sValue = dSection(sKey)
Select Case sKey
Case "FindText"
' text to find
Print "Setting " & sKey & " = " & sValue
sFindText = sValue
Case "MatchCase"
' whether to match case
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchCase = CBool(sValue)
Case "MatchWholeWord"
' whether to match only entire words
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchWholeWord = CBool(sValue)
Case "MatchWildcards"
' whether the find text can include wildcards
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchWildcards = CBool(sValue)
Case "MatchSoundsLike"
' whether to match words that sound similar
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchSoundsLike = CBool(sValue)
Case "MatchAllWordForms"
' True to match all word forms
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchAllWordForms = CBool(sValue)
Case "Forward"
' whether to search forward
Print "Setting " & sKey & " = " & CBool(sValue)
bForward = CBool(sValue)
Case "FindStyle"
' style to find
Print "Setting " & sKey & " = " & sValue
sFindStyle = sValue
oFind.Style = sFindStyle
Case "ReplaceStyle"
' style for replacement
Print "Setting " & sKey & " = " & sValue
sReplaceStyle = sValue
oReplace.Style = sReplaceStyle
Case "Wrap"
' how to act if the search did not begin at the start of the range
Print "Setting " & sKey & " = " & CInt(sValue)
iWrap = CInt(sValue)
Case "Format"
' whether to match formatting
Print "Setting " & sKey & " = " & CBool(sValue)
bFormat = CBool(sValue)
Case "MatchKashida"
' whether to match kashidas in an Arabic language
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchKashida = CBool(sValue)
Case "MatchDiacritics"
' whether to match diacritics in a right-to-left language
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchDiacritics = CBool(sValue)
Case "MatchAlefHamza"
' whether to match Alef Hamzas in an Arabic language
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchAlefHamza = CBool(sValue)
Case "MatchControl"
' whether to match bidirectional control characters in a right-to-left language
Print "Setting " & sKey & " = " & CBool(sValue)
bMatchControl = CBool(sValue)
Case "ReplaceText"
' text to replace with
Print "Setting " & sKey & " = " & sValue
sReplaceText = sValue
Case "Replace"
' how many replacements to make
Print "Setting " & sKey & " = " & CInt(sValue)
iReplace = CInt(sValue)
Case Else
sFindStyle = oFind.Style
Set oFindFormat = oFind.ParagraphFormat
Set oFindFont = oFind.Font

sReplaceStyle = oReplace.Style
Set oReplaceFormat = oReplace.ParagraphFormat
Set oReplaceFont = oReplace.Font
End Select ' sKey
Next ' sKey

iReplaceCount = 0
' Do While oFind.Execute(sFindText, bMatchCase, bMatchWholeWord, bMatchWildcards, bMatchSoundsLike, bMatchAllWordForms, bForward, iWrap, bFormat, sReplaceText, iReplace, bMatchKashida, bMatchDiacritics, bMatchAlefHamza, bMatchControl)
bLoop = True
' Do While oFind.Execute(sFindText, bMatchCase, bMatchWholeWord, bMatchWildcards, bMatchSoundsLike, bMatchAllWordForms, bForward, iWrap, bFormat, sReplaceText, iReplace, bMatchKashida, bMatchDiacritics, bMatchAlefHamza, bMatchControl)
Do While bLoop
If oFind.Execute(sFindText, bMatchCase, bMatchWholeWord, bMatchWildcards, bMatchSoundsLike, bMatchAllWordForms, bForward, iWrap, bFormat, sReplaceText, iReplace, bMatchKashida, bMatchDiacritics, bMatchAlefHamza, bMatchControl) Then
iReplaceCount = iReplaceCount + 1
Set oRange = oDoc.Content
Set oFind = oRange.Find
Set oReplace = oFind.Replacement
If sFindText = sReplaceText Then bLoop = False
Else
bLoop = False
End If
Loop
' Print StringPlural("replacement", iReplaceCount)
If iReplaceCount = 1 Then Print iReplaceCount & "match"
If iReplaceCount <> 1 Then Print iReplaceCount & "matches"
End Select ' sSection
Next ' sSection

If Not oDoc.Saved Then
If bBackupDocx Then
sBackupDocx = FileBackup(sSourceDocx)
If Len(sBackupDocx) = 0 Then
Print "Error creating backup "
Else
Print "Creating backup " & PathGetName(sBackupDocx)
End If
End If

' If not bBackupDocx or Len(sBackupDocx) > 0 Then
If not oDoc.Saved and (not bBackupDocx or Len(sBackupDocx) > 0) Then
print "Saving " & PathGetName(sSourceDocx)
oDoc.Save
End If
End If
End If ' Not bReadOnly

oApp.NormalTemplate.Saved = True
oDoc.Close(wdDoNotSaveChanges)
oApp.Quit

If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
End If

echo "Done"
