Option Explicit
WScript.Echo"Starting DocxStyles"

Dim aStyles, aIni
Dim bResetDefault, bBackupDocx, bLogActions, bDemoStyles, bDescriptionOnly, bFound, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dStyleAttribs, d, dStyles, dStyle, dIni, dSourceIni
Dim iType, iParagraph, iStyle, iArgCount, iCount
Dim oNewStyle, oSepStyle, oStyleDoc, oParagraph, oSystem, oFile, oFormat, oFont, oStyle, oApp, oData, oDoc, oDocs, oFind, oProperty, oRange, oToc, oStyles
Dim sStyleAttrib, sType, sTargetTxt, sTargetLog, sTargetDocx, sBackupDocx, sScriptVbs, sHomerLibVbs, sDir, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

Const WdCollapseEnd = 0
Const WdDoNotSaveChanges = 0

' wdOrganizerObject enumeration
Const wdOrganizerObjectStyles = 0
Const wdOrganizerObjectAutoText = 1
Const wdOrganizerObjectCommandBars = 2
Const wdOrganizerObjectProjectItems = 3

' wd line break enumeration
Const wdColumnBreak = 8 ' Column break at the insertion point.
Const wdLineBreak = 6 ' Line break.
Const wdLineBreakClearLeft = 9 ' Line break.
Const wdLineBreakClearRight = 10 ' Line break.
Const wdPageBreak = 7 ' Page break at the insertion point.
Const wdSectionBreakContinuous = 3 ' New section without a corresponding page break.
Const wdSectionBreakEvenPage = 4 ' Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
Const wdSectionBreakNextPage = 2 ' Section break on next page.
Const wdSectionBreakOddPage = 5 ' Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
Const wdTextWrappingBreak = 11 ' Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.

' wdLineSpace rule enumeration
Const wdLineSpace1pt5 = 1 ' Space-and-a-half line spacing. Spacing is equivalent to the current font size plus 6 points.
Const wdLineSpaceAtLeast  = 3 ' Line spacing is always at least a specified amount. The amount is specified separately.
Const wdLineSpaceDouble = 2 ' Double spaced.
Const wdLineSpaceExactly = 4 ' Line spacing is only the exact maximum amount of space required. This setting commonly uses less space than single spacing.
Const wdLineSpaceMultiple = 5 ' Line spacing determined by the number of lines indicated.
Const wdLineSpaceSingle = 0 ' Single spaced. default

' wdAlignParagraph enumeration
Const wdAlignParagraphCenter=1
Const wdAlignParagraphDistribute=4
Const wdAlignParagraphJustify=3
Const wdAlignParagraphJustifyHi=7
Const wdAlignParagraphJustifyLow=8
Const wdAlignParagraphJustifyMed=5
Const wdAlignParagraphLeft=0
Const wdAlignParagraphRight=2
Const wdAlignParagraphThaiJustify=9

' wdStyleType enumeration
Const wdStyleTypeParagraph = 1
Const wdStyleTypeCharacter = 2
Const wdStyleTypeTable = 3
Const wdStyleTypeList = 4

Function FileInclude(sFile)
' With CreateObject("Scripting.FileSystemObject")
' ExecuteGlobal .openTextFile(sFile).readAll()
' End With

executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

Function CopyStyles(oStylesLocal, sSourceFile, sTargetFile)
Dim oStyleLocal

' print "Source " & PathGetName(sSourceFile)
' print "Target " & PathGetName(sTargetFile)
print "Copying all used styles from " & PathGetName(sSourceFile) & " to " & PathGetName(sTargetFile)
For Each oStyleLocal in oStylesLocal
' if oStyle.InUse And oStyle.Type = wdStyleTypeParagraph Then
if oStyleLocal.InUse Then
' print oStyleLocal.NameLocal
sSourceFile = oDoc.FullName
oApp.OrganizerCopy sSourceFile, sTargetFile, oStyleLocal.NameLocal, wdOrganizerObjectStyles 
End If
Next ' For Each oStyle in oDoc.Styles
End Function

Function InsertWithLineBreaks(oApp, sText)
Dim aLines
Dim iLine, iBound
Dim s, sLine

s = StringConvertToUnixLineBreak(sText)
aLines = Split(s, vbLf)
iLine = 0
iBound = ArrayBound(aLines)
' print "s = " & s
' print "iBound = " & iBound
For iLine = 0 to iBound
sLine = Trim(aLines(iLine))
' print "sLine = " & sLine
If Len(sLine) > 0 Then oApp.Selection.InsertAfter sLine
oApp.Selection.Collapse WdCollapseEnd
If iLine < iBound Then oApp.Selection.InsertBreak wdLineBreak
Next
End Function

Function StyleExists(sStyle)
On Error Resume Next
StyleExists=True
Set oStyle =ActiveDocument.Styles(sStyle)
If Err.Number<>0 then StyleExists=false
Err.Clear
End Function

Function SetStyleAttribute(oStyleLocal, oFormatLocal, oFontLocal, sAttribLocal, sValueLocal)
Dim sStyleLocal

sStyleLocal = oStyleLocal.NameLocal

' Now set in configuration file
' If oStyleLocal.Type = wdStyleTypeParagraph Then oStyleLocal.AutomaticallyUpdate = True  
Select Case sAttribLocal
Case "Type"
' Print "Cannot set Type property of " & sStyleLocal
Case "AutomaticallyUpdate"
If oStyleLocal = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & CBool(sValueLocal)
oStyleLocal.AutomaticallyUpdate = CBool(sValueLocal)
End If
Case "BaseStyle"
' Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
Print "Setting " & sAttribLocal & " = " & sValueLocal
' Try assigning object not string
If Len(sValueLocal) = 0 Then
oStyleLocal.BaseStyle = sValueLocal
Else
print "sValueLocal=" & sValueLocal
oStyleLocal.BaseStyle = oDoc.Styles(sValueLocal)
End If
Case "NameLocal"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oStyleLocal.NameLocal = sValueLocal
Case "BuiltIn"
' Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
Print "Cannot set BuiltIn property of " & sStyleLocal
' oStyleLocal.BuiltIn = CBool(sValueLocal)
Case "InUse"
' Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
Print "Cannot set InUse property of " & sStyleLocal
'oStyleLocal.InUse = CBool(sValueLocal)
Case "Linked"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Cannot set Linked property of " & sStyleLocal
' Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & CBool(sValueLocal)
' oStyleLocal.Linked = CBool(sValueLocal)
End If
Case "ListLevelNumber"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oStyleLocal.ListLevelNumber = CInt(sValueLocal)
Case "Locked"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & CBool(sValueLocal)
oStyleLocal.Locked = CBool(sValueLocal)
Case "NextParagraphStyle"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oStyleLocal.NextParagraphStyle = sValueLocal
Case "NoProofing"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & CBool(sValueLocal)
oStyleLocal.NoProofing = CBool(sValueLocal)
Case "NoSpaceBetweenParagraphsOfSameStyle"
If oStyleLocal = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oStyleLocal.NoSpaceBetweenParagraphsOfSameStyle = CBool(sValueLocal)
End If
Case "QuickStyle"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oStyleLocal.QuickStyle = CBool(sValueLocal)
Case "UnhideWhenUsed"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & CBool(sValueLocal)
oStyleLocal.UnhideWhenUsed = CBool(sValueLocal)
Case "Alignment"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.Alignment = CInt(sValueLocal)
End If
Case "FirstLineIndent"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.FirstLineIndent = CDbl(sValueLocal)
End If
Case "KeepTogether"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.KeepTogether = CBool(sValueLocal)
End If
Case "KeepWithNext"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.KeepWithNext = CBool(sValueLocal)
End If
Case "LeftIndent"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.LeftIndent = CDbl(sValueLocal)
End If
Case "LineSpacing"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.LineSpacing = CDbl(sValueLocal)
End If
Case "LineSpacingRule"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.LineSpacingRule = CInt(sValueLocal)
End If
Case "OutlineLevel"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.OutlineLevel = CInt(sValueLocal)
End If
Case "PageBreakBefore"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.PageBreakBefore = CBool(sValueLocal)
End If
Case "RightIndent"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.RightIndent = CDbl(sValueLocal)
End If
Case "SpaceAfter"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.SpaceAfter = CInt(sValueLocal)
End If
Case "SpaceBefore"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.SpaceBefore = CInt(sValueLocal)
End If
Case "WidowControl"
If oStyleLocal.Type = wdStyleTypeParagraph Then
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFormatLocal.WidowControl = CBool(sValueLocal)
End If
Case "Font"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Name = sValueLocal
Case "Bold"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Bold = CBool(sValueLocal)
Case "Color"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Color = CLng(sValueLocal)
Case "Hidden"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Hidden = CInt(sValueLocal)
Case "Italic"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Italic = CBool(sValueLocal)
Case "Size"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Size = CInt(sValueLocal)
Case "StrikeThrough"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.StrikeThrough = CBool(sValueLocal)
Case "Subscript"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Subscript = CBool(sValueLocal)
Case "Superscript"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Superscript = CBool(sValueLocal)
Case "Underline"
Print "Setting style " & sStyleLocal & ", " & sAttribLocal & " = " & sValueLocal
oFontLocal.Underline = CBool(sValueLocal)
End Select
End Function

Function GetStyleAttribs()
Dim dAttribs

Set dAttribs = CreateDictionary()
' dAttribs.Add "Type", ""
dAttribs.Add "BaseStyle", ""
dAttribs.Add "NameLocal", ""
' dAttribs.Add "BuiltIn", ""
' dAttribs.Add "InUse", ""
' dAttribs.Add "ListLevelNumber", ""
dAttribs.Add "NextParagraphStyle", "Normal"
dAttribs.Add "NoSpaceBetweenParagraphsOfSameStyle", False
' dAttribs.Add "QuickStyle", ""
dAttribs.Add "Alignment", 0
dAttribs.Add "FirstLineIndent", 0
dAttribs.Add "KeepTogether", False
dAttribs.Add "KeepWithNext", False
dAttribs.Add "LeftIndent", 0
' dAttribs.Add "LineSpacing", oDoc.Styles("Normal").ParagraphFormat.LineSpacing
dAttribs.Add "LineSpacing", 12
dAttribs.Add "LineSpacingRule", 0
dAttribs.Add "OutlineLevel", 10
dAttribs.Add "PageBreakBefore", False
dAttribs.Add "RightIndent", 0
dAttribs.Add "SpaceAfter", 0
dAttribs.Add "SpaceBefore", 0
dAttribs.Add "WidowControl", False
' dAttribs.Add "Font", "Default Paragraph Font"
' dAttribs.Add "Font", oDoc.Styles("Normal").Font.Name
dAttribs.Add "Font", "Cambria"
dAttribs.Add "Bold", False
dAttribs.Add "Color", -16777216
dAttribs.Add "Hidden", False
dAttribs.Add "Italic", False
' dAttribs.Add "Size", oDoc.Styles("Normal").Font.Size
' dAttribs.Add "Size", oDoc.Styles("Default Paragraph Font").Font.Size
' dAttribs.Add "Size", oDoc.Styles("Normal").Font.Size
dAttribs.Add "Size", 12
dAttribs.Add "StrikeThrough", False
dAttribs.Add "Subscript", False
dAttribs.Add "Superscript", False
dAttribs.Add "Underline", False
Set GetStyleAttribs = dAttribs
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
Print "Optionally specify a configuration .ini file as a second parameter."
Quit
End If

sSourceDocx = WScript.Arguments(0)
If InStr(sSourceDocx, "\") = 0 Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If not FileExists(sSourceDocx) Then Quit "Cannot find " & sSourceDocx

If iArgCount > 1 Then
sSourceIni = WScript.Arguments(1)
sSourceIni = GetIniFile(sSourceIni)
If not FileExists(sSourceIni) Then Quit "Cannot find " & sSourceIni
Set dSourceIni = IniToDictionary(sSourceIni)
bBackupDocx = GetGlobalValue(dSourceIni, "BackupDocx", True)
bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)
bDescriptionOnly = GetGlobalValue(dSourceIni, "DescriptionOnly", False)
bDemoStyles = GetGlobalValue(dSourceIni, "DemoStyles", False)

bReadOnly = False
ProcessTerminateAllModule "WinWord"
Else
sSourceIni = ""
Set dSourceIni = CreateDictionary()
bReadOnly = True
End If

sTargetIni = PathCombine(PathGetCurrentDirectory(), "STYLES-" & PathGetRoot(sSourceDocx) & ".ini")
sTargetLog = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceDocx) & "-DocxStyles.log")
' sTargetTxt = PathChangeExtension(sTargetLog, "txt")
sTargetTxt = PathChangeExtension(sTargetIni, "txt")
' sTargetDocx = PathChangeExtension(sTargetLog, "docx")
sTargetDocx = PathChangeExtension(sTargetIni, "docx")

Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = 0
oApp.ScreenUpdating = False

Set oDocs = oApp.Documents
bAddToRecentFiles = False
bConfirmConversions = False
Print "Opening " & PathGetName(sSourceDocx)
Set oDoc = oDocs.Open(sSourceDocx, bAddToRecentFiles, bReadOnly, bConfirmConversions)

If Not bReadOnly Then
Print "Applying " & PathGetName(sSourceIni)
' Apply configuration changes
For Each sStyle in dSourceIni.Keys()
bFound = False
Set oStyles = oDoc.Styles
For Each oStyle in oStyles
' For iStyle = (oStyles.Count) To 1 Step -1
' Set oStyle = oStyles(iStyle)
' Try character styles as well
' If oStyle.Type = wdStyleTypeParagraph and oStyle.NameLocal = sStyle Then
If (oStyle.Type = wdStyleTypeParagraph or oStyle.Type = wdStyleTypeCharacter) and oStyle.NameLocal = sStyle Then
' print "NameLocal " & oStyle.NameLocal
bFound = True
Set dStyle = dSourceIni(sStyle)

If dStyle.Count = 0 Then
Print "Deleting style " & sStyle
oStyle.Delete

Else
' Edit style
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
' If dStyle.Count > 1 Then print "Style " & sStyle
bResetDefault = False
On Error Resume Next
bResetDefault = CBool(dStyle("ResetDefault"))
On Error GoTo 0

bResetDefault = True
If bResetDefault Then
print "Resetting default"
' Set oStyle.ParagraphFormat = oDoc.Styles("Normal").ParagraphFormat
' Set oStyle.ParagraphFormat = null
' Set oStyle.Font = oDoc.Styles("Normal").Font
' Set oStyle.Font = Null
Set dStyleAttribs = GetStyleAttribs()
For Each sStyleAttrib in dStyleAttribs
' SetStyleAttribute oStyle, oStyle.ParagraphFormat, oStyle.Font, s, Null
' SetStyleAttribute oStyle, oStyle.ParagraphFormat, oStyle.Font, s, vbNullString
If oStyle.Type = wdStyleTypeParagraph Then SetStyleAttribute oStyle, oStyle.ParagraphFormat, oStyle.Font, sStyleAttrib, dStyleAttribs(sStyleAttrib)
If oStyle.Type = wdStyleTypeCharacter Then SetStyleAttribute oStyle, Nothing, oStyle.Font, sStyleAttrib, dStyleAttribs(sStyleAttrib)
Next
End If ' if bResetDefault

For Each sAttrib in dStyle.Keys()
sValue = dStyle(sAttrib)
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
SetStyleAttribute oStyle, oFormat, oFont, sAttrib, sValue
Next 'For Each sAttrib in dStyle.Keys()
End If ' If dStyle.Count = 0
Exit For
End If ' If oStyle.Type = wdStyleTypeParagraph and oStyle.NameLocal = sStyle  
Next ' For Each oStyle in oStyles

' Create style
If not bFound and sStyle <> "Global" Then
Print "Creating style " & sStyle
iType = wdStyleTypeParagraph
On Error Resume Next
iType = CInt(dSourceIni(sStyle)("Type"))
On Error GoTo 0
If iType = 0 or Err.Number Then iType = wdStyleTypeParagraph

' Set oStyle = oStyles.Add(sStyle, WdStyleTypeParagraph)
Set oNewStyle = oStyles.Add(sStyle, iType)
' Set oFormat = oStyle.ParagraphFormat
Set oFormat = Nothing
If oNewStyle.Type = wdStyleTypeParagraph Then Set oFormat = oNewStyle.ParagraphFormat
Set oFont = oNewStyle.Font
Set dStyle = dSourceIni(sStyle)
For Each sAttrib in dStyle.Keys()
sValue = dStyle(sAttrib)
SetStyleAttribute oNewStyle, oFormat, oFont, sAttrib, sValue
Next ' For Each sAttrib in dStyle.Keys()
End If ' If not bFound 

Next ' For Each sStyle in dSourceIni.Keys()

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

If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed
End If

' Create target ini
Print "Collecting style Information"
iCount = 0
Set oStyles = oDoc.Styles
For Each oStyle in oDoc.Styles
' Try reporting character styles
' If oStyle.Type = wdStyleTypeParagraph Then
If oStyle.Type = wdStyleTypeParagraph or oStyle.Type = wdStyleTypeCharacter Then
iCount = iCount + 1
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
sStyle = oStyle.NameLocal

If iCount > 1 Then AppendLine ""
AppendLine "[" & sStyle & "]"
If not bDescriptionOnly Then
AppendLine "; General"

' Whether a paragraph, character, or other style
AppendLine "Type = " & oStyle.Type

' True if changes to a style automatically update descendent styles
If oStyle.Type = wdStyleTypeParagraph Then AppendLine "AutomaticallyUpdate = " & CBool(oStyle.AutomaticallyUpdate)

' Object for an existing style inherited by this one.
AppendLine "BaseStyle = " & oStyle.BaseStyle.NameLocal

' True if a built-in style of Word.
AppendLine "BuiltIn = " & CBool(oStyle.BuiltIn)
' True if the style has been used in the document
AppendLine "InUse = " & CBool(oStyle.InUse)
' True if the style may be linked with a character style
If oStyle.Type = wdStyleTypeParagraph Then AppendLine "Linked = " & CBool(oStyle.Linked)
' Integer for ListLevelNumber if in amultiple level list
' Not generally useful property
' AppendLine "ListLevelNumber = " & CInt(oStyle.ListLevelNumber)
' True if the style is locked against editing its properties
AppendLine "Locked = " & CBool(oStyle.Locked)
' Object for automatic style of the next paragraph.
AppendLine "NextParagraphStyle = " & oStyle.NextParagraphStyle.NameLocal
' True if text formatted with the style is not proofed by spelling or grammer checks
AppendLine "NoProofing = " & CBool(oStyle.NoProofing)
' True if Word removes spacing between paragraphs with this style.
If oStyle.Type = wdStyleTypeParagraph Then AppendLine "NoSpaceBetweenParagraphsOfSameStyle = " & CBool(oStyle.NoSpaceBetweenParagraphsOfSameStyle)
' True if a quick style in the Word UI.
AppendLine "QuickStyle = " & CBool(oStyle.QuickStyle)
' True if the style is hidden from the UI until applied
AppendLine "UnhideWhenUsed = " & CBool(oStyle.UnhideWhenUsed)

If oStyle.Type = wdStyleTypeParagraph Then
' Paragraph format
AppendLine ""
AppendLine "; Paragraph"
' Integer for alignment.  WdParagraphAlignment.
AppendLine "Alignment = " & CInt(oFormat.Alignment)

' Integer in points for a first line or hanging indent.
AppendLine "FirstLineIndent = " & CDbl(oFormat.FirstLineIndent)

' True if all lines in the paragraph remain on the same page after repagination.
AppendLine "KeepTogether = " & CBool(oFormat.KeepTogether)

' True if the paragraph remains on the same page as the next one after repagination.
AppendLine "KeepWithNext = " & CBool(oFormat.KeepWithNext)

' Integer in points for the left indent.
AppendLine "LeftIndent = " & CDbl(oFormat.LeftIndent)

' Float in points for the line spacing.
AppendLine "LineSpacing = " & CDbl(oFormat.LineSpacing)
' Integer enumeration for the line spacing rule.
AppendLine "LineSpacingRule = " & CInt(oFormat.LineSpacingRule)
' Integer for the outline level
AppendLine "OutlineLevel = " & CInt(oFormat.OutlineLevel)

' True if a page break is forced before the paragraph.
AppendLine "PageBreakBefore = " & CBool(oFormat.PageBreakBefore)

' Integer in points for the right indent.
AppendLine "RightIndent = " & CDbl(oFormat.RightIndent)

' Integer in points for spacing after the paragraph.
AppendLine "SpaceAfter = " & CInt(oFormat.SpaceAfter)

' Integer in points for spacing before the paragraph.
AppendLine "SpaceBefore = " & CInt(oFormat.SpaceBefore)

' True if the first and last lines in the paragraph remain on the same page as the rest of the paragraph after repagination.
AppendLine "WidowControl = " & CBool(oFormat.WidowControl)
End If ' if wdStyleTypeParagraph

' Character formatting
AppendLine ""
AppendLine "; Font"

' String for the font name.
AppendLine "Font = " & oFont.Name

' oFont.Bold =
' True if bold.
AppendLine "Bold = " & CBool(oFont.Bold)

' Integer for the 24-bit color.   WdColor.
AppendLine "Color = " & CLng(oFont.Color)
' True if Hidden.
AppendLine "Hidden = " & cInt(oFont.Hidden)
' True if italic.
AppendLine "Italic = " & CBool(oFont.Italic)

' Integer in points for font size.
AppendLine "Size = " & CInt(oFont.Size)

' True if strike-through.
AppendLine "StrikeThrough = " & CBool(oFont.StrikeThrough)

' True if subscript.
AppendLine "Subscript = " & CBool(oFont.Subscript)

' True if superscript.
AppendLine "Superscript = " & CBool(oFont.Superscript)

' True if Underline.
AppendLine "Underline = " & CBool(oFont.Underline)

End If ' If bDescriptionOnly
' Read-only string for a description of the style.
AppendLine "Description = " & VbCrLf & oStyle.Description
End If
Next

Print "Saving " & PathGetName(sTargetIni)
StringToFile sHomerText, sTargetIni

Print "Creating " & PathGetName(sTargetTxt)
sHomerText = ""
Set dStyles = CreateDictionary
For iParagraph = 1 To oDoc.Paragraphs.Count
Set oParagraph = oDoc.Paragraphs(iParagraph)
If iParagraph <> 1 Then AppendBlank
' If iParagraph <> 1 Then AppendBlank
sStyle = oParagraph.Style.NameLocal
dStyles(sStyle) = ""
AppendLine "[" & sStyle & "]"
' AppendLine oParagraph.Range.Text
s = oParagraph.Range.Text
If StringTrail(s, vbCr, False) Then s = Mid(s, 1, Len(s) - 1)
AppendLine s
Next

aStyles = dStyles.Keys
ArraySort aStyles
' s = StringPlural("style", dStyles.Count) & " used:" & vbCrLf
s = StringPlural("style", dStyles.Count) & " used by " & StringPlural("paragraph", oDoc.Paragraphs.Count) & ":" & vbCrLf
For Each sStyle in aStyles
s = s & sStyle & vbCrLf
Next

sHomerText = s & VbCrLf & sHomerText
StringToFile sHomerText, sTargetTxt

If bDemoStyles Then
print "Creating " & PathGetName(sTargetDocx)
Set oStyleDoc = oDocs.Add
oStyleDoc.Activate

Set oSepStyle = GetObject(oDoc.Styles, "Separator")
If oSepStyle Is Nothing Then Set oSepStyle = oDoc.Styles("Intense Quote")
For Each oStyle in oDoc.Styles
' Try character styles as well
' If oStyle.Type = wdStyleTypeParagraph and oStyle.InUse Then
' Try all not just in use
' If (oStyle.Type = wdStyleTypeParagraph or oStyle.Type = wdStyleTypeCharacter) and oStyle.InUse Then
If (oStyle.Type = wdStyleTypeParagraph or oStyle.Type = wdStyleTypeCharacter) Then
oApp.Selection.Collapse WdCollapseEnd
' oApp.Selection.InsertAfter "This is style " & oStyle.NameLocal & "." & vbCr & oStyle.Description
' oApp.Selection.InsertAfter "This is style " & oStyle.NameLocal & "." & vbLf & StringConvertToUnixLineBreak(oStyle.Description)
sType = "paragraph"
If oStyle.Type = wdStyleTypeCharacter Then sType = "character"
s = "This is " & sType & " style " & oStyle.NameLocal & "." & vbLf & oStyle.Description
' InsertWithLineBreaks oApp, "This is style " & oStyle.NameLocal & "." & vbLf & oStyle.Description
InsertWithLineBreaks oApp, s
oApp.Selection.InsertParagraphAfter
' oApp.Selection.Style = oStyle
' print oStyle.NameLocal
' oApp.Selection.Paragraphs(1).Style = oStyle
oApp.Selection.Paragraphs(1).Range.Select
oApp.Selection.Style = oStyle

oApp.Selection.Collapse WdCollapseEnd
oApp.Selection.InsertAfter "* * *"
oApp.Selection.InsertParagraphAfter
' oApp.Selection.Style = "Separator"
' oApp.Selection.Style = oDoc.Styles("Separator")
oApp.Selection.Style = oSepStyle
End If
next
oStyleDoc.SaveAs(sTargetDocx)
oStyleDoc.Close wdDoNotSaveChanges
oDoc.Activate

CopyStyles oDoc.Styles, sSourceDocx, sTargetDocx
End If ' bDemoStyles

oApp.NormalTemplate.Saved = True
oApp.PrintPreview = False
oDoc.Close(wdDoNotSaveChanges)

oApp.Quit
Print "Done"
