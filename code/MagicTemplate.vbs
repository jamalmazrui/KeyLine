Const wdStyleTypeParagraph = 1
Const wdStyleTypeCharacter = 2

Function EnsureStyle(sStyle, iType)
Dim oStyle

On Error Resume Next
Set oStyle = oDoc.Styles(sStyle)
On Error GoTo 0
If Err.Number > 0 Then Set oStyle = oDoc.Styles.Add(sStyle, iType)
Set EnsureStyle = oStyle
End Function

' Main

sSourceDocx = "C:\AccAuthor\MagicTemplate.docx"
' sTargetDocx = "new_MagicTemplate.docx"
sTargetDocx = "MagicTemplate.docx"
Set oApp = CreateObject("Word.Application")
oApp.RestrictLinkedStyles = False
Set oDoc = oApp.Documents.Open(sSourceDocx)

sStyle = "Normal"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = ""
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "First Indent Justify"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

' Same as Normal
sStyle = "Body Text"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = ""
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "First Indent Justify Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("First Indent Justify")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Normal Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("First Indent Justify Plus")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Undent Justify"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Undent Justify Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Undent Justify")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Right Align"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 2
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Right Align Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Right Align")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 2
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Undent Unjustify"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 0
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Undent Unjustify Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Undent Unjustify")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 0
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "First Indent Unjustify"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 0
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "First Indent Unjustify Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("First Indent Unjustify")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 0
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Center"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 1
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 10
oFormat.LineSpacingRule = 0
oFont.Size = 10
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 10
oFormat.PageBreakBefore = False

sStyle = "Center Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Center")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 1
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Block Indent Justify"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 14.4
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Block Indent Justify Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Block Indent Justify")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 14.4
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Quote"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = True
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 14.4
oFormat.RightIndent = 14.4
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Quote Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Quote")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = True
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 14.4
oFormat.RightIndent = 14.4
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "First Paragraph"
WScript.Echo (sStyle)
Set oStyle = oDoc.Styles(sStyle)
oStyle.Delete
' Set oStyle = EnsureStyle(sStyle, wdStyleTypeParagraph)
Set oStyle = oDoc.Styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Heading 1"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
' oStyle.Delete
Set oStyle = oDoc.Styles(sStyle)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.NextParagraphStyle = "First Paragraph"
' oStyle.NextParagraphStyle = "Normal"
oFormat.alignment = 1
oFont.Bold = True
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFont.Size = 20
oFormat.SpaceBefore = 24
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = True

sStyle = "Heading 2"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oStyle.NextParagraphStyle = "First Paragraph"
' oStyle.NextParagraphStyle = "Normal"
oFormat.alignment = 1
oFont.Bold = True
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFont.Size = 16
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12

sStyle = "Author"
WScript.Echo (sStyle)
Set oStyle = EnsureStyle(sStyle, wdStyleTypeParagraph)
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Center Plus")
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Center Plus")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 1
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Title"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oStyle.NextParagraphStyle = oDoc.Styles("Subtitle")
oFormat.alignment = 1
oFont.Bold = True
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFont.Size = 24
oFormat.SpaceBefore = 24
oFormat.SpaceAfter = 12

sStyle = "Subtitle"
WScript.Echo (sStyle)
Set oStyle = oDoc.Styles(sStyle)
oStyle.Delete
' Set oStyle = oDoc.Styles.Add(sStyle, wdStyleTypeParagraph)
Set oStyle = oDoc.Styles(sStyle)
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
' oStyle.BaseStyle = ""
' oStyle.BaseStyle = "Normal"
oStyle.BaseStyle = "Heading 1"
' oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Author")
oFormat.alignment = 1
oFont.Name = "Calibri"
oFont.Bold = True
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
WScript.Echo oFont.Size
oFont.Size = 16
WScript.Echo oFont.Size
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12

sStyle = "List Bullet"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFont.Color = -16777216
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0

sStyle = "List Bullet Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("List Bullet")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFont.Color = -16777216
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12

sStyle = "List Number"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFont.Color = -16777216
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0

sStyle = "List Number Plus"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("List Number")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFont.Color = -16777216
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12

sStyle = "Compact"
WScript.Echo (sStyle)
Set oStyle = EnsureStyle(sStyle, wdStyleTypeParagraph)
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 0
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 10
oFormat.LineSpacingRule = 0
oFont.Size = 10
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 10
oFormat.PageBreakBefore = False

sStyle = "Source Code"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
' oStyle.BaseStyle = oDoc.Styles("Plain Text")
' oStyle.BaseStyle = oDoc.styles.latent_styles["Plain Text")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 0
oFont.name = "Consolas"
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFormat.LeftIndent = 14.4
oFormat.RightIndent = 0
oFormat.LineSpacing = 10
oFormat.LineSpacingRule = 0
oFont.Size = 10
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 10
oFormat.PageBreakBefore = False

sStyle = "Separator"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oFont.name = "Segoe UI Symbol"
oFormat.alignment = 1
oFont.Color = -16777216
oFormat.FirstLineIndent = 0.72
oFont.Size = 18
oFormat.SpaceBefore = 18
oFormat.SpaceAfter = 18

sStyle = "Poem"
WScript.Echo (sStyle)
Set oStyle = Nothing
Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = -14.4
oFormat.LeftIndent = 14.4
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 12
oFormat.PageBreakBefore = False

sStyle = "Endnote Reference"
WScript.Echo (sStyle)
' Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeCharacter)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFont.Size = 12
oFont.superscript = True

sStyle = "Endnote Text"
WScript.Echo (sStyle)
' Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Footnote Reference"
WScript.Echo (sStyle)
' Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeCharacter)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFont.Size = 12
oFont.superscript = True

sStyle = "Footnote Text"
WScript.Echo (sStyle)
' Set oStyle = oDoc.styles.Add(sStyle, wdStyleTypeParagraph)
Set oStyle = Nothing
Set oStyle = oDoc.Styles(sStyle)
Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oStyle.BaseStyle = oDoc.Styles("Normal")
oStyle.NextParagraphStyle = oDoc.Styles("Normal")
oFormat.alignment = 3
oFont.Bold = False
oFont.Italic = False
oFont.Color = -16777216
oFormat.FirstLineIndent = 14.4
oFormat.LeftIndent = 0
oFormat.RightIndent = 0
oFormat.LineSpacing = 12
oFormat.LineSpacingRule = 0
oFont.Size = 12
oFormat.SpaceBefore = 0
oFormat.SpaceAfter = 0
oFormat.PageBreakBefore = False

sStyle = "Verbatim Char"
WScript.Echo (sStyle)
Set oStyle = EnsureStyle(sStyle, wdStyleTypeCharacter)
Set oFormat = Nothing
If oStyle.Type = wdStyleTypeParagraph Then Set oFormat = oStyle.ParagraphFormat
Set oFont = oStyle.Font
oFont.name = "Consolas"
oFont.Bold = True
oFont.Italic = False
oFont.Color = -16777216
oFont.Size = 12

' oDoc.SaveAs(sTargetDocx)
oDoc.Save
' oDoc.Close 0
oDoc.Close
oApp.Quit

WScript.Echo ("Done")
