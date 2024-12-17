import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.enum.dml import MSO_COLOR_TYPE  
from docx.enum.dml import MSO_THEME_COLOR

# dropped attribs: builtin, delete, element, font, part, paragraph_format, style_id, type
sStyleAttribs = """
base_style
hidden
locked
name
next_paragraph_style
priority
quick_style
unhide_when_used
"""
lParStyleAttribs = sStyleAttribs.strip().replace("\r\n", "\n").split("\n")
lCharStyleAttribs = list(lParStyleAttribs)
lCharStyleAttribs.remove("next_paragraph_style")

# dropped attribs: element, tab_stops
sFormatAttribs = """
alignment
first_line_indent
keep_together
keep_with_next
left_indent
line_spacing
line_spacing_rule
page_break_before
right_indent
space_after
space_before
widow_control
"""
lFormatAttribs = sFormatAttribs.strip().replace("\r\n", "\n").split("\n")

# dropped attribs: color, element, part
sFontAttribs = """
all_caps
bold
complex_script
cs_bold
cs_italic
double_strike
emboss
hidden
highlight_color
imprint
italic
math
name
no_proof
outline
rtl
shadow
size
small_caps
snap_to_grid
spec_vanish
strike
subscript
superscript
underline
web_hidden
"""
lFontAttribs = sFontAttribs.strip().replace("\r\n", "\n").split("\n")

def ensureStyle(sStyle, iType=1, bResetDefault=True):
	try: style = doc.styles[sStyle]
	except: style = doc.styles.add_style(sStyle, iType)
	if bResetDefault: style = resetDefault(style)
	return style

def resetDefault(style):
	sName = style.name
	iType = style.type
	if not iType in (1, 2): return style
	if iType == 1:
		for sStyleAttrib in lParStyleAttribs: setattr(style, sStyleAttrib, None)
		for sFormatAttrib in lFormatAttribs: setattr(style.paragraph_format, sFormatAttrib, None)
		style.paragraph_format.tab_stops.clear_all()
	elif iType == 2:
		for sStyleAttrib in lCharStyleAttribs: setattr(style, sStyleAttrib, None)

	for sFontAttrib in lFontAttribs: setattr(style.font, sFontAttrib, None)
	style.name = sName
	return style

# Main

sSourceDocx = "MagicTemplate.docx"
# sTargetDocx = "new_MagicTemplate.docx"
sTargetDocx = "MagicTemplate.docx"
doc = docx.Document(sSourceDocx)

sStyle = "Normal"
print(sStyle)
style = doc.styles[sStyle]
format = style.paragraph_format
font = style.font
# style.base_style = None
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "First Indent Justify"
print(sStyle)
# style = doc.styles.add_style(sStyle, WD_STYLE_TYPE.PARAGRAPH)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

# Same as Normal
sStyle = "Body Text"
print(sStyle)
style = doc.styles[sStyle]
format = style.paragraph_format
font = style.font
style.base_style = None
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "First Indent Justify Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["First Indent Justify"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Normal Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["First Indent Justify Plus"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Undent Justify"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Undent Justify Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Undent Justify"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Right Align"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 2
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Right Align Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Right Align"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 2
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Undent Unjustify"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = Pt(0)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Undent Unjustify Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Undent Unjustify"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = Pt(0)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "First Indent Unjustify"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = Pt(0)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "First Indent Unjustify Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["First Indent Unjustify"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = Pt(0)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Center"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 1
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(10)
format.line_spacing_rule = Pt(0)
font.size = Pt(10)
format.space_before = Pt(0)
format.space_after = Pt(10)
format.page_break_before = False

sStyle = "Center Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Center"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 1
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Block Indent Justify"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(14.4)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Block Indent Justify Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Block Indent Justify"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(14.4)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Quote"
print(sStyle)
style = doc.styles[sStyle]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = True
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(14.4)
format.right_indent = Pt(14.4)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Quote Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Quote"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = True
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(14.4)
format.right_indent = Pt(14.4)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Heading 1"
print(sStyle)
style = doc.styles[sStyle]
style.next_paragraph_style = doc.styles["First Paragraph"]
format.alignment = 1
font.bold = True
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
font.size = Pt(20)
format.space_before = Pt(24)
format.space_after = Pt(12)
format.page_break_before = True

sStyle = "Heading 2"
print(sStyle)
style = doc.styles[sStyle]
style.next_paragraph_style = doc.styles["First Paragraph"]
format.alignment = 1
font.bold = True
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
font.size = Pt(16)
format.space_before = Pt(0)
format.space_after = Pt(12)

sStyle = "Title"
print(sStyle)
style = doc.styles[sStyle]
style.base_style = None
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Subtitle"]
format.alignment = 1
font.bold = True
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
font.size = Pt(24)
format.space_before = Pt(24)
format.space_after = Pt(12)

sStyle = "Subtitle"
print(sStyle)
style = ensureStyle(sStyle)
style = resetDefault(style)
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Author"]
format.alignment = 1
font.bold = True
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
font.size = Pt(16)
format.space_before = Pt(0)
format.space_after = Pt(12)

sStyle = "List Bullet"
print(sStyle)
style = doc.styles[sStyle]
style.next_paragraph_style = doc.styles["Normal"]
font.color.rgb = None
font.color.theme_color = None
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)

sStyle = "List Bullet Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["List Bullet"]
style.next_paragraph_style = doc.styles["Normal"]
font.color.rgb = None
font.color.theme_color = None
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)

sStyle = "List Number"
print(sStyle)
style = doc.styles[sStyle]
style.next_paragraph_style = doc.styles["Normal"]
font.color.rgb = None
font.color.theme_color = None
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)

sStyle = "List Number Plus"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["List Number"]
style.next_paragraph_style = doc.styles["Normal"]
font.color.rgb = None
font.color.theme_color = None
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)

sStyle = "Compact"
print(sStyle)
style = doc.styles[sStyle]
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = Pt(0)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(10)
format.line_spacing_rule = Pt(0)
font.size = Pt(10)
format.space_before = Pt(0)
format.space_after = Pt(10)
format.page_break_before = False

sStyle = "Source Code"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
# style.base_style = doc.styles["Plain Text"]
# style.base_style = doc.styles.latent_styles["Plain Text"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = Pt(0)
font.name = "Consolas"
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(14.4)
format.right_indent = Pt(0)
format.line_spacing = Pt(10)
format.line_spacing_rule = Pt(0)
font.size = Pt(10)
format.space_before = Pt(0)
format.space_after = Pt(10)
format.page_break_before = False

sStyle = "First Paragraph"
print(sStyle)
style = doc.styles[sStyle]
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Undent Justify"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Separator"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
font.name = "Segoe UI Symbol"
format.alignment = 1
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
font.size = Pt(18)
format.space_before = Pt(18)
format.space_after = Pt(18)

sStyle = "Poem"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(-14.4)
format.left_indent = Pt(14.4)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

sStyle = "Author"
print(sStyle)
style = doc.styles[sStyle]
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Center Plus"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 1
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(0.72)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(12)
format.page_break_before = False

# style = doc.styles["Endnote Reference"]
sStyle = "Endnote Reference"
print(sStyle)
ensureStyle(sStyle, WD_STYLE_TYPE.CHARACTER)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
font.size = Pt(12)
font.superscript = True

sStyle = "Endnote Text"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

sStyle = "Footnote Reference"
print(sStyle)
ensureStyle(sStyle, WD_STYLE_TYPE.CHARACTER)
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
font.size = Pt(12)
font.superscript = True

sStyle = "Footnote Text"
print(sStyle)
style = ensureStyle(sStyle)
format = style.paragraph_format
font = style.font
style.base_style = doc.styles["Normal"]
style.next_paragraph_style = doc.styles["Normal"]
format.alignment = 3
font.bold = False
font.italic = False
font.color.rgb = None
font.color.theme_color = None
format.first_line_indent = Pt(14.4)
format.left_indent = Pt(0)
format.right_indent = Pt(0)
format.line_spacing = Pt(12)
format.line_spacing_rule = Pt(0)
font.size = Pt(12)
format.space_before = Pt(0)
format.space_after = Pt(0)
format.page_break_before = False

# style = doc.styles["Hyperlink"]
font.bold = True
font.italic = False
font.color.rgb = None
font.color.theme_color = None
font.size = Pt(10)

sStyle = "Verbatim Char"
print(sStyle)
style = doc.styles[sStyle]
font.name = "Consolas"
font.bold = True
font.italic = False
font.color.rgb = None
font.color.theme_color = None
font.size = Pt(12)

doc.save(sTargetDocx)

print("Done")
