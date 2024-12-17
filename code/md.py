sCellDivider = " | "

def pagebreak(): return line2(r"\pagebreak")

def field(sName, xValue):
	sReturn = bold(sName) + ": " + str(xValue)
	return line2(sReturn)

def titleBlock(sTitle, sAuthor=None, sDate=None):
	sReturn = dashLine()
	sReturn += line("title: '" + sTitle + "'")
	if sAuthor: sReturn += line("author: " + sAuthor)
	if sDate: sReturn += line("date: " + sDate)
	sReturn += dashLine()
	return line(sReturn)
	
def footnote(sReference): return "[^" + sReference + ']'

def bold(sText): return "**" + sText + "**"

def italics(sText): return '*' + sText + '*'

def link(sText=None, sTarget=None, sTitle=None):
	if sText == None: sText = ""
	if sTarget == None: sTarget = ""
	sReturn = '[' + sText + "](" + sTarget
	if sTitle: sReturn += ' "' + sTitle + '")'
	return sReturn

def internalLink(sText=None, sTarget=None):
	if sText == None: sText = ""
	if sTarget == None: sTarget = ""
	sReturn = '[' + sText + "](" + sTarget + ']'
	return sReturn

def image(sText=None, sTarget=None, sTitle=None):
	if sText == None: sText = ""
	if sTarget == None: sTarget = ""
	sReturn = "![" + sText + "](" + sTarget
	if sTitle: sReturn += ' "' + sTitle + '")'
	return sReturn

def definitionList(dTerms):
	l = [k + "\n:" + v for k, v in dTerms.items()]
	sReturn = "\n\n".join(l)
	return line2(sReturn)

def bulletList(lItems):
	l = ["- " + str(v) for v in lItems]
	sReturn = "\n".join(l)
	return line2(sReturn)

def numberedList(lItems):
	l = ["1. " + str(v) for v in lItems]
	sReturn = "\n".join(l)
	return line2(sReturn)

def blockQuote(sText):
	sReturn = "> " + sText
	return line2(sReturn)

def separator(): return line("***")

def strikeThrough(sText): return "~~" + sText + "~~"

def superscript(sText): return '^' + sText + '^'

def subscript(sText): return '~' + sText + '~'

def tableCaption(sCaption): return line2(": " + sCaption)

def table(sCaption=None, lHeaders=None, lAligns=None, lRows=None):
	sReturn = ""
	if sCaption: sReturn += tableCaption(sCaption)
	sReturn += line(tableHeader(lHeaders))
	sReturn += line(tableAlign(lAligns))
	for row in lRows: sReturn += line(tableRow(row))
	return sReturn

def tableAlign(lAligns):
	sReturn = ""
	l = []
	for s in lAligns:
		if s == 'r': l.append("--:")
		elif s == 'c': l.append("-:-")
		else: l.append(":--")
	sReturn = sCellDivider.join(l)
	return line(sReturn)

def tableRow(lCells):
	l = [escape(str(v)) for v in lCells]
	sReturn = sCellDivider.join(l)
	return line(sReturn)

def tableHeader(lHeaders):
	sReturn = sCellDivider.join(lHeaders)
	return line(sReturn)

def escape(sText):
	if len(sText) >1 and sText[0:1] == r'<' and sText[-1:] == r'>': return sText
	sReturn = sText
	sEscape = r"""\`~*_{}[]()<>+-.!@#$%^&="""
	# sReturn = sReturn.replace(r'|', r"\|")
	# sReturn = sReturn.replace(r'<', r"\<")
	#sReturn = sReturn.replace(r'>', r"\>")
	# for s in sEscape: sReturn = sReturn.replace(s, r'\' + s)
	for s in sEscape: sReturn = sReturn.replace(s, chr(92) + s)
	return sReturn

def lineFeed(iCount=1): return "\n" * iCount

def line(sText="", iLineFeed=1): return sText + lineFeed(iLineFeed)

def line2(sText): return line(sText, 2)

def forceLine(sText): return line(sText + " \\")

def dashLine(): return line("---")

def yamlBlock(sText): sReturn = dashLine(); sReturn += line(sText); sReturn += dashLine(); return sReturn

def urlLine(sTarget): return line2("<" + sTarget + ">")

def forceUrlLine(sTarget): return forceLine("<" + sTarget + ">")

def graveLine(): return line("```")

def fenceBlock(sText): sReturn = graveLine(); sReturn += line(sText); sReturn += graveLine(); return sReturn

def fenceSpan(sText): return "`" + sText + "`"

def fenceLine(sText): return line(fenceSpan(sText))

def forceFenceLine(sText): return forceLine(fenceSpan(sText))

def itemLine(sText): return line("- " + sText)

def forceItemLine(sText): return forceLine("- " + sText)
 

def heading(sText, iLevel=1): sNumSigns = "#" * iLevel; return line2(sNumSigns + " " + sText)

def stringPlural(sTerm, iCount): return str(iCount) + " " + (sTerm if iCount == 1 else sTerm + "s")

def fenceItemBlock(sBlock): sReturn = fenceIndentBlock(sBlock); sReturn = "- " + sReturn[2:]; return sReturn

def fenceIndentBlock(sBlock):
	l = sBlock.split("\n");
	lItems = []
	for i, s in enumerate(l):
		if s: lItems.append("  `" + s + "`\\")
		else: lItems.append(s)
	sReturn = line("\n".join(lItems))
	return sReturn

def listItems(lItems):
	sReturn = ""
	for sItem in lItems: sReturn += itemLine(sItem)
	return sReturn

def graveListItems(lItems):
	sReturn = ""
	for sItem in lItems: sReturn += line("-  "); sReturn += fenceBlock(sItem)
	return sReturn

