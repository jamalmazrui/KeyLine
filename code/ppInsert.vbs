Option Explicit

Const ppLayoutBlank = 12 
Const ppLayoutChart = 8 
Const ppLayoutChartAndText = 6 
Const ppLayoutClipArtAndText = 10 
Const ppLayoutClipArtAndVerticalText = 26 
Const ppLayoutComparison = 34 
Const ppLayoutContentWithCaption = 35 
Const ppLayoutCustom = 32 
Const ppLayoutFourObjects = 24 
Const ppLayoutLargeObject = 15 
Const ppLayoutMediaClipAndText = 18 
Const ppLayoutMixed = -2 
Const ppLayoutObject = 16 
Const ppLayoutObjectAndText = 14 
Const ppLayoutObjectAndTwoObjects = 30 
Const ppLayoutObjectOverText = 19 
Const ppLayoutOrgchart = 7 
Const ppLayoutPictureWithCaption = 36 
Const ppLayoutSectionHeader = 33 
Const ppLayoutTable = 4 
Const ppLayoutText = 2 
Const ppLayoutTextAndChart = 5 
Const ppLayoutTextAndClipArt = 9 
Const ppLayoutTextAndMediaClip = 17 
Const ppLayoutTextAndObject = 13 
Const ppLayoutTextAndTwoObjects = 21 
Const ppLayoutTextOverObject = 20 
Const ppLayoutTitle = 1 
Const ppLayoutTitleOnly = 11 
Const ppLayoutTwoColumnText = 3 
Const ppLayoutTwoObjects = 29 
Const ppLayoutTwoObjectsAndObject = 31 
Const ppLayoutTwoObjectsAndText = 22 
Const ppLayoutTwoObjectsOverText = 23 
Const ppLayoutVerticalText = 25 
Const ppLayoutVerticalTitleAndText = 27 
Const ppLayoutVerticalTitleAndTextOverChart = 28 

Const msoAutomationSecurityLow = 1
Const msoAutomationSecurityByUI = 2
Const msoAutomationSecurityForceDisable = 3

Dim iLayout, iSlide, iBaseIndex, iBodyFirst, iBodyLast
Dim oSlide, oLayouts, oShapes, oApp, oPPTs, oPPT, oBasePpt, oBodyPpt, oSlides
Dim sDir, sBasePptx, sBodyPptx, sTargetPptx

Set oLayouts = CreateObject("Scripting.Dictionary")
oLayouts.Add 12, "ppLayoutBlank"
oLayouts.Add 8, "ppLayoutChart"
oLayouts.Add 6, "ppLayoutChartAndText"
oLayouts.Add 10, "ppLayoutClipArtAndText"
oLayouts.Add 26, "ppLayoutClipArtAndVerticalText"
oLayouts.Add 34, "ppLayoutComparison"
oLayouts.Add 35, "ppLayoutContentWithCaption"
oLayouts.Add 32, "ppLayoutCustom"
oLayouts.Add 24, "ppLayoutFourObjects"
oLayouts.Add 15, "ppLayoutLargeObject"
oLayouts.Add 18, "ppLayoutMediaClipAndText"
oLayouts.Add -2, "ppLayoutMixed"
oLayouts.Add 16, "ppLayoutObject"
oLayouts.Add 14, "ppLayoutObjectAndText"
oLayouts.Add 30, "ppLayoutObjectAndTwoObjects"
oLayouts.Add 19, "ppLayoutObjectOverText"
oLayouts.Add 7, "ppLayoutOrgchart"
oLayouts.Add 36, "ppLayoutPictureWithCaption"
oLayouts.Add 33, "ppLayoutSectionHeader"
oLayouts.Add 4, "ppLayoutTable"
oLayouts.Add 2, "ppLayoutText"
oLayouts.Add 5, "ppLayoutTextAndChart"
oLayouts.Add 9, "ppLayoutTextAndClipArt"
oLayouts.Add 17, "ppLayoutTextAndMediaClip"
oLayouts.Add 13, "ppLayoutTextAndObject"
oLayouts.Add 21, "ppLayoutTextAndTwoObjects"
oLayouts.Add 20, "ppLayoutTextOverObject"
oLayouts.Add 1, "ppLayoutTitle"
oLayouts.Add 11, "ppLayoutTitleOnly"
oLayouts.Add 3, "ppLayoutTwoColumnText"
oLayouts.Add 29, "ppLayoutTwoObjects"
oLayouts.Add 31, "ppLayoutTwoObjectsAndObject"
oLayouts.Add 22, "ppLayoutTwoObjectsAndText"
oLayouts.Add 23, "ppLayoutTwoObjectsOverText"
oLayouts.Add 25, "ppLayoutVerticalText"
oLayouts.Add 27, "ppLayoutVerticalTitleAndText"
oLayouts.Add 28, "ppLayoutVerticalTitleAndTextOverChart"

Function FileInclude(sFile)
With CreateObject("Scripting.FileSystemObject")
ExecuteGlobal .openTextFile(sFile).readAll()
End With

' executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

' Main
Dim sScriptVbs, sHomerLibVbs
Dim oSystem, oFile

sScriptVbs = WScript.ScriptFullName
Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.GetFile(sScriptVbs)
sDir = oSystem.GetParentFolderName(oFile) 
sHomerLibVbs = sDir & "\" & "HomerLib.vbs"
FileInclude sHomerLibVbs

sDir = PathGetCurrentDirectory()
sBasePptx = WScript.Arguments(0)
If Not StringContains(sBasePptx, ":", false) Then sBasePptx = PathCombine(sDir, sBasePptx)
sBodyPptx = WScript.Arguments(1)
If Not StringContains(sBodyPptx, ":", false) Then sBodyPptx = PathCombine(sDir, sBodyPptx)
sTargetPptx = WScript.Arguments(2)
If Not StringContains(sTargetPptx, ":", false) Then sTargetPptx = PathCombine(sDir, sTargetPptx)

iBaseIndex = 1
iBodyFirst = 1

Set oApp = CreateObject("PowerPoint.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.Visible = True ' Needed for automation to work

Set oPpts = oApp.Presentations

' expression.Open(FileName, ReadOnly, Untitled, WithWindow)
if false then
' if true then
Set oBodyPpt = oPpts.Open(sBodyPptx, vbTrue, vbTrue, vbFalse)
iBodyLast = oBodyPpt.Slides.Count
oBodyPpt.Close

Set oPpt = oPpts.Open(sBasePptx, vbTrue, vbTrue, vbFalse)
For iSlide = oPpt.Slides.Count to 1 step -1
set oSlide = oPpt.Slides(iSlide)
oSlide.Delete
Next
' No need to save the empty base, just use it
' oPpt.Save
' oPpt.SaveAs sBasePptx

' Insert slides from sBodyPptx into this deck.

iBaseIndex = 0
oPpt.Slides.InsertFromFile sBodyPptx, iBaseIndex, iBodyFirst, iBodyLast

' Make changes to slides after inserting.

iSlide = 0
For Each oSlide in oPpt.Slides
iSlide = iSlide + 1
On Error Resume Next
' Set the slide name to match the current order.
oSlide.Name = "Slide" & iSlide
On Error GoTo 0
' Show the slide number in the footer.
oSlide.HeadersFooters.SlideNumber.Visible = True
Next
Set oSlides = Nothing

' expression.SaveAs(Filename, FileFormat, EmbedFonts)
' oPpt.SaveAs sTargetFile, iTargetFormat
oPpt.SaveAs sTargetPptx
oPpt.Close()
end if


Set oPpt = Nothing
Set oPpts = Nothing

oApp.Quit()
Set oApp = Nothing
