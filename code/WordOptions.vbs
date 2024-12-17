Option Explicit
WScript.Echo"Starting WordOptions"

Dim aIni
Dim bBackupDocx, bLogActions, bResetNormalTemplate, bValue, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dOther, dOptions, d, dStyle, dIni, dSourceIni
Dim iValue, iArgCount, iCount
Dim oSystem, oFile, oOption, oOptions, oFormat, oFont, oStyle, oApp, oData, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sNormalTemplateDotm, sKey, sTargetLog, sScriptVbs, sHomerLibVbs, sDir, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

' Word constant enumerations used within Options
Const wdDocument = 1 
Const wdEmailMessage = 0
Const wdWebPage = 2

' WdArabicNumeral
' WdAraSpeller
' WdCellColor
' WdColorIndex
' WdCursorMovement
' WdDeletedTextMark
' WdDisableFeaturesIntroducedAfter
' WdDocumentViewDirection
' WdFrenchSpeller
' WdHebSpellStart
' WdHighAnsiText
' WdInsertedTextMark
' WdLineStyle
' WdLineWidth
' WdMeasurementUnits
' WdMonthNames
' WdMoveFromTextMark
' WdMoveToTextMark
' WdMultipleWordConversionsMode
' WdOpenFormat
' WdPaperTray
' WdPasteOptions
' WdPortugueseReform
' WdRevisedLinesMark
' WdRevisedPropertiesMark
' WdRevisionsBalloonPrintOrientation
' WdSpanishSpeller
' WdUpdateStyleListBehavior
' WdVisualSelection
' WdWrapTypeMerged

Dim WdDoNotSaveChanges: WdDoNotSaveChanges = 0

Function FileInclude(sFile)
' With CreateObject("Scripting.FileSystemObject")
' ExecuteGlobal .openTextFile(sFile).readAll()
' End With

executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(sFile).readAll()
FileInclude = True
End Function

Function ResetNormalTemplate
Print "Reset Normal Template"
ProcessTerminateAllModule "WinWord"
' print sNormalTemplateDotm
' WScript.Sleep 5000
If Not FileDelete(sNormalTemplateDotm) Then Print "Error"
End Function

Function ShowAutoCorrectEmail(oApp)
Dim iAutoCorrectEmailCount
Dim oAutoCorrectEmail
Dim sAutoCorrectEmail
iAutoCorrectEmailCount = oApp.AutoCorrectEmail.Entries.Count
if iAutoCorrectEmailCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;AutoCorrectEmail]"
' AppendLine StringPlural("AutoCorrectEmail", iAutoCorrectEmailCount)
AppendLine "[; " & StringPlural("AutoCorrectEmail", iAutoCorrectEmailCount) & "]"
For Each oAutoCorrectEmail in oApp.AutoCorrectEmail.Entries
AppendLine oAutoCorrectEmail.Name & " = " & oAutoCorrectEmail.Value
Next
End Function

Function ShowAutoCorrect(oApp)
Dim iAutoCorrectCount
Dim oAutoCorrect
Dim sAutoCorrect
iAutoCorrectCount = oApp.AutoCorrect.Entries.Count
if iAutoCorrectCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;AutoCorrect]"
' AppendLine StringPlural("AutoCorrect", iAutoCorrectCount)
AppendLine "[; " & StringPlural("AutoCorrect", iAutoCorrectCount) & "]"
For Each oAutoCorrect in oApp.AutoCorrect.Entries
AppendLine oAutoCorrect.Name & " = " & oAutoCorrect.Value
Next
End Function

Function ShowAutoCaptions(oApp)
Dim iAutoCaptionCount
Dim oAutoCaption
Dim sAutoCaption
iAutoCaptionCount = oApp.AutoCaptions.Count
if iAutoCaptionCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;AutoCaptions]"
' AppendLine StringPlural("AutoCaption", iAutoCaptionCount)
AppendLine "[; " & StringPlural("AutoCaption", iAutoCaptionCount) & "]"
For Each oAutoCaption in oApp.AutoCaptions
AppendLine oAutoCaption.Name
Next
End Function

Function ShowCaptionLabels(oApp)
Dim iCaptionLabelCount
Dim oCaptionLabel
Dim sCaptionLabel
iCaptionLabelCount = oApp.CaptionLabels.Count
if iCaptionLabelCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;CaptionLabels]"
' AppendLine StringPlural("CaptionLabel", iCaptionLabelCount)
AppendLine "[; " & StringPlural("CaptionLabel", iCaptionLabelCount) & "]"
For Each oCaptionLabel in oApp.CaptionLabels
AppendLine oCaptionLabel.Name
Next
End Function

Function ShowTaskPanes(oApp)
Dim iTaskPaneCount
Dim oTaskPane
Dim sTaskPane
iTaskPaneCount = oApp.TaskPanes.Count
if iTaskPaneCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;TaskPanes]"
' AppendLine StringPlural("TaskPane", iTaskPaneCount)
AppendLine "[; " & StringPlural("TaskPane", iTaskPaneCount) & "]"
For Each oTaskPane in oApp.TaskPanes
' Property missing
AppendLine oTaskPane.Caption
Next
End Function

Function ShowCommandBars(oApp)
Dim iCommandBarCount
Dim oCommandBar
Dim sCommandBar
iCommandBarCount = oApp.CommandBars.Count
if iCommandBarCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;CommandBars]"
' AppendLine StringPlural("CommandBar", iCommandBarCount)
AppendLine "[; " & StringPlural("CommandBar", iCommandBarCount) & "]"
For Each oCommandBar in oApp.CommandBars
AppendLine oCommandBar.Name
Next
End Function

Function ShowComAddIns(oApp)
Dim iComAddInCount
Dim oComAddIn
Dim sComAddIn
iComAddInCount = oApp.ComAddIns.Count
if iComAddInCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;ComAddIns]"
' AppendLine StringPlural("ComAddIn", iComAddInCount)
AppendLine "[; " & StringPlural("ComAddIn", iComAddInCount) & "]"
For Each oComAddIn in oApp.ComAddIns
' sComAddIn = PathCombine(oComAddIn.Path, oComAddIn.Name)
' sComAddIn = oComAddIn.Name
' print sComAddIn
AppendLine oComAddIn.ProgID & ", " & oComAddIn.Description
Next
End Function

Function ShowAddIns(oApp)
Dim iAddInCount
Dim oAddIn
Dim sAddIn
iAddInCount = oApp.AddIns.Count
if iAddInCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;AddIns]"
' AppendLine StringPlural("AddIn", iAddInCount)
AppendLine "[; " & StringPlural("AddIn", iAddInCount) & "]"
For Each oAddIn in oApp.AddIns
sAddIn = PathCombine(oAddIn.Path, oAddIn.Name)
AppendLine sAddIn
Next
End Function

Function ShowCustomDictionaries(oApp)
Dim iCustomDictionaryCount
Dim oCustomDictionary
Dim sCustomDictionary
iCustomDictionaryCount = oApp.CustomDictionaries.Count
if iCustomDictionaryCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;CustomDictionaries]"
' AppendLine StringPlural("CustomDictionary", iCustomDictionaryCount)
AppendLine "[; " & StringPlural("CustomDictionary", iCustomDictionaryCount) & "]"
For Each oCustomDictionary in oApp.CustomDictionaries
sCustomDictionary = PathCombine(oCustomDictionary.Path, oCustomDictionary.Name)
AppendLine sCustomDictionary
Next
End Function

Function ShowTemplates(oApp)
Dim iTemplateCount
Dim oTemplate
iTemplateCount = oApp.Templates.Count
if iTemplateCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;Templates]"
' AppendLine StringPlural("Template", iTemplateCount)
AppendLine "[; " & StringPlural("Template", iTemplateCount) & "]"
For Each oTemplate in oApp.Templates
AppendLine oTemplate.FullName
Next
End Function

Function ShowFileConverters(oApp)
Dim iFileConverterCount
Dim oFileConverter
iFileConverterCount = oApp.FileConverters.Count
if iFileConverterCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;FileConverters]"
' AppendLine StringPlural("FileConverter", iFileConverterCount)
AppendLine "[; " & StringPlural("FileConverter", iFileConverterCount) & "]"
For Each oFileConverter in oApp.FileConverters
AppendLine oFileConverter.FormatName & " = " & oFileConverter.Extensions
Next
End Function

Function ShowFontNames(oApp)
Dim iFontNameCount
Dim sFontName
iFontNameCount = oApp.FontNames.Count
if iFontNameCount = 0 Then Exit Function
AppendBlank
' AppendLine "[;FontNames]"
' AppendLine StringPlural("FontName", iFontNameCount)
AppendLine "[; " & StringPlural("FontName", iFontNameCount) & "]"
For Each sFontName in oApp.FontNames
AppendLine sFontName
Next
End Function

Function ShowRecentFiles(oApp)
Dim iRecentFile, iRecentFileCount
Dim dRecentFiles
Dim oRecentFile
Dim sRecentFile

iRecentFileCount = oApp.RecentFiles.Count
if iRecentFileCount = 0 Then Exit Function
AppendBlank
' Includes duplicates
' print StringPlural("RecentFile", iRecentFileCount)
Set dRecentFiles = CreateDictionary
iRecentFile = 0
For Each oRecentFile in oApp.RecentFiles
sRecentFile = PathCombine(oRecentFile.Path, oRecentFile.Name)
If Not dRecentFiles.Exists(sRecentFile) Then
iRecentFile = iRecentFile + 1
' print sRecentFile
dRecentFiles.Add sRecentFile, ""
End If
Next
' AppendLine "[;RecentFiles]"
' AppendLine StringPlural("RecentFile", iRecentFile)
AppendLine "[; " & StringPlural("RecentFile", iRecentFile) & "]"
For Each sRecentFile in dRecentFiles.Keys
AppendLine sRecentFile
Next
End Function

Function ShowMisc(oApp)
AppendBlank
AppendLine "[;Miscellaneous]"
' AppendLine "WordBuild = " & oApp.Build
AppendLine "WordVersion = " & oApp.Version
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

If iArgCount > 0 Then
sSourceIni = WScript.Arguments(0)
sSourceIni = GetIniFile(sSourceIni)
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

sTargetIni = PathCombine(PathGetCurrentDirectory(), "OPTIONS.ini")
sTargetLog = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceDocx) & "-WordOptions.log")

bResetNormalTemplate = False
Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False

Set oOptions = oApp.Options

If Not bReadOnly Then
Print "Applying " & PathGetName(sSourceIni)
If dSourceIni.Exists("Options") Then
Set dOptions = dSourceIni("Options")
For Each sAttrib in dOptions.Keys()
sValue = dOptions(sAttrib)
Select Case sAttrib
Case "AddControlCharacters"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AddControlCharacters = CBool(sValue)
Case "AddHebDoubleQuote"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AddHebDoubleQuote = CBool(sValue)
Case "AlertIfNotDefault"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AlertIfNotDefault = CBool(sValue)
Case "AllowAccentedUppercase"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowAccentedUppercase = CBool(sValue)
Case "AllowClickAndTypeMouse"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowClickAndTypeMouse = CBool(sValue)
Case "AllowCombinedAuxiliaryForms"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowCombinedAuxiliaryForms = CBool(sValue)
Case "AllowCompoundNounProcessing"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowCompoundNounProcessing = CBool(sValue)
Case "AllowDragAndDrop"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowDragAndDrop = CBool(sValue)
Case "AllowOpenInDraftView"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowOpenInDraftView = CBool(sValue)
Case "AllowPixelUnits"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowPixelUnits = CBool(sValue)
Case "AllowReadingMode"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AllowReadingMode = CBool(sValue)
Case "AnimateScreenMovements"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AnimateScreenMovements = CBool(sValue)
Case "ApplyFarEastFontsToAscii"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ApplyFarEastFontsToAscii = CBool(sValue)
Case "ArabicMode"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ArabicMode = CBool(sValue)
Case "ArabicNumeral"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ArabicNumeral = CBool(sValue)
Case "AutoCreateNewDrawings"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoCreateNewDrawings = CBool(sValue)
Case "AutoFormatApplyBulletedLists"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatApplyBulletedLists = CBool(sValue)
Case "AutoFormatApplyFirstIndents"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatApplyFirstIndents = CBool(sValue)
Case "AutoFormatApplyHeadings"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatApplyHeadings = CBool(sValue)
Case "AutoFormatApplyLists"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatApplyLists = CBool(sValue)
Case "AutoFormatApplyOtherParas"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatApplyOtherParas = CBool(sValue)
Case "AutoFormatAsYouTypeApplyBorders"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyBorders = CBool(sValue)
Case "AutoFormatAsYouTypeApplyBulletedLists"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyBulletedLists = CBool(sValue)
Case "AutoFormatAsYouTypeApplyClosings"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyClosings = CBool(sValue)
Case "AutoFormatAsYouTypeApplyDates"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyDates = CBool(sValue)
Case "AutoFormatAsYouTypeApplyFirstIndents"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyFirstIndents = CBool(sValue)
Case "AutoFormatAsYouTypeApplyHeadings"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyHeadings = CBool(sValue)
Case "AutoFormatAsYouTypeApplyNumberedLists"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyNumberedLists = CBool(sValue)
Case "AutoFormatAsYouTypeApplyTables"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeApplyTables = CBool(sValue)
Case "AutoFormatAsYouTypeAutoLetterWizard"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeAutoLetterWizard = CBool(sValue)
Case "AutoFormatAsYouTypeDefineStyles"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeDefineStyles = CBool(sValue)
Case "AutoFormatAsYouTypeDeleteAutoSpaces"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeDeleteAutoSpaces = CBool(sValue)
Case "AutoFormatAsYouTypeFormatListItemBeginning"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeFormatListItemBeginning = CBool(sValue)
Case "AutoFormatAsYouTypeInsertClosings"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeInsertClosings = CBool(sValue)
Case "AutoFormatAsYouTypeInsertOvers"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeInsertOvers = CBool(sValue)
Case "AutoFormatAsYouTypeMatchParentheses"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeMatchParentheses = CBool(sValue)
Case "AutoFormatAsYouTypeReplaceFarEastDashes"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplaceFarEastDashes = CBool(sValue)
Case "AutoFormatAsYouTypeReplaceFractions"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplaceFractions = CBool(sValue)
Case "AutoFormatAsYouTypeReplaceHyperlinks"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplaceHyperlinks = CBool(sValue)
Case "AutoFormatAsYouTypeReplaceOrdinals"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplaceOrdinals = CBool(sValue)
Case "AutoFormatAsYouTypeReplacePlainTextEmphasis"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis = CBool(sValue)
Case "AutoFormatAsYouTypeReplaceQuotes"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplaceQuotes = CBool(sValue)
Case "AutoFormatAsYouTypeReplaceSymbols"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatAsYouTypeReplaceSymbols = CBool(sValue)
Case "AutoFormatDeleteAutoSpaces"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatDeleteAutoSpaces = CBool(sValue)
Case "AutoFormatMatchParentheses"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatMatchParentheses = CBool(sValue)
Case "AutoFormatPlainTextWordMail"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatPlainTextWordMail = CBool(sValue)
Case "AutoFormatPreserveStyles"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatPreserveStyles = CBool(sValue)
Case "AutoFormatReplaceFarEastDashes"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplaceFarEastDashes = CBool(sValue)
Case "AutoFormatReplaceFractions"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplaceFractions = CBool(sValue)
Case "AutoFormatReplaceHyperlinks"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplaceHyperlinks = CBool(sValue)
Case "AutoFormatReplaceOrdinals"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplaceOrdinals = CBool(sValue)
Case "AutoFormatReplacePlainTextEmphasis"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplacePlainTextEmphasis = CBool(sValue)
Case "AutoFormatReplaceQuotes"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplaceQuotes = CBool(sValue)
Case "AutoFormatReplaceSymbols"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoFormatReplaceSymbols = CBool(sValue)
Case "AutoKeyboardSwitching"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoKeyboardSwitching = CBool(sValue)
Case "AutoWordSelection"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AutoWordSelection = CBool(sValue)
Case "BackgroundSave"
Print "Setting " & sAttrib & " = " & sValue
oOptions.BackgroundSave = CBool(sValue)
Case "BibliographySort"
Print "Setting " & sAttrib & " = " & sValue
oOptions.BibliographySort = sValue
Case "BibliographyStyle"
Print "Setting " & sAttrib & " = " & sValue
oOptions.BibliographyStyle = sValue
Case "BrazilReform"
Print "Setting " & sAttrib & " = " & sValue
oOptions.BrazilReform = CBool(sValue)
Case "ButtonFieldClicks"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ButtonFieldClicks = CBool(sValue)
Case "CheckGrammarAsYouType"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CheckGrammarAsYouType = CBool(sValue)
Case "CheckGrammarWithSpelling"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CheckGrammarWithSpelling = CBool(sValue)
Case "CheckHangulEndings"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CheckHangulEndings = CBool(sValue)
Case "CheckSpellingAsYouType"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CheckSpellingAsYouType = CBool(sValue)
Case "CloudSignInOption"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CloudSignInOption = CBool(sValue)
Case "CommentsColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CommentsColor = CBool(sValue)
Case "ConfirmConversions"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ConfirmConversions = CBool(sValue)
Case "ContextualSpeller"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ContextualSpeller = CBool(sValue)
Case "ConvertHighAnsiToFarEast"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ConvertHighAnsiToFarEast = CBool(sValue)
Case "CreateBackup"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CreateBackup = CBool(sValue)
Case "CtrlClickHyperlinkToOpen"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CtrlClickHyperlinkToOpen = CBool(sValue)
Case "CursorMovement"
Print "Setting " & sAttrib & " = " & sValue
oOptions.CursorMovement = CBool(sValue)
Case "DefaultBorderColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultBorderColor = sValue
Case "DefaultBorderColorIndex"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultBorderColorIndex = CBool(sValue)
Case "DefaultBorderLineStyle"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultBorderLineStyle = CBool(sValue)
Case "DefaultBorderLineWidth"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultBorderLineWidth = CInt(sValue)
Case "DefaultEPostageApp"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultEPostageApp = CBool(sValue)
Case "DefaultFilePath"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultFilePath = CBool(sValue)
Case "DefaultHighlightColorIndex"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultHighlightColorIndex = CBool(sValue)
Case "DefaultOpenFormat"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultOpenFormat = sValue
Case "DefaultTextEncoding"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultTextEncoding = sValue
Case "DefaultTray"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultTray = sValue
Case "DefaultTrayID"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DefaultTrayID = CBool(sValue)
Case "DeletedCellColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DeletedCellColor = CBool(sValue)
Case "DeletedTextColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DeletedTextColor = CBool(sValue)
Case "DeletedTextMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DeletedTextMark = sValue
Case "DiacriticColorVal"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DiacriticColorVal = CBool(sValue)
Case "DisableFeaturesbyDefault"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DisableFeaturesbyDefault = CBool(sValue)
Case "DisableFeaturesIntroducedAfterbyDefault"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DisableFeaturesIntroducedAfterbyDefault = CBool(sValue)
Case "DisplayAlignmentGuides"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DisplayAlignmentGuides = CBool(sValue)
Case "DisplayGridLines"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DisplayGridLines = CBool(sValue)
Case "DisplayPasteOptions"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DisplayPasteOptions = CBool(sValue)
Case "DocumentViewDirection"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DocumentViewDirection = CBool(sValue)
Case "DoNotPromptForConvert"
Print "Setting " & sAttrib & " = " & sValue
oOptions.DoNotPromptForConvert = CBool(sValue)
Case "EnableHangulHanjaRecentOrdering"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableHangulHanjaRecentOrdering = CBool(sValue)
Case "EnableLegacyIMEMode"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableLegacyIMEMode = CBool(sValue)
Case "EnableLiveDrag"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableLiveDrag = CBool(sValue)
Case "EnableLivePreview"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableLivePreview = CBool(sValue)
Case "EnableMisusedWordsDictionary"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableMisusedWordsDictionary = CBool(sValue)
Case "EnableProofingToolsAdvertisement"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableProofingToolsAdvertisement = CBool(sValue)
Case "EnableSound"
Print "Setting " & sAttrib & " = " & sValue
oOptions.EnableSound = CBool(sValue)
Case "EnvelopeFeederInstalled"
' readonly so cannot set
' Print "Setting " & sAttrib & " = " & sValue
' oOptions.EnvelopeFeederInstalled = CBool(sValue)
Case "ExpandHeadingsOnOpen"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ExpandHeadingsOnOpen = CBool(sValue)
Case "FormatScanning"
Print "Setting " & sAttrib & " = " & sValue
oOptions.FormatScanning = CBool(sValue)
Case "FrenchReform"
Print "Setting " & sAttrib & " = " & sValue
oOptions.FrenchReform = CBool(sValue)
Case "GridDistanceHorizontal"
Print "Setting " & sAttrib & " = " & sValue
' oOptions.GridDistanceHorizontal = CBool(sValue)
Case "GridDistanceVertical"
Print "Setting " & sAttrib & " = " & sValue
oOptions.GridDistanceVertical = CBool(sValue)
Case "GridOriginHorizontal"
Print "Setting " & sAttrib & " = " & sValue
oOptions.GridOriginHorizontal = CBool(sValue)
Case "GridOriginVertical"
Print "Setting " & sAttrib & " = " & sValue
oOptions.GridOriginVertical = CBool(sValue)
Case "HangulHanjaFastConversion"
Print "Setting " & sAttrib & " = " & sValue
oOptions.HangulHanjaFastConversion = CBool(sValue)
Case "HebrewMode"
Print "Setting " & sAttrib & " = " & sValue
oOptions.HebrewMode = CBool(sValue)
Case "AddBiDirectionalMarksWhenSavingTextFile"
Print "Setting " & sAttrib & " = " & sValue
oOptions.AddBiDirectionalMarksWhenSavingTextFile = CBool(sValue)
Case "IgnoreInternetAndFileAddresses"
Print "Setting " & sAttrib & " = " & sValue
oOptions.IgnoreInternetAndFileAddresses = CBool(sValue)
Case "IgnoreMixedDigits"
Print "Setting " & sAttrib & " = " & sValue
oOptions.IgnoreMixedDigits = CBool(sValue)
Case "IgnoreUppercase"
Print "Setting " & sAttrib & " = " & sValue
oOptions.IgnoreUppercase = CBool(sValue)
Case "IMEAutomaticControl"
Print "Setting " & sAttrib & " = " & sValue
oOptions.IMEAutomaticControl = CBool(sValue)
Case "InlineConversion"
Print "Setting " & sAttrib & " = " & sValue
oOptions.InlineConversion = CBool(sValue)
Case "InsertedCellColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.InsertedCellColor = CBool(sValue)
Case "InsertedTextColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.InsertedTextColor = CBool(sValue)
Case "InsertedTextMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.InsertedTextMark = sValue
Case "INSKeyForOvertype"
Print "Setting " & sAttrib & " = " & sValue
oOptions.INSKeyForOvertype = CBool(sValue)
Case "INSKeyForPaste"
Print "Setting " & sAttrib & " = " & sValue
oOptions.INSKeyForPaste = CBool(sValue)
Case "InterpretHighAnsi"
Print "Setting " & sAttrib & " = " & sValue
oOptions.InterpretHighAnsi = sValue
Case "LocalNetworkFile"
Print "Setting " & sAttrib & " = " & sValue
oOptions.LocalNetworkFile = CBool(sValue)
Case "MapPaperSize"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MapPaperSize = CBool(sValue)
Case "MarginAlignmentGuides"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MarginAlignmentGuides = CBool(sValue)
Case "MatchFuzzyAY"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyAY = CBool(sValue)
Case "MatchFuzzyByte"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyByte = CBool(sValue)
Case "MatchFuzzyCase"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyCase = CBool(sValue)
Case "MatchFuzzyDash"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyDash = CBool(sValue)
Case "MatchFuzzyDZ"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyDZ = CBool(sValue)
Case "MatchFuzzyHiragana"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyHiragana = CBool(sValue)
Case "MatchFuzzyIterationMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyIterationMark = CBool(sValue)
Case "MatchFuzzyKanji"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyKanji = CBool(sValue)
Case "MatchFuzzyKiKu"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyKiKu = CBool(sValue)
Case "MatchFuzzyOldKana"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyOldKana = CBool(sValue)
Case "MatchFuzzyProlongedSoundMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyProlongedSoundMark = CBool(sValue)
Case "MatchFuzzyPunctuation"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzyPunctuation = CBool(sValue)
Case "MatchFuzzySmallKana"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzySmallKana = CBool(sValue)
Case "MatchFuzzySpace"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MatchFuzzySpace = CBool(sValue)
Case "MeasurementUnit"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MeasurementUnit = CBool(sValue)
Case "MergedCellColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MergedCellColor = CBool(sValue)
Case "MonthNames"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MonthNames = CBool(sValue)
Case "MoveFromTextColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MoveFromTextColor = CBool(sValue)
Case "MoveFromTextMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MoveFromTextMark = sValue
Case "MoveToTextColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MoveToTextColor = CBool(sValue)
Case "MoveToTextMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MoveToTextMark = sValue
Case "MultipleWordConversionsMode"
Print "Setting " & sAttrib & " = " & sValue
oOptions.MultipleWordConversionsMode = CBool(sValue)
Case "OMathAutoBuildUp"
Print "Setting " & sAttrib & " = " & sValue
oOptions.OMathAutoBuildUp = CBool(sValue)
Case "OMathCopyLF"
Print "Setting " & sAttrib & " = " & sValue
oOptions.OMathCopyLF = CBool(sValue)
Case "OptimizeForWord97byDefault"
Print "Setting " & sAttrib & " = " & sValue
oOptions.OptimizeForWord97byDefault = CBool(sValue)
Case "Overtype"
Print "Setting " & sAttrib & " = " & sValue
oOptions.Overtype = CBool(sValue)
Case "PageAlignmentGuides"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PageAlignmentGuides = CBool(sValue)
Case "Pagination"
Print "Setting " & sAttrib & " = " & sValue
oOptions.Pagination = CBool(sValue)
Case "ParagraphAlignmentGuides"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ParagraphAlignmentGuides = CBool(sValue)
Case "PasteAdjustParagraphSpacing"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteAdjustParagraphSpacing = CBool(sValue)
Case "PasteAdjustTableFormatting"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteAdjustTableFormatting = CBool(sValue)
Case "PasteAdjustWordSpacing"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteAdjustWordSpacing = CBool(sValue)
Case "PasteFormatBetweenDocuments"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteFormatBetweenDocuments = sValue
Case "PasteFormatBetweenStyledDocuments"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteFormatBetweenStyledDocuments = sValue
Case "PasteFormatFromExternalSource"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteFormatFromExternalSource = sValue
Case "PasteFormatWithinDocument"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteFormatWithinDocument = sValue
Case "PasteMergeFromPPT"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteMergeFromPPT = CBool(sValue)
Case "PasteMergeFromXL"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteMergeFromXL = CBool(sValue)
Case "PasteMergeLists"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteMergeLists = CBool(sValue)
Case "PasteOptionKeepBulletsAndNumbers"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteOptionKeepBulletsAndNumbers = CBool(sValue)
Case "PasteSmartCutPaste"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteSmartCutPaste = CBool(sValue)
Case "PasteSmartStyleBehavior"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PasteSmartStyleBehavior = CBool(sValue)
Case "PictureEditor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PictureEditor = sValue
Case "PictureWrapType"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PictureWrapType = CBool(sValue)
Case "PortugalReform"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PortugalReform = sValue
Case "PrecisePositioning"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrecisePositioning = CBool(sValue)
Case "PreferCloudSaveLocations"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PreferCloudSaveLocations = CBool(sValue)
Case "PrintBackground"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintBackground = CBool(sValue)
Case "PrintBackgrounds"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintBackgrounds = CBool(sValue)
Case "PrintComments"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintComments = CBool(sValue)
Case "PrintDraft"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintDraft = CBool(sValue)
Case "PrintDrawingObjects"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintDrawingObjects = CBool(sValue)
Case "PrintEvenPagesInAscendingOrder"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintEvenPagesInAscendingOrder = CBool(sValue)
Case "PrintFieldCodes"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintFieldCodes = CBool(sValue)
Case "PrintHiddenText"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintHiddenText = CBool(sValue)
Case "PrintOddPagesInAscendingOrder"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintOddPagesInAscendingOrder = CBool(sValue)
Case "PrintProperties"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintProperties = CBool(sValue)
Case "PrintReverse"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintReverse = CBool(sValue)
Case "PrintXMLTag"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PrintXMLTag = CBool(sValue)
Case "PromptUpdateStyle"
Print "Setting " & sAttrib & " = " & sValue
oOptions.PromptUpdateStyle = CBool(sValue)
Case "RepeatWord"
Print "Setting " & sAttrib & " = " & sValue
oOptions.RepeatWord = CBool(sValue)
Case "ReplaceSelection"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ReplaceSelection = CBool(sValue)
Case "RevisedLinesColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.RevisedLinesColor = CBool(sValue)
Case "RevisedLinesMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.RevisedLinesMark = CBool(sValue)
Case "RevisedPropertiesColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.RevisedPropertiesColor = CBool(sValue)
Case "RevisedPropertiesMark"
Print "Setting " & sAttrib & " = " & sValue
oOptions.RevisedPropertiesMark = CBool(sValue)
Case "RevisionsBalloonPrintOrientation"
Print "Setting " & sAttrib & " = " & sValue
oOptions.RevisionsBalloonPrintOrientation = CBool(sValue)
Case "SaveInterval"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SaveInterval = CBool(sValue)
Case "SaveNormalPrompt"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SaveNormalPrompt = CBool(sValue)
Case "SavePropertiesPrompt"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SavePropertiesPrompt = CBool(sValue)
Case "SendMailAttach"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SendMailAttach = CBool(sValue)
Case "SequenceCheck"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SequenceCheck = CBool(sValue)
Case "ShowControlCharacters"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowControlCharacters = CBool(sValue)
Case "ShowDevTools"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowDevTools = CBool(sValue)
Case "ShowDiacritics"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowDiacritics = CBool(sValue)
Case "ShowFormatError"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowFormatError = CBool(sValue)
Case "ShowMarkupOpenSave"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowMarkupOpenSave = CBool(sValue)
Case "ShowMenuFloaties"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowMenuFloaties = CBool(sValue)
Case "ShowReadabilityStatistics"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowReadabilityStatistics = CBool(sValue)
Case "ShowSelectionFloaties"
Print "Setting " & sAttrib & " = " & sValue
oOptions.ShowSelectionFloaties = CBool(sValue)
Case "SmartCursoring"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SmartCursoring = CBool(sValue)
Case "SmartCutPaste"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SmartCutPaste = CBool(sValue)
Case "SmartParaSelection"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SmartParaSelection = CBool(sValue)
Case "SnapToGrid"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SnapToGrid = CBool(sValue)
Case "SnapToShapes"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SnapToShapes = CBool(sValue)
Case "SpanishMode"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SpanishMode = CBool(sValue)
Case "SplitCellColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SplitCellColor = CBool(sValue)
Case "StoreRSIDOnSave"
Print "Setting " & sAttrib & " = " & sValue
oOptions.StoreRSIDOnSave = CBool(sValue)
Case "StrictFinalYaa"
Print "Setting " & sAttrib & " = " & sValue
oOptions.StrictFinalYaa = CBool(sValue)
Case "StrictInitialAlefHamza"
Print "Setting " & sAttrib & " = " & sValue
oOptions.StrictInitialAlefHamza = CBool(sValue)
Case "StrictRussianE"
Print "Setting " & sAttrib & " = " & sValue
oOptions.StrictRussianE = CBool(sValue)
Case "StrictTaaMarboota"
Print "Setting " & sAttrib & " = " & sValue
oOptions.StrictTaaMarboota = CBool(sValue)
Case "SuggestFromMainDictionaryOnly"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SuggestFromMainDictionaryOnly = CBool(sValue)
Case "SuggestSpellingCorrections"
Print "Setting " & sAttrib & " = " & sValue
oOptions.SuggestSpellingCorrections = CBool(sValue)
Case "TabIndentKey"
Print "Setting " & sAttrib & " = " & sValue
oOptions.TabIndentKey = CBool(sValue)
Case "TypeNReplace"
Print "Setting " & sAttrib & " = " & sValue
oOptions.TypeNReplace = CBool(sValue)
Case "UpdateFieldsAtPrint"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UpdateFieldsAtPrint = CBool(sValue)
Case "UpdateFieldsWithTrackedChangesAtPrint"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UpdateFieldsWithTrackedChangesAtPrint = CBool(sValue)
Case "UpdateLinksAtOpen"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UpdateLinksAtOpen = CBool(sValue)
Case "UpdateLinksAtPrint"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UpdateLinksAtPrint = CBool(sValue)
Case "UpdateStyleListBehavior"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UpdateStyleListBehavior = CBool(sValue)
Case "UseCharacterUnit"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UseCharacterUnit = CBool(sValue)
Case "UseDiffDiacColor"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UseDiffDiacColor = CBool(sValue)
Case "UseGermanSpellingReform"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UseGermanSpellingReform = CBool(sValue)
Case "UseLocalUserInfo"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UseLocalUserInfo = CBool(sValue)
Case "UseNormalStyleForList"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UseNormalStyleForList = CBool(sValue)
Case "UseSubPixelPositioning"
Print "Setting " & sAttrib & " = " & sValue
oOptions.UseSubPixelPositioning = CBool(sValue)
Case "VisualSelection"
Print "Setting " & sAttrib & " = " & sValue
oOptions.VisualSelection = CBool(sValue)
Case "WarnBeforeSavingPrintingSendingMarkup"
Print "Setting " & sAttrib & " = " & sValue
oOptions.WarnBeforeSavingPrintingSendingMarkup = CBool(sValue)
End Select
Next
End If

if dSourceIni.Exists("Other") Then
Set dOther = dSourceIni("Other")
For Each sKey in dOther.Keys()
sValue = dOther(sKey)
bValue = ForceBool(sValue)
iValue = ForceInt(sValue)
Select Case sKey
Case "ActivePrinter"
Print "Setting " & sKey & " = " & bValue
oApp.ActivePrinter = bValue
Case "BibliographyStyle"
Print "Setting " & sKey & " = " & bValue
oApp.Bibliography.BibliographyStyle = iValue
Case "CheckLanguage"
Print "Setting " & sKey & " = " & bValue
oApp.CheckLanguage = bValue
Case "DefaultTableSeparator"
Print "Setting " & sKey & " = " & sValue
oApp.DefaultTableSeparator = sValue
Case "EmailOptions"
Print "Setting " & sKey & " = " & bValue
Case "FileValidation"
Print "Setting " & sKey & " = " & iValue
oApp.FileValidation = iValue
Case "International"
Print "Setting " & sKey & " = " & bValue
Case "OpenAttachmentsInFullScreen"
Print "Setting " & sKey & " = " & bValue
oApp.OpenAttachmentsInFullScreen = bValue
Case "PrintPreview"
Print "Setting " & sKey & " = " & bValue
oApp.PrintPreview = bValue
Case "RestrictLinkedStyles"
Print "Setting " & sKey & " = " & bValue
oApp.RestrictLinkedStyles = bValue
Case "ShowStartupDialog"
' requires a document window open
oApp.Visible = True
oApp.CommandBars("Task Pane").Visible = False
Set oDoc = oApp.Documents.Add
Print "Setting " & sKey & " = " & bValue
oApp.ShowStartupDialog = bValue
oDoc.Close wdDoNotSaveChanges
Case "UserAddress"
Print "Setting " & sKey & " = " & sValue
oApp.UserAddress = sValue
Case "UserInitials"
Print "Setting " & sKey & " = " & sValue
oApp.UserInitials = sValue
Case "UserName"
Print "Setting " & sKey & " = " & sValue
oApp.UserName = sValue
Case "Version"
Print "Setting " & sKey & " = " & bValue
Case "DefaultWebOptions"
Print "Setting " & sKey & " = " & bValue
Case "SetDefaultTheme"
Print "Setting " & sKey & " = " & sValue
oApp.SetDefaultTheme sValue, wdDocument
Case "ResetNormalTemplate"
If bValue Then
sNormalTemplateDotm = oApp.NormalTemplate.FullName
bResetNormalTemplate = True
End If
End Select
Next
End If ' dSourceIni.Exists("Other")
End If ' Not bReadOnly

' Create target ini
Print "Creating " & PathGetName(sTargetIni)
AppendLine "[Options]"
' True if Word adds bidirectional control characters when cutting and copying text.
AppendLine "AddControlCharacters = " & oOptions.AddControlCharacters

' True if Word encloses number formats in double quotation marks (").
' AppendLine "AddHebDoubleQuote = " & oOptions.AddHebDoubleQuote

' True if users are notified if Word is not the default program for viewing and editing documents.
AppendLine "AlertIfNotDefault = " & oOptions.AlertIfNotDefault

' True if accents are retained when a French language character is changed to uppercase.
' AppendLine "AllowAccentedUppercase = " & oOptions.AllowAccentedUppercase

' True if Click and Type functionality is enabled.
AppendLine "AllowClickAndTypeMouse = " & oOptions.AllowClickAndTypeMouse

' True if Word ignores auxiliary verb forms when checking spelling in a Korean language document.
AppendLine "AllowCombinedAuxiliaryForms = " & oOptions.AllowCombinedAuxiliaryForms

' True if Word ignores compound nouns when checking spelling in a Korean language document.
AppendLine "AllowCompoundNounProcessing = " & oOptions.AllowCompoundNounProcessing

' True if dragging can be used to move or copy a selection.
AppendLine "AllowDragAndDrop = " & oOptions.AllowDragAndDrop

' True to allow users to open documents in draft view.
AppendLine "AllowOpenInDraftView = " & oOptions.AllowOpenInDraftView

' True if Word uses pixels as the default unit of measurement for HTML features that support measurements.
AppendLine "AllowPixelUnits = " & oOptions.AllowPixelUnits

' True if Word opens documents in Reading Layout view.
AppendLine "AllowReadingMode = " & oOptions.AllowReadingMode

' True if Word animates mouse movements, uses animated cursors, and animates actions such as background saving and find and replace operations.
AppendLine "AnimateScreenMovements = " & oOptions.AnimateScreenMovements

' True if Word applies East Asian fonts to Latin text.
' AppendLine "ApplyFarEastFontsToAscii = " & oOptions.ApplyFarEastFontsToAscii

' Integer for the mode of the Arabic spell checker. WdAraSpeller.
' AppendLine "ArabicMode = " & oOptions.ArabicMode

' Integer for the numeral style of an Arabic language document. WdArabicNumeral.
AppendLine "ArabicNumeral = " & oOptions.ArabicNumeral

' True for Word to draw newly created shapes in a drawing canvas.
AppendLine "AutoCreateNewDrawings = " & oOptions.AutoCreateNewDrawings

' True if characters (such as asterisks, hyphens, and greater-than signs) at the beginning of list paragraphs are replaced with bullets from the Bullets and Numbering dialog box ( Format menu) when Word formats a document or range automatically.
AppendLine "AutoFormatApplyBulletedLists = " & oOptions.AutoFormatApplyBulletedLists

' True if Word replaces a space entered at the beginning of a paragraph with a first-line indent when Word formats a document or range automatically.
AppendLine "AutoFormatApplyFirstIndents = " & oOptions.AutoFormatApplyFirstIndents

' True if styles are automatically applied to headings when Word formats a document or range automatically.
AppendLine "AutoFormatApplyHeadings = " & oOptions.AutoFormatApplyHeadings

' True if styles are automatically applied to lists when Word formats a document or range automatically.
AppendLine "AutoFormatApplyLists = " & oOptions.AutoFormatApplyLists

' True if styles are automatically applied to paragraphs that aren't headings or list items when Word formats a document or range automatically.
AppendLine "AutoFormatApplyOtherParas = " & oOptions.AutoFormatApplyOtherParas

' True if a series of three or more hyphens (-), equal signs (=), or underscore characters (\_) are automatically replaced by a specific border line when the ENTER key is pressed.
AppendLine "AutoFormatAsYouTypeApplyBorders = " & oOptions.AutoFormatAsYouTypeApplyBorders

' True if bullet characters (such as asterisks, hyphens, and greater-than signs) are replaced with bullets from the Bullets And Numbering dialog box ( Format menu) as you type.
AppendLine "AutoFormatAsYouTypeApplyBulletedLists = " & oOptions.AutoFormatAsYouTypeApplyBulletedLists

' True for Word to automatically apply the Closing style to letter closings as you type.
AppendLine "AutoFormatAsYouTypeApplyClosings = " & oOptions.AutoFormatAsYouTypeApplyClosings

' True for Word to automatically apply the Date style to dates as you type.
AppendLine "AutoFormatAsYouTypeApplyDates = " & oOptions.AutoFormatAsYouTypeApplyDates

' True for Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent.
AppendLine "AutoFormatAsYouTypeApplyFirstIndents = " & oOptions.AutoFormatAsYouTypeApplyFirstIndents

' True if styles are automatically applied to headings as you type.
AppendLine "AutoFormatAsYouTypeApplyHeadings = " & oOptions.AutoFormatAsYouTypeApplyHeadings

' True if paragraphs are automatically formatted as numbered lists with a numbering scheme from the Bullets and Numbering dialog box ( Format menu), according to what's typed. For example, if a paragraph starts with "1.1" and a tab character, Word automatically inserts "1.2" and a tab character after the ENTER key is pressed.
AppendLine "AutoFormatAsYouTypeApplyNumberedLists = " & oOptions.AutoFormatAsYouTypeApplyNumberedLists

' True if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. The plus signs become the column borders, and the hyphens become the column widths.
AppendLine "AutoFormatAsYouTypeApplyTables = " & oOptions.AutoFormatAsYouTypeApplyTables

' True for Word to automatically start the Letter Wizard when the user enters a letter salutation or closing.
AppendLine "AutoFormatAsYouTypeAutoLetterWizard = " & oOptions.AutoFormatAsYouTypeAutoLetterWizard

' True if Word automatically creates new styles based on manual formatting.
AppendLine "AutoFormatAsYouTypeDefineStyles = " & oOptions.AutoFormatAsYouTypeDefineStyles

' True for Word to automatically delete spaces inserted between Japanese and Latin text as you type.
AppendLine "AutoFormatAsYouTypeDeleteAutoSpaces = " & oOptions.AutoFormatAsYouTypeDeleteAutoSpaces

' True if Word repeats character formatting applied to the beginning of a list item to the next list item.
AppendLine "AutoFormatAsYouTypeFormatListItemBeginning = " & oOptions.AutoFormatAsYouTypeFormatListItemBeginning

' True for Word to automatically insert the corresponding memo closing when the user enters a memo heading.
AppendLine "AutoFormatAsYouTypeInsertClosings = " & oOptions.AutoFormatAsYouTypeInsertClosings

' True for Word to automatically insert "" when the user enters "" or "".
AppendLine "AutoFormatAsYouTypeInsertOvers = " & oOptions.AutoFormatAsYouTypeInsertOvers

' True for Word to automatically correct improperly paired parentheses.
AppendLine "AutoFormatAsYouTypeMatchParentheses = " & oOptions.AutoFormatAsYouTypeMatchParentheses

' True for Word to automatically correct long vowel sounds and dashes.
' AppendLine "AutoFormatAsYouTypeReplaceFarEastDashes = " & oOptions.AutoFormatAsYouTypeReplaceFarEastDashes

' True if typed fractions are replaced with fractions from the current character set as you type. For example, "1/2" is replaced with ".".
AppendLine "AutoFormatAsYouTypeReplaceFractions = " & oOptions.AutoFormatAsYouTypeReplaceFractions

' True if email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically changed to hyperlinks as you type.
AppendLine "AutoFormatAsYouTypeReplaceHyperlinks = " & oOptions.AutoFormatAsYouTypeReplaceHyperlinks

' True if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript as you type. For example, "1st" is replaced with "1" followed by "st" formatted as superscript.
AppendLine "AutoFormatAsYouTypeReplaceOrdinals = " & oOptions.AutoFormatAsYouTypeReplaceOrdinals

' True if manual emphasis characters are automatically replaced with character formatting as you type. For example, "*bold*" is changed to " bold " and "*underline*" is changed to "underline.".
AppendLine "AutoFormatAsYouTypeReplacePlainTextEmphasis = " & oOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis

' True if straight quotation marks are automatically changed to smart (curly) quotation marks as you type.
AppendLine "AutoFormatAsYouTypeReplaceQuotes = " & oOptions.AutoFormatAsYouTypeReplaceQuotes

' True if two consecutive hyphens (--) are replaced with an en dash (-) or an em dash () as you type.If the hyphens are typed with leading and trailing spaces, Word replaces the hyphens with an en dash; if there are no trailing spaces, the hyphens are replaced with an em dash.
AppendLine "AutoFormatAsYouTypeReplaceSymbols = " & oOptions.AutoFormatAsYouTypeReplaceSymbols

' True if spaces inserted between Japanese and Latin text will be deleted when Word formats a document or range automatically.
AppendLine "AutoFormatDeleteAutoSpaces = " & oOptions.AutoFormatDeleteAutoSpaces

' True if improperly paired parentheses are corrected when Word formats a document or range automatically.
AppendLine "AutoFormatMatchParentheses = " & oOptions.AutoFormatMatchParentheses

' True if Word automatically formats plain-text email messages when you open them in Word.
AppendLine "AutoFormatPlainTextWordMail = " & oOptions.AutoFormatPlainTextWordMail

' True if previously applied styles are preserved when Word formats a document or range automatically.
AppendLine "AutoFormatPreserveStyles = " & oOptions.AutoFormatPreserveStyles

' True if long vowel sound and dash use is corrected when Word formats a document or range automatically.
' AppendLine "AutoFormatReplaceFarEastDashes = " & oOptions.AutoFormatReplaceFarEastDashes

' True if typed fractions are replaced with fractions from the current character set when Word formats a document or range automatically. For example, "1/2" is replaced with ".".
AppendLine "AutoFormatReplaceFractions = " & oOptions.AutoFormatReplaceFractions

' True if email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically formatted whenever Word AutoFormats a document or range.
AppendLine "AutoFormatReplaceHyperlinks = " & oOptions.AutoFormatReplaceHyperlinks

' True if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript when Word formats a document or range automatically. For example, "1st" is replaced with "1" followed by "st" formatted as superscript.
AppendLine "AutoFormatReplaceOrdinals = " & oOptions.AutoFormatReplaceOrdinals

' True if manual emphasis characters are replaced with character formatting when Word formats a document or range automatically. For example, "*bold*" is changed to "bold" and "*underline*" is changed to "underline.".
AppendLine "AutoFormatReplacePlainTextEmphasis = " & oOptions.AutoFormatReplacePlainTextEmphasis

' True if straight quotation marks are automatically changed to smart (curly) quotation marks when Word formats a document or range automatically.
AppendLine "AutoFormatReplaceQuotes = " & oOptions.AutoFormatReplaceQuotes

' True if two consecutive hyphens (--) are replaced by an en dash (-) or an em dash () when Word formats a document or range automatically.
AppendLine "AutoFormatReplaceSymbols = " & oOptions.AutoFormatReplaceSymbols

' True if Word automatically switches the keyboard language to match what you are typing at any given time.
' AppendLine "AutoKeyboardSwitching = " & oOptions.AutoKeyboardSwitching

' True if dragging selects one word at a time instead of one character at a time.
AppendLine "AutoWordSelection = " & oOptions.AutoWordSelection

' True if Word saves documents in the background. When Word is saving in the background, users can continue to type and to choose commands.
AppendLine "BackgroundSave = " & oOptions.BackgroundSave

' String for the order in which to display sources in the Source Manager dialog box.
AppendLine "BibliographySort = " & oOptions.BibliographySort

' String for the name of the style to use for formatting bibliographies.
AppendLine "BibliographyStyle = " & oOptions.BibliographyStyle

' Integer for the mode of the Brazilian Portuguese spell checker. WdPortugueseReform.
' AppendLine "BrazilReform = " & oOptions.BrazilReform

' Integer for the number of clicks (either one or two) required to run a GOTOBUTTON or MACROBUTTON field.
AppendLine "ButtonFieldClicks = " & oOptions.ButtonFieldClicks

' True if Word checks grammar and marks errors automatically as you type.
AppendLine "CheckGrammarAsYouType = " & oOptions.CheckGrammarAsYouType

' True if Word checks grammar while checking spelling.
AppendLine "CheckGrammarWithSpelling = " & oOptions.CheckGrammarWithSpelling

' True if Word automatically detects Hangul endings and ignores them during conversion from Hangul to Hanja.
' AppendLine "CheckHangulEndings = " & oOptions.CheckHangulEndings

' True if Word checks spelling and marks errors automatically as you type.
AppendLine "CheckSpellingAsYouType = " & oOptions.CheckSpellingAsYouType

' True to give users the option to sign in to Microsoft OneDrive and other cloud locations.
AppendLine "CloudSignInOption = " & oOptions.CloudSignInOption

' Integer for the color of comments in a document.  WdColorIndex.
AppendLine "CommentsColor = " & oOptions.CommentsColor

' True if Word displays the Convert File dialog box before it opens or inserts a file that isn't a Word document or template. In the Convert File dialog box, the user chooses the format to convert the file from.
AppendLine "ConfirmConversions = " & oOptions.ConfirmConversions

' True to use the contextual spell checker to check spelling based on the context of a word and the words around it.
AppendLine "ContextualSpeller = " & oOptions.ContextualSpeller

' True if Word converts text that is associated with an East Asian font to the appropriate font when it opens a document.
' AppendLine "ConvertHighAnsiToFarEast = " & oOptions.ConvertHighAnsiToFarEast

' True if Word creates a backup copy each time a document is saved.
AppendLine "CreateBackup = " & oOptions.CreateBackup

' True if Word requires holding down the Ctrl key while clicking to open a hyperlink.
AppendLine "CtrlClickHyperlinkToOpen = " & oOptions.CtrlClickHyperlinkToOpen

' Integer for how the insertion point progresses within bidirectional text. WdCursorMovement.
AppendLine "CursorMovement = " & oOptions.CursorMovement

' Integer for the default 24-bit color to use for new [Border](word.border) objects.
AppendLine "DefaultBorderColor = " & oOptions.DefaultBorderColor

' Integer for the default line color for borders. WdColorIndex.
AppendLine "DefaultBorderColorIndex = " & oOptions.DefaultBorderColorIndex

' Integer for the default border line style. WdLineStyle.
AppendLine "DefaultBorderLineStyle = " & oOptions.DefaultBorderLineStyle

' Integer for the default line width of borders. WdLineWidth.
AppendLine "DefaultBorderLineWidth = " & oOptions.DefaultBorderLineWidth

' Error (object not available)
' oOptions.DefaultEPostageApp =
' String for the path and file name of the default electronic postage application.
' AppendLine "DefaultEPostageApp = " & oOptions.DefaultEPostageApp

' Error (invalid parameters)
' oOptions.DefaultFilePath =
' String for default folders for items such as documents, templates, and graphics.
' AppendLine "DefaultFilePath = " & oOptions.DefaultFilePath

' Integer for the color used to highlight text formatted with the Highlight button ( Formatting toolbar). WdColorIndex.
AppendLine "DefaultHighlightColorIndex = " & oOptions.DefaultHighlightColorIndex

' Integer for the default file converter used to open documents. Can be a number returned by the OpenFormat property, or one of the following WdOpenFormat constants.
AppendLine "DefaultOpenFormat = " & oOptions.DefaultOpenFormat

' Integer for the MsoEncoding constant representing the code page, or character set, that Word uses for all documents saved as encoded text files.
AppendLine "DefaultTextEncoding = " & oOptions.DefaultTextEncoding

' String for the default tray your printer uses to print documents.
AppendLine "DefaultTray = " & oOptions.DefaultTray

' Integer for the default tray your printer uses to print documents. WdPaperTray.
AppendLine "DefaultTrayID = " & oOptions.DefaultTrayID

' Integer for the color for a deleted cell. WdCellColor.
AppendLine "DeletedCellColor = " & oOptions.DeletedCellColor

' Integer for the color of text that is deleted while change tracking is enabled. WdColorIndex.
AppendLine "DeletedTextColor = " & oOptions.DeletedTextColor

' Integer for the format of text that is deleted while change tracking is enabled. WdDeletedTextMark.
AppendLine "DeletedTextMark = " & oOptions.DeletedTextMark

' Integer for the 24-bit color to be used for diacritics in a right-to-left language document.
' AppendLine "DiacriticColorVal = " & oOptions.DiacriticColorVal

' True for Word to disable in all documents all features introduced after the version of Word specified in the [DisableFeaturesIntroducedAfterbyDefault](word.options.disablefeaturesintroducedafterbydefault) . The default value is False.
' AppendLine "DisableFeaturesbyDefault = " & oOptions.DisableFeaturesbyDefault

' Integer to disable all features introduced after the specified version for all documents. WdDisableFeaturesIntroducedAfter.
' AppendLine "DisableFeaturesIntroducedAfterbyDefault = " & oOptions.DisableFeaturesIntroducedAfterbyDefault

' True if alignment guides are enabled in the user interface.
AppendLine "DisplayAlignmentGuides = " & oOptions.DisplayAlignmentGuides

' True if Word displays the document grid. This property is the equivalent of the Gridlines command on the View menu.
AppendLine "DisplayGridLines = " & oOptions.DisplayGridLines

' True for Word to display the Paste Options button, which displays directly under newly pasted text.
AppendLine "DisplayPasteOptions = " & oOptions.DisplayPasteOptions

' Integer for the alignment and reading order of the entire document. WdDocumentViewDirection.
AppendLine "DocumentViewDirection = " & oOptions.DocumentViewDirection

' True to prompt a warning dialog when the Convert command is invoked for documents that are in compatibility mode.
AppendLine "DoNotPromptForConvert = " & oOptions.DoNotPromptForConvert

' True if Word displays the most recently used words at the top of the suggestions list during conversion between Hangul and Hanja.
' AppendLine "EnableHangulHanjaRecentOrdering = " & oOptions.EnableHangulHanjaRecentOrdering

' True to enable legacy IME mode.
' AppendLine "EnableLegacyIMEMode = " & oOptions.EnableLegacyIMEMode

' True if live drag is enabled.
AppendLine "EnableLiveDrag = " & oOptions.EnableLiveDrag

' True to show or hide gallery previews that appear when using galleries that support previewing. True shows a preview in your document before applying the command.
AppendLine "EnableLivePreview = " & oOptions.EnableLivePreview

' True if Word checks for misused words when checking the spelling and grammar in a document.
AppendLine "EnableMisusedWordsDictionary = " & oOptions.EnableMisusedWordsDictionary

' True if users are notified when additional proofing tools are available for download.
AppendLine "EnableProofingToolsAdvertisement = " & oOptions.EnableProofingToolsAdvertisement

' True if Word makes the computer respond with a sound whenever an error occurs.
AppendLine "EnableSound = " & oOptions.EnableSound

' True if the current printer has a special feeder for envelopes. Read-only.
AppendLine "EnvelopeFeederInstalled = " & oOptions.EnvelopeFeederInstalled

' True to expand all headings in the document when the document opens.
AppendLine "ExpandHeadingsOnOpen = " & oOptions.ExpandHeadingsOnOpen

' True for Word to keep track of all formatting in a document.
AppendLine "FormatScanning = " & oOptions.FormatScanning

' Integer for which spelling dictionary to use for regions of text with language formatting set to French.  WdFrenchSpeller.
' AppendLine "FrenchReform = " & oOptions.FrenchReform

' Float for the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents.
' AppendLine "GridDistanceHorizontal = " & oOptions.GridDistanceHorizontal

' Float for the amount of vertical space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents.
' AppendLine "GridDistanceVertical = " & oOptions.GridDistanceVertical

' Float for the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in new documents.
' AppendLine "GridOriginHorizontal = " & oOptions.GridOriginHorizontal

' Float for the the point, relative to the top of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in new documents.
' AppendLine "GridOriginVertical = " & oOptions.GridOriginVertical

' True if Word automatically converts a word with only one suggestion during conversion between Hangul and Hanja.
' AppendLine "HangulHanjaFastConversion = " & oOptions.HangulHanjaFastConversion

' Integer for the mode of the Hebrew spell checker. WdHebSpellStart.
' AppendLine "HebrewMode = " & oOptions.HebrewMode

' True if Word adds bidirectional control characters when saving a document as a text file.
' AppendLine "AddBiDirectionalMarksWhenSavingTextFile = " & oOptions.AddBiDirectionalMarksWhenSavingTextFile

' True if file name extensions, MS-DOS paths, email addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are ignored while checking spelling.
AppendLine "IgnoreInternetAndFileAddresses = " & oOptions.IgnoreInternetAndFileAddresses

' True if words that contain numbers are ignored while checking spelling.
AppendLine "IgnoreMixedDigits = " & oOptions.IgnoreMixedDigits

' True if words in all uppercase letters are ignored while checking spelling.
AppendLine "IgnoreUppercase = " & oOptions.IgnoreUppercase

' True if Word is set to automatically open and close the Japanese Input Method Editor (IME).
' AppendLine "IMEAutomaticControl = " & oOptions.IMEAutomaticControl

' True if Word displays an unconfirmed character string in the Japanese Input Method Editor (IME) as an insertion between existing (confirmed) character strings.
AppendLine "InlineConversion = " & oOptions.InlineConversion

' Integer for the color of an inserted table cell.  WdCellColor.
AppendLine "InsertedCellColor = " & oOptions.InsertedCellColor

' Integer for the color of text that is inserted while change tracking is enabled. WdColorIndex.
AppendLine "InsertedTextColor = " & oOptions.InsertedTextColor

' Integer for how Word formats inserted text while change tracking is enabled (the TrackRevisions property is True ). WdInsertedTextMark.
AppendLine "InsertedTextMark = " & oOptions.InsertedTextMark

' True if the INS key can be used for switching Overtype on and off.
AppendLine "INSKeyForOvertype = " & oOptions.INSKeyForOvertype

' True if the INS key can be used for pasting the Clipboard contents.
AppendLine "INSKeyForPaste = " & oOptions.INSKeyForPaste

' Integer for the high-ANSI text interpretation behavior. WdHighAnsiText.
AppendLine "InterpretHighAnsi = " & oOptions.InterpretHighAnsi

' True if Word creates a local copy of a file on the user's computer when editing a file stored on a network server.
AppendLine "LocalNetworkFile = " & oOptions.LocalNetworkFile

' True if documents formatted for another country's/region's standard paper size (for example, A4) are automatically adjusted so that they're printed correctly on your country's/region's standard paper size (for example, Letter).
AppendLine "MapPaperSize = " & oOptions.MapPaperSize

' True if margin alignment guides are displayed in the user interface.
' AppendLine "MarginAlignmentGuides = " & oOptions.MarginAlignmentGuides

' True if Word ignores the distinction between " ![Symbol](./images/fe289_za06051768.gif)" and " ![Symbol](./images/fe241_za06051721.gif)" following ![Symbol](./images/fe144_za06051649.gif)-row and ![Symbol](./images/fe209_za06051695.gif)-row characters during a search.
' AppendLine "MatchFuzzyAY = " & oOptions.MatchFuzzyAY

' True if Word ignores the distinction between full-width and half-width characters (Latin or Japanese) during a search.
' AppendLine "MatchFuzzyByte = " & oOptions.MatchFuzzyByte

' True if Word ignores the distinction between uppercase and lowercase letters during a search.
' AppendLine "MatchFuzzyCase = " & oOptions.MatchFuzzyCase

' True if Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search.
' AppendLine "MatchFuzzyDash = " & oOptions.MatchFuzzyDash

' True if Word ignores the distinction between " ![Symbol](./images/fe274_za06051753.gif)" and " ![Symbol](./images/fe275_za06051754.gif)" and between " ![Symbol](./images/fe276_za06051755.gif)" and " ![Symbol](./images/fe277_za06051756.gif)" during a search.
' AppendLine "MatchFuzzyDZ = " & oOptions.MatchFuzzyDZ

' True if Word ignores the distinction between hiragana and katakana during a search.
' AppendLine "MatchFuzzyHiragana = " & oOptions.MatchFuzzyHiragana

' True if Word ignores the distinction between types of repetition marks during a search.
' AppendLine "MatchFuzzyIterationMark = " & oOptions.MatchFuzzyIterationMark

' True if Word ignores the distinction between standard and nonstandard kanji ideography during a search.
' AppendLine "MatchFuzzyKanji = " & oOptions.MatchFuzzyKanji

' True if Word ignores the distinction between " ![Symbol](./images/fe107_za06051631.gif)" and " ![Symbol](./images/fe112_za06051635.gif)" before ![Symbol](./images/fe290_za06051769.gif)-row characters during a search.
' AppendLine "MatchFuzzyKiKu = " & oOptions.MatchFuzzyKiKu

' True if Word ignores the distinction between new kana and old kana characters during a search.
' AppendLine "MatchFuzzyOldKana = " & oOptions.MatchFuzzyOldKana

' True if Word ignores the distinction between short and long vowel sounds during a search.
' AppendLine "MatchFuzzyProlongedSoundMark = " & oOptions.MatchFuzzyProlongedSoundMark

' True if Word ignores the distinction between types of punctuation marks during a search.
' AppendLine "MatchFuzzyPunctuation = " & oOptions.MatchFuzzyPunctuation

' True if Word ignores the distinction between diphthongs and double consonants during a search.
' AppendLine "MatchFuzzySmallKana = " & oOptions.MatchFuzzySmallKana

' True if Word ignores the distinction between space markers used during a search.
' AppendLine "MatchFuzzySpace = " & oOptions.MatchFuzzySpace

' Integer for the standard measurement unit for Word. WdMeasurementUnits.
AppendLine "MeasurementUnit = " & oOptions.MeasurementUnit

' Integer for the color of merged table cells.  WdCellColor.
AppendLine "MergedCellColor = " & oOptions.MergedCellColor

' Integer for the direction of conversion between Hangul and Hanja. WdMonthNames.
AppendLine "MonthNames = " & oOptions.MonthNames

' Integer for the color of moved text.  WdColorIndex.
AppendLine "MoveFromTextColor = " & oOptions.MoveFromTextColor

' Integer for the type of revision mark to use for moved text. WdMoveFromTextMark.
AppendLine "MoveFromTextMark = " & oOptions.MoveFromTextMark

' Integer for the color of moved text. WdColorIndex.
AppendLine "MoveToTextColor = " & oOptions.MoveToTextColor

' Integer for the type of revision mark to use for moved text. WdMoveToTextMark.
AppendLine "MoveToTextMark = " & oOptions.MoveToTextMark

' Integer for the direction of conversion between Hangul and Hanja. WdMultipleWordConversionsMode.
AppendLine "MultipleWordConversionsMode = " & oOptions.MultipleWordConversionsMode

' True if Word automatically converts equations to professional format.
AppendLine "OMathAutoBuildUp = " & oOptions.OMathAutoBuildUp

' Boolean for how equations are represented in plain text. True if Linear Format. False if MathML.
' AppendLine "OMathCopyLF = " & oOptions.OMathCopyLF

' True if Word optimizes all new documents for viewing in Word 97 by disabling any incompatible formatting.
' AppendLine "OptimizeForWord97byDefault = " & oOptions.OptimizeForWord97byDefault

' True if Overtype mode is active.
AppendLine "Overtype = " & oOptions.Overtype

' True if page alignment guides are displayed in the user interface.
' AppendLine "PageAlignmentGuides = " & oOptions.PageAlignmentGuides

' True if Word repaginates documents in the background.
AppendLine "Pagination = " & oOptions.Pagination

' True if paragraph alignment guides are displayed in the user interface.
' AppendLine "ParagraphAlignmentGuides = " & oOptions.ParagraphAlignmentGuides

' True if Word automatically adjusts the spacing of paragraphs when cutting and pasting selections.
AppendLine "PasteAdjustParagraphSpacing = " & oOptions.PasteAdjustParagraphSpacing

' True if Word automatically adjusts the formatting of tables when cutting and pasting selections.
AppendLine "PasteAdjustTableFormatting = " & oOptions.PasteAdjustTableFormatting

' True if Word automatically adjusts the spacing of words when cutting and pasting selections.
AppendLine "PasteAdjustWordSpacing = " & oOptions.PasteAdjustWordSpacing

' Integer for how text is pasted when text is copied from another Microsoft Office Word document.  WdPasteOptions.
AppendLine "PasteFormatBetweenDocuments = " & oOptions.PasteFormatBetweenDocuments

' Integer for how text is pasted when text is copied from a document that uses styles. WdPasteOptions.
AppendLine "PasteFormatBetweenStyledDocuments = " & oOptions.PasteFormatBetweenStyledDocuments

' Integer for how text is pasted when text is copied from an external source, such as a webpage. WdPasteOptions.
AppendLine "PasteFormatFromExternalSource = " & oOptions.PasteFormatFromExternalSource

' Integer for how text is pasted when text is copied or cut and then pasted in the same document. WdPasteOptions.
AppendLine "PasteFormatWithinDocument = " & oOptions.PasteFormatWithinDocument

' True to merge text formatting when pasting from Microsoft PowerPoint.
AppendLine "PasteMergeFromPPT = " & oOptions.PasteMergeFromPPT

' True to merge table formatting when pasting from Microsoft Excel.
AppendLine "PasteMergeFromXL = " & oOptions.PasteMergeFromXL

' True to merge the formatting of pasted lists with surrounding lists.
AppendLine "PasteMergeLists = " & oOptions.PasteMergeLists

' True to keep bullets and numbering when selecting Keep text only from the Paste Options context menu.
AppendLine "PasteOptionKeepBulletsAndNumbers = " & oOptions.PasteOptionKeepBulletsAndNumbers

' True if Word intelligently pastes selections into a document.
AppendLine "PasteSmartCutPaste = " & oOptions.PasteSmartCutPaste

' True if Word intelligently merges styles when pasting a selection from a different document.
AppendLine "PasteSmartStyleBehavior = " & oOptions.PasteSmartStyleBehavior

' String for the name of the application to use to edit pictures.
AppendLine "PictureEditor = " & oOptions.PictureEditor

' Integer for how Word wraps text around pictures.  WdWrapTypeMerged  .
AppendLine "PictureWrapType = " & oOptions.PictureWrapType

' Integer for the mode of the European Portuguese spell checker. WdPortugueseReform.
' AppendLine "PortugalReform = " & oOptions.PortugalReform

' True if Word optimizes character positioning for print layout rather than on-screen readability. True disables the default setting that compresses character spacing to facilitate on-screen readability and enables character spacing for print media.
AppendLine "PrecisePositioning = " & oOptions.PrecisePositioning

' True to save new documents in web locations by default.
AppendLine "PreferCloudSaveLocations = " & oOptions.PreferCloudSaveLocations

' True if Word prints in the background.
AppendLine "PrintBackground = " & oOptions.PrintBackground

' True if background colors and images are printed when a document is printed.
AppendLine "PrintBackgrounds = " & oOptions.PrintBackgrounds

' True if Word prints comments, starting on a new page at the end of the document.
AppendLine "PrintComments = " & oOptions.PrintComments

' True if Word prints using minimal formatting.
AppendLine "PrintDraft = " & oOptions.PrintDraft

' True if Word prints drawing objects.
AppendLine "PrintDrawingObjects = " & oOptions.PrintDrawingObjects

' True if Word prints even pages in ascending order during manual duplex printing.
AppendLine "PrintEvenPagesInAscendingOrder = " & oOptions.PrintEvenPagesInAscendingOrder

' True if Word prints field codes instead of field results.
AppendLine "PrintFieldCodes = " & oOptions.PrintFieldCodes

' True if hidden text is printed.
AppendLine "PrintHiddenText = " & oOptions.PrintHiddenText

' True if Word prints odd pages in ascending order during manual duplex printing.
AppendLine "PrintOddPagesInAscendingOrder = " & oOptions.PrintOddPagesInAscendingOrder

' True if Word prints document summary information on a separate page at the end of the document.
AppendLine "PrintProperties = " & oOptions.PrintProperties

' True if Word prints pages in reverse order.
AppendLine "PrintReverse = " & oOptions.PrintReverse

' True to print the XML tags when printing a document.
AppendLine "PrintXMLTag = " & oOptions.PrintXMLTag

' True displays a message asking the user to verify whether they want to reformat a style or reapply the original style formatting when changing the formatting of styles.
AppendLine "PromptUpdateStyle = " & oOptions.PromptUpdateStyle

' True to mark words that are repeated when spelling is checked.
AppendLine "RepeatWord = " & oOptions.RepeatWord

' True if the result of typing or pasting replaces the selection.
AppendLine "ReplaceSelection = " & oOptions.ReplaceSelection

' Integer for the color of changed lines in a document with tracked changes. WdColorIndex.
AppendLine "RevisedLinesColor = " & oOptions.RevisedLinesColor

' Integer for the placement of changed lines in a document with tracked changes. WdRevisedLinesMark.
AppendLine "RevisedLinesMark = " & oOptions.RevisedLinesMark

' Integer for the color used to mark formatting changes while change tracking is enabled. WdColorIndex.
AppendLine "RevisedPropertiesColor = " & oOptions.RevisedPropertiesColor

' Integer for the mark used to show formatting changes while change tracking is enabled. WdRevisedPropertiesMark.
AppendLine "RevisedPropertiesMark = " & oOptions.RevisedPropertiesMark

' Integer for the direction of revision and comment balloons when they are printed.  WdRevisionsBalloonPrintOrientation  .
AppendLine "RevisionsBalloonPrintOrientation = " & oOptions.RevisionsBalloonPrintOrientation

' Integer for the time interval in minutes for saving AutoRecover information.
AppendLine "SaveInterval = " & oOptions.SaveInterval

' True if Word prompts the user for confirmation to save changes to the Normal template before it closes.
AppendLine "SaveNormalPrompt = " & oOptions.SaveNormalPrompt

' True if Word prompts for document property information when saving a new document.
AppendLine "SavePropertiesPrompt = " & oOptions.SavePropertiesPrompt

' True if the Send To command on the File menu inserts the active document as an attachment to a mail message.
AppendLine "SendMailAttach = " & oOptions.SendMailAttach

' True to check the sequence of independent characters for South Asian text.
AppendLine "SequenceCheck = " & oOptions.SequenceCheck

' True if bidirectional control characters are visible in the current document.
AppendLine "ShowControlCharacters = " & oOptions.ShowControlCharacters

' True if the Developer tab is displayed in the ribbon.
AppendLine "ShowDevTools = " & oOptions.ShowDevTools

' True if diacritics are visible in a right-to-left language document.
AppendLine "ShowDiacritics = " & oOptions.ShowDiacritics

' True for Word to mark inconsistencies in formatting by placing a squiggly underline beneath text formatted similarly to other formatting that is used more frequently in a document.
AppendLine "ShowFormatError = " & oOptions.ShowFormatError

' True if Word displays hidden markup when opening or saving a file.
AppendLine "ShowMarkupOpenSave = " & oOptions.ShowMarkupOpenSave

' True to display mini toolbars when the user right-clicks in the document window.
AppendLine "ShowMenuFloaties = " & oOptions.ShowMenuFloaties

' True if Word displays a list of summary statistics, including measures of readability, when it has finished checking grammar.
AppendLine "ShowReadabilityStatistics = " & oOptions.ShowReadabilityStatistics

' True if mini toolbars display when a user selects text.
AppendLine "ShowSelectionFloaties = " & oOptions.ShowSelectionFloaties

' True if smart cursoring is enabled.
AppendLine "SmartCursoring = " & oOptions.SmartCursoring

' True if Word automatically adjusts the spacing between words and punctuation when cutting and pasting occurs.
AppendLine "SmartCutPaste = " & oOptions.SmartCutPaste

' True for Word to include the paragraph mark in a selection when selecting most or all of a paragraph.
AppendLine "SmartParaSelection = " & oOptions.SmartParaSelection

' True if AutoShapes or East Asian characters are automatically aligned with an invisible grid when they are drawn, moved, or resized.
AppendLine "SnapToGrid = " & oOptions.SnapToGrid

' True if Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters.
' AppendLine "SnapToShapes = " & oOptions.SnapToShapes

' Integer for the mode for the Spanish spell checker. WdSpanishSpeller.
' AppendLine "SpanishMode = " & oOptions.SpanishMode

' Integer for the color for split table cells. WdCellColor  .
AppendLine "SplitCellColor = " & oOptions.SplitCellColor

' True for Word to assign a random number to changes in a document, each time a document is saved, to facilitate comparing and merging documents.
AppendLine "StoreRSIDOnSave = " & oOptions.StoreRSIDOnSave

' True if the spell checker uses spelling rules regarding Arabic words ending with the letter yaa.
' AppendLine "StrictFinalYaa = " & oOptions.StrictFinalYaa

' True if the spell checker uses spelling rules regarding Arabic words beginning with an alef hamza.
' AppendLine "StrictInitialAlefHamza = " & oOptions.StrictInitialAlefHamza

' True if the spell checker uses spelling rules regarding Russian words that use the strict  character.
' AppendLine "StrictRussianE = " & oOptions.StrictRussianE

' True if the spell checker uses spelling rules to flag Arabic words ending with haa instead of taa marboota.
' AppendLine "StrictTaaMarboota = " & oOptions.StrictTaaMarboota

' True if Word draws spelling suggestions from the main dictionary only.
AppendLine "SuggestFromMainDictionaryOnly = " & oOptions.SuggestFromMainDictionaryOnly

' True if Word always suggests alternative spellings for each misspelled word when checking spelling.
AppendLine "SuggestSpellingCorrections = " & oOptions.SuggestSpellingCorrections

' True if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs.
AppendLine "TabIndentKey = " & oOptions.TabIndentKey

' True for Word to replace illegal South Asian characters.
AppendLine "TypeNReplace = " & oOptions.TypeNReplace

' True if Word updates fields automatically before printing a document.
AppendLine "UpdateFieldsAtPrint = " & oOptions.UpdateFieldsAtPrint

' True if Word allows fields containing tracked changes to update before printing.
AppendLine "UpdateFieldsWithTrackedChangesAtPrint = " & oOptions.UpdateFieldsWithTrackedChangesAtPrint

' True if Word automatically updates all embedded OLE links in a document when it is opened.
AppendLine "UpdateLinksAtOpen = " & oOptions.UpdateLinksAtOpen

' True if Word updates embedded links to other files before printing a document.
AppendLine "UpdateLinksAtPrint = " & oOptions.UpdateLinksAtPrint

' Integer for the behavior Word should take when updating a style to match a selection that contains numbering or bullets.  WdUpdateStyleListBehavior.
AppendLine "UpdateStyleListBehavior = " & oOptions.UpdateStyleListBehavior

' True if Word uses characters as the default measurement unit for the current document.
AppendLine "UseCharacterUnit = " & oOptions.UseCharacterUnit

' True if you can set the color of diacritics in the current document.
AppendLine "UseDiffDiacColor = " & oOptions.UseDiffDiacColor

' True if Word uses the German post-reform spelling rules when checking spelling.
' AppendLine "UseGermanSpellingReform = " & oOptions.UseGermanSpellingReform

' True if Word identifies the document author based upon the User name and Initials settings on the General tab of the Options dialog box, and False if Word identifies the author based on the account information with which the user signed in to Office.
AppendLine "UseLocalUserInfo = " & oOptions.UseLocalUserInfo

' True if Word uses the Normal style for bullets and numbering.
AppendLine "UseNormalStyleForList = " & oOptions.UseNormalStyleForList

' True if sub-pixel positioning is enabled.
AppendLine "UseSubPixelPositioning = " & oOptions.UseSubPixelPositioning

' Integer for the selection behavior based on visual cursor movement in a right-to-left language document. WdVisualSelection.
AppendLine "VisualSelection = " & oOptions.VisualSelection

' True for Word to display a warning when saving, printing, or sending as email a document containing comments or tracked changes.
AppendLine "WarnBeforeSavingPrintingSendingMarkup = " & oOptions.WarnBeforeSavingPrintingSendingMarkup

AppendBlank
AppendLine "[Other]"
AppendLine "ActivePrinter = " & oApp.ActivePrinter
AppendLine "BibliographyStyle = " & oApp.Bibliography.BibliographyStyle
AppendLine "CheckLanguage = " & CBool(oApp.CheckLanguage)
AppendLine "DefaultTableSeparator = " & oApp.DefaultTableSeparator
AppendLine "FileValidation = " & CInt(oApp.FileValidation)
AppendLine "OpenAttachmentsInFullScreen = " & CBool(oApp.OpenAttachmentsInFullScreen)
AppendLine "PrintPreview = " & CBool(oApp.PrintPreview)
AppendLine "RestrictLinkedStyles = " & CBool(oApp.RestrictLinkedStyles)
' requires a document window open
oApp.Visible = True
oApp.CommandBars("Task Pane").Visible = False
Set oDoc = oApp.Documents.Add
AppendLine "ShowStartupDialog = " & CBool(oApp.ShowStartupDialog)
oDoc.Close wdDoNotSaveChanges
AppendLine "UserAddress = " & oApp.UserAddress
AppendLine "UserInitials = " & oApp.UserInitials
AppendLine "UserName = " & oApp.UserName
AppendLine "DefaultTheme = " & oApp.GetDefaultTheme(wdDocument)

ShowAddIns(oApp)
ShowAutoCaptions(oApp)
ShowAutoCorrect(oApp)
ShowAutoCorrectEmail(oApp)
ShowCaptionLabels(oApp)
ShowComAddIns(oApp)
ShowCommandBars(oApp)
ShowCustomDictionaries(oApp)
ShowFileConverters(oApp)
ShowFontNames(oApp)
ShowRecentFiles(oApp)
' ShowTaskPanes(oApp)
ShowTemplates(oApp)
ShowMisc(oApp)

On Error Resume Next
' oDoc.Close wdDoNotSaveChanges
On Error GoTo 0
Set oDocs = oApp.Documents
oApp.PrintPreview = False
if oDocs.Count > 0 Then oDocs.Close wdDoNotSaveChanges
If Not oApp.NormalTemplate.Saved Then oApp.NormalTemplate.Save
oApp.Quit

If bResetNormalTemplate Then ResetNormalTemplate
StringToFile sHomerText, sTargetIni

If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
End If

echo "Done"
