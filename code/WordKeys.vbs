Option Explicit
WScript.Echo"Starting WordKeys"

Dim aKeys, aIni
Dim bBackupDocx, bLogActions, bValue, bAddToRecentFiles, bConfirmConversions, bReadOnly
Dim dCommands, dKeyBindings, dCategories, dKeys, dOther, dOptions, d, dStyle, dIni, dSourceIni
Dim iCommandCount, iKeyCode2, iKeyCode, iRow, iRowCount, iKeyCategory, iKeyBinding, iKey, iKeyBindingCount, iValue, iArgCount, iCount
Dim oTableDoc, oTables, oTable, oRows, oRow, oCells, oCell, oAppKeyBinding, oAppKeyBindings, oNormalKeyBinding, oNormalKeyBindings, oKeyBinding, oKeyBindings, oKey, oSystem, oFile, oOption, oOptions, oFormat, oFont, oStyle, oApp, oData, oDoc, oDocs, oFind, oProperty, oRange, oToc
Dim sKeyString, sLine, sKeyCategory, sKeys, sCategory, sCommand, sKey, sTargetLog, sScriptVbs, sHomerLibVbs, sDir, sValue, sAttrib, s, sStyle, sName, sSourceDocx, sMatch, sTargetIni, sText, sConfigFile, sSourceIni

' Word constant enumerations used within Options
Const wdDocument = 1
Const wdEmailMessage = 0
Const wdWebPage = 2

' wdKeyCategory
Const wdKeyCategoryAutoText = 4
Const wdKeyCategoryCommand = 1
Const wdKeyCategoryDisable = 0
Const wdKeyCategoryFont = 3
Const wdKeyCategoryMacro = 2
Const wdKeyCategoryNil = -1
Const wdKeyCategoryPrefix = 7
Const wdKeyCategoryStyle = 5
Const wdKeyCategorySymbol = 6

' wdKey
Const wdKey0 = 48
Const wdKey1 = 49
Const wdKey2 = 50
Const wdKey3 = 51
Const wdKey4 = 52
Const wdKey5 = 53
Const wdKey6 = 54
Const wdKey7 = 55
Const wdKey8 = 56
Const wdKey9 = 57
Const wdKeyA = 65
Const wdKeyAlt = 1024
Const wdKeyB = 66
Const wdKeyBackSingleQuote = 192
Const wdKeyBackSlash = 220
Const wdKeyBackspace = 8
Const wdKeyC = 67
Const wdKeyCloseSquareBrace = 221
Const wdKeyComma = 188
Const wdKeyCommand = 512
Const wdKeyControl = 512
Const wdKeyD = 68
Const wdKeyDelete = 46
Const wdKeyE = 69
Const wdKeyEnd = 35
Const wdKeyEquals = 187
Const wdKeyEsc = 27
Const wdKeyEscape = 27
Const wdKeyF = 70
Const wdKeyF1 = 112
Const wdKeyF10 = 121
Const wdKeyF11 = 122
Const wdKeyF12 = 123
Const wdKeyF13 = 124
Const wdKeyF14 = 125
Const wdKeyF15 = 126
Const wdKeyF16 = 127
Const wdKeyF2 = 113
Const wdKeyF3 = 114
Const wdKeyF4 = 115
Const wdKeyF5 = 116
Const wdKeyF6 = 117
Const wdKeyF7 = 118
Const wdKeyF8 = 119
Const wdKeyF9 = 120
Const wdKeyG = 71
Const wdKeyH = 72
Const wdKeyHome = 36
Const wdKeyHyphen = 189
Const wdKeyI = 73
Const wdKeyInsert = 45
Const wdKeyJ = 74
Const wdKeyK = 75
Const wdKeyL = 76
Const wdKeyM = 77
Const wdKeyN = 78
Const wdKeyNumeric0 = 96
Const wdKeyNumeric1 = 97
Const wdKeyNumeric2 = 98
Const wdKeyNumeric3 = 99
Const wdKeyNumeric4 = 100
Const wdKeyNumeric5 = 101
Const wdKeyNumeric5Special = 12
Const wdKeyNumeric6 = 102
Const wdKeyNumeric7 = 103
Const wdKeyNumeric8 = 104
Const wdKeyNumeric9 = 105
Const wdKeyNumericAdd = 107
Const wdKeyNumericDecimal = 110
Const wdKeyNumericDivide = 111
Const wdKeyNumericMultiply = 106
Const wdKeyNumericSubtract = 109
Const wdKeyO = 79
Const wdKeyOpenSquareBrace = 219
Const wdKeyOption = 1024
Const wdKeyP = 80
Const wdKeyPageDown = 34
Const wdKeyPageUp = 33
Const wdKeyPause = 19
Const wdKeyPeriod = 190
Const wdKeyQ = 81
Const wdKeyR = 82
Const wdKeyReturn = 13
Const wdKeyS = 83
Const wdKeyScrollLock = 145
Const wdKeySemiColon = 186
Const wdKeyShift = 256
Const wdKeySingleQuote = 222
Const wdKeySlash = 191
Const wdKeySpacebar = 32
Const wdKeyT = 84
Const wdKeyTab = 9
Const wdKeyU = 85
Const wdKeyV = 86
Const wdKeyW = 87
Const wdKeyX = 88
Const wdKeyY = 89
Const wdKeyZ = 90
Const wdNoKey = 255

' Additions
Const wdKeyDown = 40
Const wdKeyUp = 38
Const wdKeyLeft = 37
Const wdKeyRight = 39

'wd enumerations
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

Function CreateCategoriesDictionary
Set dCategories = CreateDictionary
dCategories.Add "AutoText", 4
dCategories.Add "Command", 1
dCategories.Add "Disable", 0
dCategories.Add "Font", 3
dCategories.Add "Macro", 2
dCategories.Add "Nil", -1
dCategories.Add "Prefix", 7
dCategories.Add "Style", 5
dCategories.Add "Symbol", 6

Set CreateCategoriesDictionary = dCategories
End Function

Function CreateKeysDictionary()
Dim dKeys

Set dKeys = CreateDictionary
dKeys.Add "0", 48
dKeys.Add "1", 49
dKeys.Add "2", 50
dKeys.Add "3", 51
dKeys.Add "4", 52
dKeys.Add "5", 53
dKeys.Add "6", 54
dKeys.Add "7", 55
dKeys.Add "8", 56
dKeys.Add "9", 57
dKeys.Add "A", 65
dKeys.Add "Alt", 1024
dKeys.Add "B", 66
dKeys.Add "BackSingleQuote", 192
dKeys.Add "`", 192
dKeys.Add "BackSlash", 220
dKeys.Add "\", 220
dKeys.Add "Backspace", 8
dKeys.Add "C", 67
dKeys.Add "CloseSquareBrace", 221
dKeys.Add "]", 221
dKeys.Add "Comma", 188
dKeys.Add ",", 188
dKeys.Add "Command", 512
dKeys.Add "Control", 512
dKeys.Add "Ctrl", 512
dKeys.Add "D", 68
dKeys.Add "Delete", 46
dKeys.Add "E", 69
dKeys.Add "End", 35
dKeys.Add "Equals", 187
dKeys.Add "Esc", 27
dKeys.Add "Escape", 27
dKeys.Add "F", 70
dKeys.Add "F1", 112
dKeys.Add "F10", 121
dKeys.Add "F11", 122
dKeys.Add "F12", 123
dKeys.Add "F13", 124
dKeys.Add "F14", 125
dKeys.Add "F15", 126
dKeys.Add "F16", 127
dKeys.Add "F2", 113
dKeys.Add "F3", 114
dKeys.Add "F4", 115
dKeys.Add "F5", 116
dKeys.Add "F6", 117
dKeys.Add "F7", 118
dKeys.Add "F8", 119
dKeys.Add "F9", 120
dKeys.Add "G", 71
dKeys.Add "H", 72
dKeys.Add "Home", 36
dKeys.Add "Hyphen", 189
dKeys.Add "-", 189
dKeys.Add "Dash", 189
dKeys.Add "I", 73
dKeys.Add "Insert", 45
dKeys.Add "J", 74
dKeys.Add "K", 75
dKeys.Add "L", 76
dKeys.Add "M", 77
dKeys.Add "N", 78
dKeys.Add "Numeric0", 96
dKeys.Add "NumPad0", 96
dKeys.Add "Numeric1", 97
dKeys.Add "NumPad1", 97
dKeys.Add "Numeric2", 98
dKeys.Add "NumPad2", 98
dKeys.Add "Numeric3", 99
dKeys.Add "NumPad3", 99
dKeys.Add "Numeric4", 100
dKeys.Add "NumPad4", 100
dKeys.Add "Numeric5", 101
dKeys.Add "NumPad5", 101
dKeys.Add "Numeric5Special", 12
dKeys.Add "NumPadCenter", 12
dKeys.Add "Numeric6", 102
dKeys.Add "NumPad6", 102
dKeys.Add "Numeric7", 103
dKeys.Add "NumPad7", 103
dKeys.Add "Numeric8", 104
dKeys.Add "NumPad8", 104
dKeys.Add "Numeric9", 105
dKeys.Add "NumPad9", 105
dKeys.Add "NumericAdd", 107
dKeys.Add "NumPadPlus", 107
dKeys.Add "NumericDecimal", 110
dKeys.Add "NumPad.", 110
dKeys.Add "NumericDivide", 111
dKeys.Add "NumPad/", 111
dKeys.Add "NumericMultiply", 106
dKeys.Add "NumPad*", 106
dKeys.Add "NumericSubtract", 109
dKeys.Add "NumPad-", 109
dKeys.Add "O", 79
dKeys.Add "OpenSquareBrace", 219
dKeys.Add "[", 219
dKeys.Add "Option", 1024
dKeys.Add "P", 80
dKeys.Add "PageDown", 34
dKeys.Add "PageUp", 33
dKeys.Add "Pause", 19
dKeys.Add "Period", 190
dKeys.Add ".", 190
dKeys.Add "Q", 81
dKeys.Add "R", 82
dKeys.Add "Return", 13
dKeys.Add "Enter", 13
dKeys.Add "S", 83
dKeys.Add "ScrollLock", 145
dKeys.Add "SemiColon", 186
dKeys.Add ";", 186
dKeys.Add "Shift", 256
dKeys.Add "SingleQuote", 222
dKeys.Add "'", 222
dKeys.Add "Slash", 191
dKeys.Add "/", 191
dKeys.Add "Spacebar", 32
dKeys.Add "Space", 32
dKeys.Add "T", 84
dKeys.Add "Tab", 9
dKeys.Add "U", 85
dKeys.Add "V", 86
dKeys.Add "W", 87
dKeys.Add "X", 88
dKeys.Add "Y", 89
dKeys.Add "Z", 90
dKeys.Add "NoKey", 255

' Additions
dKeys.Add "DownArrow", 40
dKeys.Add "UpArrow", 38
dKeys.Add "LeftArrow", 37
dKeys.Add "RightArrow", 39

Set CreateKeysDictionary = dKeys
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

Function ShowOther(oApp)
AppendBlank
AppendLine "[;Other]"
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

sSourceDocx = ""
sSourceIni = ""

If iArgCount > 1 Then
sSourceDocx = WScript.Arguments(0)
sSourceIni = WScript.Arguments(1)
sSourceIni = GetIniFile(sSourceIni)
ElseIf iArgCount = 1 Then
If PathGetExtension(WScript.Arguments(0)) = "docx" Then
sSourceDocx = WScript.Arguments(0)
Else
sSourceIni = WScript.Arguments(0)
sSourceIni = GetIniFile(sSourceIni)
End If
End If

If Len(sSourceDocx) > 0 Then
If InStr(sSourceDocx, "\") = 0 Then sSourceDocx = PathCombine(PathGetCurrentDirectory(), sSourceDocx)
If not FileExists(sSourceDocx) Then Quit "Cannot find " & sSourceDocx
End If

If Len(sSourceIni) > 0 Then
If InStr(sSourceIni, "\") = 0 Then sSourceIni = PathCombine(PathGetCurrentDirectory(), sSourceIni)
If not FileExists(sSourceIni) Then Quit "Cannot find " & sSourceIni
Set dSourceIni = IniToDictionary(sSourceIni)
bLogActions = GetGlobalValue(dSourceIni, "LogActions", True)
bReadOnly = False
rem print "Global exists=" & dSourceIni.Exists("Global")
rem if dSourceIni.Exists("Global") Then dSourceIni.Remove("Global")
ProcessTerminateAllModule "WinWord"
Else
Set dSourceIni = CreateDictionary()
bReadOnly = True
End If

sTargetIni = PathCombine(PathGetCurrentDirectory(), "KEYS.ini")
sTargetLog = PathCombine(PathGetCurrentDirectory(), "WordKeys.log")

Set oApp = CreateObject("Word.Application")
oApp.Visible = False
oApp.Visible = True
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False

Set oDocs = oApp.Documents
bAddToRecentFiles = False
bConfirmConversions = False
If Len(sSourceDocx) > 0 Then Print "Opening " & PathGetName(sSourceDocx)
If Len(sSourceDocx) > 0 Then Set oDoc = oDocs.Open(sSourceDocx, bAddToRecentFiles, bReadOnly, bConfirmConversions)

Set dCategories = CreateCategoriesDictionary
Set dKeys = CreateKeysDictionary
If Not bReadOnly Then
s = ""
If PathGetFolder(sSourceIni) <> PathGetCurrentDirectory() Then s = "global "
print "Applying " & s & PathGetName(sSourceIni)
If Len(sSourceDocx) = 0 Then
oApp.CustomizationContext = oApp.NormalTemplate
Else
oApp.CustomizationContext = oDoc
End If
Set oNormalKeyBindings = oApp.KeyBindings
For Each sKeyCategory in dCategories
If dSourceIni.Exists(sKeyCategory) Then
iKeyCategory = dCategories(sKeyCategory)
print sKeyCategory & " category"
iCommandCount = dSourceIni(sKeyCategory).Count
If iCommandCount = 0 Then
For Each oKeyBinding in oApp.KeyBindings
If oKeyBinding.KeyCategory = iKeyCategory Then
print "Clearing " & oKeyBinding.KeyString & " from " & oKeyBinding.Command
oKeyBinding.Clear
End If 'oKeyBinding.KeyCategory = iKeyCategory
Next ' oKeyBinding in oApp.KeyBindings
End If ' iCommandCount = 0
For Each sCommand in dSourceIni(sKeyCategory).Keys
' print "sCommand " & sCommand
sKeys = dSourceIni(sKeyCategory)(sCommand)
sKeys = Replace(sKeys, " ", "")
aKeys = Split(sKeys, "+")
iKeyCode = 0
' print "sKeys = " & sKeys
For Each sKey in aKeys
' print "sKey = " & sKey
iKey = dKeys(sKey)
iKeyCode = iKeyCode + iKey
Next ' sKey in aKeys
' iKeyCode = oApp.BuildKeyCode(dKeys(aKeys(0)), dKeys(aKeys(1)), dKeys(aKeys(2)))
Set oKeyBinding = oApp.FindKey(iKeyCode)
' sCommand = Replace(sCommand, " ", "")
If oKeyBinding Is Nothing Then
print "Adding " & sKey & " = " & sCommand
oNormalKeyBindings.Add iKeyCategory, sCommand, iKeyCode
ElseIf Len(sKeys) = 0 Then
print "Clearing " & oKeyBinding.KeyString & " from " & oKeyBinding.Command
oKeyBinding.Clear
ElseIf oKeyBinding.Command = sCommand and oKeyBinding.KeyCategory = iKeyCategory Then
' Do nothing
Else
print "Changing " & oKeyBinding.KeyString & " from " & oKeyBinding.Command & " to " & sCommand
' oKeyBinding.Clear
' oKeyBinding.Disable
' print oKeyBinding.Protected
' print oKeyBinding.Context.Name
Err.Clear
On Error Resume Next
oKeyBinding.Rebind iKeyCategory, sCommand
If Len(Err.Description) > 0 Then
sCommand = Replace(sCommand, " ", "")
oKeyBinding.Rebind iKeyCategory, sCommand
If Len(Err.Description) > 0 Then print Err.Description
End If
On Error GoTo 0
End If ' oKeyBinding Is Nothing
Next ' sKeys in dSourceIni(sCategory)
End If ' dSourceIni.Exists(sCategory
Next 'sCategory in dCategories

If Not oApp.NormalTemplate.Saved Then oApp.NormalTemplate.Save
If Len(sSourceDocx) > 0 Then
If Not oDoc.Saved Then oDoc.Save
End If
End If ' Not bReadOnly

' Create target ini
Print "Creating " & PathGetName(sTargetIni)

' AppendLine "[Commands]"
' oApp.ListCommands True
oApp.ListCommands bReadOnly
Set oTableDoc = oApp.ActiveDocument
Set dCommands = CreateDictionary
For Each oTable in oTableDoc.Tables
iRowCount = oTable.Rows.Count
' AppendBlank
' AppendLine "[; " & StringPlural("command", iRowCount - 1) & " with keys]"
If Len(sSourceDocx) = 0 Then AppendLine "[; " & StringPlural("command", iRowCount - 1) & " with keys]"
If Len(sSourceDocx) > 0 Then AppendLine "[; " & StringPlural("command", iRowCount - 1) & "]"
For iRow = 2 to iRowCount
Set oRow = oTable.Rows(iRow)
' sLine = StringTrimWhiteSpace(oApp.CleanString(oRow.Cells(2).Range.Text)) & StringTrimWhiteSpace(oApp.CleanString(oRow.Cells(3).Range.Text)) & " = " & StringTrimWhiteSpace(oApp.CleanString(oRow.Cells(1).Range.Text))
' sLine = StringTrimWhiteSpace(oApp.CleanString(oRow.Cells(1).Range.Text)) & " = " & StringTrimWhiteSpace(oApp.CleanString(oRow.Cells(2).Range.Text)) & StringTrimWhiteSpace(oApp.CleanString(oRow.Cells(3).Range.Text))
sCommand = StringChopRight(oRow.Cells(1).Range.Text, 2)
sKeys = StringChopRight(oRow.Cells(2).Range.Text, 2) & StringChopRight(oRow.Cells(3).Range.Text, 2)
If Len(sCommand) > 0 and Len(sKeys) > 0 and not dCommands.Exists(sKeys) Then dCommands.Add sKeys, sCommand
' sLine = StringChopRight(oRow.Cells(1).Range.Text, 2) & " = " & StringChopRight(oRow.Cells(2).Range.Text, 2) & StringChopRight(oRow.Cells(3).Range.Text, 2)
' sLine = StringChopRight(oRow.Cells(1).Range.Text, 2) & " = " & StringChopRight(oRow.Cells(2).Range.Text, 2) & StringChopRight(oRow.Cells(3).Range.Text, 2)
' AppendLine sLine
AppendLine sCommand & " = " & sKeys
Next 'oRow in oTable.Rows
Next ' oTable in oApp.Tables

' Sort by Keys
AppendBlank
AppendLine "[; " & StringPlural("key", dCommands.Count) & " with commands]"
For Each sKeys in ArraySort(dCommands.Keys)
AppendLine sKeys & " = " & dCommands(sKeys)
Next

' oTableDoc.SaveAs "C:\AccAuthor\WordCommands.docx"
oApp.NormalTemplate.Saved = True
On Error Resume Next
oApp.PrintPreview = False
oTableDoc.Close wdDoNotSaveChanges
On Error GoTo 0
If Len(sSourceDocx) > 0 Then oDoc.Activate

if False Then
' Set oApp.CustomizationContext = oApp
Set oAppKeyBindings = oApp.KeyBindings

for each sKeyCategory in dCategories
AppendBlank
AppendLine "[App" & sKeyCategory & "]"
iKeyCategory = dCategories(sKeyCategory)
' print "sKeyCategory " & sKeyCategory
for each oAppKeyBinding in oAppKeyBindings
If oAppKeyBinding.KeyCategory = iKeyCategory Then
AppendLine oAppKeyBinding.KeyString & " = " & oAppKeyBinding.Command
End If ' oAppKeyBinding.KeyCategory = iKeyCategory
Next 'oAppKeyBinding in oAppKeyBindings
next ' sCategory in dCategories
End If ' False

If Len(sSourceDocx) = 0 Then
oApp.CustomizationContext = oApp.NormalTemplate
Else
oApp.CustomizationContext = oDoc
End If

Set oNormalKeyBindings = oApp.KeyBindings
for each sKeyCategory in dCategories
iKeyCategory = dCategories(sKeyCategory)
set dKeyBindings = CreateDictionary
for each oNormalKeyBinding in oNormalKeyBindings
If oNormalKeyBinding.KeyCategory = iKeyCategory Then
' AppendLine oNormalKeyBinding.KeyString & " = " & oNormalKeyBinding.Command
' AppendLine oNormalKeyBinding.Command & " = " & oNormalKeyBinding.KeyString
' print oNormalKeyBinding.Command
' dKeyBindings.Add oNormalKeyBinding.Command, oNormalKeyBinding.KeyString
dKeyBindings(oNormalKeyBinding.Command) = oNormalKeyBinding.KeyString
End If ' oNormalKeyBinding.KeyCategory = iKeyCategory
Next 'oNornalKeyBinding in oNornalKeyBindings

If dKeybindings.Count > 0 Then
AppendBlank
' AppendLine "[Normal" & sKeyCategory & "]"
AppendLine "[" & sKeyCategory & "]"
For Each sCommand in dKeyBindings.Keys
sKeyString = dKeyBindings(sCommand)
AppendLine sCommand & " = " & sKeyString
Next 'sCommand in dKeyBindings.Keys
End If ' dKeyBindings.Count > 0
next ' sCategory in dCategories

oApp.PrintPreview = False
if oDocs.Count > 0 Then oDocs.Close wdDoNotSaveChanges
If Not oApp.NormalTemplate.Saved Then oApp.NormalTemplate.Save
oApp.Quit
StringToFile sHomerText, sTargetIni

If bLogActions Then
Echo "Saving " & PathGetName(sTargetLog)
StringAppendToFile sHomerLog, sTargetLog, vbFormFeed & vbCrLf
End If

echo "Done"

' Built-in Alt+Letter keys
' Alt+C = Accessibility tab
' Alt+H = Home tab
' Alt+N = Insert tab
' Alt+R = Review tab
' Alt+S = References tab
' Alt+W = View tab
