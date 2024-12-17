Option Explicit
' HomerLib by Jamal Mazrui
' November 30, 2022
' Use only in development mode
' WScript.Echo "Loading HomerLib"

Dim sHomerText : sHomerText = ""

Dim dtNow
Dim sHomerLog

dtNow = Now
sHomerLog = "Logging on " & FormatDateTime(dtNow, vbLongDate) & " at " & FormatDateTime(dtNow, vbLongTime) & vbCrLf

' Constants
Const xMute = ""
Const xSpace = " "
Const xComma = ","
Const xCommaSpace = ", "
Const xBar = "|"
Const xQuote = """"
Const xApostrophe = "'"
Const xSlash = "/"
Const xBackslash = "\"
Const xEquals = "="
Const xColon = ":"
Const xSemicolon = ";"
Const xLessThan = "<"
Const xGreaterThan = ">"
Const xLeftBrace = "{"
Const xRightBrace = "}"
Const xUTF8 = "EFBBBF"
Const xUTF16 = "FFFE"

function clipGetText()
dim oHtml

set oHtml = createObject("htmlfile")
clipGetText = oHtml.ParentWindow.ClipboardData.getData("text")
Set oHtml = Nothing
exit function

Dim oApp
Set oApp = CreateObject("InternetExplorer.Application")
oApp.Navigate("about:blank")
clipGetText = oApp.document.parentwindow.clipboardData.GetData("text")
oApp.Quit
set oApp = nothing
End Function

function clipSetText(sText)
dim bWait
dim iWindowStyle
dim sCommand, sFile
dim oShell, oExec, oIn

Set oShell = CreateObject("WScript.Shell")
Set oExec = oShell.Exec("clip")
Do While oExec.Status = 0
Set oIn = oExec.stdIn
oIn.Write sText
oIn.Close
    WScript.Sleep 100
Loop
exit function

' Previous method
iWindowStyle = 0
bWait = true
Set oShell = CreateObject("WScript.Shell")
sFile = pathGetTempFile
stringToFile sText, sFile
sCommand = "cmd /c clip <" & Chr(34) & sFile & chr(34)
' sCommand = replace(sCommand, "\\", "\")
shellRun sCommand, iWindowStyle, bWait
FileDelete(sFile)
' exit function
end function

Function Iif(iCondition, v1, v2)
' Return one of two values depending on a condition

IIf = v2
If iCondition Then IIf = v1
End Function

function cmdChoice(sPrompt, sDefault)
dim sReturn

sReturn = cmdPrompt(sPrompt)
sReturn = lCase(trim(sReturn))
if len(sReturn) > 0 then sReturn = left(sReturn, 1)
cmdChoice = sReturn
end function

function CmdPrompt(sPrompt)
dim sReturn
sReturn = ""
wscript.stdout.write sPrompt
sReturn = wscript.stdin.readline
CmdPrompt = sReturn
end function

Function DialogConfirm(sTitle, sMessage, sDefault)
' Get choice from a standard Yes, No, or Cancel message box

Dim iFlags,iChoice

DialogConfirm = ""
iFlags = vbYesNoCancel
iFlags = iFlags or vbQuestion	' 32 query icon
iFlags = iFlags Or vbSystemModal ' 4096	System modal
If sTitle = "" Then sTitle = "Confirm"
If sDefault = "N" Then iFlags = iFlags Or vbDefaultButton2
iChoice = MsgBox(sMessage, iFlags, sTitle)
If iChoice = vbCancel Then Exit Function

DialogConfirm = IIf(iChoice = vbYes, "Y", "N")
End Function

Function DialogInput(sTitle, sField, sValue)
' Get input from a single edit box

DialogInput = ""
If sField <> "" And Right(sField, 1) <> ":" Then sField = sField & ":"
DialogInput = InputBox(sField, sTitle, sValue)
End Function

Function DialogShow(oTitle, oText)
Dim iFlags
Dim sTitle, sText

' Show string version of two parameters in the title and prompt of a message box

sTitle = Str(oTitle)
sText = Str(oText)
If sTitle = "Alert" Or sTitle = "Error" Then
iFlags = vbExclamation ' 48	warning icon
Else
iFlags = vbInformation ' 64 information icon
End If

iFlags = iFlags Or vbSystemModal ' 4096	System modal
MsgBox sText, iFlags, sTitle
End Function

Function DialogShowType(sExp)
' Show expression and its type

DialogShow sExp, TypeName(Eval(sExp))
End Function

Function FormatShortDateTime(dt)
Dim sYear, sMonth, sDay, sHour, sMinute

sYear = "" & Year(dt)
sMonth = "" & Month(dt)
If Len(sMonth) = 1 Then sMonth = "0" & sMonth
sDay = "" & Day(dt)
If Len(sDay) = 1 Then sDay = "0" & sDay
sHour = "" & Hour(dt)
If Len(sHour) = 1 Then sHour = "0" & sHour
sMinute = "" & Minute(dt)
If Len(sMinute) = 1 Then sMinute = "0" & sMinute
FormatShortDateTime = sYear & "-" & sMonth & "-" & sDay & " " & sHour & "-" & sMinute
End Function

Function CloseOtherPanes(oApp)
If oApp.ActiveDocument.ActiveWindow.Panes.Count > 1 Then
For i = 2 To oApp.ActiveDocument.ActiveWindow.Panes.Count
oApp.ActiveDocument.ActiveWindow.Panes(i).Close
Next
End If
End Function

Function GetIniFile(sName)

Dim aInis
Dim sScriptVbs, sScriptRoot, sIni, sIniDir, sBinDir, sPlainKeysDir
 
GetIniFile = sName
sScriptVbs = Wscript.ScriptFullName
sBinDir = PathGetFolder(sScriptVbs)
sPlainKeysDir = PathGetFolder(sBinDir)
sIniDir = PathCombine(sPlainKeysDir, "ini")
sScriptRoot = PathGetRoot(sScriptVbs)

aInis = Array()
ArrayAdd aInis, sName
ArrayAdd aInis, sName & ".ini"
ArrayAdd aInis, sName & "-" & sScriptRoot & ".ini"

ArrayAdd aInis, PathCombine(sIniDir, sName)
ArrayAdd aInis, PathCombine(sIniDir, sName & ".ini")
ArrayAdd aInis, PathCombine(sIniDir, sName & "-" & sScriptRoot & ".ini")

For Each sIni in aInis
' print sIni
If FileExists(sIni) Then
GetIniFile = sIni
Exit Function
End If
Next 
End Function

Function GetGlobalValue(dIni, sKey, bDefault)
GetGlobalValue = bDefault
If not dIni.Exists("Global") Then Exit Function
If not dIni("Global").Exists(sKey) Then Exit Function
GetGlobalValue = CBool(dIni("Global")(sKey))
End Function

Function GetObject(oCollection, sName)
Dim o

Set GetObject = Nothing
On Error Resume Next
Set GetObject = oCollection(sName)
On Error GoTo 0
Exit Function

' Name property is not necessarily defined
For Each o In oCollection
If o.Name = sName Then
Set GetObject = o
Exit For
End If
Next
 End Function

Function DeleteObject(oCollection, sName)
Dim o

Set o = GetObject(oCollection, sName)
If Not o Is Nothing Then o.Delete
End Function

Function ForceBool(v)
ForceBool = v
On Error Resume Next
ForceBool = CBool(v)
On Error GoTo 0
End Function

Function ForceInt(v)
ForceInt = v
On Error Resume Next
ForceInt = CInt(v)
On Error GoTo 0
End Function

Function ForceSng(v)
ForceSng = v
On Error Resume Next
ForceSng = CSng(v)
On Error GoTo 0
End Function

Function ShowError()
Dim sReturn
sReturn = "Error " & Err.Number & ", " & Err.Description
sReturn = sReturn & "Source: " & Err.Source & vbCrLf
sReturn = sReturn & "ScriptLine: " & Err.ScriptLine & vbCrLf
Print(sReturn)
ShowError=sReturn
End Function

Function Quit(sMessage)
WScript.Echo sMessage
WScript.Quit
End Function

Function PrintBlank()
Print ""
End Function

Function PrintVar(sVar)
print sVar & " = " & Eval(sVar)
End Function

Function Print(sText)
WScript.Echo sText
Log sText
End Function

Function Log(sText)
sHomerLog = sHomerLog & sText & vbCrLf
End Function

Function Bail()
WScript.Quit
End Function

Function Echo(sText)
WScript.Echo sText
End Function

Function AppendText(sText)
sHomerText = sHomerText & sText
AppendText = sHomerText
End Function

Function AppendEcho(sText)
WScript.Echo sText
AppendLine sText
AppendEcho = sHomerText
End Function

Function AppendLine(sText)
Dim s

s = sText & vbCrLf
AppendText s
AppendLine = sHomerText
End Function

Function AppendBlank()
AppendLine ""
AppendBlank = sHomerText
End Function

Function CreateDictionary()
Dim oDictionary
Set oDictionary = CreateObject("Scripting.Dictionary")
oDictionary.CompareMode = vbTextCompare
Set CreateDictionary = oDictionary
End Function

Function IsSomething(o)
IsSomething = False
If Not IsObject(o) Then Exit Function
If o Is Nothing Then Exit Function
IsSomething = True
End Function

Function IsString(v)
IsString = (TypeName(v) = "String")
End Function

Function IsBlank(s)
IsBlank = (IsString(s) and Len(s) = 0)
End Function

Function IsNonBlank(s)
IsNonBlank = not IsBlank(s)
End Function

Function IsZero(n)
IsZero = (IsNumeric(n) and n = 0)
End Function

Function IsNonZero(n)
IsNonZero = not IsZero(n)
End Function

Function Min(v1, v2)
' Get minimum of two values

Min = v1
If v2 < v1 Then Min = v2
End Function

Function Str(v)
Str = ""
On Error Resume Next
Str = "" & v
On Error GoTo 0
End Function

Function Num(v)
Num = 0
On Error Resume Next
Num = 0 + v
On Error GoTo 0
End Function

Function VarGetTypeText(sType)
' Get the description of a data type

Dim oTypes

VarGetTypeText = ""
Set oTypes = CreateDictionary
oTypes.Add "Empty", "Uninitialized"
oTypes.Add "Null", "No valid data"
oTypes.Add "Object", "Generic object"
oTypes.Add "Unknown", "Unknown object type"
oTypes.Add "Nothing", "Object variable that doesn't yet refer to an object instance"
oTypes.Add "Error", "Error"
If oTypes.Exists(sType) Then VarGetTypeText = oTypes.Item(sType)
Set oTypes = Nothing
End Function

Function ArrayAdd(a, v)
' Add item to Array

Dim iBound

iBound = ArrayBound(a) + 1
Redim Preserve a(iBound)
a(iBound) = v
ArrayAdd = a
End Function

Function ArrayBound(a)
' Fix for UBound not working on empty dynamic Array

Dim iBound
iBound = -1
On Error Resume Next
iBound = UBound(a)
On Error GoTo 0
ArrayBound = iBound
End Function

Function ArrayClear(a)
' Remove all items from Array

a = Array()
ArrayClear = a
End Function

Function ArrayContains(a, v, bIgnoreCase)
' Test whether Array contains a value

ArrayContains = False
If ArrayIndex(a, v, bIgnoreCase) >= 0 Then ArrayContains = True
End Function

Function ArrayCopy(a)
' Copy an Array, not just a reference to the same items

Dim aReturn
Dim i, iBound

iBound = ArrayBound(a)
Redim aReturn(iBound)
For i = 0 to iBound
aReturn(i) = a(i)
Next
ArrayCopy = aReturn
End Function

Function ArrayCount(a)
' Return number of items in Array

ArrayCount = ArrayBound(a) + 1
End Function

Function ArrayEval(a, sExp)
' Return an Array transformed by an expression

Dim i, iBound
Dim v

iBound = ArrayBound(a)
For i = 0 To iBound
v = a(i)
a(i) = Eval(sExp)
Next
ArrayEval = a
End Function

Function ArrayFilter(a, sMatch)
' Return Array of matches of a wildcard filter expression

Const DataType = 202 ' adVarWChar
Const MaxLength = 260 ' maximum length of a file path
Dim oRS, oFields, oItem
Dim s

ArrayFilter = a
Set oRS = CreateObject("AdoDb.RecordSet")
Set oFields = oRS.Fields
Call oFields.Append("Item", DataType, MaxLength)
oRS.Open()

For Each s in a
Call oRS.AddNew("Item", s)
Next
oRS.Update()
a = ArrayClear(a)

oRS.Filter = "Item LIKE '" + sMatch + "'"
Set oItem = oFields("Item")
If Not oRS.EOF Then oRS.MoveFirst()
Do While Not ORS.EOF
ArrayAdd a, oItem.Value
oRS.MoveNext
Loop
oRS.Close
ArrayFilter = a

Set oItem = Nothing
Set oFields = Nothing
Set oRS = Nothing
End Function

Function ArrayIndex(a, v, bIgnoreCase)
' Get index of string in Array

Dim i, iBound

If bIgnoreCase <> 0 Then bIgnoreCase = 1
iBound = ArrayBound(a)
For i = 0 To iBound
If StrComp(a(i), v, bIgnoreCase) = 0 Then
ArrayIndex = i
Exit Function
End If
Next
ArrayIndex = -1
End Function

Function ArrayInsert(a, v, iIndex)
' Insert item in Array at index

Dim i, iBound

iBound = ArrayBound(a) + 1
Redim Preserve a(iBound)
For i = iBound To  iIndex + 1 Step -1
a(i) = a(i - 1)
Next
a(iIndex) = v
ArrayInsert = a
End Function

Function ArrayPop(a)
' Remove last item of Array and return the item

Dim iBound
iBound = ArrayBound(a)
ArrayPop = Empty
If iBound = -1 Then Exit Function

ArrayPop = a(iBound)
Redim Preserve a(iBound - 1)
End Function

Function ArrayRemove(a, iIndex)
' Remove item in Array at index

Dim i, iBound

iBound = ArrayBound(a)
For i = iIndex to iBound - 1
a(i) = a(i + 1)
Next
Redim Preserve a(iBound - 1)
ArrayRemove = a
End Function

Function ArrayReverse(a)
' Return Array with elements reversed

Dim aCopy
Dim i, iBound

iBound = ArrayBound(a)
aCopy = ArrayCopy(a)
For i = iBound To 0 Step - 1
a(iBound - i) = aCopy(i)
Next
ArrayReverse = a
End Function

Function ArraySort(a)
' Sort Array alphabetically

Const DataType = 202 ' adVarWChar
Const MaxLength = 260 ' maximum length of a file path
Dim oRS, oFields, oItem
Dim s

ArraySort = a
If ArrayBound(a) < 0 Then Exit Function

Set oRS = CreateObject("AdoDb.RecordSet")
Set oFields = oRS.Fields
Call oFields.Append("Item", DataType, MaxLength)
oRS.Open()

For Each s in a
Call oRS.AddNew("Item", s)
Next
oRS.Update()
a = ArrayClear(a)

oRS.Sort = "Item"
Set oItem = oFields("Item")
oRS.MoveFirst()
Do While Not oRS.EOF
ArrayAdd a, oItem.Value
oRS.MoveNext
Loop
oRS.Close
ArraySort = a

Set oItem = Nothing
Set oFields = Nothing
Set oRS = Nothing
End Function

Function ArrayUnique(a, bIgnoreCase)
' Return Array with unique items

Dim i, iBound, iIndex

iBound = ArrayBound(a)
for i = iBound To 0 Step -1
iIndex = ArrayIndex(a, a(i), bIgnoreCase)
If iIndex >=0 And iIndex < i Then Call ArrayRemove(a, i)
Next
ArrayUnique = a
End Function

' File

Function FileBackup(sSource)
Dim sDir, sRoot, sExt, sReturn

sDir = PathGetFolder(sSource)
sRoot = PathGetRoot(sSource)
sExt = PathGetExtension(sSource)
sReturn = PathGetUnique(sDir, sRoot, sExt)
If not FileCopy(sSource, sReturn) then sReturn = ""
FileBackup = sReturn
End Function

Function FileCopy(sSource, sTarget)
' Copy source to destination file, replacing if it exists

Dim oSystem

FileCopy = False
If Not FileDelete(sTarget) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
On Error Resume Next
oSystem.CopyFile sSource, sTarget
On Error GoTo 0
FileCopy = FileExists(sTarget)

Set oSystem = Nothing
End Function

Function FileDelete(sFile)
' Delete a file if it exists, and test whether it is subsequently absent
' either because it was successfully deleted or because it was not present in the first place

Dim oSystem

FileDelete = True
If Not FileExists(sFile) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Call oSystem.DeleteFile(sFile, True)
FileDelete = Not FileExists(sFile)

Set oSystem = Nothing
End Function

Function FileExists(sFile)
' Test whether File exists

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
FileExists =Not oSystem.FolderExists(sFile) And oSystem.FileExists(sFile)

Set oSystem =Nothing
End Function

Function FileGetDate(sFile)
' Get date of a file

Dim oSystem, oFile

FileGetDate = vbNull
If Not FileExists(sFile) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Set oFile =oSystem.GetFile(sFile)
FileGetDate =oFile.DateLastModified

Set oFile = Nothing
Set oSystem = Nothing
End Function

Function FileGetSize(sFile)
' Get size of a file

Dim oSystem, oFile

FileGetSize = 0
If Not FileExists(sFile) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Set oFile =oSystem.GetFile(sFile)
FileGetSize =oFile.size

Set oFile = Nothing
Set oSystem = Nothing
End Function

Function FileGetType(sFile)
' Get file type

Dim oSystem, oFile

FileGetType = ""
If Not FileExists(sFile) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Set oFile =oSystem.GetFile(sFile)
FileGetType =oFile.Type

Set oFile = Nothing
Set oSystem = Nothing
End Function

Function FileIsUTF8(sFile)
' Test whether file is UTF-8
Const ForReading = 1
Const ASCII = 0
Const Unicode = -1
Dim oSystem, oFile
Dim s1, s2, s3

FileIsUTF8 = False
If FileGetSize(sFile) < 3 Then Exit Function

Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.OpenTextFile(sFile, ForReading, ASCII)
s1 = Hex(AscB(MidB(oFile.Read(1), 1, 1)))
s2 = Hex(AscB(MidB(oFile.Read(1), 1, 1)))
s3 = Hex(AscB(MidB(oFile.Read(1), 1, 1)))
oFile.Close
If s1 & s2 & s3 = xUTF8 Then FileIsUTF8 = True
Set oFile = Nothing
Set oSystem = Nothing
End Function

Function FileIsUnicode(sFile)
' Test whether file is Unicode
Const ForReading = 1
Const ASCII = 0
Const Unicode = -1
Dim oSystem, oFile
Dim s1, s2

FileIsUnicode = False
If FileGetSize(sFile) < 2 Then Exit Function

Set oSystem = CreateObject("Scripting.FileSystemObject")
Set oFile = oSystem.OpenTextFile(sFile, ForReading, ASCII)
s1 = Hex(AscB(MidB(oFile.Read(1), 1, 1)))
s2 = Hex(AscB(MidB(oFile.Read(1), 1, 1)))
oFile.Close
' msgbox xutf16, 0, s1 & s2
If s1 & s2 = xUTF16 Then FileIsUnicode = True
If s2 & s1 = xUTF16 Then FileIsUnicode = True
Set oFile = Nothing
Set oSystem = Nothing
End Function

Function FileMove(sSource, sTarget)
' Move source to destination file, replacing if it exists

Dim oSystem

FileMove = False
If Not FileDelete(sTarget) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Call oSystem.MoveFile(sSource, sTarget)
FileMove = FileExists(sTarget)

Set oSystem = Nothing
End Function

Function FileToArray(sFile)
Dim sText

FileToArray = ""
sText = FileToString(sFile)
FileToArray = StringToArray(sText)
End Function

Function FileToDictionary(sFile)
Dim a, aParts
Dim i, iBound
Dim oData
Dim s, sKey, sValue

a = FileToArray(sFile)
Set oData = CreateDictionary()
iBound = ArrayBound(a)
For i = 0 to iBound
s = a(i)
If InStr(s, "=") Then
aParts = Split(s, "=")
sKey = Trim(aParts(0))
sValue = Trim(aParts(1))
oData.Add sKey, sValue
End If
Next
Set FileToDictionary = oData
End Function

Function FileToString(sFile)
' Get content of text file

Const ForReading = 1
Const ASCII = 0
Const Unicode = -1
Dim oSystem, oFile

FileToString = ""
If FileGetSize(sFile) = 0 Then Exit Function

Set oSystem =CreateObject("Scripting.FilesystemObject")
' if 1 then
If FileIsUnicode(sFile) Then
' DialogShow "unicode", ""
Set oFile = oSystem.OpenTextFile(sFile, ForReading, False, Unicode)
Else
' DialogShow "ascii", ""
Set oFile = oSystem.OpenTextFile(sFile, ForReading, False, ASCII)
End If
FileToString =oFile.ReadAll
oFile.Close

Set oFile = Nothing
Set oSystem = Nothing
End Function

Function IniToDictionary(sIni)
Dim aIni, aParts
Dim bAppendVerbatim, bSkipSection, bAppendValue
Dim d, dIni
Dim iSection, iBound, iLength, iLine
Dim sKey, sLine, sSection, sValue

aIni = FileToArray(sIni)
Set dIni = CreateDictionary()
iBound = ArrayBound(aIni)
bSkipSection = False
bAppendValue = False
sSection = "Global"
dIni.Add sSection, CreateDictionary()
iSection = 0
For iLine = 0 to iBound
sLine = trim(aIni(iLine))
' print "iLine " & iLine
' print "sLine " & sLine
iLength = Len(sLine)
' If iLength > 1 and Left(sLine, 2) = "[;" Then
If iLength > 2 and Left(sLine, 2) = "[;" and Right(sLine, 1) ="]" Then
bSkipSection = True
' ElseIf iLength > 0 and Left(sLine, 1) = "[" Then
ElseIf iLength > 1 and Left(sLine, 1) = "[" and Right(sLine, 1) = "]" Then
bSkipSection = False
bAppendValue = False
sSection = Trim(Mid(sLine, 2, iLength - 2))
iSection = iSection + 1
If Len(sSection) = 0 Then sSection = "Section" & iSection
dIni.Add sSection, CreateDictionary()
ElseIf bSkipSection or Left(sLine, 1) = ";" Then
' Do nothing
Else
If InStr(sLine, "=") > 0 and Left(sLine,1) <> ";" Then
	aParts = Split(sLine, "=", 2)
sKey = trim(aParts(0))
sValue = Trim(aParts(1))
' print "section=" & sSection
' print "key=" & sKey
' print "value=" & sValue
' DialogShow "", "section=" & sSection & vbCrLf & "key=" & sKey & vbCrLf & "Value=" & sValue
If sValue= Chr(34) & Chr(34) & Chr(34) Then
bAppendVerbatim = True
ElseIf Len(sValue) = 0 Then
bAppendValue = True
' sValue = vbLf
Else
bAppendValue = False
sValue = StringUnquote(sValue)
End If
ElseIf bAppendValue Then
sValue = dIni(sSection)(sKey)
' Remove leading line break
' If sValue <> vbLf Then sValue = sValue & vbLf
If Len(sValue) > 0 and sValue <> vbLf Then sValue = sValue & vbLf
sValue = sValue & aIni(iLine)
If Left(sValue, 1) = vbLf Then sValue = StringChopLeft(sValue, 1)
End If
' print "sSection" & sSection
' print "sKey" & sKey
' print "sValue " & sValue
' dIni(sSection)(sKey) = sValue
If Len(sKey) > 0 Then dIni(sSection)(sKey) = sValue
' If len(sKey) = 0 then DialogShow "", "section=" & sSection & vbCrLf & "key=" & sKey & vbCrLf & "Value=" & sValue
End If
Next

If dIni("Global").Count = 0 Then dIni.Remove "Global"
Set IniToDictionary = dIni
End Function



Function tried_IniToDictionary(sIni)
Dim aIni, aParts
Dim bSkipSection, bAppendValue
Dim d, dIni
Dim iBound, iLength, iLine
Dim sKey, sLine, sSection, sValue

aIni = FileToArray(sIni)
Set dIni = CreateDictionary()
iBound = ArrayBound(aIni)
bSkipSection = False
bAppendValue = False
sSection = "Global"
dIni.Add sSection, CreateDictionary()
For iLine = 0 to iBound
sLine = trim(aIni(iLine))
print "iLine " & iLine
print "sLine " & sLine
iLength = Len(sLine)
If iLength > 1 and Left(sLine, 2) = "[;" Then
bSkipSection = True
ElseIf iLength > 0 and Left(sLine, 1) = "[" Then
print "[ found"
bSkipSection = False
bAppendValue = False
sSection = Trim(Mid(sLine, 2, iLength - 2))
If Len(sSection) =0 Then sSection = CStr(iLine)
print "sSection=" & sSection
dIni.Add sSection, CreateDictionary()
ElseIf Not bSkipSection And InStr(sLine, "=") > 0 and Left(sLine,1) <> ";" Then
	aParts = Split(sLine, "=", 2)
sKey = trim(aParts(0))
sValue = Trim(aParts(1))
If Len(sValue) = 0 Then
print "Start appending"
bAppendValue = True
sValue = vbCrLf
ElseIf bAppendValue Then
print "Appending"
sValue = dIni(sSection)(sKey)
If sValue <> vbCrLf Then sValue = sValue & vbCrlf
sValue = sValue & aIni(iLine)
dIni(sSection)(sKey) = sValue
bAppendValue = False
Else
print "Not append"
bAppendValue = False
sValue = StringUnquote(sValue)
End If
print "sSection" & sSection
print "sKey" & sKey
print "sValue " & sValue
' make duplicate key replace previous one rather than cause dictionary error
' dIni(sSection).Add sKey, sValue
dIni(sSection)(sKey) = sValue
End If
Next
If dIni("Global").Count = 0 Then dIni.Remove "Global"
Set IniToDictionary = dIni
End Function

Function DictionaryToIni(dIni, sIni)
Dim sText

FileDelete(sIni)
sText = DictionaryToString(dIni)
StringToFile sText, sIni
DictionaryToIni = FileExists(sIni)
End Function

Function DictionaryToString(dIni)
Dim dSection
Dim sSection, sKey, sValue, sText

sText = ""
For Each sSection in dIni.Keys
If Len(sText) > 0 Then sText = sText & vbCrLf
sSection = Trim(sSection)
sText = sText & "[" & sSection & "]" & vbCrLf
Set dSection = dIni(sSection)
For Each sKey in dSection.Keys
sValue = dSection(sKey)
If StringLead(sValue, xSpace, False) Or StringTrail(sValue, xSpace, False) Then sValue = StringQuote(sValue)
sValue = StringConvertToWinLineBreak(sValue)
If InStr(sValue, vbCrLf) > 0 Then sValue = vbCrLf + sValue
sText = sText & sKey & " = " & sValue & vbCrLf
Next
Next
Set dSection = Nothing
' sText = Mid(sText, 3)
DictionaryToString = sText
End Function

' Folder

Function FolderCopy(sSource, sTarget)
' Copy source to destination Folder, replacing if it exists

Dim oSystem

FolderCopy = False
If Not FolderExists(sTarget) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Call oSystem.CopyFolder(sSource, sTarget)
FolderCopy = FolderExists(sTarget)

Set oSystem = Nothing
End Function

Function FolderCreate(sFolder)
' Create folder

Dim oSystem

FolderCreate = True
If FolderExists(sFolder) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
oSystem.CreateFolder sFolder
FolderCreate = FolderExists(sFolder)

Set oSystem = Nothing
End Function

Function FolderDelete(sFolder)
' Delete a Folder if it exists, and test whether it is subsequently absent
' either because it was successfully deleted or because it was not present in the first place

Dim oSystem

FolderDelete = True
If Not FolderExists(sFolder) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Call oSystem.DeleteFolder(sFolder, True)
FolderDelete =  Not FolderExists(sFolder)

Set oSystem = Nothing
End Function

Function FolderExists(sFolder)
' Test whether folder exists

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
FolderExists =oSystem.FolderExists(sFolder)

Set oSystem =Nothing
End Function

Function FolderGetDate(sFolder)
' Get date of a Folder

Dim oSystem, oFolder

FolderGetDate = vbNull
If Not FolderExists(sFolder) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Set oFolder =oSystem.GetFolder(sFolder)
FolderGetDate =oFolder.DateLastModified

Set oFolder = Nothing
Set oSystem = Nothing
End Function

Function FolderGetSize(sFolder)
' Get size of folder, summing the sizes of files and subfolders it contains

Dim oSystem, oFolder

FolderGetSize = 0
If Not FolderExists(sFolder) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Set oFolder =oSystem.GetFolder(sFolder)
FolderGetSize =oFolder.size

Set oFolder = Nothing
Set oSystem = Nothing
End Function

Function FolderMove(sSource, sTarget)
' Move source to destination Folder, replacing if it exists

Dim oSystem

FolderMove = False
If Not FolderExists(sSource) Then Exit Function

Set oSystem =CreateObject("Scripting.FileSystemObject")
Call oSystem.MoveFolder(sSource, sTarget)
FolderMove = FolderExists(sTarget)

Set oSystem = Nothing
End Function

' Path

Function PathCombine(sFolder, sName)
' Combine folder and name to form a valid path

Dim sPath

sPath = Trim(sFolder) & "\" & Trim(sName)
PathCombine = Replace(sPath, "\\", "\")
End Function

' Path

Function PathCreateTempFolder()
' Create temporary folder and return its full path

Dim sFolder

PathCreateTempFolder = ""
sFolder = PathGetTempFolder() & "\" & PathGetTempName()
If FolderCreate(sFolder) Then PathCreateTempFolder = sFolder
End Function

Function PathExists(sPath)
' Test whether path exists

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
PathExists =oSystem.FolderExists(sPath) Or oSystem.FileExists(sPath)

Set oSystem =Nothing
End Function

Function PathGetRoot(sPath)
' Get root name of a file or folder

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetRoot =oSystem.GetBaseName(sPath)

Set oSystem = Nothing
End Function

Function PathGetCurrentDirectory()
' Get current directory of active process

Dim oShell

Set oShell =CreateObject("Wscript.Shell")
PathGetCurrentDirectory =oShell.CurrentDirectory

Set oShell = Nothing
End Function

Function PathChangeExtension(sPath, sExt)
Dim s

s = sExt
If Len(s) > 0 And Left(s, 1) <> "." Then s = "." + s
PathChangeExtension = PathCombine(PathGetFolder(sPath), PathGetRoot(sPath) + s)
End Function

Function PathGetExtension(sPath)
' Get extention of file or folder

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetExtension =oSystem.GetExtensionName(sPath)

Set oSystem = Nothing
End Function

Function PathGetFolder(sPath)
' Get the parent folder of a file or folder

Dim oSystem
Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetFolder =oSystem.GetParentFolderName(sPath)

Set oSystem = Nothing
End Function

Function PathGetFull(sPath)
' Get the full path of a file or folder

Dim oSystem
Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetFull =oSystem.GetAbsolutePathName(sPath)

Set oSystem = Nothing
End Function

Function PathGetLong(sPath)
' Get long name of file or folder

Dim oShell, oShortcut

Set oShell = CreateObject("WScript.Shell")
Set oShortcut = oShell.CreateShortcut("temp.lnk")
oShortcut.TargetPath = sPath
PathGetLong = oShortcut.TargetPath

Set oShortcut = Nothing
Set oShell = Nothing
End Function

Function PathGetName(sPath)
' Get the file or folder name at the end of a path

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetName =oSystem.GetFileName(sPath)

Set oSystem = Nothing
End Function

Function PathGetShort(sPath)
' Get short path (8.3 style) of a file or folder

Dim oSystem, oFile, oFolder

Set oSystem =CreateObject("Scripting.FileSystemObject")
If FolderExists(sPath) Then
Set oFolder =oSystem.GetFolder(sPath)
PathGetShort =oFolder.ShortPath
Else
Set oFile =oSystem.GetFile(sPath)
PathGetShort =oFile.ShortPath
End If

Set oFile = Nothing
Set oFolder = Nothing
Set oSystem = Nothing
End Function

Function PathGetSpec(sDir, sWildcards, sFlags)
' Get an Array of paths, specifying folder, wild card pattern, and sort order

Const WindowStyle = 0 'hidden
Const Wait = True
Dim aReturn
Dim i, iBound
Dim s, sCommand,sTempFile, sReturn

sCommand = "%COMSPEC% /c dir /b " &  sFlags & " " & Chr(34) & sDir & "\" & sWildcards & Chr(34)
sCommand = replace(sCommand, "\\", "\")
sTempFile = PathGetTempFile()
sCommand = sCommand & " >" & sTempFile
ShellRun sCommand, WindowStyle, Wait
sReturn = StringTrimWhiteSpace(FileToString(sTempFile))
FileDelete(sTempFile)
PathGetSpec = Array()
If Len(sReturn) = 0 Then Exit Function

aReturn = Split(sReturn, vbCrLf)
iBound = ArrayBound(aReturn)
For i = 0 To iBound
s = aReturn(i)
If Not InStr(s, ":") Then aReturn(i) = PathCombine(sDir, s)
Next
PathGetSpec = aReturn
End Function

Function PathGetSpecialFolder(sFolder)
' Get a special folder of Windows
Dim oShell, oFolders
Dim s

PathGetSpecialFolder = ""
Set oShell =CreateObject("WScript.Shell")
Set oFolders =oShell.SpecialFolders
For Each s In oFolders
If StringTrail("\" & s, sFolder, False) Then
PathGetSpecialFolder = s
End If
Next

Set oFolders = Nothing
Set oShell = Nothing
End Function

Function PathGetTempFile()
' Get full path of a temporary file

PathGetTempFile = PathGetTempFolder() & "\" & PathGetTempName()
End Function

Function PathGetTempFolder()
' Get Windows folder for temporary files

Const TempFolder = 2
Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetTempFolder =oSystem.GetSpecialFolder(TempFolder).path

Set oSystem = Nothing
End Function

Function PathGetTempName()
' Get Name for temporary file or folder

Dim oSystem

Set oSystem =CreateObject("Scripting.FileSystemObject")
PathGetTempName = oSystem.GetTempName()

Set oSystem = Nothing
End Function

Function PathGetUnique(sDir, sRoot, sExt)
Dim i, iCount
Dim s, sIllegal, sLine, sPrintable, sSourceDir, sSourceExt, sSourceBase, sTargetDir, sTargetExt, sTargetFile, sTargetBase, sViewable

sIllegal = "@%*+\|':'<>/?" & Chr(34)
sViewable = "!#$%&'()*+,-./0123456789:'<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
sPrintable = " " & sViewable

sSourceDir = sDir
sSourceBase = sRoot
sSourceExt = sExt
If Len(sSourceExt) > 0 And Not StringLead(sSourceExt, ".", False) Then sSourceExt = "." & sSourceExt
sLine = sSourceBase
iCount = Len(sIllegal)
For i = 1 To iCount
s = Mid(sIllegal, i, 1)
If InStr(sLine, s) Then sLine = Replace(sLine, s, " ")
Next

sLine = StringReplaceAll(sLine, "  ", " ")
sLine = Trim(sLine)

sTargetBase = sLine
sTargetFile = sSourceDir & "\" & sTargetBase & sSourceExt
If FileExists(sTargetFile) Then
s = "-01"
sTargetFile = sSourceDir & "\" & sTargetBase & s &sSourceExt
i = 1
Do While FileExists(sTargetFile) And i <= 999
If i < 10 Then
s = "-00" & i
ElseIf i < 100 Then
s = "-0" & i
Else
s = "-" & i
End If

sTargetFile = sSourceDir & "\" & sTargetBase & s & sSourceExt
i = i +1
Loop
End If
PathGetUnique = sTargetFile
End Function

Function PathSetCurrentDirectory(sDir)
' Set current directory of active process

Dim oShell

Set oShell =CreateObject("Wscript.Shell")
oShell.CurrentDirectory = sDir

Set oShell = Nothing
End Function

' Process

' Thanks to Doug Lee for fixes and enhancements with process related functions

Function ProcessTerminateAllModule(sModule)
Dim iModule
Dim sCommand

sCommand = "tasklist /nh /fi ""imagename eq WinWord.exe"" | find /i ""WinWord.exe"" >nul && (echo Terminating WinWord.exe & taskkill /f /im WinWord.exe)"
' print sCommand
ShellExec("KillWord.cmd")
Exit Function

For iModule = 1 to 10
If Not ProcessIsModuleActive(sModule) Then exit For
If iModule = 1 then print "Terminating " & sModule
ShellExec("TaskKill.exe /f /im " & sModule)
Exit Function
On Error Resume Next
ProcessTerminateModule(sModule)
On Error GoTo 0
WScript.Sleep 1000
Next
If iModule = 10 then print "Error"
End Function ' ProcessTerminateAllModule Function

Function ProcessGetModules()
Dim oProcesses, oProcess, oModules

Set oModules = CreateDictionary

Set oProcesses = ProcessQueryName("")
For Each oProcess in oProcesses
oModules.Add oProcess.ProcessID, oProcess.name
Next
Set ProcessGetModules = oModules
Set oProcesses = Nothing
End Function

Function ProcessIsModuleActive(sName)
Dim oProcesses

Set oProcesses = ProcessQueryName(sName)
ProcessIsModuleActive = False
' if oProcesses.Count > 0 then ProcessIsModuleActive = True
If Not oProcesses Is Nothing Then
If oProcesses.Count > 0 then ProcessIsModuleActive = True
End If
End Function

Function ProcessQueryName(sName)
' Returns a collection of 0 or more Win32_Process objects.
' If sName is null, all processes are returned; otherwise, just the matching ones.
' If sName contains no dot, .exe is assumed.
' Note:  Only ProcessID and Name are pulled here, for efficiency.

dim oWMIService
dim sQuery, sComputer  

Set ProcessQueryName = Nothing
sComputer = "."
Set oWMIService = Nothing
On Error Resume Next
Set oWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & sComputer & "\root\cimv2")
On Error GoTo 0
If oWMIService Is Nothing Then Exit Function

' Only the properties needed by this class are pulled here.
sQuery = "select name,ProcessID from win32_process"
if len(sName) > 0 then
sQuery = sQuery & " where name='" & sName
if inStr(sName, ".") = 0 then sQuery = sQuery & ".exe"
sQuery = sQuery & "'"
end if

On Error Resume Next
Set ProcessQueryName = oWMIService.ExecQuery(sQuery)
On Error GoTo 0
End Function

Function ProcessTerminateModule(sName)
' Terminate a process by name.
' If sName contains no dot, .exe is assumed.
' Returns True on success and False on failure.
' Failure includes failure to find exactly one matching process.

Dim iResult
Dim oProcesses, oProcess
processTerminateModule = False
Set oProcesses = ProcessQueryName(sName)
if oProcesses Is Nothing Then exit function
if oProcesses.Count <> 1 then exit function

for each oProcess in oProcesses
iResult = oProcess.Terminate
' Non-zero result means error terminating process.
If iResult <> 0 then exit function
Exit For
next
processTerminateModule = True
End Function

' Regular expressions

Function RegexReplace(sText, sMatch, sReplace)
' Replace text matching a regular expression
' where sText is the string to search
' sMatch is the regular expression to match
' sReplace is the replacement text
' bIgnoreCase indicates whether capitalization matters

Dim oExp

Set oExp = CreateObject("VBScript.RegExp")
oExp.Pattern = sMatch
oExp.Ignorecase = false
oExp.Multiline = true
oExp.Global = True
RegexReplace = oExp.Replace(sText, sReplace)

Set oExp = Nothing
End Function

Function RegExpContains(sText, sMatch, bIgnoreCase)
' Get Array containing the starting index and text of the first match of a regular expression
' where sText is the string to search
' sMatch is the regular expression to match
' bIgnoreCase indicates whether capitalization matters

Dim iIndex, iCount
Dim oExp, oMatches, oMatch
Dim sValue

RegExpContains = Array(0, "")
Set oExp = CreateObject("VBScript.RegExp")
oExp.Pattern = sMatch
oExp.Ignorecase = bIgnoreCase
oExp.Multiline = False
oExp.Global = False

Set oMatches = oExp.Execute(sText)
iCount = oMatches.Count
If iCount Then
Set oMatch = oMatches.Item(0)
iIndex = oMatch.FirstIndex + 1
sValue = oMatch.Value
RegExpContains = Array(iIndex, sValue)
End If

Set oMatch = Nothing
Set oMatches = Nothing
Set oExp = Nothing
End Function

Function RegExpContainsLast(sText, sMatch, bIgnoreCase)
' Get Array containing the starting index and text of the last match of a regular expression
' where sText is the string to search
' sMatch is the regular expression to match
' bIgnoreCase indicates whether capitalization matters

Dim iIndex, iCount
Dim oExp, oMatches, oMatch
Dim sValue

Set oExp = CreateObject("VBScript.RegExp")
oExp.Pattern = sMatch
oExp.Ignorecase = bIgnoreCase
oExp.Multiline = False
oExp.Global = True

RegExpContainsLast = Array(0, "")
Set oMatches = oExp.Execute(sText)
iCount = oMatches.Count
If iCount Then
Set oMatch = oMatches.Item(iCount - 1)
iIndex = oMatch.FirstIndex + 1
sValue = oMatch.Value
RegExpContainsLast = Array(iIndex, sValue)
End If

Set oMatch = Nothing
Set oMatches = Nothing
Set oExp = Nothing
End Function

Function RegExpCount(sText, sMatch, bIgnoreCase)
' Count matches of a regular expression
' where sText is the string to search
' sMatch is the regular expression to match
' bIgnoreCase indicates whether capitalization matters

Dim oExp, oMatches, oMatch

Set oExp = CreateObject("VBScript.RegExp")
oExp.Pattern = sMatch
oExp.Ignorecase = bIgnoreCase
oExp.Multiline = False
oExp.Global = True

Set oMatches = oExp.Execute(sText)
RegExpCount = oMatches.Count

Set oMatches = Nothing
Set oExp = Nothing
End Function

Function RegExpExtract(sText, sMatch, bIgnoreCase)
' Get Array containing matches of a regular expression
' where sText is the string to search
' sMatch is the regular expression to match
' bIgnoreCase indicates whether capitalization matters

Dim aReturn()
Dim i, iCount
Dim oExp, oMatches, oMatch

Set oExp = CreateObject("VBScript.RegExp")
oExp.Pattern = sMatch
oExp.Ignorecase = bIgnoreCase
oExp.Multiline = False
oExp.Global = True

Set oMatches = oExp.Execute(sText)
iCount = oMatches.Count
Redim aReturn(iCount - 1)
i = 0
Do While i < iCount
Set oMatch = oMatches.Item(i)
aReturn(i) = oMatch.Value
i = i + 1
Loop
RegExpExtract = aReturn

Set oMatch = Nothing
Set oMatches = Nothing
Set oExp = Nothing
End Function

Function RegExpReplace(sText, sMatch, sReplace, bIgnoreCase)
' Replace text matching a regular expression
' where sText is the string to search
' sMatch is the regular expression to match
' sReplace is the replacement text
' bIgnoreCase indicates whether capitalization matters

Dim oExp

Set oExp = CreateObject("VBScript.RegExp")
oExp.Pattern = sMatch
oExp.Ignorecase = bIgnoreCase
oExp.Multiline = False
oExp.Global = True
RegExpReplace = oExp.Replace(sText, sReplace)

Set oExp = Nothing
End Function

' Registry

Function RegistryGetString(iRootKey, sSubKey, sValueName)
Dim oRegistry
Dim sValueData

RegistryGetString = ""
Set oRegistry = Nothing
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
On Error GoTo 0
If oRegistry Is Nothing Then Exit Function

oRegistry.GetStringValue iRootKey, sSubKey, sValueName, sValueData
If Typename(sValueData) <> "Null" Then RegistryGetString = sValueData
Set oRegistry = Nothing
End Function

' Registry

Function RegistryRead(sKey)
' Get a string from the registry

Dim oShell

RegistryRead = ""
Set oShell =CreateObject("Wscript.Shell")
On Error Resume Next
RegistryRead =oShell.RegRead(sKey)
Set oShell = Nothing
End Function

Function RegistrySetString(iRootKey, sSubKey, sValueName, sValueData)
Dim oRegistry

RegistrySetString = False
Set oRegistry = Nothing
On Error Resume Next
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
On Error GoTo 0
If oRegistry Is Nothing Then Exit Function

RegistrySetString = oRegistry.SetStringValue(iRootKey, sSubKey, sValueName, sValueData)
Set oRegistry = Nothing
End Function

Function RegistryWrite(sKey, sValue)
' Write a string to the registry

Dim oShell

RegistryWrite = False
Set oShell =CreateObject("Wscript.Shell")
On Error Resume Next
RegistryWrite =oShell.RegWrite(sKey, sValue, "REG_SZ")
On Error GoTo 0
Set oShell = Nothing
End Function

' Shell

Function ShellCreateShortcut(sFile, sTargetPath, sWorkingDirectory, iWindowStyle, sHotkey)
' Create a .lnk or .url file

Dim oShell, oShortcut

ShellCreateShortcut = False
If Not FileDelete(sFile) Then Exit Function

Set oShell = CreateObject("WScript.Shell")
Set oShortcut = oShell.CreateShortcut(sFile)
oShortcut.TargetPath = sTargetPath
oShortcut.WorkingDirectory = sWorkingDirectory
oShortcut.WindowStyle = iWindowStyle
oShortcut.Hotkey = sHotkey
oShortcut.Save()
ShellCreateShortcut = FileExists(sFile)

Set oShortcut = Nothing
Set oShell = Nothing
End Function

Function ShellExec(sCommand)
' Run a console mode command and return its standard output

Dim oShell, oExec, oOutput

Set oShell =CreateObject("Wscript.Shell")
Set oExec =oShell.Exec(sCommand)
Do While oExec.Status =0
WScript.Sleep(10)
Loop

Set oOutput =oExec.StdOut
ShellExec =oOutput.ReadAll()
oExec.Terminate()

Set oOutput = Nothing
Set oExec = Nothing
Set oShell = Nothing
End Function

Function ShellExecute(sFile, sParams, sFolder, sVerb, iWindowStyle)
ShellExecute = True
Set oShell = CreateObject("Shell.Application")
On Error Resume Next
oShell.ShellExecute sFile, sParams, sFolder, sVerb, iWindowStyle
On Error GoTo 0
If Err.Number Then ShellExecute = False
End Function

Function ShellExpandEnvironmentVariables(sText)
' Replace environment variables with their values

Dim oShell

Set oShell =CreateObject("Wscript.Shell")
ShellExpandEnvironmentVariables =oShell.ExpandEnvironmentStrings(sText)

Set oShell = Nothing
End Function

Function ShellGetDrives()
' Return a string sequence of drives that are ready for access

Dim i
Dim oSystem, oDrive
Dim sReturn, sDrive,sDrives

sReturn = ""
Set oSystem = CreateObject("Scripting.FileSystemObject")
sDrives = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
i = 1
Do While i <= 26
sDrive = Mid(sDrives, i, 1)
If oSystem.DriveExists(sDrive) Then
Set oDrive = oSystem.GetDrive(sDrive)
If oDrive.IsReady Then
sReturn = sReturn & sDrive
End If

Set oDrive = Nothing
End If
i = i + 1
Loop
ShellGetDrives = sReturn

Set oSystem = Nothing
End Function

Function ShellGetEnvironmentVariable(sVariable)
' Get the value of an environment variable

Dim oShell, oEnv

Set oShell =CreateObject("Wscript.Shell")
Set oEnv =oShell.Environment
ShellGetEnvironmentVariable =oEnv.Item(sVariable)

Set oEnv = Nothing
Set oShell = Nothing
End Function

Function ShellGetRecentPaths(sType)
Dim aPaths, aLinks
Dim bAdd
Dim oPaths
Dim sFolder, sWildcards, sFlags, sLink, sPath

sType = LCase(Trim(sType))
If sType = "" Then sType = "both"

ArrayClear aPaths
Set oPaths = CreateDictionary

sFolder = PathGetSpecialFolder("Recent")
sWildcards = "*.lnk"
sFlags = "/o:-d"
aLinks = PathGetSpec(sFolder, sWildcards, sFlags)
For Each sLink in aLinks
sPath = ShellGetShortcutTargetPath(sLink)
bAdd = False
If oPaths.Exists(sPath) Then
' Do nothing
ElseIf sType = "both" And PathExists(sPath) Then
bAdd = True
ElseIf sType = "folders" And FolderExists(sPath) Then
bAdd = True
ElseIf sType = "files" And FileExists(sPath) Then
bAdd = True
End If

If bAdd Then
oPaths.Add sPath, ""
ArrayAdd aPaths, sPath
End If
Next

ShellGetRecentPaths = aPaths
End Function

Function ShellGetShortcutTargetPath(sFile)
' Get the target path of a shortcut file

Dim oShell, oShortcut

Set oShell = CreateObject("WScript.Shell")
Set oShortcut = oShell.CreateShortcut(sFile)
ShellGetShortcutTargetPath = oShortcut.TargetPath
End Function

Function ShellGetSpecialFolder(vFolder)
Dim oShell, oNamespace, oFolder

Set oShell = CreateObject("Shell.Application")
Set oNamespace = oShell.Namespace(vFolder)
Set oFolder = oNamespace.Self
ShellGetSpecialFolder = oFolder.Path
Set oFolder = Nothing
Set oNamespace = Nothing
Set oShell = Nothing
End Function

Function ShellGetSpecialFolders()
Dim aNames, aValues
Dim i, iCount
Dim oShell, oFolder, oFolders, oPaths, oNames, oItem
Dim s, sName, sValue

Set oShell =CreateObject("WScript.Shell")
Set oFolders =oShell.SpecialFolders
iCount = oFolders.Count
aNames = Array()
aValues = Array()
Set oPaths = CreateDictionary
Set oNames = CreateDictionary
For i = 0 To iCount - 1
sValue = oFolders.Item(i)
sName = PathGetName(sValue)
If InStr(1, sValue, "\All Users\", vbTextCompare) Then
sName = "Common " & sName
ElseIf InStr(1, sValue, "\Users\", vbTextCompare) Then
sName = "My " & sName
ElseIf oNames.Exists(sName) And InStr(1, sValue, "\Local Settings\", vbTextCompare) Then
sName = "Local " & sName
End If
sName = Replace(sName, "Common My ", "Common ")
s = Left(sName, 1)
' If Not oPaths.Exists(sValue) And Not IsNumeric(sName) And s <> ":" And s <> "{" Then
If Not oPaths.Exists(sValue) And FolderExists(sValue) And Not IsNumeric(sName) Then
oPaths.Add sValue, sName
If Not oNames.Exists(sName) Then oNames.Add sName, ""
End If
Next

Set oShell = CreateObject("Shell.Application")
For i = 0 To 100
Set oFolder = oShell.Namespace(i)
If IsSomething(oFolder) Then
Set oItem = oFolder.Self
sValue = oItem.Path
sName = PathGetName(sValue)
If InStr(1, sValue, "\All Users\", vbTextCompare) Then
sName = "Common " & sName
ElseIf InStr(1, sValue, "\Users\", vbTextCompare) Then
sName = "My " & sName
ElseIf oNames.Exists(sName) And InStr(1, sValue, "\Local Settings\", vbTextCompare) Then
sName = "Local " & sName
End If
sName = Replace(sName, "Common My ", "Common ")
s = Left(sName, 1)
' If Not oPaths.Exists(sValue) And Not IsNumeric(sName) And s <> ":" And s <> "{" Then
If Not oPaths.Exists(sValue) And FolderExists(sValue) And Not IsNumeric(sName) Then
oPaths.Add sValue, sName
If Not oNames.Exists(sName) Then oNames.Add sName, ""
End If
End If
Next

OPaths.Add PathGetTempFolder, "Temp"
oPaths.Add ClientInformation.ScriptPath, "Window-Eyes User Profile"
Set ShellGetSpecialFolders = oPaths
Set oPaths = Nothing
Set oNames = Nothing
Set oShell = Nothing
End Function

Function ShellGetWindowsNTName()
' Get name of the Windows NT version installed
Dim iKey
Dim sSubkey, sValueName

iKey = 2 ' HKEY_LOCAL_MACHINE
sSubkey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
sValueName = "ProductName"
ShellGetWindowsNTName = RegistryRead("HKEY_LOCAL_MACHINE\" & sSubkey & "\" & sValueName)
End Function



Function ShellInvokeVerb(sPath, sVerb)
Dim oShell, oFolder, oName
Dim sFolder, sName

sFolder = PathGetFolder(sPath)
sName = PathGetName(sPath)
Set oShell = CreateObject("Shell.Application")
Set oFolder = oShell.Namespace(sFolder)
Set oName = oFolder.ParseName(sName)
oName.InvokeVerb sVerb
End Function

Function ShellRun(sFile, iStyle, bWait)
' Launch a program or file, indicating its window style and whether to wait before returning
' window styles:
' 0 Hides the window and activates another window.
' 1 Activates and displays a window. If the window is minimized or maximized, the
' system restores it to its original size and position. This flag should be used
' when specifying an application for the first time.
' 2 Activates the window and displays it minimized.
' 3 Activates the window and displays it maximized.
' 4 Displays a window in its most recent size and position. The active window
' remains active.
' 5 Activates the window and displays it in its current size and position.
' 6 Minimizes the specified window and activates the next top-level window in the Z
' order.
' 7 Displays the window as a minimized window. The active window remains active.
' 8 Displays the window in its current state. The active window remains active.
' 9 Activates and displays the window. If it is minimized or maximized, the system
' restores it to its original size and position. An application should specify
' this flag when restoring a minimized window.
' 10 Sets the show state based on the state of the program that started the
' application.

Dim oShell

Set oShell =CreateObject("Wscript.Shell")
ShellRun = -2
On Error Resume Next
ShellRun =oShell.Run(sFile, iStyle, bWait)
On Error GoTo 0

Set oShell = Nothing
End Function

Function ShellRunCommandPrompt(sDir)
' Open a command prompt in the directory specified

ShellRun "%COMSPEC% /k cd " & Chr(34) & sDir & Chr(34), 1, False
End Function

Function ShellRunExplorerWindow(sDir)
' Open Windows Explorer in the directory specified

ShellOpen sDir
' ShellRun "explorer.exe " & Chr(34) & sDir & Chr(34), 1, False
End Function

Function ShellOpen(sPath)
ShellRun StringQuote(sPath), 1, False
End Function

Function ShellOpenWith(sExe, sParam)
ShellOpenWith = ShellRun(StringQuote(sExe) & " " & StringQuote(sParam), 1, False)
End Function

Function ShellUrlToFile(sUrl, sFile)
Dim oXhttp, oAdoDb
Dim sBody, sExe

ShellUrlToFile = False
If Not FileDelete(sFile) Then Exit Function

On Error Resume Next
Set oXhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
' Set oXhttp = CreateAjaxObject()
Const WinHttpRequestOption_SslErrorIgnoreFlags  = 4  
oXhttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300 ' ignore all server certificate errors  
Const WinHttpRequestOption_EnableHttpsToHttpRedirects = 12  
oXhttp.Option(WinHttpRequestOption_EnableHttpsToHttpRedirects  ) = True

oXhttp.Open "GET", sUrl, False
oXhttp.Send

If oXhttp.Status = 200 Then
sBody = oXhttp.ResponseBody
Set oAdoDb = CreateObject("ADODB.Stream")
oAdoDb.Type = 1 ' binary
oAdoDb.Open
oAdoDb.Write sBody
oAdoDb.SaveToFile sFile, 1 ' create if not exist
oAdoDb.Close
End If
On Error GoTo 0

sExe = PathGetShort(ClientInformation.ScriptPath) & "\url2file.exe"
If FileGetSize(sFile) = 0 And FileExists(sExe) Then
ShellRun sExe & " " & sUrl & " " & StringQuote(sFile), 0, True
End If

sExe = PathGetShort(ClientInformation.ScriptPath) & "\NetUrl2File.exe"
If FileGetSize(sFile) = 0 And FileExists(sExe) Then
ShellRun sExe & " " & sUrl & " " & StringQuote(sFile), 0, True
End If

If FileGetSize(sFile) > 0 Then ShellUrlToFile = True
Set oXhttp = Nothing
Set oAdoDb = Nothing
End Function

Function ShellWait(sCommand)
ShellRun sCommand, 0, True
End Function

Function OldShellUrlToFile(sUrl, sFile)
' Download an Internet URL to a disk file

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1
Dim oXml, oStream

ShellUrlToFile = False
If Not FileDelete(sFile) Then Exit Function

Set oXML = CreateObject("MSXML2.XMLHTTP")
Call oXML.Open("GET", sURL, False)
oXML.Send()
Set oStream = CreateObject("Adodb.Stream")
oStream.Type = adTypeBinary
oStream.open()
oStream.Write(oXML.ResponseBody)
Call oStream.SaveToFile(sFile, adSaveCreateOverWrite)
oStream.Close()
ShellUrlToFile = FileExists(sFile)

Set oStream = Nothing
Set oXML = Nothing
End Function

' String

Function StringAppendToFile(sText, sFile, sDivider)
' Append string to File, omitting divider if the first one

If FileExists(sFile) Then sText = FileToString(sFile) & sDivider & sText
StringToFile sText, sFile
StringAppendToFile = FileExists(sFile)
End Function

Function StringChopLeft(sText, iCount)
' Remove iCount characters from left of sText

iCount = Min(iCount, Len(sText))
StringChopLeft = Mid(sText, iCount + 1)
End Function

Function StringCapitalize(sText)
' Capitalize first letter of sText
StringCapitalize = sText
If Len(sText) > 0 Then StringCapitalize = UCase(Left(sText, 1)) & Mid(sText, 2)
End Function

Function StringPadLeft(sText, iLength, sChar)
Dim i

i = 0
If iLength > Len(sText) Then i = iLength - len(sText)
StringPadLeft = String(i, sChar) & sText
End Function

Function StringPadRight(sText, iLength, sChar)
Dim i

i = 0
If iLength > Len(sText) Then i = iLength - len(sText)
StringPadRight = sText & String(i, sChar)
End Function

Function StringProper(sText)
' Capitalize first letter of sText and lowercase rest
StringProper = sText
If Len(sText) > 0 Then StringProper = UCase(Left(sText, 1)) & LCase(Mid(sText, 2))
End Function

Function StringChopRight(sText, iCount)
' Remove iCount characters from Right of sText

iCount = Min(iCount, Len(sText))
StringChopRight = Left(sText, Len(sText) - iCount)
End Function

Function StringContains(sText, sMatch, bIgnoreCase)
StringContains = False
If bIgnoreCase <> 0 Then bIgnoreCase = vbTextCompare
If InStr(1, sText, sMatch, bIgnoreCase) > 0 Then StringContains = True
End Function

Function StringConvertToMacLineBreak(sText)
' Convert to Macintosh line break, \r

Dim sMatch, sReplace

sMatch = vbCrLf
sReplace = vbCr
sText = Replace(sText, sMatch, sReplace)
sMatch = vbLf
sText = Replace(sText, sMatch, sReplace)
StringConvertToMacLineBreak = sText
End Function

Function StringConvertToUnixLineBreak(sText)
' Convert to Unix line break, \n

Dim sMatch, sReplace

sMatch = vbCrLf
sReplace = vbLf
sText = Replace(sText, sMatch, sReplace)
sMatch = vbCr
sText = Replace(sText, sMatch, sReplace)
StringConvertToUnixLineBreak = sText
End Function

Function StringConvertToWinLineBreak(sText)
' Convert to standard Windows line break, \r\n

Dim sMatch, sReplace

sMatch = vbCrLf
sReplace = vbLf
sText = Replace(sText, sMatch, sReplace)
sMatch = vbCr
sText = Replace(sText, sMatch, sReplace)
sMatch = vbLf
sReplace = vbCrLf
sText = Replace(sText, sMatch, sReplace)
StringConvertToWinLineBreak = sText
End Function

Function StringCount(sText, sChar)
' Count occurrences of a character in a string

StringCount = Len(sText) - Len(Replace(sText, sChar, ""))
End Function

Function StringEncode(sText)
' Encode a string for HTML or XML

Dim sReturn

sReturn = Replace(sText, "<", "&lt;")
sReturn = Replace(sText, ">", "&gt;")
sReturn = Replace(sText, "&", "&amp;")
sReturn = Replace(sText, ";", "&sc;")

StringEncode = sReturn
End Function

Function StringEndsWith(sText, sTrail, bIgnoreCase)
StringEndsWith = StringTrail(sText, sTrail, bIgnoreCase)
End Function

Function StringEqual(s1, s2)
StringEqual = False
If StrComp(s1, s2, vbBinaryCompare) = 0 Then StringEqual = True
End Function

Function StringEquiv(s1, s2)
StringEquiv = False
If StrComp(s1, s2, vbTextCompare) = 0 Then StringEquiv = True
End Function

Function StringGetASCII(sText)
' Get space delimited ASCII codes for characters in string

Dim i, iCount
Dim s, sReturn

sReturn = ""
iCount = Len(sText)
For i = 1 To iCount
s = Mid(sText, i, 1)
sReturn = sReturn & " " & ASC(s)
Next
StringGetASCII = Trim(sReturn)
End Function

Function StringIsUTF8(sText)
' Test whether a string is UTF-8
Dim s1, s2, s3

StringIsUTF8 = False
If Len(sText) < 3 Then Exit Function

s1 = Hex(AscB(MidB(sText, 1, 1)))
s2 = Hex(AscB(MidB(sText, 2, 1)))
s3 = Hex(AscB(MidB(sText, 3, 1)))
If s1 & s2 & s3 = xUTF8 Then StringIsUTF8 = True
End Function

Function StringIsUnicode(sText)
' Test whether a string is Unicode
Dim i, iCount, iCode
Dim s, s1, s2

StringIsUnicode = False
If Len(sText) <> LenB(sText) Then StringIsUnicode = True
Exit Function
If Len(sText) < 2 Then Exit Function

s1 = Hex(AscB(MidB(sText, 1, 1)))
s2 = Hex(AscB(MidB(sText, 2, 1)))
if 0 then
' If ((s1 & s2) = xUTF16) Or ((s2 & s1) = xUTF16) Then
StringIsUnicode = True
Exit Function
End If

iCount = Min(1000, Len(sText))
For i = 1 To iCount
s = Mid(sText, i, 1)
' iCode = Str("&h" & ASCW(s))
iCode = ASCW(s)
If iCode > 255 Then 
StringIsUnicode = True
Exit Function
End If
Next
End Function

function stringIsUpper(s)
stringIsUpper = false
if s = ucase(s) then stringIsUpper = true
end function

Function StringLead(sText, sLead, bIgnoreCase)
' Test whether first string starts with second one

Dim iText, iLead

StringLead = False
iText = Len(sText)
iLead = Len(sLead)
If iLead > iText Then Exit Function

If bIgnoreCase Then bIgnoreCase = 1
If (StrComp(Left(sText, iLead), sLead, bIgnoreCase)) <> 0 Then Exit Function
StringLead = True
End Function

Function StringPrependUTF8(sText)
Dim i
Dim s, sPrefix

sPrefix = ""
For i = 1 To Len(xUTF8) Step 2
s = Mid(xUTF8, i, 2)
s = "&H" & s
sPrefix = sPrefix & ChrB(s)
Next

StringPrependUTF8 = sPrefix & sText
End Function

Function StringPlural(sItem, iCount)
' Return singular or plural form of a string, depending on whether count equals one

Dim sReturn

sReturn = CStr(iCount) & " " & sItem
If iCount <> 1 Then sReturn = sReturn & "s"
StringPlural = sReturn
End Function

Function StringQuote(sText)
' Quote a string

Dim sReturn

sReturn = Chr(34) & sText & Chr(34)
StringQuote = sReturn
End Function

Function StringSingleQuote(sText)
' Quote a string

Dim sReturn

sReturn = Chr(39) & sText & Chr(39)
StringSingleQuote = sReturn
End Function

Function StringReplaceAll(sText, sMatch, sReplace)
Dim sReturn

StringReplaceAll = sText
If InStr(sReplace, sMatch) > 0 Then Exit Function

sReturn = sText
Do While InStr(sReturn, sMatch)
sReturn = Replace(sReturn, sMatch, sReplace)
Loop
StringReplaceAll = sReturn
End Function

Function StringStartsWith(sText, sLead, bIgnoreCase)
StringStartsWith = StringLead(sText, sLead, bIgnoreCase)
End Function

Function StringToArray(sText)

Dim a
Dim s

StringToArray = Array()
s = StringTrimWhiteSpace(sText)
s = StringConvertToUnixLineBreak(s)
StringToArray = Split(s, vbLf)
End Function

Function FileReadUtf8(sFile)
Dim oStream, sReturn
Set oStream = CreateObject("ADODB.Stream")
oStream.CharSet = "utf-8"
oStream.Open
oStream.LoadFromFile(sFile)
sReturn = oStream.ReadText()
FileReadUtf8 = sReturn
End Function

Function FileWriteUtf8(sFile, sText)
Dim oStream
Set oStream = CreateObject("ADODB.Stream")
oStream.Type = 2 ' Text
oStream.CharSet = "utf-8"
oStream.Open
oStream.WriteText(sText)
oStream.SaveToFile sFile, 2
oStream.Close()
FileWriteUtf8 = FileExists(sFile)
End Function

Function StringToFile(sText, sFile)
StringToFile StringPrependUTF8(sText), sFile
End Function

Function StringToFile(sText, sFile)
' Saves string to text file, replacing if it exists

Dim bReplace
Dim oSystem, oFile

StringToFile = False
If Not FileDelete(sFile) Then Exit Function

bReplace = True
Set oSystem =CreateObject("Scripting.FilesystemObject")
' DialogShow StringIsUnicode(sText), sText
if False Then
' If StringIsUnicode(sText) Then
Set oFile =oSystem.CreateTextFile(sFile, bReplace, True)
Else
Set oFile =oSystem.CreateTextFile(sFile, bReplace, False)
End If
On Error Resume Next
oFile.Write sText
On Error GoTo 0
oFile.Close
StringToFile = FileExists(sFile)

Set oFile = Nothing
Set oSystem = Nothing
End Function

Function StringTrail(sText, sTrail, bIgnoreCase)
' Test whether first string ends with second one

Dim iText, iTrail

StringTrail = False
iText = Len(sText)
iTrail = Len(sTrail)
If iTrail > iText Then Exit Function

If bIgnoreCase Then bIgnoreCase = 1
If StrComp(Right(sText, iTrail), sTrail, bIgnoreCase) <> 0 Then Exit Function
StringTrail = True
End Function

Function StringTrimWhiteSpace(sText)
' Trim leading and trailing white space characters

Dim sReturn

sReturn = Trim(sText)
sReturn = RegExpReplace(sReturn, "^\s+", "", False)
sReturn = RegExpReplace(sReturn, "\s+$", "", False)
StringTrimWhiteSpace = sReturn
End Function

Function StringUnquote(sText)
' Unquote a string

Dim sReturn

sReturn = sText
If StringLead(sReturn, Chr(34) , False) and StringTrail(sReturn, chr(34), False) Then sReturn = mid(sReturn, 2, len(sReturn) - 2)
StringUnquote = sReturn
Exit Function

' Old way
Do While Left(sReturn, 1) = Chr(34)
sReturn = StringChopLeft(sReturn, 1)
Loop

Do While Right(sReturn, 1) = Chr(34)
sReturn = StringChopRight(sReturn, 1)
Loop

StringUnquote = sReturn
End Function

Function StringSingleUnquote(sText)
' SingleUnquote a string

Dim sReturn

sReturn = sText
If StringLead(sReturn, Chr(39) , False) and StringTrail(sReturn, chr(39), False) Then sReturn = mid(sReturn, 2, len(sReturn) - 2)
StringSingleUnquote = sReturn
end Function

Function StringWrap(sText, iMaxLine) 
Dim aLines, aWords
Dim i, j
Dim sReturn, sLines, sLine, sWords, sWord

aLines = Split(sText, vbCrLf)
sReturn = ""
For i = 0 To UBound(aLines)
sLine = aLines(i)
If Len(sLine) > iMaxLine Then
aWords = Split(sLine, " ")
sLine = ""
For j = 0 To UBound(aWords)
sWord = aWords(j)
If Len(sLine & sWord) > iMaxLine Then
sReturn = sReturn & RTrim(sLine) & vbCrLf
sLine = sWord & " "
Else 
sLine = sLine & sWord & " "
End If
Next
Else
sReturn = sReturn & RTrim(sLine) & vbCrLf
End If
Next
StringWrap = sReturn
End Function

' Window

' XML

Function XMLGetTags(sXml)
' Get a dictionary of tag names and text in an XML file

Dim oScript, oDoc, oXml, oItem
Dim sName, sText

Set oScript = CreateObject("Scripting.Dictionary")
Set oDoc = CreateObject("Microsoft.XMLDOM")
oDoc.async = "false"
oDoc.Load sXml

Set oXml = oDoc.getElementsByTagName("*")
For Each oItem in oXml
sName = oItem.nodeName
sText = oItem.text
If Not oScript.Exists(sName) Then oScript.Add sName, sText
' Clipboard.Text = Clipboard.Text & sName & "=" & sText & vbCrLf
Next

Dim i, aKeys, aItems
aKeys = oScript.Keys
aItems = oScript.Items

Set XMLGetTags = oScript
Set oDoc = Nothing
Set oXml = Nothing
End Function

Function IniToCsv(sSourceIni, sFields, sSectionField, sRepeatFields)
Dim aRepeatFields, aSections, aFields, aKeys
Dim dRepeatFields, dIni, dSection
Dim iSectionField, iSection, iField, iRow, iCol, iFieldCount, iSectionCount
Dim oApp, oBook, oSheet
Dim sRepeatField, sField, sSection, sKey, sValue, sTargetCsv, sTargetXlsx

' xl vertical alignment enumeration
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

Const xlWorkbookDefault = 51
Const xlCSV = 6

If InStr(sSourceIni, "\") = 0 Then sSourceIni = PathCombine(PathGetCurrentDirectory(), sSourceIni)
sTargetCsv = PathCombine(PathGetFolder(sSourceIni), PathGetRoot(sSourceIni) & ".csv")
sTargetXlsx = PathChangeExtension(sTargetCsv, "xlsx")

Set dIni = IniToDictionary(sSourceIni)
sFields = StringReplaceAll(sFields, ", ", ",")
aFields = Split(sFields, ",")
iFieldCount = ArrayCount(aFields)
iSectionField = ArrayIndex(aFields, sSectionField, False)
if iSectionField = -1 Then Quit "Cannot find " & sSectionField & "in " & sFields

sRepeatFields = StringReplaceAll(sRepeatFields, ", ", ",")
aRepeatFields = Split(sRepeatFields, ",")
Set dRepeatFields = CreateDictionary
For Each sRepeatField in aRepeatFields
dRepeatFields.Add sRepeatField, ""
Next

Set oApp = CreateObject("Excel.Application")
Set oBook = oApp.Workbooks.Add
Set oSheet = oBook.Sheets(1)

' Set column headers
For iField = 0 to iFieldCount - 1
sField = aFields(iField)
oSheet.Cells(1, iField + 1) = sField
Next

' Populate rows of data
iRow = 1
For Each sSection in dIni.Keys
iRow = iRow + 1
' print "row " & iRow & " section " & sSection
' Set value of the field that serves as ini section header rather than key in section
oSheet.Cells(iRow, iSectionField + 1) = sSection

Set dSection = dIni(sSection)
' Iterate through keys in ini section
For iField = 0 to iFieldCount - 1
sField = aFields(iField)
if dSection.Exists(sField) Then
sValue = dSection(sField)
If StringLead(sValue, vbCrLf, False) Then sValue = Mid(sValue, 3)
oSheet.Cells(iRow, iField + 1) = sValue
if dRepeatFields.Exists(sField) Then dRepeatFields(sField) = sValue
ElseIf dRepeatFields.Exists(sField) Then
sValue = dRepeatFields(sField)
oSheet.Cells(iRow, iField + 1) = sValue
End If
Next
Next
' FileDelete sTargetCsv
oSheet.Columns.AutoFit
oSheet.Cells.WrapText = True
oSheet.Rows.AutoFit
oSheet.Rows.VerticalAlignment = xlVAlignTop  
FileDelete sTargetXlsx
' oSheet.SaveAs sTargetCsv, xlCSV
oSheet.SaveAs sTargetXlsx, xlWorkbookDefault  
oBook.Close 0
oApp.Quit

IniToCsv = FileExists(sTargetCsv)
End Function

Function IniToTables(sSourceIni, sFields, sRepeatFields, sExtensions)
Dim a, aRepeatFields, aSections, aFields, aKeys
Dim dExtensions, dFields, dRepeatFields, dIni, dSection
Dim iRowCount, iSectionField, iSection, iField, iRow, iCol, iFieldCount, iSectionCount
Dim oRange, oTable, oDoc, oDocApp, oApp, oBook, oSheet
Dim sTargetHtml, sTargetHtm, sTargetDocx, sRepeatField, sField, sSection, sKey, sValue, sTargetCsv, sTargetXlsx

' wdAutoFit enumeration
Const wdAutoFitContent = 1 'The table is automatically sized to fit the content contained in the table.
Const wdAutoFitFixed = 0 'The table is set to a fixed size, regardless of the content, and is not automatically sized.
Const wdAutoFitWindow = 2 'The table is automatically sized to the width of the active window.

Const wdFormatFilteredHTML = 10
Const wdFormatDocumentDefault = 16

' xl vertical alignment enumeration
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

Const xlWorkbookDefault = 51
Const xlCSV = 6
Const msoAutomationSecurityLow = 1

If InStr(sSourceIni, "\") = 0 Then sSourceIni = PathCombine(PathGetCurrentDirectory(), sSourceIni)
sTargetXlsx = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceIni) + ".xlsx")
sTargetCsv = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceIni) + ".csv")
sTargetDocx = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceIni) + ".docx")
sTargetHtm = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceIni) + ".htm")
sTargetHtml = PathCombine(PathGetCurrentDirectory(), PathGetRoot(sSourceIni) + ".html")

Set dIni = IniToDictionary(sSourceIni)
sFields = StringReplaceAll(sFields, ", ", ",")
sFields = StringReplaceAll(sFields, " ,", ",")
sFields = Trim(sFields)
if len(sFields) = 0 Then
Set dFields = CreateDictionary()
For Each sSection in dIni.Keys
Set dSection = dIni(sSection)
For Each sField in dSection.Keys
If Not dFields.Exists(sField) Then dFields.Add sField, ""
Next
Next
aFields = dFields.Keys
Else
aFields = Split(sFields, ",")
End If

iFieldCount = ArrayCount(aFields)

sRepeatFields = StringReplaceAll(sRepeatFields, ", ", ",")
aRepeatFields = Split(sRepeatFields, ",")
Set dRepeatFields = CreateDictionary
For Each sRepeatField in aRepeatFields
dRepeatFields.Add sRepeatField, ""
Next

Set oDocApp = CreateObject("Word.Application")
oDocApp.Visible = False
oDocApp.DisplayAlerts = False
oDocApp.ScreenUpdating = False

Set oDoc = oDocApp.Documents.Add
Set oRange = oDoc.Content
' iRowCount = 0
' Start at 1 because of row of column headers
iRowCount = 1
For Each sSection in dIni.Keys
iRowCount = iRowCount + 1
' iRowCount = iRowCount + dIni(sSection).Count
Next

print "rowCount=" & iRowCount
print "fieldCount=" & iFieldCount
Set oTable = oDoc.Tables.Add(oRange, iRowCount, iFieldCount)

Set oApp = CreateObject("Excel.Application")
oApp.AutomationSecurity = msoAutomationSecurityLow
oApp.DisplayAlerts = False
oApp.ScreenUpdating = False
oApp.Visible = False

Set oBook = oApp.Workbooks.Add
Set oSheet = oBook.Sheets(1)

' Set column headers
For iField = 0 to iFieldCount - 1
sField = aFields(iField)
oSheet.Cells(1, iField + 1) = sField
oTable.Rows(1).Cells(iField + 1).Range.Text = sField
Next

' oSheet.Cells(1,1).Name = "ColumnTitle"
' oBook.Names.Add oSheet.Cells(1,1).Range, "ColumnTitle"
Set oRange = oSheet.Range("A1")
oSheet.Names.Add "ColumnTitle", oRange

' Populate rows of data
iRow = 1
For Each sSection in dIni.Keys
iRow = iRow + 1
' print "row " & iRow & " section " & sSection

Set dSection = dIni(sSection)
' Iterate through keys in ini section
For iField = 0 to iFieldCount - 1
sField = aFields(iField)
if dSection.Exists(sField) Then
sValue = dSection(sField)
If StringLead(sValue, vbCrLf, False) Then sValue = Mid(sValue, 3)
' Remove initial line break
If StringLead(sValue, vbLf, False) Then sValue = Mid(sValue, 2)
oSheet.Cells(iRow, iField + 1) = sValue
' print "tableRows=" & oTable.Rows.Count
' print "tableColumns=" & oTable.Columns.Count
sValue = StringTrimWhiteSpace(sValue)
oTable.Rows(iRow).Cells(iField + 1).Range.Text = sValue
if dRepeatFields.Exists(sField) Then dRepeatFields(sField) = sValue
ElseIf dRepeatFields.Exists(sField) Then
sValue = dRepeatFields(sField)
oSheet.Cells(iRow, iField + 1) = sValue
oTable.Rows(iRow).Cells(iField + 1).Range.Text = sValue
End If
Next
Next

Set dExtensions = CreateDictionary
sExtensions = Replace(sExtensions, ".", " ")
sExtensions = StringReplaceAll(sExtensions, "  ", " ")
sExtensions = LCase(Trim(sExtensions))
a = Split(sExtensions, " ")
For Each s in A
dExtensions.Add s, ""
Next

oTable.AllowAutoFit = True
oTable.AutoFitBehavior wdAutoFitContent
' oTable.Columns.AutoFit
oTable.Rows(1).HeadingFormat = True
oTable.ApplyStyleHeadingRows = True
oDoc.Bookmarks.Add "ColumnTitle", oTable.Rows(1).Cells(2).Range
if dExtensions.Exists("docx") Then oDoc.SaveAs sTargetDocx, wdFormatDocumentDefault
if dExtensions.Exists("htm") Then oDoc.SaveAs sTargetHtm, wdFormatFilteredHTML  
if dExtensions.Exists("html") Then oDoc.SaveAs sTargetHtml, wdFormatFilteredHTML  
oDoc.Close 0
If Not oDocApp.NormalTemplate.Saved Then oDocApp.NormalTemplate.Saved = True
oDocApp.Quit

oSheet.Columns.AutoFit
oSheet.Cells.WrapText = True
oSheet.Rows.AutoFit
oSheet.Rows.VerticalAlignment = xlVAlignTop  

if dExtensions.Exists("csv") Then oSheet.SaveAs sTargetCsv, xlCsv
if dExtensions.Exists("xlsx") Then oSheet.SaveAs sTargetXlsx, xlWorkbookDefault  

oBook.Close 0
oApp.Quit

End Function

Function TableToDictionary(oTable)
Dim a, aRepeatFields, aSections, aFields, aKeys
Dim dExtensions, dFields, dRepeatFields, dIni, dSection
Dim iRowCount, iSectionField, iSection, iField, iRow, iCol, iFieldCount, iSectionCount
Dim oRange, oDoc, oDocApp, oApp, oBook, oSheet
Dim sTargetHtml, sTargetHtm, sTargetDocx, sRepeatField, sField, sSection, sKey, sValue, sTargetCsv, sTargetXlsx

' wdAutoFit enumeration
Const wdAutoFitContent = 1 'The table is automatically sized to fit the content contained in the table.
Const wdAutoFitFixed = 0 'The table is set to a fixed size, regardless of the content, and is not automatically sized.
Const wdAutoFitWindow = 2 'The table is automatically sized to the width of the active window.

Const wdFormatFilteredHTML = 10
Const wdFormatDocumentDefault = 16

' xl vertical alignment enumeration
Const xlVAlignBottom = -4107
Const xlVAlignCenter = -4108
Const xlVAlignDistributed = -4117
Const xlVAlignJustify = -4130
Const xlVAlignTop = -4160

Const xlWorkbookDefault = 51
Const xlCSV = 6
Const msoAutomationSecurityLow = 1

Set dTable = CreateDictionary
Set dFields = CreateDictionary
aFields = Array()
iFieldCount = oTable.Columns.Count
iRowCount = oTable.Rows.Count

iStartRow = 1
For iRow = 1 to iRowCount
If oTable.Rows(iRow).Cells.Count = iFieldCount Then
For iField = 1 to iFieldCount
sField = oTable.Cell(iRow, iField).Range.Text
sField = StringChopRight(sField, 2)
sField = Replace(sField, vbCr, " ")
sField = Replace(sField, vbLf, " ")
sField = Replace(sField, "  ", " ")
ArrayAdd aFields, sField
Next ' iField
Exit For
End If
iStartRow = iStartRow + 1
Next '' iRow
iStartRow = iStartRow + 1

For iRow = iStartRow to iRowCount
' sSection = "Section" & iRow
sSection = "Row" & (iRow - 1)
' dTable(sSection) = CreateDictionary()
dTable.Add sSection, CreateDictionary()
For iField = 1 to iFieldCount
sField = aFields(iField - 1)
sValue = oTable.Cell(iRow, iField)
sValue = StringChopRight(sValue, 2)
sValue = StringTrimWhiteSpace(sValue)
dTable(sSection)(sField) = sValue
Next ' iField
Next ' iRow
Set TableToDictionary = dTable
End Function

' Main
' IniToCsv "c:\accauthor\test.ini", "section, title, steps, type, priority", "title", "section, type, priority"
' IniToTables "c:\accauthor\test.ini", "", ""
' Dim d
' Set d = IniToDictionary("test.ini")
' DictionaryToIni d, "temp.ini"

' print "Done"
