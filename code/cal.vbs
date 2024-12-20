Option Explicit

dim dEvents, bRest
dim dtStart, dtEnd
dim iCharCount, iDashCount, iDay, iCount, iEvent, iItem
dim lEvents
dim oAccount, oAccounts, oEvent, oFolder, oItem, oItems, oNamespace, oOutlook, oRestrictItems
dim s, sAccount, sFileTxt, sFilter, sHelp, sKey, sShow, sArg1, sArg2, sText

const olFolderCalendar = 9
const sDefaultParam = "today"

sHelp = "Usage: OutlookCalendarQuery.vbs [<parameter>]" & vbCrLf & _
"Parameters:" & vbCrLf & _
"  today     : Events for the entire day today." & vbCrLf & _
"  tomorrow  : Events for the entire day tomorrow." & vbCrLf & _
"  yesterday : Events for the entire day yesterday." & vbCrLf & _
"  rest      : Events for the rest of today, excluding already ended events." & vbCrLf & _
"  <datetime>: A specific date or date-time in standard format (e.g., 2024-11-10T14:30)." & vbCrLf & _
"If no parameter is provided, 'today' is assumed."

function FileSaveUtf8Bom(sFile, sText)
dim oStream, oSystem

set oSystem = createObject("Scripting.FileSystemObject")

on error resume next
fileSaveUtf8Bom = false
set oStream = oSystem.createTextFile(sFile, True, True)
oStream.write chrW(&HEF) & chrW(&HBB) & chrW(&HBF)
oStream.Write sText
oStream.close
fileSaveUtf8Bom = true
end function

function fileSaveUtf8b(sFile, sText)

dim aUtf8BOM
dim iIndex
dim oStream

fileSaveUtf8b = false
on error resume next
set oStream = createObject("ADODB.Stream")
oStream.type = 2 ' Text
oStream.charset = "utf-8"
oStream.open

for each iIndex in aUtf8BOM
rem oStream.writeText chrB(iIndex)
rem oStream.writeText chr(iIndex)
oStream.write chrw(iIndex)
next

oStream.writeText sText
oStream.saveToFile sFile, 2 ' Overwrite the file if it exists
fileSaveUtf8b = true
on error goto 0
end function

function dateStart(dt)
dateStart = cdate(formatDateTime(dt, vbShortDate) & " 12:00:00 AM")
end function

function dateEnd(dt)
dateEnd = cdate(formatDateTime(dt, vbShortDate) & " 11:59:59 PM")
end function

function formatShortDateTime(dt)
formatShortDateTime = formatDateTime(dt, vbShortDate) & " " & formatDateTime(dt, vbShortTime)
end function

function formatLongDateTime(dt)
formatLongDateTime = formatDateTime(dt, vbLongDate) & " " & formatDateTime(dt, vbLongTime)
end function

Function getNextDay(sWeekday)
dim a
dim dt
dim i, iDay
dim s, sDay

getNextDay = date
a = array("sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday")
for i = 1 to 7
dt = date + i
iDay = weekday(dt) - 1
sDay = a(iDay)
if sWeekday = sDay then
getNextDay = dt
exit function
end if
next
end function

Function charCount(sText, sChar)
Dim i, iCount
iCount = 0
For i = 1 To len(sText)
If Mid(sText, i, 1) = sChar Then
iCount = iCount + 1
End If
Next
charCount = iCount
End Function

Function LPad(sText, iLen, sChar)
Dim i : i = 0
If iLen > Len(sText) Then i = iLen - len(sText)
LPad = String(i, sChar) & sText
End Function

Function RPad(sText, iLen, sChar)
Dim i : i = 0
If iLen > Len(sText) Then i = iLen - len(sText)
RPad = sText & String(i, sChar)
End Function

Function dt2s(dt)
Dim iYear, iMonth, iDay, iHour, iMinute, iSecond
Dim sDateTime

iYear = Year(dt)
iMonth = Month(dt)
iDay = Day(dt)
iHour = Hour(dt)
iMinute = Minute(dt)
iSecond = Second(dt)

sDateTime = _
lpad(iYear, 4, "0") & "-" & _
lpad(iMonth, 2, "0") & "-" & _
lpad(iDay, 2, "0") & "T" & _
lpad(iHour, 2, "0") & ":" & _
lpad(iMinute, 2, "0") & ":" & _
lpad(iSecond, 2, "0")

dt2s = sDateTime
End Function

' Main
sArg1 = ""
sArg2 = ""
' wscript.echo wscript.arguments.count
if wscript.arguments.count > 1 then sArg2 = trim(lcase(wscript.arguments(1)))
if wscript.arguments.count > 0 then sArg1 = trim(lcase(wscript.arguments(0)))
if len(sArg1) = 0 then     sArg1 = sDefaultParam

select case sArg1
case "today"
dtStart = date
dtEnd = date + 1
case "tomorrow"
dtStart = date + 1
dtEnd = date + 2
case "yesterday"
dtStart = date - 1
dtEnd = date
case "week"
dtStart = date + 1
dtEnd = dtStart + 7
case "month"
dtStart = date + 1
dtEnd = dtStart + 30
case "now"
dtStart = now
dtEnd = dateAdd("h", 1, dtStart)
case "rest"
dtStart = now
dtEnd = dateEnd(dtStart)
case else
dtStart = getNextDay(sArg1)
if dtStart = date then
on error resume next
iDashCount = charCount(sArg1, "-")
iColonCount = charCount(sArg1, ":")
if iDashCount = 0 then sArg1 = date & " " & sArg1
if iDashCount = 1 then sArg1 = year(date) & "-" & sArg1
dtStart = cDate(sArg1)
dtEnd = dtStart + 1
on error goto 0

if not isDate(dtStart) then
wscript.echo "Invalid parameter." & vbCrLf & sHelp
wscript.quit 1
end if
end if
end select

if sArg2 <> ""then
select case sArg2
case "today"
dtEnd = date
case "tomorrow"
dtEnd = date + 1
case "yesterday"
dtEnd = date - 1
case "week"
dtEnd = date + 7
case "month"
dtEnd = date + 30
case "rest"
dtEnd = Date + 1
case else
dtEnd = getNextDay(sArg2)
s = formatDateTime(dtEnd, vbShortTime)

s = formatDateTime(dtStart, vbShortTime)
if s = "00:00" then dtStart = dateAdd("s", 1, dtStart)

if s = "00:00" then dtEnd = dateAdd("s", -1, dtEnd + 1)
' if dtEnd <> date then dtEnd = dtEnd + 1

if dtEnd = date then
on error resume next
iDashCount = charCount(sArg2, "-")
iColonCount = charCount(sArg2, ":")
if iDashCount = 0 then sArg2 = date & " " & sArg2
if iDashCount = 1 then sArg2 = year(date) & "-" & sArg2
dtEnd = cDate(sArg2)
on error goto 0

if not isDate(dtEnd) then
wscript.echo "Invalid parameter." & vbCrLf & sHelp
wscript.quit 1
end if
end if
end select
end if

set oOutlook = createObject("Outlook.Application")
set oNamespace = oOutlook.getNamespace("MAPI")
set lEvents = createObject("System.Collections.ArrayList")
set dEvents = createObject("Scripting.Dictionary")
Set oAccounts = oNamespace.Accounts

if sArg2 <> "" then
' sFilter = "[Start] >= '" & dtStart & "' and [End] <= '" & dtEnd & "'"
sFilter = "[Start] >= '" & formatShortDateTime(dtStart) & "' and [End] <= '" & formatShortDateTime(dtEnd) & "'"
else
' sFilter = "[Start] >= '" & dtStart & "' and [End] < '" & dtEnd & "'"
sFilter = "[Start] >= '" & formatShortDateTime(dtStart) & "' and [End] < '" & formatShortDateTime(dtEnd) & "'"
end if
For Each oAccount In oAccounts
sAccount = oAccount.DeliveryStore.DisplayName
Set oFolder = oAccount.DeliveryStore.GetDefaultFolder(olFolderCalendar)

' wscript.echo sFilter
set     oRestrictItems = oFolder.Items.restrict(sFilter)
oRestrictItems.includeRecurrences = true
oRestrictItems.sort "[Start]"

for each oItem in oRestrictItems
if oItem.Start >= dtStart and oItem.End < dtEnd and oItem.Class = 26 Then 'olAppointment
sKey = dt2s(oItem.Start) & "_" & dt2s(oItem.End) & "_" & oItem.Subject
lEvents.Add sKey
dEvents.add sKey, oItem
end if
next

lEvents.Sort
iCount = lEvents.Count
s = iCount & " Event"
if iCount <> 1 then s = s & "s"
sText = s & vbCrLf

sFileTxt = "cal.txt"
sShow = s & vbCrLf
for iEvent = 0 to iCount - 1
sKey = lEvents.Item(iEvent)
set oEvent = dEvents(sKey)
sText = sText & vbCrLf
sShow = sShow & vbCrLf
if iCount = 1 then
s= "Event:"
else
s = Chr(12) & vbCrLf & "Event " & (iEvent + 1) & ":"
end if
if not isEmpty(oEvent.subject) then s = s & " " & oEvent.Subject
sText = sText & s & vbCrLf
sShow = sShow & s & vbCrLf
if not isEmpty(oEvent.location) then sText = sText & "  Location: " & oEvent.location & vbCrLf
if not isEmpty(oEvent.location) then sShow = sShow & "  Location: " & oEvent.location & vbCrLf
if not isEmpty(oEvent.start) then sText = sText & "  Start: " & oEvent.start & vbCrLf
if not isEmpty(oEvent.start) then sShow = sShow & "  Start: " & oEvent.start & vbCrLf
if not isEmpty(oEvent.end) then sText = sText & "  End: " & oEvent.end & vbCrLf
if not isEmpty(oEvent.end) then sShow = sShow & "  End: " & oEvent.end & vbCrLf
if not isEmpty(oEvent.body) then sText = sText & "  Body: " & oEvent.body & vbCrLf
rem sText = sText & vbCrLf
rem sShow = sShow & vbCrLf
next
next

sShow = replace(sShow, chr(12) & vbCrLf, "")
rem sShow = replace(sShow, vbCrLf & vbCrLf, vbCrLf)
wscript.echo sShow
fileSaveUtf8b sFileTxt, sText
rem sFile = “cal.txt”
rem fileSaveUtf8Bom sFile, sShow
set oRestrictItems = nothing
set oItems = nothing
set oFolder = nothing
set oNamespace = nothing
set oOutlook = nothing
