Option Explicit
' cal.vbs - Query Outlook calendar for events on a specific date or time

' Variable definitions
dim dEvents, bRest
dim dtStart, dtEnd
dim iCharCount, iDashCount, iDay, iCount, iEvent, iItem
dim lEvents
dim oAccount, oAccounts, oEvent, oFolder, oItem, oItems, oNamespace, oOutlook, oRestrictItems
dim s, sAccount, sFilter, sHelp, sKey, sParam, sText

' Constants
const olFolderCalendar = 9
const sDefaultParam = "today"

' Help message
sHelp = "Usage: OutlookCalendarQuery.vbs [<parameter>]" & vbCrLf & _
               "Parameters:" & vbCrLf & _
               "  today     : Events for the entire day today." & vbCrLf & _
               "  tomorrow  : Events for the entire day tomorrow." & vbCrLf & _
               "  yesterday : Events for the entire day yesterday." & vbCrLf & _
               "  rest      : Events for the rest of today, excluding already ended events." & vbCrLf & _
               "  <datetime>: A specific date or date-time in standard format (e.g., 2024-11-10T14:30)." & vbCrLf & _
               "If no parameter is provided, 'today' is assumed."

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

' Format date-time components
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
' Get the parameter from the command line
sParam = ""
if wscript.arguments.count > 0 then sParam = trim(lcase(wscript.arguments(0)))
if len(sParam) = 0 then     sParam = sDefaultParam

' Determine the target date or time
bRest = false
select case sParam
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
dtStart = date
dtEnd = date + 7
case "month"
dtStart = date
dtEnd = date + 30
    case "rest"
        dtStart = now
dtEnd = Date + 1
    case else
dtStart = getNextDay(sParam)
dtEnd = dtStart + 1
rem wscript.echo dt2s(dtStart)
if dtStart = date then
on error resume next
iDashCount = charCount(sParam, "-")
iColonCount = charCount(sParam, ":")
if iDashCount = 0 then sParam = date & " " & sParam
if iDashCount = 1 then sParam = year(date) & "-" & sParam
' if iColonCount = 1 then sParam = sParam & ":00"
rem wscript.echo sParam
dtStart = cDate(sParam)
dtEnd = dtStart + 1
on error goto 0

if not isDate(dtStart) then
            wscript.echo "Invalid parameter." & vbCrLf & sHelp
            wscript.quit 1
        end if
end if
end select
rem wscript.echo dt2s(dtStart)

' Initialize Outlook
set oOutlook = createObject("Outlook.Application")
set oNamespace = oOutlook.getNamespace("MAPI")
set lEvents = createObject("System.Collections.ArrayList")
set dEvents = createObject("Scripting.Dictionary")
Set oAccounts = oNamespace.Accounts
For Each oAccount In oAccounts
sAccount = oAccount.DeliveryStore.DisplayName
Set oFolder = oAccount.DeliveryStore.GetDefaultFolder(olFolderCalendar)
rem set oFolder = oNamespace.getDefaultFolder(9) ' 9 = Calendar

sFilter = "[Start] >= '" & dtStart & "' and [End] < '" & dtEnd & "'"
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
for iEvent = 0 to iCount - 1
sKey = lEvents.Item(iEvent)
set oEvent = dEvents(sKey)
sText = sText & vbCrLf
rem if iCount > 1 then sText = sText & "Event " & (iEvent + 1) & vbCrLf
rem if iCount > 1 then sText = sText & chr(12) & vbCrLf & "Event " & (iEvent + 1) & vbCrLf
if iCount = 1 then
s= "Event:"
else
s = Chr(12) & vbCrLf & "Event " & (iEvent + 1) & ":"
end if
rem         if not isEmpty(oEvent.subject) then sText = sText & "  Subject: " & oEvent.subject & vbCrLf
if not isEmpty(oEvent.subject) then s = s & " " & oEvent.Subject
sText = sText & s & vbCrLf
        if not isEmpty(oEvent.location) then sText = sText & "  Location: " & oEvent.location & vbCrLf
        if not isEmpty(oEvent.start) then sText = sText & "  Start: " & oEvent.start & vbCrLf
        if not isEmpty(oEvent.end) then sText = sText & "  End: " & oEvent.end & vbCrLf
        if not isEmpty(oEvent.body) then sText = sText & "  Body: " & oEvent.body & vbCrLf
sText = sText & vbCrLf
next
next

wscript.echo sText
set oRestrictItems = nothing
set oItems = nothing
set oFolder = nothing
set oNamespace = nothing
set oOutlook = nothing
