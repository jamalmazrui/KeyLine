Option Explicit
' cal.vbs - Query Outlook calendar for events on a specific date or time

' Variable definitions
dim bRest
dim dtStart, dtEnd
dim oFolder, oItem, oItems, oNamespace, oOutlook, oRestrictItems
dim sFilter, sHelp, sParam

' Constants
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
on error resume next
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
dtEnd = date - 2
    case "rest"
        dtStart = now
dtEnd = Date + 1
    case else
        if isDate(sParam) then
            dtStart = cdate(sParam)
        else
            wscript.echo "Invalid parameter." & vbCrLf & sHelp
            wscript.quit 1
        end if
end select
on error goto 0

' Initialize Outlook
set oOutlook = createObject("Outlook.Application")
set oNamespace = oOutlook.getNamespace("MAPI")
set oFolder = oNamespace.getDefaultFolder(9) ' 9 = Calendar

' Retrieve calendar items
set oItems = oFolder.items
sFilter = "[Start] >= '" & formatDateTime(dtStart, 2) & "' AND [End] < '" & formatDateTime(dtEnd, 2) & "'"
'  sFilter = "[Start] >= '" & dt2s(dtStart) & "' AND [End] < '" & dt2s(dtEnd) & "'"
wscript.echo sFilter
 set     oRestrictItems = oItems.restrict(sFilter)
 oItems.includeRecurrences = true
 oItems.sort "[Start]"
if oItems.count = 0 then
    wscript.echo "No events found for the specified date-time."
else
wscript.echo oItems.Count & " events"
'     for each oItem in oRestrictItems
dim iItem
' for iItem = 1 to oRestrictItems.count
' set oItem = oRestrictItems(iItem)
for each oItem in oRestrictItems
' wscript.echo oItem.Start & " or " & dtStart
' wscript.echo oItem.End & " or " & dtEnd
 if oItem.Start >= dtStart and oItem.End < dtEnd Then
' if oItem.Class = 26 Then ' 26 = olAppointment
wscript.echo
wscript.echo "Event"
        if not isEmpty(oItem.subject) then wscript.echo "  Subject: " & oItem.subject
        if not isEmpty(oItem.location) then wscript.echo "  Location: " & oItem.location
        if not isEmpty(oItem.start) then wscript.echo "  Start: " & oItem.start
        if not isEmpty(oItem.end) then wscript.echo "  End: " & oItem.end
        if not isEmpty(oItem.body) then wscript.echo "  Body: " & oItem.body
        wscript.echo ""
end if
    next
end if

set oRestrictItems = nothing
set oItem = nothing
set oItems = nothing
set oFolder = nothing
set oNamespace = nothing
set oOutlook = nothing
