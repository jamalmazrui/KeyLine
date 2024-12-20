option explicit

dim sNumber, sSequence, sReturn

function phoneAlpha2digits(sText)
dim dLetters 
dim iIndex
dim sChar, sReplace, sReturn

set dLetters = createObject("Scripting.Dictionary")
dLetters.compareMode = 1
dLetters.add "A", 2
dLetters.add "B", 2
dLetters.add "C", 2
dLetters.add "D", 3
dLetters.add "E", 3
dLetters.add "F", 3
dLetters.add "G", 4
dLetters.add "H", 4
dLetters.add "I", 4
dLetters.add "J", 5
dLetters.add "K", 5
dLetters.add "L", 5
dLetters.add "M", 6
dLetters.add "N", 6
dLetters.add "O", 6
dLetters.add "P", 7
dLetters.add "Q", 7
dLetters.add "R", 7
dLetters.add "S", 7
dLetters.add "T", 8
dLetters.add "U", 8
dLetters.add "V", 8
dLetters.add "W", 9
dLetters.add "X", 9
dLetters.add "Y", 9
dLetters.add "Z", 9

sReturn = ""
for iIndex = 1 to len(sText)
sChar = mid(sText, iIndex, 1)
' sChar = uCase(sChar)
if dLetters.exists(sChar) then
sReplace = dLetters(sChar)
elseif isNumeric(sChar) then
sReplace = sChar
else
sReplace = sChar
end if
sReturn = sReturn & sReplace

' wscript.echo iIndex & " " & sChar & " " & sReplace
next

set dLetters = nothing
phoneAlpha2digits = sReturn
end function

if wscript.arguments.count <> 1 then
    wscript.echo "Usage: phoneNumber.vbs <AlphaNumeric>"
    wscript.quit 1
end if

sSequence = wscript.arguments(0)
sNumber = phoneAlpha2digits(sSequence)
wscript.echo sNumber

