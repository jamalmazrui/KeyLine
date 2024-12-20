dim sExpression, sReturn

function phoneAlpha2digits(sText)
dim dLetterToDigit 
dim iIndex
dim sCharacter, sReturn

set dLetterToDigit = createObject("Scripting.Dictionary")
dLetterToDigit.add "A", 2
dLetterToDigit.add "B", 2
dLetterToDigit.add "C", 2
dLetterToDigit.add "D", 3
dLetterToDigit.add "E", 3
dLetterToDigit.add "F", 3
dLetterToDigit.add "G", 4
dLetterToDigit.add "H", 4
dLetterToDigit.add "I", 4
dLetterToDigit.add "J", 5
dLetterToDigit.add "K", 5
dLetterToDigit.add "L", 5
dLetterToDigit.add "M", 6
dLetterToDigit.add "N", 6
dLetterToDigit.add "O", 6
dLetterToDigit.add "P", 7
dLetterToDigit.add "Q", 7
dLetterToDigit.add "R", 7
dLetterToDigit.add "S", 7
dLetterToDigit.add "T", 8
dLetterToDigit.add "U", 8
dLetterToDigit.add "V", 8
dLetterToDigit.add "W", 9
dLetterToDigit.add "X", 9
dLetterToDigit.add "Y", 9
dLetterToDigit.add "Z", 9

sReturn = ""
for iIndex = 1 to len(sText)
sCharacter = mid(sText, iIndex, 1)
sCharacter = uCase(sCharacter)
if dLetterToDigit.exists(sCharacter) then
sReturn = sReturn & dLetterToDigit(sCharacter)
elseif isNumeric(sCharacter) then
sReturn = sReturn & sCharacter
else
' Append non-alphanumeric characters as-is (e.g., hyphen, space)
sReturn = sReturn & sCharacter
end if
next

set dLetterToDigit = nothing
phoneAlpha2digits = sReturn
end function

if wscript.arguments.count <> 1 then
    wscript.echo "Usage: phoneNumber.vbs <AlphaNumeric>"
    wscript.quit 1
end if

sExpression = wscript.arguments(0)
sReturn = phoneAlpha2digits(sExpression)
wscript.echo sReturn

