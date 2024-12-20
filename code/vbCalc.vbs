Option Explicit

dim sExpression
dim vResult

' Ensure a single argument is provided
if wscript.arguments.count = 0 then
    wscript.echo "Usage: vbCalc.vbs ""<ExcelExpression>"""
    wscript.quit(1)
end if

sExpression = wscript.arguments(0)

on error resume next
vResult = eval(sExpression)
if err.number <> 0 then
    wscript.echo err.description
else
wscript.echo vResult
end if
on error goto 0
