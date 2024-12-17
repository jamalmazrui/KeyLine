' xlCalc.vbs - Evaluate an Excel expression from the command line

' Constant definitions
const iErrorCodeInvalidArguments = 1
const iErrorCodeEvaluationFailed = 2

' Variable definitions
dim oExcel
dim sExpression
dim vResult

' Ensure a single argument is provided
if wscript.arguments.count <> 1 then
    wscript.echo "Usage: xlCalc.vbs ""<ExcelExpression>"""
    wscript.quit iErrorCodeInvalidArguments
end if

' Get the expression from command-line arguments
sExpression = wscript.arguments(0)

' Create Excel application object
set oExcel = createObject("Excel.Application")

' Ensure Excel is hidden
oExcel.visible = false

' Evaluate the expression
on error resume next
vResult = oExcel.evaluate(sExpression)
if err.number <> 0 then
    wscript.echo "Error evaluating expression: " & err.description
    oExcel.quit
    set oExcel = nothing
    wscript.quit iErrorCodeEvaluationFailed
else
wscript.echo vResult
end if
on error goto 0

' Clean up
oExcel.quit
set oExcel = nothing
