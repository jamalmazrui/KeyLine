var sExpression
var vResult

if (WScript.Arguments.length == 0) {
WScript.Echo("Usage: jsCalc \"<expression>\"")

WScript.quit(1)
} // if

sExpression = WScript.Arguments(0)

try {
vResult = eval(sExpression)
WScript.Echo(vResult)
} // try
catch (e) {
WScript.Echo(e)
} // catch
