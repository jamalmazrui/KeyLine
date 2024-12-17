Option Explicit ' Check if the correct number of arguments is provided
If WScript.Arguments.Count <> 1 Then
WScript.Echo "Usage: cscript accessibility_checker.vbs <path_to_docx_file>"
WScript.Quit 1
End If
Dim oChecker, docPath, outputFormat, outputFileName, wordApp, doc, issues, issue, issueDetails, fs, ts
docPath = WScript.Arguments(0)
outputFileName = "Accessibility_Report"
outputFormat = "csv" ' Default output format
' Create Word application object
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
' Open the document
On Error Resume Next
Set doc = wordApp.Documents.Open(docPath)
If Err.Number <> 0 Then
WScript.Echo "Failed to open the document. Make sure the file path is correct."
WScript.Quit 1
End If
On Error GoTo 0
' Run the Accessibility Checker
wordApp.CommandBars("Accessibility").Visible = True
Set oChecker = doc.AccessibilityChecker
oChecker.Check
set issues = oChecker.Issues

' Prompt user to select output format
outputFormat = InputBox("Enter output format (csv, docx, json, html, md):", "Select Output Format", "csv")
' Prepare to save the output
Set fs = CreateObject("Scripting.FileSystemObject")
Select Case LCase(outputFormat)
Case "csv"
outputFileName = outputFileName & ".csv"
Set ts = fs.CreateTextFile(outputFileName, True)
ts.WriteLine "Issue,Description,Severity,Location"
For Each issue In issues
issueDetails = issue.Title & "," & issue.Description & "," & issue.Severity & "," & issue.Context
ts.WriteLine issueDetails
Next
ts.Close
Case "docx"
outputFileName = outputFileName & ".docx"
Set ts = wordApp.Documents.Add
ts.Content.Text = "Accessibility Issues Report" & vbCrLf & vbCrLf
For Each issue In issues
ts.Content.InsertAfter "Issue: " & issue.Title & vbCrLf & _
"Description: " & issue.Description & vbCrLf & _
"Severity: " & issue.Severity & vbCrLf & _
"Location: " & issue.Context & vbCrLf & vbCrLf
Next
ts.SaveAs2 outputFileName
ts.Close
Case "json"
outputFileName = outputFileName & ".json"
Set ts = fs.CreateTextFile(outputFileName, True)
ts.WriteLine "{""AccessibilityIssues"":["
For Each issue In issues
issueDetails = "{""Issue"":""" & issue.Title & """,""Description"":""" & issue.Description & """,""Severity"":""" & issue.Severity & """,""Location"":""" & issue.Context & """}"
ts.WriteLine issueDetails & IIf(issue Is issues(issues.Count), "",",")
Next
ts.WriteLine "]}"
ts.Close
Case "html"
outputFileName = outputFileName & ".html"
Set ts = fs.CreateTextFile(outputFileName, True)
ts.WriteLine "<html><head><title>Accessibility Issues Report</title></head><body>"
ts.WriteLine "<h1>Accessibility Issues Report</h1>"
For Each issue In issues
ts.WriteLine "<h3>Issue: " & issue.Title & "</h3>" 
ts.WriteLine "<p>Description: " & issue.Description & "</p>"
ts.WriteLine "<p>Severity: " & issue.Severity & "</p>"
ts.WriteLine "<p>Location: " & issue.Context & "</p><hr>"
Next
ts.WriteLine "</body></html>"
ts.Close
Case "md"
outputFileName = outputFileName & ".md"
Set ts = fs.CreateTextFile(outputFileName, True)
ts.WriteLine "# Accessibility Issues Report" & vbCrLf
For Each issue In issues
ts.WriteLine "## Issue: " & issue.Title & vbCrLf & _
"**Description:** " & issue.Description & vbCrLf & _
"**Severity:** " & issue.Severity & vbCrLf & _
"**Location:** " & issue.Context & vbCrLf & vbCrLf
Next
ts.Close
Case Else
WScript.Echo "Invalid format selected. Please choose between csv, docx, json, html, or md."
WScript.Quit 1
End Select
' Close the document without saving
doc.Close False
wordApp.Quit
WScript.Echo "Accessibility report saved as " & outputFileName
