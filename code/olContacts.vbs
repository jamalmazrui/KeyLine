Option Explicit

Const fileName = "Contacts.csv"

dim aHeaders
Dim oContact, oContacts, oExcel, oNamespace, oOutlook, oWorkbook, oWorksheet
Dim sFilePath, sMessage
Set oExcel = CreateObject("Excel.Application")
Set oWorkbook = oExcel.Workbooks.Add
Set oWorksheet = oWorkbook.Sheets(1)
Set oOutlook = CreateObject("Outlook.Application")
Set oNamespace = oOutlook.GetNamespace("MAPI")
Set oContacts = oNamespace.GetDefaultFolder(10).Items ' Folder 10 corresponds to Contacts

rem WScript.Echo "Initialization complete

aHeaders = Array("First Name", "Last Name", "Company", "Work Email", "Home Email", "Office Phone", "Home Phone", "Mobile Phone", "Website")

Dim iHeader
For iHeader = 0 To UBound(aHeaders)
oWorksheet.Cells(1, iHeader + 1).Value = aHeaders(iHeader)
Next
rem WScript.Echo "Headers written"

Dim iRow
iRow = 2
Dim iContactCount
iContactCount = 0
wscript.echo "Getting contacts"
For Each oContact In oContacts
If oContact.Class = 40 Then ' Ensure it's a contact item
oWorksheet.Cells(iRow, 1).Value = oContact.FirstName
oWorksheet.Cells(iRow, 2).Value = oContact.LastName
oWorksheet.Cells(iRow, 3).Value = oContact.CompanyName
oWorksheet.Cells(iRow, 4).Value = oContact.Email1Address
oWorksheet.Cells(iRow, 5).Value = oContact.Email2Address
oWorksheet.Cells(iRow, 6).Value = oContact.BusinessTelephoneNumber
oWorksheet.Cells(iRow, 7).Value = oContact.HomeTelephoneNumber
oWorksheet.Cells(iRow, 8).Value = oContact.MobileTelephoneNumber
oWorksheet.Cells(iRow, 9).Value = oContact.WebPage
iRow = iRow + 1
iContactCount = iContactCount + 1
If iContactCount Mod 50 = 0 Then
WScript.Echo iContactCount
End If
End If
Next
WScript.Echo "Total " & iContactCount

Dim oSort
Set oSort = oWorksheet.Sort
With oSort
.SortFields.Clear
.SortFields.Add oWorksheet.Columns(2), 0, 1
.SortFields.Add oWorksheet.Columns(3), 0, 1
.SetRange oWorksheet.UsedRange
.Header = 1
.MatchCase = False
.Apply
End With
rem WScript.Echo "Sorted"

sFilePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") & "\" & fileName
sFilePath = Replace(sFilePath, "\\", "\")
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")
If oFSO.FileExists(sFilePath) Then oFSO.DeleteFile sFilePath
oWorkbook.SaveAs sFilePath, 6 ' xlCSV = 6
sMessage = (iRow - 2) & " contacts exported to " & fileName
rem WScript.Echo sMessage

oWorkbook.Close False
oExcel.Quit
Set oFSO = Nothing
Set oContacts = Nothing
Set oNamespace = Nothing
Set oOutlook = Nothing
Set oWorksheet = Nothing
Set oWorkbook = Nothing
Set oExcel = Nothing
