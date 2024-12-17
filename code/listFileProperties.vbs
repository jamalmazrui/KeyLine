' List Detailed Summary Information for a File

Set objShell = CreateObject ("Shell.Application")
Dim oShell, sCd
set oShell = WScript.CreateObject("WScript.Shell")
' sCd = oShell.ExpandEnvironmentStrings("%WINDIR%")
sCd = oShell.CurrentDirectory
set oShell = Nothing

Set objFSO = CreateObject("Scripting.FileSystemObject")
' Set objTextFile = objFSO.CreateTextFile(".\FileProperties.txt", True)
Set objTextFile = objFSO.CreateTextFile(".\FileProperties.txt", True, True)

wScript.echo "Listing file properties in " & sCd
objTextFile.WriteLine "Listing file properties in " & sCd
objTextFile.WriteLine
Set objFolder = objShell.Namespace (sCd)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim arrHeaders(33)

For i = 0 to 33
arrHeaders(i) = objFolder.GetDetailsOf (objFolder.Items, i)
Next

For Each strFileName in objFolder.Items
For i = 0 to 33
' If i <> 9 then
' Wscript.echo arrHeaders(i) & ": " & objFolder.GetDetailsOf (strFileName, i)
dim s, sName, sValue
sName = arrHeaders(i)
sValue = "" & objFolder.GetDetailsOf (strFileName, i)
if sName <> "Owner" and len(sValue) > 0 and sValue <> "Unknown" and sValue <> "Unrated" and sValue <> "Unspecified" then
if sName = "Item type" then sName = "Type"
if sName = "Date created" then sName = "Created"
if sName = "Date modified" then sName = "Modified"
if sName = "Date accessed" then sName = "Accessed"
if sValue = "File folder" then sValue = "Folder"
s = sName & ": " & sValue
objTextFile.WriteLine s
End If
Next
objTextFile.WriteLine
Next
wScript.echo "Saving FileProperties.txt"
objTextFile.Close
