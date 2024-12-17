Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(".\InstalledPrograms.txt", True)
wScript.echo "Getting list of programs"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSoftware = objWMIService.ExecQuery _
("Select * from Win32_Product")

objTextFile.WriteLine "Caption" & vbtab & _
"Description" & vbtab & "Identifying Number" & vbtab & _
"Install Date" & vbtab & "Install Location" & vbtab & _
"Install State" & vbtab & "Name" & vbtab & _
"Package Cache" & vbtab & "SKU Number" & vbtab & "Vendor" & vbtab _
& "Version"

wScript.Echo colSoftware.Count & " programs"
Dim i
i = 1
For Each objSoftware in colSoftware
' wscript.echo i
i = i + 1
objTextFile.WriteLine objSoftware.Caption & vbtab & _
objSoftware.Description & vbtab & _
objSoftware.IdentifyingNumber & vbtab & _
objSoftware.InstallDate2 & vbtab & _
objSoftware.InstallLocation & vbtab & _
objSoftware.InstallState & vbtab & _
objSoftware.Name & vbtab & _
objSoftware.PackageCache & vbtab & _
objSoftware.SKUNumber & vbtab & _
objSoftware.Vendor & vbtab & _
objSoftware.Version
Next
objTextFile.Close
wScript.echo "Saving InstalledPrograms.txt"
