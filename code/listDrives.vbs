' List Disk Drive Properties Using FSO


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives

dim bFirst : bFirst = true
For Each objDrive in colDrives
if not bFirst then wscript.echo
Wscript.Echo "Drive " & objDrive.DriveLetter
dim sDrive
select case objDrive.DriveType
case 0
sDrive = "Unknown"
case 1
sDrive = "Removable"
case 2
sDrive = "Fixed"
case 3
sDrive = "Network"
case 4
sDrive = "CD-ROM"
case 5
sDrive = "RAM Disk"
case else
sDrive = "Unknown"
end select
wScript.echo "Type: " & sDrive

if not objDrive.IsReady then
wScript.Echo "Drive not ready"
else
' Wscript.Echo "Available space: " & objDrive.AvailableSpace
' Wscript.Echo "File system: " & objDrive.FileSystem
' Wscript.Echo "Free space: " & objDrive.FreeSpace
' Wscript.Echo "Path: " & objDrive.Path
' Wscript.Echo "Root folder: " & objDrive.RootFolder
' Wscript.Echo "Serial number: " & objDrive.SerialNumber
' Wscript.Echo "Share name: " & objDrive.ShareName
' Wscript.Echo "Total size: " & objDrive.TotalSize
Wscript.Echo "Name: " & objDrive.VolumeName
end if
bFirst = false
Next
