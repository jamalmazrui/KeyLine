strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colStartupCommands = objWMIService.ExecQuery("Select * from Win32_StartupCommand")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile(".\StartupCommands.txt", True, True)
objTextFile.WriteLine colStartupCommands.Count & " startup commands"
objTextFile.WriteLine
wscript.echo colStartupCommands.Count & " startup commands"
For Each objStartupCommand in colStartupCommands
objTextFile.WriteLine "Name: " & objStartupCommand.Name
Wscript.Echo "Name: " & objStartupCommand.Name
if objStartupCommand.Description <> objStartupCommand.Name then Wscript.Echo "Description: " & objStartupCommand.Description
if objStartupCommand.Description <> objStartupCommand.Name then objTextFile.WriteLine "Description: " & objStartupCommand.Description
objTextFile.WriteLine "Command: " & objStartupCommand.Command
Wscript.Echo "Command: " & objStartupCommand.Command
' Wscript.Echo "Location: " & objStartupCommand.Location
' Wscript.Echo "Setting ID: " & objStartupCommand.SettingID
' Wscript.Echo "User: " & objStartupCommand.User
wscript.echo
objTextFile.WriteLine
Next

wScript.echo "Saving StartupCommands.txt"
objTextFile.Close

