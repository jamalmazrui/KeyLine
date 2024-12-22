@echo off
REM Ensure the script is run with administrative privileges
whoami /groups | find "S-1-16-12288" >nul
if errorlevel 1 (
    echo This script requires administrative privileges.
    echo Please right-click the batch file and select "Run as Administrator."
    pause
    exit /b
)

REM Task 1: Set DIRCMD to "/b" persistently
echo Setting DIRCMD environment variable to "/b"...
setx DIRCMD /b /M >nul 2>&1
if %errorlevel% equ 0 (
    echo Successfully set DIRCMD to "/b". The change will take effect in new Command Prompt windows.
) else (
    echo Failed to set DIRCMD. Please ensure you have administrative privileges.
)

REM Task 2: Configure older cmd.exe as the default terminal application
echo Configuring the older cmd.exe environment as the default terminal application...
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Terminal\Settings" /v DefaultTerminalApplication /t REG_SZ /d "{0caa0dad-35be-5f56-a8ff-afceeeaa6101}" /f >nul 2>&1
if %errorlevel% equ 0 (
    echo Successfully set the old cmd.exe as the default terminal application.
) else (
    echo Failed to set the old cmd.exe as the default terminal application. Please ensure you have administrative privileges.
)

REM Task 3: Restore the old notepad.exe and prevent updates
set "sourcePath=.\notepad.exe"
set "targetPath=C:\Windows\System32\notepad.exe"
set "backupPath=C:\Windows\System32\notepad_new.exe"

REM Backup the existing Notepad
if exist "%targetPath%" (
    echo Backing up the existing Notepad...
    copy "%targetPath%" "%backupPath%" /Y >nul
    if errorlevel 1 (
        echo Failed to back up the existing Notepad. Ensure you have write permissions.
        pause
        exit /b
    )
    echo Backup created at %backupPath%.
) else (
    echo No existing Notepad found at %targetPath%. Skipping backup.
)

REM Replace with the old Notepad
if exist "%sourcePath%" (
    echo Replacing Notepad with the old version...
    copy "%sourcePath%" "%targetPath%" /Y >nul
    if errorlevel 1 (
        echo Failed to replace Notepad. Ensure you have write permissions.
        pause
        exit /b
    )
    echo Replacement successful.
) else (
    echo The old Notepad was not found in the current directory.
    pause
    exit /b
)

REM Protect Notepad from updates
echo Protecting Notepad from being replaced...
attrib +r +s "%targetPath%" >nul
if errorlevel 1 (
    echo Failed to set attributes on Notepad.
    pause
    exit /b
)
echo File attributes set to Read-Only and System.

REM Restrict permissions to prevent replacement
icacls "%targetPath%" /deny "SYSTEM:(W)" >nul
if errorlevel 1 (
    echo Failed to restrict permissions on Notepad.
    pause
    exit /b
)
echo Permissions restricted to prevent modifications by the system.

REM Completion message
echo All tasks completed. Please restart your system for all changes to take effect.
pause
Explanation of the Tasks
Setting DIRCMD to /b:

Uses setx to set the environment variable persistently, avoiding direct registry edits.
Changes take effect in new Command Prompt windows.
Configuring the Old cmd.exe:

Uses reg add to set the default terminal application to the Windows Console Host (old cmd.exe).
The GUID {0caa0dad-35be-5f56-a8ff-afceeeaa6101} corresponds to the old console environment.
Restoring and Protecting notepad.exe:

Backups the current notepad.exe.
Replaces it with the older version from the current directory.
Sets file attributes (+r +s) to make the file read-only and system-protected.
Uses icacls to deny the SYSTEM account write permissions, preventing updates.
Usage
Save this script as configure.cmd.
Place the old notepad.exe in the same directory as the script.
Right-click configure.cmd and select Run as Administrator.
Follow the on-screen instructions and ensure each task completes successfully.
Verification
Test DIRCMD:

echo %DIRCMD%
It should output /b.

Open a new Command Prompt and check if the old console is used (no tabs or modern features).

Check notepad.exe in C:\Windows\System32:

Confirm it's the older version.
Run:
attrib C:\Windows\System32\notepad.exe
Ensure it has the R and S attributes.
Let me know if you need further refinements!













ChatGPT can make mistakes. Check important info.
?