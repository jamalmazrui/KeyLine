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

    REM Verify the registry key value
    echo Verifying the registry key value...
    for /f "tokens=2*" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Terminal\Settings" /v DefaultTerminalApplication 2^>nul') do (
        if "%%B"=="{0caa0dad-35be-5f56-a8ff-afceeeaa6101}" (
            echo Registry key verified successfully.
        ) else (
            echo Registry key value is incorrect. Expected "{0caa0dad-35be-5f56-a8ff-afceeeaa6101}" but found "%%B".
            pause
            exit /b
        )
    )
) else (
    echo Failed to set the old cmd.exe as the default terminal application. Please ensure you have administrative privileges.
)

REM Task 3: Restore the old notepad.exe and prevent updates
set "sourcePath=.\notepad.exe"
set "targetPath=C:\Windows\System32\notepad.exe"
set "backupPath=C:\Windows\System32\notepad_new.exe"

REM Take ownership and grant permissions for System32 folder
echo Taking ownership of System32 directory...
takeown /f "%targetPath%" >nul 2>&1
icacls "%targetPath%" /grant "%username%:F" >nul 2>&1
if errorlevel 1 (
    echo Failed to set permissions on the Notepad target path. Ensure you have administrative privileges.
    pause
    exit /b
)

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
