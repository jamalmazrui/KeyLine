REM Restore the old notepad.exe and prevent updates
set "sourcePath=.\notepad.exe"
set "targetPath=C:\Windows\System32\notepad.exe"
set "backupPath=C:\Windows\System32\notepad_new.exe"

REM Take ownership and grant permissions for System32 folder
echo Taking ownership of System32 directory...
takeown /f "%targetPath%" >nul 2>&1
icacls "%targetPath%" /grant "%username%:F" >nul 2>&1
if errorlevel 1 (
    echo Failed to set permissions on the Notepad target path. Ensure you have administrative privileges.
    exit /b
)

REM Backup the existing Notepad
if exist "%targetPath%" (
    echo Backing up the existing Notepad...
    copy "%targetPath%" "%backupPath%" /Y >nul
    if errorlevel 1 (
        echo Failed to back up the existing Notepad. Ensure you have write permissions.
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
        exit /b
    )
    echo Replacement successful.
) else (
    echo The old Notepad was not found in the current directory.
    exit /b
)

REM Protect Notepad from updates
echo Protecting Notepad from being replaced...
attrib +r +s "%targetPath%" >nul
if errorlevel 1 (
    echo Failed to set attributes on Notepad.
    exit /b
)
echo File attributes set to Read-Only and System.

REM Restrict permissions to prevent replacement
icacls "%targetPath%" /deny "SYSTEM:(W)" >nul
if errorlevel 1 (
    echo Failed to restrict permissions on Notepad.
    exit /b
)
echo Permissions restricted to prevent modifications by the system.
