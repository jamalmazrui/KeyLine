# Thursday, June 18, 2020

Added commands to create PowerPoint files, e.g., docx2pptx.cmd and md2pptx.cmd.

Added delJunk.cmd, which deletes miscellaneous, extraneous files after a download of many files via durlAll.cmd or saveWebDir.cmd

Added docxProof.cmd.  Run this with a .docx file as a parameter in order to produce various output files based on the document, including grammatical errors, spelling errors, styles used, embedded comments, and changes from revision tracking.

# Tuesday, June 2, 2020

Unfortunately, the Choco package manager does not seem to install on Windows 7 -- at least not with install.cmd that works with Windows 10.  Due to that problem and other incompatibilities discovered with Windows 7, I have decided to target Windows 10 only for KeyLine.  Some functionality will work on previous Windows versions, but they are not officially supported.

Made various improvements to the manual.

install.cmd now sets the dircmd environment variable so that the dir command lists file and directory names without sizes and times (extraneous speech for a screen reader user).

install.cmd now associates .ini, .inix, .md, and .mdx files with the default text editor.  This lets you open and read a Markdown (.md) file easily from File Explorer.

# Friday, May 22, 2020

Added pkh.cmd as an alias for KeyLineHelp.cmd, making it easier to open the manual in a browser.

Changed commands involving Pandoc so it can be found in other possible locations from the Choco installer.

Fixed various bugs in table-related conversions involving .inix, .htm, and .xlsx files.  The intent is to have a screen reader automatically identify column titles when navigating horizontally through a table.

# Wednesday, May 20, 2020

Added a table of contents to the manual, and updated it to describe the new installation process.  To summarize, the best way is to open a command prompt as administrator (e.g., press WindowsKey, type "cmd" and press Control+Shift+Enter).  Then change to the C:\KeyLine directory and run install.cmd.

Added the following commands, which are described in the manual:  ChromeHistory, combine, restartWindows, shutdownWindows, openDrive, closeDrive, screenOff, screenOn, saveWebDir, and move Unique.

The total number of KeyLine commands is now about 250.  I have completed programming everything envisioned for the first, release version.  I still need to fix any bugs reported and improve the learning materials.  Your help is appreciated, in order for me to subit a solid piece of work for review (and hopefully approval) as open source software contributed to the Windows field and accessibility community.

# Tuesday, May 19, 2020

Fixed bugs or improved error checking in the following commands:  ppt2htm, installSearchPath, saveAppointments, and saveContacts.

Renamed installAdmin.cmd to checkSoftwareAdmin.cmd.

The installation process has been simplified.  install.cmd now checks whether you are running as administrator.  If not, you are told that installing Calibre, Pandoc, and LibreOffice will be more automated if you re-open the cmd environment as administrator.  You can press Control+C to cancel and re-open cmd as admin, or you can press Enter to continue.  

If you continue, installDesktopShortcut.cmd and installSearchPath.cmd will be run, which do not require admin rights.  Then, depending on admin status, either checkSoftware.cmd will be run, involving manual installs of the other software packages, or checkSoftwareAdmin.cmd will be run, involving the Choco command-line package manager.

You do not need to have followed the preceding installation logic.  Just try running the install command from the KeyLine directory.  If you do so as admin, the whole KeyLine installation can be done with that single command, e.g., for installing KeyLine on another one of your computers.

Please note that I have not yet updated the manual to reflect the above changes.  I will do so and add more examples of KeyLine commands in its next revision.

Added an appointment date to the name of saved appointment files, so you know more about their content.

Fixed some typos in the KeyLine manual.

# Wednesday, May 13, 2020

This is a reminder to please test InstallAdmin.cmd (With admin rights, install Chocolatey, Calibre, Pandoc, and LibreOffice).

The KeyLine manual (C:\KeyLine\help\KeyLine.md or C:\KeyLine\help\KeyLine.htm) now describes the above command as well as the following new ones:

saveContacts.cmd. = Save all Outlook contacts

saveAppointments.cmd. = Save all Outlook appointments

moveUnique.cmd. = Move files in current directory tree to a single directory with unique file names

saveWebDir.cmd. = Save web directory to disk

ChromeHistory.cmd. = Open web page with all links in Chrome browser history

# Wednesday, May 6, 2020

In a recent KeyLine distribution, I inadvertently included files and subdirectories under the KeyLine\work directory.  These were created when I was testing the commands to get Outlook attachments.  The attachments were not intended to be shared.  Please delete them.  One way is to delete, then recreate, the work directory, if you do not need to save any other files there.

Added documentation in the introductory section of the KeyLine manual, and corrected errors in the sections about commands which control docx files.

Renamed the Outlook attachment commands so that the word "Remove" is used instead of "Delete" (to minimize confusion, since "Deleted is a folder name).

For the commands to list, save, or remove attachments, added an optional parameter specifying the file extensions to handle, e.g., "docx pdf" would only handle .docx and .pdf extensions (omit the leading dot and enclose the parameter in quotes if it contains more than one extension).

Added IniForm.txt to the KeyLine\help directory and IniForm.dll to the KeyLine\bin directory.

Added InstallAdmin.cmd, which tries to install the latest Calibre, LibreOffice, and Pandoc automatically, using the Chocolatey package manager (named after the popular candy with a "y" suffix, which is essentially the main, general software package manager for Windows, developed by a 3rd party in conformance with Microsoft conventions).

That command is what I would most like tested, if possible.  I can test a lot of commands well by myself, but naturally, installation commands especially need to be tested on different computers.  

Press WindowsKey, type "cmd" (either with or without the quotes), and then press Control+Shift+Enter, which will run the cmd environment with administrative rights.  Then change to the KeyLine directory, and enter the InstallAdmin command.

# Wednesday, April 29, 2020

- Added commands to list, save, or delete Outlook attachments.  See the Miscellaneous Commands section of documentation (which may be opened in your default browser with KeyLineHelp.cmd).  
