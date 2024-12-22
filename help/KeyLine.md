---
title: KeyLine
subtitle: Windows Productivity through Keyboard, Line-Oriented Commands.
author: Jamal Mazrui, Consultant, Access Success LLC
date: December 22, 2024
---

# Overview

*KeyLine* is a toolkit for productively performing various tasks on a Windows computer from the command line.  Many tasks become practical that would probably not be via a graphical user interface.  Although there is overlap, the collection of tools may be broadly divided into the following categories:  batch conversion of multiple files between formats (BatCon), accessible authoring of individual documents (AccAuthor), and miscellaneous functionality (Misc).

The name of this toolkit, KeyLine, tries to capture an operating philosophy.  The *Key* part means use of a keyboard rather than mouse. The "Line" part means that one or more lines of text are used, whenever possible, as means of input and output. This includes configuration settings that are pairs of keys and values in a .ini file.

After a graphical user interface (GUI) became popular in the 1990s, many computer users lost an appreciation of how simple and powerful a command line interface (CLI) can be.  In recent years, however, there has been a resurgence of interest in CLI-based apps, out of the desire to develop programs that are cross-platform in nature -- running with the same functionality whether on Windows, MacOS, or Linux.  This is because developing a cross-platform GUI is much more technically difficult and expensive than developing a cross-platform CLI.

Almost any Windows user can benefit from some of KeyLine functionality if effort is put into learning command-line operations, just like the effort needed to learn other nontrivial, new software.  Users of screen readers, however, are especially likely to benefit because equivalent, GUI-based techniques have often been built in a manner that is not compatible with assistive technology.  Thus, KeyLine commands for tasks may be the only accessible solution.  Some screen reader users who are particularly advanced might find workarounds to GUI accessibility obstacles, but the lack of conformance with accessibility guidelines makes such GUI apps impractical otherwise.

## Installation

KeyLine is distributed in the manner of *XCopy deployment*, where software may be installed simply by copying files into a directory (without needing an executable installer that accesses the Windows registry and protected directories).  Its single, distribution archive has a name starting with KeyLine, followed by a version number, and then the .zip file extension.  The file may be unarchived to any Windows directory.  The example for reference is ``C:\KeyLine`` but any directory, whether new or existing, will work, as long as it has read, write, and execute permissions for the files inside it.  Please note, however, that, in this test version of KeyLine, initialization commands assume the reference directory, and they have not yet been made flexible enough to initialize other locations.

After unarchiving, the installation directory of KeyLine will have five subdirectories:  `bin`, `ini`, `eg`, `help`, and `work`, which store, respectably, binary/executable files, ini/configuration files, e.g./sample files, help/documentation files, and work files.  The work directory is a default workspace for performing KeyLine commands.  It is where a command prompt opens when the KeyLine desktop shortcut is activated.  Any other directory, however, may also be set as the current directory for KeyLine commands.

## Directory Layout

Within the root KeyLine directory are nearly 300 .cmd files corresponding to KeyLine commands.  In fact, this directory only contains .cmd files.  

This file extension is synonymous with .bat, which was used when DOS (Disk Operating System) was common on personal computers.  The [Windows command prompt](https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/windows-commands) supports almost all DOS commands, with either a .bat or .cmd extension, plus many other enhancements to the Windows batch/command language.

KeyLine commands typically run utilities located in the ``C:\KeyLine\bin`` subdirectory, sometimes in combination with configuration files in the `C:\KeyLine\ini` subdirectory.  Most utilities are developed by the KeyLine author, using the [AutoIt](https://www.autoitscript.com/site/), [C#](https://docs.microsoft.com/en-us/dotnet/csharp/), [Python](https://www.python.org/), and [VBScript](https://docs.microsoft.com/en-us/previous-versions//t0aew7h6(v=vs.85)?redirectedfrom=MSDN) programming languages.  Many of the commands use the APIs of Microsoft Office applications, including Word, Excel, PowerPoint, and Outlook.  Some free, open source utilities by other authors are also included.

The syntax of KeyLine commands is intended to be simple, powerful, and consistent.  Many commands can operate over multiple files, e.g., to convert all files in a directory from one format to another.

See the section on initialization commands for ways of making KeyLine easier to use after it has been unarchived.

# Brief Introduction to the Command Prompt

The command prompt is a *shell* or *console mode* environment presented by a Windows executable file called cmd.exe.  This user interface (UI) is also known as a read-evaluate-print loop (REPL), where you type a command followed by Enter, which is read and evaluated by cmd.exe, before printing results to the screen.

To invoke a command prompt, press the Windows key for the Cortana assistant, type cmd, and press Enter.  Alternatively, press WindowsKey+R for the Run dialog, type cmd, and press Enter.

The command prompt shows the current working directory on the line before a command that you enter.  You can repeat a command by pressing UpArrow or DownArrow until you reach the one you want, and then press Enter.  You can press Control+V to paste text into a command from the Windows clipboard.

Commands often involve one or more parameters that are specified after the command name and separated by spaces.  If a parameter contains a space character, itself, then it should be enclosed in quote marks so that cmd.exe recognizes it as a single parameter.  In the commands below, a placeholder name for a parameter is enclosed in angle brackets.
  
Often, the first parameter of a KeyLine command is a file specification, \<Spec\>, which may include the asterisk \(\*\) and question mark \(?\) *wildcard* characters in order to reference multiple files matching a pattern.  The \* character matches 0 or more characters, and the ? character matches a single character.  For example, Chapter?.\* matches files that start with Chapter, followed by a single character (e.g., a digit), and then any extension.

A few common commands for directory and file management follow:

- dir = List files in the current directory.

- cd \<DirectoryName\> = Change focus to another directory.

- copy \<FileName\> \<DirectoryName\> = Copy a file to another directory.

- ren \<FileName\> \<NewName\> = Rename a file.

- del \<FileName\> = Delete a file.

- exit = Exit and close the command prompt window.

# Initialization Commands

The following commands help to configure KeyLine for convenient use at any command prompt.  For these commands to be found, you may need to execute them after making `C:\KeyLine` the current directory.

## install.cmd = Consecutively run installDesktopShortcut.cmd, installSearchPath.cmd, and checkSoftware.cmd

Syntax:\
install

The subcommands are described subsequently in this section.

## installDesktopShortcut.cmd = Install a desktop shortcut for KeyLine

Syntax:\
installDesktopShortcut

The command copies KeyLine.lnk from the ``C:\KeyLine\ini`` subdirectory to your Windows desktop directory.  The keyboard shortcut Alt+Control+P will then invoke this shortcut, opening a command prompt in the KeyLine\work subdirectory.  It also opens a directory view using the default GUI file manager (typically File Explorer), so you can easily examine files there.  Press Alt+Tab to switch between the directory view and command prompt.

## installSearchPath.cmd = Install KeyLine on the Windows search path

Syntax:\
installSearchPath

The command prepends the `C:\KeyLine` directory to your Windows search path so that KeyLine commands may be invoked from a command prompt regardless of the current directory.  The command uses the [pathEd](https://github.com/awaescher/pathEd) utility.

In order for the `dir` command to list file and directory names without extraneous information, the command also sets the `DIRCMD` environment variable to a value of `/b` (meaning bare list).
  
The command also associates a few file extensions used by KeyLine with the default text editor:  ini, inix, md, and mdx.

## cal.cmd = Query Outlook calendar

Syntax:\
cal \<DateTime\>

The parameter can be "today", "tomorrow", "yesterday", "rest", or a specific date and optional time in the format `2024-12-14T16:30`. Calendar events on that date are listed. If the parameter is "rest" or a numeric date-time, events are listed for the remaining part of the day. If no parameter is passed, "today" is assumed.

## checkSoftwareAdmin.cmd = Check for open source software, using the Chocolatey package manager with admin rights

Syntax:\
checkSoftwareAdmin

The command requires administrative rights.  You can open the cmd environment with administrative rights by pressing WindowsKey, then typing "cmd", followed by Control+Shift+Enter.

Chocolatey (named after the candy with a 'y' suffix) is the general-purpose, package manager for Windows software.  It is developed by a third party in conformance with Microsoft guidelines.  Once Choco is installed, the Calibre, Pandoc, and LibreOffice software will be installed.  A prompt will pause for confirmation before each of these installations.

## checkSoftware.cmd = Check for open source software packages

Syntax:\
checkSoftware

The command has the same purpose as checkSoftwareAdmin.cmd, without requiring admin rights, but involving manual steps of downloading and running installers.  It checks your computer for installations of [Calibre](https://calibre-ebook.com/), [LibreOffice](https://www.libreoffice.org/), and [Pandoc](https://pandoc.org/) software packages, and it offers to open their download pages so you can get the latest versions.  KeyLine relies on these packages for the best conversions among various file formats.  From the download pages that are opened consecutively, choose the latest, 64-bit versions of the installers.  

Find and run the installers from your download directory.  Accept all default choices.  KeyLine will then find these software packages under the ``C:\Program` Files` directory tree.

## kl.cmd = Add KeyLine to search path

Syntax:\
kl

If you do not wish to install the desktop shortcut or search path for KeyLine, you can still make it available from any directory by executing this command after opening a session of the command prompt.  The command prepends `C:\KeyLine` to the search path of the current session.  This command is not needed if installSearchPath.cmd has been executed (and Windows has been restarted afterward).

## CheckSoftwareAdmin.cmd = With admin rights, install Chocolatey, Calibre, Pandoc, and LibreOffice

Syntax:\
CheckSoftwareAdmin

The command is an alternative to CheckSoftware.cmd, enabling supporting software for KeyLine to be installed without finding download links on web pages.  It installs the latest Calibre, LibreOffice, and Pandoc software automatically, after installing the Chocolatey package manager (named after the candy with a "y" suffix).  Chocolatey (or Cho for short) is is essentially the main, general software package manager for Windows, developed by a 3rd party in conformance with Microsoft conventions.

This task requires a command prompt with administrative rights, which may be opened by pressing WindowsKey, typeing "cmd" (either with or without quote marks), and then pressing Control+Shift+Enter.  Remember to make `C:\KeyLine` the current directory before running the CheckSoftwareAdmin command.

# Select Commands and Examples

KeyLineHelp.cmd opens this manual in the default web browser.  pkh.cmd is an alias that is quicker to type.

To upgrade KeyLine, simply unzip the new archive into the same `C:\KeyLine` directory.  Updates files with the same names will replace prior versions, and new files will be added.  If the zip file has focus in File Explorer, you can find an unarchive command through the context menu, Shift+F10.  If the KeyLine directory has focus on the command line, you can use a command like the following:

```
unarchive <KeyLineZipFile>
```

If you use the above command, you can ignore error messages about not being able to replace the 7zip.exe and 7zip.dll files (since they are executing the unzipping process, they cannot be replaced, themselves).

Note that you can often copy and paste the following examples into the command line.  For ease of reference, they assume that you start at the `C:\KeyLine\work` command prompt.

Download many of the slide presentations from the 2019 CSUN conference on assistive technology:

```
durl http://bit.ly/csunatc19
```

Get documents (like .docx, .pdf, and .pptx) from <www.Section508.gov>:

```
saveWebDirDocs https://www.section508.gov/
```

Change to the www.section508.gov directory that was created, and then move all downloaded files to another directory, regardless of how deep they are in the tree that was created:

```
moveUnique * C:\KeyLine\work\Sec508
```

Change to the Sec508 subdirectory, and then rename all files based on title metadata:

```
Rentitle
```

Convert them all to HTML:

```
2htm
```

The following example illustrates several commands.

As we know, Google search results put the most popular matches first, which correlates with the quality of those results.  Although PDF has been a hindrance to accessibility in other ways, a strength it has is that people generally do not publish in PDF format unless a work is relatively polished in content.  Thus, to get free articles or books in PDF format that are likely (but not guaranteed) to be quality pieces on a topic, do a Google search that specifies .pdf as the file extension for results.  

For example, try the following Google search:

```
ext:pdf intitle:personal finance" investment stocks
```

On the results page, press Alt+D for the address bar and then Control+C to copy selected text to the clipboard (usually the URL is automatically selected).

Then go to a command prompt and enter the following:

```
durl “<PastedURL>”
```

The enclosing quotes are usually not needed, but it makes a difference in some URLs.  You can paste with Control+V on the command line.

durl.cmd will analyze each link on the Google results page and download all document types by default, which includes .pdf.  Several files will probably be downloaded.

Next, run delDupes.cmd to make sure that none of the downloads are duplicates of one another.  

Then, use rentitle.cmd as follows:

```
Rentitle *.pdf
```

Generally (though not always), this improves file names because they are typically more cryptic than the title metadata inside them.  

Finally, get HTML versions by running the following command:

```
Pdf2htm
```

If plain text versions are also desired, enter the following command:

```
Htm2txt
```

# Batch Conversion of File Formats (BatCon)

Many commands convert files from one format to another.  In general, these commands invoke programs that can automatically detect the format of the source file.  Sometimes, they assume a format based on the file extension, e.g., .docx for a Word document.  Thus, for best conversion results, it is worth consistently naming files with extensions that correspond to their formats or types.

Since screen reader support for reading tends to be strongest in a web browser with HTML content, the .htm or .html extension is a common target format.  By convention, web servers tend to use a .html extension whereas local Windows computers tend to use the .htm extension, but the extensions refer to the same file type and may be used anywhere.

When a .docx, .pdf, or .epub file is converted to .htm and opened in a browser like Google Chrome or Mozilla Firefox, it is easy to navigate by heading, do global searches, and follow links.  If the .htm file is very large, e.g., converted from a .pdf book, it may take a half minute or more to load the file, but once loaded, reading, navigating, and selecting text can be highly efficient.  While a large file is loading, you can switch to another application window to do other work in the meantime.

The KeyLine file conversion commands use standard extensions for convenient reference to file types.  For example, pdf2htm.cmd converts files from .pdf to .htm format.  The digit "2" is an abbreviation for the word"to."  The abbreviation also enables the command to be read distinctly with a screen reader.  If "pdftohtm" were written instead, a screen reader would not say the source and target extensions distinctly, since all the characters would be interpreted to be part of the same word.  The command "pdftohtm" could be written instead, since Windows ignores case in this context, but that means additional work pressing the shift key in order to create commands in mixed case.  Such capitalization is used, however, in names of other commands besides straightforward file conversions, e.g., delDupes.cmd to delete duplicate files in a directory.

Each file conversion command can convert many files with a single command.  The source files to be converted may be specified using the \* and ? wildcard characters, e.g.,

```
pdf2htm chapter*.pdf
```

This command converts all files that start with chapter and end with .pdf.  If no wildcard spec is given, a logical, default one is assumed, e.g., \*.pdf is assumed for pdf2htm.

The file conversions use a variety of utilities behind the scenes, depending on the source and target file formats.  These utilities include the open source packages Calibre, LibreOffice, and Pandoc.  Other utilities are developed by this author.

When a file conversion command is executed, a consistent pattern of console output occurs.  For each file meeting the wildcard spec, a check is made for whether the target file already exists in the current directory, that is, a file with the same root name as the source but the target extension instead of the source one.  If so, that source file is skipped on the assumption that it has already been converted.  You can cause the conversion to occur by deleting the target file first.  If no target file is found, the name of the source file is printed to the console screen before the conversion is attempted.  If the conversion is successful, no further message appears related to that conversion.  If the target file is not created, however, the message "Error" is printed to inform you that something went wrong with that conversion.

Such commands support the following source extensions:  chm, csv, doc, docx, epub, htm, html, ipynb, ini, inix, md, mdx, pdf, ppt, pptx, rtf, vtt, xls, and xlsx.

Such commands support the following target extensions:  csv, docx, epub, htm, html, inix, md, pdf, pptx, tables, txt, xlsx, and xwiki.

If you want to convert various source extensions to a single target extension, this is often also possible with a single command.  For example, 2htm.cmd tries to convert every to .htm format, regardless of the source extension (which is absent from the command name).  With such commands, \*.\* is assumed as a wildcard spec, though another may optionally be given instead, e.g.,

```
2htm chapter*.*
```

This command converts all files starting with chapter to .htm format, regardless of source extension.

Such commands include the following:  2docx, 2epub, 2htm, 2html, 2inix, 2kfx, 2pdf, 2rtf, and 2txt.

A couple of commands convert a source extension to all target extensions, where the target extension is omitted from the command name.  inix2.cmd converts to .csv, .docx, .htm, and .xlsx.  md2cmd converts to .docx and .htm.  The intent is to convert to target formats that may be compared as a way of checking the source for possible errors.

# Accessible Authoring of Individual Documents (AccAuthor)

Several KeyLine commands support accessible authoring with Microsoft word .docx files and Markdown .md files.  The *AccAuthor* abbreviation is used in this context, especially for settings recommended by the author, e.g., Microsoft Word may be configured with AccAuthor options, and its keyboard interface may be enhanced with AccAuthor keyboard shortcuts.

Among other things, AccAuthor enables you, as an author, to produce a well-structured, well-formatted Kindle book in both .docx and .kfx formats..  a KFX file may be sideloaded into the Kindle for PC app.  The book may be read there by a screen reader user, either as a reader or as an author wishing to inspect the book for possible errors before uploading it to Kindle Direct Publishing (KDP).

a File in .docx format, used by Microsoft Word and other software, is a key component in the pipeline.  The utilities offer extensive control over formatting and metadata in DOCX files through open source, plain text techniques.

Most AccAuthor utilities are written in the VBScript language and executed by Windows batch files with the CScript.exe interpreter that is built into Windows.  They all import a library of script functions called HomerLib (HomerLib.vbs).

Each command may be passed an input configuration file in .ini format.  Most utilities also report the current configuration, after optionally applying changes, in an output .ini file.  If the .docx file is modified by a utility, a backup of the source is created (e.g., test-01.docx is a backup of the original test.docx before modification).  (A setting can prevent creating such backup files if not wanted.)

For convenient, command-line invocation, the commands that accept an input configuration file support an abbreviated syntax.  If a partial file name is specified, which is not found, the command adds a suffix consisting of its root name and the .ini extension.  If no such file is found in the current directory, the file is sought in the `C:\KeyLine\Ini` subdirectory.  For example, you can configure the AccAuthor keyboard interface with the command "wordKeys AccAuthor".  Presumably a file called AccAuthor is not found in the current directory, so the command looks for AccAuthor-wordKeys.ini.  If that file is not found, the command checks the global directory for .ini configuration files, ``C:\KeyLine\ini`,` and does the same sequence of checks for a file.

By default, AccAuthor commands that use .ini configuration files create a backup copy of a .docx file that is modified, using a numeric suffix after the root name, e.g., if test.docx is modified, the backup might be test-003.docx if test-001.docx and test-002.docx already exist.  In addition, an ongoing log file is maintained for changes made to that .docx file.  If you do not want these additional files created, you can add the following settings to an initial, unnamed, global section of the .ini file:

```
BackupDocx=False
LogActions=False
```

Later sections of this guide give details on the sophisticated commands for controlling Microsoft Word and .docx files.

The following commands help authors when writing documents in Word or other formats.

## definer.cmd = Show definitions for a word

Syntax:\
definer \<word\>

A phrase of multiple words may also be defined if surrounded by quotes.  Output is saved in a text file with the DEFINE- prefix followed by the term defined.  The command uses the dictionary.com website.

## docxProof.cmd = Generate various proofing reports of a docx file.

Syntax:\
docxProof \<spec\>

For each file matching spec (e.g., \*.docx), the following commands (explained subsequently) are run:  docxGrammar, docxProperties, docxSpelling, and docxStyles.  docx2md.cmd is also run if a similar file with a .md extension does not already exist, since a Markdown version will show heading structure, links, bold, italic, and other formatting through textual symbols.  The command uses the Word API.

## docxGrammar.cmd = Show sentences that contain grammatical errors

Syntax:\
docxGrammar \<SourceFile\>

The target file has the same root name, with a GRAMMAR- prefix, and a .txt extension.

## docxIt.cmd = Perform miscellaneous tasks on a file

Syntax:\
docxIt \<SourceFile\> \<ConfigFile\>

Possible tasks include generating a table of contents, as well as search and replace.

## docxProperties.cmd = Show or change properties of a .docx file

Syntax:\
docxProperties \<SourceFile\> \<ConfigFile\>

The target file begins with the "PROPERTIES-" prefix and has a .ini extension.  For example, PROPERTIES-Work.ini results from Work.docx.

## docxSpelling.cmd = Show spelling errors and suggestions

Syntax:\
docxSpelling \<SourceFile\>

The target file has the same root name, with a SPELLING- prefix, and a .txt extension.

## docxStyles.cmd = Show or change styles of a .docx file

Syntax:\
docxStyles \<SourceFile\> \<ConfigFile\>

The target file begins with the "STYLES-" prefix and has a .ini extension.  For example, STYLES-Work.ini results from Work.docx.

## termIt.cmd = Show synonyms and antonyms for a term

Syntax:\
termIt \<term\>

A term can be either a single word or a quoted phrase.

## testPage.cmd = Test web pages for accessibility

Syntax:\
testPage \<URL\>

The page specified by URL will be tested using the API of the IBM Accessibility Checker. The results are stored in a folder named after the title of the page. Multiple URLs can be tested by passing the name of a text file with one URL per line.

## testUrl.cmd = Test web pages for accessibility

Syntax:\
testUrl \<URL\>

The page specified by URL will be tested using the API of Deque Axe. The results are output as an HTML report named after the title of the page. Multiple URLs can be tested by passing the name of a text file with one URL per line.

## wordKeys.cmd = Show or change keyboard shortcuts in Word

Syntax:\
wordKeys \<SourceFile\> \<ConfigFile\>

The target file is KEYS.ini.

## wordOptions.cmd = Show or change application options in Word

Syntax:\
wordOptions \<SourceFile\> \<ConfigFile\>

The target file is OPTIONS.ini.

The following commands help you reset Microsoft Word to default settings (which may be challenging to do through the GUI):

## moveWordAddIns.cmd = Move Word add-in files to the current directory

Syntax:\
moveWordAddIns

The command should be executed when Word is closed.  Add-ins will no longer be loaded when Word next opens.  Since the files are moved, not deleted, you can restore them if necessary.

## moveWordNormalTemplate.cmd = Move the Normal template of Word, Normal.dotm, to the current directory.

Syntax:\
moveWordNormalTemplate

The command should be executed when Word is closed.  The default template will be recreated by Word when next opened.Since the file is moved, not deleted, you can restore it if necessary.

## moveWordStartUp.cmd = Move Word start-up extension files to the current directory

Syntax:\
moveWordStartUp

The command should be executed when Word is closed.  start-up extensions will no longer be loaded when Word next opens.  Since the files are moved, not deleted, you can restore them if necessary.

# Miscellaneous Functionality (Misc)

Use the KeyLineHelp command to open this documentation in your default web browser.

If you use Microsoft Outlook, you can list, save, or remove attachments in most of your folders:  Calendar, Contacts, Deleted, Drafts, Inbox, Journal, Junk, Notes, Outbox, Sent, and Tasks.

## GetAttachments = Get attachments from folder

Syntax:\
GetAttachments \<Action\> \<Folder\>

The command takes three case-insensitive parameters corresponding to an action (List, Save, or Remove), a folder (one of the above), and an optional, space-separated list of file extensions, e.g., "docx pdf pptx" for Word, PDF, and PowerPoint files.  Only folder items having attachments are shown, including the name of the item, e.g., the subject of a message, and the file name of each attachment.  Output is also stored in the text file Attachments.log so you can review information that scrolls off the screen.

For convenience, commands without parameters have been defined for the Inbox, Sent, and Deleted folders, e.g., SaveInboxAttachments.cmd, RemoveSentAttachments.cmd, and ListDeletedAttachments.cmd.  The Save action saves attachments in the current directory with unique file names, adding a numeric suffix when needed.  Many of these files may be duplicates, so consider running delDupes.cmd afterward.

A few more Outlook commands follow, which may require an adjustment in Outlook security settings for macros in order to work (if so, restore the previous setting after running the command).

## saveContacts.cmd. = Save all Outlook contacts

Syntax:\
saveContacts

The command saves all Outlook contacts in the current directory in both .vcf and .txt formats.  Each file name is based on the first name, last name, and company fields.

## saveAppointments.cmd. = Save all Outlook appointments

Syntax:\
saveAppointments

The command saves all Outlook calendar appointments in the current directory in both .ics and .txt formats.  Each file name is based on the fields for subject and start time.

Several commands enable you to force common applications to close, in case they are not closing by other attempts, and their stuck state may be inhibiting KeyLine functionality.  These commands are as follows:  KillAdobe (Adobe Reader), KillCalibre, KillChrome, KillEdSharp (a [text and code editor](https://github.com/jamalmazrui/edsharp) by this author), KillExcel, KillFileDir (a [file and directory manager](https://github.com/jamalmazrui/filedir) by this author), KillFirefox, killJAWS, KillLibre (LibreOffice), killNVDA, KillOutlook, KillPandoc, KillPowerPoint, killTeams, killSlack, KillWget (used by SaveWebDir.cmd), and KillWord.

The following commands perform miscellaneous, high-level tasks.

## archive.cmd = Archive files

Syntax:\
archive \<FileName\> \<Spec\>

If not specified, the default archive extension is .zip.  The default archive file name is archive.zip.  The default spec is \*.\*.  The command uses the [7-Zip](https://www.7-zip.org/) utility.

## anagram.cmd = find anagrams from a set of letters

Syntax:\
anagram \<letters\>

The command checks every possible sequence of the letters, looking for valid words that Microsoft Word does not consider to be spelling errors.  Any such anagrams found are printed to the screen and saved in a text file.  This is useful, for example, if you are trying to find an acronym as a mnemonic for a set of words.

## BrowserHistory.cmd. = create and open file with links in browser history

Syntax:\
BrowserHistory <days>

The command creates the file BrowserHistory.htm in the current directory, using the BrowsingHistoryView.exe utility in the `C:\KeyLine\bin` subdirectory.  Histories of the following web browsers are examined:  Chrome, Edge, Firefox, Internet Explorer (IE), and Safari.  

The output file is then opened in the default web browser.  Links are in reverse chrononological order -- most recent first.  Depending on the size of the file, it may take significant time to open.  Load time can be reduced by passing the number of days as a parameter, e.g., 7 for the past week of browser history.  

The content of BrowserHistory.htm is a table with three columns, corresponding to the page title, URL, and browser.  To find a link of interest, you can down arrow through the content or search for words.  Screen reader keys for table navigation work as expected.  

To view history for a particular browser, rather than them all, specify one of the following, related commands instead:  ChromeHistory.cmd, EdgeHistory.cmd, or FirefoxHistory.cmd.  With these commands, The output table eliminates the column for browser, since it is already known.
 
## cleanText.cmd = produce Unicode file with clean encoding

Syntax:\
cleanText \<SourceFile\>

The command eliminates encoding errors that may have resulted from previous conversions, e.g., a gibberish symbol or ? character where a translation error occurred.  UTF-8 is the default encoding.  For details, see ftfy.htm in the KeyLine\help subdirectory.

## combine.cmd. = Combine HTML files into a single file with a table of contents

Syntax:\
combine \<spec\> \<TargetRootName\>

The command combines HTML files matching spec into a single HTML file with the root name given, e.g., \
combine *.htm MySubject

would produce MySubject.htm, where the content of each component file begins with a heading level 1.  The first section of MySubject.htm is a table of contents consisting of links to all level 1 and 2 headings in the file.

The combined file is automatically opened in the default web browser.  The command uses the `pandoc.exe` utility in the `C:\KeyLine\bin` subdirectory.

## delDupes.cmd = Delete duplicate files

Syntax:\
delDupes

The command examines all files in the current directory tree, that is, the current directory, its subdirectories, and any further descendent directories.  Files are only considered to be duplicates if they match byte-for-byte, regardless of file name.  No destructive action is taken without your confirmation.  

This is one of the few KeyLine commands that present a graphical user interface (GUI) dialog.  A multiselection listbox allows you to individually select which files to delete -- or the *All* button may be chosen to select all duplicates.  A subsequent dialog tells you how many files would be deleted and asks for confirmation before doing so.

The command uses the [Swiss File Knife]are pathEd .net path
	(http://stahlworks.com/dev/swiss-file-knife.html) utility.w

## delExtras.cmd = Delete extra files from a Google search

Syntax:\
delExtras

The command deletes files like `Search.html` that are found if you download all files from the results page of a Google search.

## delJunk.cmd = Delete extraneous, junk files

Syntax:\
delJunk

The command deletes graphics and other support files that are typically downloaded from a web page when other files are sought such as documents and archives.  Files considered junk include .css, .js, and .png files, as well as .html files with "Page not found" in their title.  The command is useful, for example, after running durlAll.cmd or saveWebDir.cmd.

## delNameless.cmd = Delete template files from a GitHub repository

Syntax:\
delNameless

The command deletes files with root names that do not include any alphabetic characters, e.g., `2-.htm`.

## delSimilar.cmd = Delete similar files

Syntax:\
delSimilar <spec>

The command examines files matching a wildcard specification in the current directory, or all files if such a parameter is not specified.  Files are considered to be similar if they have the same base names except for a numeric suffix that is assumed to indicate another version of the same content.  For example, content.epub and content.pdf would be considered different files because they do not have the same extension.  However, content.pdf, content-1.pdf, content_2.pdf, content(3).pdf, and content[4].pdf would all be considered similar.

In the case of similar files, the largest file, in terms of bytes, is assumed to be the one to keep, since it is likely to contain the most information.  No destructive action is taken, however, without your confirmation.  

This is one of the few KeyLine commands that present a graphical user interface (GUI) dialog.  A multiselection listbox allows you to individually select which files to delete -- or the *All* button may be chosen to select all similar files.  A subsequent dialog tells you how many files would be deleted and asks for confirmation before doing so.

## delTemplates.cmd = Delete template files from a GitHub repository

Syntax:\
delTemplates

The command deletes files like `license.md` that are found if you download and unzip an archived repository from GitHub.com.

## durl.cmd = Download files from a URL

Syntax:\
durl \<URL\>
or
durl \<Spec\>

The command shows the number of links found in the HTML source and then tries to determine the file type of each one.  By default, file types in the doc category are downloaded.  Additional parameters can specify other file types.  For example,
durl \<URL\> -t "archive data doc"
will download archive, (e.g., .zip), data (e.g., .json), and document (e.g., .pdf) file types.  durlData.cmd is a command that includes these particular parameters.  

If you encounter a web page with direct links to files of interest, e.g., .docx or .pdf files, the durl command can be an efficient way of downloading them all.  After opening the page in a web browser, press Alt+D for the addressbar and then Control+C to copy the URL to the clipboard (the URL is usually selected automatically).  Then go to a KeyLine command prompt and press Control+V to paste the URL as a parameter to durl.cmd or one of its variations (explained subsequently).  For best results, type a quote before and after the URL that you paste.  Note that the technique only works with public URLs, not those that require a log-in to reach the page (due to *cookies* or other security features).

If a file spec is given as the first parameter, instead of a URL, the command will analyze each HTML file found.

File extensions are grouped into the following categories:

- archive = 7z, cab, gz, rar, tar, tgz, z, zip
- audio = aac, dtb, m4a, mid, mp3, mp4a, oga, ogg, wav, wma
- data = cnf, csv, db, dbf, dbt, fpt, ini, json, mdb, sdf, tsv, xml, xls, xlsx, yaml
- doc = azw, azw3, brf, brl, chm, daisy, doc, docx, dotx, dtb, epub, hlp, kf8, kfx, lit, lrf, mobi, odf, pdf, ppt, pptx, rst, rtf, txt
- executable = exe, msi
- image = bmp, gif, ico, icon, jpeg, jpg, mdi, png, svg, ttf, wmf
- video = avi, flv, mkv, mp4, mpg, ogv, webm, wmv
- web = asp, aspx, cfm, css, htm, html, js, markdown, md, php, xhtml

The `-r` parameter will reverse the order of the list of links collected.  The `-m` parameter can match a regular expression against the base file name (root.ext) being considered for download.

## durlAll.cmd = Download files from a URL

Syntax:\
durlAll \<URL\>
or
durlAll \<Spec\>

The command tells durl to download files that it recognizes from all available categories:  archive, audio, data, doc, executable, image, video, and web.  A variation, durlAllReverse.cmd, processes links in reverse order.

## durlData.cmd = Download files from a URL

Syntax:\
durlData \<URL\>
or
durlData \<Spec\>

The command tells durl to download the following file types:  archive, data, and doc.  Archive files (like .zip) are included because data files are sometimes provided in a compressed, archived collection.  A variation, durlDataReverse.cmd, processes links in reverse order.

## durlExecutable.cmd = Download files from a URL

Syntax:\
durlExecutable \<URL\>
or
durlExecutable \<Spec\>

The command tells durl to download the following file types:  archive, audio, data, doc, and executable.  Besides the executable category, files from the archive data and doc categories are included, since they are likely to also be of interest if you want installers or other executable downloads.  Also, for security, an executable file is sometimes packaged within an archive for download.

## durlImage.cmd = Download files from a URL

Syntax:\
durlImage \<URL\>
or
durlImage \<Spec\>

The command tells durl to download files from the image category.

## durlVideo.cmd = Download files from a URL

Syntax:\
durlVideo \<URL\>
or
durlVideo \<Spec\>

The command tells durl to download the following file types:  archive, audio, data, doc, and video.  Since video files (like .mp4) tend to be large in size, other, notable files from the same page are also included at little additional cost from the archive data and doc categories, since they are likely to be related to a video presentation.  The audio category is also included, since if you are interested in multimedia, you would probably want such files as well.

## durlWeb.cmd = Download files from a URL

Syntax:\
durlWeb \<URL\>
or
durlWeb \<Spec\>

The command tells durl to download the following file types:  archive, data, doc, and web.  Besides the web category, files from the archive data and doc categories are included, since they are likely to also be of interest if you want HTML downloads.  For each HTML file, the command also shows a summary based on metadata extracted.

## findBooks = Find metadata about books

Syntax:\
findBooks \<Keywords\>

Keywords may be words in a book title or author name.  If multiple words, surround them with quotes.  The target .xlsx file has rows corresponding to books, and columns with metadata about them.  For details, see Metabook.md in the ``C:\KeyLine\help`` directory.

## getBookData = Get metadata about a book

Syntax:\
getBookData \<authors\> \<title\>

Pass the book author(s) as the first parameter and the title as an optional second parameter.

## getPDFDocs = Get PDF documentation for software

Syntax:\
getPDFDocs \<Package\>

The command tries to get full documentation for a software package as a .pdf from the ReadTheDocs.org website.

## isTagged.cmd = Determine if a PDF file is tagged for accessibility

Syntax:\
isTagged \<Spec\>

The default spec is \*.pdf.

## killApp.cmd = Kill an application

Syntax:\
killApp <ExecutableName>

The command forcefully closes an application specified by its executable name, without the .exe extension, e.g., iexplore for Internet Explorer.  If multiple instances of the app are loaded in memory, all are closed.

## listDrives.cmd = List disk drive letters, types, and names

Syntax:\
listDrives

The command lists all disk drives, showing the letter and type for each.  If a disk is ready for viewing, its volume name is also shown.

## getFileProperties.cmd = List properties of each file or folder within the current directory

Syntax:\
getFileProperties

The command shows all properties of each file or folder within the current directory.  The list is also saved in the text file FileProperties.txt.

## listInstalledPrograms.cmd = List all installed programs

Syntax:\
listInstalledPrograms

The command shows metadata of each program installed on the computer. The list is also saved in the text file InstalledPrograms.txt.

## listStartupCommands.cmd = List all startup commands

Syntax:\
listStartupCommands

The command shows the name and executable statement of each startup command on the computer. The list is also saved in the text file StartupCommands.txt.

## mainly.cmd = Extract the main content of an HTML file

Syntax:\
mainly \<Spec\>

The target file has the same root name and a .htm extension.  The extraction is similar to what some web browsers offer for removing extraneous material around an article on a page.

## moveUnique.cmd. = Move files in current directory tree to a single directory with unique file names

Syntax:\
moveUnique \<spec\> \<Directory\>

The command looks for files matching the spec in the current directory and all its subdirectories, which are moved to the target directory.  If a file with the same name already exists, a random string is added to the root name of a file before being moved, thereby preventing files with the same names from replacing each other.  The command may be useful after saveWebDir.cmd creates multiple levels of subdirectories.  delDupes.cmd could then be run in the target directory to remove files with identical content.  

copyUnique.cmd does the same thing except that files are copied rather than moved.

## MSAA.cmd = Show an accessibility tree for an application window

Syntax:\
msaa

The command presents a GUI dialog for picking one of the application windows that are currently open.  It extracts information using the Microsoft Active Accessibility (MSAA) API for each control found in the window.

## numberFiles.cmd = rename files to achieve a sort order

Syntax:\
numberFiles \<fileName\>

fileSpec is a text file containing the full paths of file names, one per line.  The command renames them by adding a numeric prefix to ensure an order when sorted alphabetically.  When another command subsequently operates on these files, they will be processed in that order, since alphanumeric sequence is the default.

## open.cmd = Open a file

Syntax:\
open \<Spec\>

The first file matching the spec is opened, so you can specify part of a file name if you do not know the precise spelling.  For example, suppose you have converted files from PowerPoint to HTML using the pptx2cmd command.  You know that a file of interest contains "intro" as part of its name, so you enter "open \*intro\*.htm" to view the file in the default web browser.

Variations of this command open a particular web browser with the file specified:  chrome.cmd, and firefox.cmd.  They enable you to open an HTML file with a browser other than the default.

## openDrive.cmd = Open a drive


Syntax:\
openDrive <Letter>

The command opens the drive corresponding to the letter passed as a parameter, assuming the drive holds portable media.  closeDrive.cmd command does the reverse.

## PasswordHistory.cmd. = create and open a file with saved passwords for websites

Syntax:\
PasswordHistory

The command creates the file PasswordHistory.htm in the current directory, using the WebBrowserPassView.exe utility in the `C:\KeyLine\bin` subdirectory.  Configurations are examined for as many browsers as possible.  Naturally, this command should only be invoked with keen attention to personal privacy and security.  It can help you log back into a website where you lost the credetntials and the password recovery process poses accessibillity challenges.

The output file, PasswordHistory.htm, is opened automatically in the default web browser.  It consists of a table with four columns, corresponding to the website URL, user name, password, and browser.  Records are shown in reverse chronological order (the most recently saved password is first).  To find a password of interest, you can down arrow through the content or search for words.  Screen reader keys for table navigation work as expected.    

## query.cmd = perform queries on data files

Syntax:\
query \<parameters ...\>

The command lets you perform queries in Structured Query Language (SQL) against files in .csv or .tsv tabular formats.  For details, see q.htm in the KeyLine\help subdirectory.

## Regexer.cmd = Show or change matches of regular expressions

Syntax:\
Regexer \<Spec\> \<ConfigFile\>
or
Regexer \<SourceFile\> \<ConfigFile\> \<TargetFile\>

Each section of the .ini configuration file specifies a regular expression for extracting or replacing text.  Extractions are copied to the clipboard.  If a target file is not specified, replacements occur in the source file.

## Rentitle.cmd = Rename files based on title metadata

Syntax:\
Rentitle \<Spec\>

The command uses the [Exiftool](https://exiftool.org/) utility to extract metadata from almost any file type (*EXIF* refers to Exchangeable Image File format).  

## reorder.cmd = reorder files so they sort better

Syntax:\
reorder <spec>

The command operates on all files in the current directory or to files matching a wildcard specification if one is passed.  A 0 prefix is added to file names that begin with a single digit.  Suppose 11 files began with a number in the sequence from 1 to 11, e.g., 1name.htm, 2name.htm ... 11name.htm.  When sorted alphabetically, the order would begin 1name.htm, 11name.htm, 2.htm, putting the last file after the first one instead of at the end.

After running this command, however, the intended order would result:  01name.htm, 02name.htm, ... 11.htm.  In addition, files that are typically front matter -- such as ReadMe.md -- are prefixed with 0#so they appear at the start of an alphetized list, and back matter files such as contributing.md -- are prefixed with z- so they appear at the end.

The command is useful, for example, before running combin.cmd to consolidate multiple HTML files into a single one.

## restartWindows.cmd = Restart Windows

Syntax:\
restartWindows <Seconds>

The command restarts Windows forcefully after pausing for the number of seconds passed as a parameter.  If no parameter is passed, the default is 5 seseconds.  Use 0 as a parameter for no pause.  shutdownWindows.cmd is similar except that the computer is turned off instead of restarting Windows.

## SayClipboard.cmd = Say content of the clipboard

Syntax:\
SayClipboard

The command uses the `nircmd.exe` utility in the `C:\KeyLine\bin` subdirectory.  
 
## SayFile.cmd = Say the content of a text file

Syntax:\
SayFile \<FileName\>

The command uses the `SayFile.exe` utility in the `C:\KeyLine\bin` subdirectory.

## SayLine.cmd = Say a line of text

Syntax:\
SayLine \<textLine\>

The command uses the `SayLine.exe` utility in the `C:\KeyLine\bin` subdirectory.

## saveVideoText.cmd = Save transcripts of videos

Syntax:\
saveVideoText \<URL\>

The command searches the page of a video on YouTube, Vimeo, or another source and downloads transcripts based on captions.  If manually created captions do not exist, auto-captions, of lesser quality, are usually available.  If a playlist is found on the page, its videos are checked for transcripts.

A related command, `saveVideoTextList.cmd`, takes a file name rather than a URL as a parameter. URLs to be checked for transcripts are listed, one per line.  `

`saveVideoTextReverse.cmd` and `saveVideoTextListReverse.cmd` are similar commands that process URLs in reverse order.  This helps in the situation where transcripts for videos later in the sequence are ignored because of an excessive number of requests to YouTube within a short time frame.

## saveWebDir.cmd. = Save web directory to disk

Syntax:\
SaveWebDir \<URL\>

The command takes a URL as a parameter, which may end with a path to a directory on the website, e.g., `https://www.example.com/tutorials/index.html`, which would download all files in the `tutorials` subdirectory of the site.  The command downloads files and directories that are located under the most descendant parent directory, that is, the directory before the last slash (/) character -- translated to a backslash (\) character on your Windows computer.  

You can capture a URL of interest by browsing to a relevant web page, pressing Alt+D for the addressbar, Control+C to copy, and then Control+V to paste on the command line.  Relevant parts of the site's directory structure are recreated on disk.  The command uses the `wget.exe` utility in the `C:\KeyLine\bin` subdirectory.

A variant, saveWebDirDocs.cmd, downloads only document-related file types such as .docx, .pdf, and .pptx.

## screenOff.cmd = Turn off the screen temporarily
Syntax:\
screenOff

The command makes the screen go dark, so content is not visible, without powering off the monitor.  This action is reversed by screenOn.cmd.

## stampFiles.cmd = Give files a current time stamp

Syntax:\
stampFiles

The command changes the last modified date and time of each file so that it reflects the moment when this action occurs.

## vtt2txt.cmd = Convert from vtt to plain text format

Syntax:\
vtt2txt \<Spec\>

The command converts files from the VTT format for video captions.  To make the content more readable, time stamps and other extraneous text snippets are removed, and a blank line is inserted between each line of text.  Punctuation may be lacking.

## Unarchive.cmd = Unarchive files

Syntax:\
unarchive  \<FileName\> \<Spec\>

The command unarchives \<FileName\> to the current directory, regardless of archive format (e.g., .7z, .tgz, or .zip).  The default spec is \*.\*.

# Low-Level Commands

These low-level commands are generally not needed, but advanced users might find them useful.

Consider the command pandoc2htm.cmd.  The name indicates that it uses Pandoc (pandoc.exe) to convert a source file to .htm format -- an HTML file.  The target file will have the same root name as the source and have a .htm extension.  It will be created in the current directory.  Multiple files may be converted with a single command, e.g.,

pandoc2htm \*.md

This command converts all files with a .md extension, that is, files in Markdown format.

Commands that begin with "calibre2" use a program called Calibre for conversions (ebook-convert.exe), and those that begin with "libre2" use LibreOffice (soffice.exe).  Commands that begin with "wd2" use a program called wdVert (wdVert64.exe), which perform conversions with that API of Microsoft Word.  Similarly, the "pp2" prefix indicates use of the PowerPoint API and "xl2" indicates use of the Excel API.

In general, a conversion only occurs if the target file is not present.  The batch file prints the name of each file that it attempts to convert.  If the conversion fails, the word "Error" is printed.  So, for example, if you copy additional .pdf files into a directory and then run the "wd2htm \*.pdf" command, conversions will not be redone that already succeeded.

Often, a target format may be achieved with more than one tool.  If a conversion fails with one tool, it might succeed with another.  If more than one tool succeeds, there may be differences in the quality of the output, e.g., a .htm file with heading structure as opposed to one without such structure.

Here are tips on conversions that tend to produce the best output:

- In general, the "wd2" commands produce the best output if a target format is supported, e.g., wd2pdf.cmd to produce a .pdf from a .docx file or wd2htm.cmd to produce a .htm from a .pdf file.

- libre2htm.cmd produces the best conversion of .xlsx workbooks, since the HTML file has a table for each sheet of the workbook, not just the first sheet.

- pandoc2docx.cmd produces the best .docx output from a .md file.

- calibre2htm.cmd is the only way of converting from a .chm to a .htm file.

In a few cases, a conversion is a two-step process involving a temporary, intermediate file that is automatically deleted (from the temp directory).  For example, pp_wd2htm.cmd converts a .pptx file to PDF using the PowerPoint API, and then converts from .pdf to .htm using the Word API.

The following are low-level commands upon which simple conversion commands are based.  The simple commands, e.g., pptx2htm.cmd, do not reference the underlying utility -- or utilities -- that are used for conversion.

## ansi.cmd = Convert file to ANSI encoding

Syntax:\
ansi \<SourceFile\>

The command auto-detects the current file encoding, and if not ANSI, converts to it.

## addImageTag.cmd = Add tag to image file

Syntax:\
addImageTag \<SourceFile\> \<TagName\> \<TagValue\>

The command adds the tag name and value to the image file, e.g., a .jpg or .png file.

## addMediaTag.cmd = Add tag to media file

Syntax:\
addMediaTag \<SourceFile\> \<TagName\> \<TagValue\>

The command adds the tag name and value to the media file, e.g., a .mp3 or .mp4 file.

## calibre_pandoc2htm.cmd = Convert to .htm format using Calibre, then Pandoc

Syntax:\
calibre_pandoc2htm \<Spec\>

A .docx file is used as a temporary, intermediate format.

## calibre_pandoc2html.cmd = Convert to .html format using Calibre, then Pandoc

Syntax:\
calibre_pandoc2html \<Spec\>

A .epub file is used as a temporary, intermediate format.

## calibre_wd2htm.cmd = Convert to .htm format using Calibre then Word

Syntax:\
calibre_wd2htm \<Spec\>

A .docx file is created as a temporary, intermediate format.

## calibre_wd2html.cmd = Convert to .html format using Calibre, then Word

Syntax:\
calibre_wd2html \<Spec\>

A .pdf file is created as a temporary, intermediate format.

## calibre2docx.cmd = Convert to .docx format using Calibre

Syntax:\
calibre2docx \<Spec\>

The command auto-detects the source format.

## calibre2epub.cmd = Convert to .epub format using Calibre

Syntax:\
calibre2epub \<Spec\>

The command auto-detects the source format.

## calibre2htmlz.cmd = Convert to .htmlz format using Calibre

Syntax:\
calibre2htmlz \<Spec\>

The command auto-detects the source format.  The target format is a zip archive containing .html, .css, and .jpg files.

## calibre2kfx.cmd = Convert to .kfx format using Calibre

Syntax:\
calibre2kfx \<Spec\>

The command auto-detects the source format.  The target, Kindle format may be sideloaded into a Kindle reading app.

## calibre2mobi.cmd = Convert to .mobi format using Calibre

Syntax:\
calibre2mobi \<Spec\>

The command auto-detects the source format.  The target, Kindle format may be sideloaded into a Kindle reading app.

## calibre2pdf.cmd = Convert to .pdf format using Calibre

Syntax:\
calibre2pdf \<Spec\>

The command auto-detects the source format.

## calibre2rtf.cmd = Convert to .rtf format using Calibre

Syntax:\
calibre2rtf \<Spec\>

The command auto-detects the source format.

## calibre2txt.cmd = Convert to .txt format using Calibre

Syntax:\
calibre2txt \<Spec\>

The command auto-detects the source format.

## enc.cmd = Show file encoding

Syntax:\
enc \<Spec\>

The command auto-detects the following encodings:  ansi, utf-8b, utf-8n, and utf-16.

## encoding.cmd = Show or change file encoding

Syntax:\
encoding \<Spec\>
or
encoding \<Spec\> \<TargetEncoding\>

Many encodings are possible.  For details, see Encoding.txt in the KeyLine\help subdirectory.

## HomerLib.cmd = Test the Homer script library

Syntax:\
HomerLib

This library of VBScript functions is used by other .vbs programs that are part of KeyLine.

## IniForm.cmd = Present a GUI form

Syntax:\
IniForm \<InputFile\>

The command uses .ini files for input and output.  For details, see IniForm.txt in the ``C:\KeyLine\help`` subdirectory.  For an example input .ini file, see PickFiles_input.ini in the ``C:\KeyLine\ini`` directory.

## libre2docx.cmd = Convert to .docx format using LibreOffice

Syntax:\
libre2docx \<Spec\>

The command auto-detects the source format.

## libre2htm.cmd = Convert to .htm format using LibreOffice

Syntax:\
libre2htm \<Spec\>

The command auto-detects the source format.

## libre2html.cmd = Convert to .html format using LibreOffice

Syntax:\
libre2html \<Spec\>

The command auto-detects the source format.

## libre2pdf.cmd = Convert to .pdf format using LibreOffice

Syntax:\
libre2pdf \<Spec\>

The command auto-detects the source format.

## libre2pptx.cmd = Convert to .pptx format using LibreOffice

Syntax:\
libre2pptx \<Spec\>

The command auto-detects the source format.

## libre2rtf.cmd = Convert to .rtf format using LibreOffice

Syntax:\
libre2rtf \<Spec\>

The command auto-detects the source format.

## libre2txt.cmd = Convert to .txt format using LibreOffice

Syntax:\
libre2txt \<Spec\>

The command auto-detects the source format.

## libre2xlsx.cmd = Convert to .xlsx format using LibreOffice

Syntax:\
libre2xlsx \<Spec\>

The command auto-detects the source format.

## metabook.cmd = Show metadata about books using the Goodreads API

Syntax:\
metabook \<Param1\> \<Param2\> ...

The target .xlsx file has rows corresponding to books, and columns with metadata about them.  For details, see Metabook.htm in the ``C:\KeyLine\help`` directory.

## pandoc2docx.cmd = Convert to .docx format using Pandoc

Syntax:\
pandoc2docx \<Spec\>

The command auto-detects the source format.

## pandoc2epub.cmd = Convert to .epub format using Pandoc

Syntax:\
pandoc2epub \<Spec\>

The command auto-detects the source format.

## pandoc2epub2.cmd = Convert to .epub V2 format using Pandoc

Syntax:\
pandoc2epub2 \<Spec\>

The command auto-detects the source format.

## pandoc2epub3.cmd = Convert to .epub V3 format using Pandoc

Syntax:\
pandoc2epub3 \<Spec\>

The command auto-detects the source format.

## pandoc2htm.cmd = Convert to .htm format using Pandoc

Syntax:\
pandoc2htm \<Spec\>

The command auto-detects the source format.  A variant, Pandoc_toc2htm.cmd, prepends a table of contents, based on heading levels 1 and 2.

## pandoc2html.cmd = Convert to .html format using Pandoc

Syntax:\
pandoc2html \<Spec\>

The command auto-detects the source format.  A variant, Pandoc_toc2html.cmd, prepends a table of contents, based on heading levels 1 and 2.


## pandoc2md.cmd = Convert to .md format using Pandoc

Syntax:\
pandoc2md \<Spec\>

The command auto-detects the source format.

## pandoc2pptx.cmd = Convert to .pptx format using Pandoc

Syntax:\
pandoc2pptx \<Spec\>

The command auto-detects the source format.

## pandoc2rtf.cmd = Convert to .rtf format using Pandoc

Syntax:\
pandoc2rtf \<Spec\>

The command auto-detects the source format.

## pandoc2txt.cmd = Convert to .txt format using Pandoc

Syntax:\
pandoc2txt \<Spec\>

The command auto-detects the source format.

## phoneNumber.cmd = Convert letters in a phone number to digits

Syntax:\
phoneNumber \<alphanumeric sequence\>

Pass a phone number with letters representing digits as a mnemonic, e.g., 1-800-flowers, and the conversion will be output, e.g., 1-800-3569377.

## pp_calibre2htm.cmd = Convert to .htmlz format using PowerPoint, then Calibre

Syntax:\
pp_calibre2htm \<Spec\>

A .pdf file is used as a temporary, intermediate format.

## pp_wd2htm.cmd = Convert to .htm format using PowerPoint, then Word

Syntax:\
pp_wd2htm \<Spec\>

A .rtf file is used as a temporary, intermediate format.

## pp_wd2html.cmd = Convert to .html format using PowerPoint, then Word

Syntax:\
pp_wd2html \<Spec\>

A .pdf file is used as a temporary, intermediate format.

## pp2pdf.cmd = Convert to .pdf format using PowerPoint

Syntax:\
pp2pdf \<Spec\>

The command auto-detects the source format.

## pp2ppt.cmd = Convert to .ppt format using PowerPoint

Syntax:\
pp2ppt \<Spec\>

The command auto-detects the source format.

## pp2pptx.cmd = Convert to .pptx format using PowerPoint

Syntax:\
pp2pptx \<Spec\>

The command auto-detects the source format.

## pp2rtf.cmd = Convert to .rtf format using PowerPoint

Syntax:\
pp2rtf \<Spec\>

The command auto-detects the source format.

## pp2txt.cmd = Convert to .txt format using PowerPoint

Syntax:\
pp2txt \<Spec\>

The command auto-detects the source format.

## utf16.cmd = Convert file to UTF-16 encoding

Syntax:\
utf16 \<SourceFile\>

The command auto-detects the current file encoding, and if not UTF-16, converts to it.  The encoding has an initial *byte order mark* (BOM).

## utf8b.cmd = Convert file to UTF-8B encoding

Syntax:\
utf8b \<SourceFile\>

The command auto-detects the current file encoding, and if not UTF-8B, converts to it.  This encoding is *UTF-8* with an initial *byte order mark* (BOM).  KeyLine makes regular use of the encoding as the most reliable one on Windows.  It supports all Unicode characters, identifies a file encoding unambiguously by an initial sequence of bytes, and is generally understood by Linux-based software, which defaults to UTF-8 encoding without a BOM.

## utf8n.cmd = Convert file to UTF-8N encoding

Syntax:\
utf8n \<SourceFile\>

The command auto-detects the current file encoding, and if not UTF-8N, converts to it.  The encoding is UTF-8 with no initial *byte order mark* (BOM).

## wd2doc.cmd = Convert to .doc format using Word

Syntax:\
wd2doc \<Spec\>

The command auto-detects the source format.

## wd2docx.cmd = Convert to .docx format using Word

Syntax:\
wd2docx \<Spec\>

The command auto-detects the source format.

## wd2htm.cmd = Convert to .htm format using Word

Syntax:\
wd2htm \<Spec\>

The command auto-detects the source format.

## wd2html.cmd = Convert to .html format using Word

Syntax:\
wd2html \<Spec\>

The command auto-detects the source format.

## wd2pdf.cmd = Convert to .pdf format using Word

Syntax:\
wd2pdf \<Spec\>

The command auto-detects the source format.

## wd2rtf.cmd = Convert to .rtf format using Word

Syntax:\
wd2rtf \<Spec\>

The command auto-detects the source format.

## wd2txt.cmd = Convert to .txt format using Word

Syntax:\
wd2txt \<Spec\>

The command auto-detects the source format.

## xl2csv.cmd = Convert to .csv format using Excel

Syntax:\
xl2csv \<Spec\>

The command auto-detects the source format.

## xl2htm.cmd = Convert to .htm format using Excel

Syntax:\
xl2htm \<Spec\>

The command auto-detects the source format.

## xl2html.cmd = Convert to .html format using Excel

Syntax:\
xl2html \<Spec\>

The command auto-detects the source format.

## xl2txt.cmd = Convert to .txt format using Excel

Syntax:\
xl2txt \<Spec\>

The command auto-detects the source format.

## xl2xls.cmd = Convert to .xls format using Excel

Syntax:\
xl2xls \<Spec\>

The command auto-detects the source format.

## xl2xlsx.cmd = Convert to .xlsx format using Excel

Syntax:\
xl2xlsx \<Spec\>

The command auto-detects the source format.

## xlCalc.cmd = Calculate an expression with Excel

Syntax:\
xlCalc \<Expression\>

Pass a quoted expression containing an Excel function or formula, and the result of its evaluation will be output to the console.

## xlFormat.cmd = Convert to auto-formatted Excel tables

Syntax:\
xlFormat \<Spec\>

The source file is converted to a target .xlsx format with table cells in all worksheets auto-sized, word-wrapped, and top-aligned. Empty columns are also dropped.  A visual user of the spreadsheet may still need to tweak column settings for optimum readability, but such effort is intended to be minimized.

## xlHeaders.cmd = Define names for column headers and row labels, so a screen reader identifies them when navigating an Excel a table

Syntax:\
xlHeaders \<File\>

The .xlsx file passed as a parameter will be examined for uniform data tables. Column headers and row labels will be named with "ColumnTitle, "RowTitle, or "Title" as appropriate.

## xlStruct.cmd = Describe the structure of an Excel workbook

Syntax:\
xlStruct \<File\>

The .xlsx file passed as a parameter will be described, including structural information about its sheets and data regions.

## jsCalc.cmd = Calculate an expression with jsScript

Syntax:\
jsCalc \<Expression\>

Pass a quoted expression in the jScript language (an earlier, Microsoft version of JavaScript), and the result of its evaluation will be output to the console.

## vbCalc.cmd = Calculate an expression with VBScript

Syntax:\
vbCalc \<Expression\>

Pass a quoted expression in the VBScript language, and the result of its evaluation will be output to the console.

## docxFormat.cmd = Convert to auto-formatted Word tables

Syntax:\
docxFormat \<Spec\>
The source file is converted to a target .docx format with all tables auto-formatted.

# docxIt Command

The docxIt command requires a DOCX file as the first parameter and an INI file as the second parameter.  It can perform many different tasks that transform the content or format of the document, e.g., generating a table of contents and performing search and replace operations.  For an example .ini file, see AccAuthor-docxIt.ini in the `C:\KeyLine\ini` subdirectory.

Two possible section names in the INI file have special meaning.  The [Tasks] section can include various, miscellaneous tasks, such as removing personal information from metadata of the document.  The [TOC] section enables a table of contents to be automatically generated.  All other section names of the INI file denote search and replace operations using this powerful capability of Word.

The following settings are possible in the [Tasks] section of a .ini file:

- AcceptAllRevisions = Accepts all tracked changes in the specified document.

- ApplyQuickStyleSet
- AttachedTemplate = Returns a Template object that represents the template attached to the specified document. Read/write Variant
- AutoFormatDocument = Automatically formats a document.
- ClearCharacterDirectFormatting
- ClearParagraphDirectFormatting
- CopyStylesFromTemplate = Copies all styles from the attached template into the document, overwriting like styles and adding unique template styles.
- DeleteAllComments = Deletes all comments from the Comments collection in a document.
- DeleteUnusedStyles
- FitToPages = Decreases the font size of text just enough so that the document will fit on one fewer pages.
- PrintOut = Prints all or part of the specified document.
- PrintRevisions = True if revision marks are printed with the document. False if revision marks aren't printed (that is, tracked changes are printed as if they'd been accepted). Read/write Boolean.
- RejectAllRevisions = Rejects all tracked changes in the specified document.
- RemoveDateAndTime = Sets or returns a Boolean indicating whether a document stores the date and time metadata for tracked changes.
- RemoveNumbers = Removes numbers or bullets from the specified document.
- Repaginate = Repaginates the entire document.
- SaveAsQuickStyleSet
- SetTableHeaders = ApplyStyleHeadingRows, True for Microsoft Word to apply heading-row formatting to the first row of the selected table. Read/write Boolean.
- UnprotectDocument = Removes protection from the specified document.
- UpgradeFormat

The following settings are possible in the TOC] section of a .ini file:

- HeadingStyles
- HidePageNumbersInWeb
- IncludePageNumbers
- LowerHeadingLevel
- RightAlignPageNumbers
- UpperHeadingLevel
- UseFields
- UseHeadingStyles
- UseHyperlinks
- UseOutlineLevels

The following settings are possible in all other sections of a .ini file, which are interpreted as find-and-replace operations:

- FindText = text to find.
- MatchCase = whether to match case.
- MatchWholeWord = whether to match only entire words.
- MatchWildcards = whether the find text can include wildcards.
- MatchSoundsLike = whether to match words that sound similar.
- MatchAllWordForms = whether to match all word forms.
- Forward = whether to search forward.
- FindStyle = style to find.
- ReplaceStyle = style for replacement.
- Wrap = how to act if the search did not begin at the start of the range.
- Format = whether to match formatting.
- ReplaceText = text to replace with.
- Replace = how many replacements to make.

# docxProperties Command

The docxProperties command lets you manage Builtin, Custom, or Miscellaneous properties of a .docx file.  The first parameter is the name of the DOCX file to open and optionally change.  If only a single parameter is passed, the result is a file with the same root name but a "PROPERTIES-" prefix and a ".ini" extension, e.g., test.docx produces PROPERTIES-test.ini.  If a second parameter is passed to docxProperties, it should be the name of an INI file to change properties in the document.    

The following settings are possible in the [Builtin] section of a .ini file:

- Title
- Author
- Comments
- Keywords
- Category

The following settings are possible in the [Misc] section of a .ini file:

- ActiveWritingStyle = Returns or sets the writing style for a specified language in the specified document. Read/write String.
- ApplyQuickStyleSet = Applies the specified StyleSet to the document.
- Compatibility = True if the compatibility option specified by the Type argument is enabled. Compatibility options affect how a document is displayed in Microsoft Word. Read/write Boolean.
- DefaultTabStop = Returns or sets the interval (in points) between the default tab stops in the specified document. Read/write Single.
- EmbedTrueTypeFonts = True if Microsoft Word embeds TrueType fonts in a document when it is saved. Read/write Boolean.
- JustificationMode = Returns or sets the character spacing adjustment for the specified document. Read/write WdJustificationMode.
- PrintPreview = Switches the view to print preview.
- RemoveDateAndTime = Sets or returns a Boolean indicating whether a document stores the date and time metadata for tracked changes.
- RemovePersonalInformation = True if Microsoft Word removes all user information from comments, revisions, and the Properties dialog box upon saving a document. Read/write Boolean.
- SaveSubsetFonts = True if Microsoft Word saves a subset of the embedded TrueType fonts with the document. Read/write Boolean.
- ShowGrammaticalErrors = True if grammatical errors are marked by a wavy green line in the specified document. Read/write Boolean.
- StyleSortMethod = Returns or sets aWdStyleSort constant that represents the sort method to use when sorting styles in the Styles task pane. Read/write.
- TextLineEnding = Returns or sets a WdLineEndingType constant indicating how Microsoft Word marks the line and paragraph breaks in documents saved as text files. Read/write.
- TrackMoves = Returns or sets a Boolean that represents whether to mark moved text when Track Changes is turned on. Read/write.
- UpdateStylesOnOpen = True if the styles in the specified document are updated to match the styles in the attached template each time the document is opened. Read/write Boolean.

# docxStyles Command

The docxStyles command lets you manage paragraph and character styles in a DOCX file.  The first parameter is the name of the DOCX file to open and optionally change.  If only a single parameter is passed, the result is a file with the same root name but a "STYLES-" prefix and a ".ini" extension, e.g., test.docx produces STYLES-test.ini.  If a second parameter is passed to docxStyles, it should be the name of an INI file to change styles in the document.

Each section of the source .ini file is the name of a style, e.g., "Heading 1]" for the highest level heading style.  Within each style section, pars of keys and values, separated by an equals sign, set properties of the style, e.g., "BuiltIn = True" indicates that the style is built into Word rather than a custom style (that property is read-only).  For an example .ini file, see AccAuthor-docxStyles.ini in the ``C:\KeyLine\ini`` subdirectory.

If a style name is not found in the document, docxStyles creates it with the properties given.  If the style name exists but no properties are given, docxStyles assumes that style should be deleted (only custom styles can be deleted, not built-in Word styles).

The following settings for paragraph and character styles are possible in a .ini file:

- AutomaticallyUpdate = True if the style is automatically redefined based on the selection. Read/write Boolean.
- BaseStyle = Returns or sets an existing style on which you can base the formatting of another style. Read/write Variant.
- BuiltIn = True if the specified object is one of the built-in styles or caption labels in Word. Read-only Boolean.
- Font = Returns or sets a Font object that represents the character formatting of the specified style. Read/write Font.
- InUse = True if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. Read-only Boolean.
- Linked = Returns or sets a Boolean that represents whether a style is a linked style that can be used for both paragraph and character formatting. Read-only.
- ListLevelNumber = Returns the list level for the specified style. Read-only Long.
- Locked = True if a style cannot be changed or edited. Read/write Boolean.
- NameLocal = Returns the name of a built-in style in the language of the user. Read/write String.
- NextParagraphStyle = Returns or sets the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style. Read/write Variant.
- NoSpaceBetweenParagraphsOfSameStyle = True for Microsoft Word to remove spacing between paragraphs that are formatted using the same style. Read/write Boolean.
- ParagraphFormat = Returns or sets a ParagraphFormat object that represents the paragraph settings for the specified style. Read/write.
- QuickStyle = Returns or sets a Boolean that represents whether the style corresponds to an available quick style. Read/write.
- Type = Returns the style type. Read-only WdStyleType.

; Paragraph Format Settings

- Alignment = Returns or sets a WdParagraphAlignment constant that represents the alignment for the specified paragraphs. Read/write.
- FirstLineIndent = Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write Single.
- KeepTogether = True if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. Read/write Long.
- KeepWithNext = True if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. Read/write Long.
- LeftIndent = Returns or sets a Single that represents the left indent value (in points) for the specified paragraph formatting. Read/write.
- LineSpacing = Returns or sets the line spacing (in points) for the specified paragraphs. Read/write Single.
- LineSpacingRule = Returns or sets the line spacing for the specified paragraph formatting. Read/write WdLineSpacing.
- OutlineLevel = Returns or sets the outline level for the specified paragraphs. Read/write WdOutlineLevel.
- PageBreakBefore = True if a page break is forced before the specified paragraphs. Can be True, False, or wdUndefined. Read/write Long.
- RightIndent = Returns or sets the right indent (in points) for the specified paragraphs. Read/write Single.
- SpaceAfter = Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write Single.
- SpaceBefore = Returns or sets the spacing (in points) before the specified paragraphs. Read/write Single.
- WidowControl = True if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be True, False or wdUndefined. Read/write Long.

; Font Settings

- Bold = True if the font is formatted as bold. Read/write Long.
- ColorIndex = Returns or sets a WdColorIndex constant that represents the color for the specified font. Read/write .
- Hidden = True if the font is formatted as hidden text. Read/write Long.
- Italic = True if the font or range is formatted as italic. Read/write Long.
- Size = Returns or sets the font size, in points. Read/write Single.
- StrikeThrough = True if the font is formatted as StrikeThrough text. Read/write Long.
- Subscript = True if the font is formatted as subscript. Read/write Long.
- Superscript = True if the font is formatted as superscript. Read/write Long.

# wordKeys Command

The wordKeys command lets you show or change keyboard shortcuts in Microsoft Word.  For an example .ini file, see AccAuthor-wordKeys.ini in the ``C:\KeyLine\ini`` subdirectory.

# wordOptions Command

The wordOptions command lets you manage options of the Microsoft Word application, itself, rather than a particular document.  With no parameters passed, it produces OPTIONS.ini as output of existing Word options and their values (which may also be found in the Options dialog of Word).  With a single parameter, wordOptions configures the application according to options in that INI file.  For an example .ini file, see AccAuthor-wordOptions.ini in the `C:\KeyLine\Ini` subdirectory.

The following settings are possible in the [Options] section of a .ini file:

- DefaultEPostageApp = String for the path and file name of the default electronic postage application.
- DefaultFilePath = String for default folders for items such as documents, templates, and graphics.
- MarginAlignmentGuides = True if margin alignment guides are displayed in the user interface.
- MatchFuzzyCase = True if Word ignores the distinction between uppercase and lowercase letters during a search.
- MatchFuzzyDash = True if Word ignores the distinction between minus signs, long vowel sounds, and dashes during a search.
- MatchFuzzyPunctuation = True if Word ignores the distinction between types of punctuation marks during a search.
- MatchFuzzySpace = True if Word ignores the distinction between space markers used during a search.
- PageAlignmentGuides = True if page alignment guides are displayed in the user interface.
- ParagraphAlignmentGuides = True if paragraph alignment guides are displayed in the user interface.

- AddControlCharacters
- AddHebDoubleQuote
- AlertIfNotDefault
- AllowAccentedUppercase
- AllowClickAndTypeMouse
- AllowCombinedAuxiliaryForms
- AllowCompoundNounProcessing
- AllowDragAndDrop
- AllowOpenInDraftView
- AllowPixelUnits
- AllowReadingMode
- AnimateScreenMovements
- ArabicNumeral
- AutoCreateNewDrawings
- AutoFormatApplyBulletedLists
- AutoFormatApplyFirstIndents
- AutoFormatApplyHeadings
- AutoFormatApplyLists
- AutoFormatApplyOtherParas
- AutoFormatAsYouTypeApplyBorders
- AutoFormatAsYouTypeApplyBulletedLists
- AutoFormatAsYouTypeApplyClosings
- AutoFormatAsYouTypeApplyDates
- AutoFormatAsYouTypeApplyFirstIndents
- AutoFormatAsYouTypeApplyHeadings
- AutoFormatAsYouTypeApplyNumberedLists
- AutoFormatAsYouTypeApplyTables
- AutoFormatAsYouTypeAutoLetterWizard
- AutoFormatAsYouTypeDefineStyles
- AutoFormatAsYouTypeDeleteAutoSpaces
- AutoFormatAsYouTypeFormatListItemBeginning
- AutoFormatAsYouTypeInsertClosings
- AutoFormatAsYouTypeInsertOvers
- AutoFormatAsYouTypeMatchParentheses
- AutoFormatAsYouTypeReplaceFractions
- AutoFormatAsYouTypeReplaceHyperlinks
- AutoFormatAsYouTypeReplaceOrdinals
- AutoFormatAsYouTypeReplacePlainTextEmphasis
- AutoFormatAsYouTypeReplaceQuotes
- AutoFormatAsYouTypeReplaceSymbols
- AutoFormatDeleteAutoSpaces
- AutoFormatMatchParentheses
- AutoFormatPlainTextWordMail
- AutoFormatPreserveStyles
- AutoFormatReplaceFractions
- AutoFormatReplaceHyperlinks
- AutoFormatReplaceOrdinals
- AutoFormatReplacePlainTextEmphasis
- AutoFormatReplaceQuotes
- AutoFormatReplaceSymbols
- AutoWordSelection
- BackgroundSave
- BibliographySort
- BibliographyStyle
- ButtonFieldClicks
- CheckGrammarAsYouType
- CheckGrammarWithSpelling
- CheckSpellingAsYouType
- CloudSignInOption
- CommentsColor
- ConfirmConversions
- ContextualSpeller
- CreateBackup
- CtrlClickHyperlinkToOpen
- CursorMovement
- DefaultBorderColor
- DefaultBorderColorIndex
- DefaultBorderLineStyle
- DefaultBorderLineWidth
- DefaultEPostageApp
- DefaultFilePath
- DefaultHighlightColorIndex
- DefaultOpenFormat
- DefaultTextEncoding
- DefaultTray
- DefaultTrayID
- DeletedCellColor
- DeletedTextColor
- DeletedTextMark
- DisplayAlignmentGuides
- DisplayGridLines
- DisplayPasteOptions
- DocumentViewDirection
- DoNotPromptForConvert
- EnableLiveDrag
- EnableLivePreview
- EnableMisusedWordsDictionary
- EnableProofingToolsAdvertisement
- EnableSound
- EnvelopeFeederInstalled
- ExpandHeadingsOnOpen
- FormatScanning
- GridDistanceHorizontal
- GridDistanceVertical
- GridOriginHorizontal
- GridOriginVertical
- IgnoreInternetAndFileAddresses
- IgnoreMixedDigits
- IgnoreUppercase
- InlineConversion
- InsertedCellColor
- InsertedTextColor
- InsertedTextMark
- INSKeyForOvertype
- INSKeyForPaste
- InterpretHighAnsi
- LocalNetworkFile
- MapPaperSize
- MarginAlignmentGuides
- MatchFuzzyCase
- MatchFuzzyDash
- MatchFuzzyPunctuation
- MatchFuzzySpace
- MeasurementUnit
- MergedCellColor
- MonthNames
- MoveFromTextColor
- MoveFromTextMark
- MoveToTextColor
- MoveToTextMark
- MultipleWordConversionsMode
- Overtype
- PageAlignmentGuides
- Pagination
- ParagraphAlignmentGuides
- PasteAdjustParagraphSpacing
- PasteAdjustTableFormatting
- PasteAdjustWordSpacing
- PasteFormatBetweenDocuments
- PasteFormatBetweenStyledDocuments
- PasteFormatFromExternalSource
- PasteFormatWithinDocument
- PasteMergeFromPPT
- PasteMergeFromXL
- PasteMergeLists
- PasteOptionKeepBulletsAndNumbers
- PasteSmartCutPaste
- PasteSmartStyleBehavior
- PictureEditor
- PictureWrapType
- PrecisePositioning
- PreferCloudSaveLocations
- PrintBackground
- PrintBackgrounds
- PrintComments
- PrintDraft
- PrintDrawingObjects
- PrintEvenPagesInAscendingOrder
- PrintFieldCodes
- PrintHiddenText
- PrintOddPagesInAscendingOrder
- PrintProperties
- PrintReverse
- PrintXMLTag
- PromptUpdateStyle
- RepeatWord
- ReplaceSelection
- RevisedLinesColor
- RevisedLinesMark
- RevisedPropertiesColor
- RevisedPropertiesMark
- RevisionsBalloonPrintOrientation
- SaveInterval
- SaveNormalPrompt
- SavePropertiesPrompt
- SendMailAttach
- SequenceCheck
- ShowControlCharacters
- ShowDevTools
- ShowDiacritics
- ShowFormatError
- ShowMarkupOpenSave
- ShowMenuFloaties
- ShowReadabilityStatistics
- ShowSelectionFloaties
- SmartCursoring
- SmartCutPaste
- SmartParaSelection
- SnapToGrid
- SnapToShapes
- SplitCellColor
- StoreRSIDOnSave
- SuggestFromMainDictionaryOnly
- SuggestSpellingCorrections
- TabIndentKey
- TypeNReplace
- UpdateFieldsAtPrint
- UpdateFieldsWithTrackedChangesAtPrint
- UpdateLinksAtOpen
- UpdateLinksAtPrint
- UpdateStyleListBehavior
- UseCharacterUnit
- UseLocalUserInfo
- UseNormalStyleForList
- UseSubPixelPositioning
- VisualSelection
- WarnBeforeSavingPrintingSendingMarkup

The following settings are possible in the [Options] section of a .ini file:

- ActivePrinter
- AddBiDirectionalMarksWhenSavingTextFile
- AddControlCharacters
- AlertIfNotDefault
- AllowClickAndTypeMouse
- AllowCompoundNounProcessing
- AllowOpenInDraftView
- AllowReadingMode
- ArabicNumeral
- AutoFormatApplyBulletedLists
- AutoFormatApplyHeadings
- AutoFormatApplyOtherParas
- AutoFormatAsYouTypeApplyBulletedLists
- AutoFormatAsYouTypeApplyDates
- AutoFormatAsYouTypeApplyHeadings
- AutoFormatAsYouTypeApplyTables
- AutoFormatAsYouTypeDefineStyles
- AutoFormatAsYouTypeFormatListItemBeginning
- AutoFormatAsYouTypeInsertOvers
- AutoFormatAsYouTypeReplaceHyperlinks
- AutoFormatAsYouTypeReplacePlainTextEmphasis
- AutoFormatAsYouTypeReplaceSymbols
- AutoFormatMatchParentheses
- AutoFormatPreserveStyles
- AutoFormatReplaceFractions
- AutoFormatReplaceOrdinals
- AutoFormatReplaceQuotes
- BackgroundSave
- BibliographyStyle
- ButtonFieldClicks
- CheckGrammarWithSpelling
- CheckLanguage
- CheckSpellingAsYouType
- CommentsColor
- ContextualSpeller
- CreateBackup
- CursorMovement
- DefaultBorderColorIndex
- DefaultBorderLineWidth
- DefaultFilePath
- DefaultOpenFormat
- DefaultTray
- DefaultWebOptions
- DeletedCellColor
- DeletedTextMark
- DisplayAlignmentGuides
- DisplayPasteOptions
- DoNotPromptForConvert
- EmailOptions
- EnableLivePreview
- EnableProofingToolsAdvertisement
- EnvelopeFeederInstalled
- FileValidation
- FormatScanning
- GridDistanceHorizontal
- GridOriginHorizontal
- IgnoreMixedDigits
- InsertedCellColor
- InsertedTextMark
- INSKeyForPaste
- International
- LocalNetworkFile
- MarginAlignmentGuides
- MatchFuzzyDash
- MatchFuzzyPunctuation
- MatchFuzzySpace
- MergedCellColor
- MoveFromTextColor
- MoveToTextColor
- MultipleWordConversionsMode
- OpenAttachmentsInFullScreen
- Overtype
- Pagination
- PasteAdjustParagraphSpacing
- PasteAdjustWordSpacing
- PasteFormatBetweenStyledDocuments
- PasteFormatWithinDocument
- PasteMergeFromXL
- PasteOptionKeepBulletsAndNumbers
- PasteSmartStyleBehavior
- PictureWrapType
- PrecisePositioning
- PrintBackground
- PrintComments
- PrintDrawingObjects
- PrintFieldCodes
- PrintOddPagesInAscendingOrder
- PrintReverse
- PromptUpdateStyle
- ReplaceSelection
- RestrictLinkedStyles
- RevisedLinesMark
- RevisedPropertiesMark
- SaveInterval
- SavePropertiesPrompt
- SequenceCheck
- SetDefaultTheme
- ShowDevTools
- ShowFormatError
- ShowMenuFloaties
- ShowSelectionFloaties
- SmartCutPaste
- SnapToGrid
- StoreRSIDOnSave
- SuggestSpellingCorrections
- TypeNReplace
- UpdateFieldsWithTrackedChangesAtPrint
- UpdateLinksAtPrint
- UseCharacterUnit
- UseNormalStyleForList
- UserAddress
- UserName
- VisualSelection

# Appendix A:  INI File Format

Although not officially defined as a file format, there has been general consensus about features of the [.ini file format](https://en.wikipedia.org/wiki/INI_file), typically used to store configuration data for an application.  In programming languages, a *dictionary* is a set of key-value pairs, e.g., a set of terms and corresponding definitions.  A .ini file is a dictionary of dictionaries.  The top-level keys are names of sections in the file, each of which occurs on a line by itself surrounded by square brackets, e.g.,

```
[Section 1]
```

The value of each section key is a dictionary, itself.  It contains a set of text lines that each have a key and value separated by an equals sign, e.g.,

```
[Section 1]
key1 = value1
key2 = value2
```

Space before or after  a section, key, or value is not significant.  Thus, a key-value pair can also be written as follows:

```
key1 = value1
```

If you want a value to include a leading or trailing space, enclose it in quote marks.  An outer pair of quotes is ignored, e.g.,

```
key1= " value1 "
```

You can turn any line of the file into a comment, which is ignored, by prefixing it with a semicolon character (;)e.g.,

```
;key1=value1
```

# Appendix B:  INIX File Format

Note that this is an advanced topic.  For example .inix files and their target conversions, see the ``C:\KeyLine\eg`` subdirectory.

The .inix format (short for .ini extended) extends the capability of storing a dictionary of dictionaries, and it adds the capability of storing a list of dictionaries.  This enables a .inix file to be converted to a table of data in .csv, .docx, .html, or .xlsx file format.  KeyLine auto-formats  tables, as much as possible, so that rows and columns fit the data, and so that screen readers identify column headers when navigating in Excel, Word, or a web browser.

A string value can occupy multiple lines, not just a single line, by making the equals sign after the key as the last character on that line, and then starting the value on the next line, e.g.,

```
key1=
line1 of value1
line2 of value2
line3 of value3
```

A line of such a value can contain any character except for an equals sign, since this would be interpreted as the next key in the section.  A semicolon may also be used to comment out a line of the value, e.g.,

```
key1=
line1 of value1
;line2 of value1
line3 of value1
```

Not only can a single line be commented out, a whole section can be commented out by placing a semicolon after the left bracket on the line of the section name, e.g.,

```
[;Section1]
```

The first section of the file can omit its name, which defaults to to "[Global]" as the implied name.  The Global section may be useful for specifying overall configuration settings that apply regardless of individual section.  For example, if each section defines a search-and-replace operation, the Global section might include a setting that specifies case insensitivity for the subsequent searches, e.g.,

```
IgnoreCase = True

[Replace dog with cat]
Find=dog
replace=Cat

[Replace sheep with goat]
find=Sheep
replace=goat
```

A .inix file is also extended to be able to store a list of dictionaries.  Here, each section lacks a name.  Only its position is significant in the sequence of sections, e.g.,

```
[]
key1=value1
key2=value2

[]
key1=value3
key2=
line1 of value4
line2 of value4
```

This list capability allows a .inix file to store tabular data.  Column headers are assumed to occupy the first row of the table, where each header corresponds to a key in a section.  Each subsequent row of the table corresponds to a section with a value for each key.  In the above example, the column headers are key1 and key2.  The first row contains value 1 and value2.  The second row contains value3 and a multiline value4.

KeyLine can convert a .inix file to a table in another format, including .csv, .docx, .htm, and .xlsx.  Thus, for example, you could create a Word table in a .docx file and then copy and paste the table into another Word document.

# Appendix C:  MDX File Format

Note that this is an advanced topic.  For an example .mdx file, see example.mdx in the ``C:\KeyLine\eg`` subdirectory.  For a Markdown tutorial, see `The Markdown Guide` in the `C:\KeyLine\help` subdirectory.

KeyLine extends the [Markdown](https://en.wikipedia.org/wiki/Markdown) format, .md, with enhancements for convenient application of styles in a target Microsoft Word .docx file.  Suppose you want to apply the *Book Title* style.  One Markdown syntax follows:

```
<div custom-style="Book Title">This text is in Book Title style.</div>
```

Alternate syntax in a .mdx file follows:

```
::: Book Title ::: This text is in Book Title style.
```

The styled text could be multiple sentences that wrap over multiple lines before a hard return is encountered.

Another Markdown syntax is useful for applying a style to multiple paragraphs, e.g.,

```
::: custom-style="Abstract"}
This is paragraph1 of an abstract.  Here is sentence 2.

This is paragraph2 of the abstract.  Its second sentence is here.
:::
```

The cleaner, .mdx equivalent follows:

```
::: Abstract
This is paragraph1 of an abstract.  Here is sentence 2.

This is paragraph2 of the abstract.  Its second sentence is here.
:::
```

In .md format, a comment uses HTML syntax, e.g.,

```
<!-- This is a comment. -->
```

The .mdx equivalent uses the semicolon character instead, like the .ini convention, e.g.,

```
; This is a comment.
```

The KeyLine command mdx2md.cmd will convert from a .mdx to a .md file.  The .md file can be converted to various formats including .htm and .docx.  A shortcut command, mdx2docx.cmd, does the intermediate conversion for you.



