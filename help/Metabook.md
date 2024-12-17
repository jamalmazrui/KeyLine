---
title: Metabook 0.8
author: Jamal Mazrui
date: January 30, 2019
---

# Description

Metabook is a utility to get metadata about books using the [Goodreads](http://goodreads.com) API.  Search results are saved in an Excel spreadsheet, which is opened automatically by default.  The name of the Excel file indicates the type of search that was performed.  If a potential file name already exists, a numeric suffix is added to create a new file.

Search results are not instantaneous because the Goodreads API requires a delay of a second between calls.  After each call that finds a book matching search criteria, its title, author, rating, and format are printed to the console screen.  If no books matched the search criteria, an output file is not created.

If Metabook is run without options on the command line, it opens a dialog window for input instead.  All search options may be entered in this way except for specifying a file containing ISBN or book id codes.  The dialog input values are remembered the next time it is invoked.

# Installation

Unzip the Metabook archive into any directory.  It may be a new directory or one containing other utilities.  metabook.exe is the only file required for running the program.  If the dialog mode is used for input, its most recent values will be stored in the file metabook.ini, created in the same directory.

# Operation

At a command prompt, enter metabook.exe (or just metabook), followed by one or more pairs of options and values that specify search criteria for books.  Most options may be specified by either a single dash followed by a single character or a double dash followed by a word.  Here are the possible options:

* -a or --author <name> = Search by author name.

* --author-id <id> = Search by author id.

* -b or --book-id-file <file name> = Search by a book id list in a text file, one per line.

* -d or --description <text> = Filter by text in description, case-insensitive.

* --dir <directory name> = Directory for results spreadsheet.

* -f or --format <text> = Filter by format, space-separated, case-insensitive.  An example is "audio paperback".  Use "." to match a blank value for the book format.

* -h or --help = show help message and exit.

* -i or --isbn-file <file name> = Search by an ISBN list in a text file, one per line.

* -l or --language <text> = Filter by language, space-separated, case-insensitive.  An example is "en eng en-us en-uk".  Use "." to match a blank value for the book language.

* -o or --order <column names> = Column names for the sort order of search results, space-separated, case-insensitive.  Use a trailing "-" to reverse the order of a column.  Use "." to specify the default order returned by the Goodreads API.  

  If an order is not specified, it is set based on the type of search.  A search based on a file of ISBN or book id codes will set the order to "." (the order of codes in the file).  A search by author or author id will set the order to "Author Book_Rating- Title".  Other searches will set the order to "Author Title Book_Rating-".

* --open-xlsx = Whether to open the results spreadsheet automatically in Excel (default is true).

* -X or --no-open-xlsx = Whether to automatically open the results spreadsheet in Excel.

* -p or --page <number> = Page number of search results to retrieve (default is 1).

* -r or --rating <number> = Filter by rating of book to meet or exceed (default is 0.0).

* -s or --search <text> = Search for text in either title or author.

* -t or --title <text1> -a or --author <text2> = Search for text in title and for other text in author.

* -t or --title <text> = Search by title.

* -k or --kindle-edition = Filter by whether a Kindle edition is available.

* -v or --version = Show program version and exit.

# Examples

Search for books by J. K. Rowling as an author: metabook -a "J. K. Rowling"

Search for books with Harry Potter in the title: metabook -t "Harry Potter"

Search for books with Stephen Hawking in either the author or title data: metabook -s "Stephen Hawking"

Modify the search so that books are ordered chronologically:  metabook -s "Stephen Hawking" -o "Year Title"

Filter the search by black hole in the description: metabook -s "Stephen Hawking" -d "black hole"

Filter the search by audio format and do not open the results spreadsheet automatically: metabook -s "Stephen Hawking" -f "audio" -X

Filter the search by Kindle books only and return the second page of results: metabook -s "Stephen Hawking" -k -p 2

Search for books based on an ISBN list: metabook --isbn-file books.txt
