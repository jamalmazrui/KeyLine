# TestURL

by Jamal Mazrui \
Consultant, Access Success LLC

TestURL is a free, open source tool for running accessibility tests on web pages specified by URL. It generates test results in HTML format.

## Installation

Install Node.js and the Node Package Manager (NPM) from the [Node website](https://nodejs.org/en).

Clone the repository: \
<   git clone https://github.com/JamalMazrui/TestURL>

Change to the installation directory: \
`cd TestURL`

Install support packages: \
`npm install`

## Operation

Run the program: \
`node testUrl.js <URL>`

where \<URL\> is the web address of the page to be tested for conformance to the [Web Content Accessibility Guidelines](https://www.w3.org/TR/WCAG22/) (WCAG).

Rather than passing a URL as a parameter, a file name may be passed instead. Such a text file lists multiple URLs, one per line: \
`node testUrl.js <File>`

On Windows, a batch file, TestURL.cmd, is also available: \
`TestURL.cmd <URL or File>`

The program will run the Microsoft Edge browser with the API of [Deque Axe](https://www.npmjs.com/package/axe-core). The resulting HTML report will have a unique file name in the current directory, based on the title of the web page tested, and if needed, a numeric suffix.

If a URL was passed as a parameter, the HTML report will be opened when the program ends. If a file was passed instead, a report is not opened automatically.

Note that automated testing only catches about a third of accessibility errors, so manual testing is also needed to evaluate WCAG conformance.

The TestURL project is available on the web at \
<http://GitHub.com/JamalMazrui/TestURL>

The project may be downloaded in a single zip archive from \
<http://GitHub.com/JamalMazrui/TestURL/archive/main.zip>
