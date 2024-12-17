# TestPage

by Jamal Mazrui \
Consultant, Access Success LLC

TestPage is a free, open source tool for running accessibility tests on web pages specified by URL. It generates test results in CSV, JSON, HTML, and XLSX formats.

## Installation

Install Node.js and the Node Package Manager (NPM) from the [Node website](https://nodejs.org/en).

Clone the repository: \
<   git clone https://github.com/JamalMazrui/TestPage>

Change to the installation directory: \
`cd TestPage`

Install support packages: \
`npm install`

## Operation

Run the program: \
`node testPage.js <URL>`

where \<URL\> is the web address of the page to be tested for conformance to the [Web Content Accessibility Guidelines](https://www.w3.org/TR/WCAG22/) (WCAG).

On Windows, a batch file, TestPage.cmd, is also available that can test multiple URLs listed in a text file: \
`TestPage.cmd <File>`

where \<File\> is a text file containing a URL on each line.

The program will run the Microsoft Edge browser with the API of [IBM Accessibility Checker](https://www.npmjs.com/package/accessibility-checker). A unique subdirectory will be created for each page tested, based on its title, and if needed, a numeric suffix. The test results will be contained there in various file formats. To review the source of the results, the HTML content, screenshot, and accessibility tree of the page tested are also saved.

Note that automated testing only catches about a third of accessibility errors, so manual testing is also needed to evaluate WCAG conformance.

The TestPage project is available on the web at \
<http://GitHub.com/JamalMazrui/TestPage>

The project may be downloaded in a single zip archive from \
<http://GitHub.com/JamalMazrui/TestPage/archive/main.zip>
