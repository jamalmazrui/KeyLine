let puppeteer = require("puppeteer-core");
let fs = require("fs");
let path = require("path");
let process = require("process");

let sOutputDir = "./results";
sOutputDir = path.resolve(sOutputDir);

function sleep(iSeconds) {
  return new Promise((resolve) => {
    setTimeout(resolve, iSeconds * 1000);
  });
}

function generateUniqueDirName(sRootName) {
  let sDirPath = path.join(process.cwd(), `${sRootName}`);
  let iCount = 1;
  while (fs.existsSync(sDirPath)) {
    sDirPath = path.join(
      process.cwd(),
      `${sRootName}-${String(iCount).padStart(2, "0")}`,
    );
    iCount++;
  }
  return sDirPath;
}

// main

(async () => {
  let sUrl = process.argv[2];
  if (!sUrl) {
    console.error("Please provide a URL as a command-line argument.");
    process.exit(1);
  }

  let aLines = await [sUrl];
  if (await !sUrl.includes("://")) {
    let sLines = await fs.readFileSync(sUrl, "utf-8");
    aLines = await sLines.split("\n");
  } // else

  let sStartDir = await __dirname;
  let sCurDir = await process.cwd();
  let sBaseYml = await ".achecker.yml";
  let sSourceYml = await path.join(sStartDir, sBaseYml);
  let sTargetYml = await path.join(sCurDir, sBaseYml);

  let browser = await puppeteer.launch({
    executablePath:
      "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe", // Adjust this path if needed
    headless: true,
  });
  let checker = null;
  while (aLines.length > 0) {
    sUrl = await aLines.shift();
    sUrl = await sUrl.trim();
    if (await !sUrl) continue;

    let page = await browser.newPage();
    await page.goto(sUrl);
    let sPageTitle = await page.title();
    sPageTitle = await sPageTitle.replace(/[^a-zA-Z0-9_,#\.\-]/gi, " ");
    sPageTitle = await sPageTitle.replace(/\s+/gi, " ").trim();

    sOutputDir = await generateUniqueDirName(sPageTitle);
    await fs.mkdirSync(sOutputDir);

    let sScreenshotPng = await path.join(sOutputDir, "screenshot.png");
    await page.screenshot({ path: sScreenshotPng, fullPage: true });
    let sPageContent = await page.content();
    let sPageHtml = await path.join(sOutputDir, "page.html");
    await fs.writeFileSync(sPageHtml, sPageContent);
    let sTreeJson = await path.join(sOutputDir, "tree.json");
  let dTree = await page.accessibility.snapshot();
await fs.writeFileSync(sTreeJson, JSON.stringify(dTree));

    console.log(path.basename(sOutputDir));
    console.log();
    let sOldYml = await fs.readFileSync(sSourceYml, "utf-8");
    let sNewYml = await sOldYml.replace(
      /outputFolder.+/,
      `outputFolder: ${sOutputDir}`,
    );
    await fs.writeFileSync(sTargetYml, sNewYml);

    checker = require("accessibility-checker");
    let dResults = await checker.getCompliance(page, "results");
    let dReport = await dResults.report;
    await checker.close();
    if (sTargetYml != sSourceYml) await fs.unlinkSync(sTargetYml);
    await fs.writeFileSync(sSourceYml, sOldYml);

    await page.close();
  } // while
  await browser.close();
  // console.log("Done");
})();
