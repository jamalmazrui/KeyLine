const axeCore = require("axe-core");
const fs = require("fs");
const opener = require("opener");
const path = require("path");
const process = require("process");
const puppeteer = require("puppeteer");
const { reporterFactory } = require("accessibility-insights-report");

function getUniqueFile(sDir, sRoot, sExt) {
  let sPath = path.join(sDir, `${sRoot}.${sExt}`);
  let iCount = 1;
  while (fs.existsSync(sPath)) {
    sPath = path.join(
      sDir,
      `${sRoot}-${String(iCount).padStart(2, "0")}.${sExt}`,
    );
    iCount++;
  }
  return sPath;
}

async function main(sUrl) {
  let aLines = await [sUrl];
  let bUrl = true;
  if (await !sUrl.includes("://")) {
    bUrl = false;
    let sLines = await fs.readFileSync(sUrl, "utf-8");
    aLines = await sLines.split("\n");
  }

  const sCurDir = await process.cwd();
  let sOutputHtml;
  const browser = await puppeteer.launch({
    executablePath:
      "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    headless: true,
  });

  while (aLines.length > 0) {
    sUrl = await aLines.shift();
    sUrl = await sUrl.trim();
    if (await !sUrl) continue;

    console.log();
    console.log("Testing " + sUrl);
    const page = await browser.newPage();
    await page.goto(sUrl);

    let sPageTitle = await page.title();
    sPageTitle = await sPageTitle.replace(/[^a-zA-Z0-9_,#\.\-]/gi, " ");
    sPageTitle = await sPageTitle.replace(/\s+/gi, " ").trim();
    console.log(sPageTitle);

    await page.evaluate(axeCore.source);
    const results = await page.evaluate(() => {
      return axe.run();
    });

    const reporter = reporterFactory();
    const sHtml = reporter
      .fromAxeResult({
        results,
        description: "description of report",
        serviceName: "service name",
        scanContext: {
          pageTitle: sPageTitle,
        },
      })
      .asHTML();

    sOutputHtml = await getUniqueFile(sCurDir, sPageTitle, "htm");
    await fs.writeFileSync(sOutputHtml, sHtml);
    await page.close();
  } // while
  await browser.close();
  if (bUrl) await opener(sOutputHtml);
}

const sUrl = process.argv[2];
if (!sUrl) {
  console.error("Please provide a URL or file as a command line argument");
  process.exit(1);
}
main(sUrl);
