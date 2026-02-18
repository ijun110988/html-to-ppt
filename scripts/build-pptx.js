const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");

const { chromium } = require("playwright");
const pptxgen = require("pptxgenjs");

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function listSlidesHtml(projectRoot) {
  const entries = fs.readdirSync(projectRoot, { withFileTypes: true });
  const slides = entries
    .filter((e) => e.isFile() && /^slide\d+\.html$/i.test(e.name))
    .map((e) => e.name)
    .sort((a, b) => {
      const an = Number(a.match(/^slide(\d+)\.html$/i)[1]);
      const bn = Number(b.match(/^slide(\d+)\.html$/i)[1]);
      return an - bn;
    });

  return slides.map((f) => path.join(projectRoot, f));
}

async function renderSlideToPng(page, htmlPath, outPngPath) {
  const fileUrl = pathToFileURL(htmlPath).toString();

  await page.goto(fileUrl, { waitUntil: "networkidle" });

  // Ensure fonts are loaded before screenshot
  await page.evaluate(async () => {
    if (document.fonts && document.fonts.ready) {
      await document.fonts.ready;
    }
  });

  // Prefer `.slide-container` (present in your HTML). Fallback to full page.
  const container = await page.$(".slide-container");
  if (container) {
    await container.screenshot({ path: outPngPath });
  } else {
    await page.screenshot({ path: outPngPath, fullPage: true });
  }
}

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const distDir = path.join(projectRoot, "dist");
  const framesDir = path.join(distDir, "frames");
  const outPptxPath = path.join(distDir, "presentasi.pptx");

  ensureDir(framesDir);

  const slideHtmlPaths = listSlidesHtml(projectRoot);
  if (slideHtmlPaths.length === 0) {
    console.error("No slide*.html files found in project root.");
    process.exitCode = 1;
    return;
  }

  const browser = await chromium.launch();
  const context = await browser.newContext({
    viewport: { width: 1280, height: 720 },
    deviceScaleFactor: 1,
  });
  const page = await context.newPage();

  // Render each slide HTML to PNG
  const pngPaths = [];
  for (const htmlPath of slideHtmlPaths) {
    const base = path.basename(htmlPath, path.extname(htmlPath));
    const outPng = path.join(framesDir, `${base}.png`);
    await renderSlideToPng(page, htmlPath, outPng);
    pngPaths.push(outPng);
    process.stdout.write(`Rendered ${path.basename(htmlPath)} -> ${path.relative(projectRoot, outPng)}\n`);
  }

  await browser.close();

  // Build PPTX (16:9) and place each PNG to fill slide.
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";

  for (const pngPath of pngPaths) {
    const slide = pptx.addSlide();
    slide.addImage({ path: pngPath, x: 0, y: 0, w: 13.333, h: 7.5 });
  }

  ensureDir(distDir);
  await pptx.writeFile({ fileName: outPptxPath });
  process.stdout.write(`\nCreated: ${path.relative(projectRoot, outPptxPath)}\n`);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
