const path = require("path");
const express = require("express");
const puppeteer = require("puppeteer");

const app = express();
const PORT = Number(process.env.PORT || 3000);

app.use(express.json({ limit: "2mb" }));

// Serve the existing static app
app.use(express.static(path.join(__dirname)));

function safeFilename(name) {
  const base = String(name || "inspection-order.pdf")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return base.toLowerCase().endsWith(".pdf") ? base : `${base}.pdf`;
}

app.post("/api/io/pdf", async (req, res) => {
  const entry = req.body?.entry ?? req.body;
  const filename = safeFilename(req.body?.filename);

  if (!entry || typeof entry !== "object") {
    res.status(400).json({ error: "Missing or invalid entry payload." });
    return;
  }

  let browser;
  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ["--disable-dev-shm-usage"],
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 794, height: 1122, deviceScaleFactor: 1 }); // ~A4 at 96dpi
    await page.emulateMediaType("print");

    // Inject entry data before any page scripts run.
    await page.evaluateOnNewDocument((data) => {
      window.__FSIS_IO_ENTRY__ = data;
    }, entry);

    const url = `http://127.0.0.1:${PORT}/io_fsis.html?render=pdf`;
    await page.goto(url, { waitUntil: "networkidle0" });

    // Ensure fonts/layout settle
    await page.waitForSelector(".page");

    // Force a solid white background to avoid any viewer "black/gray" artifacts.
    await page.evaluate(() => {
      document.documentElement.style.background = "#ffffff";
      document.body.style.background = "#ffffff";
    });

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      omitBackground: false,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
    });

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.status(200).send(pdf);
  } catch (err) {
    console.error("IO PDF generation failed:", err);
    res.status(500).json({ error: "Failed to generate IO PDF." });
  } finally {
    try {
      await browser?.close();
    } catch (_) {}
  }
});

app.listen(PORT, () => {
  console.log(`FSIS server running at http://127.0.0.1:${PORT}`);
});

