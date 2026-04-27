const path = require("path");
const os = require("os");
const fs = require("fs/promises");
const { execFile } = require("child_process");
const { promisify } = require("util");
const express = require("express");
const puppeteer = require("puppeteer");
const execFileAsync = promisify(execFile);

const app = express();
const PORT = Number(process.env.PORT || 3000);

app.use(express.json({ limit: "2mb" }));

// Serve the existing static app
app.use(express.static(path.join(__dirname)));

async function runExiftoolJson(filePath) {
  const candidates = ["exiftool", "exiftool.exe", "exiftool(-k).exe"];
  let lastErr = null;
  for (const cmd of candidates) {
    try {
      const { stdout } = await execFileAsync(cmd, [
        "-j",
        "-n",
        "-api",
        "largefilesupport=1",
        filePath,
      ]);
      const parsed = JSON.parse(stdout);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed[0];
      return {};
    } catch (err) {
      lastErr = err;
    }
  }
  throw lastErr || new Error("ExifTool executable not found.");
}

function pickFirstFiniteNumber(values) {
  for (const v of values) {
    const n =
      typeof v === "number" ? v : typeof v === "string" ? Number.parseFloat(v) : NaN;
    if (Number.isFinite(n)) return n;
  }
  return null;
}

function toIsoStringFromExifDate(value) {
  if (!value || typeof value !== "string") return null;
  const m = value
    .trim()
    .match(/^(\d{4}):(\d{2}):(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (!m) {
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d.toISOString();
  }
  const d = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

app.post(
  "/api/exif/read",
  express.raw({ type: "application/octet-stream", limit: "25mb" }),
  async (req, res) => {
    const body = req.body;
    if (!body || !Buffer.isBuffer(body) || body.length === 0) {
      res.status(400).json({ error: "Missing image payload." });
      return;
    }

    const rawName = decodeURIComponent(String(req.headers["x-file-name"] || "photo.bin"));
    const ext = path.extname(rawName) || ".bin";
    const tmpPath = path.join(
      os.tmpdir(),
      `fsis-exif-${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`
    );

    try {
      await fs.writeFile(tmpPath, body);
      const tags = await runExiftoolJson(tmpPath);

      const lat = pickFirstFiniteNumber([
        tags.GPSLatitude,
        tags.CompositeGPSLatitude,
        tags["Composite:GPSLatitude"],
      ]);
      const lng = pickFirstFiniteNumber([
        tags.GPSLongitude,
        tags.CompositeGPSLongitude,
        tags["Composite:GPSLongitude"],
      ]);

      const takenAt =
        toIsoStringFromExifDate(tags.DateTimeOriginal) ||
        toIsoStringFromExifDate(tags.CreateDate) ||
        toIsoStringFromExifDate(tags.ModifyDate) ||
        null;

      const gps =
        lat != null && lng != null && Math.abs(lat) <= 90 && Math.abs(lng) <= 180
          ? { lat, lng }
          : null;

      res.status(200).json({
        gps,
        takenAt,
        hasExif: Object.keys(tags || {}).length > 1 || Boolean(gps || takenAt),
        source: "exiftool",
      });
    } catch (err) {
      console.error("EXIF read failed:", err);
      res.status(500).json({
        error:
          "Failed to read EXIF using ExifTool. Ensure exiftool is installed and available in PATH.",
      });
    } finally {
      try {
        await fs.unlink(tmpPath);
      } catch (_) {}
    }
  }
);

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

