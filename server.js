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

function safeDecodeHeader(value) {
  const s = String(value || "");
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

function readTag(tags, names) {
  for (const name of names) {
    if (Object.prototype.hasOwnProperty.call(tags, name) && tags[name] != null) {
      return tags[name];
    }
  }
  return null;
}

function parseExifCoordValue(value) {
  const direct = pickFirstFiniteNumber([value]);
  if (direct != null) return direct;
  if (typeof value !== "string") return null;
  const s = value.trim();
  if (!s) return null;

  // DMS style fallback (e.g. "8 deg 22' 09.94\"")
  const dms = s.match(
    /^(-?\d+(?:\.\d+)?)\D+(\d+(?:\.\d+)?)\D+(\d+(?:\.\d+)?)(?:\D|$)/i
  );
  if (!dms) return null;
  const deg = Number.parseFloat(dms[1]);
  const min = Number.parseFloat(dms[2]);
  const sec = Number.parseFloat(dms[3]);
  if (![deg, min, sec].every(Number.isFinite)) return null;
  const sign = deg < 0 ? -1 : 1;
  const absDeg = Math.abs(deg);
  return sign * (absDeg + min / 60 + sec / 3600);
}

function applyGpsRefSign(value, ref, axis) {
  if (!Number.isFinite(value)) return null;
  const r = String(ref || "").trim().toUpperCase();
  if (!r) return value;
  if (axis === "lat" && (r === "S" || r === "SOUTH")) return -Math.abs(value);
  if (axis === "lng" && (r === "W" || r === "WEST")) return -Math.abs(value);
  if (axis === "lat" && (r === "N" || r === "NORTH")) return Math.abs(value);
  if (axis === "lng" && (r === "E" || r === "EAST")) return Math.abs(value);
  return value;
}

function pickGpsFromTags(tags) {
  if (!tags || typeof tags !== "object") return null;
  const latRaw = readTag(tags, [
    "GPSLatitude",
    "CompositeGPSLatitude",
    "Composite:GPSLatitude",
  ]);
  const lngRaw = readTag(tags, [
    "GPSLongitude",
    "CompositeGPSLongitude",
    "Composite:GPSLongitude",
  ]);
  if (latRaw == null || lngRaw == null) return null;

  const latRef = readTag(tags, ["GPSLatitudeRef", "Composite:GPSLatitudeRef"]);
  const lngRef = readTag(tags, ["GPSLongitudeRef", "Composite:GPSLongitudeRef"]);
  const lat = applyGpsRefSign(parseExifCoordValue(latRaw), latRef, "lat");
  const lng = applyGpsRefSign(parseExifCoordValue(lngRaw), lngRef, "lng");
  if (lat == null || lng == null) return null;
  if (Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
  return { lat, lng };
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

    const rawName = safeDecodeHeader(req.headers["x-file-name"] || "photo.bin");
    const ext = path.extname(rawName) || ".bin";
    const tmpPath = path.join(
      os.tmpdir(),
      `fsis-exif-${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`
    );

    try {
      await fs.writeFile(tmpPath, body);
      const tags = await runExiftoolJson(tmpPath);

      const takenAt =
        toIsoStringFromExifDate(
          readTag(tags, ["DateTimeOriginal", "EXIF:DateTimeOriginal", "SubSecDateTimeOriginal"])
        ) ||
        toIsoStringFromExifDate(readTag(tags, ["CreateDate", "EXIF:CreateDate"])) ||
        toIsoStringFromExifDate(readTag(tags, ["ModifyDate", "EXIF:ModifyDate"])) ||
        toIsoStringFromExifDate(readTag(tags, ["GPSDateTime", "XMP:GPSDateTime"])) ||
        null;

      const gps = pickGpsFromTags(tags);

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

