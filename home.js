const SESSION_KEY = "fsis.session";

function getSession() {
  const raw = sessionStorage.getItem(SESSION_KEY) ?? localStorage.getItem(SESSION_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    sessionStorage.removeItem(SESSION_KEY);
    localStorage.removeItem(SESSION_KEY);
    return null;
  }
}

function clearSession() {
  sessionStorage.removeItem(SESSION_KEY);
  localStorage.removeItem(SESSION_KEY);
}

function requireSession() {
  const session = getSession();
  if (!session?.username) {
    window.location.replace("./index.html");
    return null;
  }
  return session;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = value;
}

function toFriendlyDate(isoString) {
  if (!isoString) return "—";
  const d = new Date(isoString);
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleString();
}

function getCurrentView() {
  const hash = (window.location.hash || "#map").replace(/^#/, "");
  if (hash === "inspection" || hash === "fsec" || hash === "map") return hash;
  return "map";
}

const MAP_CENTER = [8.369, 124.864];
const MAP_ZOOM = 12;
let mapInstance = null;
let userLocationWatchId = null;
let userLocationMarker = null;
let userLocationAccuracyCircle = null;
let hasCenteredOnUser = false;
let lastUserLatitude = null;
let lastUserLongitude = null;

let currentExifLat = null;
let currentExifLng = null;
let currentExifPreviewUrl = null;
let currentExifTakenAt = null;
let currentExifFile = null;

let inspectionMarkersLayer = null;
let inspectionDataLoaded = false;
let inspectionActiveTab = "with-location";

function resizeMapLayout() {
  const mapSection = document.querySelector('[data-view="map"]');
  const layout = document.querySelector(".map-layout");
  const mapEl = document.getElementById("map");

  if (!mapSection || !layout || !mapEl) return;

  const available = window.innerHeight;
  if (available <= 0) return;

  layout.style.height = available + "px";
  mapSection.style.height = available + "px";
  mapEl.style.height = "100%";

  if (mapInstance) {
    setTimeout(() => mapInstance.invalidateSize(), 0);
  }
}

function initLeafletMap() {
  if (mapInstance || !window.L) return;
  const el = document.getElementById("map");
  if (!el) return;

  resizeMapLayout();

  el.innerHTML = "";
  mapInstance = L.map(el, { zoomControl: false }).setView(MAP_CENTER, MAP_ZOOM);

  // Layer to hold all inspection markers so we can manage them together
  inspectionMarkersLayer = L.layerGroup().addTo(mapInstance);

  // Base layers: OpenStreetMap (faster, lighter) + Google-style satellite with labels
  const osmLayer = L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution:
      '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  });

  // Uses Google tiles for a clearer satellite view with labels (heavier on data).
  // To keep it usable on slow mobile data, we cap zoom and disable retina tiles.
  const satelliteLayer = L.tileLayer(
    "https://{s}.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
    {
      maxZoom: 19,
      maxNativeZoom: 18,
      detectRetina: false,
      subdomains: ["mt0", "mt1", "mt2", "mt3"],
      attribution:
        '&copy; <a href="https://www.google.com/maps">Google Maps</a>',
    }
  );

  // Start with the satellite layer as the default visible layer
  satelliteLayer.addTo(mapInstance);

  // Layer switcher so you can toggle between views
  L.control
    .layers(
      {
        "Road map": osmLayer,
        Satellite: satelliteLayer,
      },
      {},
      { position: "topright" }
    )
    .addTo(mapInstance);

  // Start tracking user's current location in real time
  startUserLocationTracking();

  // If inspection data is already loaded, render any markers that have coordinates
  if (Array.isArray(inspectionData) && inspectionData.length > 0) {
    inspectionData.forEach((row) => {
      if (row.lat != null && row.lng != null) {
        addInspectionMarkerFromEntry(row);
      }
    });
  }

  // Make sure the map fully renders after layout
  setTimeout(() => {
    mapInstance.invalidateSize();
  }, 0);
}

function resetMapView() {
  if (!mapInstance) return;
  mapInstance.setView(MAP_CENTER, MAP_ZOOM);
}

function handleFabAddInspection() {
  // Stay on the map view and open the inspection modal in-place
  if (typeof inspectionOpenModal === "function") {
    inspectionOpenModal();
  }
}

function showView(name) {
  const views = Array.from(document.querySelectorAll("[data-view]"));
  for (const v of views) {
    const viewName = v.getAttribute("data-view");
    const isActive = viewName === name;
    v.hidden = !isActive;
    // Ensure sections are truly removed from layout when inactive
    v.style.display = isActive ? "" : "none";
  }

  document.body.classList.toggle("is-map-view", name === "map");

  if (name === "inspection") {
    setInspectionTab(inspectionActiveTab);
  }

  if ((name === "map" || name === "inspection") && !inspectionDataLoaded) {
    inspectionInitData();
  } else if (name === "fsec" && !fsecDataLoaded) {
    fsecInitData();
  }

  const links = Array.from(document.querySelectorAll("[data-view-link]"));
  for (const link of links) {
    const target = link.getAttribute("data-view-link");
    link.classList.toggle("is-active", target === name);
  }

  if (name === "map") {
    if (!mapInstance) {
      resizeMapLayout();
      initLeafletMap();
    } else {
      // Ensure map resizes correctly when returning to the tab
      resizeMapLayout();
    }
  } else {
    // When leaving the map view, clear any map-specific heights
    const mapSection = document.querySelector('[data-view="map"]');
    const layout = document.querySelector(".map-layout");
    if (mapSection) mapSection.style.height = "";
    if (layout) layout.style.height = "";
  }

  // Always reset scroll to top when changing main views
  window.scrollTo({ top: 0, behavior: "auto" });
}

function startUserLocationTracking() {
  if (!navigator.geolocation || !mapInstance) return;

  if (userLocationWatchId !== null) return; // already tracking

  userLocationWatchId = navigator.geolocation.watchPosition(
    (pos) => {
      const { latitude, longitude, accuracy } = pos.coords;
      const latlng = [latitude, longitude];
      lastUserLatitude = latitude;
      lastUserLongitude = longitude;

      if (!userLocationMarker) {
        userLocationMarker = L.marker(latlng, {
          title: "Your location",
        }).addTo(mapInstance);
      } else {
        userLocationMarker.setLatLng(latlng);
      }

      if (!userLocationAccuracyCircle) {
        userLocationAccuracyCircle = L.circle(latlng, {
          radius: accuracy || 30,
          color: "#2563EB",
          fillColor: "#3B82F6",
          fillOpacity: 0.15,
        }).addTo(mapInstance);
      } else {
        userLocationAccuracyCircle.setLatLng(latlng);
        userLocationAccuracyCircle.setRadius(accuracy || 30);
      }

      // On first successful fix, center the map on the user's location
      if (!hasCenteredOnUser) {
        hasCenteredOnUser = true;
        mapInstance.setView(latlng, 16);
      }
    },
    (err) => {
      console.warn("Geolocation error:", err);
      hasCenteredOnUser = false;

      // Provide a clear, non-blocking message and gracefully fall back to the default view
      if (err.code === err.PERMISSION_DENIED) {
        logbookShowToast(
          "inspection-toast",
          "Location permission was denied. Showing default map view for Manolo Fortich."
        );
      } else {
        logbookShowToast(
          "inspection-toast",
          "Could not determine your location. Showing default map view."
        );
      }
    },
    {
      enableHighAccuracy: true,
      maximumAge: 5000,
      timeout: 10000,
    }
  );
}

function openNavSidebar() {
  const sidebar = document.getElementById("navSidebar");
  const overlay = document.getElementById("navSidebarOverlay");
  const burger = document.getElementById("burgerBtn");
  if (sidebar) sidebar.classList.add("is-open");
  if (overlay) overlay.classList.add("is-open");
  if (burger) {
    burger.setAttribute("aria-expanded", "true");
    burger.setAttribute("aria-label", "Close menu");
  }
}

function closeNavSidebar() {
  const sidebar = document.getElementById("navSidebar");
  const overlay = document.getElementById("navSidebarOverlay");
  const burger = document.getElementById("burgerBtn");
  if (sidebar) sidebar.classList.remove("is-open");
  if (overlay) overlay.classList.remove("is-open");
  if (burger) {
    burger.setAttribute("aria-expanded", "false");
    burger.setAttribute("aria-label", "Open menu");
  }
}

function toggleNavSidebar() {
  const sidebar = document.getElementById("navSidebar");
  if (sidebar?.classList.contains("is-open")) closeNavSidebar();
  else openNavSidebar();
}

function initViewRouting() {
  const applyFromHash = () => showView(getCurrentView());

  window.addEventListener("hashchange", applyFromHash);

  document.addEventListener("click", (e) => {
    const target = e.target instanceof Element ? e.target : null;
    const link = target?.closest?.("[data-view-link]");
    if (!link) return;

    const view = link.getAttribute("data-view-link");
    if (!view) return;

    e.preventDefault();
    showView(view);
    if (getCurrentView() !== view) {
      window.location.hash = view;
    }
    closeNavSidebar();
  });

  applyFromHash();
}

function init() {
  const session = requireSession();
  if (!session) return;

  const name = session.displayName || session.username || "User";
  setText("userName", name);
  const lastEl = document.getElementById("lastLogin");
  if (lastEl) {
    lastEl.textContent = session.issuedAt ? "Signed in " + toFriendlyDate(session.issuedAt) : "";
    lastEl.setAttribute("aria-hidden", lastEl.textContent ? "false" : "true");
  }

  const logoutBtn = document.getElementById("logoutBtn");
  logoutBtn?.addEventListener("click", () => {
    clearSession();
    window.location.replace("./index.html");
  });

  const burgerBtn = document.getElementById("burgerBtn");
  burgerBtn?.addEventListener("click", toggleNavSidebar);
  const navOverlay = document.getElementById("navSidebarOverlay");
  navOverlay?.addEventListener("click", closeNavSidebar);

  // Inspection sub‑nav (With location / No location yet)
  document.addEventListener("click", (e) => {
    const btn = e.target.closest?.(".inspection-subnav-btn");
    if (!btn) return;
    const tab = btn.getAttribute("data-inspection-tab");
    if (!tab) return;
    setInspectionTab(tab);
  });

  // Keep the map sized correctly on window resize
  window.addEventListener("resize", () => {
    if (getCurrentView() === "map") {
      resizeMapLayout();
    }
  });

  initViewRouting();
  initInspectionPhotoExif();
  refreshStorageBadge();

  const initialView = getCurrentView();
  if ((initialView === "map" || initialView === "inspection") && !inspectionDataLoaded) {
    inspectionInitData();
    if (initialView === "inspection") {
      setInspectionTab(inspectionActiveTab);
    }
  } else if (initialView === "fsec" && !fsecDataLoaded) {
    fsecInitData();
  }
}

document.addEventListener("DOMContentLoaded", init);

// -----------------------------
// Shared Supabase + utilities
// -----------------------------

const SUPABASE_URL = "https://drqgbkninqpvhvhnbatk.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRycWdia25pbnFwdmh2aG5iYXRrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE5MTQzODQsImV4cCI6MjA4NzQ5MDM4NH0.CliUD5Ow17OXvaqDzYdAbi-rrTg_u-e4OyomcGrgZk0";

const supabaseClient =
  window.supabase?.createClient && SUPABASE_URL && SUPABASE_ANON_KEY
    ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
    : null;

function isSupabaseEnabled() {
  return !!supabaseClient;
}

function setStorageBadge(mode) {
  const el = document.getElementById("storageBadge");
  if (!el) return;
  const isDb = mode === "db";
  el.textContent = isDb ? "DB" : "LOCAL";
  el.classList.toggle("db", isDb);
  el.classList.toggle("local", !isDb);
  el.title = isDb ? "Database connected" : "Offline mode (local storage)";
}

async function refreshStorageBadge() {
  if (!supabaseClient) {
    setStorageBadge("local");
    return;
  }

  try {
    const { error } = await supabaseClient
      .from("inspection_logbook")
      .select("id")
      .limit(1);
    if (error) throw error;
    setStorageBadge("db");
  } catch (err) {
    console.warn("Database check failed, using local mode:", err);
    setStorageBadge("local");
  }
}

function logbookEsc(str) {
  if (!str) return "—";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function logbookFormatDate(d) {
  if (!d) return "—";
  return new Date(d + "T00:00:00").toLocaleDateString("en-PH", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

function logbookShowToast(id, msg) {
  const t = document.getElementById(id);
  if (!t) return;
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 3000);
}

function showSaveIndicator(message) {
  const el = document.getElementById("saveIndicator");
  if (!el) return;
  el.textContent = message;
  el.setAttribute("aria-hidden", "false");
  el.classList.add("is-visible");
  if (showSaveIndicator._timer) {
    clearTimeout(showSaveIndicator._timer);
  }
  showSaveIndicator._timer = setTimeout(() => {
    el.classList.remove("is-visible");
    el.setAttribute("aria-hidden", "true");
  }, 2200);
}

function setInspectionTab(tab) {
  inspectionActiveTab = tab;

  const panelWith = document.getElementById("panel-inspection");
  const panelNoLocation = document.getElementById("panel-inspection-nophoto");
  const buttons = document.querySelectorAll(".inspection-subnav-btn");

  buttons.forEach((btn) => {
    const t = btn.getAttribute("data-inspection-tab");
    const isActive = t === tab;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  if (panelWith) panelWith.style.display = tab === "with-location" ? "" : "none";
  if (panelNoLocation) panelNoLocation.style.display = tab === "no-location" ? "" : "none";
}

// -----------------------------
// Inspection logbook module
// -----------------------------

const INSPECTION_STORAGE_KEY = "bfp_inspection";
let inspectionData = [];
let inspectionEditingIdx = null;
let inspectionEditingId = null;

function inspectionLoadFromLocal() {
  inspectionData = JSON.parse(
    localStorage.getItem(INSPECTION_STORAGE_KEY) || "[]"
  );
}

function inspectionSaveToLocal() {
  // Avoid storing huge inline image data URLs in localStorage (they quickly exceed quota).
  const safe = inspectionData.map((row) => {
    const copy = { ...row };
    if (typeof copy.photo_url === "string" && copy.photo_url.startsWith("data:")) {
      copy.photo_url = null;
    }
    return copy;
  });

  try {
    localStorage.setItem(INSPECTION_STORAGE_KEY, JSON.stringify(safe));
  } catch (err) {
    console.warn("Failed to persist inspection cache to localStorage:", err);
    // If caching fails, we silently skip it; the in-memory data and database writes still succeed.
  }
}

function inspectionSave() {
  if (!isSupabaseEnabled()) inspectionSaveToLocal();
}

function inspectionFormatAddressDisplay(row) {
  const addrLine = row.addr_line;
  const addrBarangay = row.addr_barangay;
  if (
    addrLine != null &&
    addrBarangay != null &&
    (addrLine !== "" || addrBarangay !== "")
  ) {
    return [
      addrLine,
      addrBarangay,
      "Manolo Fortich",
      "Bukidnon",
      "Region X",
    ]
      .filter(Boolean)
      .join(", ");
  }

  const full = (row.insp_address || row.fsec_address || "").toString().trim();
  if (!full) return "—";
  const parts = full.split(/,\s*/);
  if (parts.length >= 5) {
    const line = parts[4];
    const barangay = (parts[3] || "")
      .replace(/^Barangay\s+/i, "")
      .trim();
    const municipal = parts[2] || "";
    const province = parts[1] || "";
    const region = parts[0] || "";
    return [line, barangay, municipal, province, region]
      .filter(Boolean)
      .join(", ");
  }
  return full;
}

function inspectionRenderTable() {
  const tbody = document.getElementById("tbody-inspection");
  const empty = document.getElementById("empty-inspection");
  const tableWrap = document.getElementById("table-inspection")?.closest(".table-wrap");
  const tbodyNoPhoto = document.getElementById("tbody-inspection-nophoto");
  const emptyNoPhoto = document.getElementById("empty-inspection-nophoto");
  const tableWrapNoPhoto = document.getElementById("table-inspection-nophoto")?.closest(".table-wrap");
  const panelNoPhoto = document.getElementById("panel-inspection-nophoto");

  if (!tbody || !empty) return;

  tbody.innerHTML = "";
  if (tbodyNoPhoto) tbodyNoPhoto.innerHTML = "";
  let withLocationCount = 0;
  let noLocationCount = 0;

  if (inspectionData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
  } else {
    empty.style.display = "none";
    if (tableWrap) tableWrap.style.display = "";

    inspectionData.forEach((row, i) => {
      const hasLocation = row.lat != null && row.lng != null;

      const baseRowHtml = `
        <td data-label="#">${i + 1}</td>
        <td class="td-io" data-label="IO Number">${logbookEsc(row.io_number)}</td>
        <td data-label="Name of Owner">${logbookEsc(row.insp_owner)}</td>
        <td data-label="Business / Establishment"><strong>${logbookEsc(row.business_name)}</strong></td>
        <td data-label="Address">${logbookEsc(inspectionFormatAddressDisplay(row))}</td>
        <td class="td-date" data-label="Date Inspected">${logbookFormatDate(row.date_inspected)}</td>
        <td class="td-fsic" data-label="FSIC Number">${logbookEsc(row.fsic_number)}</td>
        <td data-label="Inspected By">${logbookEsc(row.inspected_by)}</td>
      `;

      if (hasLocation) {
        withLocationCount++;
        const tr = document.createElement("tr");
        tr.innerHTML = `
          ${baseRowHtml}
          <td class="col-action" data-label="Action">
            <div class="tbl-actions">
              <button class="btn-edit" onclick="inspectionEditEntry(${i})">Edit</button>
              <button class="btn-del" onclick="inspectionDeleteEntry(${i})">Delete</button>
              <button class="btn-edit" onclick="inspectionDownloadPdf(${i})">Download PDF</button>
              <button class="btn-edit" onclick="inspectionAddPhoto(${i})">Add photo</button>
            </div>
          </td>
        `;
        tbody.appendChild(tr);
      } else {
        noLocationCount++;
        if (tbodyNoPhoto) {
          const tr2 = document.createElement("tr");
          tr2.innerHTML = `
            ${baseRowHtml}
            <td class="col-action" data-label="Action">
              <div class="tbl-actions">
                <button class="btn-edit" onclick="inspectionEditEntry(${i})">Edit</button>
                <button class="btn-del" onclick="inspectionDeleteEntry(${i})">Delete</button>
                <button class="btn-edit" onclick="inspectionDownloadPdf(${i})">Download PDF</button>
                <button class="btn-edit" onclick="inspectionAddPhoto(${i})">Add photo</button>
              </div>
            </td>
          `;
          tbodyNoPhoto.appendChild(tr2);
        }
      }
    });
  }

  if (tbodyNoPhoto && emptyNoPhoto && panelNoPhoto) {
    if (noLocationCount === 0) {
      emptyNoPhoto.style.display = "block";
      if (tableWrapNoPhoto) tableWrapNoPhoto.style.display = "none";
      panelNoPhoto.style.display = "none";
    } else {
      emptyNoPhoto.style.display = "none";
      if (tableWrapNoPhoto) tableWrapNoPhoto.style.display = "";
      panelNoPhoto.style.display = "";
    }
  }
}

function inspectionEditEntry(idx) {
  const row = inspectionData[idx];
  if (!row) return;

  inspectionEditingIdx = idx;
  inspectionEditingId = row.id || null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };

  setVal("inspection_io_number", row.io_number);
  setVal("inspection_fsic_number", row.fsic_number);
  setVal("inspection_owner", row.insp_owner);
  setVal("inspection_business_name", row.business_name);
  setVal("inspection_date_inspected", row.date_inspected);
  setVal("inspection_inspected_by", row.inspected_by);

  const addr = (row.insp_address || "").toString();
  const brgyMatch = addr.match(/Barangay\s+([^,]+)/i);
  setVal(
    "inspection_addr_barangay",
    row.addr_barangay || (brgyMatch ? brgyMatch[1].trim() : "")
  );
  setVal("inspection_addr_line", row.addr_line || "");

  const overlay = document.getElementById("inspection-modal-overlay");
  if (overlay) overlay.classList.add("open");
  setText("inspection-modal-title", "Edit Inspection Record");
  setText("inspection-modal-subtitle", "Inspection Logbook");
  const btn = document.getElementById("inspection-btn-save");
  if (btn) btn.textContent = "Update Record";
}

function inspectionAddPhoto(idx) {
  inspectionEditEntry(idx);
  const photoInput = document.getElementById("inspection_photo");
  if (photoInput) {
    photoInput.scrollIntoView({ behavior: "smooth", block: "center" });
    photoInput.focus();
  }
}

async function inspectionDownloadPdf(idx) {
  const row = inspectionData[idx];
  if (!row || !window.PDFLib) return;

  try {
    const url = "./io_fsis.pdf"; // base template
    const existingPdfBytes = await fetch(url).then((res) => res.arrayBuffer());

    const { PDFDocument, StandardFonts } = PDFLib;
    const pdfDoc = await PDFDocument.load(existingPdfBytes);
    const pages = pdfDoc.getPages();
    const page = pages[0];

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const fontSize = 9;

    // NOTE: x,y positions are placeholders and should be tuned against your template.
    page.drawText(row.io_number || "", {
      x: 120,
      y: 640,
      size: fontSize,
      font,
    });

    page.drawText(row.insp_owner || "", {
      x: 120,
      y: 625,
      size: fontSize,
      font,
    });

    page.drawText(row.business_name || "", {
      x: 120,
      y: 610,
      size: fontSize,
      font,
    });

    page.drawText(inspectionFormatAddressDisplay(row), {
      x: 120,
      y: 595,
      size: fontSize,
      font,
      maxWidth: 360,
    });

    page.drawText(logbookFormatDate(row.date_inspected), {
      x: 120,
      y: 580,
      size: fontSize,
      font,
    });

    page.drawText(row.fsic_number || "", {
      x: 120,
      y: 565,
      size: fontSize,
      font,
    });

    page.drawText(row.inspected_by || "", {
      x: 120,
      y: 550,
      size: fontSize,
      font,
    });

    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `inspection-${row.io_number || row.id || "form"}.pdf`;
    link.click();
    URL.revokeObjectURL(link.href);
  } catch (err) {
    console.error("Failed to generate inspection PDF", err);
    logbookShowToast(
      "inspection-toast",
      "Could not generate PDF for this inspection."
    );
  }
}

function inspectionDeleteEntry(idx) {
  if (!confirm("Delete this record?")) return;

  if (!isSupabaseEnabled()) {
    inspectionData.splice(idx, 1);
    inspectionSave();
    inspectionRenderTable();
    logbookShowToast("inspection-toast", "Record deleted.");
    return;
  }

  const row = inspectionData[idx];
  if (!row?.id) {
    logbookShowToast(
      "inspection-toast",
      "⚠️ Cannot delete: missing record id."
    );
    return;
  }

  (async () => {
    try {
      const { error } = await supabaseClient
        .from("inspection_logbook")
        .delete()
        .eq("id", row.id);
      if (error) throw error;
      await inspectionLoadFromSupabase();
      inspectionRenderTable();
      logbookShowToast("inspection-toast", "Record deleted.");
    } catch (err) {
      logbookShowToast(
        "inspection-toast",
        "⚠️ Delete failed: " + (err?.message || err)
      );
    }
  })();
}

function inspectionOpenModal() {
  inspectionEditingIdx = null;
  inspectionEditingId = null;

  // Reset any previously extracted EXIF coordinates and photo data
  currentExifLat = null;
  currentExifLng = null;
  currentExifPreviewUrl = null;
  currentExifTakenAt = null;

  const overlay = document.getElementById("inspection-modal-overlay");
  if (overlay) overlay.classList.add("open");

  setText("inspection-modal-title", "Add Inspection Record");
  setText("inspection-modal-subtitle", "Inspection Logbook");
  const btn = document.getElementById("inspection-btn-save");
  if (btn) btn.textContent = "Save Record";

  inspectionClearForm();
  const date = document.getElementById("inspection_date_inspected");
  if (date) date.value = new Date().toISOString().slice(0, 10);
  
  // Clear photo input
  const photoInput = document.getElementById("inspection_photo");
  if (photoInput) photoInput.value = "";
}

function inspectionCloseModal() {
  const overlay = document.getElementById("inspection-modal-overlay");
  if (overlay) overlay.classList.remove("open");
}

function inspectionCloseOnOverlay(e) {
  const overlay = document.getElementById("inspection-modal-overlay");
  if (e.target === overlay) inspectionCloseModal();
}

function inspectionClearForm() {
  [
    "inspection_io_number",
    "inspection_fsic_number",
    "inspection_owner",
    "inspection_business_name",
    "inspection_addr_barangay",
    "inspection_addr_line",
    "inspection_date_inspected",
    "inspection_inspected_by",
  ].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
}

function inspectionSaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();

  const region =
    (document.getElementById("inspection_addr_region")?.value || "X").trim();
  const province =
    (document.getElementById("inspection_addr_province")?.value ||
      "Bukidnon").trim();
  const municipal =
    (document.getElementById("inspection_addr_municipal")?.value ||
      "Manolo Fortich").trim();
  const barangay = (
    document.getElementById("inspection_addr_barangay") || { value: "" }
  ).value.trim();
  const line = (
    document.getElementById("inspection_addr_line") || { value: "" }
  ).value.trim();

  const mergedAddress = `Region ${region}, ${province}, ${municipal}, Barangay ${barangay}, ${line}`;

  const entry = {
    io_number: (
      document.getElementById("inspection_io_number") || { value: "" }
    ).value.trim(),
    fsic_number: (
      document.getElementById("inspection_fsic_number") || { value: "" }
    ).value.trim(),
    insp_owner: (
      document.getElementById("inspection_owner") || { value: "" }
    ).value.trim(),
    business_name: (
      document.getElementById("inspection_business_name") || { value: "" }
    ).value.trim(),
    insp_address: mergedAddress,
    addr_barangay: barangay,
    addr_line: line,
    date_inspected: (
      document.getElementById("inspection_date_inspected") || { value: "" }
    ).value,
    inspected_by:
      (document.getElementById("inspection_inspected_by") || { value: "" })
        .value.trim(),
    // Optional coordinates and photo metadata extracted from EXIF / geolocation
    lat: currentExifLat,
    lng: currentExifLng,
    photo_url: currentExifPreviewUrl,
    photo_taken_at: currentExifTakenAt,
    created_at: new Date().toISOString(),
  };

  // If the photo has no GPS EXIF, fall back to the user's current geolocation
  if (entry.lat == null && entry.lng == null && lastUserLatitude != null && lastUserLongitude != null) {
    entry.lat = lastUserLatitude;
    entry.lng = lastUserLongitude;
  }

  // If we still have no coordinates, treat this as \"no location yet\" and don't keep a photo reference
  if (entry.lat == null || entry.lng == null) {
    entry.photo_url = null;
    entry.photo_taken_at = null;
  }

  if (!entry.business_name || !barangay || !line || !entry.date_inspected) {
    logbookShowToast(
      "inspection-toast",
      "⚠️ Please fill in at least Business name, Barangay, House/Street, and Date inspected."
    );
    return;
  }

  const isOnline = isSupabaseEnabled();

  // Optimistic local update so the UI responds immediately
  if (inspectionEditingIdx !== null) {
    const prev = inspectionData[inspectionEditingIdx] || {};
    inspectionData[inspectionEditingIdx] = {
      ...prev,
      ...entry,
      id: prev.id || null,
      created_at: prev.created_at || entry.created_at,
    };
  } else {
    inspectionData.push({
      ...entry,
      id: null,
    });
  }
  inspectionSaveToLocal();
  inspectionRenderTable();
  addInspectionMarkerFromEntry(entry);
  inspectionCloseModal();
  showSaveIndicator("Inspection record saved");

  if (!isOnline) {
    logbookShowToast(
      "inspection-toast",
      "Saved on this device only (offline mode)."
    );
    return;
  }

  (async () => {
    try {
      // If we have a file and Supabase Storage, upload to the 'storage' bucket
      if (currentExifFile && supabaseClient?.storage) {
        let uploadFile = currentExifFile;
        if (uploadFile.size > 1024 * 1024) {
          try {
            uploadFile = await compressInspectionImage(uploadFile);
          } catch (compressErr) {
            console.warn("Image compression failed, using original file:", compressErr);
            uploadFile = currentExifFile;
          }
        }

        const fileExt =
          (currentExifFile.name && currentExifFile.name.split(".").pop()) ||
          "jpg";
        const safeExt = fileExt.toLowerCase().replace(/[^a-z0-9]/g, "") || "jpg";
        const path = `inspection-photos/${Date.now()}-${Math.random()
          .toString(36)
          .slice(2)}.${safeExt}`;
        try {
          const { error: uploadError } = await supabaseClient.storage
            .from("storage")
            .upload(path, uploadFile, {
              cacheControl: "3600",
              upsert: false,
              contentType: uploadFile.type || currentExifFile.type || "image/jpeg",
            });
          if (!uploadError) {
            const { data: urlData } = supabaseClient.storage
              .from("storage")
              .getPublicUrl(path);
            if (urlData?.publicUrl) {
              entry.photo_url = urlData.publicUrl;
            }
          } else {
            console.warn("Photo upload failed:", uploadError);
          }
        } catch (uploadErr) {
          console.warn("Photo upload threw error:", uploadErr);
        }
      }

      const payload = {
        io_number: entry.io_number,
        owner_name: entry.insp_owner,
        business_name: entry.business_name,
        address: entry.insp_address,
        date_inspected: entry.date_inspected,
        fsic_number: entry.fsic_number,
        inspected_by: entry.inspected_by || null,
        latitude: entry.lat ?? null,
        longitude: entry.lng ?? null,
        photo_url: entry.photo_url ?? null,
        photo_taken_at: entry.photo_taken_at ?? null,
      };
      const q = supabaseClient.from("inspection_logbook");

      const runWrite = async (data) =>
        inspectionEditingId
          ? await q.update(data).eq("id", inspectionEditingId)
          : await q.insert(data);

      let { error } = await runWrite(payload);
      if (error) {
        const msg = (error?.message || String(error)).toLowerCase();
        const missingGeo = msg.includes("latitude") || msg.includes("longitude");
        if (missingGeo) {
          const payloadNoGeo = {
            io_number: entry.io_number,
            owner_name: entry.insp_owner,
            business_name: entry.business_name,
            address: entry.insp_address,
            date_inspected: entry.date_inspected,
            fsic_number: entry.fsic_number,
            inspected_by: entry.inspected_by || null,
          };
          const retry = await runWrite(payloadNoGeo);
          if (retry.error) throw retry.error;
        } else {
          throw error;
        }
      }
      await inspectionLoadFromSupabase();
      inspectionRenderTable();
      // Place a marker based on the coordinates available on this entry.
      addInspectionMarkerFromEntry(entry);
      inspectionCloseModal();
      logbookShowToast("inspection-toast", "Saved to database.");
    } catch (err) {
      const msg = err?.message || String(err);
      const hint =
        msg.includes("policy") ||
        msg.includes("RLS") ||
        err?.code === "42501"
          ? " Check Supabase: add anon RLS policy (see fsis.logger.sql)."
          : "";
      logbookShowToast(
        "inspection-toast",
        "Save failed: " + msg + hint
      );
    }
  })();
}

function inspectionSetPrintDate() {
  const el = document.getElementById("inspection-print-date");
  if (!el) return;
  el.textContent = new Date().toLocaleDateString("en-PH", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function openInspectionDetailPanel(entry) {
  const panel = document.getElementById("map-detail-panel");
  if (!panel) return;

  const businessEl = document.getElementById("map-detail-business");
  const addressEl = document.getElementById("map-detail-address");
  const coordsEl = document.getElementById("map-detail-coords");
  const dateEl = document.getElementById("map-detail-date");
  const takenEl = document.getElementById("map-detail-taken");
  const inspectorEl = document.getElementById("map-detail-inspector");
  const photoWrap = document.getElementById("map-detail-photo-wrapper");
  const photoImg = document.getElementById("map-detail-photo");
  const directionsLink = document.getElementById("map-detail-directions");
  const copyCoordsBtn = document.getElementById("map-detail-copy-coords");

  if (businessEl) {
    businessEl.textContent =
      entry.business_name || entry.insp_owner || entry.io_number || "Inspection";
  }
  if (addressEl) {
    addressEl.textContent = inspectionFormatAddressDisplay(entry);
  }
  if (coordsEl) {
    if (entry.lat != null && entry.lng != null) {
      coordsEl.textContent = `${entry.lat.toFixed(6)}, ${entry.lng.toFixed(6)}`;
    } else {
      coordsEl.textContent = "—";
    }
  }
  if (dateEl) {
    dateEl.textContent = logbookFormatDate(entry.date_inspected);
  }
  if (takenEl) {
    takenEl.textContent = entry.photo_taken_at ? String(entry.photo_taken_at) : "—";
    takenEl.classList.toggle("meta-muted", !entry.photo_taken_at);
  }
  if (inspectorEl) {
    inspectorEl.textContent = entry.inspected_by || "—";
  }

  if (directionsLink) {
    if (entry.lat != null && entry.lng != null) {
      const dest = `${entry.lat},${entry.lng}`;
      const url = `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(
        dest
      )}&travelmode=driving`;
      directionsLink.href = url;
      directionsLink.style.display = "";
    } else {
      directionsLink.removeAttribute("href");
      directionsLink.style.display = "none";
    }
  }

  if (copyCoordsBtn) {
    if (entry.lat != null && entry.lng != null) {
      copyCoordsBtn.style.display = "";
      copyCoordsBtn.onclick = async () => {
        const text = `${entry.lat.toFixed(6)}, ${entry.lng.toFixed(6)}`;
        try {
          await navigator.clipboard.writeText(text);
          logbookShowToast("inspection-toast", "Coordinates copied.");
        } catch {
          // Fallback
          try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
            logbookShowToast("inspection-toast", "Coordinates copied.");
          } catch {
            logbookShowToast("inspection-toast", "Copy failed.");
          }
        }
      };
    } else {
      copyCoordsBtn.style.display = "none";
      copyCoordsBtn.onclick = null;
    }
  }

  if (photoWrap && photoImg) {
    if (entry.photo_url) {
      photoImg.src = entry.photo_url;
      photoImg.alt = entry.business_name || "Inspection photo";
      photoWrap.style.display = "";
    } else {
      photoWrap.style.display = "none";
      photoImg.removeAttribute("src");
      photoImg.alt = "";
    }
  }

  panel.classList.add("is-open");
  panel.setAttribute("aria-hidden", "false");
}

function closeInspectionDetailPanel() {
  const panel = document.getElementById("map-detail-panel");
  if (!panel) return;
  panel.classList.remove("is-open");
  panel.setAttribute("aria-hidden", "true");
}

function initInspectionPhotoExif() {
  const input = document.getElementById("inspection_photo");
  if (!input || !window.EXIF) return;

  input.addEventListener("change", () => {
    currentExifLat = null;
    currentExifLng = null;
    currentExifPreviewUrl = null;
    currentExifTakenAt = null;
    currentExifFile = null;

    const file = input.files && input.files[0];
    if (!file) return;
    currentExifFile = file;

    const reader = new FileReader();
    reader.onload = function (e) {
      const dataUrl = e.target?.result;
      if (typeof dataUrl === "string") {
        currentExifPreviewUrl = dataUrl;
      }

      const img = new Image();
      img.onload = function () {
        try {
          window.EXIF.getData(img, function () {
            const lat = window.EXIF.getTag(this, "GPSLatitude");
            const latRef = window.EXIF.getTag(this, "GPSLatitudeRef");
            const lng = window.EXIF.getTag(this, "GPSLongitude");
            const lngRef = window.EXIF.getTag(this, "GPSLongitudeRef");
            const takenAt = window.EXIF.getTag(this, "DateTimeOriginal");

            if (takenAt) {
              currentExifTakenAt = takenAt;
            }

            if (!lat || !lng || !latRef || !lngRef) return;

            currentExifLat = dmsToDecimal(lat, latRef);
            currentExifLng = dmsToDecimal(lng, lngRef);
          });
        } catch (err) {
          console.error("Failed to read EXIF data", err);
        }
      };
      img.src = dataUrl;
    };
    reader.readAsDataURL(file);
  });
}

async function compressInspectionImage(file) {
  return new Promise((resolve, reject) => {
    try {
      if (!window.FileReader || !document.createElement) {
        resolve(file);
        return;
      }

      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          try {
            const maxDimension = 1600;
            const scale =
              Math.max(img.width, img.height) > maxDimension
                ? maxDimension / Math.max(img.width, img.height)
                : 1;

            const canvas = document.createElement("canvas");
            canvas.width = Math.round(img.width * scale);
            canvas.height = Math.round(img.height * scale);
            const ctx = canvas.getContext("2d");
            if (!ctx) {
              resolve(file);
              return;
            }
            ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

            canvas.toBlob(
              (blob) => {
                if (!blob || blob.size >= file.size) {
                  resolve(file);
                  return;
                }
                const compressed = new File([blob], file.name, {
                  type: blob.type || "image/jpeg",
                });
                resolve(compressed);
              },
              "image/jpeg",
              0.75
            );
          } catch (err) {
            console.warn("Canvas compression error:", err);
            resolve(file);
          }
        };
        img.onerror = () => resolve(file);
        img.src = e.target?.result;
      };
      reader.onerror = () => resolve(file);
      reader.readAsDataURL(file);
    } catch (err) {
      reject(err);
    }
  });
}

function dmsToDecimal(dms, ref) {
  if (!dms || dms.length !== 3) return null;
  const deg = dms[0];
  const min = dms[1];
  const sec = dms[2];
  const sign = ref === "S" || ref === "W" ? -1 : 1;
  const value = sign * (deg + min / 60 + sec / 3600);
  return isFinite(value) ? value : null;
}

function addInspectionMarkerFromEntry(entry) {
  if (!mapInstance || !inspectionMarkersLayer) return;
  if (entry.lat == null || entry.lng == null) return;

  let icon;
  if (entry.photo_url) {
    icon = L.divIcon({
      className: "inspection-marker inspection-marker--photo",
      html: `<div class="inspection-marker-thumb" style="background-image:url('${entry.photo_url}')"></div>`,
      iconSize: [40, 40],
      iconAnchor: [20, 40],
      popupAnchor: [0, -36],
    });
  } else {
    icon = L.divIcon({
      className: "inspection-marker inspection-marker--default",
      html: '<div class="inspection-marker-pin"></div>',
      iconSize: [28, 36],
      iconAnchor: [14, 36],
      popupAnchor: [0, -30],
    });
  }

  const marker = L.marker([entry.lat, entry.lng], { icon }).addTo(
    inspectionMarkersLayer
  );

  marker.on("click", () => {
    openInspectionDetailPanel(entry);
  });
}

const INSPECTION_MARKER_BATCH_SIZE = 40;

function renderInspectionMarkersBatched() {
  if (!mapInstance || !inspectionMarkersLayer || !Array.isArray(inspectionData)) return;
  inspectionMarkersLayer.clearLayers();
  const withCoords = inspectionData.filter((row) => row.lat != null && row.lng != null);
  let i = 0;
  function addBatch() {
    const end = Math.min(i + INSPECTION_MARKER_BATCH_SIZE, withCoords.length);
    for (; i < end; i++) addInspectionMarkerFromEntry(withCoords[i]);
    if (i < withCoords.length) requestAnimationFrame(addBatch);
  }
  addBatch();
}

function inspectionPrintPanel() {
  inspectionSetPrintDate();
  const oldTitle = document.title;
  document.title = "";
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.title = oldTitle;
    }, 500);
  }, 0);
}

async function inspectionLoadFromSupabase() {
  const selectWithGeo =
    "id, io_number, owner_name, business_name, address, date_inspected, fsic_number, inspected_by, latitude, longitude, photo_url, photo_taken_at, created_at";
  const selectWithoutGeo =
    "id, io_number, owner_name, business_name, address, date_inspected, fsic_number, inspected_by, created_at";

  const INSPECTION_FETCH_LIMIT = 2000;
  const run = async (select) =>
    await supabaseClient
      .from("inspection_logbook")
      .select(select)
      .order("created_at", { ascending: true })
      .limit(INSPECTION_FETCH_LIMIT);

  let rows;
  {
    const { data, error } = await run(selectWithGeo);
    if (!error) rows = data;
    else {
      const msg = (error?.message || String(error)).toLowerCase();
      const missingGeo = msg.includes("latitude") || msg.includes("longitude");
      if (!missingGeo) throw error;

      // Backward compatibility: database exists but hasn't been migrated yet
      const retry = await run(selectWithoutGeo);
      if (retry.error) throw retry.error;
      rows = retry.data;
    }
  }

  inspectionData = (rows || []).map((r) => ({
    id: r.id,
    io_number: r.io_number,
    fsic_number: r.fsic_number,
    insp_owner: r.owner_name,
    business_name: r.business_name,
    insp_address: r.address,
    date_inspected: r.date_inspected,
    inspected_by: r.inspected_by || "",
    lat: r.latitude ?? null,
    lng: r.longitude ?? null,
    photo_url:
      r.latitude != null && r.longitude != null ? r.photo_url ?? null : null,
    photo_taken_at:
      r.latitude != null && r.longitude != null ? r.photo_taken_at ?? null : null,
    created_at: r.created_at,
  }));
}

function inspectionInitData() {
  if (inspectionDataLoaded) return;
  inspectionDataLoaded = true;

  // Show cached data immediately so map and table feel instant
  inspectionLoadFromLocal();
  inspectionSetPrintDate();
  inspectionRenderTable();
  renderInspectionMarkersBatched();
  setInspectionTab(inspectionActiveTab);

  if (!isSupabaseEnabled()) return;

  // Refresh from database in background; update UI when done
  (async () => {
    try {
      await inspectionLoadFromSupabase();
      inspectionSetPrintDate();
      inspectionRenderTable();
      renderInspectionMarkersBatched();
      setInspectionTab(inspectionActiveTab);
    } catch (err) {
      console.warn("Inspection refresh from database failed:", err);
      logbookShowToast("inspection-toast", "Using cached data. Could not refresh from server.");
    }
  })();
}

// -----------------------------
// FSEC logbook module
// -----------------------------

const FSEC_STORAGE_KEY = "bfp_fsec";
let fsecData = [];
let fsecEditingIdx = null;
let fsecEditingId = null;
let fsecDataLoaded = false;

function fsecLoadFromLocal() {
  fsecData = JSON.parse(localStorage.getItem(FSEC_STORAGE_KEY) || "[]");
}

function fsecSaveToLocal() {
  localStorage.setItem(FSEC_STORAGE_KEY, JSON.stringify(fsecData));
}

function fsecSave() {
  if (!isSupabaseEnabled()) fsecSaveToLocal();
}

function fsecFormatAddressDisplay(row) {
  const addrLine = row.addr_line;
  const addrBarangay = row.addr_barangay;
  if (
    addrLine != null &&
    addrBarangay != null &&
    (addrLine !== "" || addrBarangay !== "")
  ) {
    return [
      addrLine,
      addrBarangay,
      "Manolo Fortich",
      "Bukidnon",
      "Region X",
    ]
      .filter(Boolean)
      .join(", ");
  }

  const full = (row.fsec_address || "").toString().trim();
  if (!full) return "—";
  const parts = full.split(/,\s*/);
  if (parts.length >= 5) {
    const line = parts[4];
    const barangay = (parts[3] || "")
      .replace(/^Barangay\s+/i, "")
      .trim();
    const municipal = parts[2] || "";
    const province = parts[1] || "";
    const region = parts[0] || "";
    return [line, barangay, municipal, province, region]
      .filter(Boolean)
      .join(", ");
  }
  return full;
}

function fsecRenderTable() {
  const tbody = document.getElementById("tbody-fsec");
  const empty = document.getElementById("empty-fsec");
  const tableWrap = document.getElementById("table-fsec")?.closest(".table-wrap");
  if (!tbody || !empty) return;

  tbody.innerHTML = "";
  if (fsecData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  empty.style.display = "none";
  if (tableWrap) tableWrap.style.display = "";
  fsecData.forEach((row, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td data-label="#">${i + 1}</td>
      <td data-label="Name of Owner">${logbookEsc(row.fsec_owner)}</td>
      <td data-label="Proposed Project"><strong>${logbookEsc(row.proposed_project)}</strong></td>
      <td data-label="Address">${logbookEsc(fsecFormatAddressDisplay(row))}</td>
      <td class="td-date" data-label="Date">${logbookFormatDate(row.fsec_date)}</td>
      <td data-label="Contact Number">${logbookEsc(row.contact_number)}</td>
      <td class="col-action" data-label="Action">
        <div class="tbl-actions">
          <button class="btn-edit" onclick="fsecEditEntry(${i})">Edit</button>
          <button class="btn-del" onclick="fsecDeleteEntry(${i})">Delete</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function fsecEditEntry(idx) {
  const row = fsecData[idx];
  if (!row) return;

  fsecEditingIdx = idx;
  fsecEditingId = row.id || null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };

  setVal("fsec_owner", row.fsec_owner);
  setVal("proposed_project", row.proposed_project);
  setVal("fsec_date", row.fsec_date);
  setVal("contact_number", row.contact_number);

  const addr = (row.fsec_address || "").toString();
  const brgyMatch = addr.match(/Barangay\s+([^,]+)/i);
  setVal(
    "fsec_addr_barangay",
    row.addr_barangay || (brgyMatch ? brgyMatch[1].trim() : "")
  );
  setVal("fsec_addr_line", row.addr_line || "");

  const overlay = document.getElementById("fsec-modal-overlay");
  if (overlay) overlay.classList.add("open");
  setText("fsec-modal-title", "Edit FSEC Building Plan Record");
  setText("fsec-modal-subtitle", "FSEC Building Plan Logbook");
  const btn = document.getElementById("fsec-btn-save");
  if (btn) btn.textContent = "Update Record";
}

function fsecDeleteEntry(idx) {
  if (!confirm("Delete this record?")) return;

  if (!isSupabaseEnabled()) {
    fsecData.splice(idx, 1);
    fsecSave();
    fsecRenderTable();
    logbookShowToast("fsec-toast", "Record deleted.");
    return;
  }

  const row = fsecData[idx];
  if (!row?.id) {
    logbookShowToast("fsec-toast", "⚠️ Cannot delete: missing record id.");
    return;
  }

  (async () => {
    try {
      const { error } = await supabaseClient
        .from("fsec_building_plan_logbook")
        .delete()
        .eq("id", row.id);
      if (error) throw error;
      await fsecLoadFromSupabase();
      fsecRenderTable();
      logbookShowToast("fsec-toast", "Record deleted.");
    } catch (err) {
      logbookShowToast(
        "fsec-toast",
        "⚠️ Delete failed: " + (err?.message || err)
      );
    }
  })();
}

function fsecOpenModal() {
  fsecEditingIdx = null;
  fsecEditingId = null;

  const overlay = document.getElementById("fsec-modal-overlay");
  if (overlay) overlay.classList.add("open");

  setText("fsec-modal-title", "Add FSEC Building Plan Record");
  setText("fsec-modal-subtitle", "FSEC Building Plan Logbook");
  const btn = document.getElementById("fsec-btn-save");
  if (btn) btn.textContent = "Save Record";

  fsecClearForm();
  const date = document.getElementById("fsec_date");
  if (date) date.value = new Date().toISOString().slice(0, 10);
}

function fsecCloseModal() {
  const overlay = document.getElementById("fsec-modal-overlay");
  if (overlay) overlay.classList.remove("open");
}

function fsecCloseOnOverlay(e) {
  const overlay = document.getElementById("fsec-modal-overlay");
  if (e.target === overlay) fsecCloseModal();
}

function fsecClearForm() {
  [
    "fsec_owner",
    "proposed_project",
    "fsec_addr_barangay",
    "fsec_addr_line",
    "fsec_date",
    "contact_number",
  ].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
}

function fsecSaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();

  const region =
    (document.getElementById("fsec_addr_region")?.value || "X").trim();
  const province =
    (document.getElementById("fsec_addr_province")?.value || "Bukidnon").trim();
  const municipal =
    (document.getElementById("fsec_addr_municipal")?.value ||
      "Manolo Fortich").trim();
  const barangay = (
    document.getElementById("fsec_addr_barangay") || { value: "" }
  ).value.trim();
  const line = (
    document.getElementById("fsec_addr_line") || { value: "" }
  ).value.trim();

  const mergedAddress = `Region ${region}, ${province}, ${municipal}, Barangay ${barangay}, ${line}`;

  const entry = {
    fsec_owner: (document.getElementById("fsec_owner") || { value: "" }).value.trim(),
    proposed_project: (document.getElementById("proposed_project") || { value: "" }).value.trim(),
    fsec_address: mergedAddress,
    addr_barangay: barangay,
    addr_line: line,
    fsec_date: (document.getElementById("fsec_date") || { value: "" }).value,
    contact_number: (document.getElementById("contact_number") || { value: "" }).value.trim(),
    created_at: new Date().toISOString(),
  };

  if (
    !entry.fsec_owner ||
    !entry.proposed_project ||
    !barangay ||
    !line ||
    !entry.fsec_date ||
    !entry.contact_number
  ) {
    logbookShowToast(
      "fsec-toast",
      "⚠️ Please fill in all required fields."
    );
    return;
  }

  const isOnline = isSupabaseEnabled();

  // Optimistic local update so the UI responds immediately
  if (fsecEditingIdx !== null) {
    const prev = fsecData[fsecEditingIdx] || {};
    fsecData[fsecEditingIdx] = {
      ...prev,
      ...entry,
      id: prev.id || null,
      created_at: prev.created_at || entry.created_at,
    };
  } else {
    fsecData.push({
      ...entry,
      id: null,
    });
  }
  fsecSaveToLocal();
  fsecRenderTable();
  fsecCloseModal();
  showSaveIndicator("FSEC record saved");

  if (!isOnline) {
    logbookShowToast(
      "fsec-toast",
      "Saved on this device only (offline mode)."
    );
    return;
  }

  (async () => {
    try {
      const payload = {
        owner_name: entry.fsec_owner,
        proposed_project: entry.proposed_project,
        address: entry.fsec_address,
        date: entry.fsec_date,
        contact_number: entry.contact_number,
      };
      const q = supabaseClient.from("fsec_building_plan_logbook");
      const { error } = fsecEditingId
        ? await q.update(payload).eq("id", fsecEditingId)
        : await q.insert(payload);
      if (error) throw error;
      // Database write completes in the background; UI was already updated optimistically
      logbookShowToast("fsec-toast", "Saved to database.");
    } catch (err) {
      const msg = err?.message || String(err);
      const hint =
        msg.includes("policy") ||
        msg.includes("RLS") ||
        err?.code === "42501"
          ? " Check Supabase: add anon RLS policy (see fsis.logger.sql)."
          : "";
      logbookShowToast("fsec-toast", "Save failed: " + msg + hint);
    }
  })();
}

function fsecSetPrintDate() {
  const el = document.getElementById("fsec-print-date");
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleDateString("en-PH", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function fsecPrintPanel() {
  fsecSetPrintDate();
  const oldTitle = document.title;
  document.title = "";
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.title = oldTitle;
    }, 500);
  }, 0);
}

async function fsecLoadFromSupabase() {
  const { data: rows, error } = await supabaseClient
    .from("fsec_building_plan_logbook")
    .select(
      "id, owner_name, proposed_project, address, date, contact_number, created_at"
    )
    .order("created_at", { ascending: true });
  if (error) throw error;
  fsecData = (rows || []).map((r) => ({
    id: r.id,
    fsec_owner: r.owner_name,
    proposed_project: r.proposed_project,
    fsec_address: r.address,
    fsec_date: r.date,
    contact_number: r.contact_number,
    created_at: r.created_at,
  }));
}

async function fsecInitData() {
  if (fsecDataLoaded) return;
  fsecDataLoaded = true;
  try {
    fsecLoadFromLocal();
    fsecSetPrintDate();
    fsecRenderTable();
    if (isSupabaseEnabled()) {
      await fsecLoadFromSupabase();
      fsecSetPrintDate();
      fsecRenderTable();
    }
  } catch (err) {
    fsecLoadFromLocal();
    console.warn("FSEC load failed, using local storage:", err);
    logbookShowToast("fsec-toast", "Using data stored on this device.");
    fsecSetPrintDate();
    fsecRenderTable();
  }
}
