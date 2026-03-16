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

function ensureSelectOption(id, value) {
  if (!value) return;
  const el = document.getElementById(id);
  if (!el || el.tagName !== "SELECT") return;
  const opts = Array.from(el.options).map((o) => o.value);
  if (!opts.includes(value)) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    el.appendChild(opt);
  }
  el.value = value;
}

function toFriendlyDate(isoString) {
  if (!isoString) return "—";
  const d = new Date(isoString);
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleString();
}

function getCurrentView() {
  const hash = (window.location.hash || "#map").replace(/^#/, "");
  if (
    hash === "inspection" ||
    hash === "fsec" ||
    hash === "conveyance" ||
    hash === "occupancy" ||
    hash === "map"
  )
    return hash;
  return "map";
}

const BARANGAYS = [
  "Agusan Canyon", "Alae", "Dahilayan", "Dalirig", "Damilag", "Diclum",
  "Guilang-guilang", "Kalugmanan", "Lindaban", "Lingion", "Lunocan", "Maluko",
  "Mambatangan", "Mampayag", "Minsuro", "Mantibugao", "Tankulan (Poblacion)",
  "San Miguel", "Sankanan", "Santiago", "Santo Niño", "Ticala",
];

const FIRE_PERSONNEL_META = [
  { name: "SF03 Rafael I. Corona Jr", rank: "SF03" },
  { name: "SF01 Mark Ferdinand B. Cariaga", rank: "SF01" },
  { name: "SF01 Cedric B. Gamolo", rank: "SF01" },
  { name: "FO3 Rey Edward S. Descallar", rank: "FO3" },
  { name: "FO3 Jun Ray D. Abarquez", rank: "FO3" },
  { name: "FO3 Clyde Q. Rejas", rank: "FO3" },
  { name: "FO3 Juan M. Derayunan II", rank: "FO3" },
  { name: "FO3 Julious G. Cloma", rank: "FO3" },
  { name: "FO3 Luigi C. Cajes", rank: "FO3" },
  { name: "FO2 Michael S. Guyan", rank: "FO2" },
  { name: "FO2 Rhea Mae B. Lambago", rank: "FO2" },
  { name: "FO1 John Ansel P. Labinay", rank: "FO1" },
  { name: "FO1 Jessel Joy C. Paca", rank: "FO1" },
  { name: "FO1 Adoniram C. Nacilla", rank: "FO1" },
  { name: "FO1 Johnremar B. Cinchez", rank: "FO1" },
  { name: "FO1 Moctar M. Manarinta", rank: "FO1" },
  { name: "FO1 Sairah Ville L. Sante", rank: "FO1" },
  { name: "FO1 Lester V. Villarta", rank: "FO1" },
  { name: "Cherry Mae N. Lusno", rank: "Fire Aide" },
];

const FIRE_PERSONNEL = FIRE_PERSONNEL_META.map((p) => p.name);
const FIRE_PERSONNEL_BY_NAME = Object.fromEntries(
  FIRE_PERSONNEL_META.map((p) => [p.name, { rank: p.rank }])
);

function rankCodeToTitle(code) {
  const c = String(code || "").trim().toUpperCase();
  if (c === "FO1") return "Fire Officer I";
  if (c === "FO2") return "Fire Officer II";
  if (c === "FO3") return "Fire Officer III";
  if (c === "SF01" || c === "SFO1") return "Senior Fire Officer I";
  if (c === "SF02" || c === "SFO2") return "Senior Fire Officer II";
  if (c === "SF03" || c === "SFO3") return "Senior Fire Officer III";
  if (c === "SF04" || c === "SFO4") return "Senior Fire Officer IV";
  if (c === "FIRE AIDE") return "Fire Aide";
  return code ? String(code) : "";
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
let currentExifProcessingPromise = null;

// Occupancy photo/EXIF state (separate from inspection)
let occupancyExifLat = null;
let occupancyExifLng = null;
let occupancyExifPreviewUrl = null;
let occupancyExifTakenAt = null;
let occupancyExifFile = null;

let inspectionMarkersLayer = null;
let occupancyMarkersLayer = null;
let inspectionDataLoaded = false;
let inspectionActiveTab = "with-location";
let inspectionFocusMapAfterSave = false;

let mapMarkerFilter = "all"; // all | businesses | residential

function applyMapMarkerFilter(next) {
  mapMarkerFilter = next || "all";
  const buttons = Array.from(document.querySelectorAll("[data-map-filter]"));
  buttons.forEach((b) => b.classList.toggle("is-active", b.getAttribute("data-map-filter") === mapMarkerFilter));

  if (!mapInstance) return;
  const showInspection = mapMarkerFilter === "all" || mapMarkerFilter === "businesses";
  const showOccupancy = mapMarkerFilter === "all" || mapMarkerFilter === "residential";

  if (inspectionMarkersLayer) {
    if (showInspection) inspectionMarkersLayer.addTo(mapInstance);
    else mapInstance.removeLayer(inspectionMarkersLayer);
  }
  if (occupancyMarkersLayer) {
    if (showOccupancy) occupancyMarkersLayer.addTo(mapInstance);
    else mapInstance.removeLayer(occupancyMarkersLayer);
  }
}

function initMapMarkerFilterUi() {
  document.addEventListener("click", (e) => {
    const el = e.target instanceof Element ? e.target : null;
    const btn = el?.closest?.("[data-map-filter]");
    if (!btn) return;
    const f = btn.getAttribute("data-map-filter") || "all";
    applyMapMarkerFilter(f);
  });
}

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
  // Layer to hold occupancy markers (blue) so we can toggle/manage separately
  occupancyMarkersLayer = L.layerGroup().addTo(mapInstance);

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

  // If occupancy data is already loaded, render any markers that have coordinates
  if (Array.isArray(occupancyData) && occupancyData.length > 0) {
    occupancyData.forEach((row) => {
      if (row.lat != null && row.lng != null) {
        addOccupancyMarkerFromEntry(row);
      }
    });
  }

  // Make sure the map fully renders after layout
  setTimeout(() => {
    mapInstance.invalidateSize();
  }, 0);

  initMapSearch();
  applyMapMarkerFilter(mapMarkerFilter);
}

function resetMapView() {
  if (!mapInstance) return;
  mapInstance.setView(MAP_CENTER, MAP_ZOOM);
}

function handleFabAddInspection() {
  openMapAddChooser();
}

function openMapAddChooser() {
  const overlay = document.getElementById("map-add-modal-overlay");
  if (!overlay) return;
  overlay.style.display = "";
  overlay.classList.add("open");
}

function closeMapAddChooser() {
  const overlay = document.getElementById("map-add-modal-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.style.display = "none";
}

function mapAddCloseOnOverlay(e) {
  const overlay = document.getElementById("map-add-modal-overlay");
  if (!overlay) return;
  if (e.target === overlay) closeMapAddChooser();
}

function mapAddChoose(type) {
  closeMapAddChooser();
  if (type === "occupancy") {
    occupancyOpenModal?.();
    return;
  }
  inspectionOpenModal?.();
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
    // Map also shows residential (occupancy) markers
    if (name === "map" && !occupancyDataLoaded) {
      occupancyInitData();
    }
  } else if (name === "fsec" && !fsecDataLoaded) {
    fsecInitData();
  } else if (name === "conveyance" && !conveyanceDataLoaded) {
    conveyanceInitData();
  } else if (name === "occupancy" && !occupancyDataLoaded) {
    occupancyInitData();
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

function populateModalDropdowns() {
  function fillSelect(id, options, placeholder) {
    const el = document.getElementById(id);
    if (!el || !el.options) return;
    el.innerHTML = "";
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = placeholder;
    el.appendChild(opt0);
    for (const o of options) {
      const opt = document.createElement("option");
      opt.value = o;
      opt.textContent = o;
      el.appendChild(opt);
    }
  }
  fillSelect("inspection_addr_barangay", BARANGAYS, "Select barangay");
  fillSelect("fsec_addr_barangay", BARANGAYS, "Select barangay");
  fillSelect("inspection_inspected_by", FIRE_PERSONNEL, "Select inspector");
  fillSelect("inspection_included_personnel_name", FIRE_PERSONNEL, "Select personnel (optional)");
  fillSelect("conveyance_inspector_select", FIRE_PERSONNEL, "Select inspector");
  fillSelect("occupancy_inspected_by", FIRE_PERSONNEL, "Select inspector");
  fillSelect("inspection-filter-barangay", BARANGAYS, "All barangays");
  fillSelect("inspection-filter-personnel", FIRE_PERSONNEL, "All personnel");

  // Auto-fill rank/position when fire personnel is selected
  bindInspectionPersonnelAutoFill();
}

function getFirePersonnelRankPositionByName(name) {
  if (!name || typeof name !== "string") return "";
  const meta = FIRE_PERSONNEL_BY_NAME[name];
  if (!meta) return "";
  return rankCodeToTitle(meta.rank) || "";
}

function bindInspectionPersonnelAutoFill() {
  const inspectedBy = document.getElementById("inspection_inspected_by");
  const inspectorPos = document.getElementById("inspection_inspector_position");
  const includedPersonnel = document.getElementById("inspection_included_personnel_name");
  const includedPos = document.getElementById("inspection_included_personnel_position");

  inspectedBy?.addEventListener("change", () => {
    if (inspectorPos) inspectorPos.value = getFirePersonnelRankPositionByName(inspectedBy.value);
  });
  includedPersonnel?.addEventListener("change", () => {
    if (includedPos) includedPos.value = getFirePersonnelRankPositionByName(includedPersonnel.value);
  });
}

function conveyanceAddInspector() {
  const sel = document.getElementById("conveyance_inspector_select");
  const ta = document.getElementById("conveyance_inspectors");
  if (!sel || !ta || !sel.value) return;
  const current = (ta.value || "").trim();
  const sep = current ? "\n" : "";
  ta.value = current + sep + sel.value;
  sel.selectedIndex = 0;
}

// Legacy helper kept for compatibility; no longer used by the current UI.
function occupancyAddInspector() {}

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

// -----------------------------
// In-app browser detection for file upload warning
// -----------------------------

function isInAppBrowser() {
  const ua = navigator.userAgent || "";
  // Detect common in-app WebViews where file input is restricted
  return (
    /FBAN|FBAV|FB_IAB|FBIOS/.test(ua) ||   // Facebook
    /Instagram/.test(ua) ||                  // Instagram
    /Snapchat/.test(ua) ||                   // Snapchat
    /TikTok/.test(ua) ||                     // TikTok
    /Twitter/.test(ua) ||                    // Twitter/X
    /LinkedInApp/.test(ua) ||                // LinkedIn
    /Line\//.test(ua) ||                     // LINE messenger
    (ua.includes("wv") && ua.includes("Android") && !ua.includes("Chrome/")) // Generic Android WebView
  );
}

function showInAppBrowserBanner() {
  if (!isInAppBrowser()) return;

  // Don't show more than once per session
  if (sessionStorage.getItem("fsis.inapp_banner_dismissed")) return;

  const banner = document.createElement("div");
  banner.id = "inapp-browser-banner";
  banner.style.cssText = [
    "position:fixed",
    "top:0",
    "left:0",
    "right:0",
    "z-index:9999",
    "background:#C1272D",
    "color:#fff",
    "font-family:'DM Sans',sans-serif",
    "font-size:0.85rem",
    "padding:10px 48px 10px 16px",
    "line-height:1.4",
    "box-shadow:0 2px 12px rgba(0,0,0,0.4)",
  ].join(";");

  const pageUrl = window.location.href;
  banner.innerHTML = `
    <strong>⚠️ File uploads may not work in this browser.</strong>
    Open this page in your system browser (Chrome/Safari) for full functionality.
    <a href="${pageUrl}" target="_blank" rel="noopener"
       style="color:#fde68a;font-weight:700;text-decoration:underline;margin-left:6px;">
      Open in browser ↗
    </a>
    <button onclick="this.parentElement.remove();sessionStorage.setItem('fsis.inapp_banner_dismissed','1')"
      style="position:absolute;top:50%;right:10px;transform:translateY(-50%);background:none;border:none;color:#fff;font-size:1.2rem;cursor:pointer;line-height:1;padding:4px 6px;"
      aria-label="Dismiss">×</button>
  `;
  document.body.prepend(banner);
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

  // Warn users accessing the app from in-app browsers (e.g. Facebook, Messenger)
  // where file input is often blocked or restricted.
  showInAppBrowserBanner();

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

  populateModalDropdowns();
  initViewRouting();
  initTableFilters();
  initMapMarkerFilterUi();
  initInspectionPhotoExif();
  initOccupancyPhotoExif();
  refreshStorageBadge();

  // Mobile safety net: ensure Save button always triggers save handler.
  // Some mobile browsers can drop inline handlers in certain contexts.
  const saveBtn = document.getElementById("inspection-btn-save");
  if (saveBtn) {
    saveBtn.addEventListener("click", (ev) => void inspectionSaveEntry(ev));
    // touchend can fire without click on some devices
    saveBtn.addEventListener(
      "touchend",
      (ev) => {
        try { ev.preventDefault(); } catch {}
        void inspectionSaveEntry(ev);
      },
      { passive: false }
    );
  }

  // Surface runtime errors on mobile (otherwise it looks like "Save does nothing").
  window.addEventListener("error", (ev) => {
    const msg = ev?.error?.message || ev?.message || "Unknown error";
    logbookShowToast("inspection-toast", "⚠️ Error: " + msg);
  });
  window.addEventListener("unhandledrejection", (ev) => {
    const reason = ev?.reason;
    const msg = reason?.message || String(reason || "Unknown error");
    logbookShowToast("inspection-toast", "⚠️ Error: " + msg);
  });

  const initialView = getCurrentView();
  if ((initialView === "map" || initialView === "inspection") && !inspectionDataLoaded) {
    inspectionInitData();
    if (initialView === "map" && !occupancyDataLoaded) {
      occupancyInitData();
    }
    if (initialView === "inspection") {
      setInspectionTab(inspectionActiveTab);
    }
  } else if (initialView === "fsec" && !fsecDataLoaded) {
    fsecInitData();
  } else if (initialView === "conveyance" && !conveyanceDataLoaded) {
    conveyanceInitData();
  } else if (initialView === "occupancy" && !occupancyDataLoaded) {
    occupancyInitData();
  }
}

document.addEventListener("DOMContentLoaded", init);

// -----------------------------
// Shared Supabase + utilities
// -----------------------------

const SUPABASE_URL = "https://ezpwbgpbveazutmrnzlf.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImV6cHdiZ3BidmVhenV0bXJuemxmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE4Nzk2NDksImV4cCI6MjA4NzQ1NTY0OX0.C5118CQPYAqay0FhtmKdJyl9LKUHFzMnN5ecnAx1NU8";

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

function normalizeQuery(s) {
  return String(s || "").toLowerCase().trim();
}

function inDateRange(dateStr, fromStr, toStr) {
  if (!dateStr) return false;
  const t = new Date(dateStr + "T00:00:00").getTime();
  if (!isFinite(t)) return false;
  if (fromStr) {
    const f = new Date(fromStr + "T00:00:00").getTime();
    if (isFinite(f) && t < f) return false;
  }
  if (toStr) {
    const to = new Date(toStr + "T00:00:00").getTime();
    if (isFinite(to) && t > to) return false;
  }
  return true;
}

function initTableFilters() {
  const debounce = (fn, wait = 150) => {
    let t = null;
    return (...args) => {
      if (t) clearTimeout(t);
      t = setTimeout(() => fn(...args), wait);
    };
  };

  const bind = (ids, onChange) => {
    const run = debounce(onChange, 120);
    for (const id of ids) {
      const el = document.getElementById(id);
      if (!el) continue;
      el.addEventListener("input", run);
      el.addEventListener("change", run);
    }
  };

  bind(
    [
      "inspection-filter-q",
      "inspection-filter-barangay",
      "inspection-filter-personnel",
      "inspection-filter-from",
      "inspection-filter-to",
    ],
    () => inspectionRenderTable()
  );
  bind(["fsec-filter-q", "fsec-filter-from", "fsec-filter-to"], () =>
    fsecRenderTable()
  );
  bind(
    ["conveyance-filter-q", "conveyance-filter-from", "conveyance-filter-to"],
    () => conveyanceRenderTable()
  );
  bind(
    ["occupancy-filter-q", "occupancy-filter-from", "occupancy-filter-to"],
    () => occupancyRenderTable()
  );
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

// Short address for table display — strips the repeated municipality/province/region
// to keep the address column compact. Full version still used for print and detail panels.
function inspectionFormatAddressShort(row) {
  const addrLine = (row.addr_line || "").trim();
  const addrBarangay = (row.addr_barangay || "").trim();
  if (addrLine || addrBarangay) {
    const parts = [];
    if (addrLine) parts.push(addrLine);
    if (addrBarangay) parts.push("Brgy. " + addrBarangay);
    return parts.join(", ") || "—";
  }
  // Legacy free-text fallback: just use the raw string (it's already compact)
  const full = (row.insp_address || "").toString().trim();
  return full || "—";
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

  const q = normalizeQuery(document.getElementById("inspection-filter-q")?.value);
  const brgy = (document.getElementById("inspection-filter-barangay")?.value || "").trim();
  const personnel = (document.getElementById("inspection-filter-personnel")?.value || "").trim();
  const from = (document.getElementById("inspection-filter-from")?.value || "").trim();
  const to = (document.getElementById("inspection-filter-to")?.value || "").trim();

  const filtered = inspectionData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (from || to) {
        if (!inDateRange(row.date_inspected, from, to)) return false;
      }
      if (brgy) {
        const rowBrgy = (row.addr_barangay || "").toString().trim();
        if (rowBrgy !== brgy) return false;
      }
      if (personnel) {
        const rowP = (row.inspected_by || "").toString().trim();
        if (rowP !== personnel) return false;
      }
      if (!q) return true;
      const hay = normalizeQuery(
        [
          row.io_number,
          row.insp_owner,
          row.insp_owner_phone,
          row.business_name,
          inspectionFormatAddressDisplay(row),
          row.fsic_number,
          row.inspected_by,
        ].join(" | ")
      );
      return hay.includes(q);
    });

  if (inspectionData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
  } else {
    empty.style.display = "none";
    if (tableWrap) tableWrap.style.display = "";

    filtered.forEach(({ row, idx }, displayIdx) => {
      const hasLocation = row.lat != null && row.lng != null;

      const baseRowHtml = `
        <td data-label="#">${displayIdx + 1}</td>
        <td class="td-io" data-label="IO Number">${logbookEsc(row.io_number)}</td>
        <td data-label="Name of Owner">${logbookEsc(row.insp_owner)}</td>
        <td data-label="Owner phone">${logbookEsc(row.insp_owner_phone)}</td>
        <td data-label="Business / Establishment"><strong>${logbookEsc(row.business_name)}</strong></td>
        <td data-label="Address">${logbookEsc(inspectionFormatAddressShort(row))}</td>
        <td class="td-date" data-label="Date Inspected">${logbookFormatDate(row.date_inspected)}</td>
        <td class="td-fsic" data-label="FSIC Number">${logbookEsc(row.fsic_number)}</td>
        <td data-label="Inspected By">${logbookEsc(row.inspected_by)}</td>
      `;

      if (hasLocation) {
        withLocationCount++;
        const tr = document.createElement("tr");
        tr.id = `inspection-row-${idx}`;
        tr.innerHTML = `
          ${baseRowHtml}
          <td class="col-action" data-label="Action">
            <select class="action-select"
              aria-label="Row actions"
              onchange="inspectionHandleAction(this.value, ${idx}); this.selectedIndex = 0;"
            >
              <option value="">Actions…</option>
              <option value="view_on_map">View on map</option>
              <option value="edit">Edit</option>
              <option value="open_io_html">Open IO (HTML)</option>
              <option value="open_clearance_html">Release clearance (FSIC)</option>
              <option value="add_photo">Add photo</option>
              <option value="delete">Delete</option>
            </select>
          </td>
        `;
        tbody.appendChild(tr);
      } else {
        noLocationCount++;
        if (tbodyNoPhoto) {
          const tr2 = document.createElement("tr");
          tr2.id = `inspection-row-${idx}`;
          tr2.innerHTML = `
            ${baseRowHtml}
            <td class="col-action" data-label="Action">
              <select class="action-select"
                aria-label="Row actions"
                onchange="inspectionHandleAction(this.value, ${idx}); this.selectedIndex = 0;"
              >
                <option value="">Actions…</option>
                <option value="edit">Edit</option>
                <option value="open_io_html">Open IO (HTML)</option>
                <option value="open_clearance_html">Release clearance (FSIC)</option>
                <option value="add_photo">Add photo</option>
                <option value="delete">Delete</option>
              </select>
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

  // ── Record count badges ──────────────────────────────────────────────
  const countBadge = document.getElementById("inspection-record-count");
  if (countBadge) countBadge.textContent = withLocationCount;

  const noPhotoBadge = document.getElementById("inspection-nophoto-record-count");
  if (noPhotoBadge) noPhotoBadge.textContent = noLocationCount;

  // ── Filter result info bars ──────────────────────────────────────────
  const isFiltered = !!(q || brgy || personnel || from || to);
  const resultsBadge = document.getElementById("inspection-results-badge");
  if (resultsBadge) {
    if (isFiltered && inspectionData.length > 0) {
      resultsBadge.textContent =
        `Showing ${withLocationCount} of ${inspectionData.filter(r => r.lat != null && r.lng != null).length} records (with location)`;
      resultsBadge.removeAttribute("hidden");
    } else {
      resultsBadge.setAttribute("hidden", "");
    }
  }

  const noPhotoResultsBadge = document.getElementById("inspection-nophoto-results-badge");
  if (noPhotoResultsBadge) {
    if (isFiltered && inspectionData.length > 0) {
      noPhotoResultsBadge.textContent =
        `Showing ${noLocationCount} of ${inspectionData.filter(r => r.lat == null || r.lng == null).length} records (no location)`;
      noPhotoResultsBadge.removeAttribute("hidden");
    } else {
      noPhotoResultsBadge.setAttribute("hidden", "");
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
  setVal("inspection_owner_phone", row.insp_owner_phone);
  setVal("inspection_business_name", row.business_name);
  setVal("inspection_date_inspected", row.date_inspected);
  // Optional IO-specific fields (may not exist on older records or in the DOM)
  setVal("inspection_inspector_position", row.inspector_position);
  setVal("inspection_included_personnel_name", row.included_personnel_name);
  setVal(
    "inspection_included_personnel_position",
    row.included_personnel_position
  );
  setVal("inspection_duration_start", row.duration_start);
  setVal("inspection_duration_end", row.duration_end);

  const addr = (row.insp_address || "").toString();
  const brgyMatch = addr.match(/Barangay\s+([^,]+)/i);
  const barangayVal = row.addr_barangay || (brgyMatch ? brgyMatch[1].trim() : "");
  ensureSelectOption("inspection_addr_barangay", barangayVal);
  setVal("inspection_addr_line", row.addr_line || "");

  ensureSelectOption("inspection_inspected_by", row.inspected_by);
  ensureSelectOption("inspection_included_personnel_name", row.included_personnel_name);

  const overlay = document.getElementById("inspection-modal-overlay");
  if (overlay) overlay.classList.add("open");
  setText("inspection-modal-title", "Edit Inspection Record");
  setText("inspection-modal-subtitle", "Inspection Logbook");
  const btn = document.getElementById("inspection-btn-save");
  if (btn) btn.textContent = "Update Record";
}

function inspectionAddPhoto(idx) {
  inspectionFocusMapAfterSave = true;
  inspectionEditEntry(idx);
  const photoInput = document.getElementById("inspection_photo");
  if (photoInput) {
    photoInput.scrollIntoView({ behavior: "smooth", block: "center" });
    photoInput.focus();
  }
}

function inspectionHandleAction(action, idx) {
  if (!action) return;
  if (action === "view_on_map") {
    inspectionViewOnMap(idx);
    return;
  }
  if (action === "edit") {
    inspectionEditEntry(idx);
    return;
  }
  if (action === "open_io_html") {
    inspectionOpenIoHtml(idx);
    return;
  }
   if (action === "open_clearance_html") {
    inspectionOpenClearanceHtml(idx);
    return;
  }
  if (action === "add_photo") {
    inspectionAddPhoto(idx);
    return;
  }
  if (action === "delete") {
    inspectionDeleteEntry(idx);
  }
}

function inspectionViewOnMap(idx) {
  const row = inspectionData[idx];
  if (!row || row.lat == null || row.lng == null) return;
  showView("map");
  window.location.hash = "map";
  closeNavSidebar();
  setTimeout(() => {
    if (mapInstance) {
      mapInstance.setView([row.lat, row.lng], 16);
      openInspectionDetailPanel(row);
    }
  }, 100);
}

function inspectionOpenIoHtml(idx) {
  const row = inspectionData[idx];
  if (!row) return;
  try {
    sessionStorage.setItem("fsis.io.current", JSON.stringify(row));
  } catch {
    // If sessionStorage is unavailable, we still open the template;
    // it will show a friendly notice instead of data.
  }
  window.open("./inspection_io_fsis.html", "_blank");
}

function inspectionOpenClearanceHtml(idx) {
  const row = inspectionData[idx];
  if (!row) return;
  try {
    sessionStorage.setItem("fsis.clearance.current", JSON.stringify(row));
  } catch {
    // If sessionStorage is unavailable, we still open the template.
  }
  window.open("./fsis_clearance.html", "_blank");
}

// Receive clearance edits from `fsis_clearance.html` and persist them.
window.addEventListener("message", (ev) => {
  const d = ev?.data;
  if (!d || d.type !== "fsis_clearance_save") return;

  const payload = d.payload || {};
  const entryId = d.entryId || null;
  const ioNumber = d.io_number || null;

  function respond(ok, message) {
    try {
      ev.source?.postMessage({ type: "fsis_clearance_save_result", ok, message }, "*");
    } catch {
      // ignore
    }
  }

  function normalizeDate(value) {
    if (!value) return null;
    const s = String(value).trim();
    if (!s) return null;
    // If already ISO date, keep it
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const dt = new Date(s);
    if (Number.isNaN(dt.getTime())) return null;
    return dt.toISOString().slice(0, 10);
  }

  function normalizeAmount(value) {
    if (value == null || value === "") return null;
    const n = typeof value === "number" ? value : Number(String(value).replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  }

  const idx = Array.isArray(inspectionData)
    ? inspectionData.findIndex((r) => (entryId && r?.id === entryId) || (ioNumber && r?.io_number === ioNumber))
    : -1;

  if (idx < 0) {
    respond(false, "Cannot find matching inspection record.");
    return;
  }

  // Build a *partial* update payload: only fields actually included in the
  // incoming message are allowed to update the record. This prevents
  // unrelated fields from being overwritten to null/empty when the UI only
  // edits a subset of fields.
  const updates = {};
  const has = (k) => Object.prototype.hasOwnProperty.call(payload, k);

  // Core fields (allow editing on the clearance template)
  if (has("business_name")) {
    const v = (payload.business_name || "").toString().trim();
    if (v) updates.business_name = v;
  }
  if (has("owner_name")) {
    const v = (payload.owner_name || "").toString().trim();
    if (v) updates.owner_name = v;
  }
  if (has("address")) {
    const v = (payload.address || "").toString().trim();
    if (v) updates.address = v;
  }

  // FSIC clearance / fees
  if (has("fsic_number")) updates.fsic_number = (payload.fsic_number || "").toString().trim();
  if (has("fsic_purpose")) updates.fsic_purpose = payload.fsic_purpose || null;
  if (has("fsic_permit_type")) updates.fsic_permit_type = payload.fsic_permit_type || null;
  if (has("fsic_valid_from")) updates.fsic_valid_from = normalizeDate(payload.fsic_valid_from);
  if (has("fsic_valid_until")) updates.fsic_valid_until = normalizeDate(payload.fsic_valid_until);
  if (has("fsic_fee_amount")) updates.fsic_fee_amount = normalizeAmount(payload.fsic_fee_amount);
  if (has("fsic_fee_or_number")) updates.fsic_fee_or_number = payload.fsic_fee_or_number || null;
  if (has("fsic_fee_date")) updates.fsic_fee_date = normalizeDate(payload.fsic_fee_date);

  // Update in-memory + local cache immediately for responsiveness
  // Update in-memory fields using app naming (insp_owner / insp_address).
  const uiUpdates = { ...updates };
  if (uiUpdates.owner_name != null) {
    uiUpdates.insp_owner = uiUpdates.owner_name;
    delete uiUpdates.owner_name;
  }
  if (uiUpdates.address != null) {
    uiUpdates.insp_address = uiUpdates.address;
    delete uiUpdates.address;
  }
  inspectionData[idx] = { ...inspectionData[idx], ...uiUpdates };
  inspectionSaveToLocal();
  inspectionRenderTable?.();

  if (!isSupabaseEnabled()) {
    logbookShowToast?.("inspection-toast", "Saved on this device only (offline mode).");
    respond(true, "Saved locally (offline mode).");
    return;
  }

  const row = inspectionData[idx];
  if (!row?.id) {
    logbookShowToast?.("inspection-toast", "⚠️ Save failed: missing record id.");
    respond(false, "Missing record id.");
    return;
  }

  (async () => {
    try {
      const { error } = await supabaseClient.from("inspection_logbook").update(updates).eq("id", row.id);
      if (error) throw error;
      logbookShowToast?.("inspection-toast", "Saved to database.");
      respond(true, "");
    } catch (err) {
      const msg = err?.message || String(err);
      logbookShowToast?.("inspection-toast", "Save failed: " + msg);
      respond(false, msg);
    }
  })();
});

async function inspectionDownloadPdf(idx) {
  const row = inspectionData[idx];
  if (!row) return;
  const filename = `inspection-${row.io_number || row.id || "form"}.pdf`;

  // Prefer backend rendering when served over http(s).
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    try {
      const resp = await fetch("./api/io/pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ entry: row, filename }),
      });
      if (resp.ok) {
        const blob = await resp.blob();
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
        return;
      }
    } catch (err) {
      console.warn("Backend PDF generation failed; falling back.", err);
    }
  }

  // Fallback: client-side HTML-to-PDF in a new tab (works even in file:// mode).
  try {
    sessionStorage.setItem("fsis.io.current", JSON.stringify(row));
    sessionStorage.setItem("fsis.io.downloadFilename", filename);
  } catch {
    // If sessionStorage is unavailable, open without auto-download
  }
  window.open("./inspection_io_fsis.html?download=pdf", "_blank");
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
  inspectionFocusMapAfterSave = false;

  // Reset any previously extracted EXIF coordinates and photo data
  currentExifLat = null;
  currentExifLng = null;
  currentExifPreviewUrl = null;
  currentExifTakenAt = null;
  currentExifFile = null;
  currentExifProcessingPromise = null;

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
  const photoLibraryInput = document.getElementById("inspection_photo_library");
  if (photoLibraryInput) photoLibraryInput.value = "";
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
    "inspection_owner_phone",
    "inspection_business_name",
    "inspection_addr_barangay",
    "inspection_addr_line",
    "inspection_date_inspected",
    "inspection_inspected_by",
    "inspection_inspector_position",
    "inspection_included_personnel_name",
    "inspection_included_personnel_position",
    "inspection_duration_start",
    "inspection_duration_end",
  ].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
}

async function inspectionSaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();
  if (inspectionSaveEntry._lastRun && Date.now() - inspectionSaveEntry._lastRun < 800) return;
  inspectionSaveEntry._lastRun = Date.now();

  // Immediate feedback so "Save" never feels dead.
  logbookShowToast("inspection-toast", "Saving...");

  // If the user just attached a photo, GPS extraction may still be running.
  // Wait briefly so we don't save null lat/lng and lose the map pin.
  if (currentExifProcessingPromise) {
    try {
      await Promise.race([
        currentExifProcessingPromise,
        new Promise((resolve) => setTimeout(resolve, 2500)),
      ]);
    } catch {
      // ignore
    }
  }

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

  const mergedAddress = [
    `Region ${region}`,
    province,
    municipal,
    `Barangay ${barangay}`,
    line,
  ]
    .filter((p) => String(p || "").trim())
    .join(", ");

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
    insp_owner_phone: (
      document.getElementById("inspection_owner_phone") || { value: "" }
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
    inspector_position:
      (document.getElementById("inspection_inspector_position") || {
        value: "",
      }).value.trim(),
    included_personnel_name: (
      document.getElementById("inspection_included_personnel_name") || {
        value: "",
      }
    ).value.trim(),
    included_personnel_position: (
      document.getElementById("inspection_included_personnel_position") || {
        value: "",
      }
    ).value.trim(),
    duration_start: (
      document.getElementById("inspection_duration_start") || { value: "" }
    ).value,
    duration_end: (
      document.getElementById("inspection_duration_end") || { value: "" }
    ).value,
    remarks: "",
    // Optional coordinates and photo metadata extracted from EXIF / geolocation
    lat: currentExifLat,
    lng: currentExifLng,
    photo_url: currentExifPreviewUrl,
    photo_taken_at: currentExifTakenAt,
    created_at: new Date().toISOString(),
  };

  // When editing:
  // - If user didn't pick a new photo, keep existing photo URL/meta and coordinates.
  // - If user picked a new photo and EXIF has GPS, USE the photo's coordinates (move pin to where photo was taken).
  // - If EXIF has no GPS, keep existing lat/lng so the pin doesn't disappear.
  if (inspectionEditingIdx !== null) {
    const prev = inspectionData[inspectionEditingIdx] || {};
    const hasNewPhoto = !!currentExifFile;

    if (!hasNewPhoto) {
      // No new file: keep previous photo + coordinates.
      entry.photo_url = prev.photo_url ?? null;
      entry.photo_taken_at = prev.photo_taken_at ?? null;
      entry.lat = prev.lat ?? entry.lat;
      entry.lng = prev.lng ?? entry.lng;
    } else {
      // New photo chosen:
      const hasExifGps =
        Number.isFinite(currentExifLat) && Number.isFinite(currentExifLng);
      if (hasExifGps) {
        // EXIF GPS present: move pin to photo capture location.
        entry.lat = currentExifLat;
        entry.lng = currentExifLng;
      } else {
        // No EXIF GPS: keep existing pin location if we have one.
        if (prev.lat != null && prev.lng != null) {
          entry.lat = prev.lat;
          entry.lng = prev.lng;
        }
      }
    }
  }

  // If the photo has no GPS EXIF, fall back to the user's current geolocation
  if (entry.lat == null && entry.lng == null && lastUserLatitude != null && lastUserLongitude != null) {
    entry.lat = lastUserLatitude;
    entry.lng = lastUserLongitude;
  }

  // Keep the attached photo even if coordinates are missing.
  // (Coordinates may come from the device location or be added later.)

  const isPlaceholderBarangay = !barangay || /^select\s+barangay$/i.test(String(barangay).trim());
  if (!entry.business_name || isPlaceholderBarangay || !entry.date_inspected) {
    const missing = [
      !entry.business_name ? "Business name" : null,
      isPlaceholderBarangay ? "Barangay" : null,
      !entry.date_inspected ? "Date inspected" : null,
    ].filter(Boolean);
    logbookShowToast("inspection-toast", `⚠️ Missing: ${missing.join(", ")}`);

    // Bring the missing field into view on mobile so it's obvious.
    const firstMissingId = !entry.business_name
      ? "inspection_business_name"
      : isPlaceholderBarangay
        ? "inspection_addr_barangay"
        : "inspection_date_inspected";
    const el = document.getElementById(firstMissingId);
    if (el?.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "center" });
    if (el?.focus) {
      try { el.focus({ preventScroll: true }); } catch { el.focus(); }
    }
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
    if (inspectionFocusMapAfterSave && entry.lat != null && entry.lng != null) {
      inspectionFocusMapAfterSave = false;
      showView("map");
      window.location.hash = "map";
      closeNavSidebar();
      setTimeout(() => {
        if (mapInstance) {
          mapInstance.setView([entry.lat, entry.lng], 16);
          openInspectionDetailPanel(entry);
        }
      }, 100);
    }
    return;
  }

  (async () => {
    try {
      // If we have a file and Supabase Storage, upload to the 'storage' bucket
      if (currentExifFile && supabaseClient?.storage) {
        let uploadFile = currentExifFile;
        try {
          // Always pass through the sanitization module so EXIF/metadata
          // are stripped before the image can ever be embedded into PDFs.
          uploadFile = await sanitizeInspectionImage(uploadFile);
        } catch (sanitizeErr) {
          console.warn("Image sanitization failed, using original file:", sanitizeErr);
          uploadFile = currentExifFile;
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
        owner_phone: entry.insp_owner_phone || null,
        business_name: entry.business_name,
        address: entry.insp_address,
        date_inspected: entry.date_inspected,
        fsic_number: entry.fsic_number,
        inspected_by: entry.inspected_by || null,
        inspector_position: entry.inspector_position || null,
        included_personnel_name: entry.included_personnel_name || null,
        included_personnel_position: entry.included_personnel_position || null,
        duration_start: entry.duration_start || null,
        duration_end: entry.duration_end || null,
        remarks: null,
        latitude: entry.lat ?? null,
        longitude: entry.lng ?? null,
        photo_url: entry.photo_url ?? null,
        photo_taken_at: entry.photo_taken_at ?? null,
      };

      // On UPDATE: don't send photo_url when it's still a data URL (e.g. upload failed),
      // so we don't overwrite the existing DB photo with a huge string or null.
      if (inspectionEditingId && typeof payload.photo_url === "string" && payload.photo_url.startsWith("data:")) {
        delete payload.photo_url;
      }

      // On UPDATE, don't send empty/null optional fields — otherwise Supabase will
      // overwrite existing DB values to null when the user didn't intend to edit them.
      if (inspectionEditingId) {
        const requiredKeys = new Set([
          "io_number",
          "owner_name",
          "business_name",
          "address",
          "date_inspected",
          "fsic_number",
        ]);
        Object.keys(payload).forEach((k) => {
          if (requiredKeys.has(k)) return;
          const v = payload[k];
          if (v == null) delete payload[k];
          else if (typeof v === "string" && v.trim() === "") delete payload[k];
        });
      }
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
          const payloadNoGeo = { ...payload };
          delete payloadNoGeo.latitude;
          delete payloadNoGeo.longitude;
          const retry = await runWrite(payloadNoGeo);
          if (retry.error) throw retry.error;
        } else {
          throw error;
        }
      }
      await inspectionLoadFromSupabase();
      inspectionRenderTable();
      renderInspectionMarkersBatched();
      inspectionCloseModal();
      logbookShowToast("inspection-toast", "Saved to database.");

      const hasLocation = entry.lat != null && entry.lng != null;

      if (hasLocation) {
        // Always jump to the map to show the new pin
        showView("map");
        window.location.hash = "map";
        closeNavSidebar();
        setTimeout(() => {
          if (mapInstance) {
            mapInstance.setView([entry.lat, entry.lng], 16);
            openInspectionDetailPanel(entry);
          }
        }, 100);
      } else {
        // No location found: Force them to the logbook "No location" tab so they know it saved!
        // We highlight it without fully filtering so they see it in context.
        logbookShowToast("inspection-toast", "Saved! (No location found in photo)");
        showView("inspection");
        window.location.hash = "inspection";
        setInspectionTab("no-location");
        inspectionRenderTable();
        
        setTimeout(() => {
          // Find the newly saved row based on io_number or highest id
          const idx = inspectionData.findIndex((r) => r.io_number === entry.io_number);
          if (idx >= 0) {
            const rowEl = document.getElementById(`inspection-row-${idx}`);
            if (rowEl) {
              rowEl.classList.add("row-highlight");
              rowEl.scrollIntoView({ behavior: "smooth", block: "center" });
              setTimeout(() => rowEl.classList.remove("row-highlight"), 2500);
            }
          }
        }, 200);
      }
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

function inspectionClearFilters() {
  const q = document.getElementById("inspection-filter-q");
  const brgy = document.getElementById("inspection-filter-barangay");
  const personnel = document.getElementById("inspection-filter-personnel");
  const from = document.getElementById("inspection-filter-from");
  const to = document.getElementById("inspection-filter-to");
  if (q) q.value = "";
  if (brgy) brgy.value = "";
  if (personnel) personnel.value = "";
  if (from) from.value = "";
  if (to) to.value = "";
  inspectionRenderTable();
}

function openInspectionDetailPanel(entry) {
  const panel = document.getElementById("map-detail-panel");
  if (!panel) return;

  const titleEl = panel.querySelector?.(".map-detail-title");
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
  const viewBtn = document.getElementById("map-detail-view-logbook");
  const openIoBtn = document.getElementById("map-detail-open-io");

  if (titleEl) titleEl.textContent = "Inspection details";
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

  if (viewBtn) {
    viewBtn.textContent = "View in Inspection logbook";
    viewBtn.onclick = () => viewInspectionInLogbook(entry);
  }
  if (openIoBtn) {
    openIoBtn.style.display = "";
    openIoBtn.textContent = "Open IO (HTML)";
    openIoBtn.onclick = () => {
      try {
        sessionStorage.setItem("fsis.io.current", JSON.stringify(entry));
      } catch {}
      window.open("./inspection_io_fsis.html", "_blank");
    };
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

function openOccupancyDetailPanel(entry) {
  const panel = document.getElementById("map-detail-panel");
  if (!panel) return;

  const titleEl = panel.querySelector?.(".map-detail-title");
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
  const viewBtn = document.getElementById("map-detail-view-logbook");
  const openIoBtn = document.getElementById("map-detail-open-io");

  if (titleEl) titleEl.textContent = "Residential details";
  if (businessEl) businessEl.textContent = entry.owner_name || entry.io_number || "Residential";
  if (addressEl) {
    const parts = [
      entry.io_number ? `IO: ${entry.io_number}` : null,
      entry.remarks_signature ? String(entry.remarks_signature) : null,
    ].filter(Boolean);
    addressEl.textContent = parts.join(" · ") || "—";
  }
  if (coordsEl) {
    coordsEl.textContent =
      entry.lat != null && entry.lng != null ? `${entry.lat.toFixed(6)}, ${entry.lng.toFixed(6)}` : "—";
  }
  if (dateEl) dateEl.textContent = logbookFormatDate(entry.log_date);
  if (takenEl) {
    takenEl.textContent = entry.photo_taken_at ? String(entry.photo_taken_at) : "—";
    takenEl.classList.toggle("meta-muted", !entry.photo_taken_at);
  }
  if (inspectorEl) inspectorEl.textContent = entry.inspectors || "—";

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
          logbookShowToast("occupancy-toast", "Coordinates copied.");
        } catch {
          try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.left = "-9999px";
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
            logbookShowToast("occupancy-toast", "Coordinates copied.");
          } catch {
            logbookShowToast("occupancy-toast", "Copy failed.");
          }
        }
      };
    } else {
      copyCoordsBtn.style.display = "none";
      copyCoordsBtn.onclick = null;
    }
  }

  if (viewBtn) {
    viewBtn.textContent = "View in Occupancy logbook";
    viewBtn.onclick = () => viewOccupancyInLogbook(entry);
  }
  if (openIoBtn) {
    openIoBtn.style.display = "";
    openIoBtn.textContent = "Open IO (HTML)";
    openIoBtn.onclick = () => {
      try {
        sessionStorage.setItem("fsis.io.current", JSON.stringify(entry));
      } catch {}
      window.open("./occupancy_io_fsis.html", "_blank");
    };
  }

  if (photoWrap && photoImg) {
    if (entry.photo_url) {
      photoImg.src = entry.photo_url;
      photoImg.alt = entry.owner_name || "Occupancy photo";
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

function viewOccupancyInLogbook(entry) {
  showView("occupancy");
  if (getCurrentView() !== "occupancy") window.location.hash = "occupancy";

  const qEl = document.getElementById("occupancy-filter-q");
  if (qEl) qEl.value = (entry.io_number || entry.owner_name || "").trim();
  occupancyRenderTable?.();
  document.getElementById("table-occupancy")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function viewInspectionInLogbook(entry) {
  // Go to Inspection view and focus the matching entry.
  showView("inspection");
  if (getCurrentView() !== "inspection") window.location.hash = "inspection";

  // Prefill filters to narrow results to the selected entry.
  const qEl = document.getElementById("inspection-filter-q");
  if (qEl) qEl.value = (entry.io_number || entry.business_name || "").trim();

  const brgyEl = document.getElementById("inspection-filter-barangay");
  if (brgyEl) brgyEl.value = (entry.addr_barangay || "").trim();

  const pEl = document.getElementById("inspection-filter-personnel");
  if (pEl) pEl.value = (entry.inspected_by || "").trim();

  setInspectionTab(entry.lat != null && entry.lng != null ? "with-location" : "no-location");
  inspectionRenderTable();

  const idx = inspectionData.findIndex((r) => {
    if (entry.id && r.id) return r.id === entry.id;
    if (entry.io_number && r.io_number) return r.io_number === entry.io_number;
    return false;
  });
  if (idx < 0) return;

  const rowEl = document.getElementById(`inspection-row-${idx}`);
  if (!rowEl) return;
  rowEl.classList.add("row-highlight");
  rowEl.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => rowEl.classList.remove("row-highlight"), 2500);
}

function closeInspectionDetailPanel() {
  const panel = document.getElementById("map-detail-panel");
  if (!panel) return;
  panel.classList.remove("is-open");
  panel.setAttribute("aria-hidden", "true");
}

function initInspectionPhotoExif() {
  const inputCamera = document.getElementById("inspection_photo");
  const inputLibrary = document.getElementById("inspection_photo_library");
  if (!inputCamera && !inputLibrary) return;

  async function onPhotoChange(sourceInput) {
    const file = sourceInput?.files?.[0];
    if (!file) return;

    // Track async work so Save can wait briefly if needed.
    currentExifProcessingPromise = (async () => {
      // Reset extracted photo/exif state.
      currentExifLat = null;
      currentExifLng = null;
      currentExifPreviewUrl = null;
      currentExifTakenAt = null;
      currentExifFile = file;

      // Clear the other input so only one is active.
      if (sourceInput === inputCamera && inputLibrary) inputLibrary.value = "";
      if (sourceInput === inputLibrary && inputCamera) inputCamera.value = "";

      // Set preview URL (for immediate UI feedback); upload uses the File.
      try {
        const dataUrl = await new Promise((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => resolve(r.result);
          r.onerror = () => reject(new Error("FileReader failed"));
          r.readAsDataURL(file);
        });
        if (typeof dataUrl === "string") currentExifPreviewUrl = dataUrl;
      } catch (e) {
        console.warn("Preview read failed:", e);
      }

      // Read GPS from EXIF if present.
      try {
        const gps = await readGpsFromFile(file);
        if (gps && Number.isFinite(gps.lat) && Number.isFinite(gps.lng)) {
          currentExifLat = gps.lat;
          currentExifLng = gps.lng;
          logbookShowToast("inspection-toast", `GPS found: ${gps.lat.toFixed(5)}, ${gps.lng.toFixed(5)}`);
        } else {
          logbookShowToast("inspection-toast", "No GPS data found in this photo file. Your device may be stripping it.");
        }
      } catch (e) {
        console.warn("GPS read failed:", e);
        logbookShowToast("inspection-toast", "Error reading photo EXIF data.");
      }

      // If EXIF GPS is not available, fall back to the device's current location.
      if ((currentExifLat == null || currentExifLng == null) && navigator.geolocation) {
        try {
          await new Promise((resolve) => {
            navigator.geolocation.getCurrentPosition(
              (pos) => {
                const { latitude, longitude } = pos.coords || {};
                if (latitude != null && longitude != null) {
                  currentExifLat = latitude;
                  currentExifLng = longitude;
                }
                resolve();
              },
              () => resolve(),
              { enableHighAccuracy: true, timeout: 6000, maximumAge: 30_000 }
            );
          });
        } catch {
          // ignore
        }
      }
    })();

    // Fire and forget, but keep the promise for Save to await.
    void currentExifProcessingPromise;
  }

  inputCamera?.addEventListener("change", () => void onPhotoChange(inputCamera));
  inputLibrary?.addEventListener("change", () => void onPhotoChange(inputLibrary));
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

// Security/privacy: re‑encode an image via canvas so EXIF/metadata
// are stripped before the image is ever stored or used in PDFs.
async function sanitizeInspectionImage(file) {
  return new Promise((resolve) => {
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
            const longest = Math.max(img.width, img.height) || 1;
            const scale = longest > maxDimension ? maxDimension / longest : 1;

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
                if (!blob) {
                  resolve(file);
                  return;
                }
                const baseName =
                  (file.name && file.name.replace(/\.[^.]+$/, "")) || "photo";
                const sanitized = new File([blob], baseName + ".jpg", {
                  type: "image/jpeg",
                });
                resolve(sanitized);
              },
              "image/jpeg",
              0.8
            );
          } catch (err) {
            console.warn("sanitizeInspectionImage error:", err);
            resolve(file);
          }
        };
        img.onerror = () => resolve(file);
        img.src = e.target?.result;
      };
      reader.onerror = () => resolve(file);
      reader.readAsDataURL(file);
    } catch (err) {
      console.warn("sanitizeInspectionImage outer error:", err);
      resolve(file);
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

// Shared GPS extraction (used by both inspection and occupancy) to ensure consistent behavior.
async function readGpsFromFile(file) {
  const exifrApi = window.exifr?.default || window.exifr;
  
  const toFinite = (v) => {
    if (v == null) return null;
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    if (typeof v === "string") {
      const n = Number.parseFloat(v);
      return Number.isFinite(n) ? n : null;
    }
    return null;
  };

  const tryExifr = async (input) => {
    if (!exifrApi) return null;
    try {
      // Aggressive parse: check all common blocks where Android might stuff GPS data during gallery intent
      if (typeof exifrApi.parse === "function") {
        const parsed = await exifrApi.parse(input, { tiff: true, exif: true, gps: true });
        if (parsed) {
          const plat = toFinite(parsed.latitude ?? parsed.lat ?? parsed.GPSLatitude);
          const plng = toFinite(parsed.longitude ?? parsed.lng ?? parsed.GPSLongitude);
          if (plat != null && plng != null) return { lat: plat, lng: plng };
        }
      }
      
      // Fallback to the dedicated gps() method if parse() didn't find it or wasn't available
      if (typeof exifrApi.gps === "function") {
        const gps = await exifrApi.gps(input);
        const lat = toFinite(gps?.latitude ?? gps?.lat ?? gps?.GPSLatitude);
        const lng = toFinite(gps?.longitude ?? gps?.lng ?? gps?.GPSLongitude);
        if (lat != null && lng != null) return { lat, lng };
      }
    } catch (e) {
      console.warn("exifr read attempt failed:", e);
    }
    return null;
  };

  if (exifrApi) {
    // 1st attempt: ArrayBuffer (most reliable for raw binary)
    try {
      if (typeof file?.arrayBuffer === "function") {
        const buf = await file.arrayBuffer();
        const result = await tryExifr(buf);
        if (result) return result;
      }
    } catch (e) {
      console.warn("exifr arrayBuffer read failed:", e);
    }

    // 2nd attempt: Raw File object (sometimes needed depending on parser internals)
    try {
      const result = await tryExifr(file);
      if (result) return result;
    } catch (e) {
      console.warn("exifr file read failed:", e);
    }
  }

  // 3rd attempt: Fallback to exif-js if available
  if (window.EXIF) {
    let objectUrl = null;
    try {
      objectUrl = URL.createObjectURL(file);
      const img = new Image();
      
      await new Promise((resolve) => {
        img.onload = resolve;
        img.onerror = resolve;
        img.src = objectUrl;
      });

      return await new Promise((resolve) => {
        const done = (result) => {
          clearTimeout(timer);
          if (objectUrl) URL.revokeObjectURL(objectUrl);
          resolve(result);
        };
        const timer = setTimeout(() => done(null), 5000);
        try {
          window.EXIF.getData(img, function () {
            const lat = window.EXIF.getTag(this, "GPSLatitude");
            const latRef = window.EXIF.getTag(this, "GPSLatitudeRef");
            const lng = window.EXIF.getTag(this, "GPSLongitude");
            const lngRef = window.EXIF.getTag(this, "GPSLongitudeRef");
            if (lat && lng && latRef && lngRef) {
              done({ lat: dmsToDecimal(lat, latRef), lng: dmsToDecimal(lng, lngRef) });
            } else {
              done(null);
            }
          });
        } catch (e) {
          console.warn("exif-js sync block failed:", e);
          done(null);
        }
      });
    } catch (e) {
      console.warn("exif-js outer block failed:", e);
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    }
  }
  return null;
}

function initOccupancyPhotoExif() {
  const inputCamera = document.getElementById("occupancy_photo");
  const inputLibrary = document.getElementById("occupancy_photo_library");
  if (!inputCamera && !inputLibrary) return;

  async function onPhotoChange(sourceInput) {
    const file = sourceInput?.files?.[0];
    if (!file) return;

    occupancyExifLat = null;
    occupancyExifLng = null;
    occupancyExifPreviewUrl = null;
    occupancyExifTakenAt = null;
    occupancyExifFile = file;

    // Clear the other input so only one is active.
    if (sourceInput === inputCamera && inputLibrary) inputLibrary.value = "";
    if (sourceInput === inputLibrary && inputCamera) inputCamera.value = "";

    try {
      const dataUrl = await new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result);
        r.onerror = () => reject(new Error("FileReader failed"));
        r.readAsDataURL(file);
      });
      if (typeof dataUrl === "string") occupancyExifPreviewUrl = dataUrl;
    } catch (e) {
      console.warn("Occupancy preview read failed:", e);
    }

    try {
      const gps = await readGpsFromFile(file);
      if (gps) {
        occupancyExifLat = gps.lat;
        occupancyExifLng = gps.lng;
        logbookShowToast("occupancy-toast", `GPS found: ${gps.lat.toFixed(5)}, ${gps.lng.toFixed(5)}`);
      } else {
        logbookShowToast("occupancy-toast", "No GPS data found in this photo file. Your device may be stripping it.");
      }
    } catch (e) {
      console.warn("Occupancy GPS read failed:", e);
      logbookShowToast("occupancy-toast", "Error reading photo EXIF data.");
    }
  }

  inputCamera?.addEventListener("change", () => void onPhotoChange(inputCamera));
  inputLibrary?.addEventListener("change", () => void onPhotoChange(inputLibrary));
}

function addInspectionMarkerFromEntry(entry) {
  if (!mapInstance || !inspectionMarkersLayer) return;
  if (!Number.isFinite(entry.lat) || !Number.isFinite(entry.lng)) return;

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

  const tooltipText = entry.business_name || entry.insp_owner || entry.io_number || "Inspection";
  marker.bindTooltip(tooltipText, {
    permanent: false,
    direction: "top",
    offset: [0, -36],
    opacity: 0.95,
    className: "inspection-marker-tooltip",
  });

  marker.on("mouseover", () => {
    marker.setZIndexOffset(1000);
    const el = marker.getElement?.();
    if (el) el.classList.add("is-hover");
  });
  marker.on("mouseout", () => {
    marker.setZIndexOffset(0);
    const el = marker.getElement?.();
    if (el) el.classList.remove("is-hover");
  });
  marker.on("click", () => {
    openInspectionDetailPanel(entry);
  });
}

function addOccupancyMarkerFromEntry(entry) {
  if (!mapInstance || !occupancyMarkersLayer) return;
  if (!Number.isFinite(entry.lat) || !Number.isFinite(entry.lng)) return;

  let icon;
  if (entry.photo_url) {
    icon = L.divIcon({
      className: "occupancy-marker occupancy-marker--photo",
      html: `<div class="occupancy-marker-thumb" style="background-image:url('${entry.photo_url}')"></div>`,
      iconSize: [40, 40],
      iconAnchor: [20, 40],
      popupAnchor: [0, -36],
    });
  } else {
    icon = L.divIcon({
      className: "occupancy-marker occupancy-marker--default",
      html: '<div class="occupancy-marker-pin"></div>',
      iconSize: [28, 36],
      iconAnchor: [14, 36],
      popupAnchor: [0, -30],
    });
  }

  const title = entry.owner_name || entry.io_number || "Occupancy";
  const marker = L.marker([entry.lat, entry.lng], { icon }).addTo(occupancyMarkersLayer);
  marker.bindPopup(`<strong>${logbookEsc(title)}</strong><br>${logbookEsc(entry.io_number || "")}`);
  marker.on("click", () => {
    openOccupancyDetailPanel(entry);
  });
}

function renderOccupancyMarkersBatched() {
  if (!mapInstance || !occupancyMarkersLayer || !Array.isArray(occupancyData)) return;
  occupancyMarkersLayer.clearLayers();
  const withCoords = occupancyData.filter((row) => row.lat != null && row.lng != null);
  let i = 0;
  const batchSize = 40;
  function addBatch() {
    const end = Math.min(i + batchSize, withCoords.length);
    for (; i < end; i++) addOccupancyMarkerFromEntry(withCoords[i]);
    if (i < withCoords.length) requestAnimationFrame(addBatch);
  }
  addBatch();
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

function getMarkedInspectionEntries() {
  if (!Array.isArray(inspectionData)) return [];
  return inspectionData.filter((row) => row.lat != null && row.lng != null);
}

function getMarkedOccupancyEntries() {
  if (!Array.isArray(occupancyData)) return [];
  return occupancyData.filter((row) => row.lat != null && row.lng != null);
}

function searchMapLocations(query) {
  const q = normalizeQuery(query);
  const includeInspection = mapMarkerFilter === "all" || mapMarkerFilter === "businesses";
  const includeOccupancy = mapMarkerFilter === "all" || mapMarkerFilter === "residential";

  const candidates = [
    ...(includeInspection ? getMarkedInspectionEntries().map((r) => ({ type: "inspection", r })) : []),
    ...(includeOccupancy ? getMarkedOccupancyEntries().map((r) => ({ type: "occupancy", r })) : []),
  ];

  const filtered = !q
    ? candidates
    : candidates.filter(({ type, r }) => {
        const hay =
          type === "inspection"
            ? normalizeQuery(
                [
                  r.io_number,
                  r.insp_owner,
                  r.insp_owner_phone,
                  r.business_name,
                  inspectionFormatAddressDisplay(r),
                  r.fsic_number,
                  r.inspected_by,
                ].join(" | ")
              )
            : normalizeQuery(
                [
                  r.io_number,
                  r.owner_name,
                  r.inspectors,
                  r.remarks_signature,
                  r.lat,
                  r.lng,
                ].join(" | ")
              );
        return hay.includes(q);
      });

  // Convert back to a unified shape used by the UI list.
  return filtered
    .slice(0, 20)
    .map(({ type, r }) =>
      type === "inspection"
        ? { ...r, _mapType: "inspection" }
        : {
            ...r,
            insp_owner: r.owner_name,
            business_name: "Residential",
            fsic_number: "",
            inspected_by: r.inspectors,
            _mapType: "occupancy",
          }
    );
}

function initMapSearch() {
  const input = document.getElementById("map-search-input");
  const resultsEl = document.getElementById("map-search-results");
  if (!input || !resultsEl) return;

  let debounceTimer = null;
  function runSearch() {
    const query = (input.value || "").trim();
    const entries = searchMapLocations(query);
    resultsEl.hidden = false;
    resultsEl.innerHTML = "";
    if (entries.length === 0) {
      const empty = document.createElement("div");
      empty.className = "map-search-result-empty";
      empty.textContent = query ? "No marked locations match." : "Type to search business, IO number, address, owner…";
      resultsEl.appendChild(empty);
      return;
    }
    entries.forEach((entry) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "map-search-result-item";
      btn.setAttribute("role", "option");
      const title = entry.business_name || entry.insp_owner || entry.io_number || "Inspection";
      const sub =
        entry._mapType === "occupancy"
          ? [entry.io_number, entry.insp_owner].filter(Boolean).join(" · ") || "—"
          : [entry.io_number, inspectionFormatAddressDisplay(entry)].filter(Boolean).join(" · ") || "—";
      const strong = document.createElement("strong");
      strong.textContent = title;
      const span = document.createElement("span");
      span.textContent = sub;
      btn.appendChild(strong);
      btn.appendChild(span);
      btn.onclick = () => {
        if (mapInstance && entry.lat != null && entry.lng != null) {
          mapInstance.setView([entry.lat, entry.lng], 16);
          if (entry._mapType === "occupancy") openOccupancyDetailPanel(entry);
          else openInspectionDetailPanel(entry);
        }
        input.value = "";
        resultsEl.hidden = true;
        resultsEl.innerHTML = "";
        input.focus();
      };
      resultsEl.appendChild(btn);
    });
  }

  input.addEventListener("input", () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(runSearch, 180);
  });
  input.addEventListener("focus", () => {
    runSearch();
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      resultsEl.hidden = true;
      resultsEl.innerHTML = "";
      input.blur();
    }
  });

  document.addEventListener("click", (e) => {
    if (resultsEl.hidden) return;
    if (!input.contains(e.target) && !resultsEl.contains(e.target)) {
      resultsEl.hidden = true;
      resultsEl.innerHTML = "";
    }
  });
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
  const baseSelect =
    "id, io_number, owner_name, owner_phone, business_name, address, date_inspected, fsic_number, inspected_by, inspector_position, included_personnel_name, included_personnel_position, duration_start, duration_end, remarks, fsic_purpose, fsic_permit_type, fsic_valid_for, fsic_valid_until, fsic_fee_amount, fsic_fee_or_number, fsic_fee_date, photo_url, photo_taken_at, created_at";
  // Geo columns vary across deployments:
  // - New schema: latitude/longitude
  // - Legacy schema: lat/lng
  const selectWithGeo = `${baseSelect}, latitude, longitude`;
  const selectWithGeoLegacy = `${baseSelect}, lat, lng`;
  const selectWithoutGeo = baseSelect;

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
      const missingNewGeo = msg.includes("latitude") || msg.includes("longitude");
      if (!missingNewGeo) throw error;

      // Backward compatibility: try legacy lat/lng columns
      const legacy = await run(selectWithGeoLegacy);
      if (!legacy.error) rows = legacy.data;
      else {
        // Last fallback: no geo columns at all
        const retry = await run(selectWithoutGeo);
        if (retry.error) throw retry.error;
        rows = retry.data;
      }
    }
  }

  inspectionData = (rows || []).map((r) => ({
    id: r.id,
    io_number: r.io_number,
    fsic_number: r.fsic_number,
    insp_owner: r.owner_name,
    insp_owner_phone: r.owner_phone ?? "",
    business_name: r.business_name,
    insp_address: r.address,
    date_inspected: r.date_inspected,
    inspected_by: r.inspected_by || "",
    inspector_position: r.inspector_position || "",
    included_personnel_name: r.included_personnel_name || "",
    included_personnel_position: r.included_personnel_position || "",
    duration_start: r.duration_start || null,
    duration_end: r.duration_end || null,
    remarks: r.remarks || "",
    // FSIC clearance / fees
    fsic_purpose: r.fsic_purpose ?? null,
    fsic_permit_type: r.fsic_permit_type ?? null,
    fsic_valid_for: r.fsic_valid_for ?? null,
    fsic_valid_until: r.fsic_valid_until ?? null,
    fsic_fee_amount: r.fsic_fee_amount ?? null,
    fsic_fee_or_number: r.fsic_fee_or_number ?? null,
    fsic_fee_date: r.fsic_fee_date ?? null,
    lat: r.latitude ?? r.lat ?? null,
    lng: r.longitude ?? r.lng ?? null,
    photo_url: r.photo_url ?? null,
    photo_taken_at: r.photo_taken_at ?? null,
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

  const q = normalizeQuery(document.getElementById("fsec-filter-q")?.value);
  const from = (document.getElementById("fsec-filter-from")?.value || "").trim();
  const to = (document.getElementById("fsec-filter-to")?.value || "").trim();

  const filtered = fsecData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (from || to) {
        if (!inDateRange(row.fsec_date, from, to)) return false;
      }
      if (!q) return true;
      const hay = normalizeQuery(
        [
          row.fsec_owner,
          row.proposed_project,
          fsecFormatAddressDisplay(row),
          row.contact_number,
        ].join(" | ")
      );
      return hay.includes(q);
    });

  tbody.innerHTML = "";
  if (fsecData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  if (filtered.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  empty.style.display = "none";
  if (tableWrap) tableWrap.style.display = "";
  filtered.forEach(({ row, idx }, displayIdx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td data-label="#">${displayIdx + 1}</td>
      <td data-label="Name of Owner">${logbookEsc(row.fsec_owner)}</td>
      <td data-label="Proposed Project"><strong>${logbookEsc(row.proposed_project)}</strong></td>
      <td data-label="Address">${logbookEsc(fsecFormatAddressDisplay(row))}</td>
      <td class="td-date" data-label="Date">${logbookFormatDate(row.fsec_date)}</td>
      <td data-label="Contact Number">${logbookEsc(row.contact_number)}</td>
      <td class="col-action" data-label="Action">
        <select class="action-select"
          aria-label="Row actions"
          onchange="fsecHandleAction(this.value, ${idx}); this.selectedIndex = 0;"
        >
          <option value="">Actions…</option>
          <option value="edit">Edit</option>
          <option value="delete">Delete</option>
        </select>
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
  const barangayVal = row.addr_barangay || (brgyMatch ? brgyMatch[1].trim() : "");
  ensureSelectOption("fsec_addr_barangay", barangayVal);
  setVal("fsec_addr_line", row.addr_line || "");

  const overlay = document.getElementById("fsec-modal-overlay");
  if (overlay) overlay.classList.add("open");
  setText("fsec-modal-title", "Edit FSEC Building Plan Record");
  setText("fsec-modal-subtitle", "FSEC Building Plan Logbook");
  const btn = document.getElementById("fsec-btn-save");
  if (btn) btn.textContent = "Update Record";
}

function fsecHandleAction(action, idx) {
  if (!action) return;
  if (action === "edit") {
    fsecEditEntry(idx);
    return;
  }
  if (action === "delete") {
    fsecDeleteEntry(idx);
  }
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

  const mergedAddress = [
    `Region ${region}`,
    province,
    municipal,
    `Barangay ${barangay}`,
    line,
  ]
    .filter((p) => String(p || "").trim())
    .join(", ");

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

function fsecClearFilters() {
  const q = document.getElementById("fsec-filter-q");
  const from = document.getElementById("fsec-filter-from");
  const to = document.getElementById("fsec-filter-to");
  if (q) q.value = "";
  if (from) from.value = "";
  if (to) to.value = "";
  fsecRenderTable();
}

function conveyanceSetPrintDate() {
  const el = document.getElementById("conveyance-print-date");
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleDateString("en-PH", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function conveyancePrintPanel() {
  conveyanceSetPrintDate();
  const oldTitle = document.title;
  document.title = "";
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.title = oldTitle;
    }, 500);
  }, 0);
}

function conveyanceClearFilters() {
  const q = document.getElementById("conveyance-filter-q");
  const from = document.getElementById("conveyance-filter-from");
  const to = document.getElementById("conveyance-filter-to");
  if (q) q.value = "";
  if (from) from.value = "";
  if (to) to.value = "";
  conveyanceRenderTable();
}

function occupancySetPrintDate() {
  const el = document.getElementById("occupancy-print-date");
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleDateString("en-PH", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function occupancyPrintPanel() {
  occupancySetPrintDate();
  const oldTitle = document.title;
  document.title = "";
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.title = oldTitle;
    }, 500);
  }, 0);
}

function occupancyClearFilters() {
  const q = document.getElementById("occupancy-filter-q");
  const from = document.getElementById("occupancy-filter-from");
  const to = document.getElementById("occupancy-filter-to");
  if (q) q.value = "";
  if (from) from.value = "";
  if (to) to.value = "";
  occupancyRenderTable();
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

// -----------------------------
// Conveyance logbook module
// -----------------------------

const CONVEYANCE_STORAGE_KEY = "bfp_conveyance";
let conveyanceData = [];
let conveyanceEditingIdx = null;
let conveyanceEditingId = null;
let conveyanceDataLoaded = false;

function conveyanceLoadFromLocal() {
  conveyanceData = JSON.parse(localStorage.getItem(CONVEYANCE_STORAGE_KEY) || "[]");
}

function conveyanceSaveToLocal() {
  localStorage.setItem(CONVEYANCE_STORAGE_KEY, JSON.stringify(conveyanceData));
}

function conveyanceRenderTable() {
  const tbody = document.getElementById("tbody-conveyance");
  const empty = document.getElementById("empty-conveyance");
  const tableWrap = document.getElementById("table-conveyance")?.closest(".table-wrap");
  if (!tbody || !empty) return;

  const q = normalizeQuery(document.getElementById("conveyance-filter-q")?.value);
  const from = (document.getElementById("conveyance-filter-from")?.value || "").trim();
  const to = (document.getElementById("conveyance-filter-to")?.value || "").trim();

  const filtered = conveyanceData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (from || to) {
        if (!inDateRange(row.log_date, from, to)) return false;
      }
      if (!q) return true;
      const hay = normalizeQuery(
        [row.io_number, row.owner_name, row.inspectors, row.remarks_signature].join(" | ")
      );
      return hay.includes(q);
    });

  tbody.innerHTML = "";
  if (conveyanceData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  if (filtered.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  empty.style.display = "none";
  if (tableWrap) tableWrap.style.display = "";

  filtered.forEach(({ row, idx }, displayIdx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td data-label="#">${displayIdx + 1}</td>
      <td class="td-date" data-label="Date">${logbookFormatDate(row.log_date)}</td>
      <td data-label="IO Number">${logbookEsc(row.io_number)}</td>
      <td data-label="Name of Owner">${logbookEsc(row.owner_name)}</td>
      <td data-label="Name of Inspectors"><div class="cell-pre">${logbookEsc(row.inspectors)}</div></td>
      <td data-label="Remarks / Signature"><div class="cell-pre">${logbookEsc(row.remarks_signature)}</div></td>
      <td class="col-action" data-label="Action">
        <select class="action-select" aria-label="Row actions" onchange="conveyanceHandleAction(this.value, ${idx}); this.selectedIndex = 0;">
          <option value="">Actions…</option>
          <option value="edit">Edit</option>
          <option value="delete">Delete</option>
        </select>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function conveyanceHandleAction(action, idx) {
  if (!action) return;
  if (action === "edit") return conveyanceEditEntry(idx);
  if (action === "delete") return conveyanceDeleteEntry(idx);
}

function conveyanceEditEntry(idx) {
  const row = conveyanceData[idx];
  if (!row) return;

  conveyanceEditingIdx = idx;
  conveyanceEditingId = row.id || null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };
  setVal("conveyance_date", row.log_date);
  setVal("conveyance_io_number", row.io_number);
  setVal("conveyance_owner_name", row.owner_name);
  setVal("conveyance_inspectors", row.inspectors);
  setVal("conveyance_remarks_signature", row.remarks_signature);

  setText("conveyance-modal-title", "Edit Conveyance Record");
  const btn = document.getElementById("conveyance-btn-save");
  if (btn) btn.textContent = "Update Record";

  const overlay = document.getElementById("conveyance-modal-overlay");
  overlay?.classList.add("open");
}

function conveyanceOpenModal() {
  conveyanceEditingIdx = null;
  conveyanceEditingId = null;

  setText("conveyance-modal-title", "Add Conveyance Record");
  const btn = document.getElementById("conveyance-btn-save");
  if (btn) btn.textContent = "Save Record";

  const date = document.getElementById("conveyance_date");
  if (date) date.value = new Date().toISOString().slice(0, 10);
  const io = document.getElementById("conveyance_io_number");
  if (io) io.value = "";
  const owner = document.getElementById("conveyance_owner_name");
  if (owner) owner.value = "";
  const insp = document.getElementById("conveyance_inspectors");
  if (insp) insp.value = "";
  const rem = document.getElementById("conveyance_remarks_signature");
  if (rem) rem.value = "";

  const overlay = document.getElementById("conveyance-modal-overlay");
  overlay?.classList.add("open");
}

function conveyanceCloseModal() {
  const overlay = document.getElementById("conveyance-modal-overlay");
  overlay?.classList.remove("open");
}

function conveyanceCloseOnOverlay(e) {
  if (e?.target?.id === "conveyance-modal-overlay") conveyanceCloseModal();
}

function conveyanceDeleteEntry(idx) {
  if (!confirm("Delete this record?")) return;

  const row = conveyanceData[idx];
  if (!row) return;

  // Optimistic local delete
  conveyanceData.splice(idx, 1);
  conveyanceSaveToLocal();
  conveyanceRenderTable();
  logbookShowToast("conveyance-toast", "Record deleted.");

  if (!isSupabaseEnabled()) return;
  if (!row.id) {
    logbookShowToast("conveyance-toast", "⚠️ Cannot delete: missing record id.");
    return;
  }

  (async () => {
    try {
      const { error } = await supabaseClient
        .from("conveyance_logbook")
        .delete()
        .eq("id", row.id);
      if (error) throw error;
    } catch (err) {
      logbookShowToast("conveyance-toast", "⚠️ Delete failed: " + (err?.message || err));
    }
  })();
}

function conveyanceSaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();

  const entry = {
    log_date: (document.getElementById("conveyance_date") || { value: "" }).value,
    io_number: (document.getElementById("conveyance_io_number") || { value: "" }).value.trim(),
    owner_name: (document.getElementById("conveyance_owner_name") || { value: "" }).value.trim(),
    inspectors: (document.getElementById("conveyance_inspectors") || { value: "" }).value.trim(),
    remarks_signature: (document.getElementById("conveyance_remarks_signature") || { value: "" }).value.trim(),
    created_at: new Date().toISOString(),
  };

  if (!entry.log_date || !entry.io_number || !entry.inspectors) {
    logbookShowToast("conveyance-toast", "⚠️ Please fill in Date, IO Number, and Inspectors.");
    return;
  }

  const isOnline = isSupabaseEnabled();

  if (conveyanceEditingIdx !== null) {
    const prev = conveyanceData[conveyanceEditingIdx] || {};
    conveyanceData[conveyanceEditingIdx] = {
      ...prev,
      ...entry,
      id: prev.id || null,
      created_at: prev.created_at || entry.created_at,
    };
  } else {
    conveyanceData.push({ ...entry, id: null });
  }

  conveyanceSaveToLocal();
  conveyanceRenderTable();
  conveyanceCloseModal();
  showSaveIndicator("Conveyance record saved");

  if (!isOnline) {
    logbookShowToast("conveyance-toast", "Saved on this device only (offline mode).");
    return;
  }

  (async () => {
    try {
      const payload = {
        log_date: entry.log_date,
        io_number: entry.io_number,
        owner_name: entry.owner_name || null,
        inspectors: entry.inspectors,
        remarks_signature: entry.remarks_signature,
      };
      // On UPDATE, don't overwrite optional owner_name with null/blank.
      if (conveyanceEditingId && (!payload.owner_name || String(payload.owner_name).trim() === "")) {
        delete payload.owner_name;
      }
      const q = supabaseClient.from("conveyance_logbook");
      const { error } = conveyanceEditingId
        ? await q.update(payload).eq("id", conveyanceEditingId)
        : await q.insert(payload);
      if (error) throw error;
      logbookShowToast("conveyance-toast", "Saved to database.");
    } catch (err) {
      const msg = err?.message || String(err);
      const hint =
        msg.includes("policy") || msg.includes("RLS") || err?.code === "42501"
          ? " Check Supabase: add anon RLS policy (see fsis.logger.sql)."
          : "";
      logbookShowToast("conveyance-toast", "Save failed: " + msg + hint);
    }
  })();
}

async function conveyanceLoadFromSupabase() {
  const { data: rows, error } = await supabaseClient
    .from("conveyance_logbook")
    .select("id, log_date, io_number, owner_name, inspectors, remarks_signature, created_at")
    .order("created_at", { ascending: true })
    .limit(2000);
  if (error) throw error;
  conveyanceData = (rows || []).map((r) => ({
    id: r.id,
    log_date: r.log_date,
    io_number: r.io_number,
    owner_name: r.owner_name || "",
    inspectors: r.inspectors,
    remarks_signature: r.remarks_signature,
    created_at: r.created_at,
  }));
  conveyanceSaveToLocal();
}

async function conveyanceInitData() {
  if (conveyanceDataLoaded) return;
  conveyanceDataLoaded = true;
  try {
    conveyanceLoadFromLocal();
    conveyanceRenderTable();
    if (isSupabaseEnabled()) {
      await conveyanceLoadFromSupabase();
      conveyanceRenderTable();
    }
  } catch (err) {
    conveyanceLoadFromLocal();
    console.warn("Conveyance load failed, using local storage:", err);
    logbookShowToast("conveyance-toast", "Using data stored on this device.");
    conveyanceRenderTable();
  }
}

// -----------------------------
// Occupancy logbook module
// -----------------------------

const OCCUPANCY_STORAGE_KEY = "bfp_occupancy";
let occupancyData = [];
let occupancyEditingIdx = null;
let occupancyEditingId = null;
let occupancyDataLoaded = false;

function occupancyLoadFromLocal() {
  occupancyData = JSON.parse(localStorage.getItem(OCCUPANCY_STORAGE_KEY) || "[]");
}

function occupancySaveToLocal() {
  // Avoid storing huge inline image data URLs in localStorage (they quickly exceed quota).
  const safe = occupancyData.map((row) => {
    const copy = { ...row };
    if (typeof copy.photo_url === "string" && copy.photo_url.startsWith("data:")) {
      copy.photo_url = null;
    }
    return copy;
  });

  try {
    localStorage.setItem(OCCUPANCY_STORAGE_KEY, JSON.stringify(safe));
  } catch (err) {
    console.warn("Failed to persist occupancy cache to localStorage:", err);
  }
}

function occupancyRenderTable() {
  const tbody = document.getElementById("tbody-occupancy");
  const empty = document.getElementById("empty-occupancy");
  const tableWrap = document.getElementById("table-occupancy")?.closest(".table-wrap");
  if (!tbody || !empty) return;

  const q = normalizeQuery(document.getElementById("occupancy-filter-q")?.value);
  const from = (document.getElementById("occupancy-filter-from")?.value || "").trim();
  const to = (document.getElementById("occupancy-filter-to")?.value || "").trim();

  const filtered = occupancyData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (from || to) {
        if (!inDateRange(row.log_date, from, to)) return false;
      }
      if (!q) return true;
      const hay = normalizeQuery(
        [
          row.io_number,
          row.owner_name,
          row.inspectors,
          row.remarks_signature,
          row.lat,
          row.lng,
        ].join(" | ")
      );
      return hay.includes(q);
    });

  tbody.innerHTML = "";
  if (occupancyData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  if (filtered.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
    return;
  }

  empty.style.display = "none";
  if (tableWrap) tableWrap.style.display = "";

  filtered.forEach(({ row, idx }, displayIdx) => {
    const tr = document.createElement("tr");
    tr.id = `occupancy-row-${idx}`;
    tr.innerHTML = `
      <td data-label="#">${displayIdx + 1}</td>
      <td class="td-date" data-label="Date">${logbookFormatDate(row.log_date)}</td>
      <td data-label="IO Number">${logbookEsc(row.io_number)}</td>
      <td data-label="Name of Owner">${logbookEsc(row.owner_name)}</td>
      <td data-label="Name of Inspectors"><div class="cell-pre">${logbookEsc(row.inspectors)}</div></td>
      <td data-label="Remarks / Signature"><div class="cell-pre">${logbookEsc(row.remarks_signature)}</div></td>
      <td class="col-action" data-label="Action">
        <select class="action-select" aria-label="Row actions" onchange="occupancyHandleAction(this.value, ${idx}); this.selectedIndex = 0;">
          <option value="">Actions…</option>
          <option value="edit">Edit</option>
          <option value="open_io_html">Open IO (HTML)</option>
          <option value="delete">Delete</option>
        </select>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function occupancyHandleAction(action, idx) {
  if (!action) return;
  if (action === "edit") return occupancyEditEntry(idx);
  if (action === "open_io_html") return occupancyOpenIoHtml(idx);
  if (action === "delete") return occupancyDeleteEntry(idx);
}

function occupancyOpenIoHtml(idx) {
  const row = occupancyData[idx];
  if (!row) return;
  try {
    sessionStorage.setItem("fsis.io.current", JSON.stringify(row));
  } catch {
    // If sessionStorage is unavailable, we still open the template;
    // it will show a friendly notice instead of data.
  }
  window.open("./occupancy_io_fsis.html", "_blank");
}

function occupancyEditEntry(idx) {
  const row = occupancyData[idx];
  if (!row) return;

  occupancyEditingIdx = idx;
  occupancyEditingId = row.id || null;

  // Preserve existing location/photo for edits so we don't accidentally
  // reuse EXIF state from a different modal action.
  occupancyExifLat = row.lat ?? null;
  occupancyExifLng = row.lng ?? null;
  occupancyExifPreviewUrl = row.photo_url ?? null;
  occupancyExifTakenAt = row.photo_taken_at ?? null;
  occupancyExifFile = null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };
  setVal("occupancy_date", row.log_date);
  setVal("occupancy_io_number", row.io_number);
  setVal("occupancy_owner_name", row.owner_name);
  setVal("occupancy_inspected_by", row.inspectors);

  setText("occupancy-modal-title", "Edit Occupancy Record");
  const btn = document.getElementById("occupancy-btn-save");
  if (btn) btn.textContent = "Update Record";

  const overlay = document.getElementById("occupancy-modal-overlay");
  overlay?.classList.add("open");
}

function occupancyOpenModal() {
  occupancyEditingIdx = null;
  occupancyEditingId = null;

  // Reset any previously extracted EXIF coordinates and photo data
  occupancyExifLat = null;
  occupancyExifLng = null;
  occupancyExifPreviewUrl = null;
  occupancyExifTakenAt = null;
  occupancyExifFile = null;

  setText("occupancy-modal-title", "Add Occupancy Record");
  const btn = document.getElementById("occupancy-btn-save");
  if (btn) btn.textContent = "Save Record";

  const date = document.getElementById("occupancy_date");
  if (date) date.value = new Date().toISOString().slice(0, 10);
  const io = document.getElementById("occupancy_io_number");
  if (io) io.value = "";
  const owner = document.getElementById("occupancy_owner_name");
  if (owner) owner.value = "";
  const insp = document.getElementById("occupancy_inspected_by");
  if (insp) insp.selectedIndex = 0;

  const photoInput = document.getElementById("occupancy_photo");
  if (photoInput) photoInput.value = "";
  const photoLibraryInput = document.getElementById("occupancy_photo_library");
  if (photoLibraryInput) photoLibraryInput.value = "";

  const overlay = document.getElementById("occupancy-modal-overlay");
  overlay?.classList.add("open");
}

function occupancyCloseModal() {
  const overlay = document.getElementById("occupancy-modal-overlay");
  overlay?.classList.remove("open");
}

function occupancyCloseOnOverlay(e) {
  if (e?.target?.id === "occupancy-modal-overlay") occupancyCloseModal();
}

function occupancyDeleteEntry(idx) {
  if (!confirm("Delete this record?")) return;

  const row = occupancyData[idx];
  if (!row) return;

  // Optimistic local delete
  occupancyData.splice(idx, 1);
  occupancySaveToLocal();
  occupancyRenderTable();
  logbookShowToast("occupancy-toast", "Record deleted.");

  if (!isSupabaseEnabled()) return;
  if (!row.id) {
    logbookShowToast("occupancy-toast", "⚠️ Cannot delete: missing record id.");
    return;
  }

  (async () => {
    try {
      const { error } = await supabaseClient
        .from("occupancy_logbook")
        .delete()
        .eq("id", row.id);
      if (error) throw error;
    } catch (err) {
      logbookShowToast("occupancy-toast", "⚠️ Delete failed: " + (err?.message || err));
    }
  })();
}

function occupancySaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();

  const entry = {
    log_date: (document.getElementById("occupancy_date") || { value: "" }).value,
    io_number: (document.getElementById("occupancy_io_number") || { value: "" }).value.trim(),
    owner_name: (document.getElementById("occupancy_owner_name") || { value: "" }).value.trim(),
    inspectors: (document.getElementById("occupancy_inspected_by") || { value: "" }).value.trim(),
    remarks_signature: "",
    // Optional coordinates and photo metadata extracted from EXIF / geolocation
    lat: occupancyExifLat,
    lng: occupancyExifLng,
    photo_url: occupancyExifPreviewUrl,
    photo_taken_at: occupancyExifTakenAt,
    created_at: new Date().toISOString(),
  };

  // If the photo has no GPS EXIF, fall back to the user's current geolocation (same as inspection)
  if (entry.lat == null && entry.lng == null && lastUserLatitude != null && lastUserLongitude != null) {
    entry.lat = lastUserLatitude;
    entry.lng = lastUserLongitude;
  }

  // If we still have no coordinates, clear coordinates but keep the photo
  // (the photo may still be useful even without GPS location data)
  if (entry.lat == null || entry.lng == null) {
    entry.lat = null;
    entry.lng = null;
  }

  if (!entry.log_date || !entry.io_number || !entry.inspectors) {
    logbookShowToast("occupancy-toast", "⚠️ Please fill in Date, IO Number, and Inspectors.");
    return;
  }

  const isOnline = isSupabaseEnabled();

  if (occupancyEditingIdx !== null) {
    const prev = occupancyData[occupancyEditingIdx] || {};
    occupancyData[occupancyEditingIdx] = {
      ...prev,
      ...entry,
      id: prev.id || null,
      created_at: prev.created_at || entry.created_at,
    };
  } else {
    occupancyData.push({ ...entry, id: null });
  }

  occupancySaveToLocal();
  occupancyRenderTable();
  occupancyCloseModal();
  showSaveIndicator("Occupancy record saved");
  addOccupancyMarkerFromEntry(entry);

  if (!isOnline) {
    logbookShowToast("occupancy-toast", "Saved on this device only (offline mode).");
    return;
  }

  (async () => {
    try {
      // If we have a photo file and Supabase Storage, upload and set entry.photo_url to public URL
      if (occupancyExifFile && supabaseClient?.storage) {
        let uploadFile = occupancyExifFile;
        try {
          uploadFile = await sanitizeInspectionImage(uploadFile);
        } catch (sanitizeErr) {
          console.warn("Occupancy image sanitization failed, using original:", sanitizeErr);
          uploadFile = occupancyExifFile;
        }
        const fileExt =
          (occupancyExifFile.name && occupancyExifFile.name.split(".").pop()) || "jpg";
        const safeExt = fileExt.toLowerCase().replace(/[^a-z0-9]/g, "") || "jpg";
        const path = `occupancy-photos/${Date.now()}-${Math.random().toString(36).slice(2)}.${safeExt}`;
        try {
          const { error: uploadError } = await supabaseClient.storage
            .from("storage")
            .upload(path, uploadFile, {
              cacheControl: "3600",
              upsert: false,
              contentType: uploadFile.type || occupancyExifFile.type || "image/jpeg",
            });
          if (!uploadError) {
            const { data: urlData } = supabaseClient.storage.from("storage").getPublicUrl(path);
            if (urlData?.publicUrl) entry.photo_url = urlData.publicUrl;
          } else {
            console.warn("Occupancy photo upload failed:", uploadError);
          }
        } catch (uploadErr) {
          console.warn("Occupancy photo upload threw:", uploadErr);
        }
      }

      const payload = {
        log_date: entry.log_date,
        io_number: entry.io_number,
        owner_name: entry.owner_name || null,
        inspectors: entry.inspectors,
        remarks_signature: null,
        latitude: entry.lat ?? null,
        longitude: entry.lng ?? null,
        photo_url: entry.photo_url ?? null,
        photo_taken_at: entry.photo_taken_at ?? null,
      };
      // On UPDATE, don't overwrite optional owner_name with null/blank.
      if (occupancyEditingId && (!payload.owner_name || String(payload.owner_name).trim() === "")) {
        delete payload.owner_name;
      }
      const q = supabaseClient.from("occupancy_logbook");
      const { error } = occupancyEditingId
        ? await q.update(payload).eq("id", occupancyEditingId)
        : await q.insert(payload);
      if (error) {
        const msg = (error?.message || String(error)).toLowerCase();
        const missingGeo =
          msg.includes("latitude") ||
          msg.includes("longitude") ||
          msg.includes("photo_url") ||
          msg.includes("photo_taken_at");
        if (!missingGeo) throw error;

        // Backward compatibility: database exists but hasn't been migrated yet
        const payloadCompat = { ...payload };
        delete payloadCompat.latitude;
        delete payloadCompat.longitude;
        delete payloadCompat.photo_url;
        delete payloadCompat.photo_taken_at;
        if (
          occupancyEditingId &&
          (!payloadCompat.owner_name || String(payloadCompat.owner_name).trim() === "")
        ) {
          delete payloadCompat.owner_name;
        }
        const retry = occupancyEditingId
          ? await q.update(payloadCompat).eq("id", occupancyEditingId)
          : await q.insert(payloadCompat);
        if (retry.error) throw retry.error;
      }
      await occupancyLoadFromSupabase();
      occupancyRenderTable();
      renderOccupancyMarkersBatched();
      logbookShowToast("occupancy-toast", "Saved to database.");

      const hasLocation = entry.lat != null && entry.lng != null;
      if (hasLocation) {
        showView("map");
        window.location.hash = "map";
        closeNavSidebar();
        setTimeout(() => {
          if (mapInstance) {
            mapInstance.setView([entry.lat, entry.lng], 16);
            openOccupancyDetailPanel(entry);
          }
        }, 100);
      } else {
        showView("occupancy");
        window.location.hash = "occupancy";
        setTimeout(() => {
          const idx = occupancyData.findIndex((r) => r.io_number === entry.io_number);
          if (idx >= 0) {
            const rowEl = document.getElementById(`occupancy-row-${idx}`);
            if (rowEl) {
              rowEl.classList.add("row-highlight");
              rowEl.scrollIntoView({ behavior: "smooth", block: "center" });
              setTimeout(() => rowEl.classList.remove("row-highlight"), 2500);
            }
          }
        }, 200);
      }
    } catch (err) {
      const msg = err?.message || String(err);
      const hint =
        msg.includes("policy") || msg.includes("RLS") || err?.code === "42501"
          ? " Check Supabase: add anon RLS policy (see fsis.logger.sql)."
          : "";
      logbookShowToast("occupancy-toast", "Save failed: " + msg + hint);
    }
  })();
}

async function occupancyLoadFromSupabase() {
  const selectWithGeo =
    "id, log_date, io_number, owner_name, inspectors, remarks_signature, latitude, longitude, photo_url, photo_taken_at, created_at";
  const selectWithoutGeo =
    "id, log_date, io_number, owner_name, inspectors, remarks_signature, created_at";

  const run = async (select) =>
    await supabaseClient
      .from("occupancy_logbook")
      .select(select)
      .order("created_at", { ascending: true })
      .limit(2000);

  let rows;
  {
    const { data, error } = await run(selectWithGeo);
    if (!error) rows = data;
    else {
      const msg = (error?.message || String(error)).toLowerCase();
      const missingGeo =
        msg.includes("latitude") ||
        msg.includes("longitude") ||
        msg.includes("photo_url") ||
        msg.includes("photo_taken_at");
      if (!missingGeo) throw error;

      const retry = await run(selectWithoutGeo);
      if (retry.error) throw retry.error;
      rows = retry.data;
    }
  }

  occupancyData = (rows || []).map((r) => ({
    id: r.id,
    log_date: r.log_date,
    io_number: r.io_number,
    owner_name: r.owner_name || "",
    inspectors: r.inspectors,
    remarks_signature: r.remarks_signature,
    lat: r.latitude ?? null,
    lng: r.longitude ?? null,
    photo_url:
      r.latitude != null && r.longitude != null ? r.photo_url ?? null : null,
    photo_taken_at:
      r.latitude != null && r.longitude != null ? r.photo_taken_at ?? null : null,
    created_at: r.created_at,
  }));
  occupancySaveToLocal();
}

async function occupancyInitData() {
  if (occupancyDataLoaded) return;
  occupancyDataLoaded = true;
  try {
    occupancyLoadFromLocal();
    occupancyRenderTable();
    renderOccupancyMarkersBatched();
    if (isSupabaseEnabled()) {
      await occupancyLoadFromSupabase();
      occupancyRenderTable();
      renderOccupancyMarkersBatched();
    }
  } catch (err) {
    occupancyLoadFromLocal();
    console.warn("Occupancy load failed, using local storage:", err);
    logbookShowToast("occupancy-toast", "Using data stored on this device.");
    occupancyRenderTable();
    renderOccupancyMarkersBatched();
  }
}
