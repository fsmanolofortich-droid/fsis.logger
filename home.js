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

/** Coerce sheet/map coordinates to finite numbers (EXIF or Sheets may yield strings). */
function normalizeGeoNumber(v) {
  if (v == null || v === "") return null;
  const n = typeof v === "number" ? v : Number.parseFloat(String(v).trim());
  return Number.isFinite(n) ? n : null;
}

function getCurrentView() {
  const hash = (window.location.hash || "#map").replace(/^#/, "");
  if (
    hash === "inspection" ||
    hash === "fsec" ||
    hash === "conveyance" ||
    hash === "fire_drill" ||
    hash === "occupancy"
  ) {
    return hash;
  }
  return "map";
}

function toggleFilters(logbookType) {
  const filterPanel = document.getElementById(`filter-${logbookType}`);
  if (!filterPanel) return;

  if (filterPanel.style.display === "none") {
    filterPanel.style.display = "flex";
  } else {
    filterPanel.style.display = "none";
  }
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
let occupancyExifProcessingPromise = null;

/** Pending photo picker flow: confirm in preview modal before committing. */
let photoPreviewContext = null;

let inspectionMarkersLayer = null;
let occupancyMarkersLayer = null;
let inspectionDataLoaded = false;
let inspectionActiveTab = "with-location";
let inspectionFocusMapAfterSave = false;

let mapMarkerFilter = "all"; // all | businesses | occupancies | Mercantile | Storage | etc

function applyMapMarkerFilter(next) {
  mapMarkerFilter = next || "all";
  if (!mapInstance) return;
  
  // Re-render all markers with the new filter
  renderInspectionMarkersBatched();
  renderOccupancyMarkersBatched();

  // If there's an active map search, refresh it
  const q = document.getElementById("map-search-input")?.value;
  if (q) searchMapLocations(q);
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
      {
        "Inspection Flags": inspectionMarkersLayer,
        "Occupancy Flags": occupancyMarkersLayer,
      },
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

// Used so a second showView(sameName) from hashchange does not scroll to top
// (e.g. after "View in inspection logbook" scrolls the matching row into view).
let lastAppliedViewName = null;

function showView(name) {
  const viewChanged = lastAppliedViewName !== name;

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
  } else if (name === "fire_drill" && !fireDrillDataLoaded) {
    fireDrillInitData();
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

  // Reset scroll only when switching to a different main view — not when hashchange
  // re-applies the same view (that would cancel scrollIntoView on a logbook row).
  if (viewChanged) {
    window.scrollTo({ top: 0, behavior: "auto" });
  }
  lastAppliedViewName = name;
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
  const el = document.getElementById("navSidebar");
  if (!el) return;
  // Bootstrap 5 Offcanvas API
  let bsOffcanvas = bootstrap?.Offcanvas?.getInstance(el);
  if (!bsOffcanvas) bsOffcanvas = new bootstrap.Offcanvas(el);
  bsOffcanvas.show();
}

function closeNavSidebar() {
  const el = document.getElementById("navSidebar");
  if (!el) return;
  const bsOffcanvas = bootstrap?.Offcanvas?.getInstance(el);
  if (bsOffcanvas) bsOffcanvas.hide();
}

function toggleNavSidebar() {
  const el = document.getElementById("navSidebar");
  if (!el) return;
  let bsOffcanvas = bootstrap?.Offcanvas?.getInstance(el);
  if (!bsOffcanvas) bsOffcanvas = new bootstrap.Offcanvas(el);
  bsOffcanvas.toggle();
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
  fillSelect("occupancy_addr_barangay", BARANGAYS, "Select barangay");
  fillSelect("fire_drill_addr_barangay", BARANGAYS, "Select barangay");

  fillSelect("inspection_inspected_by", FIRE_PERSONNEL, "Select inspector");
  fillSelect("inspection_included_personnel_name", FIRE_PERSONNEL, "Select personnel (optional)");
  fillSelect("occupancy_inspected_by", FIRE_PERSONNEL, "Select inspector");
  fillSelect("occupancy_included_personnel_name", FIRE_PERSONNEL, "Select personnel (optional)");
  fillSelect("conveyance_inspected_by", FIRE_PERSONNEL, "Select inspector");
  fillSelect("conveyance_included_personnel_name", FIRE_PERSONNEL, "Select personnel (optional)");
  
  // Auto-fill rank/position when fire personnel is selected
  bindInspectionPersonnelAutoFill();
  bindOccupancyPersonnelAutoFill();
  bindConveyancePersonnelAutoFill();
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

function bindOccupancyPersonnelAutoFill() {
  const inspectedBy = document.getElementById("occupancy_inspected_by");
  const inspectorPos = document.getElementById("occupancy_inspector_position");
  const includedPersonnel = document.getElementById("occupancy_included_personnel_name");
  const includedPos = document.getElementById("occupancy_included_personnel_position");

  inspectedBy?.addEventListener("change", () => {
    if (inspectorPos) inspectorPos.value = getFirePersonnelRankPositionByName(inspectedBy.value);
  });
  includedPersonnel?.addEventListener("change", () => {
    if (includedPos) includedPos.value = getFirePersonnelRankPositionByName(includedPersonnel.value);
  });
}

function bindConveyancePersonnelAutoFill() {
  const inspectedBy = document.getElementById("conveyance_inspected_by");
  const inspectorPos = document.getElementById("conveyance_inspector_position");
  const includedPersonnel = document.getElementById("conveyance_included_personnel_name");
  const includedPos = document.getElementById("conveyance_included_personnel_position");

  inspectedBy?.addEventListener("change", () => {
    if (inspectorPos) inspectorPos.value = getFirePersonnelRankPositionByName(inspectedBy.value);
  });
  includedPersonnel?.addEventListener("change", () => {
    if (includedPos) includedPos.value = getFirePersonnelRankPositionByName(includedPersonnel.value);
  });
}

// Legacy helper kept for compatibility; no longer used by the current UI.
function conveyanceAddInspector() { }
function occupancyAddInspector() { }

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

  // Burger button open/close is handled by Bootstrap Offcanvas (data-bs-toggle).
  // We only need to close the sidebar when a nav overlay is clicked — BS handles that via its own backdrop.

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
  const fdIssued = document.getElementById("fire_drill_date_issued");
  if (fdIssued) {
    fdIssued.addEventListener("change", fireDrillSyncIssuanceFieldsFromDate);
    fdIssued.addEventListener("input", fireDrillSyncIssuanceFieldsFromDate);
  }
  const fdValidityType = document.getElementById("fire_drill_validity_type");
  if (fdValidityType) {
    fdValidityType.addEventListener("change", fireDrillSyncValidityDate);
    fdValidityType.addEventListener("input", fireDrillSyncValidityDate);
  }
  initViewRouting();  // Map markers ui filter initialization removed as it uses the select dropdown now.
  initTableFilters();
  initPhotoPreviewModal();
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
        try { ev.preventDefault(); } catch { }
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
  } else if (initialView === "fire_drill" && !fireDrillDataLoaded) {
    fireDrillInitData();
  } else if (initialView === "occupancy" && !occupancyDataLoaded) {
    occupancyInitData();
  }
}

document.addEventListener("DOMContentLoaded", init);

// ── Burger button 10-click easter egg → Dashboard ──────────────────────────
let _burgerClickCount = 0;
let _burgerClickTimer = null;
function burgerDashTrigger() {
  _burgerClickCount++;
  if (_burgerClickTimer) clearTimeout(_burgerClickTimer);
  if (_burgerClickCount >= 10) {
    _burgerClickCount = 0;
    window.open('./dashboard.html', '_blank');
    return;
  }
  // Reset counter if no click within 2s
  _burgerClickTimer = setTimeout(() => { _burgerClickCount = 0; }, 2000);
}

// -----------------------------
// Shared GAS client + utilities
// -----------------------------

// Paste your deployed Google Apps Script Web App URL here
const GAS_URL = "https://script.google.com/macros/s/AKfycbwJmqg6lRB_W95VNY9XfAyAovcbJrm8VpPXXg1pP1ujFD10k85xTpbwO5v8RVyy8Bpc/exec";

function isGasEnabled() {
  return Boolean(GAS_URL);
}

// Alias so existing callers (isSupabaseEnabled) still work
function isSupabaseEnabled() {
  return isGasEnabled();
}

/**
 * Send a request to the GAS Web App backend.
 * All actions go via HTTP POST with JSON body.
 */
async function gasRequest(action, payload) {
  const res = await fetch(GAS_URL, {
    method: "POST",
    body: JSON.stringify({ action, ...(payload || {}) }),
  });
  const json = await res.json();
  if (json.error) throw new Error(json.error);
  return json;
}

/**
 * Convert a File or Blob to a base64 string (without the data: prefix)
 * for sending to the GAS upload action.
 */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
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
  if (!isGasEnabled()) {
    setStorageBadge("local");
    return;
  }

  try {
    await gasRequest("ping");
    setStorageBadge("db");
  } catch (err) {
    console.warn("GAS ping failed, using local mode:", err);
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
  let dateObj;
  if (typeof d === 'string' && !d.includes('T')) {
    dateObj = new Date(d + "T00:00:00");
  } else {
    dateObj = new Date(d);
  }
  if (isNaN(dateObj.getTime())) return "Invalid Date";

  return dateObj.toLocaleDateString("en-PH", {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
}

/**
 * Returns a YYYY-MM-DD string for use in <input type="date"> value.
 */
function logbookFormatDateForInput(d) {
  if (!d) return "";
  let dateObj;
  if (typeof d === 'string' && !d.includes('T')) {
    // Already in a simple date format?
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
    dateObj = new Date(d + "T00:00:00");
  } else {
    dateObj = new Date(d);
  }
  if (isNaN(dateObj.getTime())) return "";
  
  // Use local components to avoid timezone shifts for simple calendar dates
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const day = String(dateObj.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
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
  let t;
  if (typeof dateStr === 'string' && !dateStr.includes('T')) {
    t = new Date(dateStr + "T00:00:00").getTime();
  } else {
    t = new Date(dateStr).getTime();
  }
  if (!isFinite(t)) return false;

  if (fromStr) {
    const f = new Date(fromStr + "T00:00:00").getTime();
    if (isFinite(f) && t < f) return false;
  }
  if (toStr) {
    const to = new Date(toStr + "T23:59:59").getTime();
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
    ["inspection-filter-q", "inspection-filter-from", "inspection-filter-to"],
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
    ["fire_drill-filter-q", "fire_drill-filter-from", "fire_drill-filter-to"],
    () => fireDrillRenderTable()
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

function extractAddressParts(row) {
  let addrLine = row.addr_line || "";
  let addrBarangay = row.addr_barangay || "";
  let addrMunicipal = row.addr_municipal || "";
  let addrProvince = row.addr_province || "";
  let addrRegion = row.addr_region || "";
  const fullAddr = (row.insp_address || row.fsec_address || row.address || "").toString().trim();
  
  if ((!addrLine || !addrBarangay || !addrMunicipal || !addrProvince || !addrRegion) && fullAddr) {
    const parts = fullAddr.split(/,\s*/).map((p) => String(p || "").trim()).filter(Boolean);
    if (parts.length > 1 && parts[0].toLowerCase().indexOf("region") !== -1) {
      parts.reverse();
    }
    if (!addrBarangay) {
      const brgy = parts.find(p => /^(barangay|brgy)/i.test(p));
      if (brgy) addrBarangay = brgy.replace(/^(barangay|brgy)\.?\s+/i, "");
    }
    if (parts.length >= 3) {
      addrLine = addrLine || parts[0];
      addrMunicipal = addrMunicipal || parts[parts.length - 3];
      addrProvince = addrProvince || parts[parts.length - 2];
      addrRegion = addrRegion || parts[parts.length - 1];
    }
  }

  const finalParts = [];
  if (addrLine) finalParts.push(addrLine);
  if (addrBarangay) {
    const cleanBrgy = addrBarangay.replace(/^(barangay|brgy)\.?\s+/i, "");
    finalParts.push("Barangay " + cleanBrgy);
  }
  if (addrMunicipal) finalParts.push(addrMunicipal);
  if (addrProvince) finalParts.push(addrProvince);
  if (addrRegion) finalParts.push(addrRegion);

  const seen = new Set();
  const uniqueParts = [];
  finalParts.forEach(p => {
    const lower = p.toLowerCase().trim();
    if (!seen.has(lower)) {
      seen.add(lower);
      uniqueParts.push(p);
    }
  });

  return { uniqueParts, fullAddr, addrLine, cleanBrgy: addrBarangay.replace(/^(barangay|brgy)\.?\s+/i, "") };
}

function inspectionFormatAddressDisplay(row) {
  const { uniqueParts, fullAddr } = extractAddressParts(row);
  return uniqueParts.length > 0 ? uniqueParts.join(", ") : (fullAddr || "—");
}

// Short address for table display — strips the repeated municipality/province/region
// to keep the address column compact. Full version still used for print and detail panels.
function inspectionFormatAddressShort(row) {
  const { addrLine, cleanBrgy, fullAddr } = extractAddressParts(row);
  if (addrLine || cleanBrgy) {
    const parts = [];
    if (addrLine) parts.push(addrLine);
    if (cleanBrgy) parts.push("Brgy. " + cleanBrgy);
    return parts.join(", ") || "—";
  }
  return fullAddr || "—";
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
  const totalWithLocation = inspectionData.filter((r) => r.lat != null && r.lng != null).length;
  const totalNoLocation = inspectionData.filter((r) => r.lat == null || r.lng == null).length;

  const countBadge = document.getElementById("inspection-record-count");
  if (countBadge) countBadge.textContent = String(totalWithLocation);
  const noPhotoBadge = document.getElementById("inspection-nophoto-record-count");
  if (noPhotoBadge) noPhotoBadge.textContent = String(totalNoLocation);

  const q = normalizeQuery(document.getElementById("inspection-filter-q")?.value);
  const from = (document.getElementById("inspection-filter-from")?.value || "").trim();
  const to = (document.getElementById("inspection-filter-to")?.value || "").trim();
  const filtered = inspectionData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (from || to) {
        if (!inDateRange(row.date_inspected, from, to)) return false;
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

  // If there are inspection records but the current filters hide everything,
  // show the empty state (otherwise it looks like "search/filter not working").
  if (inspectionData.length > 0 && filtered.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";

    if (emptyNoPhoto) emptyNoPhoto.style.display = "block";
    if (tableWrapNoPhoto) tableWrapNoPhoto.style.display = "none";

    return;
  }

  if (inspectionData.length === 0) {
    empty.style.display = "block";
    if (tableWrap) tableWrap.style.display = "none";
  } else {
    empty.style.display = "none";
    if (tableWrap) tableWrap.style.display = "";

    filtered.forEach(({ row, idx }) => {
      const hasLocation = row.lat != null && row.lng != null;
      // "#" is 1,2,3… within each tab (with location vs no location), not across both.
      let rowNum;
      if (hasLocation) {
        withLocationCount++;
        rowNum = withLocationCount;
      } else {
        noLocationCount++;
        rowNum = noLocationCount;
      }

      const baseRowHtml = `
        <td data-label="#">${rowNum}</td>
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
    } else {
      emptyNoPhoto.style.display = "none";
      if (tableWrapNoPhoto) tableWrapNoPhoto.style.display = "";
    }
  }

  // ── Filter result info bars ──────────────────────────────────────────
  const isFiltered = !!(q || from || to);
  const resultsBadge = document.getElementById("inspection-results-badge");
  if (resultsBadge) {
    if (isFiltered && inspectionData.length > 0) {
      resultsBadge.textContent =
        `Showing ${withLocationCount} of ${totalWithLocation} records (with location)`;
      resultsBadge.removeAttribute("hidden");
    } else {
      resultsBadge.setAttribute("hidden", "");
    }
  }

  const noPhotoResultsBadge = document.getElementById("inspection-nophoto-results-badge");
  if (noPhotoResultsBadge) {
    if (isFiltered && inspectionData.length > 0) {
      noPhotoResultsBadge.textContent =
        `Showing ${noLocationCount} of ${totalNoLocation} records (no location)`;
      noPhotoResultsBadge.removeAttribute("hidden");
    } else {
      noPhotoResultsBadge.setAttribute("hidden", "");
    }
  }
}

async function inspectionEditEntry(idx) {
  const oldRow = inspectionData[idx];
  if (!oldRow) return;

  if (isGasEnabled()) {
    logbookShowToast("inspection-toast", "Refreshing record data...");
    try {
      await inspectionLoadFromSupabase();
    } catch (err) {
      console.warn("Refresh failed:", err);
    }
  }

  // Re-find the row in case indices changed or data was refreshed
  const row = (oldRow.id) ? inspectionData.find(r => r.id === oldRow.id) : inspectionData[idx];
  if (!row) return;

  inspectionEditingIdx = inspectionData.indexOf(row);
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
  setVal("inspection_date_inspected", logbookFormatDateForInput(row.date_inspected));
  // Optional IO-specific fields (may not exist on older records or in the DOM)
  setVal("inspection_inspector_position", row.inspector_position);
  setVal("inspection_included_personnel_name", row.included_personnel_name);
  setVal(
    "inspection_included_personnel_position",
    row.included_personnel_position
  );
  setVal("inspection_duration_start", logbookFormatDateForInput(row.duration_start));
  setVal("inspection_duration_end", logbookFormatDateForInput(row.duration_end));

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

  // Reset to first step
  updateModalStepUI('inspection', 1);
}

/**
 * Dedicated "add photo late" — picks a file, uploads to Drive, patches only
 * photo_url in the sheet. Bypasses the full edit-form flow entirely.
 */
async function inspectionAddPhoto(idx) {
  const row = inspectionData[idx];
  if (!row) return;
  if (!row.id) {
    logbookShowToast("inspection-toast", "⚠️ Record must be saved to database first.");
    return;
  }
  if (!isGasEnabled()) {
    logbookShowToast("inspection-toast", "⚠️ Database not connected.");
    return;
  }

  // Trigger native file picker — no modal needed
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.onchange = async function () {
    const file = input.files[0];
    if (!file) return;

    void beginPhotoAttachFromPicker("inspection_direct", file, null, {
      onConfirm: async ({ file, dataUrl, meta }) => {
        logbookShowToast("inspection-toast", "Uploading photo to Drive…");
        try {
          let uploadFile = file;
          try { uploadFile = await sanitizeInspectionImage(file); } catch (e) {}

          const base64Data = await fileToBase64(uploadFile);
          const uploadResult = await gasRequest("upload", {
            filename: `inspection-${Date.now()}.${(file.name || "file.bin").split(".").pop() || "bin"}`,
            mimeType: file.type || "application/octet-stream",
            base64Data,
          });

          if (!uploadResult?.data?.url) {
            logbookShowToast("inspection-toast", "⚠️ Upload returned no URL — check Drive folder permissions.");
            return;
          }

          const driveUrl = uploadResult.data.url;
          const exLat = normalizeGeoNumber(meta?.gps?.lat);
          const exLng = normalizeGeoNumber(meta?.gps?.lng);

          await gasRequest("patch_photo_url", {
            table: "inspection_logbook",
            id: row.id,
            url: driveUrl,
          });

          if (exLat != null && exLng != null) {
            try {
              await gasRequest("patch_lat_lng", {
                table: "inspection_logbook",
                id: row.id,
                latitude: exLat,
                longitude: exLng,
              });
            } catch (err) { }
          }

          if (inspectionData[idx]) {
            inspectionData[idx].photo_url = driveUrl;
            if (exLat != null && exLng != null) {
              inspectionData[idx].lat = exLat;
              inspectionData[idx].lng = exLng;
            }
            inspectionSaveToLocal();
            inspectionRenderTable();
            renderInspectionMarkersBatched();
          }
          logbookShowToast("inspection-toast", "✓ Photo & EXIF saved to Drive and sheet.");
        } catch (err) {
          console.error("inspectionAddPhoto error:", err);
          logbookShowToast("inspection-toast", "⚠️ Failed: " + (err?.message || err));
        }
      }
    });
  };
  input.click();
}

/**
 * Dedicated "add photo late" for Occupancy records.
 */
async function occupancyAddPhoto(idx) {
  const row = occupancyData[idx];
  if (!row) return;
  if (!row.id) {
    logbookShowToast("occupancy-toast", "⚠️ Record must be saved to database first.");
    return;
  }
  if (!isGasEnabled()) {
    logbookShowToast("occupancy-toast", "⚠️ Database not connected.");
    return;
  }

  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.onchange = async function () {
    const file = input.files[0];
    if (!file) return;

    void beginPhotoAttachFromPicker("occupancy_direct", file, null, {
      onConfirm: async ({ file, dataUrl, meta }) => {
        logbookShowToast("occupancy-toast", "Uploading photo to Drive…");
        try {
          let uploadFile = file;
          try { uploadFile = await sanitizeInspectionImage(file); } catch (e) {}

          const base64Data = await fileToBase64(uploadFile);
          const uploadResult = await gasRequest("upload", {
            filename: `occupancy-${Date.now()}.${(file.name || "file.bin").split(".").pop() || "bin"}`,
            mimeType: file.type || "application/octet-stream",
            base64Data,
          });

          if (!uploadResult?.data?.url) {
            logbookShowToast("occupancy-toast", "⚠️ Upload returned no URL — check Drive folder permissions.");
            return;
          }

          const driveUrl = uploadResult.data.url;
          const exLat = normalizeGeoNumber(meta?.gps?.lat);
          const exLng = normalizeGeoNumber(meta?.gps?.lng);

          await gasRequest("patch_photo_url", {
            table: "occupancy_logbook",
            id: row.id,
            url: driveUrl,
          });

          if (exLat != null && exLng != null) {
            try {
              await gasRequest("patch_lat_lng", {
                table: "occupancy_logbook",
                id: row.id,
                latitude: exLat,
                longitude: exLng,
              });
            } catch (err) { }
          }

          if (occupancyData[idx]) {
            occupancyData[idx].photo_url = driveUrl;
            if (exLat != null && exLng != null) {
              occupancyData[idx].lat = exLat;
              occupancyData[idx].lng = exLng;
            }
            occupancyRenderTable();
            renderOccupancyMarkersBatched();
          }
          logbookShowToast("occupancy-toast", "✓ Photo & EXIF saved to Drive and sheet.");
        } catch (err) {
          console.error("occupancyAddPhoto error:", err);
          logbookShowToast("occupancy-toast", "⚠️ Failed: " + (err?.message || err));
        }
      }
    });
  };
  input.click();
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

let clearanceEditingIdx = null;

function inspectionOpenClearanceHtml(idx) {
  const row = inspectionData[idx];
  if (!row) return;
  clearanceEditingIdx = idx;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };

  setVal("clearance_fsic_number", row.fsic_number);
  setVal("clearance_purpose", row.fsic_purpose);
  setVal("clearance_valid_until", row.fsic_valid_until);
  setVal("clearance_fee_amount", row.fsic_fee_amount);
  setVal("clearance_fee_or_number", row.fsic_fee_or_number);
  setVal("clearance_fee_date", row.fsic_fee_date);

  const overlay = document.getElementById("clearance-modal-overlay");
  if (overlay) overlay.classList.add("open");
}

function clearanceCloseModal() {
  const overlay = document.getElementById("clearance-modal-overlay");
  if (overlay) overlay.classList.remove("open");
}

function clearanceCloseOnOverlay(e) {
  const overlay = document.getElementById("clearance-modal-overlay");
  if (e.target === overlay) clearanceCloseModal();
}

function clearanceProceed(e) {
  if (e?.preventDefault) e.preventDefault();

  if (clearanceEditingIdx === null) return;
  const row = inspectionData[clearanceEditingIdx];
  if (!row) return;

  const getVal = (id) => (document.getElementById(id) || { value: "" }).value.trim();

  // We save the inputs back to the row object that gets passed to the template
  const payload = {
    ...row,
    fsic_number: getVal("clearance_fsic_number"),
    fsic_purpose: getVal("clearance_purpose"),
    fsic_valid_from: row.business_name, // Hardcoded to business_name per user request
    fsic_valid_until: getVal("clearance_valid_until"),
    fsic_fee_amount: getVal("clearance_fee_amount"),
    fsic_fee_or_number: getVal("clearance_fee_or_number"),
    fsic_fee_date: getVal("clearance_fee_date"),
  };

  try {
    sessionStorage.setItem("fsis.clearance.current", JSON.stringify(payload));
  } catch {
    // If sessionStorage is unavailable, we still open the template.
  }

  clearanceCloseModal();
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

  let sourceLogbook = "inspection";
  let dataset = inspectionData;
  if (d.payload?._sourceType === "occupancy") {
    sourceLogbook = "occupancy";
    dataset = occupancyData;
  }

  const idx = Array.isArray(dataset)
    ? dataset.findIndex((r) => (entryId && r?.id === entryId) || (ioNumber && r?.io_number === ioNumber))
    : -1;

  if (idx < 0) {
    respond(false, `Cannot find matching ${sourceLogbook} record.`);
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
  if (has("fsic_purpose")) updates.fsic_purpose = (payload.fsic_purpose || "").toString().trim();
  if (has("fsic_valid_until")) updates.fsic_valid_until = normalizeDate(payload.fsic_valid_until);
  if (has("fsic_fee_amount")) updates.fsic_fee_amount = normalizeAmount(payload.fsic_fee_amount);
  if (has("fsic_fee_or_number")) updates.fsic_fee_or_number = (payload.fsic_fee_or_number || "").toString().trim();
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

  if (sourceLogbook === "inspection") {
    inspectionData[idx] = { ...inspectionData[idx], ...uiUpdates };
    inspectionSaveToLocal();
    inspectionRenderTable?.();
  } else {
    occupancyData[idx] = { ...occupancyData[idx], ...uiUpdates };
    occupancySaveToLocal();
    occupancyRenderTable?.();
  }

  if (!isSupabaseEnabled()) {
    logbookShowToast?.(`${sourceLogbook}-toast`, "Saved on this device only (offline mode).");
    respond(true, "Saved locally (offline mode).");
    return;
  }

  const datasetRef = sourceLogbook === "inspection" ? inspectionData : occupancyData;
  const row = datasetRef[idx];
  if (!row?.id) {
    logbookShowToast?.(`${sourceLogbook}-toast`, "⚠️ Save failed: missing record id.");
    respond(false, "Missing record id.");
    return;
  }

  (async () => {
    try {
      await gasRequest("update", { table: `${sourceLogbook}_logbook`, id: row.id, row: updates });
      logbookShowToast?.(`${sourceLogbook}-toast`, "Saved to database.");
      respond(true, "");
    } catch (err) {
      const msg = err?.message || String(err);
      logbookShowToast?.(`${sourceLogbook}-toast`, "Save failed: " + msg);
      respond(false, msg);
    }
  })();
});

// Receive IO edits from `inspection_io_fsis.html` and `occupancy_io_fsis.html` and persist them.
window.addEventListener("message", (ev) => {
  const d = ev?.data;
  if (!d || d.type !== "fsis_io_save") return;

  const payload = d.payload || {};
  const entryId = d.entryId || null;
  const ioNumber = d.io_number || null;
  const table = d.table || "inspection_logbook"; 

  function respond(ok, message) {
    try {
      ev.source?.postMessage({ type: "fsis_io_save_result", ok, message }, "*");
    } catch {
      // ignore
    }
  }

  function normalizeDate(value) {
    if (!value) return null;
    const s = String(value).trim();
    if (!s) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const dt = new Date(s);
    if (Number.isNaN(dt.getTime())) return null;
    return dt.toISOString().slice(0, 10);
  }

  let ds, idx;
  if (table === "occupancy_logbook") {
    ds = occupancyData;
    idx = Array.isArray(ds) ? ds.findIndex((r) => (entryId && r?.id === entryId) || (ioNumber && r?.io_number === ioNumber)) : -1;
  } else {
    ds = inspectionData;
    idx = Array.isArray(ds) ? ds.findIndex((r) => (entryId && r?.id === entryId) || (ioNumber && r?.io_number === ioNumber)) : -1;
  }

  if (idx < 0) {
    respond(false, "Cannot find matching IO record.");
    return;
  }

  const updates = {};
  const has = (k) => Object.prototype.hasOwnProperty.call(payload, k);

  if (has("io_number")) updates.io_number = (payload.io_number || "").toString().trim();
  if (has("date_inspected") && table === "inspection_logbook") updates.date_inspected = normalizeDate(payload.date_inspected);
  if (has("log_date") && table === "occupancy_logbook") updates.log_date = normalizeDate(payload.log_date);
  
  if (has("inspectors")) updates.inspectors = (payload.inspectors || "").toString().trim();
  if (has("inspector_position")) updates.inspector_position = (payload.inspector_position || "").toString().trim();
  if (has("included_personnel_name")) updates.included_personnel_name = (payload.included_personnel_name || "").toString().trim();
  if (has("included_personnel_position")) updates.included_personnel_position = (payload.included_personnel_position || "").toString().trim();
  if (has("duration_start")) updates.duration_start = normalizeDate(payload.duration_start);
  if (has("duration_end")) updates.duration_end = normalizeDate(payload.duration_end);
  if (has("io_remarks")) updates.io_remarks = (payload.io_remarks || "").toString().trim();

  // Update in-memory + local cache
  ds[idx] = { ...ds[idx], ...updates };
  
  if (table === "occupancy_logbook") {
    occupancySaveToLocal();
    occupancyRenderTable?.();
  } else {
    inspectionSaveToLocal();
    inspectionRenderTable?.();
  }

  if (!isSupabaseEnabled()) {
    logbookShowToast?.("inspection-toast", "Saved on this device only (offline mode).");
    respond(true, "Saved locally (offline mode).");
    return;
  }

  const row = ds[idx];
  if (!row?.id) {
    respond(false, "No valid ID for cloud update.");
    return;
  }

  app_save(row.id, updates, table)
    .then((res) => {
      if (res?.error) throw new Error(res.error);
      respond(true, "Saved to database!");
    })
    .catch((err) => {
      console.error("IO Save error:", err);
      respond(false, err?.message || "Failed to save to cloud");
    });
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
      await gasRequest("delete", { table: "inspection_logbook", id: row.id });
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

  const indicator = document.getElementById("inspection-photo-indicator");
  if (indicator) {
    indicator.className = "photo-attach-indicator";
    indicator.textContent = "";
  }

  // Reset to first step
  updateModalStepUI('inspection', 1);
}

function inspectionCloseModal() {
  photoPreviewCancel();
  const overlay = document.getElementById("inspection-modal-overlay");
  if (overlay) overlay.classList.remove("open");
}

/* ── Stepped Modal Logic ────────────────────────────────────────────── */

/**
 * Universal step UI updater
 * @param {string} prefix - 'inspection' or 'occupancy'
 * @param {number} step - Current active step (1-based)
 */
function updateModalStepUI(prefix, step) {
  const modalElem = document.getElementById(`${prefix}-modal-overlay`);
  if (!modalElem) return;

  // Track state on the element
  modalElem.setAttribute('data-current-step', step);

  // Toggle step content
  const contents = modalElem.querySelectorAll('.step-content');
  contents.forEach(div => {
    const s = parseInt(div.getAttribute('data-step'), 10);
    div.classList.toggle('active', s === step);
  });

  // Toggle step indicators
  const indicators = modalElem.querySelectorAll('.step-item');
  indicators.forEach(item => {
    const s = parseInt(item.getAttribute('data-step'), 10);
    item.classList.toggle('active', s === step);
    item.classList.toggle('completed', s < step);
  });

  // Buttons management
  const btnBack = document.getElementById(`${prefix}-btn-back`);
  const btnCancel = document.getElementById(`${prefix}-btn-cancel`);
  const btnNext = document.getElementById(`${prefix}-btn-next`);
  const btnSave = document.getElementById(`${prefix}-btn-save`);

  const maxSteps = contents.length;

  if (btnBack) btnBack.style.display = (step > 1) ? 'inline-block' : 'none';
  if (btnCancel) btnCancel.style.display = (step === 1) ? 'inline-block' : 'none';
  if (btnNext) btnNext.style.display = (step < maxSteps) ? 'inline-block' : 'none';
  if (btnSave) btnSave.style.display = (step === maxSteps) ? 'inline-block' : 'none';

  // Smooth scroll modal to top on step change
  const modalBody = modalElem.querySelector('.modal-body');
  if (modalBody) modalBody.scrollTop = 0;
}

function inspectionModalStep(delta) {
  const overlay = document.getElementById("inspection-modal-overlay");
  let current = parseInt(overlay.getAttribute('data-current-step') || '1', 10);
  current += delta;
  if (current < 1) current = 1;
  const max = 4; // Inspection has 4 steps
  if (current > max) current = max;
  updateModalStepUI('inspection', current);
}

function occupancyModalStep(delta) {
  const overlay = document.getElementById("occupancy-modal-overlay");
  let current = parseInt(overlay.getAttribute('data-current-step') || '1', 10);
  current += delta;
  if (current < 1) current = 1;
  const max = 4; // Occupancy has 4 steps
  if (current > max) current = max;
  updateModalStepUI('occupancy', current);
}

function fsecModalStep(delta) {
  const overlay = document.getElementById("fsec-modal-overlay");
  let current = parseInt(overlay.getAttribute('data-current-step') || '1', 10);
  current += delta;
  if (current < 1) current = 1;
  const max = 3; // FSEC has 3 steps
  if (current > max) current = max;
  updateModalStepUI('fsec', current);
}

function conveyanceModalStep(delta) {
  const overlay = document.getElementById("conveyance-modal-overlay");
  let current = parseInt(overlay.getAttribute('data-current-step') || '1', 10);
  current += delta;
  if (current < 1) current = 1;
  const max = 2; // Conveyance has 2 steps
  if (current > max) current = max;
  updateModalStepUI('conveyance', current);
}

function fireDrillModalStep(delta) {
  const overlay = document.getElementById("fire_drill-modal-overlay");
  if (!overlay) return;
  let current = parseInt(overlay.getAttribute("data-current-step") || "1", 10);
  current += delta;
  if (current < 1) current = 1;
  const max = 2;
  if (current > max) current = max;
  updateModalStepUI("fire_drill", current);
  if (current === 2 && delta > 0) {
    const iss = document.getElementById("fire_drill_date_issued");
    const cert = document.getElementById("fire_drill_certificate_date");
    if (iss && cert && !iss.value && cert.value) {
      iss.value = cert.value;
      fireDrillSyncIssuanceFieldsFromDate();
    }
  }
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

  // If the user just confirmed a photo, wait for lat/lng + optional geolocation to finish.
  if (currentExifProcessingPromise) {
    try {
      await currentExifProcessingPromise;
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
    line,
    barangay ? `Barangay ${barangay}` : null,
    municipal,
    province,
    `Region ${region}`
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
    lat: normalizeGeoNumber(currentExifLat),
    lng: normalizeGeoNumber(currentExifLng),
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
      entry.lat = normalizeGeoNumber(prev.lat ?? entry.lat);
      entry.lng = normalizeGeoNumber(prev.lng ?? entry.lng);
    } else {
      // New photo chosen:
      const exLat = normalizeGeoNumber(currentExifLat);
      const exLng = normalizeGeoNumber(currentExifLng);
      const hasExifGps = exLat != null && exLng != null;
      if (hasExifGps) {
        entry.lat = exLat;
        entry.lng = exLng;
      } else if (normalizeGeoNumber(prev.lat) != null && normalizeGeoNumber(prev.lng) != null) {
        entry.lat = normalizeGeoNumber(prev.lat);
        entry.lng = normalizeGeoNumber(prev.lng);
      }
    }
  }

  // Removed fallback to user's current location if photo has no GPS EXIF
  // as per user request: "do not put lat long if photo do not have lat long"

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
    const isEdit = inspectionEditingIdx !== null;
    if (entry.lat != null && entry.lng != null && !isEdit) {
      // Stay in logbook instead of jumping to map as per user request
      showView("inspection");
      window.location.hash = "inspection";
      setInspectionTab("with-location");
      inspectionRenderTable();
    }
    return;
  }

  (async () => {
    try {
      // ── Photo upload ──────────────────────────────────────────────────────
      let photoUploadedUrl = null;
      if (currentExifFile && isGasEnabled()) {
        let uploadFile = currentExifFile;
        try {
          uploadFile = await sanitizeInspectionImage(uploadFile);
        } catch (sanitizeErr) {
          console.warn("Image sanitization failed, using original file:", sanitizeErr);
          uploadFile = currentExifFile;
        }
        logbookShowToast("inspection-toast", "Uploading photo to Drive…");
        try {
          const base64Data = await fileToBase64(uploadFile);
          const uploadResult = await gasRequest("upload", {
            filename: `inspection-${Date.now()}.${(uploadFile.name || "file.bin").split(".").pop() || "bin"}`,
            mimeType: uploadFile.type || "application/octet-stream",
            base64Data,
          });
          if (uploadResult?.data?.url) {
            photoUploadedUrl = uploadResult.data.url;
            entry.photo_url = photoUploadedUrl;
            // Sync Drive URL back into the local in-memory record immediately
            const localIdx = inspectionEditingIdx !== null
              ? inspectionEditingIdx
              : inspectionData.length - 1;
            if (inspectionData[localIdx]) {
              inspectionData[localIdx].photo_url = photoUploadedUrl;
              inspectionSaveToLocal();
            }
            logbookShowToast("inspection-toast", "Photo uploaded ✓");
          } else {
            logbookShowToast("inspection-toast", "⚠️ Upload returned no URL. Check Drive folder permissions.");
          }
        } catch (uploadErr) {
          console.error("Photo upload error:", uploadErr);
          logbookShowToast("inspection-toast", "⚠️ Photo upload failed: " + (uploadErr?.message || uploadErr));
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
      payload.latitude = normalizeGeoNumber(payload.latitude);
      payload.longitude = normalizeGeoNumber(payload.longitude);

      // On UPDATE: don't send photo_url when it's still a data URL (e.g. upload failed),
      // so we don't overwrite the existing DB photo with a huge string.
      if (inspectionEditingId && typeof payload.photo_url === "string" && payload.photo_url.startsWith("data:")) {
        delete payload.photo_url;
      }

      // On UPDATE, don't send empty/null optional fields — keep existing DB values.
      // photo_url and latitude/longitude are also protected so they're never stripped.
      if (inspectionEditingId) {
        const requiredKeys = new Set([
          "io_number",
          "owner_name",
          "business_name",
          "address",
          "date_inspected",
          "fsic_number",
          "photo_url",     // protect — upload may have just set this
          "latitude",
          "longitude",
        ]);
        Object.keys(payload).forEach((k) => {
          if (requiredKeys.has(k)) return;
          const v = payload[k];
          if (v == null) delete payload[k];
          else if (typeof v === "string" && v.trim() === "") delete payload[k];
        });
      }

      console.log("[FSIS] DB payload photo_url:", payload.photo_url ?? "(none)");

      let savedId = inspectionEditingId; // will be set for edits
      if (inspectionEditingId) {
        await gasRequest("update", { table: "inspection_logbook", id: inspectionEditingId, row: payload });
      } else {
        const insertResult = await gasRequest("insert", { table: "inspection_logbook", row: payload });
        savedId = insertResult?.data?.id ?? null;
      }

      // ── Guaranteed photo_url patch ─────────────────────────────────────
      // If a photo was uploaded this session, write the Drive URL via the
      // dedicated patch_photo_url action which auto-creates the column if
      // missing and writes directly to the correct cell — no silent skips.
      if (photoUploadedUrl && savedId) {
        try {
          await gasRequest("patch_photo_url", {
            table: "inspection_logbook",
            id: savedId,
            url: photoUploadedUrl,
          });
        } catch (patchErr) {
          console.warn("[FSIS] photo_url patch failed:", patchErr);
        }
      }

      if (
        savedId &&
        Number.isFinite(entry.lat) &&
        Number.isFinite(entry.lng)
      ) {
        try {
          await gasRequest("patch_lat_lng", {
            table: "inspection_logbook",
            id: savedId,
            latitude: entry.lat,
            longitude: entry.lng,
          });
        } catch (patchErr) {
          console.warn("[FSIS] patch_lat_lng failed:", patchErr);
        }
      }

      await inspectionLoadFromSupabase();
      inspectionRenderTable();
      renderInspectionMarkersBatched();
      inspectionCloseModal();
      logbookShowToast("inspection-toast", photoUploadedUrl ? "Saved + photo linked ✓" : "Saved to database.");

      const hasLocation =
        Number.isFinite(entry.lat) && Number.isFinite(entry.lng);

      const isEdit = inspectionEditingId != null;
      if (hasLocation && !isEdit) {
        // No longer jumping to map as per user request.
        // Stay in the logbook and highlight the new row.
        showView("inspection");
        window.location.hash = "inspection";
        setInspectionTab("with-location");
        inspectionRenderTable();

        setTimeout(() => {
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
  const from = document.getElementById("inspection-filter-from");
  const to = document.getElementById("inspection-filter-to");
  if (q) q.value = "";
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
      } catch { }
      window.open("./inspection_io_fsis.html", "_blank");
    };
  }

  if (photoWrap && photoImg) {
    if (entry.photo_url) {
      photoImg.src = getGoogleDriveThumbnailUrl(entry.photo_url);
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
      } catch { }
      window.open("./occupancy_io_fsis.html", "_blank");
    };
  }

  if (photoWrap && photoImg) {
    if (entry.photo_url) {
      photoImg.src = getGoogleDriveThumbnailUrl(entry.photo_url);
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

  // Clear filters (do not prefill search — IO numbers like "0" looked like a broken default).
  inspectionClearFilters();

  setInspectionTab(entry.lat != null && entry.lng != null ? "with-location" : "no-location");

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
  const indicator = document.getElementById("inspection-photo-indicator");

  function onPhotoChange(sourceInput) {
    const file = sourceInput?.files?.[0];
    if (!file) return;

    currentExifLat = null;
    currentExifLng = null;
    currentExifPreviewUrl = null;
    currentExifTakenAt = null;
    currentExifFile = null;
    currentExifProcessingPromise = null;

    if (indicator) {
      indicator.className = "photo-attach-indicator";
      indicator.textContent = "Review preview…";
    }

    void beginPhotoAttachFromPicker("inspection", file, sourceInput);
  }

  inputCamera?.addEventListener("change", () => onPhotoChange(inputCamera));
  inputLibrary?.addEventListener("change", () => onPhotoChange(inputLibrary));
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

            let outMime = "image/jpeg";
            let outExt = "jpg";
            if (file.type === "image/png") {
              outMime = "image/png";
              outExt = "png";
            } else if (file.type === "image/webp") {
              outMime = "image/webp";
              outExt = "webp";
            }

            canvas.toBlob(
              (blob) => {
                if (!blob) {
                  resolve(file);
                  return;
                }
                const baseName =
                  (file.name && file.name.replace(/\.[^.]+$/, "")) || "photo";
                const sanitized = new File([blob], baseName + "." + outExt, {
                  type: outMime,
                });
                resolve(sanitized);
              },
              outMime,
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

/** Single EXIF rational component (number, Number object, or [num, den]). */
function exifRationalToNumber(v) {
  if (v == null || v === "") return NaN;
  if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
  if (typeof v === "string") {
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? n : NaN;
  }
  if (typeof v === "object" && "numerator" in v && "denominator" in v) {
    const d = Number(v.denominator);
    if (!d) return NaN;
    return Number(v.numerator) / d;
  }
  if (Array.isArray(v) && v.length >= 2) {
    const den = Number(v[1]);
    if (!den) return NaN;
    return Number(v[0]) / den;
  }
  return NaN;
}

function normalizeGpsRef(ref, axis) {
  if (ref == null || ref === "") return axis === "lat" ? "N" : "E";
  if (typeof ref === "number" && ref >= 32 && ref <= 127) return String.fromCharCode(ref);
  const s = String(ref).trim().toUpperCase().slice(0, 1);
  if (s === "N" || s === "S" || s === "E" || s === "W") return s;
  return axis === "lat" ? "N" : "E";
}

function dmsToDecimal(dms, ref, axis) {
  if (!dms || dms.length !== 3) return null;
  const deg = exifRationalToNumber(dms[0]);
  const min = exifRationalToNumber(dms[1]);
  const sec = exifRationalToNumber(dms[2]);
  if (![deg, min, sec].every((x) => Number.isFinite(x))) return null;
  const r = normalizeGpsRef(ref, axis || "lat");
  const sign = r === "S" || r === "W" ? -1 : 1;
  const value = sign * (deg + min / 60 + sec / 3600);
  return Number.isFinite(value) ? value : null;
}

function coordsFromExifJsTags(tags) {
  if (!tags || typeof tags !== "object") return null;
  const lat = tags.GPSLatitude;
  const lng = tags.GPSLongitude;
  const latRef = tags.GPSLatitudeRef;
  const lngRef = tags.GPSLongitudeRef;
  if (!lat || !lng) return null;
  const plat = dmsToDecimal(lat, latRef, "lat");
  const plng = dmsToDecimal(lng, lngRef, "lng");
  if (plat != null && plng != null) return { lat: plat, lng: plng };
  return null;
}

function exifToFinite(v) {
  if (v == null) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string") {
    const n = Number.parseFloat(v);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function pickGpsFromExifrParsed(parsed) {
  if (!parsed || typeof parsed !== "object") return null;
  let plat = exifToFinite(parsed.latitude ?? parsed.lat);
  let plng = exifToFinite(parsed.longitude ?? parsed.lng);
  if (plat == null || plng == null) {
    const g = parsed.gps;
    if (g && typeof g === "object") {
      plat = exifToFinite(plat ?? g.latitude ?? g.lat);
      plng = exifToFinite(plng ?? g.longitude ?? g.lng);
    }
  }
  if (plat == null || plng == null) {
    const rawLat = parsed.GPSLatitude;
    const rawLng = parsed.GPSLongitude;
    if (rawLat && rawLng) {
      plat = dmsToDecimal(rawLat, parsed.GPSLatitudeRef, "lat");
      plng = dmsToDecimal(rawLng, parsed.GPSLongitudeRef, "lng");
    }
  }
  if (plat != null && plng != null) return { lat: plat, lng: plng };
  return null;
}

function parseExifDateString(s) {
  if (!s || typeof s !== "string") return null;
  const m = s
    .trim()
    .match(/^(\d{4}):(\d{2}):(\d{2})\s+(\d{2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  const d = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  return Number.isNaN(d.getTime()) ? null : d.toISOString();
}

function exifDateFromExifrParsed(parsed) {
  if (!parsed || typeof parsed !== "object") return null;
  const candidates = [
    parsed.DateTimeOriginal,
    parsed.CreateDate,
    parsed.ModifyDate,
    parsed.DateTime,
  ];
  for (const c of candidates) {
    if (c instanceof Date && !Number.isNaN(c.getTime())) return c.toISOString();
    if (typeof c === "string") {
      const fromExif = parseExifDateString(c);
      if (fromExif) return fromExif;
      const d = new Date(c);
      if (!Number.isNaN(d.getTime())) return d.toISOString();
    }
  }
  return null;
}

function getExifrNamespace() {
  const root = window.exifr;
  if (!root) return null;
  if (typeof root.parse === "function" || typeof root.gps === "function") return root;
  const d = root.default;
  if (d && (typeof d.parse === "function" || typeof d.gps === "function")) return d;
  return null;
}

// Shared GPS extraction (used by both inspection and occupancy) to ensure consistent behavior.
async function readGpsFromFile(file) {
  const exifrApi = getExifrNamespace();

  const tryExifr = async (input) => {
    if (!exifrApi) return null;
    try {
      if (typeof exifrApi.gps === "function") {
        const gps = await exifrApi.gps(input);
        const lat = exifToFinite(gps?.latitude ?? gps?.lat);
        const lng = exifToFinite(gps?.longitude ?? gps?.lng ?? gps?.lon);
        if (lat != null && lng != null) return { lat, lng };
      }
      if (typeof exifrApi.parse === "function") {
        let parsed = await exifrApi.parse(input);
        let found = pickGpsFromExifrParsed(parsed);
        if (found) return found;
        parsed = await exifrApi.parse(input, { tiff: true, ifd0: true, exif: true, gps: true, mergeOutput: true });
        found = pickGpsFromExifrParsed(parsed);
        if (found) return found;
      }
    } catch (e) {
      console.warn("exifr read attempt failed:", e);
    }
    return null;
  };

  if (exifrApi) {
    try {
      if (typeof file?.arrayBuffer === "function") {
        const buf = await file.arrayBuffer();
        const result = await tryExifr(buf);
        if (result) return result;
      }
    } catch (e) {
      console.warn("exifr arrayBuffer read failed:", e);
    }

    try {
      const result = await tryExifr(file);
      if (result) return result;
    } catch (e) {
      console.warn("exifr file read failed:", e);
    }
  }

  if (window.EXIF) {
    try {
      if (typeof window.EXIF.readFromBinaryFile === "function" && typeof file?.arrayBuffer === "function") {
        const buf = await file.arrayBuffer();
        const tags = window.EXIF.readFromBinaryFile(buf);
        const fromBinary = coordsFromExifJsTags(tags);
        if (fromBinary) return fromBinary;
      }
    } catch (e) {
      console.warn("exif-js readFromBinaryFile failed:", e);
    }

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
            done(coordsFromExifJsTags(this.exifdata || {}));
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

/**
 * GPS, capture time, and whether EXIF-like metadata was found (for preview UI).
 */
async function readPhotoExifMetadata(file) {
  let buf = null;
  try {
    if (typeof file?.arrayBuffer === "function") buf = await file.arrayBuffer();
  } catch {
    return { gps: null, takenAt: null, hasExif: false };
  }

  const exifrApi = getExifrNamespace();
  let parsed = null;
  if (exifrApi?.parse && buf) {
    try {
      parsed = await exifrApi.parse(buf, {
        tiff: true,
        ifd0: true,
        exif: true,
        gps: true,
        mergeOutput: true,
      });
    } catch (e) {
      console.warn("exifr metadata parse failed:", e);
    }
  }

  let tagsJs = null;
  if (window.EXIF?.readFromBinaryFile && buf) {
    try {
      tagsJs = window.EXIF.readFromBinaryFile(buf);
    } catch {
      /* non-JPEG or unreadable */
    }
  }

  const hasParsed =
    parsed && typeof parsed === "object" && Object.keys(parsed).length > 0;
  const hasTagsJs =
    tagsJs && typeof tagsJs === "object" && Object.keys(tagsJs).length > 0;

  let takenAt = exifDateFromExifrParsed(parsed);
  if (!takenAt && hasTagsJs) {
    const ds = tagsJs.DateTimeOriginal || tagsJs.DateTime;
    if (typeof ds === "string") takenAt = parseExifDateString(ds);
  }

  let gps = pickGpsFromExifrParsed(parsed);
  if (!gps && hasTagsJs) gps = coordsFromExifJsTags(tagsJs);
  if (!gps) gps = await readGpsFromFile(file);

  const hasExif = Boolean(hasParsed || hasTagsJs || gps || takenAt);
  return { gps, takenAt: takenAt || null, hasExif };
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("FileReader failed"));
    r.readAsDataURL(file);
  });
}

function photoPreviewSetOpen(open) {
  const overlay = document.getElementById("photo-preview-modal-overlay");
  if (!overlay) return;
  overlay.classList.toggle("open", open);
  overlay.setAttribute("aria-hidden", open ? "false" : "true");
}

function photoPreviewCloseOnOverlay(e) {
  if (e?.target?.id === "photo-preview-modal-overlay") photoPreviewCancel();
}

function photoPreviewCancel() {
  const ctx = photoPreviewContext;
  photoPreviewSetOpen(false);
  photoPreviewContext = null;
  if (!ctx) return;
  if (ctx.inputCamera) ctx.inputCamera.value = "";
  if (ctx.inputLibrary) ctx.inputLibrary.value = "";
  const ind = ctx.indicator;
  if (ind) {
    ind.className = "photo-attach-indicator";
    ind.textContent = "";
  }
}

function photoPreviewRenderExif(meta) {
  const panel = document.getElementById("photo-preview-modal-exif");
  if (!panel) return;
  panel.removeAttribute("hidden");
  const has = meta?.hasExif;
  const gps = meta?.gps;
  const taken = meta?.takenAt;
  const takenLabel = taken ? toFriendlyDate(taken) : null;

  const rows = [];
  rows.push(
    `<dt>EXIF metadata</dt><dd>${
      has
        ? "This image contains EXIF metadata."
        : "No EXIF metadata detected (it may have been removed)."
    }</dd>`
  );
  if (gps && Number.isFinite(gps.lat) && Number.isFinite(gps.lng)) {
    rows.push(
      `<dt>Location (from photo)</dt><dd>${gps.lat.toFixed(6)}, ${gps.lng.toFixed(
        6
      )}</dd>`
    );
  } else {
    rows.push(
      `<dt>Location (from photo)</dt><dd>Not embedded. After you confirm, the app may use your device location if allowed.</dd>`
    );
  }
  rows.push(
    `<dt>Date / time taken</dt><dd>${takenLabel || "Not found in EXIF."}</dd>`
  );

  panel.innerHTML = `<dl>${rows.join("")}</dl>`;
}

async function beginPhotoAttachFromPicker(context, file, sourceInput, options = {}) {
  const isInspection = context === "inspection";
  const inputCamera = document.getElementById(
    isInspection ? "inspection_photo" : "occupancy_photo"
  );
  const inputLibrary = document.getElementById(
    isInspection ? "inspection_photo_library" : "occupancy_photo_library"
  );
  const indicator = document.getElementById(
    isInspection ? "inspection-photo-indicator" : "occupancy-photo-indicator"
  );

  if (sourceInput === inputCamera && inputLibrary) inputLibrary.value = "";
  if (sourceInput === inputLibrary && inputCamera) inputCamera.value = "";
  if (sourceInput !== inputLibrary && inputLibrary) inputLibrary.value = "";
  if (sourceInput !== inputCamera && inputCamera) inputCamera.value = "";

  photoPreviewContext = {
    context,
    file,
    sourceInput,
    inputCamera,
    inputLibrary,
    indicator,
    options,
  };

  const img = document.getElementById("photo-preview-modal-img");
  const statusEl = document.getElementById("photo-preview-modal-status");
  const panel = document.getElementById("photo-preview-modal-exif");
  const btnConfirm = document.getElementById("photo-preview-btn-confirm");
  if (btnConfirm) btnConfirm.disabled = true;
  if (statusEl) statusEl.textContent = "Reading photo…";
  if (panel) {
    panel.innerHTML = "";
    panel.setAttribute("hidden", "");
  }
  if (img) {
    img.removeAttribute("src");
    img.alt = "Selected photo preview";
  }

  photoPreviewSetOpen(true);

  let dataUrl = null;
  try {
    dataUrl = await readFileAsDataUrl(file);
  } catch (e) {
    console.warn("photo preview data URL failed:", e);
  }
  if (typeof dataUrl === "string" && img) img.src = dataUrl;
  photoPreviewContext.dataUrl = dataUrl;

  let meta = { gps: null, takenAt: null, hasExif: false };
  try {
    meta = await readPhotoExifMetadata(file);
  } catch (e) {
    console.warn("readPhotoExifMetadata:", e);
  }
  photoPreviewContext.meta = meta;

  photoPreviewRenderExif(meta);
  if (statusEl) {
    statusEl.textContent =
      "Review the preview and EXIF details below, then confirm to attach.";
  }
  if (btnConfirm) btnConfirm.disabled = false;
}

async function photoPreviewConfirm() {
  const ctx = photoPreviewContext;
  if (!ctx?.file) return;

  const {
    context,
    file,
    dataUrl,
    meta,
    indicator,
    options,
  } = ctx;

  photoPreviewSetOpen(false);
  photoPreviewContext = null;

  if (options && typeof options.onConfirm === "function") {
    return options.onConfirm({ file, dataUrl, meta });
  }

  const applyExif = async () => {
    const isInspection = context === "inspection";
    const exLat = normalizeGeoNumber(meta?.gps?.lat);
    const exLng = normalizeGeoNumber(meta?.gps?.lng);
    if (isInspection) {
      currentExifLat = exLat;
      currentExifLng = exLng;
      currentExifTakenAt = meta?.takenAt ?? null;
      currentExifPreviewUrl = typeof dataUrl === "string" ? dataUrl : null;
      currentExifFile = file;
    } else {
      occupancyExifLat = exLat;
      occupancyExifLng = exLng;
      occupancyExifTakenAt = meta?.takenAt ?? null;
      occupancyExifPreviewUrl = typeof dataUrl === "string" ? dataUrl : null;
      occupancyExifFile = file;
    }

    if (indicator) {
      const hasGps = exLat != null && exLng != null;
      indicator.className = "photo-attach-indicator";
      indicator.classList.add(hasGps ? "is-attached" : "is-missing-gps");
      let t = `Photo attached: ${file.name || "image"}`;
      if (hasGps) t += ` (GPS: ${exLat.toFixed(5)}, ${exLng.toFixed(5)})`;
      else t += " (no GPS in file)";
      indicator.textContent = t;
    }

    // Removed automatic device geolocation fallback for photos without GPS
    // as per user request: "do not put lat long if photo do not have lat long"
  };

  if (context === "inspection") {
    currentExifProcessingPromise = applyExif();
    void currentExifProcessingPromise;
  } else {
    occupancyExifProcessingPromise = applyExif();
    void occupancyExifProcessingPromise;
  }

  const toastId = context === "inspection" ? "inspection-toast" : "occupancy-toast";
  logbookShowToast(toastId, "Photo confirmed and attached.");
}

function initPhotoPreviewModal() {
  document
    .getElementById("photo-preview-btn-cancel")
    ?.addEventListener("click", () => photoPreviewCancel());
  document
    .getElementById("photo-preview-btn-close")
    ?.addEventListener("click", () => photoPreviewCancel());
  document
    .getElementById("photo-preview-btn-confirm")
    ?.addEventListener("click", () => void photoPreviewConfirm());
}

function initOccupancyPhotoExif() {
  const inputCamera = document.getElementById("occupancy_photo");
  const inputLibrary = document.getElementById("occupancy_photo_library");
  if (!inputCamera && !inputLibrary) return;
  const indicator = document.getElementById("occupancy-photo-indicator");

  function onPhotoChange(sourceInput) {
    const file = sourceInput?.files?.[0];
    if (!file) return;

    occupancyExifLat = null;
    occupancyExifLng = null;
    occupancyExifPreviewUrl = null;
    occupancyExifTakenAt = null;
    occupancyExifFile = null;
    occupancyExifProcessingPromise = null;

    if (indicator) {
      indicator.className = "photo-attach-indicator";
      indicator.textContent = "Review preview…";
    }

    void beginPhotoAttachFromPicker("occupancy", file, sourceInput);
  }

  inputCamera?.addEventListener("change", () => onPhotoChange(inputCamera));
  inputLibrary?.addEventListener("change", () => onPhotoChange(inputLibrary));
}

function getGoogleDriveThumbnailUrl(url) {
  if (!url || typeof url !== 'string') return url;
  let match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w500-h500`;
  }
  match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w500-h500`;
  }
  return url;
}

function addInspectionMarkerFromEntry(entry) {
  if (!mapInstance || !inspectionMarkersLayer) return;
  if (!Number.isFinite(entry.lat) || !Number.isFinite(entry.lng)) return;

  let icon;
  if (entry.photo_url) {
    const thumbUrl = getGoogleDriveThumbnailUrl(entry.photo_url);
    icon = L.divIcon({
      className: "inspection-marker inspection-marker--photo",
      html: `<div class="inspection-marker-thumb" style="background-image:url('${thumbUrl}')"></div>`,
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
    const thumbUrl = getGoogleDriveThumbnailUrl(entry.photo_url);
    icon = L.divIcon({
      className: "occupancy-marker occupancy-marker--photo",
      html: `<div class="occupancy-marker-thumb" style="background-image:url('${thumbUrl}')"></div>`,
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

  const tooltipText = entry.owner_name || entry.io_number || "Occupancy";
  marker.bindTooltip(tooltipText, {
    permanent: false,
    direction: "top",
    offset: [0, -36],
    opacity: 0.95,
    className: "occupancy-marker-tooltip",
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

  marker.bindPopup(`<strong>${logbookEsc(title)}</strong><br>${logbookEsc(entry.io_number || "")}`);
  marker.on("click", () => {
    openOccupancyDetailPanel(entry);
  });
}

function renderOccupancyMarkersBatched() {
  if (!mapInstance || !occupancyMarkersLayer || !Array.isArray(occupancyData)) return;
  occupancyMarkersLayer.clearLayers();
  
  const isTypeFilter = mapMarkerFilter !== "all" && mapMarkerFilter !== "businesses" && mapMarkerFilter !== "occupancies";
  const showOccupancy = mapMarkerFilter === "all" || mapMarkerFilter === "occupancies" || isTypeFilter;

  if (!showOccupancy) return;

  const withCoords = occupancyData.filter((row) => {
    if (row.lat == null || row.lng == null) return false;
    if (isTypeFilter && row.type_of_occupancy !== mapMarkerFilter) return false;
    return true;
  });
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

  const showInspection = mapMarkerFilter === "all" || mapMarkerFilter === "businesses";
  if (!showInspection) return;

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
  const isTypeFilter = mapMarkerFilter !== "all" && mapMarkerFilter !== "businesses" && mapMarkerFilter !== "occupancies";
  const includeInspection = mapMarkerFilter === "all" || mapMarkerFilter === "businesses";
  const includeOccupancy = mapMarkerFilter === "all" || mapMarkerFilter === "occupancies" || isTypeFilter;

  const candidates = [
    ...(includeInspection ? getMarkedInspectionEntries().map((r) => ({ type: "inspection", r })) : []),
    ...(includeOccupancy ? getMarkedOccupancyEntries().filter(r => !isTypeFilter || r.type_of_occupancy === mapMarkerFilter).map((r) => ({ type: "occupancy", r })) : []),
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
  const result = await gasRequest("read", { table: "inspection_logbook" });
  const rows = result.data || [];
  inspectionData = rows.map((r) => ({
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
    fsic_purpose: r.fsic_purpose ?? null,
    fsic_permit_type: r.fsic_permit_type ?? null,
    fsic_valid_for: r.fsic_valid_for ?? null,
    fsic_valid_until: r.fsic_valid_until ?? null,
    fsic_fee_amount: r.fsic_fee_amount ?? null,
    fsic_fee_or_number: r.fsic_fee_or_number ?? null,
    fsic_fee_date: r.fsic_fee_date ?? null,
    lat: r.latitude ?? null,
    lng: r.longitude ?? null,
    photo_url: r.photo_url ?? null,
    photo_taken_at: r.photo_taken_at ?? null,
    created_at: r.created_at,
  }));
}

function inspectionInitData() {
  if (inspectionDataLoaded) return;
  inspectionDataLoaded = true;

  // Clear any stale cache from the old Supabase backend
  localStorage.removeItem(INSPECTION_STORAGE_KEY);
  inspectionSetPrintDate();
  inspectionRenderTable();
  renderInspectionMarkersBatched();
  setInspectionTab(inspectionActiveTab);

  if (!isGasEnabled()) return;

  (async () => {
    try {
      await inspectionLoadFromSupabase();
      inspectionSetPrintDate();
      inspectionRenderTable();
      renderInspectionMarkersBatched();
      setInspectionTab(inspectionActiveTab);
    } catch (err) {
      console.warn("Inspection load from GAS failed:", err);
      logbookShowToast("inspection-toast", "Could not load data from server.");
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
  const countBadge = document.getElementById("fsec-record-count");
  if (!tbody || !empty) return;
  if (countBadge) countBadge.textContent = String(fsecData.length || 0);

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

async function fsecEditEntry(idx) {
  const oldRow = fsecData[idx];
  if (!oldRow) return;

  if (isGasEnabled()) {
    logbookShowToast("fsec-toast", "Refreshing record data...");
    try {
      await fsecLoadFromSupabase();
    } catch (err) {
      console.warn("Refresh failed:", err);
    }
  }

  const row = (oldRow.id) ? fsecData.find(r => r.id === oldRow.id) : fsecData[idx];
  if (!row) return;

  fsecEditingIdx = fsecData.indexOf(row);
  fsecEditingId = row.id || null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };

  setVal("fsec_owner", row.fsec_owner);
  setVal("proposed_project", row.proposed_project);
  setVal("fsec_date", logbookFormatDateForInput(row.fsec_date));
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

  // Reset to first step
  updateModalStepUI('fsec', 1);
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
      await gasRequest("delete", { table: "fsec_building_plan_logbook", id: row.id });
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

  // Reset to first step
  updateModalStepUI('fsec', 1);
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
    line,
    barangay ? `Barangay ${barangay}` : null,
    municipal,
    province,
    `Region ${region}`
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
      
      if (fsecEditingId) {
        await gasRequest("update", { table: "fsec_building_plan_logbook", id: fsecEditingId, row: payload });
      } else {
        await gasRequest("insert", { table: "fsec_building_plan_logbook", row: payload });
      }
      
      // Database write completes in the background; UI was already updated optimistically
      logbookShowToast("fsec-toast", "Saved to database.");
    } catch (err) {
      const msg = err?.message || String(err);
      logbookShowToast("fsec-toast", "Save failed: " + msg);
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
  const result = await gasRequest("read", { table: "fsec_building_plan_logbook" });
  fsecData = (result.data || []).map((r) => ({
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
  localStorage.removeItem("bfp_fsec");
  fsecSetPrintDate();
  fsecRenderTable();
  if (!isGasEnabled()) return;
  try {
    await fsecLoadFromSupabase();
    fsecSetPrintDate();
    fsecRenderTable();
  } catch (err) {
    console.warn("FSEC load from GAS failed:", err);
    logbookShowToast("fsec-toast", "Could not load data from server.");
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
  const countBadge = document.getElementById("conveyance-record-count");
  if (!tbody || !empty) return;
  if (countBadge) countBadge.textContent = String(conveyanceData.length || 0);

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

async function conveyanceEditEntry(idx) {
  const oldRow = conveyanceData[idx];
  if (!oldRow) return;

  if (isGasEnabled()) {
    logbookShowToast("conveyance-toast", "Refreshing record data...");
    try {
      await conveyanceLoadFromSupabase();
    } catch (err) {
      console.warn("Refresh failed:", err);
    }
  }

  const row = (oldRow.id) ? conveyanceData.find(r => r.id === oldRow.id) : conveyanceData[idx];
  if (!row) return;

  conveyanceEditingIdx = conveyanceData.indexOf(row);
  conveyanceEditingId = row.id || null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };
  setVal("conveyance_date", logbookFormatDateForInput(row.log_date));
  setVal("conveyance_io_number", row.io_number);
  setVal("conveyance_owner_name", row.owner_name);
  ensureSelectOption("conveyance_inspected_by", row.inspectors || "");
  setVal("conveyance_inspector_position", row.inspector_position);
  ensureSelectOption("conveyance_included_personnel_name", row.included_personnel_name || "");
  setVal("conveyance_included_personnel_position", row.included_personnel_position);
  setVal("conveyance_remarks_signature", row.remarks_signature);

  setText("conveyance-modal-title", "Edit Conveyance Record");
  const btn = document.getElementById("conveyance-btn-save");
  if (btn) btn.textContent = "Update Record";

  const overlay = document.getElementById("conveyance-modal-overlay");
  overlay?.classList.add("open");

  // Reset to first step
  updateModalStepUI('conveyance', 1);
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
  
  const inspBy = document.getElementById("conveyance_inspected_by");
  if (inspBy) inspBy.selectedIndex = 0;
  const inspPos = document.getElementById("conveyance_inspector_position");
  if (inspPos) inspPos.value = "";
  const incBy = document.getElementById("conveyance_included_personnel_name");
  if (incBy) incBy.selectedIndex = 0;
  const incPos = document.getElementById("conveyance_included_personnel_position");
  if (incPos) incPos.value = "";
  
  const rem = document.getElementById("conveyance_remarks_signature");
  if (rem) rem.value = "";

  const overlay = document.getElementById("conveyance-modal-overlay");
  overlay?.classList.add("open");

  // Reset to first step
  updateModalStepUI('conveyance', 1);
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
      await gasRequest("delete", { table: "conveyance_logbook", id: row.id });
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
    inspectors: (document.getElementById("conveyance_inspected_by") || { value: "" }).value.trim(),
    inspector_position: (document.getElementById("conveyance_inspector_position") || { value: "" }).value.trim(),
    included_personnel_name: (document.getElementById("conveyance_included_personnel_name") || { value: "" }).value.trim(),
    included_personnel_position: (document.getElementById("conveyance_included_personnel_position") || { value: "" }).value.trim(),
    remarks_signature: (document.getElementById("conveyance_remarks_signature") || { value: "" }).value.trim(),
    created_at: new Date().toISOString(),
  };

  const isIncludedPersonnelPlaceholder =
    !entry.included_personnel_name || /^select/i.test(String(entry.included_personnel_name).trim());
  if (isIncludedPersonnelPlaceholder) {
    entry.included_personnel_name = "";
    entry.included_personnel_position = "";
  }

  const isPlaceholderInspector =
    !entry.inspectors || /^select\s+inspector$/i.test(String(entry.inspectors).trim());
  if (isPlaceholderInspector) entry.inspectors = "";

  if (!entry.log_date || !entry.io_number || !entry.inspectors) {
    logbookShowToast("conveyance-toast", "⚠️ Please fill in Date, IO Number, and Inspected By.");
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
        inspector_position: entry.inspector_position || null,
        included_personnel_name: entry.included_personnel_name || null,
        included_personnel_position: entry.included_personnel_position || null,
        remarks_signature: entry.remarks_signature,
      };
      // On UPDATE, don't overwrite optional owner_name with null/blank.
      if (conveyanceEditingId && (!payload.owner_name || String(payload.owner_name).trim() === "")) {
        delete payload.owner_name;
      }
      if (conveyanceEditingId) {
        await gasRequest("update", { table: "conveyance_logbook", id: conveyanceEditingId, row: payload });
      } else {
        await gasRequest("insert", { table: "conveyance_logbook", row: payload });
      }
      logbookShowToast("conveyance-toast", "Saved to database.");
    } catch (err) {
      const msg = err?.message || String(err);
      logbookShowToast("conveyance-toast", "Save failed: " + msg);
    }
  })();
}

async function conveyanceLoadFromSupabase() {
  const result = await gasRequest("read", { table: "conveyance_logbook" });
  conveyanceData = (result.data || []).map((r) => ({
    id: r.id,
    log_date: r.log_date,
    io_number: r.io_number,
    owner_name: r.owner_name || "",
    inspectors: r.inspectors || "",
    inspector_position: r.inspector_position || "",
    included_personnel_name: r.included_personnel_name || "",
    included_personnel_position: r.included_personnel_position || "",
    remarks_signature: r.remarks_signature,
    created_at: r.created_at,
  }));
  conveyanceSaveToLocal();
}

async function conveyanceInitData() {
  if (conveyanceDataLoaded) return;
  conveyanceDataLoaded = true;
  localStorage.removeItem(CONVEYANCE_STORAGE_KEY);
  conveyanceRenderTable();
  if (!isGasEnabled()) return;
  try {
    await conveyanceLoadFromSupabase();
    conveyanceRenderTable();
  } catch (err) {
    console.warn("Conveyance load from GAS failed:", err);
    logbookShowToast("conveyance-toast", "Could not load data from server.");
    conveyanceRenderTable();
  }
}

// -----------------------------
// Fire Drill logbook module (fields align with fire_drill_certificate.html placeholders)
// -----------------------------

/** Build full address string and parts (same pattern as inspection / occupancy). */
function fireDrillMergeAddressFields() {
  const region = (document.getElementById("fire_drill_addr_region")?.value || "X").trim();
  const province = (document.getElementById("fire_drill_addr_province")?.value || "Bukidnon").trim();
  const municipal = (document.getElementById("fire_drill_addr_municipal")?.value || "Manolo Fortich").trim();
  let barangay = (document.getElementById("fire_drill_addr_barangay") || { value: "" }).value.trim();
  const line = (document.getElementById("fire_drill_addr_line") || { value: "" }).value.trim();
  if (!barangay || /^select/i.test(barangay)) barangay = "";
  const merged = [line, barangay ? `Barangay ${barangay}` : null, municipal, province, `Region ${region}`]
    .filter((p) => String(p || "").trim())
    .join(", ");
  return { merged, addr_barangay: barangay, addr_line: line };
}

/** Split stored `address` from the sheet into line + barangay when possible. */
function fireDrillParseStoredAddress(full) {
  const s = (full || "").toString().trim();
  if (!s) return { addr_line: "", addr_barangay: "" };
  const brgyMatch = s.match(/Barangay\s+([^,]+)/i);
  const addr_barangay = brgyMatch ? brgyMatch[1].trim() : "";
  let addr_line = "";
  if (brgyMatch) {
    addr_line = s.slice(0, brgyMatch.index).replace(/,\s*$/, "").trim();
  } else {
    addr_line = s.split(",")[0]?.trim() || s;
  }
  return { addr_line, addr_barangay };
}

/** Ordinal day for certificates (1st, 2nd, … 14th). */
function fireDrillOrdinalDay(dayNum) {
  const d = Math.floor(Number(dayNum));
  if (!Number.isFinite(d) || d < 1 || d > 31) return String(dayNum);
  const j = d % 10;
  const k = d % 100;
  if (k >= 11 && k <= 13) return `${d}th`;
  if (j === 1) return `${d}st`;
  if (j === 2) return `${d}nd`;
  if (j === 3) return `${d}rd`;
  return `${d}th`;
}

function fireDrillFormatMonthYearIssued(d) {
  return d.toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

function fireDrillAddMonthsClamped(dateObj, monthsToAdd) {
  const y = dateObj.getFullYear();
  const m = dateObj.getMonth();
  const day = dateObj.getDate();
  const targetMonthIndex = m + monthsToAdd;
  const lastDay = new Date(y, targetMonthIndex + 1, 0).getDate();
  const safeDay = Math.min(day, lastDay);
  return new Date(y, targetMonthIndex, safeDay, 12, 0, 0);
}

/** Infer validity type from issuance date + valid-until date when possible. */
function fireDrillInferValidityType(issuedIso, validIso) {
  if (!issuedIso || !validIso) return "";
  const issued = new Date(issuedIso + "T12:00:00");
  const valid = new Date(validIso + "T12:00:00");
  if (Number.isNaN(issued.getTime()) || Number.isNaN(valid.getTime())) return "";

  const sameDate = (a, b) =>
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();

  if (sameDate(fireDrillAddMonthsClamped(issued, 3), valid)) return "quarterly";
  if (sameDate(fireDrillAddMonthsClamped(issued, 6), valid)) return "midyear";
  if (sameDate(fireDrillAddMonthsClamped(issued, 12), valid)) return "annual";
  return "";
}

/** Auto-compute "valid until" from issuance date + selected validity period. */
function fireDrillSyncValidityDate() {
  const issuedEl = document.getElementById("fire_drill_date_issued");
  const typeEl = document.getElementById("fire_drill_validity_type");
  const validEl = document.getElementById("fire_drill_date_valid");
  if (!issuedEl || !typeEl || !validEl) return;
  const issuedVal = issuedEl.value;
  const type = (typeEl.value || "").trim();
  if (!issuedVal || !type) return;

  const issued = new Date(issuedVal + "T12:00:00");
  if (Number.isNaN(issued.getTime())) return;

  let months = 0;
  if (type === "quarterly") months = 3;
  else if (type === "midyear") months = 6;
  else if (type === "annual") months = 12;
  else return;

  const out = fireDrillAddMonthsClamped(issued, months);
  const y = out.getFullYear();
  const m = String(out.getMonth() + 1).padStart(2, "0");
  const d = String(out.getDate()).padStart(2, "0");
  validEl.value = `${y}-${m}-${d}`;
}

/** Fill day + month/year from the issuance date picker (local noon to avoid TZ shift). */
function fireDrillSyncIssuanceFieldsFromDate() {
  const el = document.getElementById("fire_drill_date_issued");
  const dayEl = document.getElementById("fire_drill_day_issued");
  const myEl = document.getElementById("fire_drill_month_year_issued");
  if (!el || !dayEl || !myEl) return;
  const v = el.value;
  if (!v) {
    dayEl.value = "";
    myEl.value = "";
    return;
  }
  const d = new Date(v + "T12:00:00");
  if (Number.isNaN(d.getTime())) return;
  dayEl.value = fireDrillOrdinalDay(d.getDate());
  myEl.value = fireDrillFormatMonthYearIssued(d);
  fireDrillSyncValidityDate();
}

/**
 * Rebuild yyyy-mm-dd for the issuance date input from stored day/month-year strings.
 */
function fireDrillIssuancePartsToDateString(dayStr, monthYearStr) {
  const dayNum = parseInt(String(dayStr || "").replace(/[^\d]/g, ""), 10);
  const my = String(monthYearStr || "").trim();
  const parts = my.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (!parts || !dayNum) return "";
  const monthName = parts[1];
  const year = parseInt(parts[2], 10);
  const months = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
  ];
  const mi = months.indexOf(monthName.toLowerCase());
  if (mi < 0) return "";
  const d = new Date(year, mi, dayNum);
  if (d.getFullYear() !== year || d.getMonth() !== mi || d.getDate() !== dayNum) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
}

const FIRE_DRILL_STORAGE_KEY = "bfp_fire_drill";
let fireDrillData = [];
let fireDrillEditingIdx = null;
let fireDrillEditingId = null;
let fireDrillDataLoaded = false;

function fireDrillLoadFromLocal() {
  fireDrillData = JSON.parse(localStorage.getItem(FIRE_DRILL_STORAGE_KEY) || "[]");
}

function fireDrillSaveToLocal() {
  localStorage.setItem(FIRE_DRILL_STORAGE_KEY, JSON.stringify(fireDrillData));
}

function fireDrillSetPrintDate() {
  const el = document.getElementById("fire_drill-print-date");
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleDateString("en-PH", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

function fireDrillPrintPanel() {
  fireDrillSetPrintDate();
  const oldTitle = document.title;
  document.title = "";
  setTimeout(() => {
    window.print();
    setTimeout(() => {
      document.title = oldTitle;
    }, 500);
  }, 0);
}

function fireDrillClearFilters() {
  const q = document.getElementById("fire_drill-filter-q");
  const from = document.getElementById("fire_drill-filter-from");
  const to = document.getElementById("fire_drill-filter-to");
  if (q) q.value = "";
  if (from) from.value = "";
  if (to) to.value = "";
  fireDrillRenderTable();
}

function fireDrillRenderTable() {
  const tbody = document.getElementById("tbody-fire_drill");
  const empty = document.getElementById("empty-fire_drill");
  const tableWrap = document.getElementById("table-fire_drill")?.closest(".table-wrap");
  const countBadge = document.getElementById("fire_drill-record-count");
  if (!tbody || !empty) return;
  if (countBadge) countBadge.textContent = String(fireDrillData.length || 0);

  const q = normalizeQuery(document.getElementById("fire_drill-filter-q")?.value);
  const from = (document.getElementById("fire_drill-filter-from")?.value || "").trim();
  const to = (document.getElementById("fire_drill-filter-to")?.value || "").trim();

  const filtered = fireDrillData
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (from || to) {
        if (!inDateRange(row.certificate_date, from, to)) return false;
      }
      if (!q) return true;
      const hay = normalizeQuery(
        [
          row.control_number,
          row.building_name,
          row.address,
          row.or_number,
          row.amount_paid,
        ]
          .filter(Boolean)
          .join(" | ")
      );
      return hay.includes(q);
    });

  tbody.innerHTML = "";
  if (fireDrillData.length === 0) {
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
      <td class="td-date" data-label="Cert. date">${logbookFormatDate(row.certificate_date)}</td>
      <td data-label="Control No.">${logbookEsc(row.control_number)}</td>
      <td data-label="Building">${logbookEsc(row.building_name)}</td>
      <td data-label="Address"><div class="cell-pre">${logbookEsc(row.address)}</div></td>
      <td data-label="Valid until">${logbookFormatDate(row.date_valid)}</td>
      <td class="col-action" data-label="Action">
        <select class="action-select" aria-label="Row actions" onchange="fireDrillHandleAction(this.value, ${idx}); this.selectedIndex = 0;">
          <option value="">Actions…</option>
          <option value="open_certificate">Open Certificate</option>
          <option value="edit">Edit</option>
          <option value="delete">Delete</option>
        </select>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function fireDrillHandleAction(action, idx) {
  if (!action) return;
  if (action === "open_certificate") return fireDrillOpenCertificate(idx);
  if (action === "edit") return fireDrillEditEntry(idx);
  if (action === "delete") return fireDrillDeleteEntry(idx);
}

function fireDrillOpenCertificate(idx) {
  const row = fireDrillData[idx];
  if (!row) return;
  try {
    sessionStorage.setItem("fsis.fire_drill.certificate", JSON.stringify(row));
  } catch (err) {
    console.warn("Could not write fire drill certificate row to sessionStorage:", err);
  }
  window.open("./fire_drill_certificate.html", "_blank");
}

async function fireDrillEditEntry(idx) {
  const oldRow = fireDrillData[idx];
  if (!oldRow) return;

  if (isGasEnabled()) {
    logbookShowToast("fire_drill-toast", "Retrieving the latest record from the server.");
    try {
      await fireDrillLoadFromSupabase();
    } catch (err) {
      console.warn("Refresh failed:", err);
    }
  }

  const row = oldRow.id ? fireDrillData.find((r) => r.id === oldRow.id) : fireDrillData[idx];
  if (!row) return;

  fireDrillEditingIdx = fireDrillData.indexOf(row);
  fireDrillEditingId = row.id || null;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };
  setVal("fire_drill_control_number", row.control_number);
  setVal("fire_drill_certificate_date", logbookFormatDateForInput(row.certificate_date));
  setVal("fire_drill_building_name", row.building_name);
  setVal("fire_drill_owner_name", row.owner_name);
  let brgy = row.addr_barangay || "";
  let line = row.addr_line || "";
  if (!brgy && !line && row.address) {
    const p = fireDrillParseStoredAddress(row.address);
    brgy = p.addr_barangay;
    line = p.addr_line;
  }
  ensureSelectOption("fire_drill_addr_barangay", brgy);
  setVal("fire_drill_addr_line", line);
  let issuedIso = fireDrillIssuancePartsToDateString(row.day_issued, row.month_year_issued);
  if (!issuedIso && row.certificate_date) {
    issuedIso = logbookFormatDateForInput(row.certificate_date);
  }
  setVal("fire_drill_date_issued", issuedIso);
  fireDrillSyncIssuanceFieldsFromDate();
  setVal("fire_drill_date_valid", logbookFormatDateForInput(row.date_valid));
  setVal("fire_drill_validity_type", fireDrillInferValidityType(issuedIso, logbookFormatDateForInput(row.date_valid)));
  setVal("fire_drill_amount_paid", row.amount_paid);
  setVal("fire_drill_or_number", row.or_number);
  setVal("fire_drill_date_paid", logbookFormatDateForInput(row.date_paid));

  setText("fire_drill-modal-title", "Amend Fire Drill Certificate Entry");
  const btn = document.getElementById("fire_drill-btn-save");
  if (btn) btn.textContent = "Update Entry";

  document.getElementById("fire_drill-modal-overlay")?.classList.add("open");
  updateModalStepUI("fire_drill", 1);
}

function fireDrillOpenModal() {
  fireDrillEditingIdx = null;
  fireDrillEditingId = null;

  setText("fire_drill-modal-title", "New Fire Drill Certificate Entry");
  const btn = document.getElementById("fire_drill-btn-save");
  if (btn) btn.textContent = "Submit";

  const d = new Date().toISOString().slice(0, 10);
  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };
  setVal("fire_drill_control_number", "");
  setVal("fire_drill_certificate_date", d);
  setVal("fire_drill_building_name", "");
  setVal("fire_drill_owner_name", "");
  setVal("fire_drill_addr_line", "");
  const fdBr = document.getElementById("fire_drill_addr_barangay");
  if (fdBr) fdBr.selectedIndex = 0;
  setVal("fire_drill_date_issued", d);
  fireDrillSyncIssuanceFieldsFromDate();
  setVal("fire_drill_validity_type", "");
  setVal("fire_drill_date_valid", "");
  setVal("fire_drill_amount_paid", "");
  setVal("fire_drill_or_number", "");
  setVal("fire_drill_date_paid", "");

  document.getElementById("fire_drill-modal-overlay")?.classList.add("open");
  updateModalStepUI("fire_drill", 1);
}

function fireDrillCloseModal() {
  document.getElementById("fire_drill-modal-overlay")?.classList.remove("open");
}

function fireDrillCloseOnOverlay(e) {
  if (e?.target?.id === "fire_drill-modal-overlay") fireDrillCloseModal();
}

// Keep modal callbacks available for inline HTML handlers after deployment.
if (typeof window !== "undefined") {
  window.fireDrillOpenModal = fireDrillOpenModal;
  window.fireDrillCloseModal = fireDrillCloseModal;
  window.fireDrillCloseOnOverlay = fireDrillCloseOnOverlay;
  window.fireDrillSaveEntry = fireDrillSaveEntry;
  window.fireDrillModalStep = fireDrillModalStep;
  window.fireDrillHandleAction = fireDrillHandleAction;
}

function fireDrillDeleteEntry(idx) {
  if (!confirm("Delete this record?")) return;

  const row = fireDrillData[idx];
  if (!row) return;

  fireDrillData.splice(idx, 1);
  fireDrillSaveToLocal();
  fireDrillRenderTable();
  logbookShowToast("fire_drill-toast", "The entry was deleted.");

  if (!isSupabaseEnabled()) return;
  if (!row.id) {
    logbookShowToast("fire_drill-toast", "This entry cannot be deleted: no record identifier.");
    return;
  }

  (async () => {
    try {
      await gasRequest("delete", { table: "fire_drill_logbook", id: row.id });
    } catch (err) {
      logbookShowToast("fire_drill-toast", "Deletion failed: " + (err?.message || err));
    }
  })();
}

function fireDrillSaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();

  const addrParts = fireDrillMergeAddressFields();

  const issuedVal = (document.getElementById("fire_drill_date_issued") || { value: "" }).value;
  let day_issued = "";
  let month_year_issued = "";
  if (issuedVal) {
    const issuedD = new Date(issuedVal + "T12:00:00");
    if (!Number.isNaN(issuedD.getTime())) {
      day_issued = fireDrillOrdinalDay(issuedD.getDate());
      month_year_issued = fireDrillFormatMonthYearIssued(issuedD);
    }
  }

  const entry = {
    control_number: (document.getElementById("fire_drill_control_number") || { value: "" }).value.trim(),
    certificate_date: (document.getElementById("fire_drill_certificate_date") || { value: "" }).value,
    building_name: (document.getElementById("fire_drill_building_name") || { value: "" }).value.trim(),
    owner_name: (document.getElementById("fire_drill_owner_name") || { value: "" }).value.trim(),
    address: addrParts.merged,
    addr_barangay: addrParts.addr_barangay,
    addr_line: addrParts.addr_line,
    day_issued,
    month_year_issued,
    date_valid: (document.getElementById("fire_drill_date_valid") || { value: "" }).value,
    amount_paid: (document.getElementById("fire_drill_amount_paid") || { value: "" }).value.trim(),
    or_number: (document.getElementById("fire_drill_or_number") || { value: "" }).value.trim(),
    date_paid: (document.getElementById("fire_drill_date_paid") || { value: "" }).value,
    created_at: new Date().toISOString(),
  };

  if (!entry.certificate_date || !entry.control_number || !entry.building_name) {
    logbookShowToast(
      "fire_drill-toast",
      "Certificate date, control number, and the name of the building or structure are required."
    );
    return;
  }

  const isOnline = isSupabaseEnabled();

  if (fireDrillEditingIdx !== null) {
    const prev = fireDrillData[fireDrillEditingIdx] || {};
    fireDrillData[fireDrillEditingIdx] = {
      ...prev,
      ...entry,
      id: prev.id || null,
      created_at: prev.created_at || entry.created_at,
    };
  } else {
    fireDrillData.push({ ...entry, id: null });
  }

  fireDrillSaveToLocal();
  fireDrillRenderTable();
  fireDrillCloseModal();
  showSaveIndicator("Fire drill certificate entry saved.");

  if (!isOnline) {
    logbookShowToast(
      "fire_drill-toast",
      "The entry was saved on this device only; the server was not available."
    );
    return;
  }

  const serverEditingId = fireDrillEditingId;
  (async () => {
    try {
      const payload = {
        control_number: entry.control_number,
        certificate_date: entry.certificate_date,
        building_name: entry.building_name,
        owner_name: entry.owner_name || null,
        address: entry.address || null,
        day_issued: entry.day_issued || null,
        month_year_issued: entry.month_year_issued || null,
        date_valid: entry.date_valid || null,
        amount_paid: entry.amount_paid || null,
        or_number: entry.or_number || null,
        date_paid: entry.date_paid || null,
      };
      if (serverEditingId) {
        await gasRequest("update", { table: "fire_drill_logbook", id: serverEditingId, row: payload });
      } else {
        await gasRequest("insert", { table: "fire_drill_logbook", row: payload });
      }
      logbookShowToast("fire_drill-toast", "The entry was saved to the central record.");
      await fireDrillLoadFromSupabase();
      fireDrillRenderTable();
    } catch (err) {
      const msg = err?.message || String(err);
      logbookShowToast("fire_drill-toast", "The entry could not be saved: " + msg);
    }
  })();
}

async function fireDrillLoadFromSupabase() {
  const result = await gasRequest("read", { table: "fire_drill_logbook" });
  fireDrillData = (result.data || []).map((r) => {
    const parsed = fireDrillParseStoredAddress(r.address || "");
    return {
      id: r.id,
      control_number: r.control_number || "",
      certificate_date: r.certificate_date || "",
      building_name: r.building_name || "",
      owner_name: r.owner_name || "",
      address: r.address || "",
      addr_barangay: parsed.addr_barangay,
      addr_line: parsed.addr_line,
      day_issued: r.day_issued || "",
      month_year_issued: r.month_year_issued || "",
      date_valid: r.date_valid || "",
      amount_paid: r.amount_paid != null ? String(r.amount_paid) : "",
      or_number: r.or_number || "",
      date_paid: r.date_paid || "",
      created_at: r.created_at,
    };
  });
  fireDrillSaveToLocal();
}

async function fireDrillInitData() {
  if (fireDrillDataLoaded) return;
  fireDrillDataLoaded = true;
  localStorage.removeItem(FIRE_DRILL_STORAGE_KEY);
  fireDrillSetPrintDate();
  fireDrillRenderTable();
  if (!isGasEnabled()) return;
  try {
    await fireDrillLoadFromSupabase();
    fireDrillSetPrintDate();
    fireDrillRenderTable();
  } catch (err) {
    console.warn("Fire Drill load from GAS failed:", err);
    logbookShowToast("fire_drill-toast", "Unable to load data from the server.");
    fireDrillRenderTable();
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
  const countBadge = document.getElementById("occupancy-record-count");
  if (!tbody || !empty) return;
  if (countBadge) countBadge.textContent = String(occupancyData.length || 0);

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
          row.owner_phone,
          row.business_name,
          row.fsic_number,
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
      <td data-label="IO Number">${logbookEsc(row.io_number)}</td>
      <td data-label="Name of Owner">${logbookEsc(row.owner_name)}</td>
      <td data-label="Owner Phone">${logbookEsc(row.owner_phone)}</td>
      <td data-label="Residential / Property"><strong>${logbookEsc(row.business_name)}</strong></td>
      <td data-label="Address">${logbookEsc(inspectionFormatAddressDisplay(row))}</td>
      <td class="td-date" data-label="Date">${logbookFormatDate(row.log_date)}</td>
      <td data-label="Type">${logbookEsc(row.type_of_occupancy)}</td>
      <td data-label="FSIC Number"><strong>${logbookEsc(row.fsic_number)}</strong></td>
      <td data-label="Inspected By"><div class="cell-pre">${logbookEsc(row.inspectors)}</div></td>
      <td class="col-action" data-label="Action">
        <select class="action-select" aria-label="Row actions" onchange="occupancyHandleAction(this.value, ${idx}); this.selectedIndex = 0;">
          <option value="">Actions…</option>
          <option value="edit">Edit</option>
          <option value="add_photo">Add photo</option>
          <option value="open_io_html">Open IO (HTML)</option>
          <option value="open_clearance_html">Release clearance (FSIC)</option>
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
  if (action === "add_photo") return occupancyAddPhoto(idx);
  if (action === "open_io_html") return occupancyOpenIoHtml(idx);
  if (action === "open_clearance_html") return occupancyClearanceOpenModal(idx);
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

async function occupancyEditEntry(idx) {
  const oldRow = occupancyData[idx];
  if (!oldRow) return;

  if (isGasEnabled()) {
    logbookShowToast("occupancy-toast", "Refreshing record data...");
    try {
      await occupancyLoadFromSupabase();
    } catch (err) {
      console.warn("Refresh failed:", err);
    }
  }

  const row = (oldRow.id) ? occupancyData.find(r => r.id === oldRow.id) : occupancyData[idx];
  if (!row) return;

  occupancyEditingIdx = occupancyData.indexOf(row);
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
  setVal("occupancy_date", logbookFormatDateForInput(row.log_date));
  setVal("occupancy_io_number", row.io_number);
  setVal("occupancy_fsic_number", row.fsic_number);
  setVal("occupancy_owner_name", row.owner_name);
  setVal("occupancy_owner_phone", row.owner_phone);
  setVal("occupancy_property_name", row.business_name);
  
  ensureSelectOption("occupancy_addr_barangay", row.addr_barangay || "");
  setVal("occupancy_addr_line", row.addr_line);
  ensureSelectOption("occupancy_type_of_occupancy", row.type_of_occupancy || "");
  setVal("occupancy_inspected_by", row.inspectors);
  setVal("occupancy_inspector_position", row.inspector_position);
  ensureSelectOption("occupancy_included_personnel_name", row.included_personnel_name || "");
  setVal("occupancy_included_personnel_position", row.included_personnel_position);
  setVal("occupancy_duration_start", logbookFormatDateForInput(row.duration_start));
  setVal("occupancy_duration_end", logbookFormatDateForInput(row.duration_end));

  setText("occupancy-modal-title", "Edit Occupancy Record");
  const btn = document.getElementById("occupancy-btn-save");
  if (btn) btn.textContent = "Update Record";

  const overlay = document.getElementById("occupancy-modal-overlay");
  overlay?.classList.add("open");

  // Reset to first step
  updateModalStepUI('occupancy', 1);
}


let occupancyClearanceIdx = null;

function occupancyClearanceOpenModal(idx) {
  const row = occupancyData[idx];
  if (!row) return;

  occupancyClearanceIdx = idx;

  const setVal = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.value = value || "";
  };

  setVal("occ_clearance_fsic_number", row.fsic_number);
  setVal("occ_clearance_purpose", row.fsic_purpose);
  setVal("occ_clearance_valid_until", row.fsic_valid_until);
  setVal("occ_clearance_fee_amount", row.fsic_fee_amount);
  setVal("occ_clearance_fee_or_number", row.fsic_fee_or_number);
  setVal("occ_clearance_fee_date", row.fsic_fee_date);

  const overlay = document.getElementById("occupancy-clearance-modal-overlay");
  overlay?.classList.add("open");
}

function occupancyClearanceCloseModal() {
  const overlay = document.getElementById("occupancy-clearance-modal-overlay");
  overlay?.classList.remove("open");
}

function occupancyClearanceCloseOnOverlay(e) {
  if (e?.target?.id === "occupancy-clearance-modal-overlay") occupancyClearanceCloseModal();
}

function occupancyClearanceProceed(e) {
  if (e?.preventDefault) e.preventDefault();

  if (occupancyClearanceIdx === null) return;
  const row = occupancyData[occupancyClearanceIdx];
  if (!row) return;

  const val = (id) => (document.getElementById(id) || { value: "" }).value.trim();

  row.fsic_number = val("occ_clearance_fsic_number");
  row.fsic_purpose = val("occ_clearance_purpose");
  row.fsic_valid_until = val("occ_clearance_valid_until");
  row.fsic_fee_amount = val("occ_clearance_fee_amount");
  row.fsic_fee_or_number = val("occ_clearance_fee_or_number");
  row.fsic_fee_date = val("occ_clearance_fee_date");

  occupancySaveToLocal();
  occupancyRenderTable();
  occupancyClearanceCloseModal();

  try {
    sessionStorage.setItem("fsis.clearance.current", JSON.stringify({ ...row, _sourceType: "occupancy" }));
  } catch (err) {
    console.warn("Could not write clearance row to sessionStorage:", err);
  }

  window.open("./fsis_clearance.html", "_blank");
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
  const getEl = (id) => document.getElementById(id);
  const clearVals = ["occupancy_io_number", "occupancy_fsic_number", "occupancy_owner_name", 
    "occupancy_owner_phone", "occupancy_property_name", "occupancy_addr_line", 
    "occupancy_inspector_position", "occupancy_included_personnel_position", 
    "occupancy_duration_start", "occupancy_duration_end"];
  clearVals.forEach(id => {
    const el = getEl(id);
    if (el) el.value = "";
  });
  const clearSelects = ["occupancy_addr_barangay", "occupancy_inspected_by", "occupancy_included_personnel_name", "occupancy_type_of_occupancy"];
  clearSelects.forEach(id => {
    const el = getEl(id);
    if (el) el.selectedIndex = 0;
  });

  const photoInput = document.getElementById("occupancy_photo");
  if (photoInput) photoInput.value = "";
  const photoLibraryInput = document.getElementById("occupancy_photo_library");
  if (photoLibraryInput) photoLibraryInput.value = "";

  const indicator = document.getElementById("occupancy-photo-indicator");
  if (indicator) {
    indicator.className = "photo-attach-indicator";
    indicator.textContent = "";
  }

  const overlay = document.getElementById("occupancy-modal-overlay");
  overlay?.classList.add("open");

  // Reset to first step
  updateModalStepUI('occupancy', 1);
}


function occupancyCloseModal() {
  photoPreviewCancel();
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
      await gasRequest("delete", { table: "occupancy_logbook", id: row.id });
    } catch (err) {
      logbookShowToast("occupancy-toast", "⚠️ Delete failed: " + (err?.message || err));
    }
  })();
}

async function occupancySaveEntry(e) {
  if (e?.preventDefault) e.preventDefault();

  if (occupancyExifProcessingPromise) {
    await occupancyExifProcessingPromise;
  }

  let barangay = (document.getElementById("occupancy_addr_barangay") || { value: "" }).value.trim();
  const line = (document.getElementById("occupancy_addr_line") || { value: "" }).value.trim();

  // Handle special 'Select barangay' residual value if someone clicks but doesn't choose
  if (!barangay || /^select/i.test(barangay)) {
    barangay = "";
  }

  const entry = {
    log_date: (document.getElementById("occupancy_date") || { value: "" }).value,
    io_number: (document.getElementById("occupancy_io_number") || { value: "" }).value.trim(),
    fsic_number: (document.getElementById("occupancy_fsic_number") || { value: "" }).value.trim(),
    owner_name: (document.getElementById("occupancy_owner_name") || { value: "" }).value.trim(),
    owner_phone: (document.getElementById("occupancy_owner_phone") || { value: "" }).value.trim(),
    business_name: (document.getElementById("occupancy_property_name") || { value: "" }).value.trim(),
    type_of_occupancy: (document.getElementById("occupancy_type_of_occupancy") || { value: "" }).value.trim(),
    addr_barangay: barangay,
    addr_line: line,
    inspectors: (document.getElementById("occupancy_inspected_by") || { value: "" }).value.trim(),
    inspector_position: (document.getElementById("occupancy_inspector_position") || { value: "" }).value.trim(),
    included_personnel_name: (document.getElementById("occupancy_included_personnel_name") || { value: "" }).value.trim(),
    included_personnel_position: (document.getElementById("occupancy_included_personnel_position") || { value: "" }).value.trim(),
    duration_start: (document.getElementById("occupancy_duration_start") || { value: "" }).value,
    duration_end: (document.getElementById("occupancy_duration_end") || { value: "" }).value,
    remarks_signature: "",
    // Optional coordinates and photo metadata extracted from EXIF / geolocation
    lat: normalizeGeoNumber(occupancyExifLat),
    lng: normalizeGeoNumber(occupancyExifLng),
    photo_url: occupancyExifPreviewUrl,
    photo_taken_at: occupancyExifTakenAt,
    created_at: new Date().toISOString(),
  };

  const isIncludedPersonnelPlaceholder =
    !entry.included_personnel_name || /^select/i.test(String(entry.included_personnel_name).trim());
  if (isIncludedPersonnelPlaceholder) {
    entry.included_personnel_name = "";
    entry.included_personnel_position = "";
  }

  if (occupancyEditingIdx !== null) {
    const prev = occupancyData[occupancyEditingIdx] || {};
    const hasNewPhoto = !!occupancyExifFile;
    if (!hasNewPhoto) {
      entry.photo_url = prev.photo_url ?? entry.photo_url;
      entry.photo_taken_at = prev.photo_taken_at ?? entry.photo_taken_at;
      entry.lat = normalizeGeoNumber(prev.lat ?? entry.lat);
      entry.lng = normalizeGeoNumber(prev.lng ?? entry.lng);
    } else {
      const exLat = normalizeGeoNumber(occupancyExifLat);
      const exLng = normalizeGeoNumber(occupancyExifLng);
      if (exLat != null && exLng != null) {
        entry.lat = exLat;
        entry.lng = exLng;
      } else if (normalizeGeoNumber(prev.lat) != null && normalizeGeoNumber(prev.lng) != null) {
        entry.lat = normalizeGeoNumber(prev.lat);
        entry.lng = normalizeGeoNumber(prev.lng);
      }
    }
  }

  // Removed fallback to user's current location if photo has no GPS EXIF
  // as per user request: "do not put lat long if photo do not have lat long"

  entry.lat = normalizeGeoNumber(entry.lat);
  entry.lng = normalizeGeoNumber(entry.lng);

  const isPlaceholderInspector =
    !entry.inspectors || /^select\s+inspector$/i.test(String(entry.inspectors).trim());
  if (isPlaceholderInspector) entry.inspectors = "";

  if (!entry.log_date || !entry.io_number) {
    logbookShowToast("occupancy-toast", "⚠️ Please fill in Date and IO Number.");
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
      // If we have a photo file and GAS is enabled, upload to Google Drive via GAS
      let occPhotoUploadedUrl = null;
      if (occupancyExifFile && isGasEnabled()) {
        let uploadFile = occupancyExifFile;
        try {
          uploadFile = await sanitizeInspectionImage(uploadFile);
        } catch (sanitizeErr) {
          console.warn("Occupancy image sanitization failed, using original:", sanitizeErr);
          uploadFile = occupancyExifFile;
        }
        logbookShowToast("occupancy-toast", "Uploading photo to Drive…");
        try {
          const base64Data = await fileToBase64(uploadFile);
          const uploadResult = await gasRequest("upload", {
            filename: `occupancy-${Date.now()}.${(uploadFile.name || "file.bin").split(".").pop() || "bin"}`,
            mimeType: uploadFile.type || "application/octet-stream",
            base64Data,
          });
          if (uploadResult?.data?.url) {
            occPhotoUploadedUrl = uploadResult.data.url;
            entry.photo_url = occPhotoUploadedUrl;
          } else {
            logbookShowToast("occupancy-toast", "⚠️ Upload returned no URL — check Drive folder.");
          }
        } catch (uploadErr) {
          console.error("Occupancy photo upload error:", uploadErr);
          logbookShowToast("occupancy-toast", "⚠️ Photo upload failed: " + (uploadErr?.message || uploadErr));
        }
      }

      const payload = {
        log_date: entry.log_date,
        io_number: entry.io_number,
        fsic_number: entry.fsic_number || null,
        owner_name: entry.owner_name || null,
        owner_phone: entry.owner_phone || null,
        business_name: entry.business_name || null,
        type_of_occupancy: entry.type_of_occupancy || null,
        address: entry.addr_line || null, // Mapping addr_line to address for consistency
        inspectors: entry.inspectors,
        inspector_position: entry.inspector_position || null,
        included_personnel_name: entry.included_personnel_name || null,
        included_personnel_position: entry.included_personnel_position || null,
        duration_start: entry.duration_start || null,
        duration_end: entry.duration_end || null,
        remarks_signature: null,
        latitude: entry.lat ?? null,
        longitude: entry.lng ?? null,
        photo_url: entry.photo_url ?? null,
        photo_taken_at: entry.photo_taken_at ?? null,
      };
      payload.latitude = normalizeGeoNumber(payload.latitude);
      payload.longitude = normalizeGeoNumber(payload.longitude);
      if (occupancyEditingId && (!payload.owner_name || String(payload.owner_name).trim() === "")) {
        delete payload.owner_name;
      }
      let occSavedId = occupancyEditingId;
      if (occupancyEditingId) {
        await gasRequest("update", { table: "occupancy_logbook", id: occupancyEditingId, row: payload });
      } else {
        const insertResult = await gasRequest("insert", { table: "occupancy_logbook", row: payload });
        occSavedId = insertResult?.data?.id ?? null;
      }

      // ── Guaranteed photo_url patch ──────────────────────────────────────
      if (occPhotoUploadedUrl && occSavedId) {
        try {
          await gasRequest("patch_photo_url", {
            table: "occupancy_logbook",
            id: occSavedId,
            url: occPhotoUploadedUrl,
          });
        } catch (patchErr) {
          console.warn("[FSIS] occ photo_url patch failed:", patchErr);
        }
      }

      if (
        occSavedId &&
        Number.isFinite(entry.lat) &&
        Number.isFinite(entry.lng)
      ) {
        try {
          await gasRequest("patch_lat_lng", {
            table: "occupancy_logbook",
            id: occSavedId,
            latitude: entry.lat,
            longitude: entry.lng,
          });
        } catch (patchErr) {
          console.warn("[FSIS] occ patch_lat_lng failed:", patchErr);
        }
      }

      await occupancyLoadFromSupabase();
      occupancyRenderTable();
      renderOccupancyMarkersBatched();
      logbookShowToast("occupancy-toast", occPhotoUploadedUrl ? "Saved + photo linked ✓" : "Saved to database.");

      const hasLocation =
        Number.isFinite(entry.lat) && Number.isFinite(entry.lng);
      const isEdit = occupancyEditingId != null;
      if (hasLocation && !isEdit) {
        // No longer jumping to map as per user request.
        // Stay in the logbook and highlight the new row.
        showView("occupancy");
        window.location.hash = "occupancy";
        occupancyRenderTable();

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
      logbookShowToast("occupancy-toast", "Save failed: " + msg);
    }
  })();
}

async function occupancyLoadFromSupabase() {
  const result = await gasRequest("read", { table: "occupancy_logbook" });
  const rows = result.data || [];
  occupancyData = rows.map((r) => ({
    id: r.id,
    log_date: r.log_date,
    io_number: r.io_number,
    fsic_number: r.fsic_number || "",
    owner_name: r.owner_name || "",
    owner_phone: r.owner_phone || "",
    business_name: r.business_name || "",
    type_of_occupancy: r.type_of_occupancy || "",
    addr_barangay: r.business_name ? "" : "", // Note: To be fully consistent, barangay mapped from address
    addr_line: r.address || "",
    inspectors: r.inspectors || "",
    inspector_position: r.inspector_position || "",
    included_personnel_name: r.included_personnel_name || "",
    included_personnel_position: r.included_personnel_position || "",
    duration_start: r.duration_start || null,
    duration_end: r.duration_end || null,
    remarks_signature: r.remarks_signature || "",
    lat: r.latitude ?? null,
    lng: r.longitude ?? null,
    photo_url: r.photo_url ?? null,
    photo_taken_at: r.photo_taken_at ?? null,
    created_at: r.created_at,
  }));
  occupancySaveToLocal();
}

async function occupancyInitData() {
  if (occupancyDataLoaded) return;
  occupancyDataLoaded = true;
  localStorage.removeItem(OCCUPANCY_STORAGE_KEY);
  occupancyRenderTable();
  renderOccupancyMarkersBatched();
  if (!isGasEnabled()) return;
  try {
    await occupancyLoadFromSupabase();
    occupancyRenderTable();
    renderOccupancyMarkersBatched();
  } catch (err) {
    console.warn("Occupancy load from GAS failed:", err);
    logbookShowToast("occupancy-toast", "Could not load data from server.");
    occupancyRenderTable();
    renderOccupancyMarkersBatched();
  }
}

// --- CSV Export Utility ---
function exportTableToCSV(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const rows = table.querySelectorAll('tr');
  const csv = [];
  for (let i = 0; i < rows.length; i++) {
    const row = [], cols = rows[i].querySelectorAll('td, th');
    for (let j = 0; j < cols.length; j++) {
      let data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, '').replace(/(\s\s)/gm, ' ');
      data = data.replace(/"/g, '""');
      row.push('"' + data + '"');
    }
    csv.push(row.join(','));
  }
  const csvString = csv.join('\n');
  const blob = new Blob([csvString], { type: "text/csv;charset=utf-8;" });
  // IE11 & Edge support
  if (navigator.msSaveBlob) {
    navigator.msSaveBlob(blob, filename);
  } else {
    // Other browsers
    const link = document.createElement("a");
    if (link.download !== undefined) {
      const url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", filename);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  }
}
