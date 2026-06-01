// ============================================================
// FSIS Logger — Admin (account management)
// ============================================================

const SESSION_KEY = "fsis.session";
const DEFAULT_GAS_URL =
  "https://script.google.com/macros/s/AKfycbwJmqg6lRB_W95VNY9XfAyAovcbJrm8VpPXXg1pP1ujFD10k85xTpbwO5v8RVyy8Bpc/exec";
const GAS_URL_STORAGE_KEY = "fsis.gas_url";

function resolveGasUrl() {
  const fromQuery = new URLSearchParams(window.location.search).get("gas_url");
  if (fromQuery) {
    const clean = fromQuery.trim();
    if (clean) {
      try {
        localStorage.setItem(GAS_URL_STORAGE_KEY, clean);
      } catch (_) {}
      return clean;
    }
  }
  try {
    const fromStorage = (localStorage.getItem(GAS_URL_STORAGE_KEY) || "").trim();
    if (fromStorage) return fromStorage;
  } catch (_) {}
  const host = window.location.hostname || "";
  const href = window.location.href || "";
  if (host.includes("script.google.com") || host.includes("script.googleusercontent.com")) {
    return href.split("?")[0];
  }
  return DEFAULT_GAS_URL;
}

const GAS_URL = resolveGasUrl();

async function gasRequest(action, payload) {
  if (!GAS_URL) throw new Error("Backend URL is not configured.");
  const res = await fetch(GAS_URL, {
    method: "POST",
    body: JSON.stringify({ action, ...(payload || {}) }),
  });
  const raw = await res.text();
  let json;
  try {
    json = JSON.parse(raw);
  } catch {
    throw new Error("Invalid response from server.");
  }
  if (json.error) throw new Error(json.error);
  return json;
}

function getSession() {
  try {
    const raw = sessionStorage.getItem(SESSION_KEY) ?? localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function clearSession() {
  sessionStorage.removeItem(SESSION_KEY);
  localStorage.removeItem(SESSION_KEY);
}

function normalizeUserRole(role) {
  const r = String(role ?? "")
    .trim()
    .toLowerCase();
  if (r === "admin" || r === "administrator") return "admin";
  return r || "user";
}

function showToast(msg, isError) {
  const el = document.getElementById("toast");
  if (!el) return;
  el.textContent = msg;
  el.classList.toggle("err", !!isError);
  el.classList.add("show");
  setTimeout(() => el.classList.remove("show"), 4000);
}

function escapeHtml(s) {
  if (s == null) return "";
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function toFriendlyDate(isoString) {
  if (!isoString) return "—";
  const d = new Date(isoString);
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleString("en-PH", { dateStyle: "short", timeStyle: "short" });
}

function requireAdminSession() {
  const session = getSession();
  if (!session?.username) {
    window.location.replace("./index.html");
    return null;
  }
  if (normalizeUserRole(session.role) !== "admin") {
    showToast("Admin access only. Sign in with an admin account.", true);
    setTimeout(() => window.location.replace("./home.html"), 1500);
    return null;
  }
  return session;
}

function openCreateUserModal() {
  const overlay = document.getElementById("create-user-overlay");
  if (!overlay) return;
  overlay.classList.add("open");
  document.getElementById("new-username")?.focus();
}

function closeCreateUserModal() {
  const overlay = document.getElementById("create-user-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
}

function adminCloseCreateOnOverlay(e) {
  if (e.target?.id === "create-user-overlay") closeCreateUserModal();
}

function exportTableToCSV(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const rows = table.querySelectorAll("tr");
  const csv = [];
  for (let i = 0; i < rows.length; i++) {
    const row = [];
    const cols = rows[i].querySelectorAll("td, th");
    for (let j = 0; j < cols.length; j++) {
      let data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, "").replace(/(\s\s)/gm, " ");
      data = data.replace(/"/g, '""');
      row.push('"' + data + '"');
    }
    csv.push(row.join(","));
  }
  const blob = new Blob([csv.join("\n")], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = "hidden";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function initNavChrome(session) {
  const name = session.displayName || session.username || "User";
  const userNameEl = document.getElementById("userName");
  const displayEl = document.getElementById("adminDisplayName");
  const lastEl = document.getElementById("lastLogin");
  if (userNameEl) userNameEl.textContent = name;
  if (displayEl) displayEl.textContent = name;
  if (lastEl) {
    lastEl.textContent = session.issuedAt ? "Signed in " + toFriendlyDate(session.issuedAt) : "";
    lastEl.setAttribute("aria-hidden", lastEl.textContent ? "false" : "true");
  }

  document.getElementById("logoutBtn")?.addEventListener("click", () => {
    clearSession();
    window.location.replace("./index.html");
  });

  const navSidebarEl = document.getElementById("navSidebar");
  const burgerBtn = document.getElementById("burgerBtn");
  if (navSidebarEl) {
    navSidebarEl.addEventListener("show.bs.offcanvas", () => {
      document.body.classList.add("nav-offcanvas-open");
      if (burgerBtn) {
        burgerBtn.setAttribute("aria-expanded", "true");
        burgerBtn.setAttribute("aria-hidden", "true");
      }
    });
    navSidebarEl.addEventListener("hide.bs.offcanvas", () => {
      document.body.classList.remove("nav-offcanvas-open");
      if (burgerBtn) {
        burgerBtn.setAttribute("aria-expanded", "false");
        burgerBtn.removeAttribute("aria-hidden");
      }
    });
  }
}

let adminUsername = "";

function setUsersCount(n) {
  const el = document.getElementById("users-count");
  if (el) el.textContent = n === 1 ? "1 user" : n + " users";
}

function loadUsers() {
  const tbody = document.getElementById("users-tbody");
  const empty = document.getElementById("users-empty");
  if (!tbody) return;

  tbody.innerHTML =
    '<tr id="users-loading-row"><td colspan="4" class="admin-table-loading">Loading users…</td></tr>';
  if (empty) empty.hidden = true;

  gasRequest("list_users", { adminUsername })
    .then((r) => {
      const rows = r.data || [];
      setUsersCount(rows.length);
      if (rows.length === 0) {
        tbody.innerHTML = "";
        if (empty) empty.hidden = false;
        return;
      }
      if (empty) empty.hidden = true;
      tbody.innerHTML = "";
      rows.forEach((row) => {
        const tr = document.createElement("tr");
        const created = row.created_at
          ? new Date(row.created_at).toLocaleDateString("en-PH", { dateStyle: "short" })
          : "—";
        const role = normalizeUserRole(row.role);
        const roleClass = role === "admin" ? "admin-role--admin" : "admin-role--user";
        tr.innerHTML =
          '<td class="username-cell">' +
          escapeHtml(row.username || "") +
          "</td>" +
          "<td>" +
          escapeHtml(row.display_name || "—") +
          "</td>" +
          '<td><span class="admin-role ' +
          roleClass +
          '">' +
          escapeHtml(role) +
          "</span></td>" +
          "<td>" +
          escapeHtml(created) +
          "</td>";
        tbody.appendChild(tr);
      });
    })
    .catch((err) => {
      tbody.innerHTML =
        '<tr><td colspan="4" class="admin-table-error">Error: ' +
        escapeHtml(err?.message || err) +
        "</td></tr>";
      setUsersCount(0);
      showToast(err?.message || "Could not load users.", true);
    });
}

function init() {
  const session = requireAdminSession();
  if (!session) return;

  adminUsername = String(session.username || "").trim();
  initNavChrome(session);

  document.getElementById("btnOpenCreateUser")?.addEventListener("click", openCreateUserModal);
  document.getElementById("btnCloseCreateUser")?.addEventListener("click", closeCreateUserModal);
  document.getElementById("btnCancelCreateUser")?.addEventListener("click", closeCreateUserModal);
  document.getElementById("btnExportCsv")?.addEventListener("click", () =>
    exportTableToCSV("table-users", "app_users.csv")
  );

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeCreateUserModal();
  });

  document.getElementById("create-user-form")?.addEventListener("submit", (e) => {
    e.preventDefault();
    const username = document.getElementById("new-username")?.value?.trim() || "";
    const displayName = document.getElementById("new-display-name")?.value?.trim() || "";
    const password = document.getElementById("new-password")?.value || "";
    const role = document.getElementById("new-role")?.value || "user";
    const submitBtn = document.getElementById("btnSubmitCreateUser");

    if (!username) {
      showToast("Username is required.", true);
      return;
    }
    if (password.length < 4) {
      showToast("Password must be at least 4 characters.", true);
      return;
    }

    if (submitBtn) submitBtn.disabled = true;
    gasRequest("create_user", {
      adminUsername,
      username,
      displayName: displayName || null,
      password,
      role,
    })
      .then(() => {
        showToast("Account created: " + username);
        document.getElementById("new-username").value = "";
        document.getElementById("new-display-name").value = "";
        document.getElementById("new-password").value = "";
        document.getElementById("new-role").value = "user";
        closeCreateUserModal();
        loadUsers();
      })
      .catch((err) => {
        showToast(err?.message || "Failed to create account", true);
      })
      .finally(() => {
        if (submitBtn) submitBtn.disabled = false;
      });
  });

  loadUsers();
}

document.addEventListener("DOMContentLoaded", init);

// Expose for inline onclick on overlay
window.adminCloseCreateOnOverlay = adminCloseCreateOnOverlay;
