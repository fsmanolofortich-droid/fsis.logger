// ============================================================
// FSIS Logger — Admin panel
// ============================================================

const SESSION_KEY = "fsis.session";
const ADMIN_AUTH_KEY = "fsis.admin.auth";
const DEFAULT_GAS_URL =
  "https://script.google.com/macros/s/AKfycbwJmqg6lRB_W95VNY9XfAyAovcbJrm8VpPXXg1pP1ujFD10k85xTpbwO5v8RVyy8Bpc/exec";
const GAS_URL_STORAGE_KEY = "fsis.gas_url";
const TABLE_COLS = 7;

/** @type {Array<Record<string, unknown>>} */
let usersCache = [];

function resolveGasUrl() {
  try {
    const fromStorage = (localStorage.getItem(GAS_URL_STORAGE_KEY) || "").trim();
    if (fromStorage) return fromStorage;
  } catch (_) { /* ignore */ }
  const host = window.location.hostname || "";
  const href = window.location.href || "";
  if (host.includes("script.google.com") || host.includes("script.googleusercontent.com")) {
    return href.split("?")[0];
  }
  return DEFAULT_GAS_URL;
}

const GAS_URL = resolveGasUrl();

async function gasRequest(action, payload) {
  const res = await fetch(GAS_URL, {
    method: "POST",
    body: JSON.stringify({ action, ...(payload || {}) }),
  });
  const raw = await res.text();
  let json;
  try {
    json = JSON.parse(raw);
  } catch {
    throw new Error("Invalid API response. Check your deployed Apps Script URL.");
  }
  if (json.error) throw new Error(json.error);
  return json;
}

function getSession() {
  const raw = sessionStorage.getItem(SESSION_KEY) ?? localStorage.getItem(SESSION_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function getAdminAuth() {
  try {
    const raw = sessionStorage.getItem(ADMIN_AUTH_KEY);
    if (!raw) return null;
    const auth = JSON.parse(raw);
    if (!auth?.username || !auth?.password) return null;
    return auth;
  } catch {
    return null;
  }
}

function clearSession() {
  sessionStorage.removeItem(SESSION_KEY);
  localStorage.removeItem(SESSION_KEY);
  sessionStorage.removeItem(ADMIN_AUTH_KEY);
}

function withAdminAuth(extra) {
  const auth = getAdminAuth();
  if (!auth) return null;
  return { username: auth.username, password: auth.password, ...(extra || {}) };
}

function isSuspended(row) {
  return row.suspended === true || String(row.suspended).toLowerCase() === "true";
}

function isSelfUser(row) {
  const session = getSession();
  if (!session?.userId || !row?.id) return false;
  return String(session.userId) === String(row.id);
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

function requireAdminAccess() {
  const session = getSession();
  if (!session?.username || String(session.role || "").toLowerCase() !== "admin") {
    window.location.replace("./index.html");
    return false;
  }
  if (!getAdminAuth()) {
    showToast("Please sign in again with your admin account.", true);
    setTimeout(() => window.location.replace("./index.html"), 1200);
    return false;
  }
  const name = session.displayName || session.username;
  const signedIn = document.getElementById("admin-signed-in");
  if (signedIn) signedIn.textContent = name;
  const navUser = document.getElementById("navUserName");
  if (navUser) navUser.textContent = name;
  return true;
}

function initAdminNav() {
  const navSidebarEl = document.getElementById("navSidebar");
  if (navSidebarEl) {
    navSidebarEl.addEventListener("show.bs.offcanvas", () => {
      document.body.classList.add("nav-offcanvas-open");
    });
    navSidebarEl.addEventListener("hidden.bs.offcanvas", () => {
      document.body.classList.remove("nav-offcanvas-open");
    });
  }
}

function setUsersLoading(loading) {
  const tbody = document.getElementById("users-tbody");
  const empty = document.getElementById("users-empty");
  const wrap = document.querySelector(".admin-table-wrap");
  if (!tbody) return;
  if (loading) {
    if (empty) empty.style.display = "none";
    if (wrap) wrap.hidden = false;
    tbody.innerHTML =
      `<tr><td colspan="${TABLE_COLS}" class="admin-table-loading">Loading users…</td></tr>`;
  }
}

function renderUsersTable(rows) {
  const tbody = document.getElementById("users-tbody");
  const empty = document.getElementById("users-empty");
  const countEl = document.getElementById("users-count");
  const wrap = document.querySelector(".admin-table-wrap");
  if (!tbody) return;

  usersCache = rows || [];
  if (countEl) countEl.textContent = String(usersCache.length);

  tbody.innerHTML = "";
  if (usersCache.length === 0) {
    if (wrap) wrap.hidden = true;
    if (empty) empty.style.display = "block";
    return;
  }

  if (wrap) wrap.hidden = false;
  if (empty) empty.style.display = "none";

  usersCache.forEach((row) => {
    const tr = document.createElement("tr");
    if (isSuspended(row)) tr.classList.add("is-suspended");

    const created = row.created_at
      ? new Date(row.created_at).toLocaleDateString("en-PH", { dateStyle: "short" })
      : "—";
    const roleClass = row.role === "admin" ? "admin-role--admin" : "admin-role--user";
    const email = row.email ? String(row.email) : "—";
    const suspended = isSuspended(row);
    const statusClass = suspended ? "admin-status--suspended" : "admin-status--active";
    const statusLabel = suspended ? "Suspended" : "Active";

    tr.innerHTML = `
      <td class="username-cell">${escapeHtml(row.username || "")}</td>
      <td>${escapeHtml(email)}</td>
      <td>${escapeHtml(row.display_name || "—")}</td>
      <td><span class="admin-role ${roleClass}">${escapeHtml(row.role || "user")}</span></td>
      <td><span class="admin-status ${statusClass}">${statusLabel}</span></td>
      <td>${escapeHtml(created)}</td>
      <td class="admin-table-actions"></td>`;

    const actions = tr.querySelector(".admin-table-actions");
    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "admin-btn admin-btn--ghost admin-btn--sm";
    editBtn.innerHTML = '<i class="bi bi-pencil" aria-hidden="true"></i> Edit';
    editBtn.addEventListener("click", () => openEditUserModal(row));
    actions?.appendChild(editBtn);

    if (!isSelfUser(row)) {
      const suspendBtn = document.createElement("button");
      suspendBtn.type = "button";
      suspendBtn.className = suspended
        ? "admin-btn admin-btn--ghost admin-btn--sm"
        : "admin-btn admin-btn--warn admin-btn--sm";
      suspendBtn.innerHTML = suspended
        ? '<i class="bi bi-check-circle" aria-hidden="true"></i> Restore'
        : '<i class="bi bi-slash-circle" aria-hidden="true"></i> Suspend';
      suspendBtn.addEventListener("click", () => {
        void toggleUserSuspended(row, !suspended);
      });
      actions?.appendChild(suspendBtn);
    }

    tbody.appendChild(tr);
  });
}

async function loadUsers() {
  if (!requireAdminAccess()) return;
  const payload = withAdminAuth();
  if (!payload) return;

  setUsersLoading(true);
  try {
    const result = await gasRequest("list_users", payload);
    renderUsersTable(result.data || []);
  } catch (err) {
    const tbody = document.getElementById("users-tbody");
    if (tbody) {
      tbody.innerHTML =
        `<tr><td colspan="${TABLE_COLS}" class="admin-table-error">${escapeHtml(err.message || "Failed to load users")}</td></tr>`;
    }
    showToast(err.message || "Failed to load users", true);
  }
}

async function toggleUserSuspended(user, suspended) {
  if (!requireAdminAccess()) return;
  const label = user.username || user.email || "this user";
  const msg = suspended
    ? `Suspend account "${label}"? They will not be able to sign in.`
    : `Restore account "${label}"?`;
  if (!window.confirm(msg)) return;

  const payload = withAdminAuth({
    userId: user.id,
    suspended: Boolean(suspended),
  });
  if (!payload) return;

  try {
    await gasRequest("suspend_user", payload);
    showToast(suspended ? "Account suspended." : "Account restored.", false);
    await loadUsers();
  } catch (err) {
    showToast(err.message || "Could not update account status", true);
  }
}

function openEditUserModal(user) {
  const overlay = document.getElementById("edit-user-overlay");
  if (!overlay) return;
  document.getElementById("edit-user-id").value = user.id || "";
  document.getElementById("edit-username").value = user.username || "";
  document.getElementById("edit-email").value = user.email || "";
  document.getElementById("edit-display-name").value = user.display_name || "";
  document.getElementById("edit-password").value = "";
  document.getElementById("edit-role").value = user.role === "admin" ? "admin" : "user";
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  document.getElementById("edit-username")?.focus();
}

function closeEditUserModal() {
  const overlay = document.getElementById("edit-user-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

function openCreateUserModal() {
  const overlay = document.getElementById("create-user-overlay");
  if (!overlay) return;
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  document.getElementById("new-username")?.focus();
}

function closeCreateUserModal() {
  const overlay = document.getElementById("create-user-overlay");
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

async function saveEditUser(e) {
  e.preventDefault();
  if (!requireAdminAccess()) return;
  const payload = withAdminAuth({
    userId: document.getElementById("edit-user-id")?.value,
    targetUsername: (document.getElementById("edit-username")?.value || "").trim(),
    email: (document.getElementById("edit-email")?.value || "").trim(),
    displayName: (document.getElementById("edit-display-name")?.value || "").trim(),
    newPassword: document.getElementById("edit-password")?.value || "",
    role: document.getElementById("edit-role")?.value || "user",
  });
  if (!payload) return;

  const btn = document.getElementById("edit-save-btn");
  if (btn) btn.disabled = true;

  try {
    await gasRequest("update_user", payload);
    const session = getSession();
    if (session && String(payload.userId) === String(session.userId) && payload.newPassword) {
      try {
        sessionStorage.setItem(
          ADMIN_AUTH_KEY,
          JSON.stringify({
            username: payload.targetUsername || session.username,
            password: payload.newPassword,
          })
        );
      } catch (_) { /* ignore */ }
    }
    showToast("Account updated.", false);
    closeEditUserModal();
    await loadUsers();
  } catch (err) {
    showToast(err.message || "Update failed", true);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function createUser(e) {
  e.preventDefault();
  if (!requireAdminAccess()) return;
  const payload = withAdminAuth({
    newUsername: (document.getElementById("new-username")?.value || "").trim(),
    email: (document.getElementById("new-email")?.value || "").trim(),
    displayName: (document.getElementById("new-display-name")?.value || "").trim(),
    newPassword: document.getElementById("new-password")?.value || "",
    role: document.getElementById("new-role")?.value || "user",
  });
  if (!payload) return;

  if (!payload.newUsername) {
    showToast("Username is required.", true);
    return;
  }
  if (!payload.newPassword || payload.newPassword.length < 4) {
    showToast("Password must be at least 4 characters.", true);
    return;
  }

  const btn = document.getElementById("create-user-btn");
  if (btn) btn.disabled = true;

  try {
    await gasRequest("create_user", payload);
    showToast("Account created: " + payload.newUsername, false);
    document.getElementById("new-username").value = "";
    document.getElementById("new-email").value = "";
    document.getElementById("new-display-name").value = "";
    document.getElementById("new-password").value = "";
    closeCreateUserModal();
    await loadUsers();
  } catch (err) {
    showToast(err.message || "Failed to create account", true);
  } finally {
    if (btn) btn.disabled = false;
  }
}

function init() {
  if (!GAS_URL) {
    showToast("Backend URL is not configured.", true);
    return;
  }

  initAdminNav();

  document.getElementById("admin-logout-btn")?.addEventListener("click", () => {
    clearSession();
    window.location.replace("./index.html");
  });

  document.getElementById("open-create-user-btn")?.addEventListener("click", openCreateUserModal);
  document.getElementById("create-user-form")?.addEventListener("submit", createUser);
  document.getElementById("create-cancel-btn")?.addEventListener("click", closeCreateUserModal);
  document.getElementById("create-cancel-btn-2")?.addEventListener("click", closeCreateUserModal);
  document.getElementById("create-user-overlay")?.addEventListener("click", (ev) => {
    if (ev.target.id === "create-user-overlay") closeCreateUserModal();
  });
  document.getElementById("edit-user-form")?.addEventListener("submit", saveEditUser);
  document.getElementById("edit-cancel-btn")?.addEventListener("click", closeEditUserModal);
  document.getElementById("edit-cancel-btn-2")?.addEventListener("click", closeEditUserModal);
  document.getElementById("edit-user-overlay")?.addEventListener("click", (ev) => {
    if (ev.target.id === "edit-user-overlay") closeEditUserModal();
  });

  if (requireAdminAccess()) {
    void loadUsers();
  }
}

document.addEventListener("DOMContentLoaded", init);

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
