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

function isUnknownActionError(err) {
  return /Unknown action/i.test(String(err?.message || err || ""));
}

const DEPLOY_HINT =
  "Redeploy Apps Script: open your Google Sheet → Extensions → Apps Script → paste Code.gs → Deploy → Manage deployments → Edit → New version → Deploy.";

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

function normalizeUserStatus(status) {
  const s = String(status ?? "")
    .trim()
    .toLowerCase();
  if (s === "suspended" || s === "inactive" || s === "disabled") return "suspended";
  return "active";
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

function openOverlay(id) {
  document.getElementById(id)?.classList.add("open");
}

function closeOverlay(id) {
  document.getElementById(id)?.classList.remove("open");
}

function closeAllModals() {
  closeOverlay("create-user-overlay");
  closeOverlay("edit-user-overlay");
  closeOverlay("delete-user-overlay");
  closeOverlay("status-user-overlay");
}

function openCreateUserModal() {
  openOverlay("create-user-overlay");
  document.getElementById("new-username")?.focus();
}

function closeCreateUserModal() {
  closeOverlay("create-user-overlay");
}

function adminCloseCreateOnOverlay(e) {
  if (e.target?.id === "create-user-overlay") closeCreateUserModal();
}

function openEditUserModal(user) {
  if (!user) return;
  document.getElementById("edit-user-id").value = user.id || "";
  const label = document.getElementById("edit-user-username-label");
  if (label) label.innerHTML = "Username: <strong>" + escapeHtml(user.username || "") + "</strong>";
  document.getElementById("edit-display-name").value = user.display_name || "";
  document.getElementById("edit-role").value = normalizeUserRole(user.role);
  document.getElementById("edit-password").value = "";
  openOverlay("edit-user-overlay");
  document.getElementById("edit-display-name")?.focus();
}

function closeEditUserModal() {
  closeOverlay("edit-user-overlay");
}

function adminCloseEditOnOverlay(e) {
  if (e.target?.id === "edit-user-overlay") closeEditUserModal();
}

function openDeleteUserModal(user) {
  if (!user) return;
  document.getElementById("delete-user-id").value = user.id || "";
  const msg = document.getElementById("delete-user-message");
  if (msg) {
    msg.innerHTML =
      "Delete account <strong>" + escapeHtml(user.username || "") + "</strong>?";
  }
  openOverlay("delete-user-overlay");
}

function closeDeleteUserModal() {
  closeOverlay("delete-user-overlay");
}

function adminCloseDeleteOnOverlay(e) {
  if (e.target?.id === "delete-user-overlay") closeDeleteUserModal();
}

function openStatusUserModal(user, nextStatus) {
  if (!user || !nextStatus) return;
  const isSuspend = nextStatus === "suspended";
  document.getElementById("status-user-id").value = user.id || "";
  document.getElementById("status-user-next").value = nextStatus;

  const heading = document.getElementById("status-user-heading");
  const icon = document.getElementById("status-user-icon");
  const msg = document.getElementById("status-user-message");
  const hint = document.getElementById("status-user-hint");
  const confirmBtn = document.getElementById("btnConfirmStatusUser");

  if (heading) heading.textContent = isSuspend ? "Suspend account" : "Activate account";
  if (icon) icon.className = isSuspend ? "bi bi-slash-circle" : "bi bi-person-check";
  if (msg) {
    msg.innerHTML = isSuspend
      ? "Suspend <strong>" + escapeHtml(user.username || "") + "</strong>?"
      : "Reactivate <strong>" + escapeHtml(user.username || "") + "</strong>?";
  }
  if (hint) {
    hint.textContent = isSuspend
      ? "Suspended users cannot sign in until reactivated."
      : "The user will be able to sign in again.";
  }
  if (confirmBtn) {
    confirmBtn.innerHTML = isSuspend
      ? '<i class="bi bi-slash-circle" aria-hidden="true"></i> Suspend account'
      : '<i class="bi bi-person-check" aria-hidden="true"></i> Activate account';
    confirmBtn.classList.toggle("admin-btn--warn", isSuspend);
    confirmBtn.classList.toggle("admin-btn--primary", !isSuspend);
  }

  openOverlay("status-user-overlay");
}

function closeStatusUserModal() {
  closeOverlay("status-user-overlay");
}

function adminCloseStatusOnOverlay(e) {
  if (e.target?.id === "status-user-overlay") closeStatusUserModal();
}

/** Call backend to delete a user by id (with fallbacks for older deployments). */
async function deleteUserAccount(id) {
  if (!id) throw new Error("User id required.");
  try {
    return await gasRequest("delete_user", { adminUsername, id });
  } catch (err) {
    if (!isUnknownActionError(err)) throw err;
    try {
      return await gasRequest("delete", { adminUsername, table: "users", id });
    } catch (err2) {
      if (isUnknownActionError(err2)) {
        throw new Error("Delete is not available on this backend. " + DEPLOY_HINT);
      }
      throw err2;
    }
  }
}

/** Call backend to set active or suspended (with fallbacks for older deployments). */
async function setUserAccountStatus(id, status) {
  if (!id) throw new Error("User id required.");
  const next = String(status || "").trim().toLowerCase();
  if (next !== "active" && next !== "suspended") {
    throw new Error("Status must be active or suspended.");
  }
  try {
    return await gasRequest("set_user_status", { adminUsername, id, status: next });
  } catch (err) {
    if (!isUnknownActionError(err)) throw err;
    try {
      return await gasRequest("update_user", { adminUsername, id, status: next });
    } catch (err2) {
      if (!isUnknownActionError(err2)) throw err2;
      try {
        return await gasRequest("update", {
          adminUsername,
          table: "users",
          id,
          row: { status: next },
        });
      } catch (err3) {
        if (isUnknownActionError(err3)) {
          throw new Error("Suspend/activate is not available on this backend. " + DEPLOY_HINT);
        }
        throw err3;
      }
    }
  }
}

/** Warn if the deployed web app is missing admin user actions. */
async function checkBackendCapabilities() {
  try {
    const getRes = await fetch(GAS_URL);
    const getJson = await getRes.json().catch(() => ({}));
    if (getJson.apiVersion >= 3) return;

    const ping = await gasRequest("ping", {});
    if (ping.apiVersion >= 3) return;

    await gasRequest("delete_user", { adminUsername, id: "__probe__" });
  } catch (err) {
    if (isUnknownActionError(err)) {
      showToast(
        "Backend is outdated (missing delete_user / set_user_status). Using fallback APIs; redeploy Code.gs for full support.",
        true
      );
    }
  }
}

/** Open delete confirmation for a user id. */
function deleteUser(userId) {
  const user = findCachedUser(userId);
  if (!user) {
    showToast("User not found.", true);
    return;
  }
  openDeleteUserModal(user);
}

/** Open suspend confirmation for a user id. */
function suspendUser(userId) {
  const user = findCachedUser(userId);
  if (!user) {
    showToast("User not found.", true);
    return;
  }
  openStatusUserModal(user, "suspended");
}

/** Open activate confirmation for a user id. */
function activateUser(userId) {
  const user = findCachedUser(userId);
  if (!user) {
    showToast("User not found.", true);
    return;
  }
  openStatusUserModal(user, "active");
}

function handleUserAction(action, userId, nextStatus) {
  if (!userId) return;

  if (action === "edit") {
    const user = findCachedUser(userId);
    if (!user) {
      showToast("User not found.", true);
      return;
    }
    openEditUserModal(user);
    return;
  }
  if (action === "delete") {
    deleteUser(userId);
    return;
  }
  if (action === "suspend") {
    suspendUser(userId);
    return;
  }
  if (action === "activate") {
    activateUser(userId);
    return;
  }
  if (action === "toggle-status") {
    const user = findCachedUser(userId);
    if (!user) {
      showToast("User not found.", true);
      return;
    }
    const next =
      nextStatus ||
      (normalizeUserStatus(user.status) === "suspended" ? "active" : "suspended");
    openStatusUserModal(user, next);
  }
}

function confirmDeleteUser() {
  const id = document.getElementById("delete-user-id")?.value || "";
  const btn = document.getElementById("btnConfirmDeleteUser");
  const user = findCachedUser(id);
  if (!id) return;

  if (btn) btn.disabled = true;
  deleteUserAccount(id)
    .then(() => {
      showToast("Account deleted" + (user?.username ? ": " + user.username : "."));
      closeDeleteUserModal();
      loadUsers();
    })
    .catch((err) => showToast(err?.message || "Failed to delete account.", true))
    .finally(() => {
      if (btn) btn.disabled = false;
    });
}

function confirmStatusChange() {
  const id = document.getElementById("status-user-id")?.value || "";
  const next = document.getElementById("status-user-next")?.value || "";
  const btn = document.getElementById("btnConfirmStatusUser");
  const user = findCachedUser(id);
  if (!id || !next) return;

  if (btn) btn.disabled = true;
  setUserAccountStatus(id, next)
    .then(() => {
      const name = user?.username || "User";
      showToast(
        next === "suspended" ? "Account suspended: " + name : "Account activated: " + name
      );
      closeStatusUserModal();
      loadUsers();
    })
    .catch((err) => showToast(err?.message || "Failed to update status.", true))
    .finally(() => {
      if (btn) btn.disabled = false;
    });
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
      if (cols[j].classList.contains("col-actions")) continue;
      let data = cols[j].innerText.replace(/(\r\n|\n|\r)/gm, "").replace(/(\s\s)/gm, " ");
      data = data.replace(/"/g, '""');
      row.push('"' + data + '"');
    }
    if (row.length) csv.push(row.join(","));
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
let adminUserId = "";
let usersCache = [];

function setUsersCount(n) {
  const el = document.getElementById("users-count");
  if (el) el.textContent = n === 1 ? "1 user" : n + " users";
}

function findCachedUser(id) {
  return usersCache.find((u) => String(u.id) === String(id));
}

function renderStatusBadge(status) {
  const s = normalizeUserStatus(status);
  if (s === "suspended") {
    return '<span class="admin-status admin-status--suspended">Suspended</span>';
  }
  return '<span class="admin-status admin-status--active">Active</span>';
}

function attrSafe(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;");
}

function renderActionsDropdown(row) {
  const id = attrSafe(row.id);
  const status = normalizeUserStatus(row.status);
  const isSelf =
    String(row.username || "").trim().toLowerCase() === adminUsername.toLowerCase();
  const statusAction = status === "suspended" ? "activate" : "suspend";
  const statusLabel = status === "suspended" ? "Activate account" : "Suspend account";
  const statusIcon = status === "suspended" ? "bi-person-check" : "bi-slash-circle";

  let menuItems =
    '<li><button type="button" class="dropdown-item" data-action="edit" data-id="' +
    id +
    '"><i class="bi bi-pencil me-2" aria-hidden="true"></i>Edit account</button></li>';

  if (!isSelf) {
    menuItems +=
      '<li><hr class="dropdown-divider" role="separator"></li>' +
      '<li><button type="button" class="dropdown-item" data-action="' +
      statusAction +
      '" data-id="' +
      id +
      '"><i class="bi ' +
      statusIcon +
      ' me-2" aria-hidden="true"></i>' +
      statusLabel +
      "</button></li>" +
      '<li><button type="button" class="dropdown-item text-danger" data-action="delete" data-id="' +
      id +
      '"><i class="bi bi-trash3 me-2" aria-hidden="true"></i>Delete account</button></li>';
  }

  return (
    '<div class="dropdown admin-row-actions">' +
    '<button type="button" class="admin-actions-toggle dropdown-toggle" data-bs-toggle="dropdown" data-bs-boundary="viewport" data-bs-display="static" aria-expanded="false" aria-label="Actions for ' +
    escapeHtml(row.username || "user") +
    '">' +
    '<i class="bi bi-three-dots-vertical" aria-hidden="true"></i><span>Actions</span>' +
    "</button>" +
    '<ul class="dropdown-menu dropdown-menu-end admin-actions-menu">' +
    menuItems +
    "</ul>" +
    "</div>"
  );
}

function renderUserRow(row) {
  const tr = document.createElement("tr");
  const status = normalizeUserStatus(row.status);
  if (status === "suspended") tr.classList.add("is-suspended");

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
    renderStatusBadge(status) +
    "</td>" +
    "<td>" +
    escapeHtml(created) +
    "</td>" +
    '<td class="col-actions"><div class="admin-table-actions">' +
    renderActionsDropdown(row) +
    "</div></td>";

  return tr;
}

function wireTableActions() {
  const table = document.getElementById("table-users");
  if (!table || table.dataset.actionsWired === "1") return;
  table.dataset.actionsWired = "1";

  table.addEventListener("click", (e) => {
    const item = e.target.closest?.("[data-action][data-id]");
    if (!item || !table.contains(item)) return;

    e.preventDefault();
    const action = item.getAttribute("data-action");
    const id = item.getAttribute("data-id");
    handleUserAction(action, id);
  });
}

function loadUsers() {
  const tbody = document.getElementById("users-tbody");
  const empty = document.getElementById("users-empty");
  if (!tbody) return;

  tbody.innerHTML =
    '<tr id="users-loading-row"><td colspan="6" class="admin-table-loading">Loading users…</td></tr>';
  if (empty) empty.hidden = true;

  gasRequest("list_users", { adminUsername })
    .then((r) => {
      usersCache = (r.data || []).filter((u) => String(u.username || "").trim() !== "");
      setUsersCount(usersCache.length);
      if (usersCache.length === 0) {
        tbody.innerHTML = "";
        if (empty) empty.hidden = false;
        return;
      }
      if (empty) empty.hidden = true;
      tbody.innerHTML = "";
      usersCache.forEach((row) => tbody.appendChild(renderUserRow(row)));
    })
    .catch((err) => {
      usersCache = [];
      tbody.innerHTML =
        '<tr><td colspan="6" class="admin-table-error">Error: ' +
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
  adminUserId = String(session.userId || "").trim();
  initNavChrome(session);
  wireTableActions();

  document.getElementById("btnOpenCreateUser")?.addEventListener("click", openCreateUserModal);
  document.getElementById("btnCloseCreateUser")?.addEventListener("click", closeCreateUserModal);
  document.getElementById("btnCancelCreateUser")?.addEventListener("click", closeCreateUserModal);
  document.getElementById("btnCloseEditUser")?.addEventListener("click", closeEditUserModal);
  document.getElementById("btnCancelEditUser")?.addEventListener("click", closeEditUserModal);
  document.getElementById("btnCloseDeleteUser")?.addEventListener("click", closeDeleteUserModal);
  document.getElementById("btnCancelDeleteUser")?.addEventListener("click", closeDeleteUserModal);
  document.getElementById("btnCloseStatusUser")?.addEventListener("click", closeStatusUserModal);
  document.getElementById("btnCancelStatusUser")?.addEventListener("click", closeStatusUserModal);
  document.getElementById("btnConfirmStatusUser")?.addEventListener("click", confirmStatusChange);
  document.getElementById("btnConfirmDeleteUser")?.addEventListener("click", confirmDeleteUser);
  document.getElementById("btnExportCsv")?.addEventListener("click", () =>
    exportTableToCSV("table-users", "app_users.csv")
  );

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeAllModals();
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
      .catch((err) => showToast(err?.message || "Failed to create account", true))
      .finally(() => {
        if (submitBtn) submitBtn.disabled = false;
      });
  });

  document.getElementById("edit-user-form")?.addEventListener("submit", (e) => {
    e.preventDefault();
    const id = document.getElementById("edit-user-id")?.value || "";
    const displayName = document.getElementById("edit-display-name")?.value?.trim() || "";
    const role = document.getElementById("edit-role")?.value || "user";
    const password = document.getElementById("edit-password")?.value || "";
    const submitBtn = document.getElementById("btnSubmitEditUser");

    if (!id) return;
    if (password && password.length < 4) {
      showToast("Password must be at least 4 characters.", true);
      return;
    }

    const payload = {
      adminUsername,
      id,
      displayName,
      role,
    };
    if (password) payload.password = password;

    if (submitBtn) submitBtn.disabled = true;
    gasRequest("update_user", payload)
      .then(() => {
        showToast("Account updated.");
        closeEditUserModal();
        loadUsers();
      })
      .catch((err) => showToast(err?.message || "Failed to update account", true))
      .finally(() => {
        if (submitBtn) submitBtn.disabled = false;
      });
  });

  checkBackendCapabilities();
  loadUsers();
}

document.addEventListener("DOMContentLoaded", init);

window.adminCloseCreateOnOverlay = adminCloseCreateOnOverlay;
window.adminCloseEditOnOverlay = adminCloseEditOnOverlay;
window.adminCloseDeleteOnOverlay = adminCloseDeleteOnOverlay;
window.adminCloseStatusOnOverlay = adminCloseStatusOnOverlay;
window.confirmDeleteUser = confirmDeleteUser;
window.confirmStatusChange = confirmStatusChange;
window.deleteUser = deleteUser;
window.suspendUser = suspendUser;
window.activateUser = activateUser;
