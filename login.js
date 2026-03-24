const SESSION_KEY = "fsis.session";

// ── Google Apps Script backend ────────────────────────────────────────────────
// Paste your deployed Web App URL below after deploying Code.gs
const GAS_URL = "https://script.google.com/macros/s/AKfycbyciiKsA8h82z8VdJHM1NFfFgdSxjWJ8kSfLHw9F6MmShYxXyJmf7qZbbLfWY3hkJyn/exec";

async function gasRequest(action, payload) {
  const res = await fetch(GAS_URL, {
    method: "POST",
    body: JSON.stringify({ action, ...(payload || {}) }),
  });
  const json = await res.json();
  if (json.error) throw new Error(json.error);
  return json;
}

function isGasEnabled() {
  return Boolean(GAS_URL);
}

function nowIso() {
  return new Date().toISOString();
}

function setSession(session) {
  const raw = JSON.stringify(session);
  if (session.rememberMe) localStorage.setItem(SESSION_KEY, raw);
  else sessionStorage.setItem(SESSION_KEY, raw);
}

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

function redirectToHome() {
  window.location.replace("./home.html#map");
}

function setError(message) {
  const el = document.getElementById("loginError");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
  el.setAttribute("aria-hidden", "false");
}

function clearError() {
  const el = document.getElementById("loginError");
  if (!el) return;
  el.textContent = "";
  el.hidden = true;
  el.setAttribute("aria-hidden", "true");
}

function normalize(s) {
  return (s ?? "").trim();
}

function init() {
  const existing = getSession();
  if (existing?.username) redirectToHome();

  const form = document.getElementById("loginForm");
  const username = document.getElementById("username");
  const password = document.getElementById("password");
  const rememberMe = document.getElementById("rememberMe");
  const adminYear = document.getElementById("adminYear");
  const adminGate = document.getElementById("adminSecretGate");
  const adminSecretInput = document.getElementById("adminSecretInput");
  const openAdminBtn = document.getElementById("openAdminBtn");

  if (!(form instanceof HTMLFormElement)) return;

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    clearError();

    const u = normalize(username?.value);
    const p = normalize(password?.value);
    const remember = Boolean(rememberMe?.checked);

    if (!u || !p) {
      setError("Please enter a username and password.");
      return;
    }

    if (!isGasEnabled()) {
      setError("Login service not available. Please configure the GAS_URL.");
      return;
    }

    try {
      const result = await gasRequest("login", { username: u, password: p });

      const user = Array.isArray(result.data) ? result.data[0] : null;
      if (!user) {
        setError("Invalid username or password.");
        return;
      }

      setSession({
        userId: user.id,
        username: user.username,
        displayName: user.display_name || user.username,
        role: user.role,
        issuedAt: nowIso(),
        rememberMe: remember,
      });

      redirectToHome();
    } catch (err) {
      console.error(err);
      setError(err.message || "Login failed. Please check your connection and try again.");
    }
  });

  // Hidden admin gate: click year 5 times to reveal secret input
  let adminYearClicks = 0;
  if (adminYear && adminGate && adminSecretInput && openAdminBtn) {
    adminYear.addEventListener("click", () => {
      adminYearClicks += 1;
      if (adminYearClicks >= 5) {
        adminGate.style.display = "block";
        adminSecretInput.focus();
      }
    });

    openAdminBtn.addEventListener("click", () => {
      const secret = normalize(adminSecretInput.value);
      if (!secret) {
        setError("Enter the admin secret.");
        return;
      }
      try {
        sessionStorage.setItem("fsis.admin.secret", secret);
      } catch (_) { }
      window.location.href = "./admin.html";
    });
  }
}

document.addEventListener("DOMContentLoaded", init);
