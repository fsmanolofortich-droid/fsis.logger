// ============================================================
// FSIS Logger — Login Script
// ============================================================

const SESSION_KEY = "fsis.session";
const ADMIN_AUTH_KEY = "fsis.admin.auth";

// ── Google Apps Script backend ────────────────────────────────────────────────
// Paste your deployed Web App URL below after deploying Code.gs
const GAS_URL = "https://script.google.com/macros/s/AKfycbwJmqg6lRB_W95VNY9XfAyAovcbJrm8VpPXXg1pP1ujFD10k85xTpbwO5v8RVyy8Bpc/exec";

// ── Toast notification system ─────────────────────────────────────────────────

/**
 * Show a toast notification.
 * @param {string} message
 * @param {'error'|'success'|'info'} type
 * @param {number} duration  ms before auto-dismiss (0 = sticky)
 */
function showToast(message, type = "info", duration = 4000) {
  const container = document.getElementById("toastContainer");
  if (!container) return;

  const icons = { error: "bi-exclamation-circle-fill", success: "bi-check-circle-fill", info: "bi-info-circle-fill" };

  const toast = document.createElement("div");
  toast.className = `fsis-toast toast-${type}`;
  toast.innerHTML = `
    <span class="toast-icon"><i class="bi ${icons[type] || icons.info}"></i></span>
    <span class="toast-msg">${message}</span>
  `;

  const dismiss = () => {
    toast.classList.add("hiding");
    toast.addEventListener("animationend", () => toast.remove(), { once: true });
  };

  toast.addEventListener("click", dismiss);
  container.appendChild(toast);

  if (duration > 0) setTimeout(dismiss, duration);
  return dismiss;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

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

function setAdminAuth(username, password) {
  try {
    sessionStorage.setItem(ADMIN_AUTH_KEY, JSON.stringify({ username, password }));
  } catch (_) { }
}

function clearAdminAuth() {
  try {
    sessionStorage.removeItem(ADMIN_AUTH_KEY);
  } catch (_) { }
}

function setSession(session) {
  const { rememberMe, ...profile } = session;
  const raw = JSON.stringify({ ...profile, rememberMe });
  if (rememberMe) localStorage.setItem(SESSION_KEY, raw);
  else sessionStorage.setItem(SESSION_KEY, raw);
}

function isAdminRole(role) {
  return String(role || "").toLowerCase() === "admin";
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

function redirectToAdmin() {
  window.location.replace("./admin.html");
}

function redirectAfterLogin(role) {
  if (isAdminRole(role)) redirectToAdmin();
  else redirectToHome();
}

function normalize(s) {
  return (s ?? "").trim();
}

// ── Inline error banner ───────────────────────────────────────────────────────

function setError(message) {
  const el = document.getElementById("loginError");
  if (!el) return;
  el.textContent = message;
  el.style.display = "block";

  // Shake the card
  const card = document.getElementById("loginCard");
  if (card) {
    card.classList.remove("shake");
    void card.offsetWidth; // reflow to restart animation
    card.classList.add("shake");
  }
}

function clearError() {
  const el = document.getElementById("loginError");
  if (!el) return;
  el.textContent = "";
  el.style.display = "none";
}

// ── Field validation ──────────────────────────────────────────────────────────

function markValid(input) {
  input.classList.remove("is-invalid");
  input.classList.add("is-valid");
}

function markInvalid(input) {
  input.classList.remove("is-valid");
  input.classList.add("is-invalid");
}

function clearValidation(...inputs) {
  inputs.forEach(i => i.classList.remove("is-valid", "is-invalid"));
}

// ── Loading state ─────────────────────────────────────────────────────────────

function setLoading(loading) {
  const btn     = document.getElementById("loginBtn");
  const text    = document.getElementById("loginBtnText");
  const spinner = document.getElementById("loginSpinner");
  const usernameInput = document.getElementById("username");
  const passwordInput = document.getElementById("password");

  if (!btn) return;

  btn.disabled = loading;
  if (text)    text.textContent = loading ? "Signing in…" : "Sign in";
  if (spinner) spinner.style.display = loading ? "block" : "none";
  if (usernameInput) usernameInput.disabled = loading;
  if (passwordInput) passwordInput.disabled = loading;
}

// ── Main init ─────────────────────────────────────────────────────────────────

function init() {
  // Already logged in? Redirect immediately.
  const existing = getSession();
  if (existing?.username) {
    showToast(`Welcome back, ${existing.displayName || existing.username}!`, "success", 2000);
    setTimeout(() => redirectAfterLogin(existing.role), 600);
    return;
  }

  const form            = document.getElementById("loginForm");
  const usernameInput   = document.getElementById("username");
  const passwordInput   = document.getElementById("password");
  const rememberMe      = document.getElementById("rememberMe");
  const togglePw        = document.getElementById("togglePw");
  const togglePwIcon    = document.getElementById("togglePwIcon");
  if (!(form instanceof HTMLFormElement)) return;

  // ── Password show / hide ──────────────────────────────
  if (togglePw && passwordInput && togglePwIcon) {
    togglePw.addEventListener("click", () => {
      const isHidden = passwordInput.type === "password";
      passwordInput.type = isHidden ? "text" : "password";
      togglePwIcon.className = isHidden ? "bi bi-eye-slash" : "bi bi-eye";
      passwordInput.focus();
    });
  }

  // ── Clear validation on input ─────────────────────────
  usernameInput?.addEventListener("input", () => {
    clearValidation(usernameInput);
    clearError();
  });
  passwordInput?.addEventListener("input", () => {
    clearValidation(passwordInput);
    clearError();
  });

  // ── Form submit ───────────────────────────────────────
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    clearError();

    const u = normalize(usernameInput?.value);
    const p = normalize(passwordInput?.value);
    const remember = Boolean(rememberMe?.checked);

    // Client-side validation
    let hasError = false;
    if (!u) {
      markInvalid(usernameInput);
      hasError = true;
    }
    if (!p) {
      markInvalid(passwordInput);
      hasError = true;
    }
    if (hasError) {
      setError("Please fill in all fields.");
      showToast("Username and password are required.", "error");
      return;
    }

    if (!isGasEnabled()) {
      setError("Login service is not configured. Please contact your administrator.");
      showToast("Backend not configured.", "error");
      return;
    }

    setLoading(true);
    const loadingToast = showToast("Signing in, please wait…", "info", 0);

    try {
      const result = await gasRequest("login", { username: u, password: p });

      const user = Array.isArray(result.data) ? result.data[0] : null;
      if (!user) {
        if (loadingToast) loadingToast();
        markInvalid(usernameInput);
        markInvalid(passwordInput);
        setError("Incorrect username or password. Please try again.");
        showToast("Login failed — check your credentials.", "error");
        setLoading(false);
        return;
      }

      // Success
      markValid(usernameInput);
      markValid(passwordInput);
      if (loadingToast) loadingToast();
      showToast(`Welcome, ${user.display_name || user.username}!`, "success", 3000);

      if (isAdminRole(user.role)) {
        setAdminAuth(u, p);
      } else {
        clearAdminAuth();
      }

      setSession({
        userId: user.id,
        username: user.username,
        displayName: user.display_name || user.username,
        role: user.role,
        issuedAt: nowIso(),
        rememberMe: remember,
      });

      const goNext = () => redirectAfterLogin(user.role);

      // Play the success audio and wait for it to finish before redirecting.
      try {
        const audio = new Audio('./fahhhhhhhhhhhhhh.mp3');
        audio.onended = goNext;

        audio.play().catch(e => {
          console.warn("Audio play failed:", e);
          setTimeout(goNext, 900);
        });
      } catch (e) {
        console.warn("Audio initialization failed:", e);
        setTimeout(goNext, 900);
      }


    } catch (err) {
      console.error(err);
      if (loadingToast) loadingToast();
      const msg = err.message?.includes("fetch")
        ? "Network error — check your internet connection."
        : (err.message || "Login failed. Please try again.");
      setError(msg);
      showToast(msg, "error");
      setLoading(false);
    }
  });

}

document.addEventListener("DOMContentLoaded", init);
