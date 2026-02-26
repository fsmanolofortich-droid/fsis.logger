const SESSION_KEY = "fsis.session";
const SUPABASE_URL = "https://drqgbkninqpvhvhnbatk.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRycWdia25pbnFwdmh2aG5iYXRrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE5MTQzODQsImV4cCI6MjA4NzQ5MDM4NH0.CliUD5Ow17OXvaqDzYdAbi-rrTg_u-e4OyomcGrgZk0";

const supabaseClient =
  window.supabase?.createClient && SUPABASE_URL && SUPABASE_ANON_KEY
    ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
    : null;

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

    if (!supabaseClient) {
      setError("Login service not available. Please try again later.");
      return;
    }

    try {
      const { data, error } = await supabaseClient.rpc("app_login", {
        p_username: u,
        p_password: p,
      });

      if (error) {
        console.error(error);
        setError("Login failed. Please try again.");
        return;
      }

      const user = Array.isArray(data) ? data[0] : null;
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
      setError("Login failed. Please check your connection and try again.");
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
      } catch (_) {}
      window.location.href = "./admin.html";
    });
  }
}

document.addEventListener("DOMContentLoaded", init);
