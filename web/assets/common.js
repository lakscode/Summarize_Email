// Shared Supabase client and auth helpers for every page.
// Load order: supabase-js UMD, config.js, assets/brand.js, then this file.

const { SUPABASE_URL, SUPABASE_ANON_KEY } = window.APP_CONFIG ?? {};
if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
  document.addEventListener("DOMContentLoaded", () =>
    showMessage("Missing Supabase settings. Fill in web/.env and run `npm run config`.", "error"),
  );
}

// Login lasts for this browser tab/window only: the session lives in sessionStorage, so a new tab must
// log in again and closing the tab signs out. Same-tab redirects (e.g. Microsoft sign-in) keep it.
// Sessions saved in localStorage by earlier versions are removed so they can't sign other tabs in.
try {
  for (const key of Object.keys(localStorage)) {
    if (key.startsWith("sb-") && key.endsWith("-auth-token")) localStorage.removeItem(key);
  }
} catch {
  // Storage blocked: nothing to clean up.
}

const sb = window.supabase.createClient(SUPABASE_URL || "http://invalid", SUPABASE_ANON_KEY || "missing", {
  auth: { storage: window.sessionStorage, persistSession: true },
});

const pageUrl = (name) => new URL(`${name}.html`, location.href).href;

async function getSession() {
  const { data } = await sb.auth.getSession();
  return data.session;
}

// Login/signup pages: skip straight to the dashboard when already signed in.
async function redirectIfSignedIn() {
  if (await getSession()) location.replace(pageUrl("dashboard"));
}

// --- Single sign-on (Supabase Auth OAuth providers; enable them under Authentication > Sign In / Providers). ---

const SSO_PROVIDERS = [
  // Azure only returns the user's email when the `email` scope is requested.
  { provider: "azure", label: "Continue with Microsoft", scopes: "email" },
];

// Fills <div id="sso"> with one button per provider. The provider sends the user back to the login page,
// which shows any error and otherwise forwards to the dashboard. New users get an account on first sign-in.
function renderSsoButtons() {
  const box = document.getElementById("sso");
  if (!box) return;
  const buttons = SSO_PROVIDERS.map(({ provider, label, scopes }) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "secondary";
    button.textContent = label;
    button.addEventListener("click", async () => {
      showMessage("");
      for (const b of buttons) b.disabled = true;
      const { error } = await sb.auth.signInWithOAuth({ provider, options: { redirectTo: pageUrl("login"), scopes } });
      if (error) {
        for (const b of buttons) b.disabled = false;
        showMessage(error.message, "error");
      }
    });
    return button;
  });
  const divider = document.createElement("div");
  divider.className = "divider";
  divider.textContent = "or";
  box.replaceChildren(...buttons, divider);
}

// Errors from a failed SSO sign-in come back in the query string or hash (e.g. the user cancelled).
function ssoErrorFromUrl() {
  const params = new URLSearchParams(location.search);
  const hash = new URLSearchParams(location.hash.slice(1));
  return params.get("error_description") || hash.get("error_description");
}

// Protected pages: bounce to login when signed out.
async function requireSession() {
  const session = await getSession();
  if (!session) {
    location.replace(pageUrl("login"));
    return null;
  }
  sb.auth.onAuthStateChange((event) => {
    if (event === "SIGNED_OUT") location.replace(pageUrl("login"));
  });
  return session;
}

// Every error shown to the user is also recorded on the Logs page.
function showMessage(text, kind = "info") {
  if (text && kind === "error") logUi("error", text);
  const el = document.getElementById("message");
  if (!el) return;
  el.textContent = text;
  el.className = `message ${kind}`;
  el.hidden = !text;
}

// Calls a Supabase Edge Function and returns { data, error } where error is the function's own
// message (e.g. "Azure isn't set up...") rather than supabase-js's generic HTTP error text.
// While it runs, the page's loading indicator is shown.
async function callApi(name, options) {
  trackLoading(1);
  try {
    const { data, error } = await sb.functions.invoke(name, options);
    if (!error) return { data, error: null };
    let message = error.message;
    try {
      const body = await error.context.json();
      if (body?.error) message = body.error;
    } catch {
      // Network failure / non-JSON response: keep supabase-js's message.
    }
    return { data: null, error: message };
  } finally {
    trackLoading(-1);
  }
}

// --- Loading indicator: a "Loading…" pill shown while any callApi request is in flight. It appears only
// after a short delay so quick requests don't flicker. ---

let pendingRequests = 0;
let loadingTimer = null;

function trackLoading(delta) {
  pendingRequests = Math.max(0, pendingRequests + delta);
  let indicator = document.getElementById("loading");
  if (!indicator) {
    indicator = document.createElement("div");
    indicator.id = "loading";
    indicator.className = "loading";
    indicator.setAttribute("role", "status");
    indicator.hidden = true;
    const spinner = document.createElement("span");
    spinner.className = "spinner";
    indicator.append(spinner, "Loading…");
    document.body.append(indicator);
  }
  clearTimeout(loadingTimer);
  if (pendingRequests > 0) loadingTimer = setTimeout(() => (indicator.hidden = false), 150);
  else indicator.hidden = true;
}

// --- UI logging (stored by the `logs` Edge Function as source 'ui') ---

let loggingInFlight = false;

async function logUi(level, message, context = {}) {
  console[level === "warn" ? "warn" : "error"]("[ui]", message, context);
  // Signed-out pages can't write logs; also never log the logger's own failures.
  if (loggingInFlight || !(await getSession())) return;
  loggingInFlight = true;
  try {
    await sb.functions.invoke("logs", {
      body: { level, message: String(message), location: location.pathname, context: { ...context, userAgent: navigator.userAgent } },
    });
  } catch {
    // Logging is best effort.
  } finally {
    loggingInFlight = false;
  }
}

window.addEventListener("error", (e) => {
  logUi("error", e.message || "Script error", { file: e.filename, line: e.lineno, column: e.colno, stack: e.error?.stack });
});

window.addEventListener("unhandledrejection", (e) => {
  const reason = e.reason;
  logUi("error", reason?.message || String(reason), { stack: reason?.stack, kind: "unhandledrejection" });
});

// --- Side menu shared by signed-in pages. Expects <aside id="sidebar"> in the page. ---

const NAV = [
  { page: "dashboard", label: "Dashboard" },
  { page: "inbox", label: "Inbox" },
  { page: "logs", label: "Logs" },
];

function renderSidebar(active, session) {
  const aside = document.getElementById("sidebar");
  const brand = document.createElement("div");
  brand.className = "brand";
  const logo = document.createElement("img");
  logo.src = "images/logo-256.png";
  logo.alt = "";
  brand.append(logo, BRAND.appName ?? "");

  const nav = document.createElement("nav");
  for (const item of NAV) {
    const a = document.createElement("a");
    a.href = pageUrl(item.page);
    a.textContent = item.label;
    if (item.page === active) a.className = "active";
    nav.append(a);
  }

  const foot = document.createElement("div");
  foot.className = "side-foot";
  const email = document.createElement("div");
  email.className = "side-email";
  email.textContent = session.user.email;
  const logout = document.createElement("button");
  logout.className = "secondary";
  logout.textContent = "Log out";
  logout.addEventListener("click", () => sb.auth.signOut());
  const legal = document.createElement("div");
  legal.className = "side-legal";
  for (const [page, label] of [["terms", "Terms"], ["privacy", "Privacy"]]) {
    const a = document.createElement("a");
    a.href = pageUrl(page);
    a.textContent = label;
    legal.append(a);
  }
  foot.append(email, logout, legal);

  aside.replaceChildren(brand, nav, foot);
}

// Disables the form while it submits; the `busy` class puts a spinner in its submit button.
function setBusy(form, busy) {
  for (const el of form.elements) el.disabled = busy;
  form.classList.toggle("busy", busy);
}
