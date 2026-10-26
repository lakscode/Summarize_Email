// Shared Supabase client and auth helpers for every page.
// Load order: supabase-js UMD, config.js, then this file.

const { SUPABASE_URL, SUPABASE_ANON_KEY } = window.APP_CONFIG ?? {};
if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
  document.addEventListener("DOMContentLoaded", () =>
    showMessage("Missing Supabase settings. Fill in web/.env and run `npm run config`.", "error"),
  );
}

const sb = window.supabase.createClient(SUPABASE_URL || "http://invalid", SUPABASE_ANON_KEY || "missing");

const pageUrl = (name) => new URL(`${name}.html`, location.href).href;

async function getSession() {
  const { data } = await sb.auth.getSession();
  return data.session;
}

// Login/signup pages: skip straight to the dashboard when already signed in.
async function redirectIfSignedIn() {
  if (await getSession()) location.replace(pageUrl("dashboard"));
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
async function callApi(name, options) {
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
  { page: "logs", label: "Logs" },
];

function renderSidebar(active, session) {
  const aside = document.getElementById("sidebar");
  const brand = document.createElement("div");
  brand.className = "brand";
  brand.textContent = "SummarizeMail";

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
  foot.append(email, logout);

  aside.replaceChildren(brand, nav, foot);
}

function setBusy(form, busy) {
  for (const el of form.elements) el.disabled = busy;
}
