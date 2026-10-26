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

function showMessage(text, kind = "info") {
  const el = document.getElementById("message");
  if (!el) return;
  el.textContent = text;
  el.className = `message ${kind}`;
  el.hidden = !text;
}

function setBusy(form, busy) {
  for (const el of form.elements) el.disabled = busy;
}
