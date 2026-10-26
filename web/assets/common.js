// Shared helpers for every page. Requires the supabase-js UMD bundle and config.js to load first.
export const sb = window.supabase.createClient(window.APP_CONFIG.SUPABASE_URL, window.APP_CONFIG.SUPABASE_ANON_KEY);

export const $ = (id) => document.getElementById(id);
export const page = (name) => new URL(name, location.href).href;

export const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

export const money = (n, cur) => n == null ? "" : new Intl.NumberFormat("en-US", {
  style: "currency", currency: /^[A-Z]{3}$/.test(cur ?? "") ? cur : "USD",
}).format(n);

export const fmtDate = (s) => s ? new Date(s).toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" }) : "";

export const PROVIDER_LABEL = {
  microsoft: "Microsoft 365 / Outlook",
  microsoft_app: "Shared Microsoft 365 mailbox",
  gmail: "Gmail / Google Workspace",
  imap: "Other (IMAP)",
};

export function toast(msg) {
  let t = $("toast");
  if (!t) {
    t = document.createElement("div");
    t.id = "toast"; t.setAttribute("role", "status");
    document.body.append(t);
  }
  t.textContent = msg; t.classList.remove("hidden");
  clearTimeout(toast.timer); toast.timer = setTimeout(() => t.classList.add("hidden"), 6000);
}

/** POST to an edge function as the signed-in user. */
export async function callFn(name, body) {
  const { data: { session } } = await sb.auth.getSession();
  const res = await fetch(`${window.APP_CONFIG.SUPABASE_URL}/functions/v1/${name}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${session?.access_token ?? ""}`,
      apikey: window.APP_CONFIG.SUPABASE_ANON_KEY,
    },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error ?? `Request failed (${res.status})`);
  return data;
}

/** Return the session, or send the visitor to the login page. */
export async function requireSession() {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) {
    location.replace(page(`login.html?next=${encodeURIComponent(location.pathname.split("/").pop() + location.search)}`));
    return new Promise(() => {}); // never resolves; the page is navigating away
  }
  sb.auth.onAuthStateChange((event) => { if (event === "SIGNED_OUT") location.replace(page("login.html")); });
  return session;
}

/** Fill <header id="appHeader"> with the brand, navigation and the signed-in user. */
export function renderHeader(active, session) {
  const header = $("appHeader");
  header.className = "app-header";
  header.innerHTML = `
    <a class="brand" href="index.html">AP Inbox</a>
    <nav>
      <a href="index.html" ${active === "dashboard" ? 'aria-current="page"' : ""}>Action items</a>
      <a href="connect.html" ${active === "connect" ? 'aria-current="page"' : ""}>Mailboxes</a>
      <span class="who">${esc(session.user.email)}</span>
      <button id="signOut">Sign out</button>
    </nav>`;
  $("signOut").onclick = () => sb.auth.signOut();
}
