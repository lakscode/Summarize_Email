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
// A page link that carries ?details: without ".html", because the local server (npx serve) answers "x.html?..." with
// a redirect to "x" that drops them (Vercel keeps them). Both serve "x" directly.
const pageLink = (name, params) => `${new URL(name, location.href).href}?${new URLSearchParams(params)}`;

async function getSession() {
  const { data } = await sb.auth.getSession();
  return data.session;
}

// Login/signup pages: skip straight to the dashboard when already signed in.
async function redirectIfSignedIn() {
  if (await getSession()) location.replace(pageUrl("inbox"));
}

// --- Single sign-on (Supabase Auth OAuth providers; enable them under Authentication > Sign In / Providers). ---

// Login and sign-up pages: false shows only "Continue with Microsoft" (the first Microsoft sign-in creates the
// account); true also shows the email / password forms, forgot password and the sign-up link.
// UI only: email sign-in stays enabled in Supabase.
const SHOW_EMAIL_LOGIN = false;

const SSO_PROVIDERS = [
  // Azure only returns the user's email when the `email` scope is requested.
  { provider: "azure", label: "Continue with Microsoft", scopes: "email" },
];

// Fills <div id="sso"> with one button per provider. The provider sends the user back to the login page,
// which shows any error and otherwise forwards to the dashboard. New users get an account on first sign-in.
// With only: true (no other form on the page) the buttons are primary and there is no "or" divider.
function renderSsoButtons({ only = false } = {}) {
  const box = document.getElementById("sso");
  if (!box) return;
  const buttons = SSO_PROVIDERS.map(({ provider, label, scopes }) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = only ? "sso-button" : "sso-button secondary";
    // Microsoft's four-square logo, then the label.
    if (provider === "azure") {
      button.insertAdjacentHTML("beforeend",
        '<svg class="sso-logo" viewBox="0 0 21 21" aria-hidden="true"><rect x="1" y="1" width="9" height="9" fill="#f25022"/>' +
        '<rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>' +
        '<rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>');
    }
    const text = document.createTextNode(label);
    button.append(text);
    // While the provider's page loads: spinner and "Opening … sign-in…", so nobody clicks again.
    button.addEventListener("click", async () => {
      showMessage("");
      for (const b of buttons) b.disabled = true;
      button.classList.add("busy");
      text.data = `Opening ${label.replace(/^Continue with /, "")} sign-in…`;
      const { error } = await sb.auth.signInWithOAuth({ provider, options: { redirectTo: pageUrl("login"), scopes } });
      if (error) {
        for (const b of buttons) b.disabled = false;
        button.classList.remove("busy");
        text.data = label;
        showMessage(error.message, "error");
      }
    });
    return button;
  });
  const divider = document.createElement("div");
  divider.className = "divider";
  divider.textContent = "or";
  box.replaceChildren(...buttons, ...(only ? [] : [divider]));
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
  await joinOrganization(session);
  return session;
}

// Once per sign-in, on whichever page comes first: loading the organization joins one by invitation or by the
// user's email domain (supabase/functions/organizations), so a new user lands in their company's organization and sees
// its mailboxes without visiting Settings first.
async function joinOrganization(session) {
  const key = `org-joined:${session.user.id}:${session.user.last_sign_in_at ?? ""}`;
  try {
    if (sessionStorage.getItem(key)) return;
  } catch {
    // Storage blocked: check on every page.
  }
  const { error } = await callApi("organizations", { method: "GET" });
  if (error) return;
  try {
    sessionStorage.setItem(key, "1");
  } catch {
    // Not remembered; checked again on the next page.
  }
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

// Icon paths are 24x24 stroked outlines; shown alone when the sidebar is collapsed.
const ICONS = {
  dashboard: '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1L7 17M17 7l2.1-2.1"/>',
  inbox: '<path d="M22 12h-6l-2 3h-4l-2-3H2"/><path d="M5.5 5h13L22 12v6a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2v-6z"/>',
  logs: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
  reports: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h5"/>',
  profile: '<circle cx="12" cy="8" r="4"/><path d="M4 21v-1a6 6 0 0 1 6-6h4a6 6 0 0 1 6 6v1"/>',
  settings: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
  bell: '<path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/>',
  logout: '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/>',
  collapse: '<path d="M15 18l-6-6 6-6"/>',
};

// Daily work first, then setup, admin, troubleshooting, and the user's own account last.
const NAV = [
  { page: "inbox", label: "Inbox" },
  // Mailbox setup: connecting mailboxes and how each one is handled.
  { page: "dashboard", label: "Settings" },
  // Organization details for everyone; what can be changed there depends on the user's role.
  { page: "settings", label: "Organization" },
  // History of the forwarded invoices reports.
  { page: "reports", label: "Reports" },
  { page: "logs", label: "Logs" },
  { page: "profile", label: "Profile" },
];


function icon(name) {
  const span = document.createElement("span");
  span.className = "icon";
  span.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name]}</svg>`;
  return span;
}

function labelled(iconName, text) {
  const label = document.createElement("span");
  label.className = "label";
  label.textContent = text;
  return [icon(iconName), label];
}

function setCollapsed(shell, toggle, collapsed) {
  shell.classList.toggle("sidebar-collapsed", collapsed);
  toggle.title = collapsed ? "Expand menu" : "Collapse menu";
  toggle.setAttribute("aria-label", toggle.title);
}

function renderSidebar(active, session) {
  const aside = document.getElementById("sidebar");
  const shell = aside.closest(".shell");
  const brand = document.createElement("div");
  brand.className = "brand";
  const logo = document.createElement("img");
  logo.src = "images/logo-256.png";
  logo.alt = "";
  const name = document.createElement("span");
  name.className = "label";
  name.textContent = BRAND.appName ?? "";
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "side-toggle";
  toggle.append(icon("collapse"));
  toggle.addEventListener("click", () =>
    setCollapsed(shell, toggle, !shell.classList.contains("sidebar-collapsed")));
  // Every page opens with the menu expanded; collapsing lasts until the next page. (Earlier versions remembered it.)
  try { localStorage.removeItem("sidebar-collapsed"); } catch {}
  setCollapsed(shell, toggle, false);
  brand.append(logo, name, notificationBell(), toggle);

  const nav = document.createElement("nav");
  for (const item of NAV) {
    const a = document.createElement("a");
    a.href = pageUrl(item.page);
    a.title = item.label;
    a.append(...labelled(item.page, item.label));
    if (item.page === active) a.className = "active";
    nav.append(a);
  }

  const foot = document.createElement("div");
  foot.className = "side-foot";
  const email = document.createElement("a");
  email.className = "side-email";
  email.href = pageUrl("profile");
  email.title = "Your profile";
  email.textContent = session.user.email;
  const logout = document.createElement("button");
  logout.className = "secondary side-logout";
  logout.title = "Log out";
  logout.append(...labelled("logout", "Log out"));
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
  loadNotifications();
}

// --- Notifications: tasks assigned to the user (inbox/index.ts assignEmail). A bell in the sidebar counts the new
// ones (not yet seen); clicking it lists the open tasks, newest first, and counts as seeing them. ---

let notifications = [];

function notificationBell() {
  const bell = document.createElement("button");
  bell.type = "button";
  bell.id = "notification-bell";
  bell.className = "side-toggle notification-bell";
  bell.title = "Notifications";
  bell.setAttribute("aria-label", "Notifications");
  bell.setAttribute("aria-haspopup", "dialog");
  bell.setAttribute("aria-expanded", "false");
  bell.append(icon("bell"));
  bell.addEventListener("click", (ev) => {
    ev.stopPropagation();
    if (document.getElementById("notification-panel")) closeNotifications();
    else openNotifications(bell);
  });
  return bell;
}

async function loadNotifications() {
  const { data, error } = await callApi("inbox?assigned=me", { method: "GET" });
  if (error || !data) return;
  notifications = data.emails;
  setNotificationCount(data.unseen);
}

// The count on the bell; 0 removes it.
function setNotificationCount(count) {
  const bell = document.getElementById("notification-bell");
  if (!bell) return;
  bell.querySelector(".nav-count")?.remove();
  bell.title = count ? `Notifications: ${count} new` : "Notifications";
  bell.setAttribute("aria-label", bell.title);
  if (!count) return;
  const badge = document.createElement("span");
  badge.className = "nav-count";
  badge.textContent = count > 99 ? "99+" : String(count);
  bell.append(badge);
}

function openNotifications(bell) {
  const panel = document.createElement("div");
  panel.id = "notification-panel";
  panel.className = "notification-panel";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-label", "Notifications");
  panel.addEventListener("click", (ev) => ev.stopPropagation());
  const title = document.createElement("strong");
  title.textContent = "Notifications";
  const head = document.createElement("div");
  head.className = "notification-head";
  head.append(title);

  const list = document.createElement("ul");
  for (const e of notifications.slice(0, 20)) {
    const li = document.createElement("li");
    if (!e.assignment_seen_at) li.className = "unread";
    const a = document.createElement("a");
    // Opens this email in the Inbox: its mailbox selected, the email scrolled to and opened.
    a.href = pageLink("inbox", { assigned: "1", mailbox: e.mailbox_id, email: e.id });
    openTaskOnClick(a, e.mailbox_id, e.id);
    const what = document.createElement("div");
    what.className = "notification-what";
    what.textContent = `${e.assigned_by_name && e.assigned_by_name !== e.assigned_to_name ? `${e.assigned_by_name} assigned you` : "Assigned to you"}: ${e.subject || "(no subject)"}`;
    const when = document.createElement("div");
    when.className = "notification-when";
    when.textContent = e.assigned_at ? new Date(e.assigned_at).toLocaleString([], { dateStyle: "short", timeStyle: "short" }) : "";
    a.append(what, when);
    li.append(a);
    list.append(li);
  }
  if (!notifications.length) {
    const empty = document.createElement("li");
    empty.className = "notification-empty";
    empty.textContent = "No notifications. Tasks assigned to you show here.";
    list.append(empty);
  }
  const all = document.createElement("a");
  all.className = "notification-all";
  // The mailbox of the newest task, as "Assigned to me" lists one mailbox at a time.
  const newest = notifications[0]?.mailbox_id;
  all.href = pageLink("inbox", { assigned: "1", ...(newest ? { mailbox: newest } : {}) });
  openTaskOnClick(all, newest, null);
  all.textContent = notifications.length > 20 ? `View all ${notifications.length} tasks` : "View my tasks";
  panel.append(head, list, all);

  // Below the bell, kept on screen.
  const rect = bell.getBoundingClientRect();
  panel.style.top = `${rect.bottom + 6}px`;
  panel.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - 348))}px`;
  document.body.append(panel);
  bell.setAttribute("aria-expanded", "true");
  document.addEventListener("click", closeNotifications);
  document.addEventListener("keydown", closeOnEscape);
  markNotificationsSeen();
}

// Already on the Inbox: open the task there (inbox.html openTaskInInbox) instead of reloading the page. Elsewhere, or
// with Ctrl / Cmd / Shift held, the link opens normally.
function openTaskOnClick(link, mailboxId, emailId) {
  link.addEventListener("click", (ev) => {
    if (typeof openTaskInInbox !== "function" || ev.ctrlKey || ev.metaKey || ev.shiftKey) return;
    ev.preventDefault();
    closeNotifications();
    openTaskInInbox(mailboxId, emailId);
  });
}

function closeNotifications() {
  document.getElementById("notification-panel")?.remove();
  document.getElementById("notification-bell")?.setAttribute("aria-expanded", "false");
  document.removeEventListener("click", closeNotifications);
  document.removeEventListener("keydown", closeOnEscape);
}

function closeOnEscape(ev) {
  if (ev.key === "Escape") closeNotifications();
}

// Opening the panel (or the Inbox's "Assigned to me" tab) counts as seeing them; the panel keeps them highlighted
// until it closes.
async function markNotificationsSeen() {
  if (!notifications.some((e) => !e.assignment_seen_at)) return;
  const { error } = await callApi("inbox", { method: "POST", body: { action: "seen" } });
  if (error) return;
  const now = new Date().toISOString();
  notifications = notifications.map((e) => ({ ...e, assignment_seen_at: e.assignment_seen_at ?? now }));
  setNotificationCount(0);
}

// Disables the form while it submits; the `busy` class puts a spinner in its submit button.
function setBusy(form, busy) {
  for (const el of form.elements) el.disabled = busy;
  form.classList.toggle("busy", busy);
}
