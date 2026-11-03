// Fills in app and company names from window.APP_CONFIG.BRAND (brand.json, via config.js).
// Load after config.js. Markup: <span data-brand="appName"></span>, <img data-brand-alt="appName">.
// The page <title> holds only the page name; " · <appName>" is appended here.

const BRAND = window.APP_CONFIG?.BRAND ?? {};

if (BRAND.appName) document.title = document.title ? `${document.title} · ${BRAND.appName}` : BRAND.appName;

function applyBrand(root = document) {
  for (const el of root.querySelectorAll("[data-brand]")) el.textContent = BRAND[el.dataset.brand] ?? "";
  for (const el of root.querySelectorAll("[data-brand-alt]")) el.alt = BRAND[el.dataset.brandAlt] ?? "";
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => applyBrand());
else applyBrand();
