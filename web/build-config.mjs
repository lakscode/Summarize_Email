// Generates web/config.js from web/.env so the static pages can read their settings.
import { existsSync, writeFileSync } from "node:fs";

const envFile = new URL(".env", import.meta.url);
if (existsSync(envFile)) process.loadEnvFile(envFile);

const config = {
  SUPABASE_URL: process.env.SUPABASE_URL ?? "",
  SUPABASE_ANON_KEY: process.env.SUPABASE_ANON_KEY ?? "",
};
const missing = Object.entries(config).filter(([, v]) => !v || v.includes("<")).map(([k]) => k);
if (missing.length) {
  // On Vercel these come from Project Settings > Environment Variables; fail rather than ship a blank config.
  if (process.env.VERCEL) {
    console.error(`Missing Vercel environment variables: ${missing.join(", ")}`);
    process.exit(1);
  }
  console.warn(`web/.env: fill in ${missing.join(", ")}`);
}

writeFileSync(
  new URL("config.js", import.meta.url),
  `// Generated from web/.env by web/build-config.mjs - do not edit.\nwindow.APP_CONFIG = ${JSON.stringify(config, null, 2)};\n`,
);
console.log("Wrote web/config.js");
