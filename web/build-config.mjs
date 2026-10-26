// Generates web/config.js from web/.env so the static dashboard can read its settings.
import { existsSync, writeFileSync } from "node:fs";

const envFile = new URL(".env", import.meta.url);
if (existsSync(envFile)) process.loadEnvFile(envFile);

const config = {
  SUPABASE_URL: process.env.SUPABASE_URL ?? "",
  SUPABASE_ANON_KEY: process.env.SUPABASE_ANON_KEY ?? "",
};
const missing = Object.entries(config).filter(([, v]) => !v || v.includes("<")).map(([k]) => k);
if (missing.length) console.warn(`web/.env: fill in ${missing.join(", ")}`);

writeFileSync(
  new URL("config.js", import.meta.url),
  `// Generated from web/.env by web/build-config.mjs - do not edit.\nwindow.APP_CONFIG = ${JSON.stringify(config, null, 2)};\n`,
);
console.log("Wrote web/config.js");
