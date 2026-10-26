import { createClient, type SupabaseClient } from "npm:@supabase/supabase-js@2.118.0";
import { decodeBase64, encodeBase64 } from "jsr:@std/encoding@1.0.10/base64";

export function env(name: string, fallback?: string): string {
  const v = Deno.env.get(name) ?? fallback;
  if (v === undefined || v === "") throw new Error(`Missing environment variable ${name}`);
  return v;
}

export function adminClient(): SupabaseClient {
  return createClient(env("SUPABASE_URL"), env("SUPABASE_SERVICE_ROLE_KEY"), {
    auth: { persistSession: false },
  });
}

/** Resolve the calling user from the Authorization header, or null. */
export async function callerUserId(req: Request): Promise<string | null> {
  const auth = req.headers.get("Authorization") ?? "";
  const token = auth.replace(/^Bearer\s+/i, "");
  if (!token) return null;
  const { data } = await adminClient().auth.getUser(token);
  return data.user?.id ?? null;
}

export const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-cron-secret",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
};

export function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

export const b64 = {
  encode: (bytes: Uint8Array) => encodeBase64(bytes),
  decode: (s: string) => decodeBase64(s),
  decodeUrl: (s: string) => decodeBase64(s.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(s.length / 4) * 4, "=")),
  encodeUrl: (bytes: Uint8Array) => encodeBase64(bytes).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""),
};

export function domainOf(address: string | null | undefined): string {
  const m = (address ?? "").toLowerCase().match(/@([^>\s]+)/);
  return m ? m[1] : "";
}

/** "billing.acme.co.uk" -> "acme.co.uk"; "mail.acme.com" -> "acme.com" (good enough for comparisons). */
export function registrableDomain(host: string): string {
  const parts = host.toLowerCase().replace(/^www\./, "").split(".").filter(Boolean);
  if (parts.length <= 2) return parts.join(".");
  const sld = parts[parts.length - 2];
  const twoPartTld = ["co", "com", "org", "net", "gov", "ac"].includes(sld) && parts[parts.length - 1].length === 2;
  return parts.slice(twoPartTld ? -3 : -2).join(".");
}

export function normalizeInvoiceNumber(n: string | null | undefined): string {
  return (n ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export function htmlToText(html: string): string {
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6])>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+/g, " ")
    .replace(/\n\s*\n\s*\n+/g, "\n\n")
    .trim();
}

export function escapeHtml(s: string | null | undefined): string {
  return (s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Run `fn` over `items` with at most `limit` in flight. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  const workers = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i]);
    }
  });
  await Promise.all(workers);
  return out;
}
