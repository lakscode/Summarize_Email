// Writes Edge Function errors to public.app_logs (source 'api') so they show on the Logs page.
// Logging never throws: a failure to log must not mask the original error.
import { json } from "./cors.ts";
import { adminClient } from "./supabase.ts";

interface LogOptions {
  userId?: string | null;
  level?: "error" | "warn" | "info";
  context?: Record<string, unknown>;
}

export async function logApi(location: string, error: unknown, opts: LogOptions = {}): Promise<void> {
  const message = error instanceof Error ? error.message : String(error);
  const stack = error instanceof Error ? error.stack?.slice(0, 4000) : undefined;
  console.error(`[${location}]`, message, opts.context ?? "");
  try {
    const { error: insertError } = await adminClient().from("app_logs").insert({
      user_id: opts.userId ?? null,
      source: "api",
      level: opts.level ?? "error",
      location,
      message: message.slice(0, 2000),
      context: { ...opts.context, ...(stack ? { stack } : {}) },
    });
    if (insertError) console.error("Could not write app log:", insertError.message);
  } catch (err) {
    console.error("Could not write app log:", err);
  }
}

// Log, then return the JSON error response the UI will show.
export async function fail(location: string, message: string, status: number, opts: LogOptions = {}): Promise<Response> {
  await logApi(location, message, { level: status >= 500 ? "error" : "warn", ...opts });
  return json({ error: message }, status);
}
