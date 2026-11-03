// Base64 / base64url helpers (same API as jsr:@std/encoding, kept local so deploys don't depend on fetching it).

function toBytes(data: string | ArrayBuffer | Uint8Array): Uint8Array {
  if (typeof data === "string") return new TextEncoder().encode(data);
  return data instanceof Uint8Array ? data : new Uint8Array(data);
}

export function encodeBase64(data: string | ArrayBuffer | Uint8Array): string {
  const bytes = toBytes(data);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

export function decodeBase64(b64: string): Uint8Array<ArrayBuffer> {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function encodeBase64Url(data: string | ArrayBuffer | Uint8Array): string {
  return encodeBase64(data).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export function decodeBase64Url(b64url: string): Uint8Array<ArrayBuffer> {
  const b64 = b64url.replace(/-/g, "+").replace(/_/g, "/");
  return decodeBase64(b64 + "=".repeat((4 - (b64.length % 4)) % 4));
}
