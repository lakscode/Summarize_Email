// Branded HTML layout for every email the app sends. Edit the markup below to change the look of all emails.
// Table-based with inline styles so it renders in Outlook, Gmail and mobile clients.
//
// Placeholders (filled by renderEmail):
//   {{title}}         <title> and the subject shown by some clients
//   {{preheader}}     hidden preview text shown next to the subject in the inbox list
//   {{heading}}       large heading at the top of the card
//   {{intro}}         paragraph under the heading (HTML allowed)
//   {{button_url}}    / {{button_label}}  call-to-action button; the whole <!--button--> block is removed when empty
//   {{content}}       main body (HTML)
//   {{footer}}        small print at the bottom
//   {{app_name}} / {{app_tagline}}  brand bar text, from brand.json
import { APP_NAME, APP_TAGLINE } from "./brand.ts";

export const EMAIL_TEMPLATE = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>{{title}}</title>
</head>
<body style="margin:0;padding:0;background:#eef2fb;">
  <div style="display:none;max-height:0;overflow:hidden;opacity:0;color:transparent;">{{preheader}}</div>
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#eef2fb;">
    <tr>
      <td align="center" style="padding:28px 12px;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:880px;">
          <!-- Brand bar -->
          <tr>
            <td style="background:#3b5bdb;background-image:linear-gradient(90deg,#3b5bdb,#6c8cff);border-radius:12px 12px 0 0;padding:18px 28px;">
              <span style="font-family:'Segoe UI',Arial,sans-serif;font-size:18px;font-weight:700;color:#ffffff;letter-spacing:0.2px;">&#9993; {{app_name}}</span>
              <span style="font-family:'Segoe UI',Arial,sans-serif;font-size:13px;color:#dfe6ff;padding-left:10px;">{{app_tagline}}</span>
            </td>
          </tr>
          <!-- Card -->
          <tr>
            <td style="background:#ffffff;padding:28px;border-radius:0 0 12px 12px;font-family:'Segoe UI',Arial,sans-serif;color:#1c2230;">
              <h1 style="margin:0 0 8px;font-size:22px;line-height:1.3;color:#1c2230;">{{heading}}</h1>
              <div style="margin:0 0 18px;font-size:15px;line-height:1.5;color:#3d4757;">{{intro}}</div>
              <!--button-->
              <table role="presentation" cellspacing="0" cellpadding="0" style="margin:0 0 22px;">
                <tr>
                  <td style="background:#3b5bdb;border-radius:8px;">
                    <a href="{{button_url}}" style="display:inline-block;padding:11px 20px;font-family:'Segoe UI',Arial,sans-serif;font-size:14px;font-weight:600;color:#ffffff;text-decoration:none;">{{button_label}}</a>
                  </td>
                </tr>
              </table>
              <!--/button-->
              {{content}}
            </td>
          </tr>
          <!-- Footer -->
          <tr>
            <td style="padding:16px 28px;font-family:'Segoe UI',Arial,sans-serif;font-size:12px;line-height:1.5;color:#8a94a3;text-align:center;">
              {{footer}}
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

export interface EmailParts {
  title: string;
  preheader: string;
  heading: string;
  intro: string; // HTML
  content: string; // HTML
  footer: string; // HTML
  button?: { url: string; label: string } | null;
}

export const escapeHtml = (s: string | null | undefined) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

// Fills the template. Plain-text fields are escaped; intro / content / footer are trusted HTML built by the caller.
export function renderEmail(p: EmailParts): string {
  let html = EMAIL_TEMPLATE;
  if (!p.button?.url) html = html.replace(/<!--button-->[\s\S]*?<!--\/button-->/, "");
  const values: Record<string, string> = {
    title: escapeHtml(p.title),
    preheader: escapeHtml(p.preheader),
    heading: escapeHtml(p.heading),
    intro: p.intro,
    content: p.content,
    footer: p.footer,
    button_url: escapeHtml(p.button?.url),
    button_label: escapeHtml(p.button?.label),
    app_name: escapeHtml(APP_NAME),
    app_tagline: escapeHtml(APP_TAGLINE),
  };
  // Single pass, so placeholder-like text inside inserted values is never expanded again.
  return html.replace(/\{\{(\w+)\}\}/g, (match, key) => (key in values ? values[key] : match));
}
