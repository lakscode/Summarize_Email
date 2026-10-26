# AP Inbox — Supabase app

Reads an accounts-payable mailbox (Microsoft 365 / Azure, Gmail, or any IMAP mailbox), classifies every
email with Claude, carries out the routine step where one is defined, and tells the employee exactly what
to do next — in a dashboard and in one summary email per run.

| Category | How it's recognised | What the app does | What the employee is told |
|---|---|---|---|
| **Invoice** | Invoice PDF/body with number + amount due | Forwards to the Smart AP inbox (if auto-forward is on), skips duplicates already sent; tags `AI-Processed` | "Confirm invoice 483739 ($1,240.00, due 15 Nov) appears in Smart AP" — or, with auto-forward off, a **Send to Smart AP** button |
| **Utility / portal bill** | "Your bill is ready", no invoice attached | Keeps the portal link only if it's on the provider's own domain; leaves the email untagged until done | "Log into the Evergy portal, download the bill for account … and send it to Smart AP" |
| **Vendor info** | W-9, COI, price list, new contact, quote… | Summarises it and the key dates | "Review and keep for later; no invoice to process yet" |
| **Account statement** | Statement of account / aging listing several invoices | Reads every line, reconciles against the AP ledger, then against invoices received by email | "Request these missing invoices from Cook's Pest Control: …" (with a table) |
| **Other** | Past-due, disputes, credit memos, remittances, non-AP | Escalates past-due / disputes / unclear ones | Specific next step, or "No AP action" |

**Fraud checks** run on every email (bank/remittance changes, SPF/DKIM/DMARC failures, reply-to mismatch,
look-alike or free-mail senders, `.zip`/`.html`/macro attachments, portal links to other domains, plus
Claude's own judgement). A flagged email is **never** forwarded automatically; the employee is told to
verify with the vendor using a phone number they already have, and must confirm that before forwarding
it from the dashboard.

## How it fits together

```
pg_cron ──► process-inbox ──fan-out──► process-inbox (one per mailbox, background)
                                          │ list new mail (Graph / Gmail API / IMAP)
                                          │ safety checks  ─┐
                                          │ Claude classify ├─► decide action ─► forward / label
                                          │ reconcile stmts ┘
                                          ├─► emails, statement_lines (Postgres)
                                          └─► summary report email to the mailbox owner
web/index.html ──► connect (OAuth / IMAP) · item-action (Send to Smart AP / Mark done) · tables via RLS
```

- `supabase/migrations/` — schema, RLS, Vault helpers (mailbox credentials are stored in Supabase Vault, never readable by the browser)
- `supabase/functions/process-inbox` — the run: batching, thread de-duplication, report
- `supabase/functions/connect` — Microsoft / Google OAuth, shared M365 mailbox, IMAP
- `supabase/functions/item-action` — dashboard buttons
- `supabase/functions/_shared` — providers, classifier prompt and schema (`classify.ts`), safety checks, reconciliation, report
- `web/` — static UI: login, mailbox setup and action-item pages (any static host)
- `supabase/cron.sql` — the schedule

## Setup

### 1. Supabase project

```bash
npx supabase login
npx supabase link --project-ref <project-ref>
npx supabase db push
```

In **Authentication → URL Configuration**, set the Site URL to where you'll host the dashboard and add
`http://localhost:5500/**` (and your hosted address with `/**`) to the redirect URLs.

Users sign up and sign in with **email and password** only (`signup.html`, `login.html`); there is no Microsoft
or Google login. The Email provider is on by default. Leave **Confirm email** on (Authentication → Sign In /
Providers → Email) so new accounts must verify their address; the login page also offers *Forgot password?*
and a one-time sign-in link. Supabase's built-in mailer is rate-limited — set up custom SMTP under
Authentication → Emails before inviting a team.

Signing in only identifies the user; connecting the mailbox (Microsoft 365, Gmail or IMAP) is a separate,
explicit step on the Mailboxes page.

### 2. Mail provider app registrations (only the ones you need)

The OAuth redirect URI for both is `https://<project-ref>.supabase.co/functions/v1/connect`.

**Microsoft 365 / Azure** — Entra ID → App registrations → New registration (Web platform, the redirect URI above).
- *Signed-in mailbox:* delegated Graph permissions `offline_access`, `User.Read`, `Mail.ReadWrite`, `Mail.Send`.
- *Shared AP mailbox (no sign-in):* application permissions `Mail.ReadWrite`, `Mail.Send` + admin consent, and set
  `MS_TENANT_ID` to your tenant. **Scope the app to the AP mailbox only** with Exchange
  [RBAC for Applications](https://learn.microsoft.com/exchange/permissions-exo/application-rbac) — otherwise it can read every mailbox in the tenant.
- Create a client secret.

**Gmail / Google Workspace** — Google Cloud console: enable the Gmail API, create an OAuth client (Web
application, the redirect URI above). Scopes used: `gmail.modify`, `gmail.send`. For a Workspace domain set
the consent screen to *Internal*; external use of these scopes requires Google verification.

**IMAP** — nothing to register. Use an app password where the provider requires one. SMTP must be on port
465 (Supabase blocks outbound 25 and 587).

### 3. Secrets and functions

There are two `.env` files, one per side of the app (both git-ignored; each has a `.env.example` template):

| File | Used by | Holds |
|---|---|---|
| `supabase/functions/.env` | Edge functions (API) | Anthropic key, Microsoft / Google OAuth credentials, `CRON_SECRET`, `OAUTH_STATE_SECRET`, `APP_URL` |
| `web/.env` | Dashboard (UI) | `SUPABASE_URL`, `SUPABASE_ANON_KEY` (public values only — never put the service-role key here) |

```bash
# fill in supabase/functions/.env, then:
npm run secrets    # supabase secrets set --env-file supabase/functions/.env
npm run deploy     # supabase functions deploy process-inbox connect item-action
```

`CRON_SECRET` and `OAUTH_STATE_SECRET` are any long random strings (e.g. `openssl rand -hex 32`).

### 4. Dashboard

Fill in `web/.env` (Supabase dashboard → Project Settings → API), then:

```bash
npm run dev        # writes web/config.js from web/.env and serves the dashboard on http://localhost:5500
```

The dashboard is a static page and can't read `.env` itself, so `npm run config` turns `web/.env` into
`web/config.js`. Re-run it after editing `web/.env`. To host it elsewhere, run `npm run config`, deploy the
`web/` folder to any static host and set `APP_URL` in `supabase/functions/.env` to that address.

| Page | What it's for |
|---|---|
| `signup.html` | Create an account (name, email, password) |
| `login.html` | Sign in with email and password, reset a forgotten password, or get a one-time sign-in link |
| `connect.html` | **Mailboxes** — one setup page per option: Microsoft 365 / Outlook (Azure), shared Microsoft 365 mailbox, Gmail / Google Workspace, and any other provider over IMAP (with presets for Yahoo, AOL, Zoho, Fastmail, GoDaddy…). Each page explains the steps, the access granted, and what an administrator must configure; options the server isn't configured for are marked *needs admin setup*. |
| `index.html` | **Action items** — run the inbox, work through the to-do list, mailbox settings, AP ledger import |

Sign in, connect the mailbox, then open **Mailbox settings** and set the Smart AP address. Auto-forward is
**off** by default: invoices wait for a one-click **Send to Smart AP** confirmation. Turn it on once you trust
the results. Optionally import an AP ledger CSV (`vendor_name, invoice_number, invoice_date, amount`) so
statements are reconciled against your AP system and not just the mailbox.

### 5. Schedule

Edit and run `supabase/cron.sql` in the SQL editor (default: weekdays 07:00 UTC). **Process inbox now** in
the dashboard runs it on demand.

## Behaviour notes

- **What's read:** Inbox mail received within the look-back window (default 24 h) that isn't already in the
  database or tagged with the processed label. Earlier messages in a thread are skipped in favour of the
  latest one.
- **Report:** one email per run to the mailbox owner (or the configured recipient), titled
  `AP Inbox Action Report – <date> (<N> action items)`, ordered: needs attention → statements → portal bills →
  review → invoices. Items still open from earlier runs are repeated until marked done. No email is sent
  when there was nothing new and nothing open.
- **Labels:** processed mail gets the `AI-Processed` Outlook category / Gmail label / IMAP keyword. Portal
  bills, flagged mail and invoices waiting for confirmation stay untagged until the employee acts.
- **Model:** `claude-opus-5-5` with structured JSON output and the server-side refusal fallback; set
  `CLAUDE_MODEL` to change it. PDFs, images and CSV/text attachments are sent to the model; other types
  (e.g. `.xlsx`) are listed for the employee to check manually.
- **Large inboxes:** each invocation handles `MAX_EMAILS_PER_INVOCATION` (12) emails, three at a time, then
  continues itself in a fresh invocation; the report goes out once the backlog is done.
- **Email content is untrusted:** the classifier is told never to follow instructions inside emails, and
  forwarding / flagging decisions are enforced in code, not left to the model.

## Development

```bash
deno test supabase/functions/tests/      # unit tests for safety checks, portal links, MIME builder
deno check supabase/functions/*/index.ts
ANTHROPIC_API_KEY=... deno run -A supabase/functions/tests/classify_smoke.ts   # live classifier check on sample emails
npx supabase functions serve --env-file supabase/functions/.env   # needs Docker
```
