# SummarizeMail

UI on **Vercel** (static pages), API on **Supabase** (Auth, Postgres, Edge Functions).
Users sign up / log in with email + password and land on a dashboard that calls the `profile` Edge Function.

```
vercel.json                           Vercel build: generates web/config.js, serves web/ with clean URLs
supabase/
  config.toml                         local Supabase settings
  migrations/*_profiles.sql           profiles table + RLS + trigger that creates a row on signup
  migrations/*_mailboxes.sql          connected mailboxes; OAuth refresh tokens stored in Vault
  functions/profile/index.ts          API: GET / PATCH the signed-in user's profile
  functions/mailboxes/index.ts        API: list / connect (Microsoft OAuth) / disconnect Azure mailboxes
  migrations/*_app_logs.sql           error log for UI and API, shown on the Logs page
  functions/logs/index.ts             API: list / write (UI errors) / clear the caller's logs
  functions/inbox/index.ts            API: process a mailbox (Claude classification + Smart AP forwarding), list / update results
  functions/_shared/classify.ts       Claude call that classifies one AP email
  functions/_shared/microsoft.ts      Microsoft Graph: token refresh, list inbox, forward
  migrations/*_ap_processing.sql      processed_emails table + Smart AP address on mailboxes
  functions/_shared/log.ts            logApi() / fail(): Edge Function errors -> app_logs
  functions/_shared/cors.ts           CORS headers (UI and API are on different origins)
web/
  index.html                          redirects to dashboard or login
  login.html / signup.html            auth forms
  dashboard.html                      protected page, loads profile via the Edge Function
  inbox.html                          processed AP emails, next steps, "Process inbox" button
  logs.html                           UI + API error log (side menu > Logs)
  assets/common.js                    Supabase client + session helpers
  build-config.mjs                    writes web/config.js from env vars (web/.env locally, Vercel env in prod)
```

## 1. Supabase (API)

1. Create a project at https://supabase.com/dashboard and note its project ref.
2. Link, apply the migration, deploy the function:
   ```
   npx supabase login
   npm run db:link -- --project-ref <project-ref>
   npm run db:push
   npm run functions:deploy
   ```
   `SUPABASE_URL` and `SUPABASE_ANON_KEY` are provided to Edge Functions automatically.
3. Authentication > URL Configuration:
   - **Site URL**: `https://<your-app>.vercel.app`
   - **Redirect URLs**: `https://<your-app>.vercel.app/**`, `https://*-<team>.vercel.app/**` (previews), `http://localhost:5500/**`

## 1b. Azure mailbox connection

1. Azure portal > Microsoft Entra ID > App registrations > New registration.
   - Supported account types: pick "any organizational directory and personal accounts" (or single tenant).
   - Redirect URI (Web): `https://<project-ref>.supabase.co/functions/v1/mailboxes`
2. Certificates & secrets > New client secret. Copy the **value**.
3. API permissions > Microsoft Graph > Delegated: `offline_access`, `User.Read`, `Mail.Read`, `Mail.Send`
   (`Mail.Send` forwards invoices; mailboxes connected before it was added must be reconnected).
4. Copy `supabase/functions/.env.example` to `supabase/functions/.env`, fill in `MS_CLIENT_ID`,
   `MS_CLIENT_SECRET`, `MS_TENANT_ID`, `ANTHROPIC_API_KEY`, `OAUTH_STATE_SECRET` and `APP_URL`, then:
   ```
   npm run secrets
   npm run functions:deploy
   ```
On the dashboard, **Add Azure mailbox** opens Microsoft sign-in in a popup (or, if popups are blocked, sends the
user there and back to `/dashboard`).

## 1c. Gmail mailbox connection (optional)

1. Google Cloud console > create or pick a project > APIs & Services > Library > enable **Gmail API**.
2. OAuth consent screen: External (or Internal for a Google Workspace org), add the scopes
   `https://www.googleapis.com/auth/gmail.readonly` and `https://www.googleapis.com/auth/gmail.send`.
   - While the app is in **Testing**, only the listed test users can connect, and Google expires their access
     after 7 days (reconnect from the dashboard). `gmail.readonly` is a restricted scope, so publishing the app
     for everyone needs Google's verification and a security assessment. An **Internal** app in a Workspace
     org avoids both.
3. Credentials > Create credentials > OAuth client ID > Web application. Authorized redirect URI:
   `https://<project-ref>.supabase.co/functions/v1/mailboxes`
4. Put the client ID and secret in `supabase/functions/.env` as `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET`, then
   `npm run secrets`, `npm run db:push` and `npm run functions:deploy`.

## 2. Vercel (UI)

1. Import the Git repo in Vercel. Leave Root Directory as the repo root; `vercel.json` sets the build.
2. Project Settings > Environment Variables (Production + Preview):
   - `SUPABASE_URL` = `https://<project-ref>.supabase.co`
   - `SUPABASE_ANON_KEY` = anon / publishable key from Supabase > Project Settings > API
3. Redeploy. The build fails with a clear message if either variable is missing.

## Local development

```
cp web/.env.example web/.env     # fill in URL + anon key
npm run dev                      # http://localhost:5500
```
With Docker installed you can run everything locally instead: `npm run db:start` (prints local URL + anon key)
and `npm run functions:serve`.

Email confirmation is on by default for hosted projects: after signing up, click the email link to reach the
dashboard. Turn it off under Authentication > Providers > Email to sign in immediately.
