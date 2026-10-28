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
   For shared mailboxes also add `Mail.Read.Shared` and `Mail.Send.Shared`.
4. Copy `supabase/functions/.env.example` to `supabase/functions/.env`, fill in `MS_CLIENT_ID`,
   `MS_CLIENT_SECRET`, `MS_TENANT_ID`, `ANTHROPIC_API_KEY`, `OAUTH_STATE_SECRET` and `APP_URL`, then:
   ```
   npm run secrets
   npm run functions:deploy
   ```
On the dashboard, **Add Azure mailbox** opens Microsoft sign-in in a popup (or, if popups are blocked, sends the
user there and back to `/dashboard`).

**Add shared mailbox** asks for a Microsoft 365 shared mailbox address (e.g. `ap@company.com`); the user then signs
in with their own account, which needs Full Access to that mailbox in Exchange. The app reads the shared inbox
through that account. Invoices are forwarded from the shared mailbox (the account also needs Send As or Send on
Behalf rights on it); the summary email is sent from the user's own mailbox to the shared one.

### Organizations (API keys) for many shared AP / AR mailboxes

**Add organization** connects a customer's Microsoft 365 tenant once with its own Azure app registration
(app-only / client credentials); its shared mailboxes are then added by address, many at a time, with no one signing in.

1. In the customer's tenant: App registrations > New registration (single tenant, no redirect URI).
2. API permissions > Microsoft Graph > **Application**: `Mail.Read` and `Mail.Send`, then **Grant admin consent**.
3. Certificates & secrets > New client secret.
4. Recommended: restrict the app to the AP/AR mailboxes only. Put them in a mail-enabled security group and run (Exchange
   Online PowerShell) `New-ApplicationAccessPolicy -AppId <client id> -PolicyScopeGroupId <group> -AccessRight RestrictAccess`,
   or use RBAC for Applications. Otherwise these permissions cover every mailbox in the tenant.
5. On the dashboard, **Add organization** with the tenant ID, client ID and secret. The app checks the keys and that
   `Mail.Read` is consented (and warns if `Mail.Send` is missing); the keys are stored in Vault and never returned.
6. **Add mailboxes**: paste the shared mailbox addresses (up to 100 at a time), choose AP or AR, the destination email
   and which categories to forward. Each address is checked for inbox access before it is added.
   **Find shared mailboxes for a person** fills the list with every shared mailbox that person has Full Access to. Graph
   can't read mailbox permissions, so this uses the Exchange Online admin API, which needs two more grants on the app:
   - API permissions > Add > APIs my organization uses > **Office 365 Exchange Online** > Application >
     `Exchange.ManageAsApp`, then **Grant admin consent**.
   - Entra ID > Roles and administrators > **Global Reader** (read-only) or **Exchange Recipient Administrator** >
     Add assignment > the app.

   It finds Full Access granted directly to the person, not through a group, and checks up to 500 shared mailboxes.
7. When the client secret expires, use **Update secret** on the organization.

Each mailbox has its own settings on the dashboard:

- **Type**: AP (invoice, utility bill, vendor info, statement) or AR (remittance, purchase order, customer inquiry).
  Claude classifies with that type's categories.
- **Forward automatically**: which categories are forwarded to the destination. Defaults: AP invoices, AR remittances.
  Anything not forwarded (or low-confidence) is listed as "action needed" with a next step.
- **Destination email**: where forwarded emails go, e.g. the AP system's invoice inbox or the cash-application team.

Forwarding and the summary email are sent from the mailbox itself. The scheduler processes three mailboxes at a time,
so dozens of mailboxes on the same schedule are worked through over a few 5-minute ticks.

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
