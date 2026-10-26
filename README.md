# SummarizeMail

UI on **Vercel** (static pages), API on **Supabase** (Auth, Postgres, Edge Functions).
Users sign up / log in with email + password and land on a dashboard that calls the `profile` Edge Function.

```
vercel.json                           Vercel build: generates web/config.js, serves web/ with clean URLs
supabase/
  config.toml                         local Supabase settings
  migrations/*_profiles.sql           profiles table + RLS + trigger that creates a row on signup
  functions/profile/index.ts          API: GET / PATCH the signed-in user's profile
  functions/_shared/cors.ts           CORS headers (UI and API are on different origins)
web/
  index.html                          redirects to dashboard or login
  login.html / signup.html            auth forms
  dashboard.html                      protected page, loads profile via the Edge Function
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
