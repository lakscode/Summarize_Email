# SummarizeMail

Static web app with Supabase email/password auth: sign up, log in, land on a dashboard.

```
supabase/
  config.toml                         local Supabase settings (auth redirect URLs -> localhost:5500)
  migrations/*_profiles.sql           profiles table + RLS + trigger that creates a row on signup
web/
  index.html                          redirects to dashboard or login
  login.html / signup.html            auth forms
  dashboard.html                      protected page, shows the user's profile
  assets/common.js                    Supabase client + session helpers
  build-config.mjs                    writes web/config.js from web/.env
```

## Setup (hosted Supabase project)

1. Create a project at https://supabase.com/dashboard.
2. Copy `web/.env.example` to `web/.env` and fill in `SUPABASE_URL` and `SUPABASE_ANON_KEY`
   (Project Settings > API).
3. Apply the migration:
   ```
   npx supabase login
   npm run db:link -- --project-ref <project-ref>
   npm run db:push
   ```
4. In the dashboard, open Authentication > URL Configuration. Set **Site URL** to `http://localhost:5500`
   and add `http://localhost:5500/**` to **Redirect URLs**.
5. `npm run dev` and open http://localhost:5500.

Email confirmation is on by default for hosted projects: after signing up, click the link in the email
and you'll be taken to the dashboard. Turn it off under Authentication > Providers > Email to sign in immediately.

## Setup (local Supabase, needs Docker)

```
npm run db:start          # prints API URL and anon key; applies migrations
```
Put the printed API URL and anon key into `web/.env`, then `npm run dev`. Email confirmation is off locally.
