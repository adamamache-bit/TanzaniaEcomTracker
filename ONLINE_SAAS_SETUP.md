# TanzaniaEcomTracker Online SaaS Setup

## 1. Create Supabase project
- Create a new Supabase project.
- Open the SQL editor.
- Run [`supabase/schema.sql`](C:\Users\Adam\Desktop\TanzaniaEcomTracker\supabase\schema.sql).

## 2. Create users
- In Supabase Auth, create your account and your cofounder's account.
- After each user exists, insert membership rows in `workspace_members` for `main-workspace`.
- Give both of you role `owner` or `admin`.

## 3. Add environment variables
Set these variables in local `.env` and in Vercel:

```env
VITE_SUPABASE_URL=...
VITE_SUPABASE_ANON_KEY=...
VITE_SUPABASE_WORKSPACE_ID=main-workspace
```

## 4. Result
- Both users open the same online app URL.
- Both sign in.
- Both read and write the same workspace.
- Changes sync in real time through Supabase Realtime.

