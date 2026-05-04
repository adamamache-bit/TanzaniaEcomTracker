**Deploy Plan**

Frontend:
- Deploy the Vite app to Vercel.
- Vercel will use [vercel.json](C:\Users\Adam\Desktop\TanzaniaEcomTracker\vercel.json).
- Set these environment variables in Vercel:
  - `VITE_META_API_BASE=https://YOUR-RENDER-BACKEND.onrender.com`
  - `VITE_SHARED_API_BASE=https://YOUR-RENDER-BACKEND.onrender.com`

Backend:
- Deploy the Express API to Render with [render.yaml](C:\Users\Adam\Desktop\TanzaniaEcomTracker\render.yaml).
- The backend exposes:
  - `/api/state`
  - `/api/state/meta`
  - `/api/meta/ad-accounts`
  - `/api/meta/insights`

Important:
- Render's filesystem is ephemeral by default.
- For shared live data, keep the persistent disk enabled as documented by Render:
  [Persistent Disks](https://render.com/docs/disks)
- The current backend stores shared app state in `DATA_DIR/shared-workspace.json`.

Recommended publish order:
1. Deploy the backend on Render.
2. Confirm `https://YOUR-BACKEND.onrender.com/api/state/meta` returns JSON.
3. Deploy the frontend on Vercel.
4. Add both Vercel env vars pointing to the backend URL.
5. Open the Vercel app from two browsers and verify that edits sync automatically.

Production note:
- This deployment is good for a practical shared MVP.
- For a more robust long-term setup, move shared data from the Render disk to a managed database such as Supabase Postgres.
