# T1D Care Hub — family app (upload version)

All files sit in ONE folder. No sub-folders, no npm, no build tools.

## Upload to GitHub
Repository → **Add file → Upload files** → select ALL files in this folder → drag them in → **Commit changes**.
Files left over from an older upload don't matter: they're ignored.

## Vercel
Vercel redeploys by itself after each commit. It needs two Environment Variables (Settings → Environment Variables):
- `VITE_SUPABASE_URL` = your Project URL (https://xxxx.supabase.co)
- `VITE_SUPABASE_ANON_KEY` = your anon / publishable key

If you change them: Deployments → ⋯ → Redeploy.
The build log should say **"T1D Care Hub built OK"**. If it says **MISSING FILES**, upload the files it names.

## Database
`schema.sql` is the Supabase setup script (SQL Editor → New query → paste → Run). Only needed once.

## Updates
Upload the changed files the same way. Phones get the new version next time the app is opened.
