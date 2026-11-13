# Lumo — family app (upload version)

All files sit in ONE folder. No sub-folders, no npm, no build tools.

## Upload to GitHub
Repository → **Add file → Upload files** → select ALL files in this folder → drag them in → **Commit changes**.
Files left over from an older upload don't matter: they're ignored.

## Vercel
Vercel redeploys by itself after each commit. It needs two Environment Variables (Settings → Environment Variables):
- `VITE_SUPABASE_URL` = your Project URL (https://xxxx.supabase.co)
- `VITE_SUPABASE_ANON_KEY` = your anon / publishable key

If you change them: Deployments → ⋯ → Redeploy.
The build log should say **"Lumo built OK"**. If it says **MISSING FILES**, upload the files it names.

## Database
`schema.sql` is the Supabase setup script (SQL Editor → New query → paste → Run). Only needed once.

## Updates
Upload the changed files the same way. Phones get the new version next time the app is opened.

## Automatic daily backups

**A. Nightly backups inside Supabase (2 minutes)**
1. Supabase → **SQL Editor → New query**. Paste all of `backup-setup.sql` and press **Run**.
2. Open the app → **More → Backup**. "Automatic daily backups" shows today's backup. The last 30 days are kept. Tap **Restore** on any day to go back to it.

**B. Off-site copy in GitHub (10 minutes)**
1. GitHub repo → **Add file → Create new file**. Name it exactly `.github/workflows/daily-backup.yml` and paste in the contents of `daily-backup.yml`. Commit.
2. Repo **Settings → Secrets and variables → Actions → New repository secret**. Add three:
   - `SUPABASE_URL`: the same URL as in Vercel
   - `SUPABASE_SECRET_KEY`: Supabase → Project Settings → API Keys → **Secret keys** (starts `sb_secret_`). Never put this one in Vercel or the app.
   - `BACKUP_PASSWORD`: a long password you make up. **Write it down.** Without it the GitHub backups can't be opened.
3. **Actions → Daily backup → Run workflow** to test it. A green tick means it worked. It then runs every night and keeps 90 days.
4. To restore: open a run and download the backup at the bottom. In the app, go to **Backup → Restore from backup**, pick the .zip and enter the password.
