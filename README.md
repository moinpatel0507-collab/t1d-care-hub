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

## Android app (free, for testing)
1. Go to **pwabuilder.com**, paste your app's web address, and press **Start**.
2. Press **Package for stores → Android → Generate package**. Keep the default settings and download the zip.
3. Upload the zip's `assetlinks.json` to the top level of this GitHub repo. Vercel redeploys on its own, and the app then opens without a browser address bar.
4. Send the `.apk` from the zip to each Android phone, open it and allow **Install unknown apps** when asked.
5. Keep the zip's `signing.keystore` and `signing-key-info.txt` safe. You need them to make updated versions of the app.

## Version 2: phone reminders (10 minutes)
1. Open the app → **More → Setup check → Turn on notifications on this phone**. The first time, it shows a secret key once. Copy it.
2. GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**. Name `VAPID_PRIVATE_KEY`, value: the key you copied.
3. GitHub repo → **Add file → Create new file**. Name it exactly `.github/workflows/reminders.yml`, paste in the contents of `reminders.yml` and commit.
4. **Actions → Reminders → Run workflow** to test it. Reminders then come about 7–8am and 6–7pm.
5. On his mum's phone: open the app → Setup check → Turn on notifications. She doesn't need a new key.
6. Also replace `.github/workflows/daily-backup.yml` with the new `daily-backup.yml`. It now tells the app when the off-site backup worked, so Setup check can show it.

iPhone: notifications only work when the app has been added to the Home Screen (Safari → Share → Add to Home Screen) and is opened from there.

## Version 7: view-only access for grandparents (5 minutes)
1. Supabase → **SQL Editor → New query**. Paste the whole of `viewer-setup.sql` and press **Run**. Safe to run again.
2. In the app: **Care plan → Family who can view → Make a family view code**.
3. New sign-ups are closed, so make their login yourself: Supabase → **Authentication → Users → Add user → Create new user**. Enter their email and a password, tick **Auto confirm user**.
4. Send them the app link, that email and password, and the code (**Share instructions** writes the message).
5. They sign in, type their name and tap **Join to view**. Then switch the code off.

They see the last 24 hours, what to do if he's low, and his care sheet. They can't add, change or delete anything, and can't see older history, glucose files, reports or backups. Remove them any time from the same panel.

## Version 6: instant alerts when the app is closed (15 minutes)
Buzzes the other parents' phones the moment someone logs a hypo or starts sick-day mode.

Before you start: notifications must be switched on in the app on each phone (**Setup check → Turn on notifications**), and you need the **private key** the app showed you then. If you've lost it, turn notifications off and on again in the app to make a new pair, and update the GitHub secret too.

1. Supabase → **Edge Functions → Deploy a new function → Via Editor**. Start from scratch, name it `parent-alert`, delete any example code, paste the whole of `parent-alert.ts`, and click **Deploy function**.
2. Open the function's **Details** tab and switch **JWT verification off**, then save. The function checks its own secret instead.
3. **Edge Functions → Secrets** (or Project Settings → Edge Functions): add
   - `VAPID_PRIVATE_KEY` = the private key from the app
   - `ALERT_SECRET` = any long random text you make up (e.g. 30 letters and numbers)
4. Open `alerts-setup.sql`, change the two lines marked **EDIT** (your Project URL, and the same ALERT_SECRET text), then paste it into **SQL Editor** and **Run**.
5. Test: on your phone, close the app. On the other phone, open **He's low** and log a hypo. Your phone should buzz within a few seconds. Delete the test entry afterwards.

If nothing arrives: Edge Functions → parent-alert → **Logs** shows why (e.g. "notifications not turned on in the app" or a missing secret).
