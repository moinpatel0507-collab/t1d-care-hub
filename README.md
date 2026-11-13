# T1D Care Hub — family app

A private app for you and his mum: food and drink, CGM graphs (Dexcom, Libre, Glooko/Omnipod, Medtronic CareLink),
supplies, growth, eye and dental check-ups, sick-day guide and backups.
It installs on both phones like a normal app and keeps both of you in sync.

**Stack:** Vite + plain JavaScript, Supabase (login + Postgres database), hosted on Vercel. Free tiers are enough.

---

## 1. Create the database (Supabase) — about 10 minutes

1. Go to supabase.com, sign up, and click **New project**.
   - Name: `t1d-care-hub`
   - Region: **London (eu-west-2)** keeps the data in the UK
   - Save the database password somewhere safe
2. When it has finished setting up, open **SQL Editor → New query**.
3. Open `supabase/schema.sql` from this folder, copy all of it, paste it in, and click **Run**.
   You should see "Success. No rows returned".

## 2. Turn on email sign-in

1. **Authentication → Sign In / Providers → Email**: make sure it is enabled (it is by default).
   Leave "Confirm email" on.
2. **Authentication → URL Configuration → Site URL**: put your app's web address here once you have it
   (step 4). This is where the confirmation email sends people back to.

## 3. Connect the app to the database

1. In Supabase, open **Project Settings → API** (on newer projects: **API Keys**).
2. Copy the **Project URL** and the **anon / publishable** key.
   (This key is meant to be public. The security rules in `schema.sql` are what protect the data.)
3. In this folder, copy `.env.example` to `.env` and paste both values in.

To try it on your computer:

```bash
npm install
npm run dev
```

Open the address it prints (usually http://localhost:5173).

## 4. Put it online (Vercel) — about 10 minutes

1. Put this folder in a **private** GitHub repository.
2. On vercel.com: **Add New → Project →** pick the repository. Vercel detects Vite automatically.
3. Under **Environment Variables**, add `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` with the same values as your `.env`.
4. Click **Deploy**. You get an address like `https://t1d-care-hub-xxxx.vercel.app`.
5. Go back to Supabase step 2 and put that address in **Site URL**.

## 5. Set up your family

The app opens on a **Sign in / Register** screen. It has show-password buttons, a password check when registering,
and **Forgot your password?** which emails a reset link (it brings you back to the app to choose a new password).

1. **You:** open the app, tap **Register**, confirm the email, sign in, enter your name and tap **Create family**.
2. The next screen shows your 8-character **family code** with **Copy** and **Share** buttons. (It's also in **Care plan → Family & account**.)
3. **His mum:** opens the same address, taps **Register**, confirms her email, signs in, enters her name and the code, and taps **Join family** (not Create family).
4. **Close the door behind you:** in Supabase go to **Authentication → Sign In / Providers** and turn off
   **Allow new users to sign up**. Your two accounts keep working, and nobody new can sign up.

## 6. Install it on your phones

- **iPhone (Safari):** open the address → Share button → **Add to Home Screen**.
- **Android (Chrome):** open the address → ⋮ menu → **Install app** (or **Add to Home screen**).

It then opens full screen from its own icon, with no browser bars.

---

## Good to know

- **Sync:** changes appear on the other phone within a second or two while both are online.
  When the app comes back from the background it refreshes itself.
- **Offline:** the app opens without signal, but saving needs an internet connection.
- **Backups:** use **Care plan → Backup → Download backup** about once a month, and keep the file in email or Google Drive.
  **Restore from backup** puts everything back.
- **Free Supabase projects are paused after about a week with no activity.** Daily use keeps it awake.
  If it ever pauses, log in to supabase.com and click **Restore project**. Your data is kept.
- **Updates:** push changes to GitHub and Vercel redeploys. The phones pick up the new version the next time the app opens.
- **Who can see the data:** only signed-in members of your family (enforced by the database rules), plus you as the Supabase account owner.
