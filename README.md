# Nova Ops Dashboard

A React + Vite project, ready to deploy to Vercel.

## Deploying to Vercel (no domain needed)

You'll get a free `your-project-name.vercel.app` address automatically — no domain purchase required.

### Option A — Vercel website (easiest, no command line)
1. Go to https://vercel.com and sign up / log in (GitHub, GitLab, or email all work).
2. Push this folder to a new GitHub repository (or use "Import" → drag-and-drop if Vercel offers it for your account type).
3. In Vercel, click **Add New → Project**, select the repo.
4. Vercel auto-detects Vite. Leave the defaults:
   - Build Command: `vite build` (auto-filled)
   - Output Directory: `dist` (auto-filled)
5. Click **Deploy**. In ~1 minute you'll get a live URL like `nova-ops-dashboard.vercel.app`.

### Option B — Vercel CLI (from this folder)
```bash
npm install -g vercel
cd nova-vercel
vercel
```
Follow the prompts (it'll ask to link/create a project) — no domain needed, it gives you a `.vercel.app` URL immediately. Run `vercel --prod` to push to the production URL after that.

## Local development
```bash
npm install
npm run dev
```

## Setting up Supabase (logins + shared data)

### 1. Create your Supabase project
1. Go to https://supabase.com → sign up (free tier is fine) → **New Project**.
2. Pick a name, a database password (save it somewhere), and a region.
3. Wait ~2 minutes for it to provision.

### 2. Run the database schema
1. In your project, go to **SQL Editor → New query**.
2. Open `supabase-schema.sql` (in this folder), copy the whole thing in, and
   **before running it**, edit the last section: replace
   `'YOUR-EMAIL@example.com'` with the email address you'll log in with.
3. Click **Run**.

### 3. Create your own login (the admin account)
1. Go to **Authentication → Users → Add User**.
2. Enter the *same email* you put in the SQL above, and a password.
3. Tick "Auto Confirm User" if offered, so you don't need to click an email link.

### 4. Get your API keys
1. Go to **Project Settings → API**.
2. Copy the **Project URL** and the **anon public** key.

### 5. Add the keys to the app
- **Local dev:** copy `.env.example` to `.env` and paste the two values in.
- **Vercel:** go to your project → **Settings → Environment Variables**, add
  `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` with those same values,
  then redeploy (Deployments tab → ⋯ → Redeploy) so the build picks them up.

### 6. Log in and add your employees
1. Open the deployed site, sign in with the admin email/password from step 3.
2. Go to the **Access Control** tab (only admins see it).
3. For each employee (e.g. Dharanivel, Anbararsan):
   - First create their login the same way as step 3 (Authentication → Users
     → Add User, in Supabase) — this app can't create Auth logins itself, only
     manage what they can see once they sign in.
   - Then in Access Control, click **Add user**, enter their email, display
     name, and tick which tabs they should see (leave "Admin" unchecked).
4. Send them their email + password — they can log in at the same URL and
   will only see the tabs you granted them.

### About data storage
Shared data (inventory, employees, payments, etc.) now lives in Supabase
Postgres (`app_storage` table) instead of browser localStorage — so every
logged-in user sees the same real data, synced across devices.

**Worth knowing:** tab visibility is enforced in the app's UI and is enough to
stop someone from *stumbling into* a tab they shouldn't see. It is **not**
airtight security — the `app_storage` table's database policy currently
allows any logged-in user to read/write all data, since the data isn't split
into separate tables per tab yet. For most small teams this is a reasonable
trade-off, but if you're storing something highly sensitive and need it
technically impossible (not just hidden) for a restricted user to reach via
direct API calls, that needs splitting `app_storage` into per-domain tables
with matching row-level security — a bigger follow-up job, not something to
assume is already covered here.
