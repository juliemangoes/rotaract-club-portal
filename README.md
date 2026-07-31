# Rotaract Club Portal — Production PWA

A full club-management portal for Rotaract clubs: meetings & event calendar, dues and
receipts, financial controls, project proposals & workspaces, membership management,
resource library, and Rotary-year rollover — installable on any phone as a PWA.

Each club's data is isolated: members sign in with email + password, join their club
with an invite code, and can only ever read or write their own club's data
(enforced by Postgres row-level security, not just the UI).

## Stack

- **Frontend:** React + Vite + Tailwind, `vite-plugin-pwa` (installable, offline shell, push)
- **Backend:** Supabase — Postgres (club data as a versioned JSONB document with
  optimistic concurrency + realtime sync), Auth, Storage (PDFs/photos/receipts),
  Edge Function for Web Push
- **Hosting:** any static host (Vercel / Netlify / Cloudflare Pages)

No Supabase configured? The app automatically runs in **local demo mode**
(seeded sample club, data stays in the browser) — handy for development.

## 1. Set up Supabase (~10 minutes)

1. Create a free project at [supabase.com](https://supabase.com).
2. Open **SQL Editor** → paste the entire contents of `supabase/schema.sql` → **Run**.
   This creates the tables, row-level security, RPCs, realtime sync, and the
   `club-files` storage bucket.
3. **Settings → API**: copy the *Project URL* and *anon public* key.
4. (Recommended) **Authentication → Providers → Email**: decide whether to require
   email confirmation. For a small club, turning confirmation off makes onboarding smoother.

## 2. Set up Web Push (optional but recommended)

1. Generate VAPID keys locally:
   ```bash
   npx web-push generate-vapid-keys
   ```
2. Install the [Supabase CLI](https://supabase.com/docs/guides/cli), then from this folder:
   ```bash
   supabase link --project-ref YOUR_PROJECT_REF
   supabase secrets set VAPID_PUBLIC_KEY=... VAPID_PRIVATE_KEY=... VAPID_SUBJECT=mailto:you@yourclub.org
   supabase functions deploy push
   ```
3. Put the **public** key in `.env` (next step). Skipping this section just means
   notifications stay in-app instead of arriving as device push.

## 3. Run locally

```bash
cp .env.example .env        # fill in VITE_SUPABASE_URL, VITE_SUPABASE_ANON_KEY, VITE_VAPID_PUBLIC_KEY
npm install
npm run dev
```

Leave `.env` empty to develop against the built-in demo mode.

## 4. Deploy

**Vercel:** import the repo → framework *Vite* → add the three `VITE_*` environment
variables → Deploy. **Netlify:** build command `npm run build`, publish directory `dist`,
same env vars. Push notifications and PWA install require HTTPS, which both provide
by default.

## 5. First run — how a club onboards

1. **President** creates an account → **Start a club** → becomes President of a fresh club.
2. In **Finance → Dues → Configure**, set monthly/district/RI dues, due day, grace period.
3. In **More → Club settings**, find the **invite code**; share it in the club group chat.
4. **Members** create accounts (with the email the club knows them by!) → **Join a club**
   → enter the code. They appear under **Members** as *Applied*.
5. Any EBOD member opens their profile → **Approve as active member**, assigns roles
   (Secretary, Treasurer, …) — the app's permissions follow those roles.
6. Members install the app: browser menu → **Add to Home Screen** (the PWA prompt).

If a member signs up with an email that already matches a member record in the club
document, their account links to that record automatically — so you can pre-create
members (via EBOD → Invite) with their email, and they land in the right profile.

## Architecture notes

- **Data model:** the entire club lives in one versioned JSONB document
  (`club_data.doc`). Saves go through the `save_club_doc` RPC which rejects stale
  writes; the app then reloads the latest and asks the user to redo the change.
  Realtime subscription keeps every open device in sync. For clubs of ~10–50 members
  this is simple, fast, and easy to back up (it's one JSON blob per club).
- **Files:** uploads (minutes PDFs, receipts, photos, logos — up to 15 MB) go to the
  public `club-files` bucket under `clubId/...`; only club members can upload into
  their club's prefix. URLs are stored in the document.
- **Push:** the client stores its push subscription per club; the `push` edge function
  verifies the caller is a club member, then fans out via VAPID Web Push and prunes
  dead subscriptions.
- **Audit trail:** the in-app audit log is append-only by construction (no delete UI,
  reversals instead of deletions). If you later need bank-grade immutability,
  the next step is mirroring financial entries into an append-only Postgres table.

## Costs

Supabase free tier (500 MB database, 1 GB storage, 500K edge invocations) and
Vercel/Netlify free tiers comfortably cover a single club. A custom domain is the
only likely expense.

## Roadmap ideas

- Email digests (Supabase cron + Resend) for members who miss push
- CSV import for migrating an existing dues spreadsheet
- Normalized financial tables + exports for district reporting
- Multi-club districts: the schema already supports one user in many clubs —
  add a club switcher in `AuthGate` when you need it.
