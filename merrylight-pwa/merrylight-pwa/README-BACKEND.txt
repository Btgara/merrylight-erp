ADDING THE SHARED TEAM BACKEND (Supabase, free tier)
=====================================================
Result: every phone and computer using the app reads and writes ONE shared
database - users, batches, orders, stock, everything stays in sync.

STEP 1 - CREATE THE DATABASE (~5 minutes, free)
  1. Go to https://supabase.com -> Start your project -> sign up (free)
  2. New project -> pick a name and password -> Create
  3. Left sidebar -> SQL Editor -> New query
  4. Paste the entire contents of setup-backend.sql -> Run
     (creates one small "kv" table the app stores everything in)

STEP 2 - GET YOUR TWO KEYS
  1. Left sidebar -> Project Settings -> API
  2. Copy "Project URL"            (looks like https://abcd1234.supabase.co)
  3. Copy the "anon public" key    (long string starting with eyJ...)

STEP 3 - PUT THEM IN THE APP
  1. Open backend-config.js in any text editor
  2. Paste the URL into url: "" and the key into anonKey: ""
  3. Save

STEP 4 - REDEPLOY
  Drag the folder onto Netlify Drop again (same as before).
  Phones that already installed the app pick up the update automatically -
  close and reopen the app twice if it seems stale.

HOW IT BEHAVES
  - The database is the source of truth; each device also keeps an offline
    cache, so the app still opens and works without signal. Changes made
    offline are saved locally and pushed when you next save while online.
  - Teammates' changes appear when you open a page / reopen the app
    (it is not live-updating in real time).
  - Conflicts resolve last-write-wins - fine for a small team; avoid two
    people editing the same order at the same moment.

SECURITY - READ THIS
  The anon key ships inside the app, so anyone who has your app URL can
  read and write the data. For an unlisted internal tool this is usually
  acceptable; do not post the URL publicly. When you want proper per-user
  authentication at the database level (so the login PINs are enforced by
  the server, not just the app), that is the next upgrade - ask Claude to
  wire Supabase Auth into the login screen.

PDF BATCH IMPORT (optional)
  The "Import from file" card on the New batch page can extract records from
  PDF reports using AI. In the Claude web-app version this works out of the
    3. Redeploy
  Spreadsheet import (.xlsx/.xlsm/.csv) works everywhere with no key.
