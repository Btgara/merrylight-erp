MERRYLIGHT PACKHOUSE ERP - PHONE APP PACKAGE (PWA)
====================================================

WHAT THIS IS
An installable web app. Host these files once, then anyone can install it
on Android and iPhone from the browser - no app store needed.

STEP 1 - HOST IT (one time, ~2 minutes, free)
Easiest: Netlify Drop
  1. Go to https://app.netlify.com/drop
  2. Drag this whole folder (or the zip's extracted contents) onto the page
  3. It gives you a URL like https://something.netlify.app
Alternatives: GitHub Pages, Cloudflare Pages, Vercel - any static host works.
It must be served over https (all of the above are).

STEP 2 - INSTALL ON PHONES
Android (Chrome):
  1. Open your URL in Chrome
  2. Tap the three-dot menu -> "Add to Home screen" / "Install app"
  3. Confirm - the Merrylight icon appears like a normal app

iPhone / iPad (Safari - must be Safari):
  1. Open your URL in Safari
  2. Tap the Share button (square with arrow)
  3. Tap "Add to Home Screen" -> Add

The app opens full-screen with its own icon, and after the first load it
also works offline (charts, entry, everything).

IMPORTANT - HOW DATA WORKS IN THIS VERSION
All data (sign-ins, batches you add, orders, stock lines) is stored ON THE
DEVICE it was entered on. Two phones will each have their own copy - there
is no automatic sync between devices in this package. The 406 workbook
batch records and the week-29 packshed snapshot are built in on every device.

If you need true team sync (one shared order book across all phones), that
requires a small backend (e.g. Supabase or Firebase) - roughly a day of
setup. Ask Claude to wire it up when you're ready.

FILES
  index.html            the entire app
  manifest.webmanifest  install metadata (name, icon, colors)
  sw.js                 service worker (offline support)
  icon-*.png            app icons
