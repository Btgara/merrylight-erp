BUILD KIT - the fast, small version of the Merrylight ERP
============================================================
Why: the standard version downloads ~4.5 MB and compiles itself in the
browser on first open - painful on slow connections. This kit produces a
precompiled version around 85% smaller that starts in a second or two.

NEEDS: any computer with Node.js 18+ (nodejs.org), internet, 10 minutes.
This is a ONE-TIME step per app update.

STEPS
  1. Open a terminal in this build-kit folder
  2. npm install          (first time only, ~1-2 min)
  3. public/backend-config.js is already pre-filled - no editing needed
  4. npm run build
  5. A "dist" folder appears - drag THAT folder onto Netlify Drop
That's it. Phones update automatically on next open (reopen twice if stale).

For each future app update from Claude: replace src/main.jsx with the new
one Claude provides, then repeat steps 4-5.
