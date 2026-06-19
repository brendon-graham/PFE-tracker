# PFE Farm Tracker

Mobile-first farm management app for Peel Forest Estate, South Canterbury, NZ.

## Live app
Once deployed via GitHub Pages, the app is accessible at:
`https://<your-github-username>.github.io/pfe-tracker/`

## What it does
- Winter crop paddock tracking (Fodder Beet, Swedes, Kale, Rape)
- Supplement inventory with frequency-aware feed-out (daily / MWF / Mon–Fri / every 2nd day)
- Daily feed-out checklist with tick boxes
- Barn TMR calculator (MA / R2 / R3 stag mobs)
- GPS-based paddock area measurement
- Break size calculator for daily fence shifting
- Daily log history
- Optional Google Sheets cloud sync

## Stack
Single HTML file (`index.html`) — vanilla JS, Leaflet 1.9.4 maps, localStorage persistence. No build step, no dependencies to install.

## How to update the app

Claude writes code fixes and pushes directly to this repo. Staff just **refresh their browser** to get the latest version. No downloading or re-sharing files needed.

### To deploy changes:
1. Make changes to `index.html`
2. `git add index.html && git commit -m "fix: description of change"`
3. `git push origin main`
4. GitHub Pages rebuilds in ~30 seconds

## Data storage
All data stored in browser `localStorage` (per device). Optional Google Sheets sync available via the ⚙ Sync settings button.
