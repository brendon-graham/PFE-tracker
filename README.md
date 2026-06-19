# PFE Farm Tracker — v7.5

Mobile-first farm management app for Peel Forest Estate, South Canterbury, NZ.

## Live app

`https://brendon-graham.github.io/PFE-tracker/`

## What it does

- 9 tabs: Overview, Week, Feed Out, Breaks, Paddocks, Barn, Jobs, Toolbox, History
- 16 crop paddocks (Fodder Beet, Swedes, Kale, Rape) with break frequency logic
- Supplement inventory with frequency-aware feed-out (daily / MWF / Mon–Fri / every 2nd day)
- 4 pasture mobs with supplement tracking
- Barn TMR calculator — 4 feeds, 3 additives, urea step-up ceiling
- Feed-out checklist with route optimisation (Leaflet.js + farm GeoJSON)
- Weekly Jobs — day-by-day template with ticks and archive
- Backlog Jobs — 6 categories
- Toolbox Minutes builder with history
- GPS area measurement tool
- Google Sheets sync — debounced push (2.5s), 30s live poll
- PWA — add to home screen, auto-update banner

## Stack

Single HTML file (`index.html`) — vanilla JS, Leaflet 1.9.4, localStorage + Google Sheets. No build step.

## How updates work

Claude writes fixes and pushes directly to this repo. Staff just **refresh their browser** to get the latest version.

## Data storage

All farm data (jobs, mobs, feed outs, paddock settings) is stored in browser `localStorage` per device and synced to Google Sheets. Code updates via Git do **not** affect farm data.
