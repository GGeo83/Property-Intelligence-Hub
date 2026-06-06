# Property Intelligence Hub

A private real estate monitoring dashboard for tracking permits, planning applications, landmarks activity, for-sale listings, and local news across a portfolio of properties in Massachusetts, Florida, Vermont, and New York.

## Overview

The PIH is a single-page HTML dashboard that aggregates property intelligence from multiple sources into a unified view organized by principal. It connects to a Google Sheets backend for live data and historical activity logging.

## Properties Monitored

### APJCJM
- 56 Beacon St — Boston, MA (Beacon Hill)
- Mandarin Oriental — 776 Boylston St W12A · W12B · 778 Boylston St APT 7G, Boston, MA
- Nantucket Estate — 1 Sandy Dr · 32B & 29 Hulbert Ave, Nantucket, MA
- Milton Estate — 1134, 1150, 1196 Canton Ave, Milton, MA
- 482 Island Drive — Palm Beach, FL
- 35 Hayride Drive — Stowe, VT
- 92 Blodgett Way — Lake Placid, NY

### ELJ
- 18 Louisburg Square — Boston, MA (Beacon Hill)
- Dover Estate — Farm St & Pegan Ln, Dover, MA
- 2929 Winding Oak Lane — Wellington, FL
- Louisburg Farm FL — 3261 & 3315 Old Hampton Dr, Wellington, FL

### EBJ
- 1 Charles River Square — Boston, MA
- 3 Charles River Square — Boston, MA

### ECJIV
- 131 Commonwealth Ave — Boston, MA (Back Bay)
- 16 Union Wharf — Boston, MA (North End / Waterfront)

## Architecture

- **Frontend:** Single-page HTML/CSS/JS (`index.html`) — no build step required
- **Backend:** Google Sheets (properties, activities, sources, scan_log, events tabs)
- **Write API:** Google Apps Script webhook (`setupSheet.js`) — handles `addProperty`, `updateProperty`, `deleteProperty`, `logScan`, and `logActivity` actions
- **Read API:** Google Sheets API (public read via API key)
- **Hosting:** Netlify (auto-deploy on push to `main`)
- **Scans:** Claude scheduled agents running every Friday at 5pm — Massachusetts, VT/NY, and Florida

## Weekly Scans

Three automated scans run every Friday at 5pm and feed findings directly into the dashboard:

| Scan | Coverage | Email |
|------|----------|-------|
| Massachusetts | Boston · Nantucket · Milton · Dover | 🏛 Massachusetts Scan — Friday [Date] |
| VT / NY | Stowe VT · Lake Placid NY · Maine (future) | 🌲 VT / NY Scan — Friday [Date] |
| Florida | Palm Beach · Wellington | 🌴 Florida Scan — Friday [Date] |

Each scan checks permits, planning applications, landmarks/commission filings, for-sale listings, local news, and deed activity — then posts findings to Google Sheets and emails a digest to the principal contact.

## Deployment

This is a static site — no build command or publish directory needed.

1. Fork or clone this repo
2. Connect to Netlify via **Import from Git**
3. Leave build settings blank → Deploy
