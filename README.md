# FRAME Medicine

Men's health telemedicine practice management system.

## Architecture

- **Database:** Google Sheets
- **Backend:** Google Apps Script (`google-sheets/Code.gs`)
- **Patient App:** Single-file PWA (`patient-app/index.html`)
- **Admin CRM:** Single-file PWA (`admin-app/index.html`)
- **SMS/OTP:** Twilio
- **Hosting:** WordPress (EasyWP) or GitHub Pages

## Project Structure

```
frame-medicine/
├── README.md
├── google-sheets/
│   ├── SCHEMA.md          <- Column definitions
│   ├── Code.gs            <- Complete Apps Script backend
│   └── data/              <- Sample CSVs (fake data only)
├── patient-app/
│   ├── index.html         <- Patient PWA (single file)
│   ├── manifest.json      <- PWA manifest
│   └── sw.js              <- Service worker
├── admin-app/
│   ├── index.html         <- Admin CRM (single file)
│   ├── manifest.json      <- PWA manifest
│   └── sw.js              <- Service worker
└── wordpress/
    └── DEPLOY.md          <- Deployment instructions
```

## Quick Start

See `wordpress/DEPLOY.md` for full deployment instructions.

## Constraints

### Apps Script (Code.gs)
- `var` only — no `const` or `let`
- No arrow functions
- No template literals
- No ES6+ features

### HTML Apps
- No `&&` in JavaScript (WordPress encoding issue)
- All CSS/JS inline in single HTML file
- Visibility via `element.style.display` only
