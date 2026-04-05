# FRAME Medicine — Deployment Guide

## Overview
- **Backend:** Google Sheets + Apps Script (deployed as web app)
- **Patient App:** Single HTML file hosted on WordPress or GitHub Pages
- **Admin App:** Single HTML file hosted on WordPress or GitHub Pages
- **SMS:** Twilio (inbound webhook points to Apps Script deploy URL)
- **Auth:** Twilio Verify (patient OTP), email+password (admin)

---

## Step 1: Google Sheets Setup

### 1.1 Create the Spreadsheet
1. Open Google Sheets and create a new spreadsheet
2. Note the spreadsheet ID from the URL: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit`
3. Update `SHEET_ID` in Code.gs with your ID

### 1.2 Create All Tabs
Create these tabs (exact names, case-sensitive):
- `Patients`
- `Leads`
- `Billing`
- `Medications`
- `Messages`
- `Check-In Responses`
- `Check-Ins`
- `Weight Log`
- `Labs`
- `Sales`
- `Refill Log`
- `Dose History`
- `Finance`
- `Overhead Items`
- `Settings`
- `Catalog`

### 1.3 Add Headers
Add header rows to each tab matching the schema in SCHEMA.md.

### 1.4 Populate Settings Tab
Add these rows to the Settings tab (Column A = Key, Column B = Value):

| Row | Key | Value |
|-----|-----|-------|
| 1 | (header) Key / Value / Notes |
| 2 | CLINIC INFO (section) | |
| 3 | Clinic Name | FRAME Medicine |
| 4 | Tom Email | tom@framemedicine.com |
| 5 | Colin Email | drsheffield@framemedicine.com |
| 6 | Tom Phone | +19044948330 |
| 7 | Twilio Number | +12393726339 |
| 8 | Tom Password | (set your password) |
| 9 | Colin Password | (set your password) |
| 10 | | |
| 11 | FINANCIAL (section) | |
| 12 | Monthly Overhead | 623 |
| 13 | Tom Split | 0.25 |
| 14 | Colin Split | 0.75 |
| 15 | | |
| 16 | ALERTS (section) | |
| 17 | Refill Alert Days Before Due | 14 |
| 18 | No-Response Escalation (hrs) | 48 |
| 19 | | |
| 20 | MEMBERSHIP RATES (section) | |
| 21 | Monthly Testosterone | 150 |
| 22 | Sponsored Testosterone | 140 |
| 23 | Monthly Tirzepatide | 300 |
| 24 | Tirz + Test | 400 |
| 25 | Monthly Semaglutide | 200 |
| 26 | Tadalafil Standalone | 59 |
| 27 | Tadalafil Add-On | 45 |
| 28 | Hair Plan Foam | 99 |
| 29 | Hair Plan Pills | 159 |

---

## Step 2: Apps Script Setup

### 2.1 Open Script Editor
1. In Google Sheets: Extensions > Apps Script
2. Delete the default code
3. Paste the entire contents of `Code.gs`
4. Save (Ctrl+S)

### 2.2 Deploy as Web App
1. Click "Deploy" > "New deployment"
2. Type: "Web app"
3. Execute as: "Me"
4. Who has access: "Anyone"
5. Click "Deploy"
6. Copy the deployment URL
7. Update `DEPLOY_URL` in Code.gs with the new URL
8. Update `API_URL` in both `patient-app/index.html` and `admin-app/index.html`

### 2.3 Set Up Triggers
1. In Apps Script, run the `setupTriggers()` function once
2. This creates all automated daily/hourly triggers
3. Authorize all permissions when prompted

### 2.4 Re-deploy After Changes
After any Code.gs changes:
1. Click "Deploy" > "Manage deployments"
2. Click the pencil icon on your deployment
3. Change "Version" to "New version"
4. Click "Deploy"
5. The URL stays the same

---

## Step 3: Twilio Setup

### 3.1 Configure SMS Webhook
1. Go to Twilio Console > Phone Numbers > Your Number (+12393726339)
2. Under "Messaging" > "A MESSAGE COMES IN"
3. Set webhook URL to your Apps Script deployment URL
4. Method: HTTP POST
5. Save

### 3.2 Verify Twilio Credentials
Make sure these match in Code.gs:
- Account SID
- Auth Token
- Phone Number
- Verify Service SID

---

## Step 4: Host the Apps

### Option A: WordPress (EasyWP)

#### Patient App
1. In WordPress, create a page at `/app` (page ID 489)
2. Add a Custom HTML block
3. Paste the contents of `patient-app/index.html`
4. Upload `patient-app/manifest.json` to Media Library
5. Update the manifest URL in the HTML to match the Media Library URL
6. Upload `patient-app/sw.js` via File Manager to the site root

#### Admin App
1. Create a page at `/admin` (page ID 493)
2. Add a Custom HTML block
3. Paste the contents of `admin-app/index.html`
4. Upload `admin-app/manifest.json` to Media Library
5. Upload `admin-app/sw.js` via File Manager to the site root

#### WordPress Tips
- Paste HTML through Notepad first to strip any encoding
- Never use `&&` in JavaScript — WordPress converts it to `&#038;&#038;`
- All visibility changes use `element.style.display`, not CSS class toggles
- Test thoroughly after pasting — WordPress can mangle special characters

### Option B: GitHub Pages (Recommended)

#### Setup
1. Create a GitHub repo (e.g., `frame-medicine-apps`)
2. Push the `patient-app/` and `admin-app/` folders
3. Go to Settings > Pages > Source: main branch
4. Custom domain: Add `app.framemedicine.com` or use path-based routing

#### DNS Configuration (Namecheap)
For subdomain approach:
- Add CNAME record: `app` -> `yourusername.github.io`
- In GitHub repo settings, set custom domain to `app.framemedicine.com`

For path-based (using WordPress):
- Create a reverse proxy or redirect in WordPress for `/app` and `/admin`
- Or use Namecheap DNS to point specific paths

#### Benefits
- Version control (see every change, roll back easily)
- No WordPress encoding issues
- Free SSL
- Auto-deploys on git push
- Better for PWA service workers

### Option C: Netlify
1. Connect your GitHub repo to Netlify
2. Set build command: (none, static files)
3. Set publish directory: `/`
4. Add custom domain in Netlify settings
5. Auto-deploys from GitHub

---

## Step 5: Verify Everything

### Backend Verification
1. Open your deploy URL in a browser with `?action=getSettings`
2. You should see JSON with your settings data
3. Test admin login: `?action=adminLogin&email=tom@framemedicine.com&password=yourpassword`

### Patient App Verification
1. Open the patient app URL
2. Enter a test phone number
3. Verify OTP flow works
4. Check all screens render correctly

### Admin App Verification
1. Open the admin app URL
2. Log in with Tom's credentials
3. Verify dashboard loads with patient data
4. Test adding a patient
5. Test messaging
6. Verify Colin cannot see P&L

### Twilio Verification
1. Send a text to your Twilio number
2. Check the Messages tab in Google Sheets
3. Verify the inbound message was logged
4. Check that notification email was sent

---

## Step 6: Go Live Checklist

- [ ] All spreadsheet tabs created with headers
- [ ] Settings tab populated with correct values
- [ ] Apps Script deployed and URL updated in both apps
- [ ] Twilio webhook configured
- [ ] Patient app accessible at framemedicine.com/app
- [ ] Admin app accessible at framemedicine.com/admin
- [ ] manifest.json and sw.js uploaded and accessible
- [ ] OTP login tested with real phone number
- [ ] Admin login tested for both Tom and Colin
- [ ] Two-way messaging tested (send and receive)
- [ ] Triggers set up (run setupTriggers())
- [ ] Daily digest email verified
- [ ] PWA install works on mobile
- [ ] Colin confirmed: P&L section hidden

---

## Ongoing Maintenance

### Monthly Tasks
- Import Sales CSV (reminder email sent on last day of month)
- Import Patients CSV (reminder email sent on last day of month)
- Lock previous month P&L after reconciliation
- Review and update overhead if needed

### If You Need to Update Code.gs
1. Edit in Apps Script editor
2. Re-deploy (Manage deployments > New version)
3. URL stays the same — no app changes needed

### If You Need to Update the Apps
**WordPress:** Re-paste the HTML into the Custom HTML block
**GitHub Pages:** Push changes, auto-deploys in ~60 seconds
