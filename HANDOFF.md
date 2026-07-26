# FRAME Medicine — Handoff / Where I Left Off

_Last updated: 2026-07-26 · Branch: `claude/organize-coding-events-niw91q` (identical to `master`)_

## What this project is

Men's-health telemedicine practice-management system.

- **Database:** Google Sheets
- **Backend:** Google Apps Script — `google-sheets/Code.gs`
- **Patient app:** single-file PWA — `patient-app/index.html`
- **Admin CRM:** single-file PWA — `admin-app/index.html` (a duplicate `admin/` copy also exists — see below)
- **SMS/OTP:** Twilio
- **Hosting:** WordPress (EasyWP) or GitHub Pages (custom domain via `CNAME`)
- **Native:** Capacitor wrapper under `native/` for the app stores

## Where it is right now

Core flows work: OTP login, patient dashboard, refill, check-in, weight logging, two-way SMS, admin CRM. Twilio is live, the PWA deploys cleanly, and the custom domain is wired up. Last code change (2026-06-10) restored the custom-domain `CNAME`. Recent work merged in patient-auth foundation, service worker, admin quick-fixes, native build hardening, and the "roadmap to 95%" audit; plus a mobile-overflow fix on the reset controls.

## What's next (from `docs/ROADMAP-TO-95.md`)

**P0 — do before promoting the app:**
1. **No auth on patient data endpoints** — `getPatientDashboard/getMessages/…` return full PHI for any phone number. OTP only gates the login screen, not the data. HIPAA-grade exposure. (Claude writes it + Tom coordinates the `Code.gs` redeploy.)
2. **Role check on financials** — Colin's admin token can POST `getPnl/updateOverhead/lockMonth` directly even though P&L is hidden in his UI.
3. **Refill confirm/decline shows "Confirmed!" even when the server fails** — frontend ignores `data.success`. Patient-safety.
4. **Editing a patient silently wipes Outstanding balance & Follow-Up flag.**
5. **Plaintext admin passwords & biometric tokens** in the sheets; hash + redact.
6. **Web push is dead** — empty `applicationServerKey`; decide real VAPID/native push or remove the dead path.

**P1 highlights:** med cost/supply math vs. SCHEMA, Catalog-sourced vial cost, `LockService` around read-modify-writes, stable patient ID (currently name-based key), lab-schedule drift, SMS consent/opt-out + "STOP" handling, Twilio inbound signature validation, email/HTML injection, iOS input zoom/autofill, service-worker drift, reconciling the diverged `index.html` copies, CSV import error reporting.

**Release / store:** Android signing PR pending a `workflow`-scope re-auth; app icons need a 1024×1024 `native/resources/icon.png`; iOS build blocked on Apple Developer enrollment.

⚠️ **Known drift to resolve:** `admin/` vs `admin-app/` are two ~2,300-line copies that have diverged; root `index.html` vs `patient-app/index.html` have also diverged. Pick a canonical source for each and delete/sync the other.

## Set up on a new computer

No build step for the web apps — they're standalone single-file HTML PWAs.

```bash
git clone https://github.com/TomATFRAME/Frame-Medicine.git
cd Frame-Medicine
# Open patient-app/index.html or admin-app/index.html directly in a browser to preview,
# or serve the folder:  python3 -m http.server 8000
```

- **Backend:** paste `google-sheets/Code.gs` into the Apps Script editor bound to the Google Sheet; redeploy after edits. Column definitions in `google-sheets/SCHEMA.md`.
- **Native build:** `cd native && npm install` (see `native/SIGNING.md`, `native/STORE_LISTING.md`).
- **Deploy:** follow `wordpress/DEPLOY.md`.

**Coding constraints (important):** in `Code.gs` use `var` only — no `const`/`let`, arrow functions, template literals, or ES6+. In the HTML apps avoid `&&` in JS (WordPress encoding), keep all CSS/JS inline, and toggle visibility via `element.style.display` only.
