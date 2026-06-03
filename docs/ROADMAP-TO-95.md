# FRAME Medicine — Path to 95%

Comprehensive audit + prioritized roadmap to take the app to 95% functionality, cleanliness, and operability. Compiled 2026-06-03 from deep audits of the backend (`Code.gs`), the patient app, and the admin CRM.

**How to read this:** items are tiered P0 (do first) -> P3 (polish). "Owner" = who has to act: **Claude** (I write the code/PR), **Tom** (deploy, accounts, or a decision). Almost every backend fix also needs **Tom to redeploy `Code.gs`** in the Apps Script editor to take effect.

---

## Snapshot

What's solid: the core flows work (OTP login, dashboard, refill, check-in, weight, two-way SMS), Twilio is live, the PWA deploys cleanly, the privacy policy is HIPAA-grade, and the brand/UX foundation is strong.

What's holding it back from "best-in-class healthcare app": **patient data is not actually access-controlled**, a few **patient-safety/data-integrity bugs**, **compliance gaps** (SMS consent, plaintext credentials), and **drift** between duplicated copies of the apps. None are huge individually; together they're the gap to 95%.

The single most important finding: **the patient data API has no authentication** — OTP only gates the login screen, not the data. Anyone who passes a valid phone number can read another patient's full record. That's a HIPAA-grade exposure and is P0 #1.

---

## P0 — Critical (security, patient safety, data integrity). Fix before promoting the app.

| # | Item | Area | Why it matters | Owner |
|---|------|------|----------------|-------|
| 1 | **No auth on patient data endpoints** — `getPatientDashboard/getMessages/getWeightLog/getCheckIns/getLabStatus/getPatient` accept any phone/name and return full PHI. OTP only gates the login UI. | Backend | Any person (or patient) can read anyone's PHI by changing a parameter. HIPAA breach risk. | Claude + Tom (coordinated deploy) |
| 2 | **Colin can see & change financials** — `handleGetPnl` + overhead/lock actions only check "is an admin," never role. P&L is hidden in Colin's UI but his token can POST `getPnl`/`updateOverhead`/`lockMonth` directly. | Backend/Admin | The 25/75 split, revenue, and med costs are exposed to the co-founder they're hidden from; he can also alter overhead. | Claude + Tom (redeploy) |
| 3 | **Refill confirm/decline show "Confirmed!" even when the server fails** — frontend ignores `data.success`. | Patient app | A failed medication refill tells the patient it succeeded — they may skip reordering. Patient-safety. | Claude |
| 4 | **Editing a patient silently wipes Outstanding balance & Follow-Up flag** — admin edit form omits those fields, backend overwrites them with 0/blank on every save. | Admin/Backend | Routine edits silently destroy financial balances and clinical follow-up flags. | Claude + Tom (redeploy) |
| 5 | **Plaintext admin passwords & biometric tokens** in the Settings/Patients sheets; `getSettings` returns passwords to any admin. | Backend | Credential theft from anyone with sheet or API access. | Claude (hash + redact) + Tom (reset pw) |
| 6 | **Web push is dead** — `applicationServerKey` is empty, so `pushManager.subscribe` always fails; no patient is ever subscribed. Native (Capacitor) push also has no APNs/FCM backend. | Patient app/Infra | "Notifications" is a headline feature that currently does nothing. Decide: real VAPID/native push, or remove the dead path for v1. | Claude + Tom (keys/decision) |

---

## P1 — High (correctness, compliance, core UX, release)

**Backend correctness**
- **Med cost/supply math violates the documented rules** — `handleNewOrder` hardcodes 200mg/vial and a `30.44` fixed-month constant; SCHEMA says Catalog-driven, real-calendar-date only. Skews P&L and supply dates. *(Claude + Tom redeploy)*
- **Vial cost never sourced from Catalog** — comes only from the client payload (defaults to 0), so med costs in P&L are whatever the client sends. *(Claude + Tom)*
- **No `LockService`** around payment/order/lab read-modify-write — concurrent requests can double-count or clobber rows. *(Claude + Tom)*
- **Name-based primary key** — joins on `name.toLowerCase()`; duplicate names merge/orphan billing/labs, and renames don't cascade. Add a stable patient ID. *(Claude + Tom; bigger effort)*
- **Lab schedule drift** — next-due dates recompute from "when marked done" instead of a fixed schedule; the **annual lab gets stuck permanently "complete"** and never re-arms. *(Claude + Tom)*

**Compliance**
- **No SMS consent/opt-out gating** — outbound texts (`sendMessage`/`sendLabReminder`/`sendLoginLink`) don't check a consent flag; there's no consent column. TCPA/A2P risk. Add consent column + block when not consented/opted-out. *(Claude + Tom)*
- **Inbound "STOP" treated as a refill decline** — collides with the carrier opt-out keyword; patient unsubscribes from all texts while the app logs a "decline." *(Claude + Tom)*
- **No Twilio inbound signature validation** — anyone hitting the webhook can spoof inbound texts and trigger refill actions/emails. *(Claude + Tom)*
- **HTML/email injection** — unsanitized patient text is embedded raw into provider emails. *(Claude + Tom)*

**Patient app core UX**
- **iOS input zoom + missing autofill** — all inputs are 14px (iOS auto-zooms), OTP field lacks `autocomplete="one-time-code"`, phone lacks `autocomplete="tel"`. Hurts the login flow everyone hits. *(Claude)*
- **Service worker drift / stale-app risk** — two SWs exist (cache-first v1 vs network-first v2). The cache-first copy can pin patients on a stale app forever. Standardize on network-first and reconcile. *(Claude)*
- **root `index.html` vs `patient-app/index.html` have diverged** (apiGet error handling, SW, manifest). Manual sync is already failing — establish one source + a copy step. *(Claude)*
- **Silent dashboard failures** — load/refresh errors drop the user with no message; no retry. *(Claude)*

**Admin operability**
- **CSV import: no per-row error reporting, no undo** — a failed row leaves partial writes silently. *(Claude + Tom)*
- **Duplicate `admin/` vs `admin-app/` codebases** (both ~2300-line files) have drifted; fixes to one miss the other. Pick canonical, delete the other. *(Claude + Tom decision on which deploys)*

**Release / store** (from the earlier evaluation)
- **Android signing** — persistent keystore fix is built (PR pending a `workflow`-scope re-auth). *(Tom re-auth, then Claude)*
- **App icons** — generation step built; needs a 1024x1024 `native/resources/icon.png`. *(Tom asset)*
- **iOS build** — blocked on Apple Developer enrollment (3 signing secrets). *(Tom)*
- **Store listing** — fix broken ToS URL, support email, Apple demo-login. *(Claude + Tom)*

---

## P2 — Medium (operability, features, resilience)

**Healthcare features the app is missing**
- Lab **results viewing** (not just due/overdue status) — expected for TRT.
- **Appointment/visit scheduling** — core telemedicine expectation.
- **Billing/payment visibility** — invoices, payment method, history; refill is a charge with no price shown.
- **Prescription detail** — dosing instructions, injection schedule, shipment tracking.
- **Profile/address editing** — patients can't update shipping address (medications are shipped).
- **In-app notification center** — no inbox of alerts/reminders.

**Operability**
- **Bulk actions** in admin (bulk lab reminders, bulk refill texts, bulk status).
- **Missed check-in detection** — nobody is flagged when a patient stops responding.
- **Overdue refill texts only fire every 3rd day** (`daysLeft % 3`) — gaps in follow-up; alert daily once overdue.
- **Refill detection depends on one logged order** — a missing order silently disables all alerts; surface "active patient, no current order."
- **Labs only tracked for plans containing the literal "test"** — GLP-1/other patients get no lab oversight.
- **"Labs Due" admin pill is a no-op** (`filter(()=>true)`).
- **Log Dose Change** has a backend but no admin button.
- **Mark-as-Paid** has no undo and no payment ledger view.

**Resilience**
- Missing tabs/columns fail silently (empty results, no error surfaced).
- `doPost` assumes a non-empty body; empty/odd POSTs throw generically.
- Import coercion (`safeNumber`/`parseDate`) turns bad data into 0/null silently.

---

## P3 — Polish / cleanliness
- Accessibility: add `aria-label` to icon-only buttons; raise muted-text contrast (some below WCAG AA) and base font to 16px; enlarge sub-44px tap targets.
- Keyboard overlaps the chat input on iOS (`visualViewport` handling).
- Consolidate the admin's two conflicting modal systems; add device back-button/history support.
- Make the "Patient Notifications ON/OFF" switch a setting in the app (today it's edited in the sheet).
- Search should cover email + preferred name; add patient-list sorting.
- Remove dead code (no-op `sendPushNotification` routing, duplicated leads loop, stale comments).
- Move the committed Apps Script deploy URL/IDs to Script Properties (or document as non-secret).
- P&L YTD recomputes every prior month live on each call — use locked Finance rows for closed months.

---

## Suggested phasing

**Phase 1 — Secure & stabilize (P0).** Patient-data auth, P&L role enforcement, refill success-handling, patient-edit preservation, credential hashing, resolve the push decision. *Outcome: safe to put in front of real patients and the co-founder.*

**Phase 2 — Correctness & compliance (P1).** Catalog-driven costs, LockService, lab-schedule fix, SMS consent + STOP handling, webhook signature, SW reconciliation + root/patient-app sync, iOS input/autofill, CSV import errors, de-duplicate admin copies, finish store/release.

**Phase 3 — UX & features (P2).** Lab results, billing visibility, profile/address, prescription detail, bulk admin actions, missed-check-in + refill-cadence fixes, lab tracking for all plans.

**Phase 4 — Launch polish (P3).** Accessibility, modal/back-button, settings toggles, dead-code cleanup, store submission.

---

## What I'm doing now (autonomously, low-risk, as separate PRs you merge from your phone)
1. **Backend hardening PR** (Code.gs): P&L/overhead role gate (P0 #2), preserve Outstanding/Follow-Up on patient edit (P0 #4), fix `editBilling` audit patient, redact passwords from `getSettings` (P0 #5 partial). *Needs your Code.gs redeploy to take effect.*
2. **Patient-app safety + mobile PR**: refill confirm/decline honor `data.success` (P0 #3), 16px inputs + OTP/tel autofill, aria-labels, top safe-area (P1/P3).
3. Holding **P0 #1 (patient-data auth)** for a coordinated change with you — it must deploy frontend + backend together or it can break login for everyone. It's the first thing to do together when you're back.
