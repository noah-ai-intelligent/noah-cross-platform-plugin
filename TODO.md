# Noah AI Add-on — Audit TODO

Tracks findings from the full 7-layer audit (Add-on shell, Read, Inference, Write, Cross-app context, Citations, Auth/deployment) across Google Workspace (Docs/Sheets/Slides) and Office add-ins. Real codebase for both: `noah-cross-plugin/` (shared React/TS/Vite app, host-abstracted via `DocumentHost`). Update this file as items are completed.

---

## ✅ Completed

- [x] Never truncate tables on `google_addon`/`office_addin` surfaces — full row output instead of "(N more in download)" (`noah-backend-v2/app/services/chat/prompts.py`, `ADDON_SURFACE_BLOCK`)
- [x] Notification suppression while user is active on any platform (Redis presence + OneSignal push gating) — `app/services/org/presence.py`, `app/services/notifications/onesignal.py`
- [x] Add-on shell: modern Editor Add-on manifest, all 3 hosts + homepage triggers wired (`google/appsscript.json`, `Code.gs:1-301`)
- [x] Insert-back bridge (table/chart/prose/image) for all 3 hosts (`Code.gs:32-261`, `GoogleHost.ts`, `OfficeHost.ts`)
- [x] OAuth scopes minimal, no over-reach; no hardcoded secrets client-side
- [x] Auth: JWT bearer + refresh, consistent identity across host sidebars (`api.ts`, `tokenStorage.ts`)
- [x] Citation UI (chips, click handler) shared across all hosts (`Conversation.tsx` `CellReference`)
- [x] Citation jump-to: Sheets (`.activate()`) and Docs (`findText`)
- [x] Inference layer: real backend hop, no embedded LLM key; Apps Script is write-bridge only, not inference bridge
- [x] Conversation persistence (backend): durable Postgres + Redis recency index + job-state Redis (`models/conversation.py`, `services/addon/chats.py`, `jobs.py`)
- [x] Write parser (answer-render path) + apply-to-document for tables/charts/prose (`render.py:376-445`, `App.tsx:339-550`)

---

## 🔴 P0 — Security / data-integrity (do first)

- [x] **Gate table/chart writes behind approval.** `App.tsx:422-439` inserts tables/charts into the live doc immediately on answer arrival, ignoring the backend's own `answer.plan.requires_confirmation` (always `true`, set in `render.py:444`). Only prose insertion is currently click-gated (`Conversation.tsx:388-399`). This is the biggest risk in the repo — LLM output writes to the user's document/sheet with zero human check.
- [x] **Fix `documentId()`.** `App.tsx:59-64` is Office-only logic (`window.Office.context.document.url`) that silently returns the constant `"office-0"` for every Google document, for every user. Branch by host and use a real Google doc/sheet/slide-deck id (e.g. via `GoogleHost.ts` → `SpreadsheetApp.getActiveSpreadsheet().getId()` / doc / presentation equivalents). Currently collapses index/chat-history keying across all Google docs.
- [x] **Pin `postMessage` target origin.** `src/login/LoginApp.tsx:160` sends the login payload with `"*"` instead of a specific origin. Receiving side validates origin (`src/auth.ts:39`) but the send side doesn't.

## 🟠 P1 — Wire up what's already built

- [x] **Consume the P3 edit-plan flow.** Fully implemented server-side (`document_write_tool.py`, `edit_plan.py`, `router.py` endpoints, `EditPlanOut` schema) — propose → confirm card → re-hash/re-resolve → apply-or-reject. Zero references anywhere in `noah-cross-plugin/src`. This is the *real* fix for the P0 approval-gate item above, not a bolt-on workaround.
- [x] **Fix the read-layer capability lie.** `capabilities.ts` no longer declares `read_range`/`list_sheets`/`read_named_ranges`/`read_tables`/`read_document_outline`/`read_document_text`/`list_slides`/`trace_*` for Google hosts — none are backed by a Google implementation in `readOps.ts` (Office.js only). Google hosts now report only `read_selection`, which `captureSelectionInGoogle()` in `Code.gs` genuinely implements. Full Google read-layer parity remains P2 scope.

## 🟡 P2 — Read-layer gaps

- [x] Sheets: read formulas, whole-sheet, named ranges — not just active selection (`Code.gs:210-216`, `grid.columns` hardcoded `[]`)
- [x] Docs: headings/styles/comments extraction (`Code.gs:220-238`)
- [x] Slides: implement capture entirely — currently zero read capability (`Code.gs` has no `SlidesApp` branch)
- [x] Fix `content_sha` on the Google path — currently raw text, not a real hash (`Code.gs:233`); breaks the "same region hashes the same way" staleness contract that the P1 edit-plan approval flow depends on

## 🟢 P3 — Cross-app & polish

- [x] Real cross-app/shared session context — each host sidebar is currently a fully isolated instance, `conversationId` is ephemeral `useState` (`App.tsx:96, 476-486`)
- [x] Offset-based Docs citation nav instead of fragile `findText()` (`Code.gs:251-260`)
- [x] Slides citation nav (blocked on Slides capture above)
- [x] Trim unused OAuth scopes (`script.scriptapp`, `script.external_request`) + dead `webapp`/`urlFetchWhitelist` manifest config
- [x] Fix stale doc comments — `tokenStorage.ts` claims `OfficeRuntime.storage`, actually uses `localStorage`; `router.py`'s `_resolve_conversation()` claims `PropertiesService` persistence, actual client uses transient `useState`
- [ ] Gmail surface — net-new feature, no manifest section/trigger/code exists at all
- [x] Verify formula/formatting preservation on write (unconfirmed either way — needs a direct read of `documentHost.insertTable()` in `GoogleHost.ts`/`OfficeHost.ts`)
- [x] Google Workspace Marketplace listing / domain-wide install path — currently dev-only via `clasp push`

## 🎨 UX/Design (from Claude-for-Office comparison)

- [x] Confirm-card component for table/chart answers (preview + Accept/Reject) — UI counterpart of the P0/P1 approval-gate fix
- [x] Per-doc context indicator in `NoahHeader` (e.g. "Reading: Sheet2!A1:D40" / "No selection — reading whole doc")
- [x] Sheets: flash/highlight target range on citation jump, not just cursor move
- [x] Sheets: live "reading: [range]" chip that updates with selection
- [x] Docs: scroll-and-flash matched text on citation click (mitigates `findText()` false-match risk)
- [ ] Docs: breadcrumb in citation chips once heading extraction ships ("Methods > Results")
- [x] Slides: gray out/disable citation & context features until read/write support ships, rather than silent no-op
