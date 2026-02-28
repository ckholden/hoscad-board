# HOSCAD Roadmap

Items below are planned future enhancements, not yet scheduled for implementation.
When ready to begin work on any item, use plan mode to design the approach.

---

## Agency Onboarding & Scaling

### GIS Address Point Import (per agency)
When onboarding a new agency, authoritative address data must be imported to power
the address typeahead in the New Call form. This replaces manual guessing with
county-E911-grade accuracy — the same data source real CAD systems use.

**Process:**
1. Identify the county GIS/E911 ArcGIS REST service URL (contact county GIS dept)
2. Add a `SOURCE_CONFIG` entry to `tools/import-address-points.js` with the new county
3. Set `SUPABASE_SERVICE_KEY` env var and run:
   ```
   SUPABASE_SERVICE_KEY=eyJ... node tools/import-address-points.js --county=<name>
   ```
4. Run `rebackfillCanonicalAddresses` from the admin panel (IT role) to normalize existing incidents
5. Refresh schedule: run monthly, or after major county address changes

**Current coverage:** Deschutes (E911 direct), Crook, Jefferson, Lake (Oregon Statewide)

**Template for new counties:** see `SOURCE_CONFIG` comment in `tools/import-address-points.js`

**ArcGIS sources:**
- County-specific E911 services (best — purpose-built for dispatch, has `cad_address` field)
- Oregon Statewide: `https://services8.arcgis.com/8PAo5HGmvRMlF2eU/arcgis/rest/services/Oregon_Address_Points/FeatureServer/0` (filter by county name)
- Each state will have its own statewide service; county E911 services are always preferred

---

### Multi-Agency Scaling (White-Label / SaaS)
HOSCAD is designed to scale to additional agencies — EMS services, hospitals,
animal control, security dispatch, etc. Key onboarding tasks per new agency:

**Infrastructure:**
- Supabase project per agency (isolation) OR shared project with row-level agency scoping
- Cloudflare Pages deployment with agency-specific subdomain (e.g., `agency.hoscad.net`)
- Cloudflare Zero Trust policy configured for agency staff

**Data setup:**
- Run all migrations on new Supabase project
- Seed positions table with agency-appropriate roles
- Import GIS address points for agency's service area (see above)
- Configure agency settings: scope, type codes, destinations, roster

**Customization checklist:**
- [ ] Positions/roles appropriate for agency type (EMS vs. hospital vs. security)
- [ ] Type code taxonomy relevant to agency (transport types, incident categories)
- [ ] Destination list (hospitals, staging areas, facilities)
- [ ] Address coverage area (county GIS import)
- [ ] Banner and diversion messages if applicable
- [ ] User accounts and initial passwords

**Agency types this system is well-suited for:**
- EMS / hospital transport (current use case — Deschutes County)
- Fire/EMS combined dispatch
- Hospital internal transport coordination
- Animal control / county code enforcement dispatch
- Private security dispatch
- Event medical / mass casualty coordination

---

## Field User Documentation Portal

A separate web portal for field personnel to document on incidents after the fact.
This is for **post-incident documentation only** — not live CAD operations.

**Scope:**
- Web browser based (mobile-friendly PWA)
- Accessible regardless of incident status (including CLOSED)
- Limited to incidents the user was assigned to or worked
- HIPAA-safe educational design — no PHI beyond what's in the incident record

**Features:**
- View incident details (read-only, or append notes if permitted)
- Add structured post-incident notes / narrative
- Review audit trail of their own actions
- Print/export incident summary
- Fully auditable — all entries logged with timestamp and actor

**Access control:**
- UNIT role users: see only incidents where they appear in `linked_units` or `unit_id`
- DP/SUPV/MGR: full access for supervisory review
- VIEWER: read-only

**Reporting:**
- Printable incident summary (PDF-friendly layout)
- Shift report filtered by user

**Design consideration:**
This portal is intentionally separate from the live CAD board to:
1. Not distract from active dispatch operations
2. Allow slower, more deliberate documentation after incident resolution
3. Be usable on personal devices without full board access

---

## Completed (as of 2026-02-28)

All items below have been implemented and deployed:

- **CLI Context Mode** — `R <INC>` binds context, `EXIT` releases, auto-clears on close, context banner
- **Location History (LH)** — LH panel with Safety Flags / Location Info / History tabs, flag create/deactivate
- **Danger/Safety Flag System** — per-address flags, FlagCache, danger-banner in modal, ⚠ queue badge, admin panel
- **Address Resolution Engine** — 5-stage E911 pipeline (EXACT→PREFIX→DIRECTION→FUZZY→USER_INPUT), PostGIS, pg_trgm, resolution audit log
- **Universal Query (!)** — active incidents, units, addresses, destinations, type codes, historical incidents, +CALL
- **F3 Search Enhancement** — location history, flags, historical incidents, +CALL button
- **Related Incident Linking (ILINK/IUNLINK)** — incident_relationships table, bidirectional
- **Hot Call Broadcast (HT)** — broadcasts incident summary to all dispatchers
- **Real-Time Incident Update Awareness** — Supabase Realtime WebSocket, per-incident channel, incUpdateNotice
- **Multi-Dispatcher Concurrency** — field-level last-write-wins + full audit trail + soft presence
- **Incident ID Auto-Linking in Messages** — clickable INC links in message panels
- **Soft Presence Indicators** — 30s heartbeat, who is viewing an incident displayed in modal

## Next Candidates — CAD Expert Audit (2026-02-28)

Prioritized by operational impact. Full report available in session history.

---

### Sprint 1 — Safety & Compliance (Do First)

**1. Mandatory disposition at incident close**
- Backend: reject `closeIncident` with null/OTHER disposition
- Field app: change `closeIncidentAction` to a `CLOSE_REQUEST` that dispatcher must confirm with disposition selection
- Files: `incidents.ts` → `closeIncident` action; `field/index.html` close handler
- Impact: Without this, every field-app close defaults to `OTHER` disposition, corrupting all disposition-based reports from day one

**2. Danger flag warning at dispatch execution time**
- Frontend: FlagCache check on `canonical_address` when ASSIGN or `D <UNIT>` dispatch executes
- Show inline confirmation: "⚠ SAFETY FLAGS ACTIVE — CONFIRM DISPATCH"
- Files: `app.js` ASSIGN handler and `D` status dispatch path
- Impact: A dispatcher acting through CLI without opening the modal currently receives no safety warning

**3. PRI-1 unassigned alert after 5 minutes**
- Frontend: scan `STATE.incidents` for PRI-1 + QUEUED + age > 5min in board render loop
- Trigger existing audio alert + on-screen indicator
- Files: `app.js` render/interval logic
- Impact: Clinical event if a CCT call sits unassigned — must be a system interrupt

---

### Sprint 2 — Response Time Reporting (Compliance Necessity)

**4. Response time analytics report**
- Backend: `getResponseTimeReport(token, startDate, endDate)` in `admin.ts`
- Compute: dispatch-to-enroute, enroute-to-arrival, on-scene time, transport time, total call time
- Aggregate by incident_type category, unit, shift, day-of-week
- Admin panel: new Reports tab with date range picker and results table
- Note: All data already exists in `incidents` columns — this is purely query + display

**5. Disposition breakdown report**
- Same endpoint as above — add disposition column aggregate
- TRANSPORTED / CANCELLED-PRIOR / PATIENT-REFUSED / OTHER breakdown week-over-week

---

### Sprint 3 — Workflow Efficiency

**6. Visual elapsed timer escalation by unit status**
- CSS class escalation: `elapsed-warn` / `elapsed-alert` thresholds per status
  - DE: warn 10min, alert 20min
  - OS: warn 30min, alert 45min
  - T: warn 30min, alert 60min
- Files: `app.js` unit board render; `styles.css` new classes
- No backend change required

**7. SCHED command — scheduled call view**
- Client-side: filter `STATE.incidents` for QUEUED with `[HOLD:]` tags, sort by hold time
- Render chronological list of upcoming calls for shift planning
- Files: `app.js` command handler only

**8. Realtime disconnected indicator**
- When WebSocket is disconnected, show visible "REALTIME OFFLINE — DATA MAY BE STALE" banner
- Protects dispatchers from acting on stale state during reconnect window
- Files: `app.js` `_rtReconTimer` / `_rtConnect` path; `index.html` + `styles.css`

---

### Sprint 4 — Structural Debt

**9. Cache dispatcher_agencies per session in getState**
- Add 60-second in-memory Map keyed by token in `state.ts`
- Eliminates redundant DB query on every state poll for every dispatcher
- File: `state.ts` lines ~163-202

**10. Migrate MA state from incident_note text tags to proper table**
- Create `incident_ma` table: (incident_id, agency_id, status, requested/acknowledged/released timestamps + actors)
- Migrate requestMA / acknowledgeMA / releaseMA / listMA to write to table
- Eliminates text-tag fragility and enables MA utilization reporting
- File: new migration + `incidents.ts` MA section

---

### Future / Lower Priority

- **Unit crew certification matching**: warn when BLS unit assigned to CCT incident (patient safety warning, not hard block)
- **Structured patient reference field**: HIPAA-safe (non-PHI) patient reference on incident row (not free text in note)
- **Session idle timeout**: automatic session expiration for abandoned workstations
- **Hospital diversion integration**: manual diversion entry at minimum; HAVBed ingest longer term
- **Recurring / scheduled call support**: dialysis runs 3x/week, recurring transports
- **Field documentation portal**: post-incident documentation for UNIT role (prerequisite: Sprint 2 reporting complete first)
- **Multi-agency scaling**: `agency_settings` table for per-agency type codes / positions / destinations
