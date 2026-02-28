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

## Major Feature Backlog

See plan file for architecture details on:
- CLI Context Mode (bind incident to CLI session)
- Location History (LH command)
- Danger/Safety Flag System (per address)
- Universal Query (! command)
- Related Incident Linking
- Hot Call Broadcast (HT command)
- Real-Time Incident Update Awareness (Supabase Realtime)
- Multi-Dispatcher Concurrency handling
- Incident ID Auto-Linking in Messages
- Soft Presence Indicators (who is viewing an incident)
