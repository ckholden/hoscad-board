# HOSCAD Board — Claude Instructions

## Project
HOSCAD frontend — Dispatcher board, field MDT, admin portal, RMS, viewer.
Production: scmc.hoscad.net (Cloudflare Pages, auto-deploy on push to master)
Backend repo: `hoscad` (Supabase edge functions)

## Permissions
- Commit and push without asking

## Stack
- Vanilla HTML/CSS/JS — no framework, no build step
- PWA with service workers (`sw.js`, `sw-field.js`)
- Cloudflare Pages hosting + Cloudflare Access (zero-trust identity gate)

## Key Files
| File | Size | Purpose |
|------|------|---------|
| `app.js` | ~563 KB | Board logic, command parser, state management |
| `api.js` | ~26 KB | API wrapper (all calls go through `API.call()`) |
| `styles.css` | ~75 KB | Global styles — shared by board + viewer (dark theme) |
| `index.html` | ~50 KB | Board shell (dispatcher console) |
| `field/index.html` | Inline | Field MDT (mobile PWA, 7000+ lines) |
| `viewer/index.html` | Inline | CADview — read-only wall display, has own inline `<style>` |
| `popout-inc/index.html` | Inline | Popout incident queue, has own inline `<style>` |
| `admin/index.html` | Inline | Admin portal |
| `rms/index.html` | Inline | Crew reporting |

## Architecture Notes

### Service Workers
- `sw.js` — Board service worker
- `sw-field.js` — Field MDT service worker (currently cache `v91`)
- Bump cache version string manually when deploying field MDT changes

### Multi-Tenant Branding
- `getTenantSettings` endpoint returns BOTH the settings key-value map AND tenant core info (`displayName`, `accentColor`, `timezone`, `incidentPrefix`, etc.)
- Board (`app.js`): applies `displayName` to `#boardTitle`, `accentColor` as `--blue` CSS override
- Field MDT: applies tenant config on login, caches to localStorage for offline boot
- Viewer/popout-inc: inherit from board via shared token/state

### Field MDT (field/index.html)
- Has JS Table of Contents comment block at top of `<script>` (~40 section references)
- Uses z-index CSS custom properties (`:root` scale from `--z-bottom-nav: 20` to `--z-session-expired: 9999`)
- Full ARIA accessibility: `role="tabpanel"`, `aria-selected` on nav, `role="dialog"` + `aria-modal` on all overlays, `aria-live` regions
- `switchTab()` toggles `aria-selected` on nav buttons
- Offline-first: cached tenant config, offline watermark, service worker

### Viewer & Popout Styles
- `viewer/index.html` has its own inline `<style>` — isolated from main board
- `popout-inc/index.html` has its own inline `<style>` — isolated from main board
- Both inherit CSS variables from `styles.css` `:root` but override specific values for their context
- Viewer is optimized for wall/TV display at 10ft viewing distance (14px base, 32px rows)

### UI Polish (March 2026)
- WCAG AA contrast on all muted text (~5.2:1 ratio)
- Button click targets sized for stressed dispatchers (min 5px padding)
- Visible active states on toolbar toggles (blue underline)
- Focus-visible: 2px outlines on all interactive elements
- Board table: 13px base font, 30px rows (compact density is escape hatch)
- Big-screen media queries at 1600px and 2400px maintain +1px font progression
- Topbar responsive: sync badges hide <1200px, scope indicator hides <1000px
- PRI-1 incident rows have red left border (priority visible without color alone)
- Stale row indicators: 4px border with high-opacity pulses

## SaaS Transition — Frontend Work Queue

### Phase 0 — Quick Wins
(No frontend changes needed — Phase 0 is backend-only)

### Phase 2c — Frontend Tenant Isolation
- [x] **Per-tenant branding** — `getTenantSettings` returns tenant core info. Board + field MDT apply display name and accent color dynamically.
- [ ] **Tenant detection** — In `api.js`, extract tenant slug from `window.location.hostname.split('.')[0]`. Pass to login. Store in sessionStorage.
- [ ] **Per-tenant position dropdown** — Already dynamic from `init()` endpoint. Will auto-scope once backend is tenant-aware.

### Phase 3 — Operational Readiness
- [ ] **Service worker cache versioning** — Automate version bumps in CI (currently manual in sw.js/sw-field.js)
- [ ] **WebSocket migration** — Subscribe to Supabase Realtime for units, incidents, messages tables. Fall back to polling if WS disconnects.

## Recent Features (March 2026)
- **Incident mask auto-refresh** — Open incident modal refreshes on poll cycle (~10s) when STATE changes
- **Topbar universal search dropdown** — Search box searches units, active incidents, addresses, historical incidents, and people inline with keyboard navigation
- **Topbar search** — Search box at top of board is a comprehensive instant search with debounced server queries

## Known Issues
- **Central Oregon address DB** — Needs verification for completeness
- **Non-address locations** — Intersections, mile posts, landmarks not yet handled
- **Branding refresh** — Admin branding changes require board page refresh to take effect (no live push yet)
- **SCMC logo** — Needs `logo_url` configured in Admin → AGENCY tab for logo to display

## Audit Findings (2026-03-03)

### Fixed
- **CB/COPY incident ID normalization** — Was padding to 4 digits instead of 5 (mismatched `resolveIncidentId`). Fixed to pad to 5.
- **inc.units type mismatch** — `_refreshIncidentModal` treated `inc.units` as array (`.join()`), but it's a comma-separated string from API. Fixed.
- **Missing CSS variables** — `--fg`, `--surface`, `--accent` used in admin/inline styles but not defined in `:root`. Added to both dark and light themes.
- **Field MDT audit trail XSS** — `e.message` and `e.actor` not HTML-escaped in audit trail render. Fixed with `esc()`.

### Deferred (tracked for future work)
- **Board `start()` event listener leak** — Re-login accumulates event listeners. Consider cleanup function.
- **`renderBoardDiff` clears innerHTML** despite row-cache hash system — Can optimize to skip clear when hash matches.
- **Popout-inc color variables** differ from main board (`--green: #4dff91` vs `#20b060`) — cosmetic, low priority.
- **Field MDT full DOM rebuild on poll** for roster/calls tabs — Consider diff-based update for performance.
- **Field MDT wake lock listener leak** — Reacquire listener can accumulate on toggle. Low impact.
- **Field MDT offline queue** doesn't include transport destination — Missing field in queued operations.

## Documentation
Full technical docs in the `hoscad` repo at `/docs/`. Key references:
- `docs/API_SURFACE.md` — All 154 endpoints (api.js calls these)
- `docs/ARCHITECTURE.md` — System design and data flow
- `docs/AUTHENTICATION.md` — Auth flows (relevant to login pages)

## Conventions
- `_headers` file controls Cloudflare security headers (CSP, HSTS, X-Frame-Options)
- `manifest.json` / `manifest-field.json` — PWA manifests
- All apps share `styles.css` at root (board + viewer inherit from it)
- Inline JS in field/admin/rms HTML files (not separate .js files)
- `--font-bump` CSS variable with `calc()` for tenant-configurable font scaling
- Density modes (compact/normal/expanded) via `cycleDensity()` — escape hatch for row height changes
- Zero border-radius everywhere — intentional retro dispatch aesthetic
- CRT scanline texture on board background — do not remove
