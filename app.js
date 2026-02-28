/**
 * HOSCAD/EMS Tracking System - Application Logic
 *
 * Main application module handling all UI interactions, state management,
 * and command processing. Uses the API module for backend communication.
 *
 * PERFORMANCE OPTIMIZATIONS (2026-01):
 * - Granular change detection: Lightweight hash per data section instead of JSON.stringify
 * - Selective rendering: Only re-render sections that actually changed
 * - DOM diffing: Board uses row-level caching, only updates changed rows
 * - Event delegation: Single click/dblclick handler on table body vs per-row
 * - Pre-computed sort keys: Timestamps computed once before sort, not in comparator
 * - Efficient selection: Uses data-unit-id attribute instead of text parsing
 */

// ============================================================
// Global State
// ============================================================
let TOKEN = localStorage.getItem('ems_token') || '';
let ACTOR = '';
let ROLE = '';
let STATE = null;
let ACTIVE_INCIDENT_FILTER = '';
let POLL = null;
let BASELINED = false;
let LAST_MAX_UPDATED_AT = '';
let LAST_NOTE_TS = '';
let LAST_ALERT_TS = '';
let LAST_INCIDENT_TOUCH = '';
let LAST_MSG_COUNT = 0;
let _holdAlertedIds = new Set(); // track HOLD calls already alerted this session
let _welfareAlertedKeys = new Set(); // track units already beeped for welfare check — key: unit_id:updated_at
let _lastAvCount = null; // track AV count transitions for coverage alerts
let _urgentIncAlertedIds = new Set(); // track PRI-1/urgent incidents already beeped on creation
let _unattendedAlertedIds = new Set(); // track QUEUED incidents already beeped for 30-min no-unit warning
let _prevActiveUnitSet = null; // previous poll's active unit IDs — null = not yet initialized
let CURRENT_INCIDENT_ID = '';
let CMD_HISTORY = [];
let CMD_INDEX = -1;
let SELECTED_UNIT_ID = null;
let UH_CURRENT_UNIT = '';
let UH_CURRENT_HOURS = 12;
let CONFIRM_CALLBACK = null;
let CONFIRM_CANCEL_CALLBACK = null;
let _newUnitResolve = null;
let _newUnitPendingNote = '';
let _MODAL_UNIT = null;
let _popoutBoardWindow = null;  // viewer popout (used by POPOUT + BOARDS commands)
let _popoutIncWindow   = null;  // incident queue popout
let _showAssisting = true; // Show assisting agency units (law/dot/support) by default
// LifeFlight ADS-B module state
let _lfnAircraft       = [];   // current fleet state array
let _lfnPrevAlt        = {};   // { tail: alt_baro } for descent detection
let _lfnPrevInbd       = {};   // { tail: true } 2-poll debounce for INBOUND
let _lfnPollTimer      = null; // setInterval handle
let _lfnInboundAlerted = {};   // { tail: true } suppress repeat INBOUND alerts
let _lfnLastSync       = null; // Date of last successful ADS-B fetch
let _lfnSyncing        = false;// prevents overlapping fetches
let _dc911LastSync     = null; // Date of last DC911 ingest (from STATE.dc911State.updatedAt)
let _lastPollAt    = 0;    // unix ms of last successful getState response — used for staleness indicator
const _expandedStacks = new Set(); // unit_ids with expanded stack rows (Phase 2D)
let _undoStack = []; // [{description, revertFn, ts}] — max 3 entries, 5-min expiry
let _RT = null;            // Supabase Realtime WebSocket
let _rtHbTimer = null;     // heartbeat interval (25s)
let _rtReconTimer = null;  // reconnect timeout (5s)
let _rtRef = 0;            // Phoenix message ref counter

// VIEW state for layout/display controls
let VIEW = {
  sidebar: false,
  incidents: true,
  messages: true,
  density: 'normal',
  sort: 'status',
  sortDir: 'asc',
  filterStatus: null,
  filterType: null,
  preset: 'dispatch',
  elapsedFormat: 'short',
  nightMode: false,
  theme: 'dark',    // 'dark' | 'night' | 'light'
  showActiveBar: false  // active calls bar below map — off by default
};

// Hardcoded fallback positions — used immediately on page load before API responds.
// Overwritten by live DB data once init() completes.
const POSITIONS_FALLBACK = [
  { position_id: 'DP1',   label: 'Dispatcher 1',    display_order: 10,  is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'DP2',   label: 'Dispatcher 2',    display_order: 20,  is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'DP3',   label: 'Dispatcher 3',    display_order: 30,  is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'DP4',   label: 'Dispatcher 4',    display_order: 40,  is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'DP5',   label: 'Dispatcher 5',    display_order: 50,  is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'DP6',   label: 'Dispatcher 6',    display_order: 60,  is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'SUPV1', label: 'Supervisor 1',    display_order: 70,  is_dispatcher: true,  is_admin: true,  can_reset: false },
  { position_id: 'SUPV2', label: 'Supervisor 2',    display_order: 80,  is_dispatcher: true,  is_admin: true,  can_reset: false },
  { position_id: 'MGR1',  label: 'Manager 1',       display_order: 90,  is_dispatcher: true,  is_admin: true,  can_reset: true  },
  { position_id: 'MGR2',  label: 'Manager 2',       display_order: 100, is_dispatcher: true,  is_admin: true,  can_reset: true  },
  { position_id: 'EMS',   label: 'EMS Coordinator', display_order: 110, is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'TCRN',  label: 'Transport RN',    display_order: 120, is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'PLRN',  label: 'Placement RN',    display_order: 130, is_dispatcher: true,  is_admin: false, can_reset: false },
  { position_id: 'IT',    label: 'IT Support',      display_order: 140, is_dispatcher: true,  is_admin: true,  can_reset: true  },
  { position_id: 'UNIT',  label: 'Field Unit',      display_order: 150, is_dispatcher: false, is_admin: false, can_reset: false },
  { position_id: 'VIEWER',label: 'Viewer',          display_order: 160, is_dispatcher: false, is_admin: false, can_reset: false },
];

// Positions metadata — populated from API on startup; drives dropdown + role checks
let POSITIONS_META = [...POSITIONS_FALLBACK];

// Admin role check — uses POSITIONS_META when loaded, falls back to hardcoded list
function isAdminRole() {
  if (POSITIONS_META.length > 0) {
    return POSITIONS_META.some(p => p.position_id === ROLE && p.is_admin);
  }
  return ['SUPV1','SUPV2','MGR1','MGR2','IT'].includes(ROLE);
}

// Returns the display label for a role/position ID
function getRoleLabel(roleId) {
  const pos = POSITIONS_META.find(p => p.position_id === roleId);
  return pos ? pos.label : roleId;
}

// Populate the login role dropdown from positions data
function _populateLoginRoleDropdown(positions) {
  const sel = document.getElementById('loginRole');
  if (!sel || !positions || !positions.length) return;
  const prev = sel.value;
  sel.innerHTML = '<option value="">SELECT POSITION...</option>';
  positions
    .slice()
    .sort((a, b) => (a.display_order || 0) - (b.display_order || 0))
    .forEach(p => {
      const opt = document.createElement('option');
      opt.value = p.position_id;
      opt.textContent = p.label || p.position_id;
      sel.appendChild(opt);
    });
  // Restore previously selected value if it still exists
  if (prev) sel.value = prev;
}

// Populate dropdown immediately with fallback so it's never empty on page load
document.addEventListener('DOMContentLoaded', function() {
  _populateLoginRoleDropdown(POSITIONS_FALLBACK.filter(p => p.is_dispatcher));
});

// Unit display name mappings
const UNIT_LABELS = {
  "JC": "JEFFERSON COUNTY FIRE/EMS",
  "CC": "CROOK COUNTY FIRE/EMS",
  "BND": "BEND FIRE/EMS",
  "BDN": "BEND FIRE/EMS",
  "RDM": "REDMOND FIRE/EMS",
  "CRR": "CROOKED RIVER RANCH FIRE/EMS",
  "LP": "LA PINE FIRE/EMS",
  "SIS": "SISTERS FIRE/EMS",
  "AL1": "AIRLINK 1 RW",
  "AL2": "AIRLINK 2 FW",
  "ALG": "AIRLINK GROUND",
  "AL": "AIR RESOURCE",
  "ADVMED": "ADVENTURE MEDICS",
  "ADVMED CC": "ADVENTURE MEDICS CRITICAL CARE"
};

const STATUS_RANK = { D: 1, DE: 2, OS: 3, T: 4, TH: 4, AV: 5, IQ: 6, OOS: 7 };
const VALID_STATUSES = new Set(['D', 'DE', 'OS', 'F', 'FD', 'T', 'TH', 'AV', 'UV', 'BRK', 'OOS', 'IQ']);
const KPI_TARGETS = { 'D→DE': 5, 'DE→OS': 10, 'OS→T': 30, 'T→AV': 20 };

// Incident type taxonomy for cascading selects (4A) — overridden by server if admin has customized it
// Transport-type focused for SCMC interfacility dispatch. Clinical body-system style.
// Priority levels (determinants):
//   PRI-1 = CCT/ALS — critical/unstable, life-threatening, requires CCT or advanced life support
//   PRI-2 = ALS     — time-sensitive, ALS monitoring required
//   PRI-3 = BLS     — stable, basic life support adequate
//   PRI-4 = BLS Routine — scheduled/non-urgent (discharge, dialysis)
//   NONE  = Administrative resolution type (cancellations, exceptions)
// Metadata fields: clinical_group, service_level, clinical_severity, legacy (bool), desc
let INC_TYPE_TAXONOMY = {
  CCT: {
    natures: {
      // ── NEW clinical types ──
      'CARDIAC-CRITICAL':    { dets: ['PRI-1'], clinical_group: 'CARDIOVASCULAR', service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Cardiac arrest, cardiogenic shock, hemodynamic failure requiring CCT-level care' },
      'STEMI':               { dets: ['PRI-1'], clinical_group: 'CARDIOVASCULAR', service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'STEMI with active intervention (IABP, pressors, CCT-level monitoring)' },
      'POST-ARREST':         { dets: ['PRI-1'], clinical_group: 'CARDIOVASCULAR', service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Post-cardiac arrest with ROSC — targeted temp management or hemodynamic instability' },
      'NEURO-CRITICAL':      { dets: ['PRI-1'], clinical_group: 'NEUROLOGICAL',   service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Hemorrhagic stroke, brain herniation, ICP crisis, post-craniotomy' },
      'STROKE-ALERT':        { dets: ['PRI-1'], clinical_group: 'NEUROLOGICAL',   service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Ischemic stroke with tPA on board or LVO requiring thrombectomy-capable center' },
      'RESPIRATORY-FAILURE': { dets: ['PRI-1'], clinical_group: 'RESPIRATORY',    service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Intubated/vented patient, high-risk airway, CPAP-dependent respiratory failure' },
      'SEPSIS-SHOCK':        { dets: ['PRI-1'], clinical_group: 'INFECTIOUS',     service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Septic shock with vasopressors — hemodynamically unstable, CCT-level monitoring' },
      'TRAUMA-CRITICAL':     { dets: ['PRI-1'], clinical_group: 'TRAUMA',         service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Unstable multi-system trauma, hemorrhagic shock, damage-control surgery post-op' },
      'BURN-CRITICAL':       { dets: ['PRI-1'], clinical_group: 'TRAUMA',         service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Severe burns requiring burn center CCT (>20% TBSA, inhalation, airway involvement)' },
      'OB-CRITICAL':         { dets: ['PRI-1'], clinical_group: 'OB',             service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'OB emergency: eclampsia, abruption, peripartum cardiomyopathy requiring CCT' },
      'PEDIATRIC-CRITICAL':  { dets: ['PRI-1'], clinical_group: 'GENERAL',        service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Neonatal/peds critical care transport (NICU/PICU level)' },
      'MULTI-SYSTEM-FAILURE':{ dets: ['PRI-1'], clinical_group: 'GENERAL',        service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'Multi-organ failure, complex CCT not fitting single system category' },
      // ── LEGACY — kept for existing incident display ──
      'VENT':             { dets: ['PRI-1'], clinical_group: 'RESPIRATORY',    service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: true, desc: '[LEGACY] Ventilator-dependent transport' },
      'MULTI-DRIP':       { dets: ['PRI-1'], clinical_group: 'GENERAL',        service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: true, desc: '[LEGACY] Multiple high-risk infusions' },
      'CRITICAL-TRAUMA':  { dets: ['PRI-1'], clinical_group: 'TRAUMA',         service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: true, desc: '[LEGACY] Critical trauma — use TRAUMA-CRITICAL' },
      'ECMO':             { dets: ['PRI-1'], clinical_group: 'CARDIOVASCULAR', service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: true, desc: '[LEGACY] ECMO transport — use CARDIAC-CRITICAL' },
      'NICU-PICU':        { dets: ['PRI-1'], clinical_group: 'GENERAL',        service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: true, desc: '[LEGACY] Neonatal/peds critical — use PEDIATRIC-CRITICAL' },
      'HIGH-RISK-AIRWAY': { dets: ['PRI-1'], clinical_group: 'RESPIRATORY',    service_level: 'CCT', clinical_severity: 'LIFE_THREATENING', legacy: true, desc: '[LEGACY] High-risk airway — use RESPIRATORY-FAILURE' },
    }
  },
  'IFT-ALS': {
    natures: {
      // ── NEW clinical types ──
      'CHEST-PAIN':           { dets: ['PRI-1','PRI-2'], clinical_group: 'CARDIOVASCULAR', service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Chest pain with ACS concern, troponin trending, ALS monitoring required' },
      'STEMI':                { dets: ['PRI-1','PRI-2'], clinical_group: 'CARDIOVASCULAR', service_level: 'ALS2', clinical_severity: 'LIFE_THREATENING', legacy: false, desc: 'STEMI transfer — not yet CCT-level but requires ALS2' },
      'STROKE':               { dets: ['PRI-1','PRI-2'], clinical_group: 'NEUROLOGICAL',   service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Stroke / neuro deficit, time-sensitive transfer' },
      'SEIZURE':              { dets: ['PRI-1','PRI-2'], clinical_group: 'NEUROLOGICAL',   service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Active or post-ictal seizure requiring ALS monitoring' },
      'AMS':                  { dets: ['PRI-1','PRI-2'], clinical_group: 'NEUROLOGICAL',   service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Altered mental status, unclear etiology, ALS monitoring' },
      'RESPIRATORY-DISTRESS': { dets: ['PRI-1','PRI-2'], clinical_group: 'RESPIRATORY',   service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Respiratory distress — not yet failure, O2 dependent, ALS airway monitoring' },
      'SEPSIS':               { dets: ['PRI-1','PRI-2'], clinical_group: 'INFECTIOUS',     service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Sepsis without shock — IV antibiotics running, close monitoring' },
      'GI-BLEED':             { dets: ['PRI-1','PRI-2'], clinical_group: 'GI',             service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Active or suspected GI bleed, hemodynamic monitoring required' },
      'OB-COMPLICATION':      { dets: ['PRI-1','PRI-2'], clinical_group: 'OB',             service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'OB complication — preterm, hypertension, bleeding, fetal concern' },
      'OVERDOSE':             { dets: ['PRI-1','PRI-2'], clinical_group: 'TOXICOLOGY',     service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Overdose/toxicological emergency requiring ALS monitoring' },
      'ENDOCRINE-METABOLIC':  { dets: ['PRI-1','PRI-2'], clinical_group: 'ENDOCRINE',      service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'DKA, HHS, severe hypo/hyperglycemia, electrolyte crisis' },
      'TRAUMA':               { dets: ['PRI-1','PRI-2'], clinical_group: 'TRAUMA',         service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'ALS trauma transfer (hemodynamically stable or borderline)' },
      'BURN':                 { dets: ['PRI-1','PRI-2'], clinical_group: 'TRAUMA',         service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: false, desc: 'Burn transfer not requiring CCT (moderate, stable airway)' },
      // ── LEGACY ──
      'CARDIAC':      { dets: ['PRI-1','PRI-2'], clinical_group: 'CARDIOVASCULAR', service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: true, desc: '[LEGACY] Cardiac instability — use CHEST-PAIN or STEMI' },
      'NEURO-STROKE': { dets: ['PRI-1','PRI-2'], clinical_group: 'NEUROLOGICAL',   service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: true, desc: '[LEGACY] Stroke/neuro — use STROKE' },
      'RESPIRATORY':  { dets: ['PRI-1','PRI-2'], clinical_group: 'RESPIRATORY',    service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: true, desc: '[LEGACY] Respiratory — use RESPIRATORY-DISTRESS' },
      'OB':           { dets: ['PRI-1','PRI-2'], clinical_group: 'OB',             service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE', legacy: true, desc: '[LEGACY] OB transfer — use OB-COMPLICATION' },
    }
  },
  'IFT-BLS': {
    natures: {
      'POST-OP':           { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Stable post-operative transfer' },
      'BASIC-MEDICAL':     { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Stable medical transfer, BLS-appropriate' },
      'DIAGNOSTIC':        { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Transfer for imaging, procedure, or diagnostic workup' },
      'PSYCH-STABLE':      { dets: ['PRI-3'], clinical_group: 'PSYCH',   service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Behavioral health / psychiatric transfer (stable, no acute medical need)' },
      'WOUND-CARE':        { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Wound care, dressing change, stable surgical follow-up' },
      'FALL-NO-INJURY':    { dets: ['PRI-3'], clinical_group: 'TRAUMA',  service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Fall with no acute injury, cleared for BLS transport' },
      'FACILITY-TRANSFER': { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Facility-to-facility transfer, non-specific stable' },
      'HOSPICE':           { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Hospice / comfort care transport' },
      'LTACH':             { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Long-term acute care hospital admission transfer' },
      'MEMORY-CARE':       { dets: ['PRI-3'], clinical_group: 'PSYCH',   service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Memory care / dementia unit transfer' },
      'ASSISTED-LIVING':   { dets: ['PRI-3'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'STABLE', legacy: false, desc: 'Assisted living or residential care facility transfer' },
      // ── LEGACY ──
      'PSYCH':         { dets: ['PRI-3'], clinical_group: 'PSYCH', service_level: 'BLS', clinical_severity: 'STABLE', legacy: true, desc: '[LEGACY] Psych transfer — use PSYCH-STABLE' },
    }
  },
  DISCHARGE: {
    natures: {
      'STRETCHER':  { dets: ['PRI-4'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Discharge requiring stretcher' },
      'WHEELCHAIR': { dets: ['PRI-4'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Discharge via wheelchair transport' },
      'AMBULATORY': { dets: ['PRI-4'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Ambulatory discharge transport' },
      'HOME':       { dets: ['PRI-4'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Discharge to home' },
      'REHAB':      { dets: ['PRI-4'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Discharge to inpatient rehab facility' },
      'SNF':        { dets: ['PRI-4'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Discharge to skilled nursing facility' },
    }
  },
  DIALYSIS: {
    natures: {
      'ROUTINE':      { dets: ['PRI-4'],         clinical_group: 'ENDOCRINE', service_level: 'BLS',  clinical_severity: 'ROUTINE',        legacy: false, desc: 'Scheduled dialysis transport' },
      'MISSED-TX':    { dets: ['PRI-3'],         clinical_group: 'ENDOCRINE', service_level: 'BLS',  clinical_severity: 'STABLE',         legacy: false, desc: 'Missed treatment — rescheduled, stable fluid overload' },
      'EMERGENT':     { dets: ['PRI-2','PRI-3'], clinical_group: 'ENDOCRINE', service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE',  legacy: false, desc: 'Urgent dialysis need — fluid overload, uremic symptoms' },
      'HYPERKALEMIA': { dets: ['PRI-2'],         clinical_group: 'ENDOCRINE', service_level: 'ALS1', clinical_severity: 'TIME_SENSITIVE',  legacy: false, desc: 'Symptomatic hyperkalemia requiring urgent dialysis + ALS cardiac monitoring' },
      'RETURN':       { dets: ['PRI-4'],         clinical_group: 'ENDOCRINE', service_level: 'BLS',  clinical_severity: 'ROUTINE',        legacy: false, desc: 'Return trip after completed dialysis session' },
    }
  },
  // ADMIN: resolution/exception codes — NOT shown in new incident picker, only in edit modal
  ADMIN: {
    natures: {
      'CANCELLED-PRE-DISPATCH':   { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Call cancelled before any unit was dispatched' },
      'CANCELLED-ON-SCENE':       { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Unit arrived; patient declined or call cancelled on scene' },
      'NO-PATIENT':               { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Unit arrived; no patient found at scene' },
      'DUPLICATE':                { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Incident created in error; merged with or superseded by another call' },
      'WEATHER-DELAY':            { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Transport delayed or cancelled due to weather conditions' },
      'ACCEPTING-FACILITY-DELAY': { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Accepting facility not ready; call held or rescheduled' },
      'NO-UNIT-AVAILABLE':        { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'No unit available at time of call; mutual aid requested or deferred' },
      'WAITING-BED':              { dets: ['NONE'], clinical_group: 'GENERAL', service_level: 'BLS', clinical_severity: 'ROUTINE', legacy: false, desc: 'Call held — destination has no available bed' },
    }
  },
};

// OLD→NEW taxonomy migration: maps old type strings to new equivalents for display/reporting
const INC_TYPE_MIGRATION = {
  // Pre-Supabase legacy prefixes
  'MED-CARDIAC': 'IFT-ALS-CHEST-PAIN', 'MED-STROKE': 'IFT-ALS-STROKE',
  'MED-RESPIRATORY': 'IFT-ALS-RESPIRATORY-DISTRESS', 'MED-SEPSIS': 'IFT-ALS-SEPSIS',
  'MED-OB': 'IFT-ALS-OB-COMPLICATION', 'MED-': 'IFT-BLS-BASIC-MEDICAL',
  'CCT-CARDIAC-DRIP': 'CCT-CARDIAC-CRITICAL', 'CCT-ICU': 'CCT-MULTI-SYSTEM-FAILURE',
  'CCT-TRAUMA': 'CCT-TRAUMA-CRITICAL', 'CCT-MISC': 'CCT-MULTI-SYSTEM-FAILURE',
  'TRAUMA-': 'IFT-ALS-TRAUMA',
  'DISCHARGE-SNF': 'DISCHARGE-SNF', 'DISCHARGE-': 'DISCHARGE-STRETCHER',
  // IFT-ALS renames (device/vague → clinical body-system)
  'IFT-ALS-CARDIAC':      'IFT-ALS-CHEST-PAIN',
  'IFT-ALS-NEURO-STROKE': 'IFT-ALS-STROKE',
  'IFT-ALS-RESPIRATORY':  'IFT-ALS-RESPIRATORY-DISTRESS',
  'IFT-ALS-OB':           'IFT-ALS-OB-COMPLICATION',
  // IFT-BLS renames
  'IFT-BLS-PSYCH':        'IFT-BLS-PSYCH-STABLE',
  // CCT renames (device-specific → clinical)
  'CCT-VENT':             'CCT-RESPIRATORY-FAILURE',
  'CCT-MULTI-DRIP':       'CCT-MULTI-SYSTEM-FAILURE',
  'CCT-CRITICAL-TRAUMA':  'CCT-TRAUMA-CRITICAL',
  'CCT-ECMO':             'CCT-CARDIAC-CRITICAL',
  'CCT-NICU-PICU':        'CCT-PEDIATRIC-CRITICAL',
  'CCT-HIGH-RISK-AIRWAY': 'CCT-RESPIRATORY-FAILURE',
};

// Border colors indexed by getIncidentTypeClass result (4B)
const INC_GROUP_BORDER = {
  'inc-type-delta':    '#ff4444',  // PRI-1 / CCT
  'inc-type-charlie':  '#ff6600',  // PRI-2 / IFT-ALS
  'inc-type-bravo':    '#ffd700',  // PRI-3 / IFT-BLS
  'inc-type-alpha':    '#4fa3e0',  // PRI-4 / DISCHARGE / DIALYSIS
  'inc-type-discharge':'#6a7a8a',
  'inc-type-other':    '#6a7a8a',
};

// LifeFlight Network fleet registry — tail# → callsign/board unit mapping
// Note: LFN may reassign aircraft to bases dynamically; update as confirmed.
// Combined air resource fleet registry — all providers, keyed by tail number
const AIR_FLEET = {
  // LifeFlight Network (LFN) — Redmond base
  'N429LF': { callsign: 'LF11',  unitId: 'LF11', type: 'HELI', provider: 'LFN'     },
  'N430LF': { callsign: 'LF45',  unitId: 'LF45', type: 'HELI', provider: 'LFN'     },
  'N450LF': { callsign: 'LFH1',  unitId: null,   type: 'HELI', provider: 'LFN'     },
  'N451LF': { callsign: 'LFH2',  unitId: null,   type: 'HELI', provider: 'LFN'     },
  'N452LF': { callsign: 'LFH3',  unitId: null,   type: 'HELI', provider: 'LFN'     },
  'N453LF': { callsign: 'LFH4',  unitId: null,   type: 'HELI', provider: 'LFN'     },
  'N661LF': { callsign: 'LFPC1', unitId: null,   type: 'PC12', provider: 'LFN'     },
  'N662LF': { callsign: 'LFPC2', unitId: null,   type: 'PC12', provider: 'LFN'     },
  'N866LF': { callsign: 'LFPC3', unitId: null,   type: 'PC12', provider: 'LFN'     },
  // AirLink CCT (MTC) — rotor + fixed wing
  'N885AL': { callsign: 'AL1',   unitId: 'AL1',  type: 'HELI', provider: 'AIRLINK' },
  'N853AL': { callsign: 'AL2',   unitId: 'AL2',  type: 'FW',   provider: 'AIRLINK' },
  'N880GT': { callsign: 'AL3',   unitId: 'AL3',  type: 'HELI', provider: 'AIRLINK' },
  'N852AL': { callsign: 'AL4',   unitId: 'AL4',  type: 'FW',   provider: 'AIRLINK' },
};
// Hospital helipads and landing zones. scmc:true = eligible for aircraft INBOUND alert.
const LFN_HELIPADS = [
  // SCMC campuses — primary receiving hospitals, trigger INBOUND alert
  { code: 'SCMC-BC', name: 'SCMC Bend',              lat: 44.0672, lon: -121.2690, type: 'ROOF',    scmc: true },
  { code: 'SCMC-RC', name: 'SCMC Redmond',           lat: 44.2704, lon: -121.1417, type: 'GROUND',  scmc: true },
  { code: 'SCMC-PC', name: 'SCMC Prineville',        lat: 44.2980, lon: -120.8253, type: 'ROOF',    scmc: true },
  { code: 'SCMC-MC', name: 'SCMC Madras',            lat: 44.6329, lon: -121.1298, type: 'GROUND',  scmc: true },
  // Regional airports — excluded from map markers and INBOUND alert
  { code: 'KRDM',    name: 'Redmond Airport',        lat: 44.2542, lon: -121.1486, type: 'AIRPORT' },
  { code: 'KBDN',    name: 'Bend Airport',           lat: 44.0955, lon: -121.2003, type: 'AIRPORT' },
  // Regional hospital helipads — map markers only
  { code: 'SKY-LAKES',  name: 'Sky Lakes MC',                   lat: 42.2530, lon: -121.7851, type: 'ROOF'   },
  { code: 'MCMC',       name: 'Mid-Columbia MC (The Dalles)',   lat: 45.5980, lon: -121.1525, type: 'ROOF'   },
  { code: 'WS-IHS',     name: 'Warm Springs IHS',              lat: 44.7636, lon: -121.2733, type: 'GROUND' },
  { code: 'HARNEY-DH',  name: 'Harney District Hospital',      lat: 43.5888, lon: -119.0601, type: 'ROOF'   },
  { code: 'BLUE-MTN',   name: 'Blue Mountain Hospital',        lat: 44.4171, lon: -118.9576, type: 'ROOF'   },
  { code: 'LAKE-DIST',  name: 'Lake District Hospital',        lat: 42.1818, lon: -120.3515, type: 'ROOF'   },
  { code: 'SALEM-HOSP', name: 'Salem Hospital',                lat: 44.9363, lon: -123.0351, type: 'ROOF'   },
  { code: 'OHSU',       name: 'OHSU Portland',                 lat: 45.4991, lon: -122.6870, type: 'ROOF'   },
];
const LFN_BASE_LAT = 44.254;
const LFN_BASE_LON = -121.150;
const LFN_POLL_INTERVAL  = 30 * 1000; // 30-second interval — aircraft status can change quickly
const LFN_QUERY_RADIUS   = 75;        // nautical miles radius for adsb.lol query


// Command hints for autocomplete
const CMD_HINTS = [
  { cmd: 'BOARDS', desc: 'BOARDS — re-open unit status board + incident queue windows' },
  { cmd: 'D <UNIT>; <NOTE>', desc: 'Dispatch unit' },
  { cmd: 'DE <UNIT>; <NOTE>', desc: 'Set enroute' },
  { cmd: 'OS <UNIT>; <NOTE>', desc: 'Set on scene' },
  { cmd: 'T <UNIT>; <NOTE>', desc: 'Set transporting' },
  { cmd: 'TH <UNIT>', desc: 'AT HOSPITAL — crew with patient at facility' },
  { cmd: 'AV <UNIT>', desc: 'Set available' },
  { cmd: 'OOS <UNIT>; <NOTE>', desc: 'Set out of service' },
  { cmd: 'BRK <UNIT>; <NOTE>', desc: 'Set on break' },
  { cmd: 'F <STATUS>', desc: 'Filter board by status' },
  { cmd: 'V SIDE', desc: 'Toggle sidebar' },
  { cmd: 'V INC', desc: 'Toggle incident queue' },
  { cmd: 'V MSG', desc: 'Toggle messages' },
  { cmd: 'SORT STATUS', desc: 'Sort by status' },
  { cmd: 'SORT ELAPSED', desc: 'Sort by elapsed time' },
  { cmd: 'DEN', desc: 'Cycle density mode' },
  { cmd: 'NIGHT', desc: 'Toggle night mode' },
  { cmd: 'NC <LOCATION>; <NOTE>; <TYPE>; <PRIORITY>; @<SCENE ADDR>', desc: 'New incident (add MA in note for mutual aid, [CB:PHONE] in note for callback, PRIORITY e.g. PRI-1, @ADDR for scene address)' },
  { cmd: 'HOLD <DEST>; <HH:MM>; [NOTE]; [TYPE]; [PRIORITY]', desc: 'Schedule a call for later — creates QUEUED with hold clock badge. Alerts when time arrives.' },
  { cmd: 'COPY <INC>', desc: 'Duplicate incident into new QUEUED call (copies dest, type, scene, priority; strips tags from note)' },
  { cmd: 'COPY', desc: 'Duplicate currently-open incident into new QUEUED call' },
  { cmd: 'R <INC>', desc: 'Review incident' },
  { cmd: 'U <INC> <NOTE>', desc: 'Append note to incident (e.g. U 0023 PT STABLE IN WTRM)' },
  { cmd: 'U <NOTE>', desc: 'Append note to currently-open incident (no INC ID needed)' },
  { cmd: 'CB <INC> <PHONE>', desc: 'Set/update callback number on incident (e.g. CB 0023 5415551234)' },
  { cmd: 'CB <PHONE>', desc: 'Set callback on currently-open incident' },
  { cmd: 'RQ <INC>', desc: 'Requeue incident (QUEUED, clears unit assignment — for reassignment)' },
  { cmd: 'RO <INC>', desc: 'Reopen closed incident (ACTIVE, keeps existing units)' },
  { cmd: 'UH <UNIT> [HOURS]', desc: 'Unit history (alias: HIST)' },
  { cmd: 'MSG <ROLE/UNIT>; <TEXT>', desc: 'Send message' },
  { cmd: 'PG <UNIT>', desc: 'Radio page unit (plays fire/EMS tone on field device)' },
  { cmd: 'WELF <UNIT>', desc: 'Welfare check — sends urgent message asking unit to confirm status' },
  { cmd: 'MA <INC> <AGENCY>', desc: 'Request mutual aid (e.g. MA 0001 BEND FIRE). MA ACK / MA REL to update status.' },
  { cmd: 'GPS <UNIT>', desc: 'Show unit on board map using current known position' },
  { cmd: 'GPSUL <UNIT>', desc: 'Request unit to ping their GPS location to the board' },
  { cmd: 'MSGDP; <TEXT>', desc: 'Message all dispatchers' },
  { cmd: 'HTDP; <TEXT>', desc: 'URGENT message all dispatchers' },
  { cmd: 'MSGU; <TEXT>', desc: 'Message all active field units' },
  { cmd: 'HTU; <TEXT>', desc: 'URGENT message all field units' },
  { cmd: 'DEST <UNIT>; <LOCATION> [NOTE]', desc: 'Set unit location/destination' },
  { cmd: 'L <UNIT>; <NOTE>', desc: 'Logon unit (LOGON also works)' },
  { cmd: 'LO <UNIT>', desc: 'Logoff unit (LOGOFF also works)' },
  { cmd: 'PRESET DISPATCH', desc: 'Dispatch view preset' },
  { cmd: 'CLR', desc: 'Clear all filters' },
  { cmd: 'INFO', desc: 'Quick reference (key numbers)' },
  { cmd: 'INFO ALL', desc: 'Full dispatch/emergency directory' },
  { cmd: 'INFO DISPATCH', desc: '911/PSAP dispatch centers' },
  { cmd: 'INFO AIR', desc: 'Air ambulance dispatch' },
  { cmd: 'INFO CRISIS', desc: 'Mental health / crisis lines' },
  { cmd: 'INFO LE', desc: 'Law enforcement direct lines' },
  { cmd: 'INFO FIRE', desc: 'Fire department admin / BC' },
  { cmd: 'ADDR', desc: 'Address directory / search' },
  { cmd: 'ADMIN', desc: 'Admin commands (SUPV/MGR/IT only)' },
  { cmd: 'REPORT SHIFT [12]', desc: 'Printable shift summary (hours, default 12)' },
  { cmd: 'REPORT INC <ID>',   desc: 'Printable per-incident report' },
  { cmd: 'REPORTUTIL <UNIT> [24]', desc: 'Per-unit utilization report (hours, default 24)' },
  { cmd: 'WHO',               desc: 'Dispatchers currently online' },
  { cmd: 'WHO [ID]',          desc: 'Look up unit crew or dispatcher by role/name' },
  { cmd: 'UR',                desc: 'Active unit roster' },
  { cmd: 'INCQ',              desc: 'Incident queue quick view (queued + active)' },
  { cmd: 'SUGGEST <INC>',     desc: 'Recommend available units for incident' },
  { cmd: 'DIVERSION ON <CODE>',  desc: 'Set hospital/facility on diversion' },
  { cmd: 'DIVERSION OFF <CODE>', desc: 'Clear hospital/facility diversion' },
  { cmd: 'STACK <INC> <UNIT>',   desc: 'Smart-assign: primary if unit free, queued if unit busy' },
  { cmd: 'ASSIGN <INC> <UNIT>',  desc: 'Force primary assignment (displaces current primary to queued)' },
  { cmd: 'QUEUE / QUE <INC> <UNIT>', desc: 'Add incident to unit queue behind primary (explicit)' },
  { cmd: 'PRIMARY <INC> <UNIT>', desc: 'Promote queued assignment to primary' },
  { cmd: 'CLEAR <INC> <UNIT>',   desc: 'Remove assignment from unit stack' },
  { cmd: 'STACK <UNIT>',         desc: 'Show unit assignment stack' },
  { cmd: 'SCOPE ALL',            desc: 'View all agencies (SUPV/MGR/IT only)' },
  { cmd: 'SCOPE AGENCY <ID>',    desc: 'Limit view to one agency' },
  { cmd: 'MSGALL <TEXT>', desc: 'Broadcast to all dispatchers + units' },
  { cmd: 'HTALL <TEXT>', desc: 'URGENT broadcast to all' },
  { cmd: 'NOTE <MESSAGE>', desc: 'Set info banner' },
  { cmd: 'NOTE CLEAR', desc: 'Clear info banner' },
  { cmd: 'ALERT <MESSAGE>', desc: 'Set alert banner (plays tone)' },
  { cmd: 'ALERT CLEAR', desc: 'Clear alert banner' },
  { cmd: 'CLR <UNIT>', desc: 'Clear unit from incident (no status change)' },
  { cmd: 'ETA <UNIT> <MINUTES>', desc: 'Set ETA for unit (e.g. ETA EMS1 8)' },
  { cmd: 'ETA <UNIT> CLR', desc: 'Clear ETA badge from unit' },
  { cmd: 'NEXT <UNIT>', desc: 'Advance unit one step in EMS chain: D→DE→OS→T→TH→AV' },
  { cmd: 'PAT <UNIT> <TEXT>', desc: 'Set patient info badge on unit (e.g. PAT EMS3 2PTS CHEST PAIN)' },
  { cmd: 'PAT <UNIT> CLR', desc: 'Clear patient info badge from unit' },
  { cmd: 'PRIORITY <INC> <PRI>', desc: 'Update incident priority (e.g. PRIORITY 0023 PRI-1)' },
  { cmd: 'STATS', desc: 'Live board summary (units, incidents)' },
  { cmd: 'SHIFT END <UNIT>', desc: 'End shift: set AV, clear assignments, deactivate' },
  { cmd: 'LINK <U1> <U2> <INC>', desc: 'Assign both units to incident' },
  { cmd: 'TRANSFER <FROM> <TO> <INC>', desc: 'Transfer incident between units' },
  { cmd: 'AVALL <INC#>', desc: 'Set all units on an incident to AV (e.g. AVALL 0071)' },
  { cmd: 'OSALL <INC#>', desc: 'Set all units on an incident to OS (e.g. OSALL 0071)' },
  { cmd: 'MASS D <DEST> CONFIRM', desc: 'Dispatch all AV units (requires CONFIRM)' },
  { cmd: 'LUI [UNIT]', desc: 'Create temp one-off unit (SUPV/MGR/IT only)' },
  { cmd: 'REL <INC1> <INC2>', desc: 'Link two HOSCAD incidents together (REL 26-0006 26-0007)' },
  { cmd: 'HELP', desc: 'Show command reference' },
  { cmd: 'POPOUT', desc: 'Open status board on secondary monitor' },
  { cmd: 'POPIN', desc: 'Restore status board to this screen' },
  { cmd: 'MAP', desc: 'Toggle map panel on board' },
  { cmd: 'MAP <UNIT>', desc: 'Focus map on unit location (opens map if closed)' },
  { cmd: 'MAP <INC>', desc: 'Focus map on incident scene address' },
  { cmd: 'LOC <UNIT> <ADDR>', desc: 'Set unit location for map (chain status: LOC M1 123 MAIN; AV M1)' },
  { cmd: 'LOC <UNIT> CLR', desc: 'Clear unit location' },
  { cmd: 'MAP <ADDRESS>', desc: 'Geocode address and focus map on it' },
  { cmd: 'MAPR', desc: 'Quick map refresh (re-render all markers)' },
  { cmd: 'MAP IN/OUT/FIT/STA/CLR/RESET', desc: 'Map zoom/view controls' },
  { cmd: 'POPMAP', desc: 'Pop map out into its own window' },
  { cmd: 'POPINC', desc: 'Pop incident queue into its own window' },
  { cmd: 'BUG', desc: 'Report a bug or system issue (opens form)' },
];
let CMD_HINT_INDEX = -1;

// ============================================================
// Address History Module — per-user localStorage autocomplete
// Key: hoscad_addr_<actor>  Max: 50 entries, deduped, newest first
// ============================================================
const AddrHistory = {
  _key() { return 'hoscad_addr_' + (ACTOR || 'anon'); },
  get() {
    try { return JSON.parse(localStorage.getItem(this._key()) || '[]'); } catch(e) { return []; }
  },
  push(addr) {
    if (!addr || addr.length < 4) return;
    const a = addr.trim().toUpperCase();
    let list = this.get().filter(x => x !== a);
    list.unshift(a);
    if (list.length > 50) list = list.slice(0, 50);
    try { localStorage.setItem(this._key(), JSON.stringify(list)); } catch(e) {}
    this._refresh();
  },
  _refresh() {
    const list = this.get();
    document.querySelectorAll('.addr-history-list').forEach(dl => {
      dl.innerHTML = list.map(a => '<option value="' + a.replace(/"/g, '&quot;') + '">').join('');
    });
  },
  attach(inputId, datalistId) {
    const inp = document.getElementById(inputId);
    const dl = document.getElementById(datalistId);
    if (!inp || !dl) return;
    dl.className = 'addr-history-list';
    inp.setAttribute('list', datalistId);
    inp.setAttribute('autocomplete', 'off');
    this._refresh();
  },
};

// ============================================================
// Address Lookup Module
// ============================================================
const AddressLookup = {
  _cache: [],
  _loaded: false,

  async load() {
    if (!TOKEN) return;
    try {
      const r = await API.getAddresses(TOKEN);
      if (r && r.ok && r.addresses) {
        this._cache = r.addresses;
        this._loaded = true;
      }
    } catch (e) {
      console.error('[AddressLookup] Load failed:', e);
    }
  },

  getById(id) {
    if (!id) return null;
    const u = String(id).trim().toUpperCase();
    return this._cache.find(a => a.id === u) || null;
  },

  search(query, limit) {
    limit = limit || 8;
    if (!query || query.length < 2) return [];
    const q = String(query).trim().toLowerCase();
    const exact = [];
    const starts = [];
    const contains = [];

    for (let i = 0; i < this._cache.length; i++) {
      const a = this._cache[i];
      const idL = a.id.toLowerCase();
      const nameL = a.name.toLowerCase();
      const aliases = a.aliases || [];

      // Exact alias/id match
      if (idL === q || aliases.indexOf(q) >= 0) {
        exact.push(a);
        continue;
      }

      // Starts-with on id, name, aliases
      if (idL.indexOf(q) === 0 || nameL.indexOf(q) === 0 || aliases.some(function(al) { return al.indexOf(q) === 0; })) {
        starts.push(a);
        continue;
      }

      // Contains in id, name, aliases, address, city
      const addressL = (a.address || '').toLowerCase();
      const cityL = (a.city || '').toLowerCase();
      if (idL.indexOf(q) >= 0 || nameL.indexOf(q) >= 0 ||
          aliases.some(function(al) { return al.indexOf(q) >= 0; }) ||
          addressL.indexOf(q) >= 0 || cityL.indexOf(q) >= 0) {
        contains.push(a);
      }
    }

    return exact.concat(starts, contains).slice(0, limit);
  },

  resolve(destValue) {
    if (!destValue) return { recognized: false, addr: null, displayText: '' };
    const v = String(destValue).trim().toUpperCase();
    const addr = this.getById(v);
    if (addr) {
      return { recognized: true, addr: addr, displayText: addr.name };
    }
    return { recognized: false, addr: null, displayText: v };
  },

  /** Parse bracket/paren note from a value: "SCMC [BED 4]" or "123 MAIN ST (BED2)" → { base, note } */
  _parseBracketNote(val) {
    if (!val) return { base: '', note: '' };
    // Check for [note] first, then (note) — used by LOC tags which store bracket notes as parens
    let idx = val.indexOf('[');
    let close = ']';
    if (idx < 0) { idx = val.lastIndexOf('('); close = ')'; }
    if (idx < 0) return { base: val, note: '' };
    const base = val.substring(0, idx).trim();
    let note = val.substring(idx + 1).trim();
    if (note.endsWith(close)) note = note.substring(0, note.length - 1).trim();
    return { base: base || val, note };
  },

  _noteBadge(note) {
    return note ? ' ' + esc(note) : '';
  },

  formatBoard(destValue) {
    if (!destValue) return '<span class="muted">\u2014</span>';
    const v = String(destValue).trim().toUpperCase();
    const { base, note } = this._parseBracketNote(v);
    const lookupKey = base || v;
    const addr = this.getById(lookupKey);
    const destObj = (STATE.destinations || []).find(d => d.code === lookupKey);
    const divBadge = destObj && destObj.diverted ? ' <span class="div-badge">DIV</span>' : '';
    const nBadge = this._noteBadge(note);
    if (addr) {
      const tip = esc(addr.address + ', ' + addr.city + ', ' + addr.state + ' ' + addr.zip + (note ? ' [' + note + ']' : ''));
      return '<span class="dest-recognized destBig" title="' + tip + '">' + esc(addr.name) + '</span>' + nBadge + divBadge;
    }
    return '<span class="destBig">' + esc(lookupKey || '\u2014') + '</span>' + nBadge + divBadge;
  },

  /**
   * Format contextual LOCATION column for a unit on the board.
   * D/DE/OS → scene address from incident
   * T/AT/TH → transport destination (hospital)
   * Fallback → destination field or [LOC:] tag or dash
   */
  formatLocation(unit) {
    const st = String(unit.status || '').toUpperCase();
    // Dispatched/on-scene: show incident scene address
    if (['D', 'DE', 'OS'].includes(st) && unit.incident && STATE) {
      const inc = (STATE.incidents || []).find(i => i.incident_id === unit.incident);
      if (inc && inc.scene_address) {
        const raw = String(inc.scene_address).trim().toUpperCase();
        const { base, note } = this._parseBracketNote(raw);
        // No JS truncation — CSS text-overflow:ellipsis clips to column width
        return '<span class="destBig" title="' + esc(raw) + '">' + esc(base) + '</span>' + this._noteBadge(note);
      }
    }
    // Transporting/at hospital: show destination (hospital)
    if (['T', 'AT', 'TH'].includes(st) && unit.destination) {
      return this.formatBoard(unit.destination);
    }
    // In Quarters: show station name
    if (st === 'IQ' && unit.station) {
      const staName = String(unit.station).trim().toUpperCase();
      return '<span class="destBig" title="IN QUARTERS">' + esc(staName) + '</span>';
    }
    // Fallback: destination if set, or incident address (when assigned), or [LOC:] GPS (only without incident), or dash
    if (unit.destination) return this.formatBoard(unit.destination);
    // When assigned to an incident, always prefer incident scene/destination over GPS tag
    if (unit.incident && STATE) {
      const inc = (STATE.incidents || []).find(i => i.incident_id === unit.incident);
      if (inc) {
        if (inc.scene_address) {
          const raw = String(inc.scene_address).trim().toUpperCase();
          const { base, note } = this._parseBracketNote(raw);
          return '<span class="destBig" title="' + esc(raw) + '">' + esc(base) + '</span>' + this._noteBadge(note);
        }
        if (inc.destination) return this.formatBoard(inc.destination);
      }
    }
    // [LOC:] GPS tag — only shown when unit is NOT assigned to an active incident
    const locTag = (unit.note || '').match(/\[LOC:([^\]]+)\]/);
    if (locTag && !unit.incident) {
      const raw = locTag[1].trim().toUpperCase();
      const { base, note } = this._parseBracketNote(raw);
      return '<span class="destBig" title="' + esc(raw) + '">' + esc(base) + '</span>' + this._noteBadge(note);
    }
    return '<span class="muted">\u2014</span>';
  }
};

// ============================================================
// Address Autocomplete Component
// ============================================================
const AddrAutocomplete = {
  attach(inputEl, options) {
    if (!inputEl || inputEl.dataset.acAttached) return;
    inputEl.dataset.acAttached = '1';
    var onSelect = options && options.onSelect;

    // Wrap input in relative container
    const wrapper = document.createElement('div');
    wrapper.className = 'addr-ac-wrapper';
    inputEl.parentNode.insertBefore(wrapper, inputEl);
    wrapper.appendChild(inputEl);

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.className = 'addr-ac-dropdown';
    wrapper.appendChild(dropdown);

    let acIndex = -1;
    let acResults = [];

    function showDropdown(results) {
      acResults = results;
      acIndex = -1;
      if (!results.length) {
        dropdown.classList.remove('open');
        dropdown.innerHTML = '';
        return;
      }
      dropdown.innerHTML = results.map(function(a, i) {
        return '<div class="addr-ac-item" data-idx="' + i + '">' +
          '<span class="addr-ac-id">' + esc(a.id) + '</span>' +
          '<span class="addr-ac-name">' + esc(a.name) + '</span>' +
          '<span class="addr-ac-detail">\u2014 ' + esc(a.address + ', ' + a.city) + '</span>' +
          '<span class="addr-ac-cat">' + esc((a.category || '').replace(/_/g, ' ')) + '</span>' +
          '</div>';
      }).join('');
      dropdown.classList.add('open');
    }

    function hideDropdown() {
      dropdown.classList.remove('open');
      dropdown.innerHTML = '';
      acResults = [];
      acIndex = -1;
    }

    function selectItem(idx) {
      if (idx < 0 || idx >= acResults.length) return;
      var a = acResults[idx];
      inputEl.value = a.name;
      inputEl.dataset.addrId = a.id;
      hideDropdown();
      if (onSelect) onSelect(a);
    }

    function highlightItem(idx) {
      var items = dropdown.querySelectorAll('.addr-ac-item');
      items.forEach(function(el) { el.classList.remove('active'); });
      if (idx >= 0 && idx < items.length) {
        items[idx].classList.add('active');
        items[idx].scrollIntoView({ block: 'nearest' });
      }
    }

    inputEl.addEventListener('input', function() {
      delete inputEl.dataset.addrId;
      var val = inputEl.value.trim();
      if (val.length < 2) {
        hideDropdown();
        return;
      }
      var results = AddressLookup.search(val);
      showDropdown(results);
    });

    inputEl.addEventListener('keydown', function(e) {
      if (!dropdown.classList.contains('open')) return;

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        acIndex = Math.min(acIndex + 1, acResults.length - 1);
        highlightItem(acIndex);
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        acIndex = Math.max(acIndex - 1, 0);
        highlightItem(acIndex);
      } else if (e.key === 'Enter') {
        if (acIndex >= 0) {
          e.preventDefault();
          selectItem(acIndex);
        } else {
          hideDropdown();
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        hideDropdown();
      }
    });

    inputEl.addEventListener('blur', function() {
      setTimeout(hideDropdown, 150);
    });

    dropdown.addEventListener('mousedown', function(e) {
      e.preventDefault(); // Prevent blur
      var item = e.target.closest('.addr-ac-item');
      if (item) {
        var idx = parseInt(item.dataset.idx);
        selectItem(idx);
      }
    });
  }
};

// ============================================================
// View State Persistence
// ============================================================
function loadViewState() {
  try {
    const saved = localStorage.getItem('hoscad_view');
    if (saved) {
      const parsed = JSON.parse(saved);
      Object.assign(VIEW, parsed);
      // Migrate legacy nightMode boolean → VIEW.theme
      if (parsed.nightMode === true && !parsed.theme) VIEW.theme = 'night';
    }
  } catch (e) { }
}

function saveViewState() {
  try {
    localStorage.setItem('hoscad_view', JSON.stringify(VIEW));
  } catch (e) { }
}

function applyViewState() {
  // Side panel
  const sp = document.getElementById('sidePanel');
  if (sp) {
    if (VIEW.sidebar) sp.classList.add('open');
    else sp.classList.remove('open');
  }

  // Bottom panels: hide when sidebar is open (messages + scratch move to sidebar)
  const bp = document.querySelector('.bottom-panels');
  if (bp) bp.style.display = VIEW.sidebar ? 'none' : '';

  // Sync scratch notes between bottom pad and side pad when transitioning
  const _scratchMain = document.getElementById('scratchPad');
  const _scratchSide = document.getElementById('scratchPadSide');
  if (VIEW.sidebar) {
    if (_scratchMain && _scratchSide) _scratchSide.value = _scratchMain.value;
  } else {
    if (_scratchMain && _scratchSide) _scratchMain.value = _scratchSide.value;
  }

  // Incident queue
  const iq = document.getElementById('incidentQueueCard');
  if (iq) {
    if (VIEW.incidents) iq.classList.remove('collapsed');
    else iq.classList.add('collapsed');
    iq.style.display = VIEW.incidents ? '' : '';
  }

  // Messages section in sidebar — always show when sidebar is open; only hide when sidebar is closed AND VIEW.messages is off
  const ms = document.getElementById('sideMsgSection');
  if (ms) ms.style.display = (VIEW.sidebar || VIEW.messages) ? '' : 'none';
  if (VIEW.sidebar) renderMessagesPanel();

  // Density
  const wrap = document.querySelector('.wrap');
  if (wrap) {
    wrap.classList.remove('density-compact', 'density-normal', 'density-expanded');
    wrap.classList.add('density-' + VIEW.density);
  }
  // Relay density to viewer popout so board rows scale accordingly
  if (_popoutBoardWindow && !_popoutBoardWindow.closed) {
    try { _popoutBoardWindow.postMessage({ type: 'HOSCAD_DENSITY', density: VIEW.density }, window.location.origin); } catch(e) {}
  }

  // Theme: apply data-theme attribute and legacy night-mode class
  const theme = VIEW.theme || (VIEW.nightMode ? 'night' : 'dark');
  document.body.setAttribute('data-theme', theme);
  if (theme === 'night') document.body.classList.add('night-mode');
  else document.body.classList.remove('night-mode');

  // Night/theme button state + label
  const nightBtn = document.getElementById('tbBtnNight');
  if (nightBtn) {
    const labels = { dark: 'DARK', night: 'NIGHT', light: 'LIGHT' };
    nightBtn.textContent = labels[theme] || 'DARK';
    if (theme !== 'dark') nightBtn.classList.add('active');
    else nightBtn.classList.remove('active');
  }

  // Toolbar button states
  updateToolbarButtons();

  // Toolbar dropdowns
  const tbFs = document.getElementById('tbFilterStatus');
  if (tbFs) tbFs.value = VIEW.filterStatus || '';

  const tbSort = document.getElementById('tbSort');
  if (tbSort) tbSort.value = VIEW.sort || 'status';

  // Column sort indicators
  updateSortHeaders();

}

function updateToolbarButtons() {
  const btns = {
    'tbBtnINC': VIEW.incidents,
    'tbBtnSIDE': VIEW.sidebar,
    'tbBtnMSG': VIEW.messages
  };
  for (const [id, active] of Object.entries(btns)) {
    const el = document.getElementById(id);
    if (el) {
      if (active) el.classList.add('active');
      else el.classList.remove('active');
    }
  }

  const denBtn = document.getElementById('tbBtnDEN');
  if (denBtn) denBtn.textContent = 'DEN: ' + VIEW.density.toUpperCase();
}

function updateSortHeaders() {
  document.querySelectorAll('.board-table th.sortable').forEach(th => {
    th.classList.remove('sort-active', 'sort-desc');
    if (th.dataset.sort === VIEW.sort) {
      th.classList.add('sort-active');
      if (VIEW.sortDir === 'desc') th.classList.add('sort-desc');
    }
  });
}

function toggleView(panel) {
  if (panel === 'sidebar' || panel === 'side') {
    VIEW.sidebar = !VIEW.sidebar;
  } else if (panel === 'incidents' || panel === 'inc') {
    VIEW.incidents = !VIEW.incidents;
  } else if (panel === 'messages' || panel === 'msg') {
    VIEW.messages = !VIEW.messages;
  } else if (panel === 'all') {
    VIEW.sidebar = true;
    VIEW.incidents = true;
    VIEW.messages = true;
  } else if (panel === 'none') {
    VIEW.sidebar = false;
    VIEW.incidents = false;
    VIEW.messages = false;
  }
  saveViewState();
  applyViewState();
}

function toggleActiveBar() {
  VIEW.showActiveBar = !VIEW.showActiveBar;
  saveViewState();
  renderActiveCallsBar();
}

function toggleNightMode() {
  const cycle = { 'dark': 'night', 'night': 'light', 'light': 'dark' };
  VIEW.theme = cycle[VIEW.theme || 'dark'] || 'dark';
  VIEW.nightMode = (VIEW.theme === 'night'); // backward compat
  saveViewState();
  applyViewState();
}

function applyFeatureFlags() {
  const ff = (STATE && STATE.featureFlags) || {};
  // DC911 — show/hide toggle button and sync badge
  const dc911Btn   = document.getElementById('btnToggleDC911');
  const dc911Badge = document.getElementById('dc911SyncBadge');
  const dc911On    = !!ff.dc911_enabled;
  if (dc911Btn)   dc911Btn.style.display   = dc911On ? '' : 'none';
  if (dc911Badge) dc911Badge.style.display = dc911On ? '' : 'none';
  // MAP — show/hide board map button
  const mapBtn = document.getElementById('tbBtnMAP');
  if (mapBtn) mapBtn.style.display = (ff.map_enabled !== false) ? '' : 'none';
  // Incident export — show/hide EXPORT button in incident panel
  const exportBtn = document.getElementById('btnExportInc');
  if (exportBtn) exportBtn.style.display = (ff.incident_export !== false) ? '' : 'none';
}

function cycleDensity() {
  const modes = ['normal', 'compact', 'expanded'];
  const idx = modes.indexOf(VIEW.density);
  VIEW.density = modes[(idx + 1) % modes.length];
  saveViewState();
  applyViewState();
}

function applyPreset(name) {
  if (name === 'dispatch') {
    VIEW.sidebar = false;
    VIEW.incidents = true;
    VIEW.messages = true;
    VIEW.density = 'normal';
    VIEW.sort = 'status';
    VIEW.sortDir = 'asc';
    VIEW.filterStatus = null;
  } else if (name === 'supervisor') {
    VIEW.sidebar = true;
    VIEW.incidents = true;
    VIEW.messages = true;
    VIEW.density = 'normal';
    VIEW.sort = 'status';
    VIEW.sortDir = 'asc';
    VIEW.filterStatus = null;
  } else if (name === 'field') {
    VIEW.sidebar = false;
    VIEW.incidents = false;
    VIEW.messages = false;
    VIEW.density = 'compact';
    VIEW.sort = 'status';
    VIEW.sortDir = 'asc';
    VIEW.filterStatus = null;
  }
  VIEW.preset = name;
  saveViewState();
  applyViewState();
  renderBoardDiff();
}

function toggleIncidentQueue() {
  VIEW.incidents = !VIEW.incidents;
  saveViewState();
  applyViewState();
}

// Toolbar event handlers
function tbFilterChanged() {
  const val = document.getElementById('tbFilterStatus').value;
  VIEW.filterStatus = val || null;
  saveViewState();
  renderBoardDiff();
}

function tbSortChanged() {
  VIEW.sort = document.getElementById('tbSort').value || 'status';
  saveViewState();
  updateSortHeaders();
  renderBoardDiff();
}

// ============================================================
// Audio Feedback (board/dispatch side)
// ============================================================
function beepChange()     { try { const a = new Audio('sounds/pg.mp3'); a.volume = 0.15; a.play().catch(() => {}); } catch(e) {} }
function _boardPlayFile(file) {
  try {
    const a = new Audio(file);
    a.volume = 0.9;
    a.play().catch(() => {});
  } catch(e) {}
}
function beepNote()       { _boardPlayFile('sounds/msg.mp3'); }
// Alert banner set — plays htmsg tone
function beepAlert()      { _boardPlayFile('sounds/htmsg.mp3'); }

// Incoming regular message
function beepMessage()    { _boardPlayFile('sounds/msg.mp3'); }
// Incoming urgent/hot message or alert
function beepHotMessage() { _boardPlayFile('sounds/htmsg.mp3'); }

// ============================================================
// Utility Functions
// ============================================================

// Parse unit_info crew string (CM1:NAME (CERT)|CM2:NAME (CERT)) into readable text
function parseCrewInfo(unitInfo) {
  if (!unitInfo) return '';
  const parts = unitInfo.split('|').map(p => p.trim()).filter(Boolean);
  const crew = parts
    .filter(p => p.startsWith('CM1:') || p.startsWith('CM2:'))
    .map(p => p.replace(/^CM[12]:/, '').trim());
  return crew.join(' / ');
}

function esc(s) {
  return String(s ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", "&#039;");
}

// Normalize a timestamp string before passing to new Date().
// Supabase returns TIMESTAMPTZ with microsecond precision, e.g. "2026-02-22T15:30:00.123456+00:00".
// ECMAScript Date.parse only guarantees parsing up to 3 fractional-second digits (milliseconds).
// Older Safari (iOS 15 / macOS 12 and earlier) returns Invalid Date for 6-digit fractional seconds.
// Truncating to 3 decimal places makes the string spec-compliant and cross-browser safe.
function _normalizeTs(i) {
  if (typeof i !== 'string') return i;
  return i.replace(/(\.\d{3})\d+/, '$1');
}

function fmtTime24(i) {
  if (!i) return '—';
  const d = new Date(_normalizeTs(i));
  if (!isFinite(d.getTime())) return '—';
  return d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
}

function minutesSince(i) {
  if (!i) return null;
  const t = new Date(_normalizeTs(i)).getTime();
  if (!isFinite(t)) return null;
  return (Date.now() - t) / 60000;
}

function formatElapsed(minutes) {
  if (minutes == null) return '—';
  if (VIEW.elapsedFormat === 'off') return '';
  const m = Math.floor(minutes);
  if (VIEW.elapsedFormat === 'long') {
    const hrs = Math.floor(m / 60);
    const mins = m % 60;
    const secs = Math.floor((minutes - m) * 60);
    if (hrs > 0) return hrs + ':' + String(mins).padStart(2, '0') + ':' + String(secs).padStart(2, '0');
    return mins + ':' + String(secs).padStart(2, '0');
  }
  // short format
  if (m >= 60) {
    const hrs = Math.floor(m / 60);
    const mins = m % 60;
    return hrs + 'H' + (mins > 0 ? String(mins).padStart(2, '0') + 'M' : '');
  }
  return m + 'M';
}

function statusRank(c) {
  return STATUS_RANK[String(c || '').toUpperCase()] ?? 99;
}

function displayNameForUnit(u) {
  const uu = String(u || '').trim().toUpperCase();
  return UNIT_LABELS[uu] || uu;
}

function canonicalUnit(r) {
  if (!r) return '';
  let u = String(r).trim().toUpperCase().replace(/[^\w\s-]/g, '').replace(/\s+/g, ' ').trim();
  const k = Object.keys(UNIT_LABELS).sort((a, b) => b.length - a.length);
  for (const kk of k) {
    if (u === kk) return kk;
  }
  return u;
}

function expandShortcutsInText(t) {
  if (!t) return '';
  return t.toUpperCase().split(/\b/).map(w => UNIT_LABELS[w.toUpperCase()] || w).join('');
}

// Levenshtein distance for fuzzy unit matching
function _levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) =>
    Array.from({ length: n + 1 }, (_, j) => i === 0 ? j : j === 0 ? i : 0)
  );
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] :
        1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

// Returns known unit IDs that are within edit distance 2 of the typed ID
function findSimilarUnits(typedId) {
  if (!STATE || !STATE.units) return [];
  const t = typedId.toUpperCase();
  return (STATE.units || [])
    .map(u => (u.unit_id || '').toUpperCase())
    .filter(uid => uid && uid !== t && _levenshtein(uid, t) <= 2)
    .slice(0, 5);
}

function getRoleColor(a) {
  const m = String(a || '').match(/@([A-Z0-9]+)$/);
  if (!m) return '';
  return 'roleColor-' + m[1];
}

function setLive(ok, txt) {
  const e = document.getElementById('livePill');
  e.className = 'pill ' + (ok ? 'live' : 'offline');
  e.textContent = txt || (ok ? 'LIVE' : 'OFFLINE');
}

function offline(e) {
  console.error(e);
  setLive(false, 'OFFLINE');
}

function autoFocusCmd() {
  setTimeout(() => document.getElementById('cmd').focus(), 100);
}

// ============================================================
// Dialog Functions
// ============================================================
function showConfirm(title, message, callback, cancelCallback, cancelLabel) {
  document.getElementById('confirmTitle').textContent = title;
  document.getElementById('confirmMessage').textContent = message;
  CONFIRM_CALLBACK = callback;
  CONFIRM_CANCEL_CALLBACK = cancelCallback || null;
  const closeBtn = document.getElementById('confirmClose');
  if (closeBtn) {
    if (cancelCallback) {
      closeBtn.textContent = cancelLabel || 'CANCEL';
      closeBtn.style.display = '';
    } else {
      closeBtn.style.display = 'none';
    }
  }
  document.getElementById('confirmDialog').classList.add('active');
}

function hideConfirm() {
  document.getElementById('confirmDialog').classList.remove('active');
  const closeBtn = document.getElementById('confirmClose');
  if (closeBtn) { closeBtn.style.display = 'none'; closeBtn.textContent = 'CLOSE'; }
  CONFIRM_CALLBACK = null;
  CONFIRM_CANCEL_CALLBACK = null;
}

function showConfirmAsync(title, msg) {
  return new Promise((resolve) => {
    showConfirm(title, msg, () => resolve(true), () => resolve(false), 'BACK');
  });
}

// ── Disposition Short Codes ────────────────────────────────────────
const DISPO_SHORT_CODES = {
  'TC':  'TRANSPORTED',
  'PR':  'PATIENT-REFUSED',
  'CAN': 'CANCELLED-ON-SCENE',
  'OCS': 'CANCELLED-ON-SCENE',
  'REF': 'PATIENT-REFUSED',
  'NP':  'NO-PATIENT-FOUND',
  'MA':  'MUTUAL-AID-TRANSFER',
  'DUP': 'DUPLICATE',
  'ERR': 'DATA-ERROR',
  'OTHER': 'OTHER',
};

// ── Disposition Picker ────────────────────────────────────────
let _dispResolve = null;
const DISPOSITION_CODES = [
  { code: 'TRANSPORTED',            label: 'TRANSPORTED' },
  { code: 'CANCELLED-PRIOR',        label: 'CANCELLED — PRIOR TO DISPATCH' },
  { code: 'CANCELLED-ON-SCENE',     label: 'CANCELLED — ON SCENE' },
  { code: 'MUTUAL-AID-TRANSFER',    label: 'MUTUAL AID TRANSFER' },
  { code: 'PATIENT-REFUSED',        label: 'PATIENT REFUSED' },
  { code: 'NO-PATIENT-FOUND',       label: 'NO PATIENT FOUND' },
  { code: 'DUPLICATE',              label: 'DUPLICATE CALL' },
  { code: 'DATA-ERROR',             label: 'DATA ENTRY ERROR' },
  { code: 'REFUSED-BY-RECEIVING',   label: 'REFUSED BY RECEIVING FACILITY' },
  { code: 'DIVERTED-EN-ROUTE',      label: 'DIVERTED EN ROUTE' },
  { code: 'CONVERTED-TO-911',       label: 'CONVERTED TO 911 RESPONSE' },
  { code: 'OTHER',                  label: 'OTHER / UNSPECIFIED' },
];

function promptDisposition(incidentId) {
  return new Promise(resolve => {
    _dispResolve = resolve;
    const overlay = document.getElementById('dispositionOverlay');
    const label = document.getElementById('dispositionIncLabel');
    const btns = document.getElementById('dispositionBtns');
    if (label) label.textContent = 'INCIDENT ' + String(incidentId).replace(/^[A-Z]*\d{2}-0*/, '');
    if (btns) {
      btns.innerHTML = DISPOSITION_CODES.map(d =>
        '<button onclick="selectDisposition(\'' + d.code.replace(/'/g, "\\'") + '\')" style="padding:8px 6px;background:var(--bg);border:1px solid var(--border);color:var(--text);font-family:inherit;font-size:10px;font-weight:900;cursor:pointer;text-align:center;letter-spacing:.04em;line-height:1.3;">' + esc(d.label) + '</button>'
      ).join('');
    }
    if (overlay) overlay.style.display = 'flex';
  });
}

function selectDisposition(code) {
  const overlay = document.getElementById('dispositionOverlay');
  if (overlay) overlay.style.display = 'none';
  if (_dispResolve) { _dispResolve(code); _dispResolve = null; }
}

function cancelDisposition() {
  const overlay = document.getElementById('dispositionOverlay');
  if (overlay) overlay.style.display = 'none';
  if (_dispResolve) { _dispResolve(null); _dispResolve = null; }
}

// ── OOS Reason Dialog ─────────────────────────────────────────
let _oosResolve = null;
const OOS_REASONS = ['MECHANICAL','FUEL','CREW REST','DOCUMENTATION','TRAINING','HOSPITAL','OTHER'];

function promptOOSReason(unitId) {
  return new Promise(resolve => {
    _oosResolve = resolve;
    document.getElementById('oosUnitLabel').textContent = unitId;
    const btns = document.getElementById('oosReasonBtns');
    btns.innerHTML = OOS_REASONS.map(r =>
      `<button class="btn-secondary" style="text-align:left;" onclick="selectOOSReason('${r}')">${r}</button>`
    ).join('');
    const dlg = document.getElementById('oosReasonDialog');
    dlg.style.display = 'flex';
  });
}

function selectOOSReason(reason) {
  document.getElementById('oosReasonDialog').style.display = 'none';
  if (_oosResolve) { _oosResolve(reason); _oosResolve = null; }
}

function cancelOOSReason() {
  document.getElementById('oosReasonDialog').style.display = 'none';
  if (_oosResolve) { _oosResolve(null); _oosResolve = null; }
}

// New unit confirmation dialog — [BACK] or [LOG ON NEW UNIT]
function showNewUnitDialog(unitId, msg, note) {
  const dlg = document.getElementById('newUnitDialog');
  if (!dlg) {
    // Fallback for cached old board.html without the dialog element
    return showConfirmAsync('NEW UNIT: ' + unitId, msg).then(ok => ok ? 'logon' : 'back');
  }
  return new Promise(resolve => {
    _newUnitResolve = resolve;
    _newUnitPendingNote = note || '';
    document.getElementById('newUnitDialogId').textContent = unitId;
    document.getElementById('newUnitDialogMsg').textContent = msg;
    dlg.style.display = 'flex';
  });
}

function _newUnitBack() {
  document.getElementById('newUnitDialog').style.display = 'none';
  const r = _newUnitResolve;
  _newUnitResolve = null;
  if (r) r('back');
}

function _newUnitOpen() {
  document.getElementById('newUnitDialog').style.display = 'none';
  const uid = document.getElementById('newUnitDialogId').textContent;
  const note = _newUnitPendingNote;
  const r = _newUnitResolve;
  _newUnitResolve = null;
  _newUnitPendingNote = '';
  // Open modal pre-filled (same as LUI <UNIT>)
  const dN = displayNameForUnit(uid);
  const fakeUnit = { unit_id: uid, display_name: dN, type: '', active: true, status: 'AV', note: note, unit_info: '', incident: '', destination: '', updated_at: '', updated_by: '' };
  openModal(fakeUnit);
  if (r) r('logon');
}

function updateScopeIndicator(scope) {
  const el = document.getElementById('scopeIndicator');
  if (!el) return;
  el.textContent = scope === 'ALL' ? 'SCOPE: ALL AGENCIES' : 'SCOPE: ' + scope.replace('AGENCY ', '');
  el.style.display = scope && scope !== 'AGENCY SCMC' ? '' : 'none';
}

function showToast(msg, type = 'info', duration = 3000) {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const el = document.createElement('div');
  el.className = 'toast toast-' + type;
  el.textContent = msg;
  container.appendChild(el);
  requestAnimationFrame(() => el.classList.add('toast-visible'));
  setTimeout(() => {
    el.classList.remove('toast-visible');
    el.addEventListener('transitionend', () => el.remove(), { once: true });
  }, duration);
}

// ============================================================
// Bug / Issue Report Modal
// ============================================================
// Auto-captured context snapshot — built when modal opens, sent with submission
let _bugContext = null;

function openBugReport() {
  const back = document.getElementById('bugReportBack');
  if (!back) return;
  document.getElementById('bugReporterLabel').textContent = ACTOR || '—';
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  document.getElementById('bugTimeLabel').textContent =
    (now.getMonth() + 1) + '/' + now.getDate() + ' ' + hh + ':' + mm;
  document.getElementById('bugDescription').value = '';
  document.getElementById('bugSeverity').value = 'MED';
  document.getElementById('bugStatus').textContent = '';

  // Capture auto-context snapshot
  _bugContext = {
    actor:           ACTOR || '',
    role:            ROLE || '',
    selectedUnit:    SELECTED_UNIT_ID || null,
    openIncident:    CURRENT_INCIDENT_ID || null,
    lastCmds:        CMD_HISTORY.slice(-8),
    unitsOnline:     STATE ? (STATE.units || []).filter(u => u.active).length : null,
    activeIncidents: STATE ? (STATE.incidents || []).filter(i => i.status === 'ACTIVE').length : null,
    online:          navigator.onLine,
    ua:              navigator.userAgent.substring(0, 120),
    vp:              window.innerWidth + 'x' + window.innerHeight,
    ts:              now.toISOString(),
  };

  // Render context preview in modal
  const ctxEl = document.getElementById('bugContextPreview');
  if (ctxEl) {
    const lines = [
      'UNIT: '    + (SELECTED_UNIT_ID || 'none'),
      'INCIDENT: '+ (CURRENT_INCIDENT_ID || 'none'),
      'LAST CMDS: ' + (_bugContext.lastCmds.slice(-4).join(' → ') || 'none'),
      'BOARD: '   + (_bugContext.unitsOnline ?? '?') + ' units / ' + (_bugContext.activeIncidents ?? '?') + ' active inc',
      'ONLINE: '  + (navigator.onLine ? 'yes' : 'NO'),
    ];
    ctxEl.textContent = lines.join('\n');
  }

  back.style.display = 'flex';
  setTimeout(() => document.getElementById('bugDescription').focus(), 80);
}

function closeBugReport() {
  const back = document.getElementById('bugReportBack');
  if (back) back.style.display = 'none';
  autoFocusCmd();
}

async function submitBugReport() {
  const sev  = document.getElementById('bugSeverity').value;
  const desc = (document.getElementById('bugDescription').value || '').trim().toUpperCase();
  const stat = document.getElementById('bugStatus');
  if (!desc) { stat.textContent = 'DESCRIPTION REQUIRED.'; return; }
  if (!TOKEN) { stat.textContent = 'NOT LOGGED IN.'; return; }
  stat.textContent = 'SUBMITTING...';
  const ctx = _bugContext ? JSON.stringify(_bugContext) : '';
  try {
    const r = await API.submitIssue(TOKEN, 'BOARD', sev, desc, ctx);
    if (!r.ok) { stat.textContent = 'ERROR: ' + (r.error || 'UNKNOWN'); return; }
    _bugContext = null;
    closeBugReport();
    showToast('ISSUE REPORT SUBMITTED. THANK YOU.', 'success');
  } catch (e) {
    stat.textContent = 'NETWORK ERROR.';
  }
}

function showAlert(title, message, style) {
  const titleEl = document.getElementById('alertTitle');
  const msgEl = document.getElementById('alertMessage');
  const dialogEl = document.getElementById('alertDialog');
  if (!titleEl || !msgEl || !dialogEl) {
    alert(title + '\n\n' + message);
    return;
  }
  titleEl.textContent = title;
  msgEl.textContent = message;
  msgEl.style.color = style === 'yellow' ? 'var(--yellow)' : '';
  dialogEl.classList.add('active');
}

function hideAlert() {
  document.getElementById('alertDialog').classList.remove('active');
}

function showErr(r) {
  if (r && r.conflict) {
    showConfirm('CONFLICT', r.error + '\n\nCURRENT: ' + r.current.status + '\nUPDATED: ' + r.current.updated_at + '\nBY: ' + r.current.updated_by, () => refresh());
    return;
  }
  showAlert('ERROR', r && r.error ? r.error : 'UNKNOWN ERROR.');
  refresh();
}

// ============================================================
// Authentication
// ============================================================
async function login() {
  const r = (document.getElementById('loginRole').value || '').trim().toUpperCase();
  const cadId = (document.getElementById('loginCadId').value || '').trim();
  const p = (document.getElementById('loginPassword').value || '').trim();
  const e = document.getElementById('loginErr');
  e.textContent = '';

  if (!r) { e.textContent = 'SELECT ROLE.'; return; }
  if (!cadId || cadId.length < 2) { e.textContent = 'ENTER CAD ID.'; return; }
  if (!p) { e.textContent = 'ENTER PASSWORD.'; return; }

  const res = await API.login(r, cadId, p, 'board');
  if (!res || !res.ok) {
    if (res && res.mustChangePassword) {
      showMustChangePassword(res.username || cadId, p);
      return;
    }
    e.textContent = (res && res.error) ? res.error : 'LOGIN FAILED.';
    return;
  }

  TOKEN = res.token;
  ACTOR = res.actor;
  ROLE = r;
  localStorage.setItem('ems_token', TOKEN);
  localStorage.setItem('ems_role', ROLE);
  localStorage.setItem('ems_actor', ACTOR);
  document.getElementById('loginBack').style.display = 'none';
  document.getElementById('userLabel').textContent = ACTOR;
  const adminLink = document.getElementById('adminLink');
  if (adminLink) adminLink.style.display = ['SUPV1','SUPV2','MGR1','MGR2','IT'].includes(ROLE) ? '' : 'none';
  start();
  setTimeout(openPopouts, 500);
}

async function showMustChangePassword(cadIdOrUsername, oldPassword) {
  const newPw = window.prompt('YOUR PASSWORD MUST BE CHANGED BEFORE LOGGING IN.\n\nENTER NEW PASSWORD (MIN 5 CHARACTERS):');
  if (!newPw || newPw.trim().length < 5) {
    showToast('PASSWORD MUST BE AT LEAST 5 CHARACTERS.', 'error');
    return;
  }
  const confirm = window.prompt('CONFIRM NEW PASSWORD:');
  if (newPw !== confirm) {
    showToast('PASSWORDS DO NOT MATCH.', 'error');
    return;
  }
  const r = await API.changePasswordNoAuth(cadIdOrUsername, oldPassword, newPw);
  if (r && r.ok) {
    showToast('PASSWORD CHANGED. LOGGING IN...', 'success');
    // Re-attempt login with the new password
    const role = (document.getElementById('loginRole').value || '').trim().toUpperCase();
    const cadId = (document.getElementById('loginCadId').value || '').trim();
    const res2 = await API.login(role, cadId, newPw, 'board');
    if (!res2 || !res2.ok) {
      document.getElementById('loginErr').textContent = (res2 && res2.error) ? res2.error : 'LOGIN FAILED AFTER PASSWORD CHANGE.';
      return;
    }
    TOKEN = res2.token;
    ACTOR = res2.actor;
    ROLE = role;
    localStorage.setItem('ems_token', TOKEN);
    localStorage.setItem('ems_role', ROLE);
    localStorage.setItem('ems_actor', ACTOR);
    document.getElementById('loginBack').style.display = 'none';
    document.getElementById('userLabel').textContent = ACTOR;
    const adminLink = document.getElementById('adminLink');
    if (adminLink) adminLink.style.display = ['SUPV1','SUPV2','MGR1','MGR2','IT'].includes(ROLE) ? '' : 'none';
    start();
  } else {
    showToast('ERROR: ' + ((r && r.error) || 'COULD NOT CHANGE PASSWORD'), 'error');
  }
}

// ============================================================
// Data Refresh
// ============================================================
// Performance: Granular change detection instead of JSON.stringify
let _lastUnitsHash = '';
let _lastIncidentsHash = '';
let _lastBannersHash = '';
let _lastMessagesHash = '';
let _refreshing = false;
let _pendingRender = false;
let _changedSections = { units: false, incidents: false, banners: false, messages: false };

// Performance: Cache for row data to enable DOM diffing
let _rowCache = new Map(); // unit_id -> { html, status, updated_at, ... }

// M-6: Banner acknowledgment tracking — persisted in sessionStorage (resets on page reload)
// Key format: "<kind>:<message>" — stale when message changes, so ack reappears for new content
const _ackedBanners = new Set(
  JSON.parse(sessionStorage.getItem('_ackedBanners') || '[]')
);
function _saveBannerAcks() {
  sessionStorage.setItem('_ackedBanners', JSON.stringify([..._ackedBanners]));
}
function _bannerKey(kind, message) {
  return kind + ':' + (message || '');
}
function _ackBanner(kind) {
  const b = (STATE && STATE.banners) ? STATE.banners : {};
  const msg = (b[kind] && b[kind].message) ? b[kind].message : '';
  const key = _bannerKey(kind, msg);
  _ackedBanners.add(key);
  _saveBannerAcks();
  renderBanners();
  // Fire-and-forget audit call — backend receives kind + actor token
  API.call('bannerAck', TOKEN, kind).catch(() => {});
}

// Compute lightweight hash for change detection (no JSON.stringify)
function _computeUnitsHash(units) {
  if (!units || !units.length) return '0';
  let h = units.length + ':';
  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    h += (u.unit_id || '') + (u.status || '') + (u.updated_at || '') + (u.incident || '') + (u.destination || '') + (u.note || '') + (u.active ? '1' : '0') + '|';
  }
  return h;
}

function _computeIncidentsHash(incidents) {
  if (!incidents || !incidents.length) return '0';
  let h = incidents.length + ':';
  for (let i = 0; i < incidents.length; i++) {
    const inc = incidents[i];
    h += (inc.incident_id || '') + (inc.status || '') + (inc.last_update || '') + '|';
  }
  return h;
}

function _computeBannersHash(banners, destinations) {
  const base = !banners ? '0' : (banners.alert?.message || '') + (banners.alert?.ts || '') + (banners.note?.message || '') + (banners.note?.ts || '');
  const divHash = (destinations || []).filter(d => d.diverted).map(d => d.code || '').sort().join(',');
  return base + '|DIV:' + divHash;
}

function _computeMessagesHash(messages) {
  if (!messages || !messages.length) return '0';
  let h = messages.length + ':';
  for (let i = 0; i < messages.length; i++) {
    h += (messages[i].message_id || '') + (messages[i].read ? '1' : '0') + '|';
  }
  return h;
}

async function refresh(forceFull) {
  if (!TOKEN || _refreshing) return;
  _refreshing = true;

  try {
    // PERF-3: Pass sinceTs on background polls to enable delta responses.
    // forceFull=true (used by forceRefresh) always requests a complete state.
    const sinceTs = (!forceFull && BASELINED && LAST_MAX_UPDATED_AT) ? LAST_MAX_UPDATED_AT : null;
    const r = await API.getState(TOKEN, sinceTs);
    if (!r || !r.ok) {
      // Detect session invalidation (kicked by another login or expired)
      if (r && r.error && (r.error.includes('NOT AUTHENTICATED') || r.error.includes('TOKEN EXPIRED'))) {
        showAlert('SESSION ENDED — ANOTHER USER LOGGED INTO THIS POSITION OR SESSION EXPIRED. PLEASE LOG IN AGAIN.');
        localStorage.removeItem('ems_token');
        localStorage.removeItem('ems_role');
        localStorage.removeItem('ems_actor');
        TOKEN = '';
        ACTOR = '';
        if (POLL) clearInterval(POLL);
        stopLfnPolling();
        _rtDisconnect();
        document.getElementById('loginBack').style.display = 'flex';
        document.getElementById('userLabel').textContent = '—';
        return;
      }
      setLive(false, 'OFFLINE');
      return;
    }

    // PERF-3: Merge delta responses into existing STATE rather than replacing entirely.
    // A delta response has isDelta=true and contains only units changed since sinceTs.
    if (r.isDelta && STATE) {
      // Merge changed/new units; keep units not in delta untouched
      if (r.units && r.units.length > 0) {
        const updatedIds = new Set(r.units.map(function(u) { return u.unit_id; }));
        const kept = (STATE.units || []).filter(function(u) { return !updatedIds.has(u.unit_id); });
        STATE.units = kept.concat(r.units);
      }
      // Always replace small payloads returned in full
      if (r.incidents !== undefined) STATE.incidents = r.incidents;
      if (r.banners !== undefined) STATE.banners = r.banners;
      if (r.destinations !== undefined) STATE.destinations = r.destinations;
      if (r.messages !== undefined) STATE.messages = r.messages;
      if (r.assignments !== undefined) STATE.assignments = r.assignments;
      if (r.dc911Config !== undefined) STATE.dc911Config = r.dc911Config;
      if (r.dc911State !== undefined) STATE.dc911State = r.dc911State;
      if (r.roster !== undefined) STATE.roster = r.roster;
      if (r.typeCodes !== undefined) { STATE.typeCodes = r.typeCodes; populateTypeCodeDatalist(); }
      if (r.featureFlags !== undefined) STATE.featureFlags = r.featureFlags;
      STATE.serverTime = r.serverTime;
      STATE.actor = r.actor || STATE.actor;
    } else {
      // Full state replace (cache hit, sinceTs=null, or forceFull)
      STATE = r;
    }
    if (!STATE.staleThresholds) STATE.staleThresholds = { WARN: 10, ALERT: 20, CRITICAL: 30 };

    if (r.incTypeTaxonomy && typeof r.incTypeTaxonomy === 'object' && Object.keys(r.incTypeTaxonomy).length > 0) {
      INC_TYPE_TAXONOMY = r.incTypeTaxonomy;
    }
    _lastPollAt = Date.now();
    setLive(true, 'LIVE • ' + fmtTime24(STATE.serverTime));
    applyFeatureFlags();
    ACTOR = STATE.actor || ACTOR;
    document.getElementById('userLabel').textContent = ACTOR;
    tryBeepOnStateChange();

    // Granular change detection — only re-render what actually changed
    const unitsHash = _computeUnitsHash(STATE.units);
    const incidentsHash = _computeIncidentsHash(STATE.incidents);
    const bannersHash = _computeBannersHash(STATE.banners, STATE.destinations);
    const messagesHash = _computeMessagesHash(STATE.messages);

    _changedSections.units = (unitsHash !== _lastUnitsHash);
    _changedSections.incidents = (incidentsHash !== _lastIncidentsHash);
    _changedSections.banners = (bannersHash !== _lastBannersHash);
    _changedSections.messages = (messagesHash !== _lastMessagesHash);

    _lastUnitsHash = unitsHash;
    _lastIncidentsHash = incidentsHash;
    _lastBannersHash = bannersHash;
    _lastMessagesHash = messagesHash;

    const anyChange = _changedSections.units || _changedSections.incidents || _changedSections.banners || _changedSections.messages;

    if (anyChange) {
      if (document.hidden) {
        _pendingRender = true;
      } else {
        renderSelective();
      }
    }
  } finally {
    _refreshing = false;
  }
}

async function forceRefresh() {
  _lastUnitsHash = null;
  _lastIncidentsHash = null;
  _lastBannersHash = null;
  _lastMessagesHash = null;
  await refresh(true); // forceFull=true: bypass delta, get complete state
  showToast('REFRESHED.', 'success');
}

function toggleAssisting() {
  _showAssisting = !_showAssisting;
  const btn = document.getElementById('btnToggleAssisting');
  if (btn) btn.textContent = _showAssisting ? 'AUX ON' : 'AUX OFF';
  if (btn) btn.style.opacity = _showAssisting ? '1' : '0.45';
  renderBoardDiff(STATE);
}

function runQuickCmd(cmd) {
  const inp = document.getElementById('cmd');
  if (inp) inp.value = cmd;
  runCommand();
}

// Performance: Selective rendering — only update changed sections
function renderSelective() {
  if (!STATE) return;

  // Populate status dropdown once
  const sS = document.getElementById('mStatus');
  if (!sS.options.length) {
    (STATE.statuses || []).forEach(s => {
      const o = document.createElement('option');
      o.value = s.code;
      o.textContent = s.code + ' — ' + s.label;
      sS.appendChild(o);
    });
  }

  // Only render what changed
  if (_changedSections.banners) renderBanners();
  if (_changedSections.units) renderStatusSummary();
  // Board re-renders on unit OR incident changes (board rows display incident notes)
  if (_changedSections.units || _changedSections.incidents) renderBoardDiff();
  if (_changedSections.incidents) renderIncidentQueue();
  if (_changedSections.units || _changedSections.incidents) renderActiveCallsBar();
  if (_changedSections.units || _changedSections.incidents) renderBoardMap();
  if (_changedSections.messages) {
    renderMessagesPanel();
    renderMessages();
    renderInboxPanel();
  }

}

function tryBeepOnStateChange() {
  let mU = '';
  (STATE.units || []).forEach(u => {
    if (u && u.updated_at && (!mU || u.updated_at > mU)) mU = u.updated_at;
  });

  const nTs = (STATE.banners && STATE.banners.note && STATE.banners.note.ts) ? STATE.banners.note.ts : '';
  const aTs = (STATE.banners && STATE.banners.alert && STATE.banners.alert.ts) ? STATE.banners.alert.ts : '';

  let mI = '';
  (STATE.incidents || []).forEach(i => {
    if (i && i.last_update && (!mI || i.last_update > mI)) mI = i.last_update;
  });

  const mC = (STATE.messages || []).length;
  const uU = (STATE.messages || []).filter(m => m.urgent && !m.read).length;

  if (!BASELINED) {
    BASELINED = true;
    LAST_MAX_UPDATED_AT = mU;
    LAST_NOTE_TS = nTs;
    LAST_ALERT_TS = aTs;
    LAST_INCIDENT_TOUCH = mI;
    LAST_MSG_COUNT = mC;
    return;
  }

  if (aTs && aTs !== LAST_ALERT_TS) {
    LAST_ALERT_TS = aTs;
    beepAlert();
    // Browser notification for alert banner
    if ('Notification' in window && Notification.permission === 'granted' && document.hidden) {
      try {
        const alertText = (STATE.banners && STATE.banners.alert && STATE.banners.alert.message) || 'ALERT';
        const n = new Notification('HOSCAD ALERT', { body: alertText, tag: 'hoscad-alert', icon: 'download.png' });
        n.onclick = function() { window.focus(); n.close(); };
        setTimeout(function() { n.close(); }, 10000);
      } catch (e) {}
    }
  }
  if (nTs && nTs !== LAST_NOTE_TS) { LAST_NOTE_TS = nTs; beepNote(); }
  if (mC > LAST_MSG_COUNT) {
    LAST_MSG_COUNT = mC;
    if (uU > 0) beepHotMessage(); else beepMessage();
  }
  if (mI && mI > LAST_INCIDENT_TOUCH) { LAST_INCIDENT_TOUCH = mI; } // no beep — unit/incident status changes are silent
  if (mU && mU > LAST_MAX_UPDATED_AT) { LAST_MAX_UPDATED_AT = mU; } // audio only for messages, alerts, banners

  // HOLD call maturity — beep once when a scheduled call's time arrives
  if (BASELINED) {
    (STATE.incidents || []).forEach(inc => {
      if (!inc.incident_note || _holdAlertedIds.has(inc.incident_id)) return;
      const holdM = (inc.incident_note || '').match(/\[HOLD:(\d{2}:\d{2})\]/i);
      if (!holdM) return;
      const parts = holdM[1].split(':').map(Number);
      const now = new Date();
      const holdDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parts[0], parts[1], 0);
      if (now >= holdDate) {
        _holdAlertedIds.add(inc.incident_id);
        beepAlert();
        showToast('HOLD CALL READY: ' + inc.incident_id + ' → ' + (inc.destination || '?') + ' (SCHED ' + holdM[1] + ')');
      }
    });
  }

  // PRI-1 / urgent new incident alert — beep once on creation (≤3 min old)
  if (BASELINED) {
    (STATE.incidents || []).forEach(inc => {
      if (_urgentIncAlertedIds.has(inc.incident_id)) return;
      if (inc.status !== 'QUEUED' && inc.status !== 'ACTIVE') return;
      const pri = (inc.priority || '').toUpperCase();
      const isUrgent = pri === 'PRI-1' || pri === 'CRITICAL' || (inc.incident_note || '').includes('[URGENT]');
      if (!isUrgent) return;
      const ageMin = minutesSince(inc.created_at);
      if (ageMin == null || ageMin > 3) return; // ignore incidents older than 3 min (pre-existing)
      _urgentIncAlertedIds.add(inc.incident_id);
      beepAlert();
      const shortId = String(inc.incident_id).replace(/^[A-Z]*\d{2}-0*/, '');
      showToast('PRI-1: ' + inc.incident_id + ' — ' + (inc.incident_type || 'INCIDENT').toUpperCase() + (inc.scene_address ? ' @ ' + inc.scene_address.toUpperCase() : ''), 'warn', 8000);
    });
  }

  // Unattended incident alert — beep once when a QUEUED incident has no assigned unit for >30 min
  if (BASELINED) {
    (STATE.incidents || []).forEach(inc => {
      if (inc.status !== 'QUEUED') return;
      if (_unattendedAlertedIds.has(inc.incident_id)) return;
      const hasPrimary = (STATE.assignments || []).some(a => a.incident_id === inc.incident_id && !a.cleared_at);
      if (hasPrimary) return;
      const waitMin = minutesSince(inc.created_at);
      if (waitMin == null || waitMin < 30) return;
      _unattendedAlertedIds.add(inc.incident_id);
      beepAlert();
      const shortId = String(inc.incident_id).replace(/^[A-Z]*\d{2}-0*/, '');
      showToast('UNATTENDED: ' + inc.incident_id + ' QUEUED ' + Math.floor(waitMin) + 'M — NO UNIT ASSIGNED', 'warn', 8000);
    });
  }

  // New unit online — beep + toast when a unit becomes active
  const curActiveSet = new Set((STATE.units || []).filter(u => u.active).map(u => u.unit_id));
  if (_prevActiveUnitSet !== null && BASELINED) {
    const newlyActive = [];
    curActiveSet.forEach(uid => { if (!_prevActiveUnitSet.has(uid)) newlyActive.push(uid); });
    if (newlyActive.length === 1) {
      beepNote();
      showToast('UNIT ONLINE: ' + newlyActive[0]);
    } else if (newlyActive.length > 1) {
      beepNote();
      showToast('UNITS ONLINE: ' + newlyActive.join(', '));
    }
  }
  _prevActiveUnitSet = curActiveSet;
}

// ============================================================
// Rendering Functions
// ============================================================
function renderAll() {
  if (!STATE) return;

  // Populate status dropdown
  const sS = document.getElementById('mStatus');
  if (!sS.options.length) {
    (STATE.statuses || []).forEach(s => {
      const o = document.createElement('option');
      o.value = s.code;
      o.textContent = s.code + ' — ' + s.label;
      sS.appendChild(o);
    });
  }

  renderBanners();
  renderStatusSummary();
  renderActiveCallsBar();
  renderIncidentQueue();
  renderMessagesPanel();
  renderMessages();
  renderInboxPanel();
  renderBoardDiff(); // Use optimized DOM diffing
  applyViewState();
}

function renderBanners() {
  const a = document.getElementById('alertBanner');
  const n = document.getElementById('noteBanner');
  const b = (STATE && STATE.banners) ? STATE.banners : { alert: null, note: null };

  // M-6: Prune acks whose banner content no longer matches current state
  // (keeps the Set small; also means message changes auto-show ACK button again)
  const validKeys = new Set();
  for (const kind of ['alert', 'note']) {
    if (b[kind] && b[kind].message) validKeys.add(_bannerKey(kind, b[kind].message));
  }
  let acksChanged = false;
  for (const k of [..._ackedBanners]) {
    if (!validKeys.has(k)) { _ackedBanners.delete(k); acksChanged = true; }
  }
  if (acksChanged) _saveBannerAcks();

  // M-6: Helper to build banner inner HTML with optional ACK button
  function bannerInner(prefix, kind, bannerObj) {
    const msg = bannerObj.message || '';
    const actor = bannerObj.actor || '';
    const key = _bannerKey(kind, msg);
    const isAcked = _ackedBanners.has(key);
    const ackHtml = isAcked
      ? ' <span class="banner-acked" title="Acknowledged">[ACK\'D]</span>'
      : ' <button class="banner-ack-btn" onclick="_ackBanner(\'' + kind + '\')" title="Acknowledge this banner">[ACK]</button>';
    return prefix + esc(msg) + ' \u2014 ' + esc(actor) + ackHtml;
  }

  if (b.alert && b.alert.message) {
    a.style.display = 'block';
    a.innerHTML = bannerInner('ALERT: ', 'alert', b.alert);
  } else {
    a.style.display = 'none';
  }

  if (b.note && b.note.message) {
    n.style.display = 'block';
    n.innerHTML = bannerInner('NOTE: ', 'note', b.note);
  } else {
    n.style.display = 'none';
  }

  renderDiversionBar();
}

function renderDiversionBar() {
  const bar = document.getElementById('diversionBar');
  const list = document.getElementById('diversionList');
  if (!bar || !list) return;
  const diverted = (STATE && STATE.destinations || []).filter(d => d.diverted);
  if (diverted.length) {
    list.textContent = diverted.map(d => d.name || d.code).join(' / ');
    bar.style.display = 'block';
  } else {
    bar.style.display = 'none';
    list.textContent = '';
  }
}

function renderStatusSummary() {
  const el = document.getElementById('statusSummary');
  if (!el) return;

  const units = (STATE.units || []).filter(u => u.active);
  const counts = { AV: 0, D: 0, DE: 0, OS: 0, T: 0, TH: 0, F: 0, BRK: 0, OOS: 0 };

  units.forEach(u => {
    const st = String(u.status || '').toUpperCase();
    if (counts[st] !== undefined) counts[st]++;
  });

  // Coverage change alert — beep on downward threshold crossings only
  if (BASELINED && _lastAvCount !== null && counts.AV < _lastAvCount) {
    if (_lastAvCount > 0 && counts.AV === 0) {
      beepAlert();
      showToast('COVERAGE: CRITICAL — 0 UNITS AVAILABLE', 'warn', 8000);
    } else if (_lastAvCount > 1 && counts.AV === 1) {
      beepAlert();
      showToast('COVERAGE: LIMITED — 1 UNIT AVAILABLE', 'warn', 6000);
    } else if (_lastAvCount > 3 && counts.AV <= 3) {
      beepNote();
      showToast('COVERAGE: REDUCED — ' + counts.AV + ' UNITS AVAILABLE', 'warn', 4000);
    }
  }
  _lastAvCount = counts.AV;

  // Coverage badge — only shown when AV count is ≤3 (low coverage)
  let coverageBadge = '';
  if (counts.AV === 0) {
    coverageBadge = '<span class="coverage-badge coverage-critical" title="NO units available">COVERAGE: CRITICAL</span>';
  } else if (counts.AV === 1) {
    coverageBadge = '<span class="coverage-badge coverage-limited" title="Only 1 unit available">COVERAGE: LIMITED</span>';
  } else if (counts.AV <= 3) {
    coverageBadge = '<span class="coverage-badge coverage-reduced" title="' + counts.AV + ' units available">COVERAGE: REDUCED</span>';
  }

  el.innerHTML = `
    <span class="sum-item sum-av" onclick="quickFilter('AV')">AV: <strong>${counts.AV}</strong></span>
    <span class="sum-item sum-d" onclick="quickFilter('D')">D: <strong>${counts.D}</strong></span>
    <span class="sum-item sum-de" onclick="quickFilter('DE')">DE: <strong>${counts.DE}</strong></span>
    <span class="sum-item sum-os" onclick="quickFilter('OS')">OS: <strong>${counts.OS}</strong></span>
    <span class="sum-item sum-t" onclick="quickFilter('T')">T: <strong>${counts.T}</strong></span>
    <span class="sum-item sum-th" onclick="quickFilter('TH')">TH: <strong>${counts.TH}</strong></span>
    <span class="sum-item sum-f" onclick="quickFilter('F')">F: <strong>${counts.F}</strong></span>
    <span class="sum-item sum-brk" onclick="quickFilter('BRK')">BRK: <strong>${counts.BRK}</strong></span>
    <span class="sum-item sum-oos" onclick="quickFilter('OOS')">OOS: <strong>${counts.OOS}</strong></span>
    <span class="sum-item sum-total" onclick="quickFilter('')">TOTAL: <strong>${units.length}</strong></span>
    ${coverageBadge}
  `;
}

function quickFilter(status) {
  VIEW.filterStatus = status || null;
  const tbFs = document.getElementById('tbFilterStatus');
  if (tbFs) tbFs.value = VIEW.filterStatus || '';
  saveViewState();
  renderBoardDiff();
}

// Active Calls Bar — compact cards for all ACTIVE incidents above the board
function renderActiveCallsBar() {
  const bar = document.getElementById('activeCallsBar');
  if (!bar || !STATE) return;

  // Update toggle button label
  const toggleBtn = document.getElementById('activeBarToggleBtn');
  if (toggleBtn) toggleBtn.textContent = VIEW.showActiveBar ? 'ACTIVE ▲' : 'ACTIVE ▼';

  const active = (STATE.incidents || []).filter(i => i.status === 'ACTIVE');
  if (!active.length || !VIEW.showActiveBar) { bar.style.display = 'none'; return; }

  // Sort by priority (PRI-1 first), then by created_at
  const priOrder = { 'PRI-1': 0, 'CRITICAL': 0, 'PRI-2': 1, 'PRI-3': 2, 'PRI-4': 3 };
  active.sort((a, b) => {
    const pa = priOrder[(a.priority || '').toUpperCase()] ?? 9;
    const pb = priOrder[(b.priority || '').toUpperCase()] ?? 9;
    if (pa !== pb) return pa - pb;
    return new Date(_normalizeTs(a.created_at)) - new Date(_normalizeTs(b.created_at));
  });

  const now = Date.now();
  const cards = active.map(inc => {
    const shortId = String(inc.incident_id).replace(/^[A-Z]*\d{2}-0*/, '');
    const elapsedMs = now - new Date(_normalizeTs(inc.dispatch_time || inc.created_at)).getTime();
    const elapsedMin = Math.floor(elapsedMs / 60000);
    const elapsedStr = elapsedMin >= 60
      ? Math.floor(elapsedMin / 60) + 'H ' + (elapsedMin % 60) + 'M'
      : elapsedMin + 'M';
    const assignedUnits = (STATE.units || []).filter(u => u.active && u.incident === inc.incident_id);
    const unitStr = assignedUnits.length
      ? assignedUnits.map(u => u.unit_id).join(' · ')
      : 'NO UNITS';
    const pri = inc.priority || '';
    const incType = (inc.incident_type || '').toUpperCase();
    const typeCl = getIncidentTypeClass(incType);
    const isUrgent = (inc.incident_note || '').includes('[URGENT]') || pri === 'PRI-1' || pri === 'CRITICAL';
    const cardCl = 'acb-card' + (isUrgent ? ' acb-urgent' : '');
    const priBadgeHtml = pri ? '<span class="acb-pri ' + esc('priority-' + pri) + '">' + esc(pri) + '</span>' : '';
    const locValAcb = (inc.level_of_care || '').toUpperCase();
    const locBadgeAcb = locValAcb ? '<span style="font-size:9px;font-weight:900;margin-left:3px;padding:1px 4px;background:rgba(121,192,255,.15);border:1px solid rgba(121,192,255,.35);color:#79c0ff;border-radius:2px;">' + esc(locValAcb) + '</span>' : '';
    const isStale = elapsedMin > 60;
    const elCl = isStale ? 'acb-elapsed stale' : elapsedMin > 30 ? 'acb-elapsed warn' : 'acb-elapsed';
    // MA badge for active calls bar
    const acbMaMatches = [...(inc.incident_note || '').matchAll(/\[MA:([^\]:]+):([^\]]+)\]/gi)];
    const acbMaTags = acbMaMatches.map(m => ({ agency: m[1].trim(), status: m[2].trim().toUpperCase() }));
    const acbMaBadge = acbMaTags.length > 0 ? acbMaTags.map(t =>
      t.status === 'ACTIVE'
        ? '<span class="ma-active-badge" style="font-size:8px;">MA:' + esc(t.agency) + '</span>'
        : '<span class="ma-badge" style="font-size:8px;">MA REQ</span>'
    ).join('') : ((inc.incident_note || '').includes('[MA]') ? '<span class="ma-badge" style="font-size:8px;">MA</span>' : '');
    // AIR badge — air resource linked to incident via LFN LINK command
    const acbAirMatches = [...(inc.incident_note || '').matchAll(/\[AIR:([^\]:]+)/gi)];
    const acbAirBadge = acbAirMatches.length > 0
      ? acbAirMatches.map(m => '<span class="air-badge" style="font-size:8px;">AIR:' + esc(m[1].trim()) + '</span>').join('')
      : '';

    return '<div class="' + cardCl + '" data-inc-id="' + esc(inc.incident_id) + '" onclick="openIncident(\'' + esc(inc.incident_id) + '\')" title="' + esc(inc.scene_address || '') + '">' +
      '<div class="acb-id">' + esc(inc.incident_id) + priBadgeHtml + locBadgeAcb + acbMaBadge + acbAirBadge + '</div>' +
      '<div class="acb-type ' + typeCl + '">' + esc(incType || '—') + '</div>' +
      '<div class="' + elCl + '">' + esc(elapsedStr) + ' · ' + assignedUnits.length + ' UNIT' + (assignedUnits.length !== 1 ? 'S' : '') + '</div>' +
      '<div class="acb-units">' + esc(unitStr) + '</div>' +
    '</div>';
  }).join('');

  bar.innerHTML = cards;
  bar.style.display = 'flex';
}

function renderMessages() {
  const m = STATE.messages || [];
  const u = m.filter(mm => !mm.read).length;
  const uu = m.filter(mm => mm.urgent && !mm.read).length;
  const b = document.getElementById('msgBadge');
  const c = document.getElementById('msgCount');

  if (u > 0) {
    b.style.display = 'inline-block';
    c.textContent = u;
    if (uu > 0) {
      b.classList.add('hasUrgent');
    } else {
      b.classList.remove('hasUrgent');
    }
  } else {
    b.style.display = 'none';
  }
}

function getIncidentTypeClass(type) {
  const t = String(type || '').toUpperCase().trim();
  // Priority-based matching (new transport taxonomy)
  if (t.endsWith('-PRI-1') || t === 'PRI-1') return 'inc-type-delta';
  if (t.endsWith('-PRI-2') || t === 'PRI-2') return 'inc-type-charlie';
  if (t.endsWith('-PRI-3') || t === 'PRI-3') return 'inc-type-bravo';
  if (t.endsWith('-PRI-4') || t === 'PRI-4') return 'inc-type-alpha';
  // Category-based fallback for partially-formed types
  if (t.startsWith('CCT')) return 'inc-type-delta';
  if (t.startsWith('IFT-ALS')) return 'inc-type-charlie';
  if (t.startsWith('IFT-BLS')) return 'inc-type-bravo';
  if (t.startsWith('DISCHARGE') || t.startsWith('DIALYSIS')) return 'inc-type-alpha';
  // Priority suffix detection: PRI-1 / PRI-2 / PRI-3 / PRI-4 (regex fallback)
  const priMatch = t.match(/PRI-?(\d)$/);
  if (priMatch) {
    const n = priMatch[1];
    if (n === '1') return 'inc-type-delta';
    if (n === '2') return 'inc-type-charlie';
    if (n === '3') return 'inc-type-bravo';
    if (n === '4') return 'inc-type-alpha';
  }
  // Old-style EMD determinants (backward compat)
  const det = t.split('-').pop();
  if (det === 'DELTA')   return 'inc-type-delta';
  if (det === 'CHARLIE') return 'inc-type-charlie';
  if (det === 'BRAVO')   return 'inc-type-bravo';
  if (det === 'ALPHA')   return 'inc-type-alpha';
  // Category-based fallback (legacy)
  if (t.startsWith('IFT'))         return 'inc-type-bravo';
  if (t.includes('STRETCHER') || t.includes('WHEELCHAIR')) return 'inc-type-discharge';
  if (t) return 'inc-type-other';
  return '';
}

// ── Phase 2D: Stack badge + stack state helpers ──────────────────────────

/**
 * Render a stack depth badge for a unit row.
 * @param {number} stackDepth - total assignments in stack (including primary)
 * @param {boolean} hasUrgent - true if any stacked assignment is PRI-1/urgent
 * @param {string} unitId - unit_id (used to check _expandedStacks)
 * @returns {string} HTML string for the badge, or '' if depth < 2
 */
function renderStackBadge(stackDepth, hasUrgent, unitId) {
  if (!stackDepth || stackDepth < 2) return '';
  const cls = hasUrgent ? 'stack-badge stack-badge-urgent' : 'stack-badge';
  const chevron = _expandedStacks.has(unitId) ? '▲' : '▼';
  return '<span class="' + cls + '" data-stack-unit="' + esc(unitId) + '">' + stackDepth + 'Q ' + chevron + '</span>';
}

/**
 * Extract stack info for a unit from STATE.
 * Reads STATE.assignments (array of {unit_id, incident_id, role, assigned_at}).
 * Returns { depth, hasUrgent } or null if no stack data available.
 * @param {string} unitId
 * @returns {{ depth: number, hasUrgent: boolean }|null}
 */
function getUnitStackData(unitId) {
  if (!STATE || !STATE.assignments || !Array.isArray(STATE.assignments)) return null;
  const unitAssignments = STATE.assignments.filter(a => a.unit_id === unitId);
  if (!unitAssignments.length) return null;
  const depth = unitAssignments.length;
  const hasUrgent = unitAssignments.some(a => {
    const inc = STATE.incidents ? STATE.incidents.find(i => i.incident_id === a.incident_id) : null;
    return inc && (inc.priority === 'PRI-1' || inc.priority === 'CRITICAL' || (inc.incident_note && inc.incident_note.includes('[URGENT]')));
  });
  return { depth, hasUrgent };
}

/**
 * Append stack sub-rows to a DocumentFragment after a unit row.
 * Called in renderBoardDiff for both cached and rebuilt rows.
 */
function _appendStackSubRows(unitId, fragment) {
  if (!_expandedStacks.has(unitId)) return;
  if (!STATE || !STATE.assignments) return;
  const assignments = STATE.assignments
    .filter(a => a.unit_id === unitId)
    .sort((a, b) => (a.assignment_order || 0) - (b.assignment_order || 0));
  // Self-heal: if unit no longer has a stack, clear the stale expanded entry
  if (!assignments.length) { _expandedStacks.delete(unitId); return; }
  assignments.forEach(a => {
    const isPrimary = a.role === 'primary';
    const roleLabel = isPrimary ? 'PRIMARY' : 'QUEUED #' + (a.assignment_order || '?');
    const inc = (STATE.incidents || []).find(i => i.incident_id === a.incident_id);
    const typeLabel = inc ? (inc.incident_type || '') : '';
    const shortId = String(a.incident_id).replace(/^[A-Z]*\d{2}-0*/, '');
    const subTr = document.createElement('tr');
    subTr.className = 'stack-detail-row';
    subTr.innerHTML =
      '<td colspan="2" class="stack-detail-role">' + esc(roleLabel) + '</td>' +
      '<td colspan="3" class="stack-detail-inc">' +
        '<span class="clickableIncidentNum" style="cursor:pointer;" data-inc="' + esc(a.incident_id) + '">' + esc(a.incident_id) + '</span> ' + esc(typeLabel) +
      '</td>' +
      '<td colspan="2" class="stack-detail-actions">' +
        (isPrimary ? '' : '<button class="stack-row-btn" onclick="event.stopPropagation();_execCmd(\'PRIMARY ' + esc(a.incident_id) + ' ' + esc(unitId) + '\')">&#9650; PRIMARY</button>') +
        '<button class="stack-row-btn" style="border-color:#f8514960;color:#f85149;" onclick="event.stopPropagation();_execCmd(\'CLEAR ' + esc(a.incident_id) + ' ' + esc(unitId) + '\')">&#x2715; REMOVE</button>' +
      '</td>';
    fragment.appendChild(subTr);
  });
}

/** Resolve AGENCY_ID from M### or C### unit ID pattern. Returns null if no match. */
function resolveAgencyFromUnitId(uid) {
  const m = String(uid || '').toUpperCase().match(/^[MC](\d+)$/);
  if (!m) return null;
  const n = parseInt(m[1], 10);
  if (n >= 100  && n <= 199)  return 'LAPINE_FD';
  if (n >= 200  && n <= 299)  return 'SUNRIVER_FD';
  if (n >= 300  && n <= 399)  return 'BEND_FIRE';
  if (n >= 400  && n <= 499)  return 'REDMOND_FIRE';
  if (n >= 500  && n <= 599)  return 'CROOK_COUNTY_FIRE';
  if (n >= 600  && n <= 699)  return 'CLOVERDALE_FD';
  if (n >= 700  && n <= 799)  return 'SISTERS_CAMP_SHERMAN';
  if (n >= 800  && n <= 899)  return 'BLACK_BUTTE_RANCH';
  if (n >= 900  && n <= 999)  return 'ALFALFA_FD';
  if (n >= 1100 && n <= 1199) return 'CRESCENT_RFPD';
  if (n >= 1200 && n <= 1299) return 'PRINEVILLE_FIRE';
  if (n >= 1300 && n <= 1399) return 'THREE_RIVERS_FD';
  if (n >= 1700 && n <= 1799) return 'JEFFCO_FIRE_EMS';
  if (n >= 2200 && n <= 2299) return 'WARM_SPRINGS_FD';
  return null;
}

function computeRecommendations() {
  const incType = (document.getElementById('newIncType')?.value || '').trim().toUpperCase();
  const pri = (document.getElementById('newIncPriority')?.value || '').trim().toUpperCase();
  const available = ((STATE && STATE.units) || []).filter(u =>
    u.active && (u.status === 'AV' || u.status === 'BRK') && u.include_in_recommendations !== false
  );
  if (!available.length) return [];

  const needsALS  = /^CCT|^IFT-ALS/.test(incType) || pri === 'PRI-1';
  const preferALS = pri === 'PRI-2';
  const blsOk     = /^IFT-BLS|^DISCHARGE|^DIALYSIS/.test(incType) || pri === 'PRI-3' || pri === 'PRI-4';

  const scored = available.map(u => {
    const level = (u.level || '').toUpperCase();
    let score = 100;
    // BRK penalty
    if (u.status === 'BRK') score -= 10;
    if (needsALS) {
      if (level === 'ALS')        score += 60;
      else if (level === 'AEMT')  score += 30;
      else if (level === 'BLS' || level === 'EMT')  score += 5;
    } else if (preferALS) {
      if (level === 'ALS')        score += 40;
      else if (level === 'AEMT')  score += 25;
      else if (level === 'BLS' || level === 'EMT')  score += 15;
    } else if (blsOk) {
      if (level === 'BLS' || level === 'EMT') score += 40;
      else if (level === 'AEMT')  score += 35;
      else if (level === 'ALS')   score += 20;
    } else {
      if (level === 'ALS')        score += 30;
      else if (level === 'AEMT')  score += 20;
      else if (level === 'BLS' || level === 'EMT')  score += 10;
    }
    return { unit: u, score };
  });

  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, 3).map(s => s.unit);
}

function renderIncSuggest() {
  const el = document.getElementById('incSuggest');
  if (!el) return;
  const recs = computeRecommendations();
  if (!recs.length) { el.innerHTML = ''; return; }

  const chips = recs.map(u => {
    const level = u.level ? '<span class="suggest-level">' + esc(u.level) + '</span>' : '';
    const sta   = u.station ? '<span class="muted" style="font-size:10px;margin-left:2px;">' + esc(u.station) + '</span>' : '';
    return '<button type="button" class="suggest-chip" onclick="selectSuggestedUnit(\'' + esc(u.unit_id) + '\')">' +
      esc(u.unit_id) + level + sta + '</button>';
  }).join('');

  el.innerHTML = '<div class="inc-suggest-row">' +
    '<span class="muted" style="font-size:11px;white-space:nowrap;">SUGGESTED:</span>' +
    chips + '</div>';
}

function selectSuggestedUnit(unitId) {
  const sel = document.getElementById('newIncUnit');
  if (!sel) return;
  sel.value = unitId;
  sel.classList.add('row-flash');
  setTimeout(() => sel.classList.remove('row-flash'), 600);
}

function renderIncidentQueue() {
  const panel = document.getElementById('incidentQueue');
  const countEl = document.getElementById('incQueueCount');
  const incidents = (STATE.incidents || []).filter(i => i.status === 'QUEUED');

  if (countEl) countEl.textContent = incidents.length > 0 ? '(' + incidents.length + ' QUEUED)' : '';

  if (!incidents.length) {
    panel.innerHTML = '<div class="muted" style="padding:8px;text-align:center;">NO QUEUED INCIDENTS</div>';
    return;
  }

  // Sort by priority (PRI-1 first), then by creation time (oldest first)
  const _qPriOrder = { 'PRI-1': 0, 'CRITICAL': 0, 'PRI-2': 1, 'PRI-3': 2, 'PRI-4': 3 };
  incidents.sort((a, b) => {
    const pa = _qPriOrder[(a.priority || '').toUpperCase()] ?? 9;
    const pb = _qPriOrder[(b.priority || '').toUpperCase()] ?? 9;
    if (pa !== pb) return pa - pb;
    return new Date(_normalizeTs(a.created_at)) - new Date(_normalizeTs(b.created_at));
  });

  let html = '<table class="inc-queue-table"><thead><tr>';
  html += '<th>INC#</th><th>LOCATION</th><th>TYPE</th><th>NOTE</th><th>SCENE</th><th>WAIT</th><th>ACTIONS</th>';
  html += '</tr></thead><tbody>';

  incidents.forEach(inc => {
    const urgent = inc.incident_note && inc.incident_note.includes('[URGENT]');
    const pri = inc.priority || '';
    let rawNote = inc.incident_note || '';
    // Parse [MA:AGENCY:STATUS] tags (new) + legacy [MA] tag
    const maTagMatches = [...rawNote.matchAll(/\[MA:([^\]:]+):([^\]]+)\]/gi)];
    const maTags = maTagMatches.map(m => ({ agency: m[1].trim(), status: m[2].trim().toUpperCase() }));
    const isMutualAidLegacy = /\[MA\](?!:)/i.test(rawNote);
    const hasMutualAid = maTags.length > 0 || isMutualAidLegacy;
    let maBadge = '';
    if (maTags.length > 0) {
      maBadge = maTags.map(t => t.status === 'ACTIVE'
        ? '<span class="ma-active-badge">MA: ' + esc(t.agency) + '</span>'
        : '<span class="ma-badge">MA REQ: ' + esc(t.agency) + '</span>'
      ).join('');
    } else if (isMutualAidLegacy) {
      maBadge = '<span class="ma-badge">MA</span>';
    }
    const cbMatch = rawNote.match(/\[CB:([^\]]+)\]/i);
    const cbBadge = cbMatch ? '<span class="cb-badge">CB:' + esc(cbMatch[1].trim()) + '</span>' : '';
    const relIds = Array.isArray(inc.related_incidents) ? inc.related_incidents : [];
    const relBadge = relIds.map(function(rid) {
      const shortRel = String(rid).replace(/^\d{2}-/, '');
      return '<span class="rel-badge" title="Linked Incident" onclick="event.stopPropagation();openIncident(\'' + esc(rid) + '\')">REL:' + esc(shortRel) + '</span>';
    }).join('');
    const isPri1 = pri === 'PRI-1' || pri === 'CRITICAL';
    let rowCl = (urgent || isPri1 ? 'inc-urgent' : '') + (isPri1 ? ' inc-pri1-queue' : '') + (hasMutualAid ? ' inc-mutual-aid' : '');
    const mins = minutesSince(inc.created_at);
    const age = mins != null ? Math.floor(mins) + 'M' : '--';
    const waitMins = Math.floor((Date.now() - new Date(_normalizeTs(inc.created_at)).getTime()) / 60000);
    const isStale = waitMins >= 240;
    const staleBadge = isStale ? '<span class="stale-badge">STALE</span>' : '';
    if (isStale) rowCl += ' inc-stale';
    const waitCls = isStale ? 'inc-stale-wait blink' : waitMins > 20 ? 'inc-overdue' : waitMins > 10 ? 'inc-wait' : '';
    // HOLD tag — scheduled call with target time
    const holdM = rawNote.match(/\[HOLD:(\d{2}:\d{2})\]/i);
    let holdBadge = '';
    if (holdM) {
      const now = new Date();
      const hParts = holdM[1].split(':').map(Number);
      const holdDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), hParts[0], hParts[1], 0);
      const holdPast = now >= holdDate;
      holdBadge = holdPast
        ? '<span class="hold-badge hold-badge-due blink">\u23f0 ' + holdM[1] + '</span>'
        : '<span class="hold-badge hold-badge-pending">\u23f0 ' + holdM[1] + '</span>';
      if (!holdPast) rowCl += ' inc-hold-pending';
    }
    const shortId = inc.incident_id.replace(/^\d{2}-/, '');
    let note = rawNote.replace(/\[[^\]]*\]/g, '').replace(/\s{2,}/g, ' ').trim();
    const incType = inc.incident_type || '';
    const typeCl = getIncidentTypeClass(incType);
    const priBadge = pri ? `<span class="priority-${esc(pri)}" style="font-size:10px;font-weight:900;margin-left:4px;">${esc(pri)}</span>` : '';
    const locVal = (inc.level_of_care || '').toUpperCase();
    const locBadge = locVal ? `<span style="font-size:9px;font-weight:900;margin-left:3px;padding:1px 4px;background:rgba(121,192,255,.15);border:1px solid rgba(121,192,255,.35);color:#79c0ff;border-radius:2px;">${esc(locVal)}</span>` : '';
    const sceneDisplay = (inc.scene_address || '').substring(0, 20) || '—';

    html += `<tr class="${rowCl}" data-inc-id="${esc(inc.incident_id)}" onclick="openIncident('${esc(inc.incident_id)}')">`;
    html += `<td class="inc-id">${urgent ? 'HOT ' : ''}${esc(inc.incident_id)}${priBadge}${locBadge}${maBadge}${cbBadge}${relBadge}${staleBadge}${holdBadge}</td>`;
    const incDestResolved = AddressLookup.resolve(inc.destination);
    const incDestDisplay = incDestResolved.recognized ? incDestResolved.addr.name : (inc.destination || 'NO DEST');
    html += `<td class="inc-dest${incDestResolved.recognized ? ' dest-recognized' : ''}">${esc(incDestDisplay)}</td>`;
    html += `<td>${incType ? '<span class="inc-type ' + typeCl + '">' + esc(incType) + '</span>' : '<span class="muted">--</span>'}</td>`;
    html += `<td class="inc-note" title="${esc(note)}">${esc(note || '--')}</td>`;
    html += `<td style="font-size:11px;color:var(--muted);">${esc(sceneDisplay)}</td>`;
    html += `<td class="${waitCls}">${waitMins}M</td>`;
    html += `<td style="white-space:nowrap;">`;
    html += `<button class="toolbar-btn toolbar-btn-accent" onclick="event.stopPropagation(); assignIncidentToUnit('${esc(inc.incident_id)}')">ASSIGN</button> `;
    html += `<button class="toolbar-btn" onclick="event.stopPropagation(); openIncident('${esc(inc.incident_id)}')">REVIEW</button> `;
    html += `<button class="btn-danger mini" style="padding:3px 6px;font-size:10px;" onclick="event.stopPropagation(); closeIncidentFromQueue('${esc(inc.incident_id)}')">CLOSE</button>`;
    html += `</td>`;
    html += '</tr>';
  });

  html += '</tbody></table>';
  panel.innerHTML = html;
  updatePopoutStats();
}

function renderMessagesPanel() {
  const panel = document.getElementById('messagesPanel');
  if (!STATE) return;
  const m = STATE.messages || [];
  const unread = m.filter(msg => !msg.read).length;
  const countEl = document.getElementById('msgPanelCount');

  if (countEl) {
    countEl.textContent = m.length > 0 ? `(${m.length} TOTAL, ${unread} UNREAD)` : '';
  }

  if (!m.length) {
    panel.innerHTML = '<div class="muted" style="padding:20px;text-align:center;">NO MESSAGES</div>';
    return;
  }

  panel.innerHTML = m.map((msg, idx) => {
    const cl = ['messageDisplayItem'];
    if (msg.urgent) cl.push('urgent');
    const fr = msg.from_initials + '@' + msg.from_role;
    const fC = getRoleColor(fr);
    const uH = msg.urgent ? '[HOT] ' : '';
    const replyCmd = 'MSG ' + msg.from_role + ' ';
    const msgIdx = idx + 1; // 1-based local number
    return `<div class="${cl.join(' ')}">
      <div class="messageDisplayHeader ${fC}"><span class="muted" style="font-size:10px;margin-right:6px;">#${msgIdx}</span>${uH}FROM ${esc(fr)} TO ${esc(msg.to_role)}</div>
      <div class="messageDisplayText">${esc(msg.message)}</div>
      <div class="messageDisplayTime">${fmtTime24(msg.ts)}<button class="btn-secondary mini" style="margin-left:10px;" onclick="replyToMessage('${esc(replyCmd)}')">REPLY</button><button class="btn-danger mini" style="margin-left:6px;" onclick="deleteMessage('${esc(msg.message_id)}')">DEL</button></div>
    </div>`;
  }).join('');
}

// ============================================================
// Inbox Panel (live message display)
// ============================================================
function renderInboxPanel() {
  const panel = document.getElementById('msgInboxList');
  if (!panel) return;
  const m = STATE.messages || [];
  const unread = m.filter(msg => !msg.read).length;
  const badge = document.getElementById('inboxBadge');
  if (badge) badge.textContent = m.length > 0 ? `(${unread} NEW / ${m.length} TOTAL)` : '(EMPTY)';

  if (!m.length) {
    panel.innerHTML = '<div class="muted" style="padding:10px;text-align:center;">NO MESSAGES</div>';
    return;
  }

  panel.innerHTML = m.map(msg => {
    const cl = ['inbox-msg'];
    if (!msg.read) cl.push('unread');
    if (msg.urgent) cl.push('urgent');
    const fr = (msg.from_initials || '?') + '@' + (msg.from_role || '?');
    const ts = msg.ts ? fmtTime24(msg.ts) : '';
    const text = String(msg.message || '').substring(0, 120);
    const replyCmd = 'MSG ' + msg.from_role + ' ';
    return `<div class="${cl.join(' ')}" onclick="readAndReplyInbox('${esc(msg.message_id)}', '${esc(replyCmd)}')">
      <div><span class="inbox-from">${msg.urgent ? 'HOT ' : ''}${esc(fr)}</span> <span class="inbox-time">${esc(ts)}</span></div>
      <div class="inbox-text">${esc(text)}</div>
    </div>`;
  }).join('');
}

async function readAndReplyInbox(msgId, replyCmd) {
  if (TOKEN && msgId) {
    await API.readMessage(TOKEN, msgId);
  }
  const cmd = document.getElementById('cmd');
  if (cmd) {
    cmd.value = replyCmd;
    cmd.focus();
    cmd.setSelectionRange(replyCmd.length, replyCmd.length);
  }
  refresh();
}

// ============================================================
// Bottom Panel Toggle
// ============================================================
function toggleBottomPanel(panel) {
  const el = document.getElementById(panel === 'msgInbox' ? 'msgInboxPanel' : 'scratchPanel');
  if (el) el.classList.toggle('collapsed');
}

// ============================================================
// Scratch Notes (localStorage, per-user)
// ============================================================
function getScratchKey() {
  return 'hoscad_scratch_' + (ACTOR || 'anon');
}

function loadScratch() {
  const val = localStorage.getItem(getScratchKey()) || '';
  const pad = document.getElementById('scratchPad');
  if (pad) {
    pad.value = val;
    if (!pad.dataset.scratchAttached) {
      pad.addEventListener('input', saveScratch);
      pad.dataset.scratchAttached = '1';
    }
  }
  const side = document.getElementById('scratchPadSide');
  if (side) side.value = val;
}

function saveScratch() {
  const pad = document.getElementById('scratchPad');
  if (!pad) return;
  localStorage.setItem(getScratchKey(), pad.value);
  const side = document.getElementById('scratchPadSide');
  if (side) side.value = pad.value;
}

function saveScratchSide() {
  const side = document.getElementById('scratchPadSide');
  if (!side) return;
  localStorage.setItem(getScratchKey(), side.value);
  const pad = document.getElementById('scratchPad');
  if (pad) pad.value = side.value;
}

function renderBoard() {
  const tb = document.getElementById('boardBody');
  const q = document.getElementById('search').value.trim().toUpperCase();
  const sI = document.getElementById('showInactive').checked;
  const boardCountEl = document.getElementById('boardCount');

  let us = (STATE.units || []).filter(u => {
    if (!sI && !u.active) return false;
    // Filter assisting agency units if toggle is off
    if (!_showAssisting) {
      const t = (u.type || '').toLowerCase();
      if (t === 'law' || t === 'dot' || t === 'support') return false;
    }
    // Filter DC911 units if feed is disabled
    if (u.source && u.source.startsWith('DC911:') && !_isDc911Enabled()) return false;
    const h = (u.unit_id + ' ' + (u.display_name || '') + ' ' + (u.note || '') + ' ' + (u.destination || '') + ' ' + (u.incident || '')).toUpperCase();
    if (q && !h.includes(q)) return false;
    if (ACTIVE_INCIDENT_FILTER && String(u.incident || '') !== ACTIVE_INCIDENT_FILTER) return false;
    // VIEW filter
    if (VIEW.filterStatus) {
      const uSt = String(u.status || '').toUpperCase();
      if (uSt !== VIEW.filterStatus.toUpperCase()) return false;
    }
    return true;
  });

  // Sort based on VIEW.sort
  us.sort((a, b) => {
    let cmp = 0;
    switch (VIEW.sort) {
      case 'unit':
        cmp = String(a.unit_id || '').localeCompare(String(b.unit_id || ''));
        break;
      case 'elapsed': {
        const mA = minutesSince(a.updated_at) ?? -1;
        const mB = minutesSince(b.updated_at) ?? -1;
        cmp = mB - mA;
        break;
      }
      case 'updated': {
        const tA = a.updated_at ? new Date(_normalizeTs(a.updated_at)).getTime() : 0;
        const tB = b.updated_at ? new Date(_normalizeTs(b.updated_at)).getTime() : 0;
        cmp = tB - tA;
        break;
      }
      case 'status':
      default: {
        const ra = statusRank(a.status);
        const rb = statusRank(b.status);
        cmp = ra - rb;
        if (cmp === 0 && String(a.status || '').toUpperCase() === 'D') {
          const ta = a.updated_at ? new Date(_normalizeTs(a.updated_at)).getTime() : 0;
          const tbb = b.updated_at ? new Date(_normalizeTs(b.updated_at)).getTime() : 0;
          cmp = tbb - ta;
        }
        if (cmp === 0) cmp = String(a.unit_id || '').localeCompare(String(b.unit_id || ''));
        break;
      }
    }
    return VIEW.sortDir === 'desc' ? -cmp : cmp;
  });

  // Stale detection — expanded to D, DE, OS, T, TH
  const STALE_STATUSES = new Set(['D', 'DE', 'OS', 'T', 'TH']);
  const staleGroups = {};
  us.forEach(u => {
    if (!u.active) return;
    const st = String(u.status || '').toUpperCase();
    if (!STALE_STATUSES.has(st)) return;
    const mi = minutesSince(u.updated_at);
    if (mi != null && mi >= STATE.staleThresholds.CRITICAL) {
      if (!staleGroups[st]) staleGroups[st] = [];
      staleGroups[st].push(u.unit_id);
    }
  });

  // Welfare alert beep — fires once when unit first crosses CRITICAL threshold
  if (BASELINED) {
    us.forEach(u => {
      if (!u.active) return;
      const st = String(u.status || '').toUpperCase();
      if (!STALE_STATUSES.has(st)) return;
      const mi = minutesSince(u.updated_at);
      if (mi == null || mi < STATE.staleThresholds.CRITICAL) return;
      const wKey = u.unit_id + ':' + (u.updated_at || '');
      if (!_welfareAlertedKeys.has(wKey)) {
        _welfareAlertedKeys.add(wKey);
        beepAlert();
        showToast('WELFARE CHECK: ' + u.unit_id + ' IN ' + st + ' FOR ' + Math.floor(mi) + 'M', 'warn', 6000);
      }
    });
  }

  const ba = document.getElementById('staleBanner');
  const staleEntries = Object.keys(staleGroups).map(s =>
    'STALE ' + s + ' (&ge;' + STATE.staleThresholds.CRITICAL + 'M): ' +
    staleGroups[s].map(uid =>
      '<span class="stale-unit-link" onclick="scrollToUnit(\'' + esc(uid) + '\')">' + esc(uid) + '</span>' +
      '<button class="stale-welf-btn" onclick="event.stopPropagation();_execCmd(\'WELF ' + esc(uid) + '\')" title="Welfare check ' + esc(uid) + '">WELF</button>'
    ).join('&nbsp; '));
  if (staleEntries.length) {
    ba.style.display = 'block';
    ba.innerHTML = staleEntries.join(' &nbsp;|&nbsp; ');
  } else {
    ba.style.display = 'none';
  }

  const activeCount = us.filter(u => u.active).length;
  if (boardCountEl) boardCountEl.textContent = '(' + activeCount + ' ACTIVE)';

  tb.innerHTML = '';
  us.forEach(u => {
    const tr = document.createElement('tr');
    const mi = minutesSince(u.updated_at);

    // Stale classes — expanded to D, DE, OS, T
    if (u.active && STALE_STATUSES.has(String(u.status || '').toUpperCase()) && mi != null) {
      if (mi >= STATE.staleThresholds.CRITICAL) tr.classList.add('stale30');
      else if (mi >= STATE.staleThresholds.ALERT) tr.classList.add('stale20');
      else if (mi >= STATE.staleThresholds.WARN) tr.classList.add('stale10');
    }

    // Status row tint
    tr.classList.add('status-' + (u.status || '').toUpperCase());

    // Selected row
    if (SELECTED_UNIT_ID && String(u.unit_id || '').toUpperCase() === SELECTED_UNIT_ID) {
      tr.classList.add('selected');
    }

    // UNIT column
    const uId = (u.unit_id || '').toUpperCase();
    const di = (u.display_name || '').toUpperCase();
    const sD = di && di !== uId;
    const lvlBadge = u.level ? ' <span class="level-badge level-' + esc(u.level) + '">' + esc(u.level) + '</span>' : '';
    const dc911Badge = (() => {
      if (!u.source || !u.source.startsWith('DC911:')) return '';
      const t = (u.type || 'EMS').toUpperCase();
      const cls = t === 'FIRE' ? 'dc911-fire' : t === 'LAW' ? 'dc911-law' : 'dc911-ems';
      return ' <span class="dc911-badge ' + cls + '">911·' + t + '</span>';
    })();
    const crewParts = u.unit_info ? String(u.unit_info).split('|').filter(p => /^CM\d:/i.test(p)) : [];
    const crewHtml = crewParts.length ? '<div class="crew-sub">' + crewParts.map(p => esc(p.replace(/^CM\d:/i, '').trim())).join(' / ') + '</div>' : '';
    const unitHtml = '<span class="unit">' + esc(uId) + '</span>' + lvlBadge + dc911Badge +
      (u.active ? '' : ' <span class="muted">(I)</span>') +
      (sD ? ' <span class="muted" style="font-size:10px;">' + esc(di) + '</span>' : '') +
      crewHtml;

    // STATUS column — CAD reader board style: large code + dimmer label
    const sL = (STATE.statuses || []).find(s => s.code === u.status)?.label || u.status;
    const stCode = (u.status || '').toUpperCase();
    const statusHtml = '<span class="statusCode status-text-' + esc(stCode) + '" title="' + esc(sL) + '">' + esc(stCode) + '</span>';

    // ELAPSED column — OS gets its own thresholds (15/30/45m); others use global stale thresholds
    const elapsedVal = formatElapsed(mi);
    let elapsedClass = 'elapsed-cell';
    if (mi != null) {
      if (stCode === 'OS') {
        if (mi >= 45) elapsedClass += ' elapsed-critical';
        else if (mi >= 30) elapsedClass += ' elapsed-os-alert';
        else if (mi >= 15) elapsedClass += ' elapsed-os-warn';
      } else if (STALE_STATUSES.has(stCode)) {
        if (STATE.staleThresholds && mi >= STATE.staleThresholds.CRITICAL) elapsedClass += ' elapsed-critical';
        else if (STATE.staleThresholds && mi >= STATE.staleThresholds.WARN) elapsedClass += ' elapsed-warn';
      }
    }

    // LOCATION column
    const destHtml = AddressLookup.formatLocation(u);

    // NOTES column — incident notes if on incident, status notes otherwise
    let noteText = '';
    if (u.incident) {
      const incObj = (STATE.incidents || []).find(i => i.incident_id === u.incident);
      if (incObj && incObj.incident_note) noteText = incObj.incident_note.replace(/\[[^\]]*\]/g, '').replace(/\s{2,}/g, ' ').trim();
    }
    if (!noteText) noteText = (u.note || '').replace(/^\[OOS:[^\]]+\]\s*/, '').replace(/\[LOC:[^\]]*\]\s*/g, '').replace(/\[ETA:\d+\]\s*/g, '');
    noteText = noteText.toUpperCase();
    const oosReason = u.oos_reason || ((u.note || '').match(/^\[OOS:([^\]]+)\]/) || [])[1] || '';
    const oosBadge = oosReason ? '<span class="oos-badge">' + esc(oosReason) + '</span>' : '';
    const patMatch = (u.note || '').match(/\[PAT:([^\]]+)\]/);
    const patBadge = patMatch ? '<span class="pat-badge">PAT:' + esc(patMatch[1]) + '</span>' : '';
    const locMatch = (u.note || '').match(/\[LOC:([^\]]+)\]/);
    const locBadge = locMatch ? '<span class="loc-badge" title="' + esc(locMatch[1]) + '">LOC</span>' : '';
    const etaMatch = (u.note || '').match(/\[ETA:(\d+)\]/);
    const etaBadge = etaMatch ? '<span class="eta-badge" title="ETA SET">ETA:' + esc(etaMatch[1]) + 'M</span>' : '';
    // ASSIST badge — for law/dot/support units or units explicitly excluded from recommendations
    const uTypeL = (u.type || '').toLowerCase();
    const isAssistType = uTypeL === 'law' || uTypeL === 'dot' || uTypeL === 'support';
    const assistBadge = isAssistType ? '<span class="cap-badge-assist">ASSIST</span>' : '';
    const noteHtml = (noteText ? '<span class="noteBig">' + esc(noteText) + '</span>' : '<span class="muted">—</span>') + oosBadge + patBadge + locBadge + etaBadge + assistBadge;

    // INC# column — with type dot
    let incHtml = '<span class="muted">—</span>';
    let groupBorderColor = '';
    if (u.incident) {
      const shortInc = String(u.incident).replace(/^\d{2}-/, '');
      let dotHtml = '';
      const incObj = (STATE.incidents || []).find(i => i.incident_id === u.incident);
      if (incObj && incObj.incident_type) {
        const typCl = getIncidentTypeClass(incObj.incident_type);
        const dotCl = typCl.replace('inc-type-', 'inc-type-dot-');
        if (dotCl) dotHtml = '<span class="inc-type-dot ' + dotCl + '"></span>';
        // Group border: show if 2+ active units share this incident
        const sharedCount = (STATE.units || []).filter(ou => ou.active && ou.unit_id !== u.unit_id && ou.incident === u.incident).length;
        if (sharedCount > 0) groupBorderColor = INC_GROUP_BORDER[typCl] || '#6a7a8a';
      }
      const stackData = getUnitStackData(u.unit_id);
      const stackBadgeHtml = stackData ? renderStackBadge(stackData.depth, stackData.hasUrgent, u.unit_id) : '';
      incHtml = dotHtml + '<span class="clickableIncidentNum" onclick="event.stopPropagation(); openIncident(\'' + esc(u.incident) + '\')">' + esc(u.incident) + '</span>' + stackBadgeHtml;
    }
    // Apply border-left: incident group border takes priority; fall back to unit type accent
    if (groupBorderColor) {
      tr.style.borderLeft = '3px solid ' + groupBorderColor;
    } else if (uTypeL === 'law') {
      tr.style.borderLeft = '3px solid #4a6fa5';
    } else if (uTypeL === 'dot') {
      tr.style.borderLeft = '3px solid #e6841a';
    } else if (uTypeL === 'support') {
      tr.style.borderLeft = '3px solid #888';
    }

    // TYPE column — incident type from incObj (already fetched in INC# block)
    let typeHtml = '<span class="muted">—</span>';
    if (u.incident) {
      const incObjForType = (STATE.incidents || []).find(i => i.incident_id === u.incident);
      if (incObjForType && incObjForType.incident_type) {
        const typCl = getIncidentTypeClass(incObjForType.incident_type);
        typeHtml = '<span class="inc-type ' + typCl + '" style="font-size:10px;">' + esc(incObjForType.incident_type) + '</span>';
      }
    }

    // UPDATED column
    const aC = getRoleColor(u.updated_by);
    const updatedHtml = fmtTime24(u.updated_at) + ' <span class="muted ' + aC + '" style="font-size:10px;">' + esc((u.updated_by || '').toUpperCase()) + '</span>';

    tr.innerHTML = '<td>' + unitHtml + '</td>' +
      '<td>' + statusHtml + '</td>' +
      '<td class="' + elapsedClass + '">' + elapsedVal + '</td>' +
      '<td>' + destHtml + '</td>' +
      '<td>' + noteHtml + '</td>' +
      '<td>' + typeHtml + '</td>' +
      '<td>' + incHtml + '</td>' +
      '<td>' + updatedHtml + '</td>';

    // Single-click = select row
    tr.onclick = (e) => {
      e.stopPropagation();
      selectUnit(u.unit_id);
    };

    // Double-click = open edit modal
    tr.ondblclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (u.source && u.source.startsWith('DC911:')) { showToast('DC911 UNIT — READ-ONLY. DATA FROM DESCHUTES 911 CAD FEED.', 'info'); return; }
      if (u.source && u.source.startsWith('PP:')) { openPpUnitActions(u); return; }
      openModal(u);
    };

    tr.style.cursor = 'pointer';
    tb.appendChild(tr);
  });
  updatePopoutStats();
}

// Performance: DOM diffing version — only updates changed rows
function renderBoardDiff() {
  const tb = document.getElementById('boardBody');
  const q = document.getElementById('search').value.trim().toUpperCase();
  const sI = document.getElementById('showInactive').checked;
  const boardCountEl = document.getElementById('boardCount');

  // Pre-compute uppercase filter status once (not in loop)
  const filterStatusUpper = VIEW.filterStatus ? VIEW.filterStatus.toUpperCase() : null;

  let us = (STATE.units || []).filter(u => {
    if (!sI && !u.active) return false;
    // Filter assisting agency units if toggle is off
    if (!_showAssisting) {
      const t = (u.type || '').toLowerCase();
      if (t === 'law' || t === 'dot' || t === 'support') return false;
    }
    // Filter DC911 units if feed is disabled
    if (u.source && u.source.startsWith('DC911:') && !_isDc911Enabled()) return false;
    const h = (u.unit_id + ' ' + (u.display_name || '') + ' ' + (u.note || '') + ' ' + (u.destination || '') + ' ' + (u.incident || '')).toUpperCase();
    if (q && !h.includes(q)) return false;
    if (ACTIVE_INCIDENT_FILTER && String(u.incident || '') !== ACTIVE_INCIDENT_FILTER) return false;
    if (filterStatusUpper) {
      if (String(u.status || '').toUpperCase() !== filterStatusUpper) return false;
    }
    return true;
  });

  // Pre-compute timestamps for sorting (avoid new Date() in comparator)
  const tsCache = new Map();
  us.forEach(u => {
    tsCache.set(u.unit_id, u.updated_at ? new Date(_normalizeTs(u.updated_at)).getTime() : 0);
  });

  us.sort((a, b) => {
    let cmp = 0;
    switch (VIEW.sort) {
      case 'unit':
        cmp = String(a.unit_id || '').localeCompare(String(b.unit_id || ''));
        break;
      case 'elapsed':
      case 'updated': {
        cmp = tsCache.get(b.unit_id) - tsCache.get(a.unit_id);
        break;
      }
      case 'status':
      default: {
        const ra = statusRank(a.status);
        const rb = statusRank(b.status);
        cmp = ra - rb;
        if (cmp === 0 && String(a.status || '').toUpperCase() === 'D') {
          cmp = tsCache.get(b.unit_id) - tsCache.get(a.unit_id);
        }
        if (cmp === 0) cmp = String(a.unit_id || '').localeCompare(String(b.unit_id || ''));
        break;
      }
    }
    return VIEW.sortDir === 'desc' ? -cmp : cmp;
  });

  // Stale detection
  const STALE_STATUSES = new Set(['D', 'DE', 'OS', 'T', 'TH']);
  const staleGroups = {};
  us.forEach(u => {
    if (!u.active) return;
    const st = String(u.status || '').toUpperCase();
    if (!STALE_STATUSES.has(st)) return;
    const mi = minutesSince(u.updated_at);
    if (mi != null && mi >= STATE.staleThresholds.CRITICAL) {
      if (!staleGroups[st]) staleGroups[st] = [];
      staleGroups[st].push(u.unit_id);
    }
  });

  // Welfare alert beep — fires once when unit first crosses CRITICAL threshold
  if (BASELINED) {
    us.forEach(u => {
      if (!u.active) return;
      const st = String(u.status || '').toUpperCase();
      if (!STALE_STATUSES.has(st)) return;
      const mi = minutesSince(u.updated_at);
      if (mi == null || mi < STATE.staleThresholds.CRITICAL) return;
      const wKey = u.unit_id + ':' + (u.updated_at || '');
      if (!_welfareAlertedKeys.has(wKey)) {
        _welfareAlertedKeys.add(wKey);
        beepAlert();
        showToast('WELFARE CHECK: ' + u.unit_id + ' IN ' + st + ' FOR ' + Math.floor(mi) + 'M', 'warn', 6000);
      }
    });
  }

  const ba = document.getElementById('staleBanner');
  const staleEntries = Object.keys(staleGroups).map(s =>
    'STALE ' + s + ' (&ge;' + STATE.staleThresholds.CRITICAL + 'M): ' +
    staleGroups[s].map(uid =>
      '<span class="stale-unit-link" onclick="scrollToUnit(\'' + esc(uid) + '\')">' + esc(uid) + '</span>' +
      '<button class="stale-welf-btn" onclick="event.stopPropagation();_execCmd(\'WELF ' + esc(uid) + '\')" title="Welfare check ' + esc(uid) + '">WELF</button>'
    ).join('&nbsp; '));
  if (staleEntries.length) {
    ba.style.display = 'block';
    ba.innerHTML = staleEntries.join(' &nbsp;|&nbsp; ');
  } else {
    ba.style.display = 'none';
  }

  const activeCount = us.filter(u => u.active).length;
  if (boardCountEl) boardCountEl.textContent = '(' + activeCount + ' ACTIVE)';

  // Build new row order
  const newOrder = us.map(u => u.unit_id);
  const existingRows = tb.querySelectorAll('tr[data-unit-id]');
  const existingMap = new Map();
  existingRows.forEach(tr => existingMap.set(tr.dataset.unitId, tr));

  // Precompute incident lookup map — O(1) per unit vs O(n) find() per unit
  const incidentMap = new Map((STATE.incidents || []).map(i => [i.incident_id, i]));

  // Track which rows we've processed
  const processedIds = new Set();

  // Build/update rows using DocumentFragment for batch insert
  const fragment = document.createDocumentFragment();

  us.forEach((u, idx) => {
    const unitId = u.unit_id;
    processedIds.add(unitId);

    // Generate row hash to check if update needed
    // Include linked incident's last_update so note changes on the incident invalidate the row
    const _incObj = u.incident && STATE.incidents ? (STATE.incidents.find(i => i.incident_id === u.incident) || {}) : {};
    const _iLU = _incObj.last_update || '';
    const _iSA = _incObj.scene_address || '';
    const rowHash = unitId + '|' + (u.status || '') + '|' + (u.updated_at || '') + '|' + (u.destination || '') + '|' + (u.note || '') + '|' + (u.incident || '') + '|' + (u.active ? '1' : '0') + '|' + (u.level || '') + '|' + _iLU + '|' + (u.unit_info || '') + '|' + (u.source || '') + '|' + _iSA + '|' + (_expandedStacks.has(unitId) ? 'E' : '');
    const cached = _rowCache.get(unitId);

    let tr = existingMap.get(unitId);

    // If row exists and hash matches, just reposition if needed
    if (tr && cached && cached.hash === rowHash) {
      // Update stale/selected classes only
      updateRowClasses(tr, u, STALE_STATUSES);
      fragment.appendChild(tr);
      _appendStackSubRows(unitId, fragment);
      return;
    }

    // Build new row HTML
    const mi = minutesSince(u.updated_at);

    // Build classes
    let rowClasses = 'status-' + (u.status || '').toUpperCase();
    const stCode = (u.status || '').toUpperCase();
    if (u.active && STALE_STATUSES.has(stCode) && mi != null) {
      if (mi >= STATE.staleThresholds.CRITICAL) rowClasses += ' stale30';
      else if (mi >= STATE.staleThresholds.ALERT) rowClasses += ' stale20';
      else if (mi >= STATE.staleThresholds.WARN) rowClasses += ' stale10';
    }
    if (SELECTED_UNIT_ID && String(unitId).toUpperCase() === SELECTED_UNIT_ID) {
      rowClasses += ' selected';
    }

    // UNIT column
    const uId = (u.unit_id || '').toUpperCase();
    const di = (u.display_name || '').toUpperCase();
    const sD = di && di !== uId;
    const lvlBadge = u.level ? ' <span class="level-badge level-' + esc(u.level) + '">' + esc(u.level) + '</span>' : '';
    const dc911Badge = (() => {
      if (!u.source || !u.source.startsWith('DC911:')) return '';
      const t = (u.type || 'EMS').toUpperCase();
      const cls = t === 'FIRE' ? 'dc911-fire' : t === 'LAW' ? 'dc911-law' : 'dc911-ems';
      return ' <span class="dc911-badge ' + cls + '">911·' + t + '</span>';
    })();
    const crewParts = u.unit_info ? String(u.unit_info).split('|').filter(p => /^CM\d:/i.test(p)) : [];
    const crewHtml = crewParts.length ? '<div class="crew-sub">' + crewParts.map(p => esc(p.replace(/^CM\d:/i, '').trim())).join(' / ') + '</div>' : '';
    const unitHtml = '<span class="unit">' + esc(uId) + '</span>' + lvlBadge + dc911Badge +
      (u.active ? '' : ' <span class="muted">(I)</span>') +
      (sD ? ' <span class="muted" style="font-size:10px;">' + esc(di) + '</span>' : '') +
      crewHtml;

    // STATUS column — CAD reader board style: large code + dimmer label
    const sL = (STATE.statuses || []).find(s => s.code === u.status)?.label || u.status;
    const statusHtml = '<span class="statusCode status-text-' + esc(stCode) + '" title="' + esc(sL) + '">' + esc(stCode) + '</span>';

    // ELAPSED column — OS gets its own thresholds (15/30/45m); others use global stale thresholds
    const elapsedVal = formatElapsed(mi);
    let elapsedClass = 'elapsed-cell';
    if (mi != null) {
      if (stCode === 'OS') {
        if (mi >= 45) elapsedClass += ' elapsed-critical';
        else if (mi >= 30) elapsedClass += ' elapsed-os-alert';
        else if (mi >= 15) elapsedClass += ' elapsed-os-warn';
      } else if (STALE_STATUSES.has(stCode)) {
        if (STATE.staleThresholds && mi >= STATE.staleThresholds.CRITICAL) elapsedClass += ' elapsed-critical';
        else if (STATE.staleThresholds && mi >= STATE.staleThresholds.WARN) elapsedClass += ' elapsed-warn';
      }
    }

    // LOCATION column
    const destHtml = AddressLookup.formatLocation(u);

    // NOTES column
    let noteText = '';
    if (u.incident) {
      const incObj = incidentMap.get(u.incident);
      if (incObj && incObj.incident_note) noteText = incObj.incident_note.replace(/\[[^\]]*\]/g, '').replace(/\s{2,}/g, ' ').trim();
    }
    if (!noteText) noteText = (u.note || '').replace(/^\[OOS:[^\]]+\]\s*/, '').replace(/\[LOC:[^\]]*\]\s*/g, '').replace(/\[ETA:\d+\]\s*/g, '');
    noteText = noteText.toUpperCase();
    const oosReason2 = u.oos_reason || ((u.note || '').match(/^\[OOS:([^\]]+)\]/) || [])[1] || '';
    const oosBadge = oosReason2 ? '<span class="oos-badge">' + esc(oosReason2) + '</span>' : '';
    const patMatch = (u.note || '').match(/\[PAT:([^\]]+)\]/);
    const patBadge = patMatch ? '<span class="pat-badge">PAT:' + esc(patMatch[1]) + '</span>' : '';
    const locMatch2 = (u.note || '').match(/\[LOC:([^\]]+)\]/);
    const locBadge2 = locMatch2 ? '<span class="loc-badge" title="' + esc(locMatch2[1]) + '">LOC</span>' : '';
    const etaMatch2 = (u.note || '').match(/\[ETA:(\d+)\]/);
    const etaBadge2 = etaMatch2 ? '<span class="eta-badge" title="ETA SET">ETA:' + esc(etaMatch2[1]) + 'M</span>' : '';
    // ASSIST badge — for law/dot/support units or units explicitly excluded from recommendations
    const uTypeL2 = (u.type || '').toLowerCase();
    const isAssistType2 = uTypeL2 === 'law' || uTypeL2 === 'dot' || uTypeL2 === 'support';
    const assistBadge2 = (isAssistType2 || u.include_in_recommendations === false) ? '<span class="cap-badge-assist">ASSIST</span>' : '';
    const noteHtml = (noteText ? '<span class="noteBig">' + esc(noteText) + '</span>' : '<span class="muted">—</span>') + oosBadge + patBadge + locBadge2 + etaBadge2 + assistBadge2;

    // INC# column
    let incHtml = '<span class="muted">—</span>';
    let groupBorderColor2 = '';
    if (u.incident) {
      const shortInc = String(u.incident).replace(/^\d{2}-/, '');
      let dotHtml = '';
      const incObj = incidentMap.get(u.incident);
      if (incObj && incObj.incident_type) {
        const typCl2 = getIncidentTypeClass(incObj.incident_type);
        const dotCl = typCl2.replace('inc-type-', 'inc-type-dot-');
        if (dotCl) dotHtml = '<span class="inc-type-dot ' + dotCl + '"></span>';
        const sharedCount2 = (STATE.units || []).filter(ou => ou.active && ou.unit_id !== u.unit_id && ou.incident === u.incident).length;
        if (sharedCount2 > 0) groupBorderColor2 = INC_GROUP_BORDER[typCl2] || '#6a7a8a';
      }
      const stackData2 = getUnitStackData(u.unit_id);
      const stackBadgeHtml2 = stackData2 ? renderStackBadge(stackData2.depth, stackData2.hasUrgent, u.unit_id) : '';
      incHtml = dotHtml + '<span class="clickableIncidentNum" data-inc="' + esc(u.incident) + '">' + esc(u.incident) + '</span>' + stackBadgeHtml2;
    }

    // Compute border-left: incident group border takes priority; fall back to unit type accent
    const typeBorderStyle = groupBorderColor2 ? '3px solid ' + groupBorderColor2
      : uTypeL2 === 'law'     ? '3px solid #4a6fa5'
      : uTypeL2 === 'dot'     ? '3px solid #e6841a'
      : uTypeL2 === 'support' ? '3px solid #888'
      : '';

    // UPDATED column
    const aC = getRoleColor(u.updated_by);
    const updatedHtml = fmtTime24(u.updated_at) + ' <span class="muted ' + aC + '" style="font-size:10px;">' + esc((u.updated_by || '').toUpperCase()) + '</span>';

    const rowHtml = '<td>' + unitHtml + '</td>' +
      '<td>' + statusHtml + '</td>' +
      '<td class="' + elapsedClass + '">' + elapsedVal + '</td>' +
      '<td>' + destHtml + '</td>' +
      '<td>' + noteHtml + '</td>' +
      '<td>' + incHtml + '</td>' +
      '<td>' + updatedHtml + '</td>';

    if (tr) {
      // Update existing row
      tr.className = rowClasses;
      tr.innerHTML = rowHtml;
      tr.style.borderLeft = typeBorderStyle;
      tr.classList.add('row-flash');
      tr.addEventListener('animationend', () => tr.classList.remove('row-flash'), { once: true });
    } else {
      // Create new row
      tr = document.createElement('tr');
      tr.dataset.unitId = unitId;
      tr.className = rowClasses;
      tr.innerHTML = rowHtml;
      tr.style.cursor = 'pointer';
      tr.style.borderLeft = typeBorderStyle;
    }

    // Cache the row
    _rowCache.set(unitId, { hash: rowHash });

    fragment.appendChild(tr);
    _appendStackSubRows(unitId, fragment);
  });

  // Clear and append all at once
  tb.innerHTML = '';
  tb.appendChild(fragment);

  // Clean up cache for removed units
  for (const key of _rowCache.keys()) {
    if (!processedIds.has(key)) _rowCache.delete(key);
  }

  // Keep quick-action bar current after every board render
  updateQuickBar();
}

// Helper: update row classes without rebuilding HTML
function updateRowClasses(tr, u, STALE_STATUSES) {
  const mi = minutesSince(u.updated_at);
  const stCode = (u.status || '').toUpperCase();

  let classes = ['status-' + stCode];

  if (u.active && STALE_STATUSES.has(stCode) && mi != null) {
    if (mi >= STATE.staleThresholds.CRITICAL) classes.push('stale30');
    else if (mi >= STATE.staleThresholds.ALERT) classes.push('stale20');
    else if (mi >= STATE.staleThresholds.WARN) classes.push('stale10');
  }

  if (SELECTED_UNIT_ID && String(u.unit_id).toUpperCase() === SELECTED_UNIT_ID) {
    classes.push('selected');
  }

  tr.className = classes.join(' ');
}

function selectUnit(unitId) {
  const id = String(unitId || '').toUpperCase();
  const prevId = SELECTED_UNIT_ID;

  if (SELECTED_UNIT_ID === id) {
    // Deselecting — collapse stack if it was auto-expanded
    SELECTED_UNIT_ID = null;
    if (prevId) _expandedStacks.delete(prevId);
  } else {
    // Deselect previous — collapse its auto-expanded stack
    if (prevId) _expandedStacks.delete(prevId);
    SELECTED_UNIT_ID = id;
    // Auto-expand stack if unit has multiple assignments
    const stackData = getUnitStackData(id);
    if (stackData && stackData.depth > 1) _expandedStacks.add(id);
  }

  // Performance: Use data-unit-id attribute for O(1) lookup instead of text parsing
  const tb = document.getElementById('boardBody');
  const rows = tb.querySelectorAll('tr[data-unit-id]');
  rows.forEach(tr => {
    if (SELECTED_UNIT_ID && tr.dataset.unitId.toUpperCase() === SELECTED_UNIT_ID) {
      tr.classList.add('selected');
    } else {
      tr.classList.remove('selected');
    }
  });
  renderBoardDiff();
  updateQuickBar();
  autoFocusCmd();
}

// Select a unit and scroll its board row into view
function scrollToUnit(unitId) {
  selectUnit(unitId);
  setTimeout(() => {
    const tb = document.getElementById('boardBody');
    if (!tb) return;
    const row = tb.querySelector('tr[data-unit-id="' + String(unitId).toUpperCase() + '"]');
    if (row) row.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }, 50); // brief delay so renderBoardDiff can update the DOM
}

function getStatusLabel(code) {
  if (!STATE || !STATE.statuses) return code;
  const s = STATE.statuses.find(s => s.code === code);
  return s ? s.label : code;
}

// ============================================================
// Column Sort Setup
// ============================================================
function setupColumnSort() {
  document.querySelectorAll('.board-table th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const sortKey = th.dataset.sort;
      if (VIEW.sort === sortKey) {
        VIEW.sortDir = VIEW.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        VIEW.sort = sortKey;
        VIEW.sortDir = 'asc';
      }
      // Sync toolbar dropdown
      const tbSort = document.getElementById('tbSort');
      if (tbSort) tbSort.value = VIEW.sort;
      saveViewState();
      updateSortHeaders();
      renderBoardDiff();
    });
  });
}

// ============================================================
// Quick Actions
// ============================================================

/** Update the quick-action bar to reflect the currently selected unit. */
function updateQuickBar() {
  const bar = document.getElementById('quickBar');
  if (!bar) return;

  if (!SELECTED_UNIT_ID) {
    bar.style.display = 'none';
    const qbNote = document.getElementById('qbNote');
    if (qbNote) qbNote.value = '';
    return;
  }

  const u = STATE && STATE.units ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === SELECTED_UNIT_ID) : null;
  if (!u) { bar.style.display = 'none'; return; }

  bar.style.display = 'flex';
  const qbUnit = document.getElementById('qbUnit');
  const qbStatus = document.getElementById('qbStatus');
  if (qbUnit) qbUnit.textContent = u.unit_id;
  if (qbStatus) qbStatus.textContent = u.status + (u.incident ? ' · ' + u.incident : '');

  // Disable the button that matches current status
  bar.querySelectorAll('.qb-btn').forEach(btn => {
    const code = btn.getAttribute('onclick').match(/'([^']+)'/)?.[1];
    btn.disabled = code === u.status;
  });
}

/**
 * Push a reversible action onto the undo stack.
 * @param {string} description  Short label shown in toast (e.g. "M1: AV→DP")
 * @param {Function} revertFn   async function that performs the reversal; throw on failure
 */
function pushUndo(description, revertFn) {
  _undoStack.push({ description, revertFn, ts: Date.now() });
  if (_undoStack.length > 3) _undoStack.shift(); // keep last 3 only
}

/** Called by quick-action bar buttons — sets selected unit to status code. */
async function qbStatus(code) {
  if (!SELECTED_UNIT_ID) return;
  const u = STATE && STATE.units ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === SELECTED_UNIT_ID) : null;
  if (!u) return;
  const btn = document.querySelector('.qb-' + code);
  if (btn) btn.disabled = true;
  const note = (document.getElementById('qbNote')?.value || '').trim().toUpperCase();
  let oosReason = null;
  if (code === 'OOS') {
    const reason = await promptOOSReason(SELECTED_UNIT_ID);
    if (!reason) { if (btn) btn.disabled = false; return; }
    oosReason = reason;
  }
  const patch = { status: code };
  if (oosReason) patch.oosReason = oosReason;
  if (note) patch.note = note;
  const _prevQb = { uid: u.unit_id, status: u.status, note: u.note || '', incident: u.incident || '', destination: u.destination || '' };
  setLive(true, 'LIVE • UPDATE');
  const r = await API.upsertUnit(TOKEN, u.unit_id, patch, u.updated_at || '');
  if (btn) btn.disabled = false;
  if (!r.ok) return showErr(r);
  pushUndo(`${_prevQb.uid}: ${_prevQb.status}→${code}`, async () => {
    const rv = await API.upsertUnit(TOKEN, _prevQb.uid, { status: _prevQb.status, note: _prevQb.note, incident: _prevQb.incident, destination: _prevQb.destination }, '');
    if (!rv.ok) throw new Error(rv.error || 'API error');
  });
  document.getElementById('qbNote').value = '';
  refresh();
}

function quickStatus(u, c) {
  const msg = 'SET ' + u.unit_id + ' → ' + c + '?' + (c === 'AV' && (u.incident || u.destination || u.note) ? '\n\nNOTE: AV CLEARS INCIDENT.' : '');
  const _prevQs = { status: u.status, note: u.note || '', incident: u.incident || '', destination: u.destination || '' };
  showConfirm('CONFIRM STATUS CHANGE', msg, async () => {
    setLive(true, 'LIVE • UPDATE');
    const r = await API.upsertUnit(TOKEN, u.unit_id, { status: c, displayName: u.display_name }, u.updated_at || '');
    if (!r.ok) return showErr(r);
    pushUndo(`${u.unit_id}: ${_prevQs.status}→${c}`, async () => {
      const rv = await API.upsertUnit(TOKEN, u.unit_id, { status: _prevQs.status, note: _prevQs.note, incident: _prevQs.incident, destination: _prevQs.destination }, '');
      if (!rv.ok) throw new Error(rv.error || 'API error');
    });
    refresh();
    autoFocusCmd();
  });
}

async function okUnit(u) {
  if (!u || !u.unit_id) return;
  setLive(true, 'LIVE • OK');
  const r = await API.touchUnit(TOKEN, u.unit_id, u.updated_at || '');
  if (!r || !r.ok) return showErr(r);
  refresh();
  autoFocusCmd();
}

function okAllOS() {
  showConfirm('CONFIRM OKALL', 'OKALL: RESET STATIC TIMER FOR ALL ON SCENE (OS) UNITS?', async () => {
    setLive(true, 'LIVE • OKALL');
    const r = await API.touchAllOS(TOKEN);
    if (!r || !r.ok) return showErr(r);
    refresh();
    autoFocusCmd();
  });
}

function undoUnit(uId) {
  showConfirm('CONFIRM UNDO', 'UNDO LAST ACTION FOR ' + uId + '?', async () => {
    setLive(true, 'LIVE • UNDO');
    const r = await API.undoUnit(TOKEN, uId);
    if (!r.ok) return showErr(r);
    refresh();
    autoFocusCmd();
  });
}

// ============================================================
// Modal Functions
// ============================================================
function openPpUnitActions(u) {
  showConfirmAsync('PP UNIT: ' + u.unit_id, 'SOURCE: ' + (u.source || 'UNKNOWN') + '\n\nLOG THIS UNIT OFF THE BOARD?')
    .then(function(confirmed) {
      if (!confirmed) return;
      API.upsertUnit(TOKEN, u.unit_id, { active: false, status: 'AV', source: null }, u.updated_at || '')
        .then(function(r) {
          if (r && r.ok) { showToast(u.unit_id + ' LOGGED OFF.'); refresh(); }
          else showToast((r && r.error) || 'ERROR LOGGING OFF UNIT.');
        });
    });
}

function openModal(u, f = false) {
  _MODAL_UNIT = u;
  const b = document.getElementById('modalBack');
  b.style.display = 'flex';
  document.getElementById('mUnitId').value = u ? u.unit_id : '';
  document.getElementById('mDisplayName').value = u ? (u.display_name || '') : '';
  document.getElementById('mType').value = u ? (u.type || '') : '';
  document.getElementById('mStatus').value = u ? u.status : 'AV';
  const destEl = document.getElementById('mDestination');
  if (u && u.destination) {
    const resolved = AddressLookup.resolve(u.destination);
    destEl.value = resolved.displayText;
    if (resolved.recognized) destEl.dataset.addrId = resolved.addr.id;
    else delete destEl.dataset.addrId;
  } else {
    destEl.value = '';
    delete destEl.dataset.addrId;
  }
  document.getElementById('mIncident').value = u ? (u.incident || '') : '';
  // Parse [LOC:...] from note
  const rawNote = u ? (u.note || '') : '';
  const locMatch = rawNote.match(/\[LOC:([^\]]*)\]/);
  const mLocationEl = document.getElementById('mLocation');
  if (mLocationEl) mLocationEl.value = locMatch ? locMatch[1] : '';
  document.getElementById('mNote').value = rawNote.replace(/\s*\[LOC:[^\]]*\]\s*/g, '').trim();
  // Parse unit_info into structured crew fields + notes
  const rawInfo = u ? (u.unit_info || '') : '';
  const infoParts = rawInfo.split('|').map((p) => p.trim()).filter(Boolean);
  let cm1Name = '', cm1Cert = '', cm2Name = '', cm2Cert = '', extraNotes = '';
  infoParts.forEach((p) => {
    const crewMatch = p.match(/^CM([12]):(.*?)\s*\(([^)]+)\)\s*$/);
    if (crewMatch) {
      if (crewMatch[1] === '1') { cm1Name = crewMatch[2].trim(); cm1Cert = crewMatch[3].trim(); }
      else                      { cm2Name = crewMatch[2].trim(); cm2Cert = crewMatch[3].trim(); }
    } else if (!p.startsWith('CM')) {
      extraNotes = p;
    }
  });
  const mCrew1Name = document.getElementById('mCrew1Name');
  const mCrew1Cert = document.getElementById('mCrew1Cert');
  const mCrew2Name = document.getElementById('mCrew2Name');
  const mCrew2Cert = document.getElementById('mCrew2Cert');
  const mUnitNotes = document.getElementById('mUnitNotes');
  if (mCrew1Name) mCrew1Name.value = cm1Name;
  if (mCrew1Cert) mCrew1Cert.value = cm1Cert;
  if (mCrew2Name) mCrew2Name.value = cm2Name;
  if (mCrew2Cert) mCrew2Cert.value = cm2Cert;
  if (mUnitNotes) mUnitNotes.value = extraNotes;
  const mLevel = document.getElementById('mLevel');
  const mStation = document.getElementById('mStation');
  if (mLevel) mLevel.value = u ? (u.level || '') : '';
  if (mStation) mStation.value = u ? (u.station || '') : '';
  document.getElementById('modalTitle').textContent = u ? 'EDIT ' + u.unit_id : 'LOGON UNIT';
  document.getElementById('modalFoot').textContent = u ? 'UPDATED: ' + (u.updated_at || '—') + ' BY ' + (u.updated_by || '—') : 'TIP: SET STATUS TO D WITH INCIDENT BLANK TO AUTO-GENERATE.';
  b.dataset.expectedUpdatedAt = u ? (u.updated_at || '') : '';
}

function closeModal() {
  const b = document.getElementById('modalBack');
  b.style.display = 'none';
  b.dataset.expectedUpdatedAt = '';
  autoFocusCmd();
}

function openLogon() {
  openModal(null);
}

async function saveModal() {
  let uId = canonicalUnit(document.getElementById('mUnitId').value);
  if (!uId) { showConfirm('ERROR', 'UNIT REQUIRED.', () => { }); return; }

  // DC911 units are read-only — never allow editing via modal
  if (_MODAL_UNIT && _MODAL_UNIT.source && _MODAL_UNIT.source.startsWith('DC911:')) {
    showToast('DC911 UNITS ARE READ-ONLY. DATA FROM DESCHUTES 911 CAD FEED.', 'warn');
    return;
  }
  if (!_MODAL_UNIT) {
    const info = await API.getUnitInfo(TOKEN, uId);
    if (info.ok && !info.everSeen) {
      const similar = findSimilarUnits(uId);
      let confirmMsg = `"${uId}" HAS NEVER LOGGED ON BEFORE.\nCONFIRM THIS IS NOT A DUPLICATE OR TYPO?`;
      if (similar.length) confirmMsg += '\n\nSIMILAR KNOWN UNITS: ' + similar.join(', ');
      const ok = await showConfirmAsync('NEW UNIT', confirmMsg);
      if (!ok) return;
    }
  }

  let dN = (document.getElementById('mDisplayName').value || '').trim().toUpperCase();
  if (!dN) dN = displayNameForUnit(uId);

  const destEl = document.getElementById('mDestination');
  const destVal = destEl.dataset.addrId || (destEl.value || '').trim().toUpperCase();

  const newStatus = document.getElementById('mStatus').value;
  let modalNote = (document.getElementById('mNote').value || '').toUpperCase();
  // Reassemble [LOC:address] from location field
  const mLocVal = (document.getElementById('mLocation')?.value || '').trim().toUpperCase();
  if (mLocVal) {
    const resolved = AddressLookup.getById(mLocVal);
    const locAddr = resolved ? (resolved.address || resolved.name || mLocVal) : mLocVal;
    modalNote = '[LOC:' + locAddr + '] ' + modalNote;
  }
  if (newStatus === 'OOS') {
    const prevStatus = _MODAL_UNIT ? (_MODAL_UNIT.status || '') : '';
    if (prevStatus !== 'OOS') {
      const reason = await promptOOSReason(uId);
      if (!reason) return;
      if (!modalNote.startsWith('[OOS:')) modalNote = `[OOS:${reason}] ` + modalNote;
    }
  }

  const p = {
    displayName: dN,
    type: (document.getElementById('mType').value || '').trim().toUpperCase(),
    status: newStatus,
    destination: destVal,
    incident: (document.getElementById('mIncident').value || '').trim().toUpperCase(),
    note: modalNote,
    unitInfo: (() => {
      const c1n = (document.getElementById('mCrew1Name')?.value || '').trim().toUpperCase();
      const c1c = (document.getElementById('mCrew1Cert')?.value || '').trim().toUpperCase();
      const c2n = (document.getElementById('mCrew2Name')?.value || '').trim().toUpperCase();
      const c2c = (document.getElementById('mCrew2Cert')?.value || '').trim().toUpperCase();
      const notes = (document.getElementById('mUnitNotes')?.value || '').trim();
      const parts = [];
      if (c1n) parts.push('CM1:' + c1n + (c1c ? ' (' + c1c + ')' : ''));
      if (c2n) parts.push('CM2:' + c2n + (c2c ? ' (' + c2c + ')' : ''));
      if (notes) parts.push(notes);
      return parts.join('|');
    })(),
    level: (document.getElementById('mLevel')?.value || '').trim().toUpperCase(),
    station: (document.getElementById('mStation')?.value || '').trim(),
    active: true
  };

  const eUA = document.getElementById('modalBack').dataset.expectedUpdatedAt || '';
  setLive(true, 'LIVE • SAVING');
  const r = await API.upsertUnit(TOKEN, uId, p, eUA);
  if (!r.ok) {
    if (r.conflict) {
      // Unit was updated by someone else — refresh modal with latest data and let dispatcher retry
      const cur = r.current;
      const msg = 'CONFLICT: UNIT UPDATED BY ' + (cur.updated_by || 'ANOTHER USER') +
        ' (' + (cur.status || '?') + '). CLOSE AND REOPEN TO RETRY.';
      showConfirm('CONFLICT', msg, () => { closeModal(); refresh(); });
    } else {
      showErr(r);
    }
    return;
  }
  closeModal();
  refresh();
}

async function confirmLogoff() {
  const uId = canonicalUnit(document.getElementById('mUnitId').value);
  if (!uId) return;
  // Always pass '' for expectedUpdatedAt on logoff — dispatcher has override authority
  // regardless of whether field app recently changed the unit's updated_at timestamp
  const currentStatus = document.getElementById('mStatus').value;
  const currentIncident = (document.getElementById('mIncident').value || '').trim().toUpperCase();

  // Check for active incident first
  if (currentIncident) {
    const okInc = await showConfirmAsync(
      'WARNING',
      'LOG OFF ' + uId + '? UNIT IS STILL ASSIGNED TO INCIDENT ' + currentIncident + '. LOG OFF ANYWAY?'
    );
    if (!okInc) return;
  } else if (['OS', 'T', 'D', 'DE'].includes(currentStatus)) {
    const okSt = await showConfirmAsync('LOG OFF', 'LOG OFF ' + uId + '? UNIT WILL BE REMOVED FROM BOARD.');
    if (!okSt) return;
  }

  setLive(true, 'LIVE • LOGOFF');
  const r = await API.logoffUnit(TOKEN, uId, '');
  if (!r.ok) return showErr(r);
  showToast('LOGGED OFF: ' + uId);
  closeModal();
  refresh();
}

function confirmRidoff() {
  const uId = canonicalUnit(document.getElementById('mUnitId').value);
  if (!uId) return;
  const eUA = document.getElementById('modalBack').dataset.expectedUpdatedAt || '';
  showConfirm('CONFIRM RIDOFF', 'RIDOFF ' + uId + '? (SETS AV + CLEARS NOTE/INCIDENT/DEST)', async () => {
    setLive(true, 'LIVE • RIDOFF');
    const r = await API.ridoffUnit(TOKEN, uId, eUA);
    if (!r.ok) return showErr(r);
    closeModal();
    refresh();
  });
}

// ============================================================
// Type Code Flat Picker
// ============================================================
function populateTypeCodeDatalist() {
  const dl = document.getElementById('typeCodeList');
  if (!dl) return;
  const codes = (STATE && STATE.typeCodes) || [];
  dl.innerHTML = codes.filter(c => (c.category || '').toUpperCase() !== 'ADMIN')
    .map(c => `<option value="${c.code}">${c.code} — ${c.name}${c.priority ? ' · PRI-' + c.priority : ''}</option>`)
    .join('');
}

function onNewIncTypeInput(val) {
  const v = (val || '').trim().toUpperCase();
  if (!v || !(STATE && STATE.typeCodes)) return;
  // Auto-set priority when a known code is typed/selected exactly
  const tc = STATE.typeCodes.find(c => c.code === v);
  if (!tc) return;
  const priEl = document.getElementById('newIncPriority');
  if (priEl && tc.priority && !priEl.value) priEl.value = 'PRI-' + tc.priority;
  const locEl = document.getElementById('newIncLoc');
  if (locEl && !locEl.value && tc.category) {
    const cat = tc.category.toUpperCase();
    if (cat === 'CARDIAC' || cat === 'NEURO' || cat === 'RESPIRATORY' || cat === 'TRAUMA') locEl.value = 'CCT';
    else if (cat === 'MEDICAL') locEl.value = 'ALS';
    else if (cat === 'TRANSFER') locEl.value = 'BLS';
  }
  // Also sync the dropdown if it has this value
  const sel = document.getElementById('newIncTypeSelect');
  if (sel) { try { sel.value = v; } catch(_) {} }
}

function populateTypeCodeSelect(selectEl, includeAdmin) {
  if (!selectEl) return;
  const codes = (STATE && STATE.typeCodes) || [];
  const filtered = includeAdmin ? codes : codes.filter(c => (c.category || '').toUpperCase() !== 'ADMIN');
  const catOrder = ['CRITICAL', 'ALS', 'BLS', 'CCT', 'IFT', 'ADMIN'];
  const groups = {};
  filtered.forEach(c => {
    const g = (c.category || 'OTHER').toUpperCase();
    if (!groups[g]) groups[g] = [];
    groups[g].push(c);
  });
  const orderedCats = catOrder.filter(g => groups[g])
    .concat(Object.keys(groups).filter(g => !catOrder.includes(g)));
  const prevVal = selectEl.value;
  selectEl.innerHTML = '<option value="">CALL TYPE...</option>';
  orderedCats.forEach(cat => {
    const grp = document.createElement('optgroup');
    grp.label = cat;
    (groups[cat] || []).forEach(tc => {
      const opt = document.createElement('option');
      opt.value = tc.code;
      const priLabel = tc.priority ? ' · PRI-' + tc.priority : '';
      opt.textContent = tc.code + ' — ' + tc.name + priLabel;
      grp.appendChild(opt);
    });
    selectEl.appendChild(grp);
  });
  // Restore previous selection if still valid
  if (prevVal) selectEl.value = prevVal;
}

function onNewIncTypeSelect(val) {
  const typeEl = document.getElementById('newIncType');
  if (typeEl) typeEl.value = val;
  if (val && STATE && STATE.typeCodes) {
    const tc = (STATE.typeCodes || []).find(c => c.code === val);
    const priEl = document.getElementById('newIncPriority');
    if (priEl && tc && tc.priority) priEl.value = 'PRI-' + tc.priority;
    const locEl = document.getElementById('newIncLoc');
    if (locEl && tc && !locEl.value) {
      // Auto-suggest LOC from category
      const cat = (tc.category || '').toUpperCase();
      if (cat === 'CRITICAL' || cat === 'CCT') locEl.value = 'CCT';
      else if (cat === 'ALS') locEl.value = 'ALS';
      else if (cat === 'BLS') locEl.value = 'BLS';
    }
  }
  renderIncSuggest();
}

function onIncTypeSelect(val) {
  const hidEl = document.getElementById('incTypeEdit');
  if (hidEl) hidEl.value = val;
  if (val && STATE && STATE.typeCodes) {
    const tc = (STATE.typeCodes || []).find(c => c.code === val);
    const priEl = document.getElementById('incPriorityEdit');
    if (priEl && tc && tc.priority) priEl.value = 'PRI-' + tc.priority;
    const locEl = document.getElementById('incLocEdit');
    if (locEl && !locEl.value && tc) {
      const cat = (tc.category || '').toUpperCase();
      if (cat === 'CRITICAL' || cat === 'CCT') locEl.value = 'CCT';
      else if (cat === 'ALS') locEl.value = 'ALS';
      else if (cat === 'BLS') locEl.value = 'BLS';
    }
  }
}

// ============================================================
// New Incident Modal
// ============================================================
function openNewIncident(prefillScene) {
  const unitSelect = document.getElementById('newIncUnit');
  unitSelect.innerHTML = '<option value="">ASSIGN UNIT (OPTIONAL)</option>';

  const units = ((STATE && STATE.units) || []).filter(u => u.active && u.status === 'AV');
  units.forEach(u => {
    const opt = document.createElement('option');
    opt.value = u.unit_id;
    opt.textContent = u.unit_id + (u.display_name && u.display_name !== u.unit_id ? ' - ' + u.display_name : '');
    unitSelect.appendChild(opt);
  });

  const newIncSceneEl = document.getElementById('newIncScene');
  if (newIncSceneEl) newIncSceneEl.value = prefillScene ? prefillScene.toUpperCase() : '';
  const newIncPriorityEl = document.getElementById('newIncPriority');
  if (newIncPriorityEl) newIncPriorityEl.value = '';
  const newIncLocEl = document.getElementById('newIncLoc');
  if (newIncLocEl) newIncLocEl.value = '';
  document.getElementById('newIncType').value = '';
  document.getElementById('newIncNote').value = '';
  // Populate flat type code picker from STATE.typeCodes (excluding ADMIN)
  const tcSelectEl = document.getElementById('newIncTypeSelect');
  populateTypeCodeSelect(tcSelectEl, false);
  if (tcSelectEl) tcSelectEl.value = '';
  populateTypeCodeDatalist();
  // Callback + MA reset
  const cbEl = document.getElementById('newIncCallback');
  if (cbEl) cbEl.value = '';
  const maEl = document.getElementById('newIncMA');
  if (maEl) maEl.checked = false;
  // legacy urgent checkbox — may not exist in newer HTML
  const newIncUrgentEl = document.getElementById('newIncUrgent');
  if (newIncUrgentEl) newIncUrgentEl.checked = false;
  document.getElementById('newIncBack').style.display = 'flex';
  AddrHistory.attach('newIncScene', 'addrHistList1');
  renderIncSuggest();
  setTimeout(() => { if (newIncSceneEl) newIncSceneEl.focus(); }, 50);
}

function ncUseLast() {
  const last = AddrHistory.get()[0];
  if (!last) { showToast('NO ADDRESS HISTORY YET.'); return; }
  const el = document.getElementById('newIncScene');
  if (el) { el.value = last; renderIncSuggest(); }
}

function closeNewIncident() {
  document.getElementById('newIncBack').style.display = 'none';
  autoFocusCmd();
}

// ============================================================
// MAP MODAL — Leaflet + Nominatim geocoding
// ============================================================

let _mapLeafletLoaded  = false;
let _mapLeafletLoading = false;
let _mapInstance       = null;
let _mapMarker         = null;
let _mapTargetFieldId  = null;
let _mapSelectedAddr   = null;

// Central Oregon viewbox for Nominatim bias
const MAP_VIEWBOX = '-122.5,43.0,-119.5,45.2';
const MAP_DEFAULT_CENTER = [44.058, -121.315]; // Bend, OR

function _loadLeaflet(cb) {
  if (_mapLeafletLoaded) { cb(); return; }
  if (_mapLeafletLoading) { const t = setInterval(() => { if (_mapLeafletLoaded) { clearInterval(t); cb(); } }, 100); return; }
  _mapLeafletLoading = true;
  const css = document.createElement('link');
  css.rel = 'stylesheet';
  css.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
  document.head.appendChild(css);
  const js = document.createElement('script');
  js.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
  js.onload = () => { _mapLeafletLoaded = true; cb(); };
  document.head.appendChild(js);
}

function openMapModal(fieldId) {
  _mapTargetFieldId = fieldId;
  _mapSelectedAddr  = null;
  const el = document.getElementById(fieldId);
  const existing = el ? el.value.trim() : '';
  document.getElementById('mapSearchInput').value = existing;
  document.getElementById('mapResults').style.display = 'none';
  document.getElementById('mapResults').innerHTML = '';
  document.getElementById('mapSelectedAddr').textContent = 'SELECT A RESULT TO USE IT';
  document.getElementById('mapUseBtn').disabled = true;
  document.getElementById('mapUseBtn').style.opacity = '.5';
  document.getElementById('mapModalBack').style.display = 'flex';
  _loadLeaflet(() => {
    _initMapInstance();
    setTimeout(() => {
      if (_mapInstance) _mapInstance.invalidateSize();
      if (existing) _mapSearch();
    }, 150);
    document.getElementById('mapSearchInput').focus();
  });
}

function closeMapModal() {
  document.getElementById('mapModalBack').style.display = 'none';
}

function _initMapInstance() {
  if (_mapInstance) return;
  const container = document.getElementById('mapContainer');
  if (!container || !window.L) return;
  _mapInstance = L.map(container).setView(MAP_DEFAULT_CENTER, 13);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
    maxZoom: 19
  }).addTo(_mapInstance);
  _mapInstance.on('click', function(e) { _mapPinAt(e.latlng.lat, e.latlng.lng, null); });
}

async function _mapSearch() {
  const raw = document.getElementById('mapSearchInput').value.trim();
  if (!raw) return;
  // Append Oregon if no state hint
  const q = /\bOR\b|\boregon\b/i.test(raw) ? raw : raw + ', Oregon';
  const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=5&countrycodes=us&q=' +
    encodeURIComponent(q) + '&viewbox=' + MAP_VIEWBOX + '&bounded=0';
  const resultsEl = document.getElementById('mapResults');
  resultsEl.innerHTML = '<div style="padding:8px;font-size:11px;color:var(--muted);">SEARCHING...</div>';
  resultsEl.style.display = 'block';
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (!data.length) { resultsEl.innerHTML = '<div style="padding:8px;font-size:11px;color:var(--muted);">NO RESULTS.</div>'; return; }
    resultsEl.innerHTML = data.map((r, i) =>
      '<div class="map-result-row" onclick="_mapSelectResult(' + i + ')" data-lat="' + r.lat + '" data-lon="' + r.lon + '" data-addr="' + esc(r.display_name) + '">' +
        esc(r.display_name) +
      '</div>'
    ).join('');
  } catch(e) {
    resultsEl.innerHTML = '<div style="padding:8px;font-size:11px;color:#e05050;">GEOCODE FAILED. CHECK CONNECTION.</div>';
  }
}

function _mapSelectResult(idx) {
  const rows = document.querySelectorAll('.map-result-row');
  const row  = rows[idx];
  if (!row) return;
  rows.forEach(r => r.classList.remove('map-result-selected'));
  row.classList.add('map-result-selected');
  const lat  = parseFloat(row.dataset.lat);
  const lon  = parseFloat(row.dataset.lon);
  const addr = row.dataset.addr;
  _mapPinAt(lat, lon, addr);
}

function _mapPinAt(lat, lon, addr) {
  if (!_mapInstance || !window.L) return;
  if (_mapMarker) _mapInstance.removeLayer(_mapMarker);
  _mapMarker = L.marker([lat, lon]).addTo(_mapInstance);
  _mapInstance.setView([lat, lon], 16);
  // Use short address: everything before the first comma that looks like a county/state
  const short = addr ? addr.split(',').slice(0, 3).join(',').trim() : (lat.toFixed(5) + ', ' + lon.toFixed(5));
  _mapSelectedAddr = short;
  document.getElementById('mapSelectedAddr').textContent = short;
  document.getElementById('mapUseBtn').disabled = false;
  document.getElementById('mapUseBtn').style.opacity = '1';
}

function _mapUseSelected() {
  if (!_mapSelectedAddr || !_mapTargetFieldId) return;
  const el = document.getElementById(_mapTargetFieldId);
  if (el) {
    el.value = _mapSelectedAddr.toUpperCase();
    el.dispatchEvent(new Event('input'));
  }
  closeMapModal();
}

// ── Incident type parsing helper (for review modal) ──────────────────────────
function parseIncType(typeStr) {
  if (!typeStr) return { cat: '', nature: '' };
  const cats = Object.keys(INC_TYPE_TAXONOMY || {});
  for (const cat of cats) {
    if (typeStr === cat) return { cat, nature: '' };
    if (typeStr.startsWith(cat + '-')) {
      const rest = typeStr.slice(cat.length + 1);
      const natures = Object.keys((INC_TYPE_TAXONOMY[cat]?.natures || INC_TYPE_TAXONOMY[cat] || {}));
      for (const nature of natures) {
        if (rest === nature || rest.startsWith(nature + '-')) return { cat, nature };
      }
      return { cat, nature: rest };
    }
  }
  return { cat: '', nature: '' };
}

function onIncEditCatChange() {
  const cat = document.getElementById('incEditCat').value;
  const natureEl = document.getElementById('incEditNature');
  const typeEl = document.getElementById('incTypeEdit');
  if (!cat || !INC_TYPE_TAXONOMY[cat]) {
    natureEl.style.display = 'none';
    natureEl.value = '';
    typeEl.value = cat || '';
    return;
  }
  const natures = Object.keys(INC_TYPE_TAXONOMY[cat]?.natures || INC_TYPE_TAXONOMY[cat] || {});
  natureEl.innerHTML = '<option value="">—</option>' +
    natures.map(n => '<option value="' + n + '">' + n + '</option>').join('');
  natureEl.style.display = '';
  natureEl.value = '';
  typeEl.value = cat;
}

function onIncEditNatureChange() {
  const cat = document.getElementById('incEditCat').value;
  const nature = document.getElementById('incEditNature').value;
  document.getElementById('incTypeEdit').value = nature ? (cat + '-' + nature) : cat;
}

function onIncCatChange() {
  const cat = document.getElementById('newIncCat').value;
  const natureEl = document.getElementById('newIncNature');
  const detEl = document.getElementById('newIncDet');
  const typeEl = document.getElementById('newIncType');
  if (!cat || !INC_TYPE_TAXONOMY[cat]) {
    natureEl.style.display = 'none';
    detEl.style.display = 'none';
    typeEl.value = cat || '';
    return;
  }
  const _natSrc = INC_TYPE_TAXONOMY[cat]?.natures || INC_TYPE_TAXONOMY[cat] || {};
  // Filter out legacy types from dispatch picker (still visible in edit modal)
  const natures = Object.keys(_natSrc).filter(n => !(_natSrc[n]?.legacy));
  natureEl.innerHTML = '<option value="">NATURE...</option>' +
    natures.map(n => '<option value="' + n + '">' + n + '</option>').join('');
  natureEl.style.display = '';
  natureEl.value = '';
  detEl.style.display = 'none';
  detEl.value = '';
  typeEl.value = cat;
  renderIncSuggest();
}

function onIncNatureChange() {
  const cat = document.getElementById('newIncCat').value;
  const nature = document.getElementById('newIncNature').value;
  const detEl = document.getElementById('newIncDet');
  const typeEl = document.getElementById('newIncType');
  if (!nature) {
    detEl.style.display = 'none';
    typeEl.value = cat;
    return;
  }
  const _natMap = INC_TYPE_TAXONOMY[cat]?.natures || INC_TYPE_TAXONOMY[cat] || {};
  const _natVal = _natMap[nature];
  const dets = (_natVal?.dets || (Array.isArray(_natVal) ? _natVal : []));
  if (dets.length) {
    detEl.innerHTML = '<option value="">DET...</option>' +
      dets.map(d => '<option value="' + d + '">' + d + '</option>').join('');
    detEl.style.display = '';
    detEl.value = '';
  } else {
    detEl.style.display = 'none';
  }
  typeEl.value = cat + '-' + nature;
  renderIncSuggest();
}

function onIncDetChange() {
  const cat = document.getElementById('newIncCat').value;
  const nature = document.getElementById('newIncNature').value;
  const det = document.getElementById('newIncDet').value;
  const typeEl = document.getElementById('newIncType');
  typeEl.value = det ? (cat + '-' + nature + '-' + det) : (cat + '-' + nature);
  // Auto-set priority if determinant is a PRI-n value
  const priMatch = det.match(/^PRI-(\d)$/);
  const priEl = document.getElementById('newIncPriority');
  if (priEl && priMatch) priEl.value = 'PRI-' + priMatch[1];
  // Direct priority determinants (new transport taxonomy)
  if (priEl) {
    if (det === 'PRI-1') priEl.value = 'PRI-1';
    else if (det === 'PRI-2') priEl.value = 'PRI-2';
    else if (det === 'PRI-3') priEl.value = 'PRI-3';
    else if (det === 'PRI-4') priEl.value = 'PRI-4';
  }
  renderIncSuggest();
}

async function createNewIncident() {
  let note = document.getElementById('newIncNote').value.trim().toUpperCase();
  const priority = (document.getElementById('newIncPriority')?.value || '').trim().toUpperCase();
  const unitId = document.getElementById('newIncUnit').value;
  const incType = (document.getElementById('newIncType').value || '').trim().toUpperCase();
  const sceneAddress = (document.getElementById('newIncScene')?.value || '').trim().toUpperCase();
  const levelOfCare = (document.getElementById('newIncLoc')?.value || '').trim().toUpperCase();
  const callback = (document.getElementById('newIncCallback')?.value || '').trim();
  const mutualAid = document.getElementById('newIncMA')?.checked || false;

  if (!sceneAddress) {
    showAlert('ERROR', 'SCENE ADDRESS REQUIRED.');
    return;
  }

  // Duplicate dispatch guard — warn if same scene+type created in last 30 seconds
  const thirtySecAgo = Date.now() - 30000;
  const dupe = (STATE && STATE.incidents || []).find(inc => {
    if (inc.status === 'CLOSED') return false;
    const created = inc.created_at ? new Date(inc.created_at).getTime() : 0;
    if (created < thirtySecAgo) return false;
    const sameScene = (inc.scene_address || inc.destination || '').toUpperCase() === sceneAddress.toUpperCase();
    const sameType = incType && (inc.incident_type || '').toUpperCase() === incType.toUpperCase();
    return sameScene && (!incType || sameType);
  });
  if (dupe) {
    const ok = await showConfirmAsync('POSSIBLE DUPLICATE',
      dupe.incident_id + ' (' + (dupe.incident_type || 'UNKNOWN') + ') AT ' +
      (dupe.scene_address || dupe.destination || 'SAME LOCATION') +
      ' WAS CREATED ' + Math.round((Date.now() - new Date(dupe.created_at).getTime()) / 1000) +
      'S AGO.\n\nCREATE ANYWAY?');
    if (!ok) return;
  }

  // Prepend prefixes (MA first, then CB)
  const prefixes = [];
  if (mutualAid) prefixes.push('[MA]');
  if (callback) prefixes.push('[CB:' + callback + ']');
  if (prefixes.length) note = prefixes.join(' ') + (note ? ' ' + note : '');

  setLive(true, 'LIVE • CREATE INCIDENT');
  const r = await API.createQueuedIncident(TOKEN, '', note, priority, unitId, incType, sceneAddress, levelOfCare);
  if (!r.ok) return showErr(r);
  if (sceneAddress) { AddrHistory.push(sceneAddress); _geoVerifyAddress(sceneAddress); }
  closeNewIncident();
  refresh();
}

async function closeIncidentFromQueue(incidentId) {
  const disposition = await promptDisposition(incidentId);
  if (!disposition) return; // user cancelled
  setLive(true, 'LIVE • CLOSE INCIDENT');
  try {
    const r = await API.closeIncident(TOKEN, incidentId, disposition);
    if (!r.ok) { showAlert('ERROR', r.error || 'FAILED TO CLOSE INCIDENT'); return; }
    showToast(incidentId + ' CLOSED — ' + disposition);
    refresh();
  } catch (e) {
    showAlert('ERROR', 'FAILED: ' + e.message);
  }
}

function assignIncidentToUnit(incidentId) {
  const shortId = incidentId.replace(/^\d{2}-/, '');

  // Score available units for this incident (same logic as suggestUnits)
  const inc = (STATE.incidents || []).find(i => i.incident_id === incidentId);
  const incType = inc ? (inc.incident_type || '').toUpperCase() : '';
  const pri = inc ? (inc.priority || '').toUpperCase() : '';
  const assigned = (STATE.assignments || []).filter(a => a.incident_id === incidentId && !a.cleared_at).map(a => a.unit_id);
  const needsALS = /^CCT|^IFT-ALS/.test(incType) || pri === 'PRI-1';
  const blsOk = /^IFT-BLS|^DISCHARGE|^DIALYSIS/.test(incType) || pri === 'PRI-3' || pri === 'PRI-4';
  const quickPicks = (STATE && STATE.units || [])
    .filter(u => u.active && (u.status === 'AV' || u.status === 'BRK') && u.include_in_recommendations !== false && !assigned.includes(u.unit_id))
    .map(u => {
      const lvl = (u.level || '').toUpperCase();
      let score = u.status === 'AV' ? 100 : 90;
      if (needsALS) { score += lvl === 'ALS' ? 60 : lvl === 'AEMT' ? 30 : 5; }
      else if (blsOk) { score += (lvl === 'BLS' || lvl === 'EMT') ? 40 : lvl === 'AEMT' ? 35 : 20; }
      else { score += lvl === 'ALS' ? 30 : lvl === 'AEMT' ? 20 : 10; }
      return { u, score };
    })
    .sort((a, b) => b.score - a.score)
    .slice(0, 4)
    .map(({ u }) => u);

  const input = document.createElement('input');
  input.type = 'text';
  input.id = 'assignUnitInput';
  input.placeholder = 'UNIT ID (E.G. EMS1, WC1)';
  input.style.cssText = 'width:100%;padding:10px;background:var(--panel);color:var(--text);border:2px solid var(--line);font-family:inherit;text-transform:uppercase;font-size:14px;margin-top:6px;';

  const message = document.createElement('div');
  let msgHtml = '<div style="margin-bottom:8px;">ASSIGN ' + esc(incidentId) + ' TO UNIT:</div>';
  if (quickPicks.length) {
    msgHtml += '<div style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:8px;">';
    quickPicks.forEach(u => {
      const lvlBadge = u.level ? ' <span style="font-size:9px;opacity:.7;">' + esc(u.level) + '</span>' : '';
      const brkBadge = u.status === 'BRK' ? ' <span style="font-size:9px;color:#ffd66b;">BRK</span>' : '';
      msgHtml += '<button type="button" data-unit-pick="' + esc(u.unit_id) + '" style="padding:7px 12px;background:#1a3a5c;border:1px solid #2f6cff;color:#e6edf3;font-family:inherit;font-size:12px;font-weight:900;cursor:pointer;letter-spacing:.05em;">' + esc(u.unit_id) + lvlBadge + brkBadge + '</button>';
    });
    msgHtml += '</div><div style="font-size:10px;color:#8b949e;margin-bottom:2px;">OR TYPE UNIT ID:</div>';
  }
  message.innerHTML = msgHtml;
  message.appendChild(input);

  document.getElementById('alertTitle').textContent = 'ASSIGN INCIDENT';
  document.getElementById('alertMessage').innerHTML = '';
  document.getElementById('alertMessage').appendChild(message);
  document.getElementById('alertDialog').classList.add('active');

  setTimeout(() => input.focus(), 100);

  const handleAssign = () => {
    const unitInput = input.value.trim();
    if (!unitInput) {
      hideAlert();
      return;
    }
    const unitId = canonicalUnit(unitInput);
    if (!unitId) {
      showAlert('ERROR', 'INVALID UNIT ID');
      return;
    }
    hideAlert();
    const cmd = `D ${unitId} ${incidentId}`;
    document.getElementById('cmd').value = cmd;
    runCommand();
  };

  // Quick-pick chip click → fill input and assign immediately
  message.addEventListener('click', e => {
    const chip = e.target.closest('[data-unit-pick]');
    if (chip) {
      input.value = chip.getAttribute('data-unit-pick');
      handleAssign();
    }
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      handleAssign();
    } else if (e.key === 'Escape') {
      hideAlert();
    }
  });
}

// ============================================================
// Incident Review Modal
// ============================================================
async function openIncidentFromServer(iId) {
  setLive(true, 'LIVE • INCIDENT REVIEW');
  const r = await API.getIncident(TOKEN, iId);
  if (!r.ok) return showErr(r);

  const inc = r.incident;
  CURRENT_INCIDENT_ID = String(inc.incident_id || '').toUpperCase();
  document.getElementById('incTitle').firstChild.textContent = 'INCIDENT ' + CURRENT_INCIDENT_ID;

  // Status badge — color-coded: ACTIVE=green, QUEUED=yellow, CLOSED=red
  const statusBadgeEl = document.getElementById('incStatusBadge');
  if (statusBadgeEl) {
    const incStatus = (inc.status || '').toUpperCase();
    const statusStyles = {
      'ACTIVE': 'background:rgba(63,185,80,.18);border:1px solid rgba(63,185,80,.5);color:#3fb950;',
      'QUEUED': 'background:rgba(210,153,34,.18);border:1px solid rgba(210,153,34,.5);color:#d2a424;',
      'CLOSED': 'background:rgba(212,48,48,.18);border:1px solid rgba(212,48,48,.5);color:#f85149;',
    };
    statusBadgeEl.textContent = incStatus || 'UNKNOWN';
    statusBadgeEl.style.cssText = (statusStyles[incStatus] || 'background:rgba(139,148,158,.18);border:1px solid rgba(139,148,158,.4);color:#8b949e;') + 'display:inline-block;font-size:12px;font-weight:900;padding:3px 10px;border-radius:3px;letter-spacing:.08em;';
  }
  document.getElementById('incUnits').textContent = (inc.units || '—').toUpperCase();

  // Show live unit status for all active units currently on this incident
  const incUnitsDetailEl = document.getElementById('incUnitsDetail');
  if (incUnitsDetailEl && STATE && STATE.units) {
    const liveUnits = (STATE.units || []).filter(u => u.active && String(u.incident || '').toUpperCase() === String(iId).toUpperCase());
    if (liveUnits.length > 0) {
      incUnitsDetailEl.innerHTML = liveUnits.map(u => {
        const stCode = u.status || '?';
        const stBadge = '<span class="S-' + esc(stCode) + '" style="display:inline-block;min-width:28px;text-align:center;font-size:10px;font-weight:900;padding:1px 4px;margin-right:4px;">' + esc(stCode) + '</span>';
        const rawNote = (u.note || '').replace(/\[(ETA|PAT|LOC|OOS|PP|ACK):[^\]]*\]\s*/gi, '').trim();
        const etaM = ((u.note || '').match(/\[ETA:(\d+)\]/) || [])[1];
        const patM = ((u.note || '').match(/\[PAT:([^\]]+)\]/) || [])[1];
        const extras = [etaM ? 'ETA:' + etaM + 'M' : '', patM ? 'PAT:' + patM : ''].filter(Boolean).join(' ');
        const noteDisplay = [rawNote, extras].filter(Boolean).join(' | ');
        return '<div style="padding:2px 0;border-bottom:1px solid rgba(255,255,255,.05);">' +
          stBadge +
          '<span style="color:#8b949e;min-width:64px;display:inline-block;">' + esc(u.unit_id || '') + '</span>' +
          (noteDisplay ? '<span style="color:#c9d1d9;"> — ' + esc(noteDisplay) + '</span>' : '') +
          '</div>';
      }).join('');
      incUnitsDetailEl.style.display = 'block';
    } else {
      incUnitsDetailEl.style.display = 'none';
    }
  } else if (incUnitsDetailEl) {
    incUnitsDetailEl.style.display = 'none';
  }

  const incDestR = AddressLookup.resolve(inc.destination);
  const incDestEl = document.getElementById('incDestEdit');
  incDestEl.value = (incDestR.recognized ? incDestR.addr.name : (inc.destination || '')).toUpperCase();
  if (incDestR.recognized) incDestEl.dataset.addrId = incDestR.addr.id;
  else delete incDestEl.dataset.addrId;
  // Populate flat type code picker for edit modal
  const incTypeRaw = (inc.incident_type || '').toUpperCase();
  const incTypeSelectEl = document.getElementById('incTypeSelect');
  populateTypeCodeSelect(incTypeSelectEl, true); // include ADMIN for edit modal
  if (incTypeSelectEl) {
    // Try to match current type code in the flat list; fall back to showing raw type
    const matchCode = (STATE.typeCodes || []).find(c => c.code === incTypeRaw);
    incTypeSelectEl.value = matchCode ? incTypeRaw : '';
  }
  document.getElementById('incTypeEdit').value = incTypeRaw;
  const priEditEl = document.getElementById('incPriorityEdit');
  if (priEditEl) priEditEl.value = (inc.priority || '').toUpperCase();
  const locEditEl = document.getElementById('incLocEdit');
  if (locEditEl) locEditEl.value = (inc.level_of_care || '').toUpperCase();
  document.getElementById('incUpdated').textContent = inc.last_update ? fmtTime24(inc.last_update) : '—';

  const bC = getRoleColor(inc.updated_by);
  const bE = document.getElementById('incBy');
  bE.textContent = (inc.updated_by || '—').toUpperCase();
  bE.className = bC;

  // Strip [DISP:] and [CB:] tags from note textarea; show each as a badge
  const rawIncNote = (inc.incident_note || '').toUpperCase();
  const dispTagMatch = rawIncNote.match(/\[DISP:([^\]]+)\]/i);
  const cbTagMatch = rawIncNote.match(/\[CB:([^\]]+)\]/i);
  document.getElementById('incNote').value = rawIncNote.replace(/\[DISP:[^\]]*\]\s*/gi, '').replace(/\[CB:[^\]]*\]\s*/gi, '').trim();
  const dispBadgeEl = document.getElementById('incDispositionBadge');
  if (dispBadgeEl) {
    const dispCode = dispTagMatch ? dispTagMatch[1].toUpperCase() : (inc.disposition || '').toUpperCase();
    const incStatus = (inc.status || '').toUpperCase();
    if (incStatus === 'CLOSED') {
      // Prominent closed banner: show status + disposition together
      let closedText = 'CLOSED';
      if (inc.closed_at) closedText += ' · ' + fmtTime24(inc.closed_at);
      if (dispCode) closedText += ' · DISPO: ' + dispCode;
      dispBadgeEl.innerHTML = '<span style="color:#f85149;font-weight:900;">■ ' + esc(closedText) + '</span>';
      dispBadgeEl.style.cssText = 'display:block;margin:4px 0 2px;font-size:11px;font-weight:900;letter-spacing:.06em;font-family:monospace;padding:4px 8px;background:rgba(212,48,48,.12);border-left:3px solid #f85149;';
    } else if (dispCode) {
      dispBadgeEl.textContent = 'DISPOSITION: ' + dispCode;
      dispBadgeEl.style.cssText = 'display:block;margin:4px 0 2px;font-size:11px;font-weight:900;letter-spacing:.06em;color:#79c0ff;font-family:monospace;';
    } else {
      dispBadgeEl.style.display = 'none';
    }
  }
  // CB badge — click to copy phone number
  const cbPillEl = document.getElementById('incCbPill');
  const cbNumEl = document.getElementById('incCbNumber');
  if (cbPillEl && cbNumEl) {
    if (cbTagMatch) {
      cbNumEl.textContent = cbTagMatch[1].toUpperCase();
      cbPillEl.style.display = '';
    } else {
      cbPillEl.style.display = 'none';
    }
  }

  const relEl = document.getElementById('incRelated');
  if (relEl) {
    const parts = [];
    const relIds = inc.related_incidents || [];
    if (relIds.length) {
      parts.push('<span class="muted" style="font-size:10px;">RELATED: </span>' +
        relIds.map(function(id) {
          const shortRel = id.replace(/^\d{2}-/, '');
          return '<button type="button" class="related-inc-chip" onclick="openIncident(\'' + esc(id) + '\')">' + esc(id) + '</button>';
        }).join(''));
    }
    if (parts.length) {
      relEl.innerHTML = parts.join('');
      relEl.style.display = '';
    } else {
      relEl.innerHTML = '';
      relEl.style.display = 'none';
    }
  }

  const incSceneEl = document.getElementById('incSceneAddress');
  if (incSceneEl) incSceneEl.value = (inc.scene_address || '').toUpperCase();
  AddrHistory.attach('incSceneAddress', 'addrHistList2');

  // Timing row — show all 6 EMS timestamps with KPI-colored elapsed deltas
  const tr2 = document.getElementById('incTimingRow');
  if (tr2) {
    const _minsApartNum = (a, b) => {
      if (!a || !b) return null;
      const d = Math.round((new Date(b) - new Date(a)) / 60000);
      return d > 0 ? d : null;
    };
    const stages = [
      { key: 'dispatch_time',    label: 'DISP',    kpi: null },
      { key: 'enroute_time',     label: 'ENRTE',   kpi: KPI_TARGETS['D→DE']  },
      { key: 'arrival_time',     label: 'ARR',     kpi: KPI_TARGETS['DE→OS'] },
      { key: 'transport_time',   label: 'TRANS',   kpi: KPI_TARGETS['OS→T']  },
      { key: 'at_hospital_time', label: 'AT HOSP', kpi: null },
      { key: 'handoff_time',     label: 'HOFF',    kpi: KPI_TARGETS['T→AV']  },
    ];
    const parts = [];
    let prevTime = null;
    stages.forEach(s => {
      const t = inc[s.key];
      if (!t) { prevTime = null; return; }
      const d = _minsApartNum(prevTime, t);
      let elapsedHtml = '';
      if (d !== null) {
        let col = '';
        if (s.kpi) {
          col = d <= s.kpi ? '#3fb950' : d <= s.kpi * 1.5 ? '#e0b040' : '#f85149';
        } else {
          col = '#8b949e';
        }
        elapsedHtml = ' <span style="color:' + col + ';font-weight:900;">(+' + d + 'M)</span>';
      }
      parts.push('<span>' + s.label + ': ' + fmtTime24(t) + elapsedHtml + '</span>');
      prevTime = t;
    });
    tr2.innerHTML = parts.join('<span style="color:#30363d;"> | </span>');
    tr2.style.display = parts.length ? '' : 'none';
  }

  renderIncidentAudit(r.audit || []);
  document.getElementById('incBack').style.display = 'flex';
  setTimeout(() => document.getElementById('incNote').focus(), 50);
}

function openIncident(iId) {
  openIncidentFromServer(iId);
}

// Agency-level approximate coordinates for proximity scoring in SUGGEST
// Used when exact geocode is unavailable — coarse but correct directionally
const _SUGGEST_AGENCY_COORDS = {
  'LAPINE_FD':            [43.6679, -121.5036],
  'SUNRIVER_FD':          [43.8758, -121.4363],
  'BEND_FIRE':            [44.0489, -121.3153],
  'REDMOND_FIRE':         [44.2783, -121.1785],
  'CROOK_COUNTY_FIRE':    [44.3050, -120.7271],
  'SISTERS_CAMP_SHERMAN': [44.2879, -121.5490],
  'AIRLINK_CCT':          [44.0613, -121.2832],
  'JEFFCO_FIRE_EMS':      [44.6289, -121.1278],
  'WARM_SPRINGS_FD':      [44.7640, -121.2660],
  'BLACK_BUTTE_RANCH':    [44.3988, -121.6282],
  'ALFALFA_FD':           [44.0084, -121.1091],
  'CRESCENT_RFPD':        [43.4644, -121.7168],
  'THREE_RIVERS_FD':      [44.3524, -121.1781],
  'SCMC':                 [44.0613, -121.2832],
};

// Central Oregon bounding box for geo-verification warnings
const _GEO_BBOX = { n: 44.72, s: 43.49, e: -120.68, w: -122.07 };

// Fire-and-forget: geocode addr and warn if outside Central Oregon
// Never blocks the caller — always resolves silently on error
function _geoVerifyAddress(addr) {
  if (!addr || addr.length < 6) return;
  const a = addr.trim().toUpperCase();
  // Check cache first — instant
  if (_bmGeoCache && _bmGeoCache[a]) {
    const c = _bmGeoCache[a];
    const lat = c.lat, lon = c.lon;
    if (lat < _GEO_BBOX.s || lat > _GEO_BBOX.n || lon < _GEO_BBOX.w || lon > _GEO_BBOX.e) {
      showToast('⚠ SCENE ADDRESS MAY BE OUTSIDE CENTRAL OREGON — VERIFY CORRECT LOCATION', 'error', 8000);
    }
    return;
  }
  // Not in cache — async Nominatim lookup (non-blocking)
  const q = /\bOR\b|\boregon\b/i.test(a) ? a : a + ', Oregon';
  const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=us&q=' + encodeURIComponent(q);
  fetch(url).then(function(res) { return res.json(); }).then(function(data) {
    if (!data || !data[0]) return;
    const lat = parseFloat(data[0].lat), lon = parseFloat(data[0].lon);
    if (isNaN(lat) || isNaN(lon)) return;
    if (lat < _GEO_BBOX.s || lat > _GEO_BBOX.n || lon < _GEO_BBOX.w || lon > _GEO_BBOX.e) {
      showToast('⚠ SCENE ADDRESS MAY BE OUTSIDE CENTRAL OREGON — VERIFY CORRECT LOCATION', 'error', 8000);
    }
  }).catch(function() {});
}

function _haversineDist(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function suggestUnits(incId) {
  let iId = (incId || CURRENT_INCIDENT_ID || '').trim().toUpperCase();
  if (!iId) { showAlert('ERROR', 'NO INCIDENT OPEN. USE: SUGGEST INC0001'); return; }
  // Normalize: INC0014 → 26-0014 (strip INC prefix, add year if no dash present)
  iId = iId.replace(/^INC-?/i, '');
  if (!iId.includes('-')) iId = (new Date().getFullYear() % 100) + '-' + iId;

  const inc = (STATE.incidents || []).find(i => i.incident_id === iId);
  if (!inc) { showAlert('NOT FOUND', 'INCIDENT ' + iId + ' NOT IN CURRENT STATE. REFRESH AND TRY AGAIN.'); return; }

  const incType  = (inc.incident_type || '').toUpperCase();
  const pri      = (inc.priority || '').toUpperCase();
  // Exclude units already assigned (primary or queued) to this incident
  const assigned = (STATE.assignments || [])
    .filter(a => a.incident_id === iId && !a.cleared_at)
    .map(a => a.unit_id);

  const available = (STATE.units || []).filter(u =>
    u.active &&
    (u.status === 'AV' || u.status === 'BRK') &&
    u.include_in_recommendations !== false &&
    !assigned.includes(u.unit_id)
  );

  // Try proximity scoring using geocache
  const sceneAddr = (inc.scene_address || '').trim().toUpperCase();
  const geoHit    = sceneAddr && _bmGeoCache ? _bmGeoCache[sceneAddr] : null;
  const incLatLon = geoHit ? [geoHit.lat, geoHit.lon] : null;

  const needsALS  = /^CCT|^IFT-ALS/.test(incType) || pri === 'PRI-1';
  const preferALS = pri === 'PRI-2';
  const blsOk     = /^IFT-BLS|^DISCHARGE|^DIALYSIS/.test(incType) || pri === 'PRI-3' || pri === 'PRI-4';

  const scored = available.map(u => {
    const level  = (u.level || '').toUpperCase();
    const agency = (u.agency_id || '').toUpperCase();
    let score = 100;

    // BRK penalty — available but less ideal than AV
    if (u.status === 'BRK') score -= 10;

    // ALS/level scoring
    if (needsALS) {
      if (level === 'ALS')                    score += 60;
      else if (level === 'AEMT')              score += 30;
      else if (level === 'BLS' || level === 'EMT') score += 5;
    } else if (preferALS) {
      if (level === 'ALS')                    score += 40;
      else if (level === 'AEMT')              score += 25;
      else if (level === 'BLS' || level === 'EMT') score += 15;
    } else if (blsOk) {
      if (level === 'BLS' || level === 'EMT') score += 40;
      else if (level === 'AEMT')              score += 35;
      else if (level === 'ALS')               score += 20;
    } else {
      if (level === 'ALS')                    score += 30;
      else if (level === 'AEMT')              score += 20;
      else if (level === 'BLS' || level === 'EMT') score += 10;
    }

    // Station proximity bonus (if geocache hit)
    if (incLatLon) {
      const agencyCoords = _SUGGEST_AGENCY_COORDS[agency];
      if (agencyCoords) {
        const dist = _haversineDist(incLatLon[0], incLatLon[1], agencyCoords[0], agencyCoords[1]);
        if      (dist < 5)  score += 30;
        else if (dist < 15) score += 20;
        else if (dist < 30) score += 10;
        else if (dist > 60) score -= 10;
      }
    }

    return { unit: u, score, agency };
  });

  scored.sort((a, b) => b.score - a.score);
  const recs = scored.slice(0, 5);

  // Render inline chips in the incident panel instead of a modal
  const resultEl = document.getElementById('incSuggestResult');
  if (resultEl) {
    if (!recs.length) {
      resultEl.innerHTML =
        '<div class="muted" style="font-size:11px;padding:4px 0;">NO AVAILABLE UNITS TO SUGGEST.</div>';
      resultEl.style.display = '';
      return;
    }
    const chips = recs.map(function(s) {
      const u = s.unit;
      const levelBadge = u.level ? ' <span style="font-size:9px;opacity:.7;">' + esc(u.level) + '</span>' : '';
      const brkBadge   = u.status === 'BRK' ? ' <span style="font-size:9px;color:#ffd66b;">BRK</span>' : '';
      return '<button type="button" class="suggest-chip" style="margin:2px 3px;" ' +
        'title="Click to fill ASSIGN command" ' +
        'onclick="fillAssignCmd(\'' + esc(u.unit_id) + '\',\'' + esc(iId) + '\')">' +
        esc(u.unit_id) + levelBadge + brkBadge +
        '</button>';
    }).join('');
    const proximity = incLatLon ? '' : ' <span class="muted" style="font-size:9px;">(no geo)</span>';
    resultEl.innerHTML =
      '<div class="inc-suggest-row">' +
      '<span class="muted" style="font-size:11px;white-space:nowrap;">SUGGESTED' + proximity + ':</span>' +
      chips +
      '<button type="button" onclick="document.getElementById(\'incSuggestResult\').style.display=\'none\'" ' +
      'style="font-size:10px;padding:1px 6px;background:transparent;border:1px solid #30363d;color:#8b949e;cursor:pointer;font-family:inherit;margin-left:4px;">✕</button>' +
      '</div>';
    resultEl.style.display = '';
  } else {
    // Fallback if panel not open
    let msg = 'TYPE: ' + (inc.incident_type || '—');
    if (inc.priority) msg += '  |  PRIORITY: ' + inc.priority;
    if (assigned.length) msg += '\nALREADY ASSIGNED: ' + assigned.join(', ');
    msg += '\n\nSUGGESTED:\n';
    recs.forEach((s, i) => {
      const u = s.unit;
      msg += (i + 1) + '. ' + u.unit_id + ' [' + u.status + ']';
      if (u.level) msg += ' ' + u.level;
      msg += '\n';
    });
    msg += '\nUSE: ASSIGN <UNIT>; ' + iId;
    showAlert('SUGGESTIONS — ' + iId, msg);
  }
}

function fillAssignCmd(unitId, incId) {
  const inp = document.getElementById('cmd');
  if (inp) {
    inp.value = 'ASSIGN ' + incId + ' ' + unitId;
    inp.focus();
    inp.select();
  }
}

function closeIncidentPanel() {
  document.getElementById('incBack').style.display = 'none';
  CURRENT_INCIDENT_ID = '';
}

// Keep old name as alias for ESC key handler etc.
function closeIncident() { closeIncidentPanel(); }

async function alertAllIncident() {
  const incId = CURRENT_INCIDENT_ID;
  if (!incId) { showAlert('ERROR', 'NO INCIDENT OPEN'); return; }
  const inc = (STATE && STATE.incidents || []).find(i => i.incident_id === incId);
  const parts = [incId];
  if (inc) {
    if (inc.priority) parts.push('[' + inc.priority + ']');
    if (inc.incident_type) parts.push(inc.incident_type);
    if (inc.destination) parts.push('DEST: ' + inc.destination);
    if (inc.scene_address) parts.push('SCENE: ' + inc.scene_address);
    if (inc.incident_note) parts.push(inc.incident_note.replace(/\[[^\]]*\]/g,'').replace(/\s{2,}/g,' ').trim());
  }
  const msg = 'CRITICAL INCIDENT ALERT — ' + parts.join(' | ');
  const ok = await showConfirmAsync('ALERT ALL?', 'Send hot message to ALL dispatchers and ALL field units:\n\n' + msg);
  if (!ok) return;
  setLive(true, 'LIVE • ALERT ALL');
  const [r1, r2] = await Promise.all([
    API.sendToDispatchers(TOKEN, msg, true),
    API.sendToUnits(TOKEN, msg, true)
  ]);
  setLive(false);
  if (!r1.ok && !r2.ok) return showErr(r1);
  const dp = (r1.ok ? r1.recipients : 0);
  const un = (r2.ok ? r2.recipients : 0);
  showToast('CRITICAL ALERT SENT — ' + dp + ' DISPATCHER(S), ' + un + ' UNIT(S)');
  refresh();
}

async function closeIncidentAction() {
  const incId = CURRENT_INCIDENT_ID;
  if (!incId) { showAlert('ERROR', 'NO INCIDENT OPEN'); return; }
  const disposition = await promptDisposition(incId);
  if (!disposition) return; // user cancelled
  setLive(true, 'LIVE • CLOSE INCIDENT');
  try {
    const r = await API.closeIncident(TOKEN, incId, disposition);
    if (!r.ok) { showAlert('ERROR', r.error || 'FAILED TO CLOSE INCIDENT'); return; }
    closeIncidentPanel();
    showToast('INCIDENT ' + String(incId).replace(/^[A-Z]*\d{2}-0*/, '') + ' CLOSED — ' + disposition);
    refresh();
  } catch (e) {
    showAlert('ERROR', 'FAILED TO CLOSE INCIDENT: ' + e.message);
  }
}

async function reopenIncidentAction() {
  const incId = CURRENT_INCIDENT_ID;
  if (!incId) { showAlert('ERROR', 'NO INCIDENT OPEN'); return; }
  const ok = await showConfirmAsync('REOPEN INCIDENT ' + incId + '?');
  if (!ok) return;
  setLive(true, 'LIVE • REOPEN INCIDENT');
  try {
    const r = await API.reopenIncident(TOKEN, incId);
    if (!r.ok) { showAlert('ERROR', r.error || 'FAILED TO REOPEN INCIDENT'); return; }
    closeIncidentPanel();
    showToast('INCIDENT ' + incId + ' REOPENED.');
    refresh();
  } catch (e) {
    showAlert('ERROR', 'FAILED TO REOPEN INCIDENT: ' + e.message);
  }
}

async function saveIncidentNote() {
  const m = (document.getElementById('incNote').value || '').trim().toUpperCase();
  const newType = (document.getElementById('incTypeEdit').value || '').trim().toUpperCase();
  const destEl = document.getElementById('incDestEdit');
  const newDest = destEl.dataset.addrId || (destEl.value || '').trim().toUpperCase();
  const newScene = (document.getElementById('incSceneAddress')?.value || '').trim().toUpperCase() || undefined;
  const newPriority = (document.getElementById('incPriorityEdit')?.value || '').trim().toUpperCase();
  const newLoc = (document.getElementById('incLocEdit')?.value || '').trim().toUpperCase();
  if (!CURRENT_INCIDENT_ID) return;

  // Get current incident to compare all fields
  const curInc = (STATE.incidents || []).find(i => i.incident_id === CURRENT_INCIDENT_ID);
  const curDest = curInc ? (curInc.destination || '') : '';
  const curScene = curInc ? (curInc.scene_address || '') : '';
  const curPriority = curInc ? (curInc.priority || '') : '';
  const curLoc = curInc ? (curInc.level_of_care || '') : '';
  const curType = curInc ? (curInc.incident_type || '') : '';
  const destChanged = newDest !== curDest.toUpperCase();
  const sceneChanged = newScene !== undefined && newScene !== curScene.toUpperCase();
  const priorityChanged = newPriority !== curPriority.toUpperCase();
  const locChanged = newLoc !== curLoc.toUpperCase();
  // Only flag typeChanged if type actually differs from current — avoids spurious [TYPE:] audit entries
  const typeChanged = newType && newType !== curType.toUpperCase();

  // Preserve [DISP:] and [CB:] tags when saving note — re-prepend from current incident
  const curDispMatch = ((curInc && curInc.incident_note) || '').match(/\[DISP:([^\]]+)\]/i);
  const curCbMatch   = ((curInc && curInc.incident_note) || '').match(/\[CB:([^\]]+)\]/i);
  let mWithDisp = curDispMatch ? ('[DISP:' + curDispMatch[1].toUpperCase() + '] ' + m).trim() : m;
  if (curCbMatch) mWithDisp = (mWithDisp + ' [CB:' + curCbMatch[1].toUpperCase() + ']').trim();

  // If anything changed, use updateIncident
  if (typeChanged || m || destChanged || sceneChanged || priorityChanged || locChanged) {
    setLive(true, 'LIVE • UPDATE INCIDENT');
    const r = await API.updateIncident(TOKEN, CURRENT_INCIDENT_ID, mWithDisp, typeChanged ? newType : '', destChanged ? newDest : undefined, sceneChanged ? newScene : undefined, priorityChanged ? newPriority : undefined, locChanged ? newLoc : undefined);
    if (!r.ok) return showErr(r);
    if (sceneChanged && newScene) { AddrHistory.push(newScene); _geoVerifyAddress(newScene); }
    closeIncidentPanel();
    refresh();
    return;
  }

  showConfirm('ERROR', 'NO CHANGES DETECTED. UPDATE TYPE, NOTE, DESTINATION, SCENE, PRIORITY, OR LEVEL OF CARE.', () => { });
}

function renderIncidentAudit(aR) {
  const e = document.getElementById('incAudit');
  const rs = aR || [];
  if (!rs.length) {
    e.innerHTML = '<div class="muted">NO HISTORY.</div>';
    return;
  }
  e.innerHTML = rs.map(r => {
    const ts = r.ts ? fmtTime24(r.ts) : '—';
    const aC = getRoleColor(r.actor);
    return `<div style="border-bottom:1px solid var(--line); padding:8px 6px;">
      <div class="muted ${aC}">${esc(ts)} • ${esc((r.actor || '').toUpperCase())}</div>
      <div style="font-weight:900; color:var(--yellow); margin-top:2px;">${esc(String(r.message || ''))}</div>
    </div>`;
  }).join('');
  // Auto-scroll to bottom — most recent entry is last
  setTimeout(() => { e.scrollTop = e.scrollHeight; }, 0);
}

// ============================================================
// Unit History Modal
// ============================================================
function closeUH() {
  document.getElementById('uhBack').style.display = 'none';
  UH_CURRENT_UNIT = '';
}

function reloadUH() {
  if (!UH_CURRENT_UNIT) return;
  const h = Number(document.getElementById('uhHours').value || 12);
  openHistory(UH_CURRENT_UNIT, h);
}

async function openHistory(uId, h) {
  if (!TOKEN) { showConfirm('ERROR', 'NOT LOGGED IN.', () => { }); return; }
  const u = canonicalUnit(uId);
  if (!u) { showConfirm('ERROR', 'USAGE: UH <UNIT> [HOURS]', () => { }); return; }

  UH_CURRENT_UNIT = u;
  UH_CURRENT_HOURS = Number(h || 12);
  document.getElementById('uhTitle').textContent = 'UNIT HISTORY';
  document.getElementById('uhUnit').textContent = u;
  document.getElementById('uhHours').value = String(UH_CURRENT_HOURS);
  document.getElementById('uhBack').style.display = 'flex';
  document.getElementById('uhBody').innerHTML = '<tr><td colspan="7" class="muted">LOADING…</td></tr>';

  setLive(true, 'LIVE • UNIT HISTORY');
  const r = await API.getUnitHistory(TOKEN, u, UH_CURRENT_HOURS);
  if (!r || !r.ok) return showErr(r);

  const rs = r.rows || [];
  if (!rs.length) {
    document.getElementById('uhBody').innerHTML = '<tr><td colspan="7" class="muted">NO HISTORY IN THIS WINDOW.</td></tr>';
    return;
  }

  rs.sort((a, b) => new Date(b.ts || 0) - new Date(a.ts || 0));
  document.getElementById('uhBody').innerHTML = rs.map(rr => {
    const ts = rr.ts ? fmtTime24(rr.ts) : '—';
    const nx = rr.next || {};
    const st = String(nx.status || '').toUpperCase();
    const aC = getRoleColor(rr.actor);
    return `<tr>
      <td>${esc(ts)}</td>
      <td>${esc((rr.action || '').toUpperCase())}</td>
      <td>${esc(st || '—')}</td>
      <td>${nx.note ? esc(String(nx.note || '').toUpperCase()) : '<span class="muted">—</span>'}</td>
      <td>${nx.incident ? esc(String(nx.incident || '').toUpperCase()) : '<span class="muted">—</span>'}</td>
      <td>${nx.destination ? esc(String(nx.destination || '').toUpperCase()) : '<span class="muted">—</span>'}</td>
      <td class="muted ${aC}">${rr.actor ? esc(String(rr.actor || '').toUpperCase()) : '<span class="muted">—</span>'}</td>
    </tr>`;
  }).join('');
}

// ============================================================
// Messages Modal
// ============================================================
function openMessages() {
  if (!TOKEN) { showConfirm('ERROR', 'NOT LOGGED IN.', () => { }); return; }
  const ms = STATE.messages || [];
  const li = document.getElementById('msgList');
  document.getElementById('msgModalCount').textContent = ms.length;

  if (!ms.length) {
    li.innerHTML = '<div class="muted" style="padding:20px; text-align:center;">NO MESSAGES</div>';
  } else {
    li.innerHTML = ms.map(m => {
      const cl = ['msgItem'];
      if (!m.read) cl.push('unread');
      if (m.urgent) cl.push('urgent');
      const fr = m.from_initials + '@' + m.from_role;
      const fC = getRoleColor(fr);
      const uH = m.urgent ? '<div class="msgUrgent">URGENT</div>' : '';
      return `<div class="${cl.join(' ')}" onclick="viewMessage('${esc(m.message_id)}')">
        <div class="msgHeader">
          <span class="msgFrom ${fC}">FROM ${esc(fr)}</span>
          <span class="msgTime">${fmtTime24(m.ts)}</span>
        </div>
        ${uH}
        <div class="msgText">${esc(m.message)}</div>
      </div>`;
    }).join('');
  }
  document.getElementById('msgBack').style.display = 'flex';
}

function closeMessages() {
  document.getElementById('msgBack').style.display = 'none';
  refresh();
}

async function viewMessage(mId) {
  const r = await API.readMessage(TOKEN, mId);
  if (!r.ok) return showErr(r);
  refresh();
}

async function deleteMessage(mId) {
  const r = await API.deleteMessage(TOKEN, mId);
  if (!r.ok) return showErr(r);
  closeMessages();
  refresh();
}

async function deleteAllMessages() {
  const r = await API.deleteAllMessages(TOKEN);
  if (!r.ok) return showErr(r);
  closeMessages();
  refresh();
}

function replyToMessage(cmd) {
  document.getElementById('cmd').value = cmd;
  document.getElementById('cmd').focus();
}

// ============================================================
// Export & Metrics
// ============================================================
async function exportCsv(h) {
  const r = await API.exportAuditCsv(TOKEN, h);
  if (!r.ok) return showErr(r);
  const b = new Blob([r.csv], { type: 'text/csv;charset=utf-8;' });
  const u = URL.createObjectURL(b);
  const a = document.createElement('a');
  a.href = u;
  a.download = r.filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(u);
}

// ============================================================
// Command Parser & Runner
// ============================================================
async function runCommand() {
  const cE = document.getElementById('cmd');
  let tx = (cE.value || '').trim();
  if (!tx) return;

  CMD_HISTORY.push(tx);
  if (CMD_HISTORY.length > 50) CMD_HISTORY.shift();
  CMD_INDEX = CMD_HISTORY.length;
  cE.value = '';

  // Command chaining: split on ; (NC uses ; internally, so consume remainder)
  const segments = [];
  let _remaining = tx;
  while (_remaining.length > 0) {
    const _trimmed = _remaining.trimStart();
    if (/^NC\s/i.test(_trimmed) || /^NC$/i.test(_trimmed)) {
      segments.push(_trimmed.trim());
      break;
    }
    const _semiIdx = _remaining.indexOf(';');
    if (_semiIdx < 0) {
      segments.push(_remaining.trim());
      break;
    }
    const _seg = _remaining.substring(0, _semiIdx).trim();
    if (_seg) segments.push(_seg);
    _remaining = _remaining.substring(_semiIdx + 1);
  }
  for (const segment of segments) {
    await _execCmd(segment);
  }
}

const _ONE_TOKEN_CMDS = new Set([
  'NC', 'MSGALL', 'HTALL', 'MSGU', 'HTU', 'HTMSU', 'MSGDP', 'HTDP',
  'NOTE', 'ALERT', 'NEWUSER', 'DELUSER', 'CLEARDATA', 'PASSWD',
  'REPORTOOS', 'REPORT', 'REPORTUTIL', 'REPORTSHIFT',
  'MAP', 'LOC'
]);

async function _execCmd(tx) {
  const tokens = tx.split(/\s+/);
  const cmd = (tokens[0] || '').toUpperCase();

  let ma, no;
  if (_ONE_TOKEN_CMDS.has(cmd)) {
    ma = cmd;
    no = tokens.slice(1).join(' ');
  } else if (tokens.length <= 2) {
    ma = tx;
    no = '';
  } else {
    const idx1 = tx.indexOf(' ');
    const idx2 = idx1 >= 0 ? tx.indexOf(' ', idx1 + 1) : -1;
    if (idx2 >= 0) {
      ma = tx.substring(0, idx2);
      no = tx.substring(idx2 + 1);
    } else {
      ma = tx;
      no = '';
    }
  }

  const mU = ma.toUpperCase();
  const nU = expandShortcutsInText(no || '');

  if (mU === 'HELP' || mU === 'H') return showHelp();
  if (mU === 'BUG') { openBugReport(); return; }
  if (mU === 'UNDO') {
    if (!_undoStack.length) { showToast('NOTHING TO UNDO.'); return; }
    const entry = _undoStack[_undoStack.length - 1];
    if (Date.now() - entry.ts > 5 * 60 * 1000) { _undoStack.pop(); showToast('UNDO EXPIRED (> 5 MIN).'); return; }
    _undoStack.pop();
    showToast('UNDOING: ' + entry.description + '...');
    try {
      await entry.revertFn();
      setLive(true, 'LIVE • UNDO');
      refresh();
      showToast('UNDONE: ' + entry.description);
    } catch (e) {
      showErr({ error: 'UNDO FAILED: ' + (e.message || String(e)) });
    }
    return;
  }
  if (mU === 'ADMIN') return showAdmin();
  if (mU === 'REFRESH') { forceRefresh(); return; }
  if (mU === 'POPOUT') { openPopout(); return; }
  if (mU === 'POPIN')  { closePopin(); return; }
  if (mU === 'POPMAP') { openPopoutMap(); return; }
  if (mU === 'POPINC') { openPopoutInc(); return; }
  if (mU === 'MAPIN')  { mapZoomIn(); return; }
  if (mU === 'MAPOUT') { mapZoomOut(); return; }
  if (mU === 'MAPFIT') { mapFitAll(); return; }
  if (mU === 'MAPSTA') { mapShowStations(); return; }
  if (mU === 'MAPINC') { mapShowIncidents(); return; }
  if (mU === 'MAPRESET') { mapReset(); return; }
  if (mU === 'MAPR') { _ensureMapOpen(() => { renderBoardMap(); showToast('MAP REFRESHED'); }); return; }
  if (mU === 'MAP') {
    const mapArg = (nU || '').toUpperCase().trim();
    if (!mapArg) { toggleBoardMap(); return; }
    if (mapArg === 'IN') { mapZoomIn(); return; }
    if (mapArg === 'OUT') { mapZoomOut(); return; }
    if (mapArg === 'FIT') { mapFitAll(); return; }
    if (mapArg === 'STA' || mapArg === 'STATIONS') { mapShowStations(); return; }
    if (mapArg === 'INC') { mapShowIncidents(); return; }
    if (mapArg === 'RESET') { mapReset(); return; }
    if (mapArg === 'R' || mapArg === 'REFRESH') {
      _ensureMapOpen(() => { renderBoardMap(); showToast('MAP REFRESHED'); });
      return;
    }
    if (mapArg === 'CLR' || mapArg === 'CLEAR') {
      _bmGeoCache = Object.assign({}, BM_KNOWN_COORDS);
      _clearSearchPin();
      _ensureMapOpen(() => { renderBoardMap(); showToast('MAP CACHE CLEARED + REFRESHED'); });
      return;
    }
    // MAP INC<id> — focus on incident scene
    if (/^INC\s*\d+$|^\d{2}-\d+$/.test(mapArg)) {
      focusIncidentOnMap(mapArg.replace(/\s+/g, ''));
      return;
    }
    // Check if arg matches a known unit — if not, treat as address to geocode
    const knownUnit = (STATE && STATE.units || []).find(u => u.unit_id === mapArg.split(/\s/)[0]);
    if (knownUnit) {
      focusUnitOnMap(mapArg.split(/\s/)[0]);
    } else {
      // MAP <address> — geocode and focus
      mapGeoFocus(mapArg);
    }
    return;
  }


  // ── VIEW / DISPLAY COMMANDS ──

  // V SIDE/MSG/MET/INC/ALL/NONE
  if (/^V\s+/i.test(mU)) {
    const panel = mU.substring(2).trim();
    if (panel === 'SIDE') toggleView('sidebar');
    else if (panel === 'MSG') toggleView('messages');
    else if (panel === 'INC') toggleView('incidents');
    else if (panel === 'ALL') toggleView('all');
    else if (panel === 'NONE') toggleView('none');
    else { showAlert('ERROR', 'USAGE: V SIDE/MSG/INC/ALL/NONE'); }
    return;
  }

  // F <STATUS> / F ALL — filter
  if (/^F\s+/i.test(mU) || mU === 'F') {
    const arg = mU.substring(2).trim();
    if (!arg || arg === 'ALL') {
      VIEW.filterStatus = null;
    } else if (VALID_STATUSES.has(arg)) {
      VIEW.filterStatus = arg;
    } else {
      showAlert('ERROR', 'USAGE: F <STATUS> OR F ALL\nVALID: D, DE, OS, F, FD, T, TH, AV, UV, BRK, OOS, IQ');
      return;
    }
    const tbFs = document.getElementById('tbFilterStatus');
    if (tbFs) tbFs.value = VIEW.filterStatus || '';
    saveViewState();
    renderBoardDiff();
    return;
  }

  // SORT STATUS/UNIT/ELAPSED/UPDATED/REV
  if (/^SORT\s+/i.test(mU)) {
    const arg = mU.substring(5).trim();
    if (arg === 'REV') {
      VIEW.sortDir = VIEW.sortDir === 'asc' ? 'desc' : 'asc';
    } else if (['STATUS', 'UNIT', 'ELAPSED', 'UPDATED'].includes(arg)) {
      VIEW.sort = arg.toLowerCase();
      VIEW.sortDir = 'asc';
    } else {
      showAlert('ERROR', 'USAGE: SORT STATUS/UNIT/ELAPSED/UPDATED/REV');
      return;
    }
    const tbSort = document.getElementById('tbSort');
    if (tbSort) tbSort.value = VIEW.sort;
    saveViewState();
    updateSortHeaders();
    renderBoardDiff();
    return;
  }

  // NIGHT — toggle night mode
  if (mU === 'NIGHT') {
    toggleNightMode();
    return;
  }

  // DEN / DEN COMPACT/NORMAL/EXPANDED
  if (/^DEN$/i.test(mU)) {
    cycleDensity();
    return;
  }
  if (/^DEN\s+/i.test(mU)) {
    const arg = mU.substring(4).trim();
    if (['COMPACT', 'NORMAL', 'EXPANDED'].includes(arg)) {
      VIEW.density = arg.toLowerCase();
      saveViewState();
      applyViewState();
    } else {
      showAlert('ERROR', 'USAGE: DEN COMPACT/NORMAL/EXPANDED');
    }
    return;
  }

  // PRESET DISPATCH/SUPERVISOR/FIELD
  if (/^PRESET\s+/i.test(mU)) {
    const arg = mU.substring(7).trim().toLowerCase();
    if (['dispatch', 'supervisor', 'field'].includes(arg)) {
      applyPreset(arg);
    } else {
      showAlert('ERROR', 'USAGE: PRESET DISPATCH/SUPERVISOR/FIELD');
    }
    return;
  }

  // ELAPSED SHORT/LONG/OFF
  if (/^ELAPSED\s+/i.test(mU)) {
    const arg = mU.substring(8).trim().toLowerCase();
    if (['short', 'long', 'off'].includes(arg)) {
      VIEW.elapsedFormat = arg;
      saveViewState();
      renderBoardDiff();
    } else {
      showAlert('ERROR', 'USAGE: ELAPSED SHORT/LONG/OFF');
    }
    return;
  }

  // CLR <UNIT> - clear unit from incident without status change
  if (mU.startsWith('CLR ')) {
    const unitId = mU.substring(4).trim().toUpperCase();
    if (unitId) {
      setLive(true, 'LIVE • CLR UNIT');
      const r = await API.clearUnitIncident(TOKEN, unitId);
      if (!r.ok) return showErr(r);
      showToast('CLEARED ' + unitId + ' FROM ' + (r.clearedIncident || 'INCIDENT'));
      refresh();
      return;
    }
  }

  // CLR - clear filters + search
  if (mU === 'CLR') {
    VIEW.filterStatus = null;
    ACTIVE_INCIDENT_FILTER = '';
    document.getElementById('search').value = '';
    const tbFs = document.getElementById('tbFilterStatus');
    if (tbFs) tbFs.value = '';
    saveViewState();
    renderBoardDiff();
    return;
  }


  // INBOX - open/focus inbox panel
  if (mU === 'INBOX') {
    const p = document.getElementById('msgInboxPanel');
    if (p && p.classList.contains('collapsed')) p.classList.remove('collapsed');
    const list = document.getElementById('msgInboxList');
    if (list) list.scrollTop = 0;
    return;
  }

  // NOTES / SCRATCH - focus scratch notes
  if (mU === 'NOTES' || mU === 'SCRATCH') {
    if (VIEW.sidebar) {
      const pad = document.getElementById('scratchPadSide');
      if (pad) pad.focus();
    } else {
      const p = document.getElementById('scratchPanel');
      if (p && p.classList.contains('collapsed')) p.classList.remove('collapsed');
      const pad = document.getElementById('scratchPad');
      if (pad) pad.focus();
    }
    return;
  }

  // IQ <UNIT> — set in quarters (returns to station address)
  if (mU.startsWith('IQ ') || (mU === 'IQ' && SELECTED_UNIT_ID)) {
    const targetUnit = mU === 'IQ' ? SELECTED_UNIT_ID : canonicalUnit(mU.substring(3).trim());
    if (!targetUnit) { showAlert('ERROR', 'USAGE: IQ <UNIT>'); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === targetUnit) : null;
    if (!uO) { showAlert('ERROR', 'UNIT NOT FOUND: ' + targetUnit); return; }
    const station = (uO.station || '').trim();
    if (!station) {
      showAlert('ERROR', targetUnit + ' HAS NO STATION/QUARTERS ASSIGNED.\nUPDATE VIA UNIT MODAL OR ADMIN ROSTER.');
      return;
    }
    setLive(true, 'LIVE \u2022 IQ');
    const patch = { status: 'IQ', displayName: uO.display_name };
    const r = await API.upsertUnit(TOKEN, targetUnit, patch, uO.updated_at || '');
    if (!r.ok) return showErr(r);
    showToast(targetUnit + ': IN QUARTERS @ ' + station);
    refresh();
    return;
  }

  // SUP <UNIT> <MESSAGE> — supplemental update to unit's incident
  if (mU.startsWith('SUP ')) {
    const supUnit = canonicalUnit(mU.substring(4).trim());
    if (!supUnit) { showAlert('ERROR', 'USAGE: SUP <UNIT> <MESSAGE>'); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === supUnit) : null;
    if (!uO) { showAlert('ERROR', 'UNIT NOT FOUND: ' + supUnit); return; }
    if (!uO.incident) { showAlert('ERROR', supUnit + ' IS NOT ASSIGNED TO AN INCIDENT.'); return; }
    if (!nU) { showAlert('ERROR', 'USAGE: SUP <UNIT> <MESSAGE>'); return; }
    setLive(true, 'LIVE \u2022 SUP');
    const r = await API.appendIncidentNote(TOKEN, uO.incident, '[SUP] ' + nU);
    if (!r.ok) return showErr(r);
    showToast('SUP SENT TO ' + supUnit + ' (' + uO.incident + ')');
    refresh();
    return;
  }

  // ── BARE STATUS CODE with selected unit ──
  if (SELECTED_UNIT_ID && VALID_STATUSES.has(mU) && !no) {
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === SELECTED_UNIT_ID) : null;
    if (uO) {
      quickStatus(uO, mU);
      return;
    }
  }

  // ── EXISTING COMMANDS (unchanged) ──

  // LUI - Create temp one-off unit (or open existing unit modal if unit is already known)
  if (mU === 'LUI' || mU.startsWith('LUI ')) {
    const luiPrefill = mU.startsWith('LUI ') ? ma.substring(4).trim().toUpperCase() : '';
    // If unit already exists on the board, open its real modal — never stamp [TEMP] on a known unit
    if (luiPrefill && STATE) {
      const existing = (STATE.units || []).find(function(u) { return u.unit_id === luiPrefill; });
      if (existing) {
        openModal(existing);
        showToast('LUI: ' + luiPrefill + ' ALREADY ON BOARD — OPENING UNIT MODAL.');
        return;
      }
    }
    const dN = luiPrefill ? displayNameForUnit(canonicalUnit(luiPrefill)) : '';
    const fakeUnit = {
      unit_id: luiPrefill,
      display_name: dN || luiPrefill,
      active: false,
      status: 'AV',
      note: '[TEMP]',
      type: 'EMS',
      level: '',
      station: ''
    };
    openModal(fakeUnit);
    showToast('LUI: CREATING TEMP UNIT. FILL IN DETAILS AND SAVE.');
    return;
  }

  // User management
  if (mU === 'NEWUSER' && nU) {
    const parts = nU.trim().split(',');
    if (parts.length !== 2) { showAlert('ERROR', 'USAGE: NEWUSER lastname,firstname'); return; }
    const r = await API.newUser(TOKEN, parts[0].trim(), parts[1].trim());
    if (!r.ok) return showErr(r);
    const collisionMsg = r.collision ? '\n\nUSERNAME COLLISION - NUMBER ADDED' : '';
    showAlert('USER CREATED', `NEW USER CREATED:${collisionMsg}\n\nNAME: ${r.firstName} ${r.lastName}\nUSERNAME: ${r.username}\nPASSWORD: ${r.password}\n\nUser can now log in with this username and password.`);
    return;
  }

  if (mU === 'DELUSER' && nU) {
    const deluserRaw = nU.trim();
    const deluserParts = deluserRaw.split(/\s+/);
    const deluserHasConfirm = deluserParts[deluserParts.length - 1].toUpperCase() === 'CONFIRM';
    const u = deluserHasConfirm ? deluserParts.slice(0, -1).join(' ') : deluserRaw;
    if (!u) { showAlert('ERROR', 'USAGE: DELUSER username CONFIRM'); return; }
    if (!deluserHasConfirm) {
      showErr({ error: 'CONFIRMATION REQUIRED. RE-RUN WITH CONFIRM. EXAMPLE: DELUSER ' + u + ' CONFIRM' });
      return;
    }
    showConfirm('CONFIRM DELETE USER', 'DELETE USER: ' + u + '?', async () => {
      const r = await API.delUser(TOKEN, u);
      if (!r.ok) return showErr(r);
      showAlert('USER DELETED', 'USER DELETED: ' + r.username);
    });
    return;
  }

  // REPORTOOS - Out of Service report
  if (mU === 'REPORTOOS' || /^REPORTOOS\d/.test(mU)) {
    const ts = (mU === 'REPORTOOS' ? (nU || '') : mU.substring(9)).trim().toUpperCase();
    let hrs = 24;
    if (ts) {
      const m = ts.match(/^(\d+)(H|D)?$/);
      if (m) {
        const n = parseInt(m[1]);
        const ut = m[2] || 'H';
        hrs = ut === 'D' ? n * 24 : n;
      } else {
        showAlert('ERROR', 'USAGE: REPORTOOS [24H|7D|30D]\nH=HOURS, D=DAYS\nExample: REPORTOOS 24H or REPORTOOS 7D');
        return;
      }
    }
    const r = await API.reportOOS(TOKEN, hrs);
    if (!r.ok) return showErr(r);
    const rp = r.report || {};
    let out = `OUT OF SERVICE REPORT\n${hrs}H PERIOD (${rp.startTime} TO ${rp.endTime})\n\n`;
    out += '='.repeat(47) + '\n';
    out += `TOTAL OOS TIME: ${rp.totalOOSMinutes} MINUTES (${rp.totalOOSHours} HOURS)\n`;
    out += `TOTAL UNITS: ${rp.unitCount}\n`;
    out += '='.repeat(47) + '\n\n';
    if (rp.units && rp.units.length > 0) {
      out += 'UNIT BREAKDOWN:\n\n';
      rp.units.forEach(u => {
        out += `${u.unit.padEnd(12)} ${String(u.oosMinutes).padStart(6)} MIN  ${u.oosHours} HRS\n`;
        if (u.periods && u.periods.length > 0) {
          u.periods.forEach(p => {
            out += `  ${p.start} -> ${p.end} (${p.duration}M)\n`;
          });
        }
        out += '\n';
      });
    } else {
      out += 'NO OOS TIME RECORDED IN THIS PERIOD\n';
    }
    showAlert('OOS REPORT', out);
    return;
  }

  // REPORT SHIFT — printable shift summary
  if (mU === 'REPORT' && /^SHIFT\b/i.test(nU)) {
    const hrs = parseFloat(nU.replace(/^SHIFT\s*/i, '').trim()) || 12;
    setLive(true, 'LIVE • SHIFT REPORT');
    const r = await API.getShiftReport(TOKEN, hrs);
    if (!r.ok) return showErr(r);
    openShiftReportWindow(r);
    return;
  }
  if (mU === 'REPORTSHIFT') {
    const hrs = parseFloat(nU || '') || 12;
    setLive(true, 'LIVE • SHIFT REPORT');
    const r = await API.getShiftReport(TOKEN, hrs);
    if (!r.ok) return showErr(r);
    openShiftReportWindow(r);
    return;
  }

  // REPORT INC — printable per-incident report
  if (mU === 'REPORT' && /^INC/i.test(nU)) {
    const iId = nU.replace(/^INC\s*/i, '').trim();
    if (!iId) { showAlert('USAGE', 'REPORT INC <ID>\nExample: REPORT INC1234'); return; }
    setLive(true, 'LIVE • INCIDENT REPORT');
    const r = await API.getIncident(TOKEN, iId);
    if (!r.ok) return showErr(r);
    openIncidentPrintWindow(r);
    return;
  }

  // REPORTUTIL — per-unit utilization report
  if (mU === 'REPORTUTIL') {
    const parts = (nU || '').trim().split(/\s+/);
    const uId  = parts[0] || '';
    const hrs  = parseFloat(parts[1]) || 24;
    if (!uId) { showAlert('USAGE', 'REPORTUTIL <UNIT> [HOURS]\nExample: REPORTUTIL EMS1 24'); return; }
    setLive(true, 'LIVE • UNIT REPORT');
    const r = await API.getUnitReport(TOKEN, uId, hrs);
    if (!r.ok) return showErr(r);
    openUnitReportWindow(r);
    return;
  }

  // SUGGEST — recommend available units for an incident
  if (mU.startsWith('SUGGEST ')) {
    const iId = ma.substring(8).trim().toUpperCase();
    if (!iId) { showAlert('USAGE', 'SUGGEST INC0001'); return; }
    return suggestUnits(iId);
  }

  // DIVERSION — set/clear hospital diversion
  if (mU.startsWith('DIVERSION ')) {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    const onOff = parts[1] || '';
    const code = parts[2] || '';
    if ((onOff !== 'ON' && onOff !== 'OFF') || !code) {
      showAlert('USAGE', 'DIVERSION ON <CODE>\nDIVERSION OFF <CODE>\n\nCodes:  SCMC-BC (Bend)   SCMC-RC (Redmond)\n        SCMC-PR (Prineville)  SCMC-MD (Madras)\n        OHSU (Portland)\n\nShort:  SCB/SBH  SCR/SRH  SCP/SPH  SCM/SMH\nExample: DIVERSION ON SCB');
      return;
    }
    const active = onOff === 'ON';
    setLive(true, 'LIVE');
    const r = await API.setDiversion(TOKEN, code, active);
    if (!r.ok) return showErr(r);
    showToast((active ? 'DIVERSION ON: ' : 'DIVERSION OFF: ') + r.code, active ? 'warn' : 'ok');
    return;
  }

  // SCOPE — set board view scope to all agencies or a specific agency
  if (mU.startsWith('SCOPE ')) {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    const scopeArg = parts.slice(1).join(' ').trim().toUpperCase();
    if (!scopeArg) {
      showAlert('ERROR', 'USAGE: SCOPE ALL | SCOPE AGENCY <ID>');
    } else {
      const r = await API.setScope(TOKEN, scopeArg);
      if (r.ok) {
        showToast('SCOPE SET: ' + r.scope, 'success');
        updateScopeIndicator(r.scope);
      } else {
        showErr(r);
      }
    }
    return;
  }

  // ── Phase 2D: Stacked Assignment Commands ──────────────────────────

  // ASSIGN <INC> <UNIT>  — set incident as unit's primary assignment
  if (mU.startsWith('ASSIGN ')) {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    if (parts.length >= 3) {
      const incArg = parts[1].replace(/^INC-?/i, '').trim();
      const unitArg = parts[2].trim().toUpperCase();
      const incId = incArg.includes('-') ? incArg : (new Date().getFullYear() % 100 + '-' + incArg);
      setLive(true, 'LIVE • ASSIGN');
      const r = await API.assignUnit(TOKEN, incId, unitArg);
      if (r.ok) { showToast(unitArg + ' ASSIGNED TO ' + incId + ' AS PRIMARY.', 'success'); refresh(); }
      else showErr(r);
    } else {
      showAlert('ERROR', 'USAGE: ASSIGN <INC> <UNIT>\nExample: ASSIGN 26-0023 EMS1');
    }
    return;
  }

  // QUEUE <INC> <UNIT>  — add incident to unit's queue (behind primary)
  // QUE / QUEU accepted as common misspellings
  if (mU.startsWith('QUEUE ') || mU.startsWith('QUE ') || mU.startsWith('QUEU ')) {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    if (parts.length >= 3) {
      const incArg = parts[1].replace(/^INC-?/i, '').trim();
      const unitArg = parts[2].trim().toUpperCase();
      const incId = incArg.includes('-') ? incArg : (new Date().getFullYear() % 100 + '-' + incArg);
      setLive(true, 'LIVE • QUEUE');
      const r = await API.queueUnit(TOKEN, incId, unitArg);
      if (r.ok) { showToast(unitArg + ': ' + incId + (r.action === 'queued' ? ' QUEUED.' : ' ASSIGNED.'), 'success'); refresh(); }
      else showErr(r);
    } else {
      showAlert('ERROR', 'USAGE: QUEUE <INC> <UNIT>\nExample: QUEUE 26-0024 EMS1');
    }
    return;
  }

  // PRIMARY <INC> <UNIT>  — promote queued assignment to primary
  if (mU.startsWith('PRIMARY ')) {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    if (parts.length >= 3) {
      const incArg = parts[1].replace(/^INC-?/i, '').trim();
      const unitArg = parts[2].trim().toUpperCase();
      const incId = incArg.includes('-') ? incArg : (new Date().getFullYear() % 100 + '-' + incArg);
      setLive(true, 'LIVE • PRIMARY');
      const r = await API.primaryUnit(TOKEN, incId, unitArg);
      if (r.ok) { showToast(incId + ' IS NOW PRIMARY FOR ' + unitArg + '.', 'success'); refresh(); }
      else showErr(r);
    } else {
      showAlert('ERROR', 'USAGE: PRIMARY <INC> <UNIT>\nExample: PRIMARY 26-0023 EMS1');
    }
    return;
  }

  // CLEAR <INC> <UNIT>  — remove assignment from unit stack
  // Note: placed before CLEARDATA check is irrelevant because CLEARDATA uses startsWith('CLEARDATA ')
  // and this check requires exactly 3 parts starting with CLEAR (not CLEARDATA).
  if (mU.startsWith('CLEAR ') && !mU.startsWith('CLEARDATA ')) {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    if (parts.length >= 3) {
      const incArg = parts[1].replace(/^INC-?/i, '').trim();
      const unitArg = parts[2].trim().toUpperCase();
      const incId = incArg.includes('-') ? incArg : (new Date().getFullYear() % 100 + '-' + incArg);
      const confirmed = await showConfirmAsync('CLEAR ASSIGNMENT', 'CLEAR ' + incId + ' FROM ' + unitArg + "'S STACK?");
      if (!confirmed) return;
      setLive(true, 'LIVE • CLEAR');
      const r = await API.clearUnitAssignment(TOKEN, incId, unitArg);
      if (r.ok) {
        const msg = r.promoted
          ? unitArg + ': ' + incId + ' CLEARED. ' + r.promoted + ' PROMOTED TO PRIMARY.'
          : unitArg + ': ' + incId + ' CLEARED.';
        showToast(msg, 'success'); refresh();
      } else {
        showErr(r);
      }
    } else {
      showAlert('ERROR', 'USAGE: CLEAR <INC> <UNIT>\nExample: CLEAR 26-0024 EMS1');
    }
    return;
  }

  // STACK <INC> <UNIT>  — smart-assign: primary if unit free, queued if unit busy
  // STACK <UNIT>        — show unit's assignment stack (read)
  // Disambiguates by checking if 2nd token looks like an incident ID
  if (mU.startsWith('STACK ')) {
    const stackParts = (mU + ' ' + nU).trim().split(/\s+/);
    const secondToken = stackParts[1] || '';
    const looksLikeInc = /^\d{2}-\d{4}$/i.test(secondToken) || /^[A-Z]{1,4}\d{2}-\d{4}$/i.test(secondToken) || /^INC/i.test(secondToken) || /^0\d{3}$/.test(secondToken);
    if (stackParts.length >= 3 && looksLikeInc) {
      // Write path: STACK <INC> <UNIT>
      const incArg = stackParts[1].replace(/^INC-?/i, '').trim();
      const unitArg = stackParts[2].trim().toUpperCase();
      const incId = incArg.includes('-') ? incArg : (new Date().getFullYear() % 100 + '-' + incArg);
      setLive(true, 'LIVE • STACK');
      const r = await API.queueUnit(TOKEN, incId, unitArg);
      if (r.ok) {
        showToast(unitArg + ': ' + incId + (r.action === 'queued' ? ' STACKED (QUEUED).' : ' STACKED (PRIMARY).'), 'success');
        refresh();
      } else { showErr(r); }
    } else if (stackParts.length >= 2) {
      // Read path: STACK <UNIT>
      const unitArg = stackParts[1].trim().toUpperCase();
      setLive(true, 'LIVE • STACK');
      const r = await API.getUnitStack(TOKEN, unitArg);
      if (!r.ok) { showErr(r); return; }
      if (!r.stack || !r.stack.length) {
        showAlert('UNIT STACK — ' + unitArg, unitArg + ' HAS NO QUEUED ASSIGNMENTS.');
      } else {
        let lines = [unitArg + ' STACK [' + r.stack.length + ' ASSIGNMENT' + (r.stack.length !== 1 ? 'S' : '') + ']:'];
        r.stack.forEach(a => {
          const lbl = a.is_primary ? '#' + a.assignment_order + ' PRIMARY ' : '#' + a.assignment_order + ' QUEUED  ';
          const dest = a.destination ? '/ ' + a.destination : '';
          lines.push('  ' + lbl + a.incident_id + '  ' + (a.incident_type || '--') + ' ' + dest);
        });
        showAlert('UNIT STACK — ' + unitArg, lines.join('\n'));
      }
    } else {
      showAlert('ERROR', 'USAGE: STACK <INC> <UNIT>  or  STACK <UNIT>');
    }
    return;
  }

  if (mU === 'LISTUSERS') {
    const r = await API.listUsers(TOKEN);
    if (!r.ok) return showErr(r);
    const users = r.users || [];
    if (!users.length) { showAlert('USERS', 'NO USERS IN SYSTEM'); return; }
    const userList = users.map(u => `${u.username} - ${u.firstName} ${u.lastName}`).join('\n');
    showAlert('SYSTEM USERS (' + users.length + ')', userList);
    return;
  }

  if (mU === 'PASSWD' && nU) {
    const parts = nU.trim().split(/\s+/);
    if (parts.length !== 2) { showAlert('ERROR', 'USAGE: PASSWD oldpassword newpassword'); return; }
    const r = await API.changePassword(TOKEN, parts[0], parts[1]);
    if (!r.ok) return showErr(r);
    showAlert('PASSWORD CHANGED', 'YOUR PASSWORD HAS BEEN CHANGED SUCCESSFULLY.');
    return;
  }

  // Search
  if (mU.startsWith('! ')) {
    const query = (ma.substring(2).trim() + (no ? ' ' + no : '')).trim().toUpperCase();
    if (!query || query.length < 2) { showAlert('ERROR', 'USAGE: ! searchtext (min 2 chars)'); return; }
    const r = await API.search(TOKEN, query);
    if (!r.ok) return showErr(r);
    const results = r.results || [];
    if (!results.length) { showAlert('SEARCH RESULTS', 'NO RESULTS FOUND FOR: ' + query); return; }
    let report = 'SEARCH RESULTS FOR: ' + query + '\n\n';
    results.forEach(res => { report += `[${res.type}] ${res.summary}\n`; });
    showAlert('SEARCH RESULTS (' + results.length + ')', report);
    return;
  }

  // Clear data (admin roles only)
  if (mU === 'CLEARDATA' && nU) {
    if (!isAdminRole()) {
      showAlert('ACCESS DENIED', 'CLEARDATA COMMANDS REQUIRE ADMIN LOGIN (SUPV/MGR/IT).');
      return;
    }
    const whatRaw = nU.trim().toUpperCase();
    const whatParts = whatRaw.split(/\s+/);
    const hasConfirm = whatParts[whatParts.length - 1] === 'CONFIRM';
    const what = hasConfirm ? whatParts.slice(0, -1).join(' ') : whatRaw;
    if (!['UNITS', 'INACTIVE', 'AUDIT', 'INCIDENTS', 'MESSAGES', 'SESSIONS', 'ALL'].includes(what)) {
      showAlert('ERROR', 'USAGE: CLEARDATA [UNITS|INACTIVE|AUDIT|INCIDENTS|MESSAGES|SESSIONS|ALL]');
      return;
    }
    if (!hasConfirm) {
      showErr({ error: 'CONFIRMATION REQUIRED. RE-RUN WITH CONFIRM.\nExample: CLEARDATA ' + what + ' CONFIRM' });
      return;
    }
    // SESSIONS uses a different API endpoint
    if (what === 'SESSIONS') {
      showConfirm('CONFIRM SESSION CLEAR', 'LOG OUT ALL USERS?\n\nTHIS WILL FORCE EVERYONE TO RE-LOGIN.', async () => {
        const r = await API.clearSessions(TOKEN);
        if (!r.ok) return showErr(r);
        showAlert('SESSIONS CLEARED', `${r.deleted} SESSIONS CLEARED. ALL USERS LOGGED OUT.`);
      });
      return;
    }
    showConfirm('CONFIRM DATA CLEAR', `CLEAR ALL ${what} DATA?\n\nTHIS CANNOT BE UNDONE!`, async () => {
      const r = await API.clearData(TOKEN, what);
      if (!r.ok) return showErr(r);
      showAlert('DATA CLEARED', `${what} DATA CLEARED: ${r.deleted} ROWS DELETED`);
      refresh();
    });
    return;
  }

  // Unit status report
  if (mU === 'US') {
    if (!STATE || !STATE.units) { showAlert('ERROR', 'NO DATA LOADED'); return; }
    const units = (STATE.units || []).filter(u => u.active).sort((a, b) => {
      const ra = statusRank(a.status);
      const rb = statusRank(b.status);
      if (ra !== rb) return ra - rb;
      return String(a.unit_id || '').localeCompare(String(b.unit_id || ''));
    });
    let report = 'UNIT STATUS REPORT\n\n';
    units.forEach(u => {
      const statusLabel = (STATE.statuses || []).find(s => s.code === u.status)?.label || u.status;
      const mins = minutesSince(u.updated_at);
      const age = mins != null ? Math.floor(mins) + 'M' : '—';
      report += `${u.unit_id.padEnd(12)} ${u.status.padEnd(4)} ${statusLabel.padEnd(20)} ${age.padEnd(6)}\n`;
      if (u.incident) report += `  INC: ${u.incident}\n`;
      if (u.destination) {
        const dr = AddressLookup.resolve(u.destination);
        report += `  DEST: ${dr.recognized ? dr.addr.name + ' [' + dr.addr.id + ']' : u.destination}\n`;
      }
      if (u.note) report += `  NOTE: ${u.note}\n`;
    });
    showAlert('UNIT STATUS', report);
    return;
  }

  // WHO [identifier] — dispatchers online, or look up specific unit/role/user
  if (mU === 'WHO' || mU.startsWith('WHO ')) {
    const filter = mU.startsWith('WHO ') ? ma.substring(4).trim().toUpperCase() : '';

    if (filter) {
      // First: check if it's an active unit on the board (field device lookup)
      const unit = ((STATE && STATE.units) || []).find(u => u.unit_id.toUpperCase() === filter);
      if (unit) {
        if (!unit.active) {
          showAlert('WHO ' + filter, filter + ' IS NOT LOGGED IN', 'yellow');
        } else {
          const crew = parseCrewInfo(unit.unit_info);
          const st = unit.status || '--';
          const inc = unit.incident ? '\nINCIDENT: ' + unit.incident : '';
          const dest = unit.destination ? '\nDESTINATION: ' + unit.destination : '';
          const crewLine = crew ? '\nCREW: ' + crew : '\nNO CREW INFO ON FILE';
          showAlert('WHO ' + filter, 'STATUS: ' + st + inc + dest + crewLine, 'yellow');
        }
        return;
      }

      // Otherwise: look up dispatcher by role or username in active sessions
      const r = await API.who(TOKEN, filter);
      if (!r.ok) return showErr(r);
      const users = r.users || [];
      if (!users.length) {
        showAlert('WHO ' + filter, filter + ' IS NOT LOGGED IN', 'yellow');
      } else {
        const userList = users.map(u => `${u.actor} (${u.minutesAgo}M AGO)`).join('\n');
        showAlert('WHO ' + filter, userList, 'yellow');
      }
      return;
    }

    // WHO with no args — list all online dispatchers
    const r = await API.who(TOKEN);
    if (!r.ok) return showErr(r);
    const users = r.users || [];
    if (!users.length) { showAlert('WHO', 'NO DISPATCHERS ONLINE', 'yellow'); return; }
    const userList = users.map(u => {
      const cad = u.cadId ? ' [' + u.cadId + ']' : '';
      return `${u.actor}${cad}  (${u.minutesAgo}M AGO)`;
    }).join('\n');
    showAlert('DISPATCHERS ONLINE (' + users.length + ')', userList, 'yellow');
    return;
  }

  // UR - active unit roster
  if (mU === 'UR') {
    const units = ((STATE && STATE.units) || []).filter(u => u.active);
    if (!units.length) { showAlert('UNIT ROSTER', 'NO ACTIVE UNITS ON BOARD', 'yellow'); return; }
    const lines = units.map(u => {
      const st = (u.status || '--').padEnd(4);
      const level = u.level ? ` [${u.level}]` : '';
      const inc = u.incident ? ` INC${u.incident}` : '';
      const dest = u.destination ? ` → ${u.destination}` : '';
      return `${String(u.unit_id).padEnd(8)} ${st}${level}${inc}${dest}`;
    }).join('\n');
    showAlert('UNIT ROSTER (' + units.length + ' ACTIVE)', lines, 'yellow');
    return;
  }

  // INCQ - incident queue quick view
  if (mU === 'INCQ') {
    const incidents = ((STATE && STATE.incidents) || []);
    const queued = incidents.filter(i => i.status === 'QUEUED');
    const active = incidents.filter(i => i.status === 'ACTIVE');
    if (!queued.length && !active.length) { showAlert('INCIDENT QUEUE', 'NO ACTIVE OR QUEUED INCIDENTS.', 'yellow'); return; }
    let report = '';
    if (queued.length) {
      report += '── QUEUED (' + queued.length + ') ──────────────────\n';
      queued.forEach(inc => {
        const shortId = String(inc.incident_id).replace(/^\d{2}-/, '');
        const pri = inc.priority ? ' ' + inc.priority : '';
        const dest = inc.destination ? ' → ' + inc.destination : '';
        const typ = inc.incident_type ? ' [' + inc.incident_type + ']' : '';
        const wait = inc.created_at ? Math.floor((Date.now() - new Date(inc.created_at).getTime()) / 60000) + 'M' : '';
        const note = (inc.incident_note || '').replace(/\[[^\]]*\]/g, '').replace(/\s{2,}/g, ' ').trim();
        report += `${inc.incident_id}${pri}${dest}${typ}  ${wait}\n`;
        if (note) report += `  ${note.substring(0, 60)}\n`;
      });
    }
    if (active.length) {
      if (report) report += '\n';
      report += '── ACTIVE (' + active.length + ') ─────────────────\n';
      active.forEach(inc => {
        const shortId = String(inc.incident_id).replace(/^\d{2}-/, '');
        const pri = inc.priority ? ' ' + inc.priority : '';
        const dest = inc.destination ? ' → ' + inc.destination : '';
        const typ = inc.incident_type ? ' [' + inc.incident_type + ']' : '';
        const units = inc.units ? ' UNITS: ' + inc.units : '';
        const scene = inc.scene_address ? ' @ ' + inc.scene_address : '';
        report += `${inc.incident_id}${pri}${dest}${typ}${units}\n`;
        if (scene) report += `  SCENE: ${scene.substring(0, 50)}\n`;
      });
    }
    showAlert('INCIDENT QUEUE (' + queued.length + ' QUEUED / ' + active.length + ' ACTIVE)', report, 'yellow');
    return;
  }

  // PURGE - clean old data + install daily trigger (admin roles only)
  if (mU === 'PURGE') {
    if (!isAdminRole()) {
      showAlert('ACCESS DENIED', 'PURGE COMMAND REQUIRES ADMIN LOGIN (SUPV/MGR/IT).');
      return;
    }
    setLive(true, 'LIVE • PURGE');
    const r = await API.runPurge(TOKEN);
    if (!r.ok) return showErr(r);
    showAlert('PURGE COMPLETE', r.message || ('DELETED ' + (r.deleted || 0) + ' OLD ROWS.'));
    return;
  }

  // INFO
  if (mU === 'INFO') {
    showAlert('SCMC HOSCAD — QUICK REFERENCE',
      'QUICK REFERENCE — MOST USED NUMBERS\n' +
      '═══════════════════════════════════════════════\n\n' +
      'DISPATCH CENTERS:\n' +
      '  DESCHUTES 911 NON-EMERG:  (541) 693-6911\n' +
      '  CROOK 911 NON-EMERG:      (541) 447-4168\n' +
      '  JEFFERSON NON-EMERG:      (541) 384-2080\n\n' +
      'AIR AMBULANCE:\n' +
      '  AIRLINK CCT:              1-800-621-5433\n' +
      '  LIFE FLIGHT NETWORK:      1-800-232-0911\n\n' +
      'CRISIS:\n' +
      '  988 SUICIDE/CRISIS:       988\n' +
      '  DESCHUTES CRISIS:         (541) 322-7500 X9\n\n' +
      'OTHER:\n' +
      '  POISON CONTROL:           1-800-222-1222\n' +
      '  OSP NON-EMERGENCY:        *677 (*OSP)\n' +
      '  ODOT ROAD CONDITIONS:     511\n\n' +
      '═══════════════════════════════════════════════\n' +
      'SUB-COMMANDS FOR DETAILED INFO:\n\n' +
      '  INFO DISPATCH    911/PSAP CENTERS\n' +
      '  INFO AIR         AIR AMBULANCE DISPATCH\n' +
      '  INFO OSP         OREGON STATE POLICE\n' +
      '  INFO CRISIS      MENTAL HEALTH / CRISIS\n' +
      '  INFO POISON      POISON CONTROL\n' +
      '  INFO ROAD        ROAD CONDITIONS / ODOT\n' +
      '  INFO LE          LAW ENFORCEMENT DIRECT\n' +
      '  INFO JAIL        JAILS\n' +
      '  INFO FIRE        FIRE DEPARTMENT ADMIN\n' +
      '  INFO ME          MEDICAL EXAMINER\n' +
      '  INFO OTHER       OTHER USEFUL NUMBERS\n' +
      '  INFO ALL         SHOW EVERYTHING\n' +
      '  INFO <UNIT>      DETAILED UNIT INFO\n');
    return;
  }

  // ADDR — Address directory / search
  if (mU === 'ADDR' || mU.startsWith('ADDR ')) {
    const addrQuery = mU === 'ADDR' ? '' : mU.substring(5).trim();
    if (!AddressLookup._loaded) {
      showAlert('ADDRESS DIRECTORY', 'ADDRESS DATA NOT YET LOADED. PLEASE TRY AGAIN.');
      return;
    }
    if (!addrQuery) {
      // Full directory grouped by category
      const cats = {};
      AddressLookup._cache.forEach(function(a) {
        const c = a.category || 'OTHER';
        if (!cats[c]) cats[c] = [];
        cats[c].push(a);
      });
      let out = 'ADDRESS DIRECTORY (' + AddressLookup._cache.length + ' ENTRIES)\n\n';
      Object.keys(cats).sort().forEach(function(c) {
        out += '═══ ' + c.replace(/_/g, ' ') + ' (' + cats[c].length + ') ═══\n';
        cats[c].forEach(function(a) {
          out += '  ' + a.id.padEnd(10) + a.name + '\n';
          out += '  ' + ''.padEnd(10) + a.address + ', ' + a.city + ', ' + a.state + ' ' + a.zip + '\n';
          if (a.phone) out += '  ' + ''.padEnd(10) + 'PH: ' + a.phone + '\n';
          if (a.notes) out += '  ' + ''.padEnd(10) + a.notes + '\n';
        });
        out += '\n';
      });
      showAlert('ADDRESS DIRECTORY', out);
    } else {
      const results = AddressLookup.search(addrQuery, 20);
      if (!results.length) {
        showAlert('ADDRESS SEARCH', 'NO RESULTS FOR: ' + addrQuery);
      } else {
        let out = 'ADDRESS SEARCH: ' + addrQuery + ' (' + results.length + ' RESULTS)\n\n';
        results.forEach(function(a) {
          out += '[' + a.id + '] ' + a.name + '\n';
          out += '  ' + a.address + ', ' + a.city + ', ' + a.state + ' ' + a.zip + '\n';
          out += '  CATEGORY: ' + (a.category || '').replace(/_/g, ' ');
          if (a.phone) out += '  |  PH: ' + a.phone;
          if (a.notes) out += '  |  ' + a.notes;
          out += '\n\n';
        });
        showAlert('ADDRESS SEARCH', out);
      }
    }
    return;
  }

  // STATUS
  if (mU === 'STATUS') {
    const r = await API.getSystemStatus(TOKEN);
    if (!r.ok) return showErr(r);
    const s = r.status;
    showConfirm('SYSTEM STATUS', 'SYSTEM STATUS\n\nUNITS: ' + s.totalUnits + ' TOTAL, ' + s.activeUnits + ' ACTIVE\n\nBY STATUS:\n  D:   ' + (s.byStatus.D || 0) + '\n  DE:  ' + (s.byStatus.DE || 0) + '\n  OS:  ' + (s.byStatus.OS || 0) + '\n  T:   ' + (s.byStatus.T || 0) + '\n  AV:  ' + (s.byStatus.AV || 0) + '\n  OOS: ' + (s.byStatus.OOS || 0) + '\n\nINCIDENTS:\n  ACTIVE: ' + s.activeIncidents + '\n  STALE:  ' + s.staleIncidents + '\n\nLOGGED IN AS: ' + s.actor, () => { });
    return;
  }

  // OKALL
  if (mU === 'OKALL') return okAllOS();

  // LO / LOGOUT — session logout (no args); LO <UNIT> routes to unit logoff below
  if (mU === 'LO' || mU === 'LOGOUT') {
    const ok = await showConfirmAsync('LOG OUT', 'LOG OUT OF HOSCAD?');
    if (!ok) return;
    const logoutResult = await API.logout(TOKEN);
    if (!logoutResult.ok) {
      showAlert('LOGOUT ERROR', logoutResult.error || 'FAILED TO LOG OUT. SESSION MAY STILL BE ACTIVE.');
    }
    localStorage.removeItem('ems_token');
    localStorage.removeItem('ems_role');
    localStorage.removeItem('ems_actor');
    TOKEN = '';
    ACTOR = '';
    document.getElementById('loginBack').style.display = 'flex';
    document.getElementById('userLabel').textContent = '—';
    if (POLL) clearInterval(POLL);
    stopLfnPolling();
    _rtDisconnect();
    return;
  }

  // OK - Touch unit or incident
  if (mU.startsWith('OK ')) {
    const re = ma.substring(3).trim().toUpperCase();
    if (re.startsWith('INC')) {
      const iId = re.replace(/^INC\s*/i, '');
      const r = await API.touchIncident(TOKEN, iId);
      if (!r.ok) return showErr(r);
      refresh();
      return;
    }
    const u = canonicalUnit(re);
    if (!u) { showConfirm('ERROR', 'USAGE: OK <UNIT> OR OK INC0001', () => { }); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === u) : null;
    if (!uO) { showConfirm('ERROR', 'UNIT NOT FOUND: ' + u, () => { }); return; }
    return okUnit(uO);
  }

  // NOTE/ALERT banners
  if (mU === 'NOTE') {
    setLive(true, 'LIVE • NOTE');
    const r = await API.setBanner(TOKEN, 'NOTE', nU || 'CLEAR');
    if (!r.ok) return showErr(r);
    beepNote();
    refresh();
    return;
  }

  if (mU === 'ALERT') {
    setLive(true, 'LIVE • ALERT');
    const r = await API.setBanner(TOKEN, 'ALERT', nU || 'CLEAR');
    if (!r.ok) return showErr(r);
    beepAlert();
    refresh();
    return;
  }

  // UI - Unit info modal
  if (mU.startsWith('UI ')) {
    const u = canonicalUnit(ma.substring(3).trim());
    if (!u) { showConfirm('ERROR', 'USAGE: UI <UNIT>', () => { }); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === u) : null;
    if (uO) openModal(uO, true);
    else openModal({ unit_id: u, display_name: displayNameForUnit(u), type: '', active: true, status: 'AV', note: '', unit_info: '', incident: '', destination: '', updated_at: '', updated_by: '' }, true);
    return;
  }

  // INFO for specific unit
  if (mU.startsWith('INFO ')) {
    const infoArg = mU.substring(5).trim();

    // INFO sub-commands for dispatch/emergency reference
    const INFO_SECTIONS = {
      'DISPATCH': {
        title: 'INFO — 911 / PSAP DISPATCH CENTERS',
        text:
          '911 / PSAP CENTERS (PUBLIC SAFETY ANSWERING POINTS)\n' +
          '═══════════════════════════════════════════════\n\n' +
          'DESCHUTES COUNTY 911\n' +
          '  NON-EMERGENCY:  (541) 693-6911\n' +
          '  ADMIN/BUSINESS: (541) 388-0185\n' +
          '  DISPATCHES FOR: BEND PD, REDMOND PD, DCSO,\n' +
          '    ALL DESCHUTES FIRE/EMS\n\n' +
          'CROOK COUNTY 911\n' +
          '  NON-EMERGENCY:  (541) 447-4168\n' +
          '  DISPATCHES FOR: PRINEVILLE PD, CCSO,\n' +
          '    CROOK COUNTY FIRE & RESCUE\n\n' +
          'JEFFERSON COUNTY DISPATCH\n' +
          '  NON-EMERGENCY:  (541) 384-2080\n' +
          '  ADMIN/BUSINESS: (541) 475-6520\n' +
          '  DISPATCHES FOR: JCSO, JEFFERSON COUNTY\n' +
          '    FIRE & EMS\n'
      },
      'AIR': {
        title: 'INFO — AIR AMBULANCE DISPATCH',
        text:
          'AIR AMBULANCE DISPATCH\n' +
          '═══════════════════════════════════════════════\n\n' +
          'AIRLINK CCT\n' +
          '  DISPATCH:  1-800-621-5433\n' +
          '  ALT:       (541) 280-3624\n' +
          '  BEND-BASED HELICOPTER (EC-135)\n' +
          '  & FIXED WING (PILATUS PC-12)\n\n' +
          'LIFE FLIGHT NETWORK\n' +
          '  DISPATCH:  1-800-232-0911\n' +
          '  REDMOND-BASED HELICOPTER (A-119)\n' +
          '  24/7 DISPATCH\n'
      },
      'OSP': {
        title: 'INFO — OREGON STATE POLICE',
        text:
          'OREGON STATE POLICE\n' +
          '═══════════════════════════════════════════════\n\n' +
          'NON-EMERGENCY:  *677 (*OSP) FROM CELL\n' +
          '  COVERS DESCHUTES, CROOK, JEFFERSON COUNTIES\n\n' +
          'TOLL-FREE:      1-800-452-7888\n' +
          '  NORTHERN COMMAND CENTER\n\n' +
          'DIRECT:         (503) 375-3555\n' +
          '  SALEM DISPATCH\n'
      },
      'CRISIS': {
        title: 'INFO — MENTAL HEALTH / CRISIS LINES',
        text:
          'MENTAL HEALTH / CRISIS LINES\n' +
          '═══════════════════════════════════════════════\n\n' +
          '988 SUICIDE & CRISIS LIFELINE\n' +
          '  CALL OR TEXT:  988\n' +
          '  24/7\n\n' +
          'DESCHUTES COUNTY CRISIS LINE\n' +
          '  (541) 322-7500 EXT. 9\n' +
          '  24/7\n\n' +
          'DESCHUTES STABILIZATION CENTER\n' +
          '  (541) 585-7210\n' +
          '  NON-EMERGENCY, WALK-IN 24/7\n\n' +
          'OREGON YOUTHLINE\n' +
          '  1-877-968-8491\n' +
          '  TEEN-TO-TEEN 4-10PM; ADULTS OTHER HOURS\n\n' +
          'VETERANS CRISIS LINE\n' +
          '  988, THEN PRESS 1\n\n' +
          'TRANS LIFELINE\n' +
          '  1-877-565-8860\n' +
          '  LIMITED HOURS\n\n' +
          'OREGON CRISIS TEXT LINE\n' +
          '  TEXT HOME TO 741741\n' +
          '  24/7\n'
      },
      'POISON': {
        title: 'INFO — POISON CONTROL',
        text:
          'POISON CONTROL\n' +
          '═══════════════════════════════════════════════\n\n' +
          'OREGON POISON CENTER\n' +
          '  1-800-222-1222\n' +
          '  24/7, MULTILINGUAL\n\n' +
          'POISONHELP.ORG\n' +
          '  ONLINE TOOL — NON-EMERGENCY\n'
      },
      'ROAD': {
        title: 'INFO — ROAD CONDITIONS / ODOT',
        text:
          'ROAD CONDITIONS / ODOT\n' +
          '═══════════════════════════════════════════════\n\n' +
          'TRIPCHECK 511\n' +
          '  511 FROM ANY PHONE IN OREGON\n\n' +
          'ODOT TOLL-FREE\n' +
          '  1-800-977-6368 (1-800-977-ODOT)\n\n' +
          'ODOT OUTSIDE OREGON\n' +
          '  (503) 588-2941\n\n' +
          'TRIPCHECK.COM\n' +
          '  LIVE CAMERAS, CONDITIONS\n'
      },
      'LE': {
        title: 'INFO — LAW ENFORCEMENT DIRECT LINES',
        text:
          'LAW ENFORCEMENT DIRECT LINES\n' +
          '═══════════════════════════════════════════════\n\n' +
          'DESCHUTES COUNTY SHERIFF   (541) 388-6655\n' +
          'CROOK COUNTY SHERIFF       (541) 447-6398\n' +
          'JEFFERSON COUNTY SHERIFF   (541) 475-6520\n' +
          'PRINEVILLE POLICE          (541) 447-4168\n' +
          '  (SHARES LINE WITH CROOK 911)\n' +
          'BEND POLICE ADMIN          (541) 322-2960\n' +
          'REDMOND POLICE             (541) 504-1810\n'
      },
      'JAIL': {
        title: 'INFO — JAILS',
        text:
          'JAILS — CONTROL ROOM NUMBERS\n' +
          '═══════════════════════════════════════════════\n\n' +
          'DESCHUTES COUNTY JAIL      (541) 388-6661\n' +
          'CROOK COUNTY JAIL          (541) 416-3620\n' +
          '  86 BEDS\n' +
          'JEFFERSON COUNTY JAIL      (541) 475-2869\n'
      },
      'FIRE': {
        title: 'INFO — FIRE DEPARTMENT ADMIN',
        text:
          'FIRE DEPARTMENT ADMIN\n' +
          '═══════════════════════════════════════════════\n\n' +
          'BEND FIRE & RESCUE         (541) 322-6300\n' +
          '  HQ: STATION 301\n' +
          'REDMOND FIRE & RESCUE      (541) 504-5000\n' +
          '  HQ: STATION 401\n' +
          'CROOK COUNTY FIRE & RESCUE (541) 447-5011\n' +
          '  HQ: PRINEVILLE\n' +
          'JEFFERSON COUNTY FIRE/EMS  (541) 475-7274\n' +
          '  HQ: MADRAS\n\n' +
          'BATTALION CHIEFS\n' +
          '═══════════════════════════════════════════════\n' +
          'BEND FIRE BC               TBD\n' +
          'REDMOND FIRE BC            TBD\n' +
          'CROOK COUNTY FIRE BC       TBD\n' +
          'JEFFERSON COUNTY FIRE BC   TBD\n'
      },
      'ME': {
        title: 'INFO — MEDICAL EXAMINER',
        text:
          'MEDICAL EXAMINER\n' +
          '═══════════════════════════════════════════════\n\n' +
          'DESCHUTES COUNTY ME\n' +
          '  MEDICAL.EXAMINER@DESCHUTES.ORG\n' +
          '  VIA DA\'S OFFICE\n\n' +
          'STATE MEDICAL EXAMINER\n' +
          '  (971) 673-8200\n' +
          '  CLACKAMAS (AUTOPSIES)\n'
      },
      'OTHER': {
        title: 'INFO — OTHER USEFUL NUMBERS',
        text:
          'OTHER USEFUL NUMBERS\n' +
          '═══════════════════════════════════════════════\n\n' +
          'DHS — ADULT PROTECTIVE SERVICES\n' +
          '  (541) 475-6773  (MADRAS)\n\n' +
          'DHS — DEVELOPMENTAL DISABILITIES\n' +
          '  (541) 322-7554  (BEND)\n\n' +
          'OUTDOOR BURN LINE (JEFFERSON CO)\n' +
          '  (541) 475-1789\n\n' +
          'COIDC (WILDFIRE DISPATCH)\n' +
          '  CENTRAL OREGON INTERAGENCY DISPATCH\n' +
          '  TBD\n'
      }
    };

    // Check for known sub-commands
    if (INFO_SECTIONS[infoArg]) {
      const sec = INFO_SECTIONS[infoArg];
      showAlert(sec.title, sec.text);
      return;
    }

    // INFO ALL — show everything
    if (infoArg === 'ALL') {
      let all = 'SCMC HOSCAD — COMPLETE REFERENCE DIRECTORY\n';
      all += '═══════════════════════════════════════════════\n\n';
      const order = ['DISPATCH', 'AIR', 'OSP', 'CRISIS', 'POISON', 'ROAD', 'LE', 'JAIL', 'FIRE', 'ME', 'OTHER'];
      order.forEach(function(k) {
        all += INFO_SECTIONS[k].text + '\n';
      });
      showAlert('INFO — COMPLETE DIRECTORY', all);
      return;
    }

    // Fall through to unit info lookup
    const u = canonicalUnit(ma.substring(5).trim());
    if (!u) { showConfirm('ERROR', 'USAGE: INFO <UNIT> OR INFO DISPATCH/AIR/CRISIS/LE/FIRE/JAIL/ALL', () => { }); return; }
    const r = await API.getUnitInfo(TOKEN, u);
    if (!r.ok) return showErr(r);
    if (!r.unit) {
      showErr({ error: 'UNIT ' + u + ' NOT FOUND IN SYSTEM.' });
      return;
    }
    const un = r.unit;
    const destR = AddressLookup.resolve(un.destination);
    const destDisplay = destR.recognized ? destR.addr.name + ' [' + destR.addr.id + ']' : (un.destination || '—');
    showConfirm('UNIT INFO: ' + un.unit_id, 'UNIT INFO: ' + un.unit_id + '\n\nDISPLAY: ' + (un.display_name || '—') + '\nTYPE: ' + (un.type || '—') + '\nSTATUS: ' + (un.status || '—') + '\nACTIVE: ' + (un.active ? 'YES' : 'NO') + '\n\nINCIDENT: ' + (un.incident || '—') + '\nDESTINATION: ' + destDisplay + '\nNOTE: ' + (un.note || '—') + '\n\nUNIT INFO:\n' + (un.unit_info || '(NONE)') + '\n\nUPDATED: ' + (un.updated_at || '—') + '\nBY: ' + (un.updated_by || '—'), () => { });
    return;
  }

  // R - Review incident
  if (mU.startsWith('R ')) {
    const iR = ma.substring(2).trim().toUpperCase();
    if (!iR) { showConfirm('ERROR', 'USAGE: R INC0001 OR R 0001', () => { }); return; }
    return openIncidentFromServer(iR);
  }

  // CB - Set/update callback number on incident
  // CB <INC> <PHONE>  — explicit incident ID
  // CB <PHONE>        — uses currently-open incident panel
  if (mU.startsWith('CB ')) {
    const cbArg = ma.substring(3).trim().toUpperCase();
    const cbPhone = nU.trim().toUpperCase();
    const isIncId = /^INC\d+$/.test(cbArg) || /^[A-Z]{1,4}\d{2}-\d{4}$/.test(cbArg) || /^\d{2}-\d{4}$/.test(cbArg) || /^\d{3,4}$/.test(cbArg);
    let cbIncId, cbNum;
    if (isIncId && cbPhone) {
      cbIncId = cbArg;
      cbNum = cbPhone;
    } else if (!isIncId && cbArg) {
      if (!CURRENT_INCIDENT_ID) { showAlert('ERROR', 'USAGE: CB <INC> <PHONE> — OR OPEN AN INCIDENT FIRST'); return; }
      cbIncId = CURRENT_INCIDENT_ID;
      cbNum = (cbArg + (cbPhone ? ' ' + cbPhone : '')).trim();
    } else {
      showAlert('ERROR', 'USAGE: CB <INC> <PHONE> — e.g. CB 0023 5415551234'); return;
    }
    if (!cbNum) { showAlert('ERROR', 'PHONE NUMBER REQUIRED'); return; }
    // Normalize cbIncId to yy-xxxx for STATE lookup
    let cbNorm = cbIncId;
    if (/^\d{3,4}$/.test(cbNorm)) {
      const yy = String(new Date().getFullYear()).slice(-2);
      cbNorm = yy + '-' + (cbNorm.length === 3 ? '0' + cbNorm : cbNorm);
    } else if (/^INC(\d+)$/.test(cbNorm)) {
      const d = cbNorm.replace(/^INC/, '');
      cbNorm = String(new Date().getFullYear()).slice(-2) + '-' + d.padStart(4, '0');
    }
    // Get current note from STATE (to preserve [DISP:] tag and strip old [CB:])
    const cbIncObj = STATE && STATE.incidents ? STATE.incidents.find(i => String(i.incident_id || '').toUpperCase() === cbNorm.toUpperCase()) : null;
    const rawCbNote = cbIncObj ? (cbIncObj.incident_note || '') : '';
    const cbDispM = rawCbNote.match(/\[DISP:([^\]]+)\]/i);
    const cbStripped = rawCbNote.replace(/\[DISP:[^\]]*\]\s*/gi, '').replace(/\[CB:[^\]]*\]\s*/gi, '').trim();
    const cbNewNote = [(cbDispM ? '[DISP:' + cbDispM[1].toUpperCase() + ']' : ''), cbStripped, '[CB:' + cbNum + ']'].filter(Boolean).join(' ');
    setLive(true, 'LIVE • SET CALLBACK');
    const r = await API.updateIncident(TOKEN, cbIncId, cbNewNote, undefined, undefined, undefined, undefined);
    if (!r.ok) return showErr(r);
    showToast('CB SET ON ' + cbNorm + ': ' + cbNum);
    refresh();
    return;
  }

  // U - Update incident note
  // U <INC> <NOTE>  — explicit incident ID
  // U <NOTE>        — uses currently-open incident panel (no ID required)
  if (mU.startsWith('U ')) {
    const iR = ma.substring(2).trim().toUpperCase();
    const isIncId = /^INC\d+$/.test(iR) || /^[A-Z]{1,4}\d{2}-\d{4}$/.test(iR) || /^\d{2}-\d{4}$/.test(iR) || /^\d{3,4}$/.test(iR);
    if (isIncId) {
      if (!nU) { showConfirm('ERROR', 'USAGE: U INC0001 MESSAGE', () => { }); return; }
      setLive(true, 'LIVE • ADD NOTE');
      const r = await API.appendIncidentNote(TOKEN, iR, nU);
      if (!r.ok) return showErr(r);
      refresh();
      return;
    } else {
      // No INC ID — use currently-open incident
      if (!CURRENT_INCIDENT_ID) { showAlert('ERROR', 'USAGE: U <INC> <NOTE> — OR OPEN AN INCIDENT FIRST'); return; }
      const fullNote = (iR + (nU ? ' ' + nU : '')).trim();
      if (!fullNote) { showAlert('ERROR', 'NOTE TEXT REQUIRED'); return; }
      setLive(true, 'LIVE • ADD NOTE');
      const r = await API.appendIncidentNote(TOKEN, CURRENT_INCIDENT_ID, fullNote);
      if (!r.ok) return showErr(r);
      showToast('NOTE ADDED TO ' + CURRENT_INCIDENT_ID);
      refresh();
      return;
    }
  }

  // COPY - Duplicate an existing incident into a new QUEUED call
  // COPY <INC>
  if (mU.startsWith('COPY ') || (mU === 'COPY' && !nU.trim())) {
    let copyId;
    if (mU === 'COPY' && !nU.trim()) {
      if (!CURRENT_INCIDENT_ID) { showAlert('ERROR', 'USAGE: COPY <INC> — OR OPEN AN INCIDENT FIRST'); return; }
      copyId = CURRENT_INCIDENT_ID;
    } else {
      copyId = (mU.startsWith('COPY ') ? ma.substring(5) : nU).trim().toUpperCase();
      if (!copyId) { showAlert('ERROR', 'USAGE: COPY <INC>'); return; }
    }
    // Normalize to yy-xxxx for STATE lookup
    let copyNorm = copyId;
    if (/^\d{3,4}$/.test(copyNorm)) {
      const yy = String(new Date().getFullYear()).slice(-2);
      copyNorm = yy + '-' + (copyNorm.length === 3 ? '0' + copyNorm : copyNorm);
    } else if (/^INC(\d+)$/.test(copyNorm)) {
      const d = copyNorm.replace(/^INC/, '');
      copyNorm = String(new Date().getFullYear()).slice(-2) + '-' + d.padStart(4, '0');
    }
    const srcInc = STATE && STATE.incidents ? STATE.incidents.find(i => String(i.incident_id || '').toUpperCase() === copyNorm.toUpperCase()) : null;
    if (!srcInc) { showAlert('ERROR', 'INCIDENT NOT FOUND: ' + copyNorm); return; }
    // Strip all bracket tags from note — keep human-readable text only
    const copyNote = (srcInc.incident_note || '').replace(/\[[^\]]*\]\s*/g, '').trim().toUpperCase();
    setLive(true, 'LIVE • COPY INCIDENT');
    const r = await API.createQueuedIncident(TOKEN, srcInc.destination || '', copyNote, srcInc.priority || '', '', srcInc.incident_type || '', srcInc.scene_address || '', srcInc.level_of_care || '');
    if (!r.ok) return showErr(r);
    showToast('COPIED → ' + (r.incidentId || ''));
    refresh();
    if (r.incidentId) setTimeout(() => openIncident(r.incidentId), 400);
    return;
  }

  // HOLD - Scheduled call in queue (creates QUEUED incident with [HOLD:HH:MM] tag)
  // HOLD <DEST>; <HH:MM>; [NOTE]; [TYPE]; [PRIORITY]
  if (mU.startsWith('HOLD ') || mU === 'HOLD') {
    const holdRaw = tx.substring(4).trim();
    if (!holdRaw) { showAlert('ERROR', 'USAGE: HOLD <DEST>; <HH:MM>; [NOTE]; [TYPE]; [PRIORITY]\nExample: HOLD SCMC; 16:30; DIALYSIS TRANSPORT; EMS-ROUTINE'); return; }
    const holdParts = holdRaw.split(';').map(p => p.trim().toUpperCase());
    const hDest = holdParts[0] || '';
    const hTimeRaw = (holdParts[1] || '').trim();
    const hNote = holdParts[2] || '';
    const hType = holdParts[3] || '';
    const hPri = holdParts[4] || '';
    if (!hDest) { showAlert('ERROR', 'HOLD: DESTINATION REQUIRED'); return; }
    const timeMatch = hTimeRaw.match(/^(\d{1,2}):(\d{2})$/);
    if (!timeMatch) { showAlert('ERROR', 'HOLD: TIME MUST BE HH:MM (24-HOUR) — e.g. 16:30'); return; }
    const hh = String(parseInt(timeMatch[1])).padStart(2, '0');
    const mm = String(parseInt(timeMatch[2])).padStart(2, '0');
    if (parseInt(hh) > 23 || parseInt(mm) > 59) { showAlert('ERROR', 'HOLD: INVALID TIME — must be 00:00–23:59'); return; }
    const holdTime = hh + ':' + mm;
    const fullNote = ('[HOLD:' + holdTime + ']' + (hNote ? ' ' + hNote : '')).trim();
    setLive(true, 'LIVE • CREATE HOLD');
    const r = await API.createQueuedIncident(TOKEN, hDest, fullNote, hPri || '', '', hType, '');
    if (!r.ok) return showErr(r);
    showToast('HOLD CALL CREATED FOR ' + holdTime + ' → ' + hDest);
    refresh();
    autoFocusCmd();
    return;
  }

  // NC - New incident in queue
  if (mU.startsWith('NC ') || mU === 'NC') {
    const ncRaw = tx.substring(2).trim();
    if (!ncRaw) { openNewIncident(); return; }
    const ncParts = ncRaw.split(';').map(p => p.trim().toUpperCase());
    const dest     = ncParts[0] || '';
    let   noteRaw  = ncParts[1] || '';
    const incType  = ncParts[2] || '';
    const priority = ncParts[3] || '';
    // Scene address: 5th segment, OR any segment prefixed with @ (e.g. @1234 MAIN ST)
    let sceneAddress = '';
    if (ncParts[4] && ncParts[4].startsWith('@')) {
      sceneAddress = ncParts[4].substring(1).trim();
    } else if (ncParts[4]) {
      sceneAddress = ncParts[4].trim();
    }
    // Also scan all segments for @-prefixed token in case dispatcher puts it elsewhere
    for (let _si = 1; _si < ncParts.length; _si++) {
      if (ncParts[_si].startsWith('@') && _si !== 4) {
        sceneAddress = ncParts[_si].substring(1).trim();
        // Remove it from noteRaw if it appeared in position 1
        if (_si === 1) noteRaw = '';
        break;
      }
    }
    if (!dest) { showAlert('ERROR', 'USAGE: NC <LOCATION>; <NOTE>; <TYPE>; <PRIORITY>; @<SCENE ADDR>'); return; }
    // MA token in note
    const isMa = /\bMA\b/.test(noteRaw.toUpperCase());
    let note = noteRaw.replace(/\bMA\b\s*/gi, '').trim();
    const prefixes = [];
    if (isMa) prefixes.push('[MA]');
    if (prefixes.length) note = prefixes.join(' ') + (note ? ' ' + note : '');
    setLive(true, 'LIVE • CREATE INCIDENT');
    const r = await API.createQueuedIncident(TOKEN, dest, note, priority || '', '', incType, sceneAddress);
    if (!r.ok) return showErr(r);
    if (sceneAddress) { AddrHistory.push(sceneAddress); _geoVerifyAddress(sceneAddress); }
    refresh();
    autoFocusCmd();
    return;
  }

  // ETA <UNIT> CLR — clear ETA tag
  const etaClrMatch = (mU + ' ' + nU).trim().match(/^ETA\s+(\S+)\s+(CLR|CLEAR)$/i);
  if (etaClrMatch) {
    const etaClrUnit = etaClrMatch[1].toUpperCase();
    const unitObj = (STATE && STATE.units || []).find(u => String(u.unit_id || '').toUpperCase() === etaClrUnit);
    if (!unitObj) { showAlert('ERROR', 'UNIT NOT FOUND: ' + etaClrUnit); return; }
    const clearedNote = (unitObj.note || '').replace(/\[ETA:\d+\]\s*/gi, '').trim();
    setLive(true, 'LIVE • CLEAR ETA');
    const r = await API.upsertUnit(TOKEN, etaClrUnit, { note: clearedNote });
    if (!r.ok) return showErr(r);
    showToast('ETA CLEARED: ' + etaClrUnit);
    refresh();
    return;
  }

  // ETA <UNIT> <MINUTES>
  const etaMatch = (mU + ' ' + nU).trim().match(/^ETA\s+(\S+)\s+(\d+)$/);
  if (etaMatch) {
    const etaUnitId = etaMatch[1].toUpperCase();
    const etaMins = etaMatch[2];
    setLive(true, 'LIVE • SET ETA');
    const r = await API.setUnitETA(TOKEN, etaUnitId, etaMins);
    if (!r.ok) return showErr(r);
    showToast('ETA ' + etaMins + 'M SET FOR ' + etaUnitId);
    refresh();
    return;
  }

  // NEXT <UNIT> — advance unit one step in EMS status chain: D→DE→OS→T→TH→AV
  if (mU === 'NEXT') {
    const nextRaw = (nU || '').trim().toUpperCase();
    const nextUnit = canonicalUnit(nextRaw);
    if (!nextUnit) { showAlert('ERROR', 'USAGE: NEXT <UNIT>\nAdvances unit one step: D→DE→OS→T→TH→AV'); return; }
    const NEXT_CHAIN = { D: 'DE', DE: 'OS', OS: 'T', T: 'TH', TH: 'AV' };
    const unitObj = (STATE && STATE.units || []).find(u => String(u.unit_id || '').toUpperCase() === nextUnit);
    if (!unitObj) { showAlert('ERROR', 'UNIT NOT FOUND: ' + nextUnit); return; }
    const curSt = (unitObj.status || '').toUpperCase();
    const nextSt = NEXT_CHAIN[curSt];
    if (!nextSt) { showAlert('ERROR', nextUnit + ' IS ' + curSt + ' — NO NEXT STEP IN EMS CHAIN.\nChain: D→DE→OS→T→TH→AV'); return; }
    const ok = await showConfirmAsync('NEXT: ' + nextUnit, 'Advance ' + nextUnit + ' from ' + curSt + ' → ' + nextSt + '?');
    if (!ok) return;
    setLive(true, 'LIVE • NEXT STATUS');
    const r = await API.upsertUnit(TOKEN, nextUnit, { status: nextSt });
    if (!r.ok) return showErr(r);
    showToast(nextUnit + ': ' + curSt + ' → ' + nextSt);
    refresh();
    return;
  }

  // PAT <UNIT> <TEXT>  or  PAT <UNIT> CLR
  const patCmdMatch = (mU + ' ' + nU).trim().match(/^PAT\s+(\S+)\s+(.+)$/i);
  if (patCmdMatch) {
    const patUnitId = patCmdMatch[1].toUpperCase();
    const patText = patCmdMatch[2].trim().toUpperCase();
    const isClear = patText === 'CLR' || patText === 'CLEAR';
    setLive(true, 'LIVE • SET PAT INFO');
    const r = await API.setUnitPAT(TOKEN, patUnitId, isClear ? '' : patText);
    if (!r.ok) return showErr(r);
    showToast(isClear ? 'PAT INFO CLEARED: ' + patUnitId : 'PAT SET: ' + patUnitId + ' — ' + patText);
    refresh();
    return;
  }

  // PRIORITY <INC> <PRI-N>
  const priMatch = (mU + ' ' + nU).trim().match(/^PRIORITY\s+(\S+)\s+(PRI-[1-4])$/);
  if (priMatch) {
    let priIncId = priMatch[1].toUpperCase().replace(/^INC/i, '');
    if (/^\d{3}$/.test(priIncId)) priIncId = '0' + priIncId;
    const pri = priMatch[2].toUpperCase();
    setLive(true, 'LIVE • SET PRIORITY');
    const r = await API.setIncidentPriority(TOKEN, priIncId, pri);
    if (!r.ok) return showErr(r);
    showToast('PRIORITY UPDATED: ' + priIncId + ' → ' + pri);
    refresh();
    return;
  }

  // STATS - Live board summary
  if (mU === 'STATS') {
    const statsUnits = (STATE && STATE.units || []).filter(u => u.active);
    const statsIncidents = (STATE && STATE.incidents || []);
    const byStatus = {};
    ['AV','D','DE','OS','T','TH','OOS','BRK','IQ','UV','F','FD'].forEach(s => byStatus[s] = 0);
    statsUnits.forEach(u => { const s = (u.status||'').toUpperCase(); if (byStatus[s] !== undefined) byStatus[s]++; });
    const activeInc = statsIncidents.filter(i => i.status === 'ACTIVE');
    const queuedInc = statsIncidents.filter(i => i.status === 'QUEUED');
    const now = Date.now();
    let longestId = null, longestMins = 0;
    activeInc.forEach(i => {
      const m = Math.floor((now - new Date(_normalizeTs(i.created_at)).getTime()) / 60000);
      if (m > longestMins) { longestMins = m; longestId = i.incident_id; }
    });
    const fmtLong = (m) => m >= 60 ? Math.floor(m/60) + 'H ' + (m%60) + 'M' : m + 'M';
    const showIf = (label, count) => count > 0 ? (label + ' ' + count) : null;
    const avCount = byStatus['AV'] || 0;
    const coverageLabel = avCount === 0 ? 'CRITICAL' : avCount === 1 ? 'LIMITED' : avCount <= 3 ? 'REDUCED' : 'NORMAL';
    const lines = [
      'BOARD STATUS SUMMARY',
      '═'.repeat(36),
      'COVERAGE:     ' + coverageLabel + ' (' + avCount + ' AV)',
      'INCIDENTS — ACTIVE: ' + activeInc.length + '  QUEUED: ' + queuedInc.length,
      '═'.repeat(36),
      'AVAILABLE:    ' + String(avCount).padStart(3),
      showIf('DISPATCHED:   ', byStatus['D']),
      showIf('EN ROUTE:     ', byStatus['DE']),
      'ON SCENE:     ' + String(byStatus['OS']).padStart(3),
      'TRANSPORT:    ' + String(byStatus['T']).padStart(3),
      showIf('AT HOSPITAL:  ', byStatus['TH']),
      showIf('IN QUARTERS:  ', byStatus['IQ']),
      'OOS:          ' + String(byStatus['OOS']).padStart(3),
      showIf('ON BREAK:     ', byStatus['BRK']),
      '═'.repeat(36),
      longestId ? 'LONGEST INC:  ' + String(longestId).replace(/^[A-Z]*\d{2}-0*/,'') + ' (' + fmtLong(longestMins) + ')' : 'NO ACTIVE INCIDENTS',
      'TOTAL UNITS:  ' + statsUnits.length,
    ].filter(l => l !== null);
    showAlert('BOARD STATS', lines.join('\n'));
    return;
  }

  // SHIFT END <UNIT>
  const _shiftFull = (mU + ' ' + nU).trim();
  const shiftEndMatch = _shiftFull.match(/^SHIFT\s+END\s+(\S+)$/);
  if (shiftEndMatch) {
    const shiftEndUnit = shiftEndMatch[1].toUpperCase();
    const confirmed = await showConfirmAsync('SHIFT END: ' + shiftEndUnit + '?', 'Set AV, clear assignments, then deactivate ' + shiftEndUnit + '?');
    if (!confirmed) return;
    setLive(true, 'LIVE • SHIFT END');
    const r = await API.ridoffUnit(TOKEN, shiftEndUnit, '');
    if (!r.ok) { setLive(false); return showErr(r); }
    const r2 = await API.logoffUnit(TOKEN, shiftEndUnit, '');
    setLive(false);
    if (!r2.ok) { showToast('RIDOFF OK, LOGOFF FAILED: ' + r2.error); }
    else showToast('SHIFT END COMPLETE: ' + shiftEndUnit + ' DEACTIVATED');
    refresh();
    return;
  }

  // LINK - Link two units to incident
  if (mU.startsWith('LINK ')) {
    const ps = (ma + ' ' + no).trim().substring(5).trim().split(/\s+/);
    if (ps.length < 3) { showConfirm('ERROR', 'USAGE: LINK UNIT1 UNIT2 INC0001', () => { }); return; }
    const inc = ps[ps.length - 1].toUpperCase();
    const u2R = ps[ps.length - 2];
    const u1R = ps.slice(0, -2).join(' ');
    const u1 = canonicalUnit(u1R);
    const u2 = canonicalUnit(u2R);
    if (!u1 || !u2) { showConfirm('ERROR', 'USAGE: LINK UNIT1 UNIT2 INC0001', () => { }); return; }
    const r = await API.linkUnits(TOKEN, u1, u2, inc);
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  // TRANSFER
  if (mU.startsWith('TRANSFER ')) {
    const ps = (ma + ' ' + no).trim().substring(9).trim().split(/\s+/);
    if (ps.length < 3) { showConfirm('ERROR', 'USAGE: TRANSFER UNIT1 UNIT2 INC0001', () => { }); return; }
    const inc = ps[ps.length - 1].toUpperCase();
    const u2R = ps[ps.length - 2];
    const u1R = ps.slice(0, -2).join(' ');
    const u1 = canonicalUnit(u1R);
    const u2 = canonicalUnit(u2R);
    if (!u1 || !u2) { showConfirm('ERROR', 'USAGE: TRANSFER UNIT1 UNIT2 INC0001', () => { }); return; }
    const r = await API.transferIncident(TOKEN, u1, u2, inc);
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  // CAN <inc> [note] — quick cancel with optional note (no picker)
  {
    const canFull = (mU + ' ' + nU).trim();
    const canMatch = canFull.match(/^CAN\s+(?:INC\s*)?([A-Z]*\d{2}-\d{4}|\d{3,4})(?:\s+(.+))?$/) ||
                     canFull.match(/^(?:INC\s*)?([A-Z]*\d{2}-\d{4}|\d{3,4})\s+CAN$/);
    if (canMatch) {
      let inc = canMatch[1];
      const note = (canMatch[2] || '').trim();
      if (/^\d{3,4}$/.test(inc)) {
        if (inc.length === 3) inc = '0' + inc;
        const yy = String(new Date().getFullYear()).slice(-2);
        inc = 'SC' + yy + '-' + inc;
      }
      if (note) await API.appendIncidentNote(TOKEN, inc, note);
      const r = await API.closeIncident(TOKEN, inc, 'CANCELLED-PRIOR');
      if (!r.ok) return showErr(r);
      showToast('INC ' + inc.replace(/^[A-Z]*\d{2}-0*/, '') + ' CANCELLED');
      refresh();
      return;
    }
  }

  // DEL <inc> [dup|err] — close with structured reason; prompts if no reason given
  {
    const delFull = (mU + ' ' + nU).trim();
    const delMatch = delFull.match(/^DEL\s+(?:INC\s*)?([A-Z]*\d{2}-\d{4}|\d{3,4})(?:\s+(\S+))?$/) ||
                     delFull.match(/^(?:INC\s*)?([A-Z]*\d{2}-\d{4}|\d{3,4})\s+DEL$/) ||
                     delFull.match(/^DEL([A-Z]*\d{2}-\d{4}|\d{3,4})$/);
    // Don't match DEL MSG... (handled by message handler below)
    if (delMatch && !/^MSG/i.test(delMatch[1])) {
      let inc = delMatch[1];
      const reasonRaw = (delMatch[2] || '').toUpperCase();
      if (/^\d{3,4}$/.test(inc)) {
        if (inc.length === 3) inc = '0' + inc;
        const yy = String(new Date().getFullYear()).slice(-2);
        inc = 'SC' + yy + '-' + inc;
      }
      let disposition = '';
      if (reasonRaw === 'DUP' || reasonRaw === 'DUPLICATE') {
        disposition = 'DUPLICATE';
      } else if (reasonRaw === 'ERR' || reasonRaw === 'ERROR') {
        disposition = 'DATA-ERROR';
      } else {
        disposition = await promptDisposition(inc);
        if (!disposition) return;
      }
      const r = await API.closeIncident(TOKEN, inc, disposition);
      if (!r.ok) return showErr(r);
      showToast('INC ' + inc.replace(/^[A-Z]*\d{2}-0*/, '') + ' DEL — ' + disposition);
      refresh();
      return;
    }
  }

  // CLOSE (no ID) — dismiss open incident panel
  if (mU === 'CLOSE' && !nU.trim()) {
    if (CURRENT_INCIDENT_ID) { closeIncidentPanel(); }
    else { showAlert('INFO', 'NO INCIDENT OPEN'); }
    return;
  }

  // CLOSE incident
  if (mU.startsWith('CLOSE ')) {
    const inc = ma.substring(6).trim().toUpperCase();
    if (!inc) { showConfirm('ERROR', 'USAGE: CLOSE 0001 OR DEL 023 OR CAN 023', () => { }); return; }
    // If shortcode provided in nU, skip picker
    const closeShort = nU.trim().toUpperCase();
    if (closeShort && DISPO_SHORT_CODES[closeShort]) {
      const dispo = DISPO_SHORT_CODES[closeShort];
      const r = await API.closeIncident(TOKEN, inc, dispo);
      if (!r.ok) return showErr(r);
      showToast('INC ' + inc.replace(/^[A-Z]*\d{2}-0*/, '') + ' CLOSED — ' + dispo);
      refresh();
      return;
    }
    const disposition = await promptDisposition(inc);
    if (!disposition) return; // user cancelled
    const r = await API.closeIncident(TOKEN, inc, disposition);
    if (!r.ok) return showErr(r);
    showToast('INC ' + inc.replace(/^[A-Z]*\d{2}-0*/, '') + ' CLOSED — ' + disposition);
    refresh();
    return;
  }

      if (mU === 'BOARDS' || mU === 'BOARD') {
        openPopouts();
        return;
      }

  // REL - Link two HOSCAD incidents together (or unlink with CLR flag)
  if (mU.startsWith('REL ') || mU === 'REL') {
    const parts = (mU + ' ' + nU).trim().split(/\s+/);
    let relIncId = (parts[1] || '').toUpperCase();
    let relTarget = (parts[2] || '').toUpperCase();
    const relFlag = (parts[3] || '').toUpperCase();
    if (!relIncId || !relTarget) { showAlert('ERROR', 'USAGE: REL <INC1> <INC2> [CLR]  — LINK TWO INCIDENTS'); return; }
    // Normalize incident IDs: add year prefix if missing (e.g. 0023 → 26-0023)
    if (/^\d+$/.test(relIncId)) relIncId = (new Date().getFullYear() % 100) + '-' + relIncId;
    if (relIncId.startsWith('INC')) relIncId = relIncId.replace(/^INC/i, '');
    if (/^\d+$/.test(relTarget)) relTarget = (new Date().getFullYear() % 100) + '-' + relTarget;
    if (relTarget.startsWith('INC')) relTarget = relTarget.replace(/^INC/i, '');
    const unlink = relFlag === 'CLR' || relFlag === 'CLEAR';
    const relRes = await API.linkIncidents(TOKEN, relIncId, relTarget, unlink ? 'UNLINK' : '');
    if (!relRes.ok) { showAlert('ERROR', relRes.error || 'ERROR.'); return; }
    showToast(unlink ? 'LINK REMOVED: ' + relIncId + ' ↔ ' + relTarget : 'LINKED: ' + relIncId + ' ↔ ' + relTarget, 'success');
    refresh();
    return;
  }

  // RQ - Requeue incident (back to QUEUED, clears unit assignment)
  if (mU.startsWith('RQ ')) {
    const inc = ma.substring(3).trim().toUpperCase();
    if (!inc) { showConfirm('ERROR', 'USAGE: RQ INC0001', () => { }); return; }
    showConfirm('REQUEUE INCIDENT', `REQUEUE INC ${inc}?\n\nThis clears the current unit assignment and sets the incident back to QUEUED for reassignment.`, async () => {
      const r = await API.requeueIncident(TOKEN, inc);
      if (!r.ok) return showErr(r);
      refresh();
    });
    return;
  }

  // RO - Reopen incident (CLOSED → ACTIVE, keeps existing units)
  if (mU.startsWith('RO ')) {
    const inc = ma.substring(3).trim().toUpperCase();
    if (!inc) { showConfirm('ERROR', 'USAGE: RO INC0001', () => { }); return; }
    const r = await API.reopenIncident(TOKEN, inc);
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  // MASS D - Mass dispatch
  if (mU.startsWith('MASS D ') || (mU === 'MASS D' && nU)) {
    const massRaw = (ma.substring(7).trim() + (nU ? ' ' + nU : '')).trim().toUpperCase();
    const massParts = massRaw.split(/\s+/);
    const massHasConfirm = massParts[massParts.length - 1] === 'CONFIRM';
    const de = massHasConfirm ? massParts.slice(0, -1).join(' ') : massRaw;
    if (!de) { showConfirm('ERROR', 'USAGE: MASS D <DESTINATION> CONFIRM', () => { }); return; }
    if (!massHasConfirm) {
      showErr({ error: 'CONFIRMATION REQUIRED. RE-RUN WITH CONFIRM. EXAMPLE: MASS D ' + de + ' CONFIRM' });
      return;
    }
    showConfirm('CONFIRM MASS DISPATCH', 'MASS DISPATCH ALL AV UNITS TO ' + de + '?', async () => {
      const r = await API.massDispatch(TOKEN, de);
      if (!r.ok) return showErr(r);
      const ct = (r.updated || []).length;
      showConfirm('MASS DISPATCH COMPLETE', 'MASS DISPATCH: ' + ct + ' UNITS DISPATCHED TO ' + de + '\n\n' + (r.updated || []).join(', '), () => { });
      refresh();
    });
    return;
  }

  // AVALL / OSALL — bulk status update for all units on a specific incident
  if (mU.startsWith('AVALL ') || mU.startsWith('OSALL ') || mU === 'AVALL' || mU === 'OSALL') {
    const isAv = mU.startsWith('AVALL') || mU === 'AVALL';
    const newStatus = isAv ? 'AV' : 'OS';
    const cmdLen = isAv ? 6 : 6; // 'AVALL ' or 'OSALL ' both 6 chars
    const rawInc = (mU.length > 6 ? mU.substring(cmdLen).trim() : '') || nU.trim();
    if (!rawInc) { showAlert('ERROR', 'USAGE: ' + (isAv ? 'AVALL' : 'OSALL') + ' <INC#>  (e.g. ' + (isAv ? 'AVALL' : 'OSALL') + ' 0071)'); return; }
    // Resolve 4-digit shorthand or full incident ID
    const _stripped = rawInc.replace(/^INC[-\s]*/i, '').trim().toUpperCase();
    let _resolvedInc = _stripped;
    if (/^\d{4}$/.test(_stripped)) {
      const _found = (STATE.incidents || []).find(i => i.incident_id.endsWith('-' + _stripped));
      if (_found) _resolvedInc = _found.incident_id;
      else { const yy = String(new Date().getFullYear()).slice(-2); _resolvedInc = 'SC' + yy + '-' + _stripped; }
    }
    const incObj = (STATE.incidents || []).find(i => i.incident_id === _resolvedInc);
    if (!incObj) { showAlert('ERROR', 'INCIDENT NOT FOUND: ' + _resolvedInc); return; }
    const assigned = (STATE.units || []).filter(u => u.active && u.incident === _resolvedInc);
    if (!assigned.length) { showAlert('ERROR', 'NO UNITS ASSIGNED TO ' + _resolvedInc); return; }
    const unitNames = assigned.map(u => u.unit_id).join(', ');
    const ok = await showConfirmAsync(
      (isAv ? 'AVALL' : 'OSALL') + ': INC ' + _resolvedInc.replace(/^[A-Z]*\d{2}-0*/, ''),
      'Set ' + assigned.length + ' unit(s) → ' + newStatus + ':\n\n' + unitNames
    );
    if (!ok) return;
    setLive(true, 'LIVE \u2022 ' + newStatus + ' ALL');
    const results = await Promise.all(assigned.map(u => {
      const patch = { status: newStatus, displayName: u.display_name };
      if (!isAv) patch.incident = _resolvedInc; // OS keeps incident assignment
      return API.upsertUnit(TOKEN, u.unit_id, patch, '');
    }));
    setLive(false);
    const failed = results.filter(r => !r.ok);
    if (failed.length) {
      showToast((assigned.length - failed.length) + '/' + assigned.length + ' UPDATED — ' + failed.length + ' FAILED', 'error', 6000);
    } else {
      showToast(assigned.length + ' UNIT' + (assigned.length > 1 ? 'S' : '') + ' → ' + newStatus + '  INC ' + _resolvedInc.replace(/^[A-Z]*\d{2}-0*/, ''), 'success');
    }
    refresh();
    return;
  }

  // UH / HIST - Unit history
  if (mU.startsWith('UH ') || mU.startsWith('HIST ')) {
    const ps = (ma + ' ' + no).trim().split(/\s+/);
    let hr = 12;
    const la = ps[ps.length - 1];
    if (/^\d+$/.test(la)) { hr = Number(la); ps.pop(); }
    const uR = ps.slice(1).join(' ').trim();
    const u = canonicalUnit(uR);
    if (!u) { showConfirm('ERROR', 'USAGE: UH <UNIT> [12|24|48|168]', () => { }); return; }
    return openHistory(u, hr);
  }

  // Alternate UH syntax: EMS1 UH 12
  {
    const ps = (ma + ' ' + no).trim().split(/\s+/).filter(Boolean);
    if (ps.length >= 2 && (ps[1].toUpperCase() === 'UH' || ps[1].toUpperCase() === 'HIST')) {
      let hr = 12;
      const la = ps[ps.length - 1];
      const hH = /^\d+$/.test(la);
      if (hH) hr = Number(la);
      const en = hH ? ps.length - 1 : ps.length;
      const uR = ps.slice(0, en).filter((x, i) => i !== 1).join(' ');
      const u = canonicalUnit(uR);
      if (!u) { showConfirm('ERROR', 'USAGE: <UNIT> UH [12|24|48|168]', () => { }); return; }
      return openHistory(u, hr);
    }
  }

  // UNDO
  if (mU.startsWith('UNDO ')) {
    const u = canonicalUnit(ma.substring(5).trim());
    if (!u) { showConfirm('ERROR', 'USAGE: UNDO <UNIT>', () => { }); return; }
    return undoUnit(u);
  }

  // L / LOGON — logon unit
  if (mU.startsWith('L ') || mU.startsWith('LOGON ')) {
    const u = canonicalUnit(ma.substring(mU.startsWith('L ') ? 2 : 6).trim());
    if (!u) { showConfirm('ERROR', 'USAGE: LOGON <UNIT> <NOTE>', () => { }); return; }
    // Check everSeen — same barrier as the modal
    setLive(true, 'LIVE • CHECK UNIT');
    const info = await API.getUnitInfo(TOKEN, u);
    if (info.ok && !info.everSeen) {
      const similar = findSimilarUnits(u);
      let msg = '"' + u + '" HAS NEVER BEEN LOGGED ON BEFORE.\nIS THIS A NEW UNIT, OR A TYPO / DUPLICATE?';
      if (similar.length) msg += '\n\nSIMILAR KNOWN UNITS: ' + similar.join(', ');
      // Show dialog: [BACK] cancels, [LOG ON NEW UNIT] opens modal pre-filled
      const choice = await showNewUnitDialog(u, msg, nU);
      autoFocusCmd();
      if (choice === 'logon') {
        const dN = displayNameForUnit(u);
        const fakeUnit = {
          unit_id: u, display_name: dN, type: '', active: true, status: 'AV',
          note: nU || '', unit_info: '', incident: '', destination: '',
          updated_at: '', updated_by: ''
        };
        openModal(fakeUnit);
      }
      return;
    }
    // Roster unit (first logon, or level back-filled from roster): open modal pre-filled so dispatcher can confirm
    if (info.ok && info.unit && (info.unit.updated_at === null || info.levelFromRoster)) {
      openModal(Object.assign({}, info.unit, { active: true, status: 'AV', note: nU || '' }));
      return;
    }
    const dN = displayNameForUnit(u);
    setLive(true, 'LIVE • LOGON');
    const r = await API.upsertUnit(TOKEN, u, { active: true, status: 'AV', note: nU, displayName: dN }, '');
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  // LO <UNIT> / LOGOFF — logoff unit from board
  if (mU.startsWith('LO ') || mU.startsWith('LOGOFF ')) {
    const u = canonicalUnit(ma.substring(mU.startsWith('LO ') ? 3 : 7).trim());
    if (!u) { showConfirm('ERROR', 'USAGE: LOGOFF <UNIT>', () => { }); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === u) : null;
    const currentStatus = uO ? uO.status : '';
    const needsConfirm = ['OS', 'T', 'D', 'DE'].includes(currentStatus);
    const doLogoff = async () => {
      setLive(true, 'LIVE • LOGOFF');
      const r = await API.logoffUnit(TOKEN, u, '');
      if (!r.ok) return showErr(r);
      showToast('LOGGED OFF: ' + u);
      refresh();
    };
    if (needsConfirm) {
      showConfirm('CONFIRM LOGOFF', 'LOGOFF ' + u + ' (CURRENTLY ' + currentStatus + ')?', doLogoff);
    } else {
      await doLogoff();
    }
    return;
  }

  // RIDOFF
  if (mU.startsWith('RIDOFF ')) {
    const u = canonicalUnit(ma.substring(7).trim());
    if (!u) { showConfirm('ERROR', 'USAGE: RIDOFF <UNIT>', () => { }); return; }
    showConfirm('CONFIRM RIDOFF', 'RIDOFF ' + u + '? (SETS AV + CLEARS NOTE/INC/DEST)', async () => {
      setLive(true, 'LIVE • RIDOFF');
      const r = await API.ridoffUnit(TOKEN, u, '');
      if (!r.ok) return showErr(r);
      refresh();
    });
    return;
  }

  // DEST <UNIT>; <LOCATION> — set unit destination
  if (mU.startsWith('DEST ')) {
    const uRaw = ma.substring(5).trim();
    const u = canonicalUnit(uRaw);
    if (!u) { showAlert('ERROR', 'USAGE: DEST <UNIT> <LOCATION>\nDEST <UNIT> (CLEAR DESTINATION)'); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === u) : null;
    if (!uO) { showAlert('ERROR', 'UNIT NOT FOUND: ' + u); return; }
    let destVal = (nU || '').trim().toUpperCase();
    if (destVal) {
      // Parse bracket note: SCMC [BED 4] → base=SCMC, note=BED 4
      const { base: destBase, note: destNote } = AddressLookup._parseBracketNote(destVal);
      const lookupVal = destBase || destVal;
      // Try to resolve to a known address ID
      const byId = AddressLookup.getById(lookupVal);
      if (byId) {
        destVal = byId.id + (destNote ? ' [' + destNote + ']' : '');
      } else {
        const results = AddressLookup.search(lookupVal, 3);
        if (results.length === 1) destVal = results[0].id + (destNote ? ' [' + destNote + ']' : '');
        else if (destNote) destVal = lookupVal + ' [' + destNote + ']';
      }
    }
    setLive(true, 'LIVE • SET DEST');
    const r = await API.upsertUnit(TOKEN, u, { destination: destVal, displayName: uO.display_name }, uO.updated_at || '');
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  // LOC <UNIT> <LOCATION> — set unit location for map + board
  if (mU === 'LOC') {
    if (!nU) { showAlert('ERROR', 'USAGE: LOC <UNIT> <ADDR>\nLOC <UNIT> CLR — clear location'); return; }
    // LOC <CALLSIGN> — single token matching an air unit: show bearing + distance info
    const locSingle = nU.trim().toUpperCase();
    if (!locSingle.includes(' ')) {
      const locAcPair = Object.entries(AIR_FLEET).find(([tail, v]) => v.callsign === locSingle || tail === locSingle);
      if (locAcPair) {
        const locAc = _lfnAircraft.find(a => a.tail === locAcPair[0]);
        if (!locAc || locAc.status === 'NOSIG' || locAc.lat == null) {
          showToast('NO POSITION \u2014 ' + locSingle + ' HAS NO ADS-B SIGNAL', 'warn', 5000);
        } else {
          const locDist = _lfnDistNm(locAc.lat, locAc.lon, LFN_BASE_LAT, LFN_BASE_LON);
          const locBear = _lfnBearing(LFN_BASE_LAT, LFN_BASE_LON, locAc.lat, locAc.lon);
          const altStr  = locAc.alt_ft  != null ? ' \u00b7 ' + locAc.alt_ft.toLocaleString() + 'ft' : '';
          const spdStr  = locAc.gs_kts  != null ? ' \u00b7 ' + locAc.gs_kts + 'kts'            : '';
          showToast(locSingle + ' \u2014 ' + locAc.status + ' \u00b7 ' + locDist.toFixed(1) + 'nm @ ' + String(Math.round(locBear)).padStart(3,'0') + '\u00b0' + altStr + spdStr, 'info', 8000);
          renderBoardMap();
        }
        return;
      }
    }
    const locRaw = nU.trim().toUpperCase();
    const spIdx = locRaw.indexOf(' ');
    const locUnit = canonicalUnit(spIdx > 0 ? locRaw.substring(0, spIdx) : locRaw);
    const locAddr = spIdx > 0 ? locRaw.substring(spIdx + 1).trim().toUpperCase() : '';
    if (!locUnit) { showAlert('ERROR', 'USAGE: LOC <UNIT> <ADDR>\nLOC <UNIT> CLR — clear location\nChain status: LOC M1 123 MAIN; AV M1'); return; }
    const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === locUnit) : null;
    if (!uO) { showAlert('ERROR', 'UNIT NOT FOUND: ' + locUnit); return; }

    // Resolve address shortcodes — parse bracket note first
    let resolvedAddr = locAddr;
    let locBracketNote = '';
    if (locAddr && locAddr !== 'CLR' && locAddr !== 'CLEAR') {
      const { base: locBase, note: locNote } = AddressLookup._parseBracketNote(locAddr);
      locBracketNote = locNote;
      const lookupAddr = locBase || locAddr;
      const byId = AddressLookup.getById(lookupAddr);
      if (byId) {
        resolvedAddr = byId.address ? (byId.address + (byId.city ? ', ' + byId.city : '')) : byId.name;
      } else {
        const results = AddressLookup.search(lookupAddr, 1);
        if (results.length === 1) resolvedAddr = results[0].address ? (results[0].address + (results[0].city ? ', ' + results[0].city : '')) : results[0].name;
        else resolvedAddr = lookupAddr;
      }
      // Append bracket note as parens inside LOC tag to avoid nested bracket issues
      if (locBracketNote) resolvedAddr += ' (' + locBracketNote + ')';
    }

    // Build note: strip old [LOC:...], add new
    let curNote = (uO.note || '').replace(/\s*\[LOC:[^\]]*\]\s*/g, '').trim();
    if (resolvedAddr && resolvedAddr !== 'CLR' && resolvedAddr !== 'CLEAR') {
      curNote = '[LOC:' + resolvedAddr + '] ' + curNote;
    }
    curNote = curNote.trim();

    const patch = { note: curNote, displayName: uO.display_name };

    setLive(true, 'LIVE • SET LOC');
    const r = await API.upsertUnit(TOKEN, locUnit, patch, uO.updated_at || '');
    if (!r.ok) return showErr(r);
    if (resolvedAddr === 'CLR' || resolvedAddr === 'CLEAR') {
      showToast(locUnit + ': LOCATION CLEARED.');
    } else {
      AddrHistory.push(resolvedAddr);
      showToast(locUnit + ': ' + resolvedAddr);
    }
    refresh();
    return;
  }

  // Messaging
  if (mU === 'MSGALL') {
    if (!nU) { showAlert('ERROR', 'USAGE: MSGALL MESSAGE TEXT'); return; }
    const r = await API.sendBroadcast(TOKEN, nU, false);
    if (!r.ok) return showErr(r);
    showAlert('MESSAGE SENT', `BROADCAST MESSAGE SENT TO ${r.recipients} RECIPIENTS`);
    refresh();
    return;
  }

  if (mU === 'HTALL') {
    if (!nU) { showAlert('ERROR', 'USAGE: HTALL URGENT MESSAGE TEXT'); return; }
    const r = await API.sendBroadcast(TOKEN, nU, true);
    if (!r.ok) return showErr(r);
    showAlert('URGENT MESSAGE SENT', `URGENT BROADCAST SENT TO ${r.recipients} RECIPIENTS`);
    refresh();
    return;
  }

  if (mU === 'MSGDP' && nU) {
    setLive(true, 'LIVE • MSG DISPATCHERS');
    const r = await API.sendToDispatchers(TOKEN, nU, false);
    if (!r.ok) return showErr(r);
    showToast('MSG SENT TO ALL DISPATCHERS');
    setLive(false);
    refresh();
    return;
  }
  if (mU === 'HTDP' && nU) {
    setLive(true, 'LIVE • HTMSG DISPATCHERS');
    const r = await API.sendToDispatchers(TOKEN, nU, true);
    if (!r.ok) return showErr(r);
    showToast('URGENT MSG SENT TO ALL DISPATCHERS');
    setLive(false);
    refresh();
    return;
  }
  if (mU === 'MSGU' && nU) {
    setLive(true, 'LIVE • MSG ALL UNITS');
    const r = await API.sendToUnits(TOKEN, nU, false);
    if (!r.ok) return showErr(r);
    showToast('MSG SENT TO ALL FIELD UNITS');
    setLive(false);
    refresh();
    return;
  }
  if ((mU === 'HTU' || mU === 'HTMSU') && nU) {
    setLive(true, 'LIVE • HTMSG ALL UNITS');
    const r = await API.sendToUnits(TOKEN, nU, true);
    if (!r.ok) return showErr(r);
    showToast('URGENT MSG SENT TO ALL FIELD UNITS');
    setLive(false);
    refresh();
    return;
  }

  if (mU.startsWith('MSG ')) {
    const tR = ma.substring(4).trim().toUpperCase();
    if (!tR || !nU) { showAlert('ERROR', 'USAGE: MSG DP2 MESSAGE TEXT'); return; }
    const r = await API.sendMessage(TOKEN, tR, nU, false);
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  if (mU.startsWith('HTMSG ')) {
    const tR = ma.substring(6).trim().toUpperCase();
    if (!tR || !nU) { showConfirm('ERROR', 'USAGE: HTMSG DP2 URGENT MESSAGE', () => { }); return; }
    const r = await API.sendMessage(TOKEN, tR, nU, true);
    if (!r.ok) return showErr(r);
    refresh();
    return;
  }

  if (/^MSG\d+$/i.test(mU)) {
    return viewMessage(mU);
  }

  if (mU.startsWith('PG ')) {
    const pgUnit = mU.substring(3).trim();
    if (!pgUnit) { showAlert('ERROR', 'USAGE: PG <UNIT>  (e.g. PG M1)'); return; }
    setLive(true, 'LIVE • PAGING ' + pgUnit);
    const r = await API.sendMessage(TOKEN, pgUnit, '[PAGE] RADIO PAGE FROM DISPATCH', true);
    setLive(false);
    if (!r.ok) return showErr(r);
    showToast('PAGE SENT TO ' + pgUnit);
    return;
  }

  if (mU.startsWith('WELF ')) {
    const welfUnit = canonicalUnit(mU.substring(5).trim());
    if (!welfUnit) { showAlert('ERROR', 'USAGE: WELF <UNIT>  (e.g. WELF M1)'); return; }
    setLive(true, 'LIVE • WELFARE CHECK ' + welfUnit);
    const r = await API.sendMessage(TOKEN, welfUnit, 'WELFARE CHECK — PLEASE CONFIRM STATUS', true);
    setLive(false);
    if (!r.ok) return showErr(r);
    showToast('WELFARE CHECK SENT TO ' + welfUnit, 'warn', 5000);
    return;
  }

  // MA — Mutual Aid: request, acknowledge, release, list
  // MA <INC> <AGENCY>          → requestMA
  // MA ACK <INC> <AGENCY>      → acknowledgeMA
  // MA REL <INC> <AGENCY>      → releaseMA
  // MA LIST <INC>              → listMA
  if (mU === 'MA' || mU.startsWith('MA ')) {
    const maParts = mU.replace(/^MA\s*/i, '').trim();
    // MA ACK or MA REL
    const maSubMatch = maParts.match(/^(ACK|REL)\s+(\S+)\s+(.+)$/i);
    if (maSubMatch) {
      const maSub = maSubMatch[1].toUpperCase();
      const maIncRaw = maSubMatch[2].trim();
      const maAgency = maSubMatch[3].trim().toUpperCase();
      setLive(true, 'LIVE • MA ' + maSub + ' ' + maIncRaw);
      const r = maSub === 'ACK'
        ? await API.acknowledgeMA(TOKEN, maIncRaw, maAgency)
        : await API.releaseMA(TOKEN, maIncRaw, maAgency);
      setLive(false);
      if (!r.ok) return showErr(r);
      const verb = maSub === 'ACK' ? 'ACKNOWLEDGED' : 'RELEASED';
      showToast('MUTUAL AID ' + verb + ': ' + maAgency + ' → ' + (r.incidentId || maIncRaw), 'good', 5000);
      return;
    }
    // MA LIST <INC>
    const maListMatch = maParts.match(/^LIST\s+(\S+)$/i);
    if (maListMatch) {
      setLive(true, 'LIVE • MA LIST');
      const r = await API.listMA(TOKEN, maListMatch[1].trim());
      setLive(false);
      if (!r.ok) return showErr(r);
      const list = r.mutualAid || [];
      if (!list.length) { showAlert('MUTUAL AID', 'NO MUTUAL AID REQUESTS FOR ' + r.incidentId); return; }
      const rows = list.map(t => '<li>' + esc(t.agency) + ' — <strong>' + esc(t.status) + '</strong></li>').join('');
      showAlert('MUTUAL AID: ' + r.incidentId, '<ul style="margin:0;padding-left:1.2em;">' + rows + '</ul>');
      return;
    }
    // MA <INC> <AGENCY> — request mutual aid
    const maReqMatch = maParts.match(/^(\S+)\s+(.+)$/);
    if (maReqMatch) {
      const maIncRaw2 = maReqMatch[1].trim();
      const maAgency2 = maReqMatch[2].trim().toUpperCase();
      setLive(true, 'LIVE • MA REQUEST ' + maIncRaw2);
      const r = await API.requestMA(TOKEN, maIncRaw2, maAgency2);
      setLive(false);
      if (!r.ok) return showErr(r);
      const shortId = String(r.incidentId || '').replace(/^[A-Z]*\d{2}-0*/, '');
      const displayName = r.agencyName || r.agency || maAgency2;
      showToast('MUTUAL AID REQUESTED: ' + displayName + ' → ' + (r.incidentId || ''), 'warn', 6000);
      return;
    }
    showAlert('USAGE',
      'MA &lt;INC&gt; &lt;AGENCY&gt; — Request mutual aid\n' +
      'MA ACK &lt;INC&gt; &lt;AGENCY&gt; — Acknowledge (agency responding)\n' +
      'MA REL &lt;INC&gt; &lt;AGENCY&gt; — Release mutual aid\n' +
      'MA LIST &lt;INC&gt; — Show all MA requests for incident');
    return;
  }

  // LFN LINK <CALLSIGN> <INC> — link an air resource to an incident
  // Appends [AIR:CALLSIGN:LINKED] tag to incident note; shows AIR badge in active calls bar
  if (mU === 'LFN LINK') {
    if (!nU) { showAlert('ERROR', 'USAGE: LFN LINK <CALLSIGN> <INC>\nExample: LFN LINK LF11 26-0042'); return; }
    const llParts = nU.trim().toUpperCase().split(/\s+/);
    if (llParts.length < 2) { showAlert('ERROR', 'USAGE: LFN LINK <CALLSIGN> <INC>\nExample: LFN LINK LF11 26-0042'); return; }
    const llCallsign = llParts[0];
    const llIncRaw   = llParts.slice(1).join(' ');
    const llAcEntry  = Object.values(AIR_FLEET).find(v => v.callsign === llCallsign);
    if (!llAcEntry) {
      showAlert('ERROR', 'UNKNOWN AIR CALLSIGN: ' + llCallsign + '\nKnown: ' + Object.values(AIR_FLEET).map(v => v.callsign).join(', '));
      return;
    }
    setLive(true, 'LIVE • LFN LINK ' + llCallsign);
    const r = await API.appendIncidentNote(TOKEN, llIncRaw, '[AIR:' + llCallsign + ':LINKED]');
    setLive(false);
    if (!r.ok) return showErr(r);
    const llShort = String(r.incidentId || llIncRaw).replace(/^[A-Z]*\d{2}-0*/, '');
    showToast('AIR LINKED: ' + llCallsign + ' \u2192 ' + (r.incidentId || llIncRaw), 'good', 5000);
    return;
  }

  // GPS <UNIT> — request GPS update from field device, then show on map
  if (mU.startsWith('GPS ')) {
    const gpsUnit = canonicalUnit(mU.substring(4).trim());
    if (!gpsUnit) { showAlert('ERROR', 'USAGE: GPS <UNIT>  (e.g. GPS M1)'); return; }
    // Send [GPS:UL] ping to field device so it reports back coords
    setLive(true, 'LIVE • GPS PING → ' + gpsUnit);
    const gpsR = await API.sendMessage(TOKEN, gpsUnit, '[GPS:UL]', true);
    setLive(false);
    if (!gpsR.ok) return showErr(gpsR);
    showToast('GPS PING SENT → ' + gpsUnit + ' — MAP WILL AUTO-UPDATE IN ~20s', 'info', 20000);
    _ensureMapOpen(() => { renderBoardMap(); });
    // Re-focus map after unit has had time to report GPS location (~20s)
    setTimeout(() => { _ensureMapOpen(() => { renderBoardMap(); focusUnitOnMap(gpsUnit); }); }, 20000);
    return;
  }

  // GPSUL <UNIT> — request unit to send GPS location update to board
  if (mU.startsWith('GPSUL ')) {
    const gpsuUnit = canonicalUnit(mU.substring(6).trim());
    if (!gpsuUnit) { showAlert('ERROR', 'USAGE: GPSUL <UNIT>  (e.g. GPSUL M1)'); return; }
    setLive(true, 'LIVE • GPS UPDATE REQUEST → ' + gpsuUnit);
    const r = await API.sendMessage(TOKEN, gpsuUnit, '[GPS:UL]', true);
    setLive(false);
    if (!r.ok) return showErr(r);
    showToast('GPS UPDATE REQUESTED → ' + gpsuUnit, 'info', 5000);
    return;
  }

  // M <UNIT> <MESSAGE> — add note to unit's incident (no alert)
  if (cmd === 'M' && mU.startsWith('M ')) {
    const mTarget = canonicalUnit(mU.substring(2).trim());
    if (mTarget) {
      const uO = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === mTarget) : null;
      if (uO && uO.incident) {
        if (!nU) { showAlert('ERROR', 'USAGE: M <UNIT> <MESSAGE>'); return; }
        setLive(true, 'LIVE \u2022 MSG');
        const r = await API.appendIncidentNote(TOKEN, uO.incident, nU);
        if (!r.ok) return showErr(r);
        showToast('NOTE ADDED TO ' + uO.incident);
        refresh();
        return;
      }
      if (uO && !uO.incident) { showAlert('ERROR', mTarget + ' IS NOT ASSIGNED TO AN INCIDENT.'); return; }
      if (!uO) { showAlert('ERROR', 'UNIT NOT FOUND: ' + mTarget); return; }
    }
  }

  if ((mU + ' ' + nU).trim().toUpperCase().startsWith('DEL ALL MSG')) {
    return deleteAllMessages();
  }

  if (mU.startsWith('DEL MSG') || (mU === 'DEL' && /^MSG/i.test(nU))) {
    const re = mU.startsWith('DEL MSG') ? mU.substring(7).trim() : nU.trim();
    if (!re) { showConfirm('ERROR', 'USAGE: DEL MSG1 OR DEL MSG #2 (inbox position)', () => { }); return; }
    const msgId = re.toUpperCase();
    if (/^MSG\d+$/i.test(msgId)) {
      return deleteMessage(msgId);
    }
    if (/^\d+$/.test(re)) {
      // Bare number = local inbox position (1-based) — resolve to actual message_id
      const idx = parseInt(re, 10) - 1;
      const msgs = STATE.messages || [];
      if (idx < 0 || idx >= msgs.length) { showConfirm('ERROR', 'MESSAGE #' + re + ' NOT FOUND. YOU HAVE ' + msgs.length + ' MESSAGE' + (msgs.length !== 1 ? 'S' : '') + '.', () => { }); return; }
      return deleteMessage(msgs[idx].message_id);
    }
    showConfirm('ERROR', 'USAGE: DEL MSG1 OR DEL MSG #2 (inbox position)', () => { });
    return;
  }

  // Parse status + unit commands (D JC, JC OS, etc.)
  const tk = ma.trim().split(/\s+/).filter(Boolean);

  function parseStatusUnit(t) {
    if (t.length >= 2 && VALID_STATUSES.has(t[0].toUpperCase())) {
      return { status: t[0].toUpperCase(), unit: t.slice(1).join(' ') };
    }
    if (t.length >= 2 && VALID_STATUSES.has(t[t.length - 1].toUpperCase())) {
      return { status: t[t.length - 1].toUpperCase(), unit: t.slice(0, -1).join(' ') };
    }
    if (t.length === 2 && VALID_STATUSES.has(t[1].toUpperCase())) {
      return { status: t[1].toUpperCase(), unit: t[0] };
    }
    if (t.length === 3 && VALID_STATUSES.has(t[0].toUpperCase())) {
      return { status: t[0].toUpperCase(), unit: t.slice(1).join(' ') };
    }
    return null;
  }

  const pa = parseStatusUnit(tk);
  if (!pa) {
    showAlert('ERROR', 'UNKNOWN COMMAND. TYPE HELP FOR ALL COMMANDS.');
    return;
  }

  const stCmd = pa.status;
  let rawUnit = pa.unit;
  let incidentId = '';

  // OOS reason intercept for command-line OOS
  let oosNotePrefix = '';
  if (stCmd === 'OOS') {
    const oosUnit = canonicalUnit(rawUnit) || rawUnit;
    const reason = await promptOOSReason(oosUnit);
    if (!reason) return;
    oosNotePrefix = `[OOS:${reason}] `;
  }

  // Resolve 4-digit shorthand to full incident_id by looking up STATE.incidents first,
  // then falling back to SC+YY- prefix (current format). Avoids FK failures from format mismatch.
  function resolveIncidentId(raw) {
    const stripped = raw.replace(/^INC[-\s]*/i, '').trim().toUpperCase();
    if (/^\d{4}$/.test(stripped)) {
      const stateInc = (STATE.incidents || []).find(i => i.incident_id.endsWith('-' + stripped));
      if (stateInc) return stateInc.incident_id;
      const yy = String(new Date().getFullYear()).slice(-2);
      return 'SC' + yy + '-' + stripped;
    }
    return stripped;
  }

  // Check for incident ID at end of unit (e.g. "D AMWC1 INC-0001" or "D AMWC1 SC26-0001")
  const incMatch = rawUnit.match(/\s+(INC\s*[A-Z]*\d{2}-\d{4}|INC\s*\d{4}|[A-Z]{1,4}\d{2}-\d{4}|\d{2}-\d{4}|\d{4})$/i);
  if (incMatch) {
    incidentId = resolveIncidentId(incMatch[1]);
    rawUnit = rawUnit.substring(0, incMatch.index).trim();
  }

  // Also check if nU looks like an incident ID (e.g. "D AMWC1 INC-0001" or "D AMWC1 SC26-0001")
  let nuUsedAsIncident = false;
  const dispatchLikeStatuses = new Set(['D', 'DE', 'AT', 'TH']);
  if (!incidentId && dispatchLikeStatuses.has(stCmd) && nU) {
    const nuIncMatch = nU.trim().match(/^(INC[-\s]?[A-Z]*\d{2}-\d{4}|INC[-\s]?\d{4}|[A-Z]{1,4}\d{2}-\d{4}|\d{2}-\d{4}|\d{4})$/i);
    if (nuIncMatch) {
      incidentId = resolveIncidentId(nuIncMatch[1]);
      nuUsedAsIncident = true;
    }
  }

  // AV FORCE / dispo check
  // Note: "AV AMBLS1 FORCE" is parsed as ma="AV AMBLS1" nU="FORCE" by the tokenizer,
  // so rawUnit comes out as "AMBLS1" and nU is "FORCE". Check both patterns.
  // "MALS1 AV TC" → ma="MALS1 AV" nU="TC" → stCmd=AV, rawUnit="MALS1", nU="TC"
  let avForce = false;
  let _avDispoClose = null; // { incidentId, dispo } — set when nU is a dispo shortcode
  if (stCmd === 'AV') {
    const forceMatch = rawUnit.match(/^(.+?)\s+FORCE$/i);
    if (forceMatch) {
      avForce = true;
      rawUnit = forceMatch[1].trim();
    } else if (nU.trim().toUpperCase() === 'FORCE') {
      avForce = true;
      // rawUnit is already just the unit name (FORCE was in nU)
    } else {
      const avNuUpper = nU.trim().toUpperCase();
      const avDispoResolved = DISPO_SHORT_CODES[avNuUpper];
      const avUnitId = canonicalUnit(rawUnit);
      const avUnitObj = (STATE && STATE.units) ? STATE.units.find(x => String(x.unit_id || '').toUpperCase() === avUnitId) : null;
      if (avDispoResolved && avUnitObj && avUnitObj.incident) {
        // Dispo-close AV: bypass FORCE restriction, close incident in background
        avForce = true;
        _avDispoClose = { incidentId: avUnitObj.incident, dispo: avDispoResolved };
      } else if (avUnitObj && avUnitObj.incident) {
        showErr({ error: 'UNIT HAS ACTIVE INCIDENT (' + avUnitObj.incident + '). USE: AV ' + rawUnit.toUpperCase() + ' FORCE  OR  ' + rawUnit.toUpperCase() + ' AV TC' });
        return;
      }
    }
  }

  const u = canonicalUnit(rawUnit);
  const boardUnit = (STATE && STATE.units) ? STATE.units.find(function(x) { return String(x.unit_id || '').toUpperCase() === u; }) : null;
  if (!boardUnit) {
    showErr({ error: 'UNIT ' + u + ' NOT ON BOARD. USE LOGON ' + u + ' TO ACTIVATE FROM ROSTER.' });
    return;
  }
  const dN = displayNameForUnit(u);
  const p = { status: stCmd, displayName: dN };
  // nU='FORCE' and nU=dispo-shortcode are consumed as flags — don't write as unit note
  const _nuNote = (avForce && (nU.trim().toUpperCase() === 'FORCE' || _avDispoClose)) ? '' : nU;
  if (oosNotePrefix || (_nuNote && !nuUsedAsIncident)) p.note = oosNotePrefix + _nuNote;
  if (incidentId) {
    // Validate incident exists in STATE before dispatching (prevents FK constraint errors)
    const incObj = (STATE.incidents || []).find(i => i.incident_id === incidentId);
    if (!incObj) {
      showErr({ error: 'INCIDENT NOT FOUND: ' + incidentId + '\nVERIFY INC# AND TRY AGAIN.' });
      return;
    }
    p.incident = incidentId;
    // Auto-copy incident destination to unit
    if (incObj.destination) {
      p.destination = incObj.destination;
    }
  }

  const _prevStat = { status: boardUnit.status, note: boardUnit.note || '', incident: boardUnit.incident || '', destination: boardUnit.destination || '' };
  setLive(true, 'LIVE • UPDATE');
  const r = await API.upsertUnit(TOKEN, u, p, '');
  if (!r.ok) return showErr(r);

  // Dispo-close AV: fire closeIncident in background after unit goes AV
  if (_avDispoClose) {
    const _adc = _avDispoClose;
    API.closeIncident(TOKEN, _adc.incidentId, _adc.dispo).then(function(cr) {
      if (cr.ok) showToast('INC ' + _adc.incidentId.replace(/^[A-Z]*\d{2}-0*/, '') + ' CLOSED — ' + _adc.dispo);
      else showToast(u + ' AV — INC CLOSE FAILED: ' + (cr.error || 'ERROR'));
      refresh();
    });
  }

  pushUndo(`${u}: ${_prevStat.status}→${stCmd}`, async () => {
    const rv = await API.upsertUnit(TOKEN, u, { status: _prevStat.status, note: _prevStat.note, incident: _prevStat.incident, destination: _prevStat.destination }, '');
    if (!rv.ok) throw new Error(rv.error || 'API error');
  });
  refresh();
  autoFocusCmd();
}

// ============================================================
// Command Hints Autocomplete
// ============================================================
function selectLocAddrHint(addr, unit) {
  const cmdEl = document.getElementById('cmd');
  if (cmdEl) {
    cmdEl.value = 'LOC ' + unit + ' ' + addr;
    cmdEl.focus();
    cmdEl.setSelectionRange(cmdEl.value.length, cmdEl.value.length);
  }
  hideCmdHints();
}

function showCmdHints(query) {
  const el = document.getElementById('cmdHints');
  if (!el) return;
  if (!query || query.length < 1) { hideCmdHints(); return; }

  const q = query.toUpperCase();

  // LOC <UNIT> <partial> — show address history suggestions
  const locM = q.match(/^LOC\s+(\S+)\s+(.*)$/);
  if (locM) {
    const unitPart = locM[1];
    const partial  = locM[2].trim();
    const hist = AddrHistory.get().filter(a => !partial || a.includes(partial)).slice(0, 6);
    if (hist.length) {
      CMD_HINT_INDEX = -1;
      el.innerHTML = hist.map((a, i) =>
        '<div class="cmd-hint-item" data-index="' + i + '" onmousedown="selectLocAddrHint(\'' + a.replace(/\\/g,'\\\\').replace(/'/g,"\\'") + '\',\'' + unitPart + '\')">' +
        '<span class="hint-cmd">LOC ' + esc(unitPart) + ' ' + esc(a) + '</span>' +
        '<span class="hint-desc">RECENT ADDRESS</span>' +
        '</div>'
      ).join('');
      el.classList.add('open');
      return;
    }
    hideCmdHints();
    return;
  }

  const matches = CMD_HINTS.filter(h => h.cmd.toUpperCase().includes(q)).slice(0, 5);

  if (!matches.length) { hideCmdHints(); return; }

  CMD_HINT_INDEX = -1;
  el.innerHTML = matches.map((h, i) =>
    '<div class="cmd-hint-item" data-index="' + i + '" onmousedown="selectCmdHint(' + i + ')">' +
    '<span class="hint-cmd">' + esc(h.cmd) + '</span>' +
    '<span class="hint-desc">' + esc(h.desc) + '</span>' +
    '</div>'
  ).join('');
  el.classList.add('open');
}

function hideCmdHints() {
  const el = document.getElementById('cmdHints');
  if (el) { el.classList.remove('open'); el.innerHTML = ''; }
  CMD_HINT_INDEX = -1;
}

function selectCmdHint(index) {
  const el = document.getElementById('cmdHints');
  if (!el) return;
  const items = el.querySelectorAll('.cmd-hint-item');
  if (index < 0 || index >= items.length) return;

  const cmdText = CMD_HINTS.filter(h => {
    const q = (document.getElementById('cmd').value || '').toUpperCase();
    return h.cmd.toUpperCase().startsWith(q);
  })[index];

  if (cmdText) {
    // Extract the fixed prefix of the command (before first <)
    const raw = cmdText.cmd;
    const angleBracket = raw.indexOf('<');
    const prefix = angleBracket > 0 ? raw.substring(0, angleBracket).trimEnd() + ' ' : raw;
    const cmdEl = document.getElementById('cmd');
    cmdEl.value = prefix;
    cmdEl.focus();
    cmdEl.setSelectionRange(prefix.length, prefix.length);
  }
  hideCmdHints();
}

function navigateCmdHints(dir) {
  const el = document.getElementById('cmdHints');
  if (!el || !el.classList.contains('open')) return false;
  const items = el.querySelectorAll('.cmd-hint-item');
  if (!items.length) return false;

  items.forEach(it => it.classList.remove('active'));
  CMD_HINT_INDEX += dir;
  if (CMD_HINT_INDEX < 0) CMD_HINT_INDEX = items.length - 1;
  if (CMD_HINT_INDEX >= items.length) CMD_HINT_INDEX = 0;
  items[CMD_HINT_INDEX].classList.add('active');
  return true;
}

function openShiftReportWindow(rpt) {
  const w = window.open('', '_blank');
  if (!w) { showAlert('BLOCKED', 'ALLOW POPUPS FOR SHIFT REPORT.'); return; }
  const av = rpt.metrics.averagesMinutes || {};
  let html = `<!DOCTYPE html><html><head><title>SHIFT REPORT</title>
  <style>body{font-family:monospace;background:#0d1117;color:#e6edf3;padding:24px}
  h2{color:#58a6ff}table{border-collapse:collapse;width:100%}
  td,th{border:1px solid #30363d;padding:6px 10px;font-size:12px}
  th{background:#161b22;text-align:left}.good{color:#7fffb2}.warn{color:#ffd66b}.bad{color:#ff6b6b}
  </style></head><body>`;
  html += `<h2>SHIFT REPORT — ${rpt.windowHours}H WINDOW</h2>`;
  html += `<p style="font-size:11px;color:#8b949e">GENERATED ${new Date(rpt.generatedAt).toLocaleString('en-US',{hour12:false})} | INCIDENTS: ${rpt.incidentCount}</p>`;

  html += '<h3>RESPONSE TIMES</h3><table><tr><th>METRIC</th><th>AVG (MIN)</th><th>TARGET</th><th>STATUS</th></tr>';
  Object.keys(KPI_TARGETS).forEach(k => {
    const val = av[k];
    const tgt = KPI_TARGETS[k];
    const cls = val == null ? '' : val <= tgt ? 'good' : val <= tgt*1.5 ? 'warn' : 'bad';
    html += `<tr><td>${k}</td><td class="${cls}">${val ?? '—'}</td><td>${tgt}</td><td class="${cls}">${val == null ? '—' : val <= tgt ? 'OK' : 'OVER'}</td></tr>`;
  });
  html += '</table>';

  if (rpt.incidents.length) {
    html += '<h3>INCIDENTS</h3><table><tr><th>ID</th><th>TYPE</th><th>PRIORITY</th><th>SCENE</th><th>UNITS</th><th>STATUS</th><th>DISPOSITION</th></tr>';
    rpt.incidents.forEach(inc => {
      const dispM = (inc.incident_note || '').match(/\[DISP:([^\]]+)\]/i);
      const disp = dispM ? dispM[1].toUpperCase() : ((inc.disposition || '').toUpperCase() || '—');
      const incIdEsc = esc(inc.incident_id);
      const incLink = `<a href="#" title="Open incident report" onclick="window.opener&&window.opener.runQuickCmd('REPORT INC ${incIdEsc}');return false;" style="color:#58a6ff;text-decoration:none;" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${incIdEsc}</a>`;
      html += `<tr><td>${incLink}</td><td>${esc(inc.incident_type||'—')}</td><td>${esc(inc.priority||'—')}</td><td>${esc(inc.scene_address||'—')}</td><td>${esc(inc.units||'—')}</td><td>${esc(inc.status)}</td><td>${esc(disp)}</td></tr>`;
    });
    html += '</table>';
  }

  if (rpt.unitSummaries.length) {
    html += '<h3>UNIT ACTIVITY</h3><table><tr><th>UNIT</th><th>DISPATCHES</th><th>D (MIN)</th><th>OS (MIN)</th><th>T (MIN)</th><th>OOS (MIN)</th></tr>';
    rpt.unitSummaries.forEach(u => {
      const ts = u.timeInStatus;
      html += `<tr><td>${esc(u.unit_id)}</td><td>${u.dispatches}</td><td>${ts['D']||0}</td><td>${ts['OS']||0}</td><td>${ts['T']||0}</td><td>${ts['OOS']||0}</td></tr>`;
    });
    html += '</table>';
  }

  html += '</body></html>';
  w.document.write(html);
  w.document.close();
}

function exportCurrentIncident() {
  if (!CURRENT_INCIDENT_ID || !TOKEN) return;
  setLive(true, 'LIVE • INCIDENT EXPORT');
  API.getIncident(TOKEN, CURRENT_INCIDENT_ID).then(r => {
    if (!r.ok) return showErr(r);
    openIncidentPrintWindow(r);
  });
}

function openIncidentPrintWindow(r) {
  const w = window.open('', '_blank');
  if (!w) { showAlert('BLOCKED', 'ALLOW POPUPS FOR INCIDENT REPORT.'); return; }
  const inc = r.incident;

  const fmt = (v) => v ? fmtTime24(v) : '—';
  const fmtMin = (a, b) => {
    if (!a || !b) return null;
    const m = Math.round((new Date(b) - new Date(a)) / 60000);
    return m > 0 ? m + ' min' : null;
  };
  const genAt = new Date().toISOString().replace('T',' ').slice(0,16) + ' UTC';
  let html = `<!DOCTYPE html><html><head><title>INCIDENT REPORT — ${inc.incident_id}</title>
  <style>
  body{font-family:monospace;background:#fff;color:#1a1a1a;padding:24px;max-width:860px;margin:0 auto;}
  h2{font-size:18px;margin:0 0 4px;}h3{font-size:13px;margin:16px 0 6px;text-transform:uppercase;border-bottom:1px solid #999;}
  table{border-collapse:collapse;width:100%;margin-bottom:12px;}
  td,th{border:1px solid #bbb;padding:5px 10px;font-size:12px;text-align:left;}
  th{background:#eee;font-weight:bold;width:180px;}
  .audit{font-size:11px;padding:3px 0;border-bottom:1px solid #eee;}
  .audit-ts{color:#555;min-width:80px;display:inline-block;}
  .audit-actor{color:#333;min-width:90px;display:inline-block;font-weight:bold;}
  .good{color:#1a7a3a;}.warn{color:#a06000;}.bad{color:#cc2222;}
  .footer{margin-top:24px;font-size:10px;color:#666;border-top:1px solid #ccc;padding-top:8px;}
  .confidential{font-weight:bold;color:#cc2222;}
  @media print{body{padding:8px;}h2,h3{break-after:avoid;}table{break-inside:avoid;}}
  </style></head><body>`;
  html += `<div style="display:flex;justify-content:space-between;align-items:flex-start;">`;
  html += `<div><h2>HOSCAD INCIDENT REPORT</h2><div style="font-size:11px;color:#555;">HOSCAD — SCMC EMS TRACKING SYSTEM</div></div>`;
  html += `<div class="confidential" style="font-size:10px;text-align:right;">FOR OFFICIAL USE ONLY<br>NOT FOR PUBLIC DISTRIBUTION</div></div>`;
  html += `<h3>Incident Details</h3>`;
  html += `<table>
    <tr><th>INCIDENT #</th><td><strong>${esc(inc.incident_id)}</strong></td></tr>
    <tr><th>STATUS</th><td>${esc(inc.status)}</td></tr>
    <tr><th>TYPE</th><td>${esc(inc.incident_type||'—')}</td></tr>
    <tr><th>PRIORITY</th><td>${esc(inc.priority||'—')}</td></tr>
    ${inc.level_of_care ? `<tr><th>LEVEL OF CARE</th><td>${esc(inc.level_of_care)}</td></tr>` : ''}
    <tr><th>SCENE ADDRESS</th><td>${esc(inc.scene_address||'—')}</td></tr>
    <tr><th>DESTINATION</th><td>${esc(inc.destination||'—')}</td></tr>
    <tr><th>UNITS ASSIGNED</th><td>${esc(inc.units||'—')}</td></tr>
    ${(() => { const dm = (inc.incident_note||'').match(/\[DISP:([^\]]+)\]/i); const dv = dm ? dm[1].toUpperCase() : (inc.disposition||'').toUpperCase(); return dv ? `<tr><th>DISPOSITION</th><td>${esc(dv)}</td></tr>` : ''; })()}
    <tr><th>CREATED</th><td>${fmt(inc.created_at)} by ${esc(inc.created_by||'?')}</td></tr>
    <tr><th>DISPATCH TIME</th><td>${fmt(inc.dispatch_time)}</td></tr>
    <tr><th>ENROUTE TIME</th><td>${fmt(inc.enroute_time)}</td></tr>
    <tr><th>ARRIVAL TIME</th><td>${fmt(inc.arrival_time)}</td></tr>
    <tr><th>TRANSPORT TIME</th><td>${fmt(inc.transport_time)}</td></tr>
    <tr><th>AT HOSPITAL TIME</th><td>${fmt(inc.at_hospital_time)}</td></tr>
    <tr><th>HANDOFF TIME</th><td>${fmt(inc.handoff_time)}</td></tr>
    <tr><th>INCIDENT NOTE</th><td>${esc((inc.incident_note||'').replace(/\[DISP:[^\]]*\]\s*/gi,'').replace(/\[[A-Z]+:[^\]]*\]\s*/gi,'').trim()||'—')}</td></tr>
  </table>`;
  // Response time calculations
  const rtRows = [
    { label: 'DISPATCH → ENROUTE', a: inc.dispatch_time,   b: inc.enroute_time,    target: KPI_TARGETS['D→DE']  || 5  },
    { label: 'ENROUTE → ON SCENE', a: inc.enroute_time,    b: inc.arrival_time,    target: KPI_TARGETS['DE→OS'] || 10 },
    { label: 'ON SCENE → TRANSPORT', a: inc.arrival_time,  b: inc.transport_time,  target: KPI_TARGETS['OS→T']  || 30 },
    { label: 'TRANSPORT → AT HOSP', a: inc.transport_time, b: inc.at_hospital_time, target: null },
    { label: 'AT HOSP → HANDOFF',  a: inc.at_hospital_time, b: inc.handoff_time,   target: KPI_TARGETS['T→AV']  || 20 },
  ];
  const rtAvail = rtRows.filter(r => r.a && r.b);
  if (rtAvail.length) {
    html += '<h3>RESPONSE TIMES</h3><table><tr><th>INTERVAL</th><th>ELAPSED</th><th>TARGET</th><th>STATUS</th></tr>';
    rtAvail.forEach(r => {
      const m = fmtMin(r.a, r.b);
      const mNum = m ? parseInt(m) : null;
      const cls = !r.target || mNum == null ? '' : mNum <= r.target ? 'good' : mNum <= r.target * 1.5 ? 'warn' : 'bad';
      const status = !r.target ? '—' : mNum == null ? '—' : mNum <= r.target ? 'OK' : 'OVER';
      html += `<tr><td>${esc(r.label)}</td><td class="${cls}">${m || '—'}</td><td>${r.target ? r.target + ' min' : '—'}</td><td class="${cls}">${status}</td></tr>`;
    });
    html += '</table>';
  }

  if (r.audit && r.audit.length) {
    html += '<h3>Incident Audit Trail</h3>';
    r.audit.forEach(a => {
      html += `<div class="audit"><span class="audit-ts">${fmt(a.ts)}</span> <span class="audit-actor">${esc((a.actor||'').toUpperCase())}</span> ${esc(a.message)}</div>`;
    });
  }

  html += `<div class="footer">
    <span class="confidential">FOR OFFICIAL USE ONLY — HOSCAD CONFIDENTIAL OPERATIONAL RECORD</span><br>
    Generated: ${genAt} | HOSCAD EMS Tracking System — SCMC<br>
    This document may contain protected health information (PHI). Handle per HIPAA requirements.
    Do not distribute outside authorized personnel.
  </div>`;
  html += '</body></html>';
  w.document.write(html);
  w.document.close();
  setTimeout(() => { try { w.print(); } catch(e) {} }, 600);
}

function openUnitReportWindow(rpt) {
  const w = window.open('', '_blank');
  if (!w) { showAlert('BLOCKED', 'ALLOW POPUPS FOR UNIT REPORT.'); return; }
  const ts = rpt.timeInStatus || {};
  const STATUS_ORDER = ['D','OS','T','OOS','AV','BRK'];
  const fmtMin = (m) => {
    if (!m) return '0M';
    if (m < 60) return m + 'M';
    return Math.floor(m / 60) + 'H ' + (m % 60) + 'M';
  };

  let html = `<!DOCTYPE html><html><head><title>UNIT REPORT — ${rpt.unit_id}</title>
  <style>body{font-family:monospace;background:#0d1117;color:#e6edf3;padding:24px}
  h2{color:#58a6ff}h3{color:#79c0ff;margin-top:20px}
  table{border-collapse:collapse;width:100%}
  td,th{border:1px solid #30363d;padding:6px 10px;font-size:12px}
  th{background:#161b22;text-align:left}
  .good{color:#7fffb2}.warn{color:#ffd66b}.bad{color:#ff6b6b}
  .audit{font-size:11px;color:#8b949e;margin:2px 0}
  </style></head><body>`;

  html += `<h2>UNIT REPORT — ${rpt.unit_id}</h2>`;
  html += `<p style="font-size:11px;color:#8b949e">WINDOW: ${rpt.windowHours}H | ` +
          `${new Date(rpt.startIso).toLocaleString('en-US',{hour12:false})} → ${new Date(rpt.endIso).toLocaleString('en-US',{hour12:false})} | ` +
          `GENERATED ${new Date(rpt.generatedAt).toLocaleString('en-US',{hour12:false})}</p>`;

  // Status time breakdown
  html += '<h3>STATUS TIME BREAKDOWN</h3><table><tr><th>STATUS</th><th>TIME</th><th>MINUTES</th></tr>';
  const allKeys = [...new Set([...STATUS_ORDER, ...Object.keys(ts)])];
  let totalMin = 0;
  allKeys.forEach(k => { totalMin += ts[k] || 0; });
  allKeys.forEach(k => {
    if (!ts[k]) return;
    const pct = totalMin ? Math.round((ts[k] / totalMin) * 100) : 0;
    html += `<tr><td>${k}</td><td>${fmtMin(ts[k])} (${pct}%)</td><td>${ts[k]}</td></tr>`;
  });
  html += `<tr><td><strong>TOTAL</strong></td><td>${fmtMin(totalMin)}</td><td>${totalMin}</td></tr>`;
  html += '</table>';

  html += `<p style="font-size:12px">DISPATCHES: <strong>${rpt.dispatches}</strong> | AUDIT EVENTS: <strong>${rpt.eventCount}</strong></p>`;

  // Incidents served
  if (rpt.incidents && rpt.incidents.length) {
    html += `<h3>INCIDENTS SERVED (${rpt.incidents.length})</h3>`;
    html += '<table><tr><th>ID</th><th>TYPE</th><th>PRI</th><th>SCENE</th><th>DEST</th><th>DISPATCH</th><th>ARRIVAL</th><th>TRANSPORT</th><th>HANDOFF</th></tr>';
    rpt.incidents.forEach(inc => {
      const fmt = (v) => v ? (() => { try { return new Date(v).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit',hour12:false}); } catch(e) { return v; } })() : '—';
      const incIdEsc = esc(inc.incident_id);
      const incLink = `<a href="#" title="Open incident report" onclick="window.opener&&window.opener.runQuickCmd('REPORT INC ${incIdEsc}');return false;" style="color:#58a6ff;text-decoration:none;" onmouseover="this.style.textDecoration='underline'" onmouseout="this.style.textDecoration='none'">${incIdEsc}</a>`;
      html += `<tr><td>${incLink}</td><td>${esc(inc.incident_type||'—')}</td><td>${esc(inc.priority||'—')}</td>` +
              `<td>${esc(inc.scene_address||'—')}</td><td>${esc(inc.destination||'—')}</td>` +
              `<td>${fmt(inc.dispatch_time)}</td><td>${fmt(inc.arrival_time)}</td>` +
              `<td>${fmt(inc.transport_time)}</td><td>${fmt(inc.handoff_time)}</td></tr>`;
    });
    html += '</table>';
  } else {
    html += '<p style="color:#8b949e;font-size:12px">NO INCIDENTS FOUND IN WINDOW.</p>';
  }

  // Audit trail
  if (rpt.events && rpt.events.length) {
    html += `<h3>AUDIT TRAIL</h3>`;
    rpt.events.forEach(e => {
      const t = (() => { try { return new Date(e.ts).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false}); } catch(x) { return e.ts; } })();
      const dest = e.new_dest ? ` → ${e.new_dest}` : '';
      const inc  = e.new_incident ? ` INC:${e.new_incident}` : '';
      html += `<div class="audit">[${t}] ${esc(e.action)} ${esc(e.prev_status)}→${esc(e.new_status)}${esc(dest)}${esc(inc)} (by ${esc(e.actor)})</div>`;
    });
  }

  html += '</body></html>';
  w.document.write(html);
  w.document.close();
}

function showHelp() {
  window.open('help/', '_blank');
}
/* showHelpLegacy removed — dead code (HV-5) */
function _showHelpLegacy_REMOVED() {
  showAlert('HELP - COMMAND REFERENCE', `SCMC HOSCAD/EMS TRACKING - COMMAND REFERENCE

═══════════════════════════════════════════════════
VIEW / DISPLAY COMMANDS
═══════════════════════════════════════════════════
V SIDE                  Toggle sidebar panel
V MSG                   Toggle messages in sidebar
V INC                   Toggle incident queue
V ALL                   Show all panels
V NONE                  Hide all panels
F <STATUS>              Filter board by status
F ALL                   Clear status filter
SORT STATUS             Sort by status
SORT UNIT               Sort by unit ID
SORT ELAPSED            Sort by elapsed time
SORT UPDATED            Sort by last updated
SORT REV                Reverse sort direction
DEN                     Cycle density (compact/normal/expanded)
DEN COMPACT             Set compact density
DEN NORMAL              Set normal density
DEN EXPANDED            Set expanded density
PRESET DISPATCH         Dispatch view preset
PRESET SUPERVISOR       Supervisor view preset
PRESET FIELD            Field view preset
ELAPSED SHORT           Elapsed: 12M, 1H30M
ELAPSED LONG            Elapsed: 1:30:45
ELAPSED OFF             Hide elapsed time
NIGHT                   Toggle night mode (dim display)
CLR                     Clear all filters + search

═══════════════════════════════════════════════════
GENERAL COMMANDS
═══════════════════════════════════════════════════
H / HELP                Show this help
STATUS                  System status summary
REFRESH                 Reload board data
INFO                    Quick reference (key numbers)
INFO ALL                Full dispatch/emergency directory
INFO DISPATCH           911/PSAP centers
INFO AIR                Air ambulance dispatch
INFO OSP                Oregon State Police
INFO CRISIS             Mental health / crisis lines
INFO POISON             Poison control
INFO ROAD               Road conditions / ODOT
INFO LE                 Law enforcement direct lines
INFO JAIL               Jails
INFO FIRE               Fire department admin / BC
INFO ME                 Medical examiner
INFO OTHER              Other useful numbers
INFO <UNIT>             Detailed unit info from server
WHO                     Dispatchers currently online
UR                      Active unit roster
US                      Unit status report (all units)
LO                      Logout session and return to login
! <TEXT>                Search audit/incidents
ADDR                    Show full address directory
ADDR <QUERY>            Search addresses / facilities

═══════════════════════════════════════════════════
PANELS
═══════════════════════════════════════════════════
INBOX                   Open/show message inbox
NOTES / SCRATCH         Open/focus scratch notepad
  (Scratch notes save per-user to your browser)

═══════════════════════════════════════════════════
UNIT OPERATIONS
═══════════════════════════════════════════════════
<STATUS> <UNIT>; <NOTE>    Set unit status with note
<UNIT> <STATUS>; <NOTE>    Alternate syntax
<STATUS>                   Apply to selected row

STATUS CODES: D, DE, OS, F, FD, T, AV, UV, BRK, OOS
  D   = Pending Dispatch (flashing blue)
  DE  = Enroute
  OS  = On Scene
  F   = Follow Up
  FD  = Flagged Down
  T   = Transporting
  AV  = Available
  UV  = Unavailable
  BRK = Break/Lunch
  OOS = Out of Service

Examples:
  D JC; MADRAS ED
  D WC1 0023              Dispatch + assign incident
  EMS1 OS; ON SCENE
  F EMS2; FOLLOW UP NEEDED
  BRK WC1; LUNCH BREAK

DEST <UNIT>; <LOCATION> [NOTE]  Set unit location/destination
  DEST EMS1; SCB           → resolves to ST. CHARLES BEND
  DEST EMS1; SCB [BED 4]  → with room/note
  DEST EMS1; BEND ED      → freeform text
  DEST EMS1               → clears destination

L <UNIT>; <NOTE>        Logon unit (LOGON also works)
LO <UNIT>               Logoff unit (LOGOFF also works)
RIDOFF <UNIT>           Set AV + clear all fields
LUI                     Open logon modal (empty)
LUI <UNIT>              Open logon modal (pre-filled)
UI <UNIT>               Open unit info modal
UNDO <UNIT>             Undo last action

═══════════════════════════════════════════════════
UNIT TIMING (STALE DETECTION)
═══════════════════════════════════════════════════
OK <UNIT>               Touch timer (reset staleness)
OKALL                   Touch all OS units

═══════════════════════════════════════════════════
INCIDENT MANAGEMENT
═══════════════════════════════════════════════════
NC <LOCATION>; <NOTE>; <TYPE>; <PRIORITY>  Create new incident
  Example: NC BEND ED; CHEST PAIN; MED; PRI-2
  Note, type, and priority are optional: NC BEND ED

DE <UNIT> <INC>         Assign queued incident to unit
  Example: DE EMS1 0023

R <INC>                 Review incident + history
  R 0001 (auto-year) or R INC26-0001

U <INC>; <MESSAGE>      Add note to incident
  U 0001; PT IN WTG RM

OK INC<ID>              Touch incident timestamp
LINK <U1> <U2> <INC>    Assign both units to incident
TRANSFER <FROM> <TO> <INC>   Transfer incident
CLOSE <INC>             Manually close incident (full picker)
CLOSE <INC> <CODE>      Close with shortcode — no picker
  CLOSE 0045 TC, CLOSE SC26-0045 PR
<UNIT> AV <CODE>        Go AV and close incident with shortcode
  MALS1 AV TC, EMS2 AV CAN
CAN <INC> [NOTE]        Quick cancel (CANCELLED — no picker)
  CAN SC26-0045, CAN 0045, CAN 0045 PER HOSPITAL
DEL <INC> [DUP|ERR]     Delete with structured reason
  DEL SC26-0045 DUP  (duplicate call)
  DEL SC26-0045 ERR  (data entry error)
  DEL SC26-0045      (opens reason picker)
DISPO SHORTCODES: TC=TRANSPORTED  PR=PATIENT-REFUSED  CAN=CANCELLED-ON-SCENE
  NP=NO-PATIENT-FOUND  MA=MUTUAL-AID-XFER  DUP=DUPLICATE  ERR=DATA-ERROR
RQ <INC>                Reopen incident

═══════════════════════════════════════════════════
UNIT HISTORY
═══════════════════════════════════════════════════
UH <UNIT> [HOURS]       View unit history (alias: HIST)
  UH EMS1 24
  HIST EMS1 24
<UNIT> UH [HOURS]       Alternate syntax
  EMS1 UH 12

═══════════════════════════════════════════════════
REPORTS
═══════════════════════════════════════════════════
REPORTOOS               OOS report (default 24H)
REPORTOOS24H            OOS report for 24 hours
REPORTOOS7D             OOS report for 7 days
REPORTOOS30D            OOS report for 30 days

REPORT SHIFT [H]        Printable shift summary (default 12H)
REPORT INC <ID>         Printable per-incident report
REPORTUTIL <UNIT> [H]   Per-unit utilization report (default 24H)
SUGGEST <INC>           Recommend available units for incident

═══════════════════════════════════════════════════
INCIDENT CREATION (EXTENDED)
═══════════════════════════════════════════════════
NC <DEST>; <NOTE>; <TYPE>; <PRIORITY>; @<SCENE ADDR>
  TYPE format: CAT-NATURE-DET (e.g. MED-CARDIAC-CHARLIE)
  Add "MA" anywhere in NOTE to flag as mutual aid
  Use [CB:PHONE] in NOTE for callback number (e.g. [CB:5415550123])
  PRIORITY = PRI-1 / PRI-2 / PRI-3 / PRI-4
  SCENE ADDR: 5th segment, prefix with @ (e.g. @1234 MAIN ST, BEND)

  Examples:
    NC ST CHARLES; MA 67 YOF CARDIAC [CB:5415550123]; MED-CARDIAC-CHARLIE; PRI-1; @5TH FLOOR TOWER B
    NC BEND RURAL; MVC WITH ENTRAPMENT; TRAUMA-MVA-DELTA; PRI-2
    NC SCMC; IFT CARDIAC; IFT-ALS-CARDIAC; PRI-2; @789 SW CANAL BLVD

═══════════════════════════════════════════════════
MASS OPERATIONS
═══════════════════════════════════════════════════
AVALL <INC#>            Set ALL units on incident → AV (clears assignment)
  AVALL 0071
OSALL <INC#>            Set ALL units on incident → OS (keeps assignment)
  OSALL 0071
MASS D <DEST>           Dispatch all AV units
  MASS D MADRAS ED

═══════════════════════════════════════════════════
BANNERS
═══════════════════════════════════════════════════
NOTE; <MESSAGE>         Set info banner
NOTE; CLEAR             Clear banner
ALERT; <MESSAGE>        Set alert banner (alert tone)
ALERT; CLEAR            Clear alert

═══════════════════════════════════════════════════
MESSAGING SYSTEM
═══════════════════════════════════════════════════
MSG <ROLE/UNIT>; <TEXT> Send normal message
  MSG DP2; NEED COVERAGE AT 1400
  MSG EMS12; CALL ME

HTMSG <ROLE/UNIT>; <TEXT> Send URGENT message (hot)
  HTMSG SUPV1; CALLBACK ASAP

MSGALL; <TEXT>          Broadcast to all active stations
  MSGALL; RADIO CHECK AT 1400

HTALL; <TEXT>           Urgent broadcast to all
  HTALL; SEVERE WEATHER WARNING

PG <UNIT>               Radio page unit (plays fire/EMS tone on field device)
  PG M1
WELF <UNIT>             Welfare check — sends urgent message asking unit to confirm status
  WELF M1
GPS <UNIT>              Show unit on board map using current known position
  GPS M1
GPSUL <UNIT>            Request unit to ping their GPS coords to the board (unit taps or auto-responds)
  GPSUL M1
MSGDP; <TEXT>           Message all dispatchers only
HTDP; <TEXT>            URGENT message all dispatchers
MSGU; <TEXT>            Message all active field units
HTU; <TEXT>             URGENT message all field units

ROLES: DP1-6, SUPV1-2, MGR1-2, EMS, TCRN, PLRN, IT

DEL ALL MSG             Delete all your messages

═══════════════════════════════════════════════════
USER MANAGEMENT
═══════════════════════════════════════════════════
NEWUSER lastname,firstname   Create new user
  NEWUSER smith,john → creates username smithj
  (Default password: 12345)

DELUSER <username>      Delete user
  DELUSER smithj

LISTUSERS               Show all system users
PASSWD <old> <new>      Change your password
  PASSWD 12345 myNewPass

═══════════════════════════════════════════════════
SESSION MANAGEMENT
═══════════════════════════════════════════════════
WHO                     Show logged-in users
LO                      Logout current session
ADMIN                   Admin commands (SUPV/MGR/IT only)

═══════════════════════════════════════════════════
INTERACTION
═══════════════════════════════════════════════════
CLICK ROW               Select unit (yellow outline)
DBLCLICK ROW            Open edit modal
TYPE STATUS CODE        Apply to selected unit
  (e.g. select EMS1, type OS → sets OS)

═══════════════════════════════════════════════════
KEYBOARD SHORTCUTS
═══════════════════════════════════════════════════
CTRL+K / F1 / F3        Focus command bar
CTRL+L                  Open logon modal
CTRL+D                  Cycle density mode
UP/DOWN ARROWS          Command history
ENTER                   Run command
F2                      New incident
F4                      Open messages
ESC                     Close dialogs`);
}

function showAdmin() {
  if (!isAdminRole()) {
    showAlert('ACCESS DENIED', 'ADMIN COMMANDS REQUIRE SUPV, MGR, OR IT LOGIN.');
    return;
  }
  showAlert('ADMIN COMMANDS', `SCMC HOSCAD - ADMIN COMMANDS
ACCESS: SUPV1, SUPV2, MGR1, MGR2, IT

═══════════════════════════════════════════════════
DATA MANAGEMENT
═══════════════════════════════════════════════════
PURGE                   Clean old data (>7 days) + install auto-purge
CLEARDATA UNITS         Clear ALL units from board
CLEARDATA INACTIVE      Clear only inactive/logged-off units
CLEARDATA AUDIT         Clear unit audit history
CLEARDATA INCIDENTS     Clear all incidents
CLEARDATA MESSAGES      Clear all messages
CLEARDATA SESSIONS      Log out all users (force re-login)
CLEARDATA ALL           Clear all data

═══════════════════════════════════════════════════
USER MANAGEMENT
═══════════════════════════════════════════════════
NEWUSER lastname,firstname   Create new user
  (Default password: 12345)
DELUSER <username>      Delete user
LISTUSERS               Show all system users

═══════════════════════════════════════════════════
NOTES
═══════════════════════════════════════════════════
• PURGE automatically runs daily once triggered
• CLEARDATA operations cannot be undone
• CLEARDATA SESSIONS will log you out too`);
}

// ============================================================
// Popout / Secondary Monitor
// ============================================================
// ── CAD Popout Windows (auto-opened on login) ─────────────────────────────
function openPopouts() {
  // Board (viewer)
  if (!_popoutBoardWindow || _popoutBoardWindow.closed) {
    _popoutBoardWindow = window.open('/viewer/', 'hoscad-board',
      'width=1280,height=800,left=0,top=0');
    if (_popoutBoardWindow) {
      monitorPopout(_popoutBoardWindow, 'board');
      // Explicit token relay (same belt-and-suspenders pattern as incident queue)
      function _relayTokenToBoard() {
        if (_popoutBoardWindow && !_popoutBoardWindow.closed && TOKEN) {
          _popoutBoardWindow.postMessage({ type: 'HOSCAD_RELAY_TOKEN', token: TOKEN }, window.location.origin);
        }
      }
      _popoutBoardWindow.addEventListener('load', _relayTokenToBoard);
      window.addEventListener('message', function _relayBoardHandler(e) {
        if (e.origin !== window.location.origin) return;
        if (e.data && e.data.type === 'HOSCAD_REQUEST_RELAY_TOKEN') {
          window.removeEventListener('message', _relayBoardHandler);
          _relayTokenToBoard();
        }
      });
    }
  } else {
    try { _popoutBoardWindow.focus(); } catch(e){}
  }
  // Incident queue
  if (!_popoutIncWindow || _popoutIncWindow.closed) {
    _popoutIncWindow = window.open('/popout-inc/', 'hoscad-inc',
      'width=480,height=950,left=0,top=0');
    if (_popoutIncWindow) {
      monitorPopout(_popoutIncWindow, 'inc');
      // Explicit token relay (same belt-and-suspenders pattern as viewer)
      function _relayTokenToInc() {
        if (_popoutIncWindow && !_popoutIncWindow.closed && TOKEN) {
          _popoutIncWindow.postMessage({ type: 'HOSCAD_RELAY_TOKEN', token: TOKEN }, window.location.origin);
        }
      }
      _popoutIncWindow.addEventListener('load', _relayTokenToInc);
      // Also relay on request (in case load fires before listener is attached)
      window.addEventListener('message', function _relayIncHandler(e) {
        if (e.origin !== window.location.origin) return;
        if (e.data && e.data.type === 'HOSCAD_REQUEST_RELAY_TOKEN') {
          window.removeEventListener('message', _relayIncHandler);
          _relayTokenToInc();
        }
      });
    }
  } else {
    try { _popoutIncWindow.focus(); } catch(e){}
  }
  if (!_popoutBoardWindow && !_popoutIncWindow) {
    showToast('POPUPS BLOCKED — TYPE BOARDS TO RETRY. ALLOW POPUPS FOR THIS SITE.', 'warn');
  } else {
    showToast('BOARDS OPENED.');
  }
}

function openPopoutInc() {
  if (_popoutIncWindow && !_popoutIncWindow.closed) {
    _popoutIncWindow.focus();
    showToast('INCIDENT QUEUE ALREADY OPEN.');
    return;
  }
  _popoutIncWindow = window.open('/popout-inc/', 'hoscad-inc',
    'width=480,height=950,left=0,top=0');
  if (!_popoutIncWindow) {
    showToast('POPUP BLOCKED — ALLOW POPUPS FOR THIS SITE.', 'warn');
    return;
  }
  monitorPopout(_popoutIncWindow, 'inc');
  function _relayTokenToInc() {
    if (_popoutIncWindow && !_popoutIncWindow.closed && TOKEN) {
      _popoutIncWindow.postMessage({ type: 'HOSCAD_RELAY_TOKEN', token: TOKEN }, window.location.origin);
    }
  }
  _popoutIncWindow.addEventListener('load', _relayTokenToInc);
  window.addEventListener('message', function _relayIncHandler(e) {
    if (e.origin !== window.location.origin) return;
    if (e.data && e.data.type === 'HOSCAD_REQUEST_RELAY_TOKEN') {
      window.removeEventListener('message', _relayIncHandler);
      _relayTokenToInc();
    }
  });
  showToast('INCIDENT QUEUE OPENED.');
}

function monitorPopout(win, name) {
  const check = setInterval(function() {
    if (win.closed) {
      clearInterval(check);
      if (name === 'board') _popoutBoardWindow = null;
      if (name === 'inc')   _popoutIncWindow   = null;
    }
  }, 3000);
}

function openPopout() {
  if (_popoutBoardWindow && !_popoutBoardWindow.closed) {
    _popoutBoardWindow.focus();
    showToast('BOARD ALREADY ON SECONDARY MONITOR.');
    return;
  }
  _popoutBoardWindow = window.open('/viewer/', 'hoscad-board', 'width=1280,height=800');
  if (!_popoutBoardWindow) {
    showErr({ error: 'POPUP BLOCKED. ALLOW POPUPS FOR THIS SITE.' });
    return;
  }
  monitorPopout(_popoutBoardWindow, 'board');
  // Relay token to viewer
  function _relayTokenToViewer() {
    if (_popoutBoardWindow && !_popoutBoardWindow.closed && TOKEN) {
      _popoutBoardWindow.postMessage({ type: 'HOSCAD_RELAY_TOKEN', token: TOKEN }, window.location.origin);
    }
  }
  _popoutBoardWindow.addEventListener('load', _relayTokenToViewer);
  // Show placeholder on main board
  const boardEl = document.getElementById('boardMain');
  const popoutPlaceholder = document.getElementById('popoutPlaceholder');
  if (boardEl) boardEl.style.display = 'none';
  if (popoutPlaceholder) popoutPlaceholder.style.display = 'flex';
  showToast('BOARD OPENED ON SECONDARY MONITOR.');
  // Poll for window close to auto-restore
  const closeCheck = setInterval(function() {
    if (_popoutBoardWindow && _popoutBoardWindow.closed) {
      clearInterval(closeCheck);
      closePopin();
    }
  }, 2000);
}

function closePopin() {
  if (_popoutBoardWindow && !_popoutBoardWindow.closed) _popoutBoardWindow.close();
  _popoutBoardWindow = null;
  const boardEl = document.getElementById('boardMain');
  const popoutPlaceholder = document.getElementById('popoutPlaceholder');
  if (boardEl) boardEl.style.display = '';
  if (popoutPlaceholder) popoutPlaceholder.style.display = 'none';
  showToast('BOARD RESTORED.');
}

function updatePopoutClock() {
  const el = document.getElementById('popoutClock');
  if (!el) return;
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  el.textContent = hh + ':' + mm;
  const dateEl = document.getElementById('popoutDate');
  if (dateEl) {
    const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    dateEl.textContent = days[now.getDay()] + ', ' + now.getDate() + ' ' + months[now.getMonth()] + ' ' + now.getFullYear();
  }
}

function updatePopoutStats() {
  const el = document.getElementById('popoutStats');
  if (!el || !STATE) return;
  const activeUnits = (STATE.units || []).filter(u => u.active).length;
  const queued = (STATE.incidents || []).filter(i => i.status === 'QUEUED').length;
  el.textContent = activeUnits + ' UNITS ACTIVE  ·  ' + queued + ' QUEUED';
}

// ════════════════════════════════════════════════════════════════════════════
// ── Search Panel (F3 / Ctrl+F) ───────────────────────────────────────────
// ════════════════════════════════════════════════════════════════════════════

let _searchPanelTimer = null;

function openSearchPanel() {
  const el = document.getElementById('searchBack');
  if (!el) return;
  el.style.display = 'flex';
  const inp = document.getElementById('searchPanelInput');
  if (inp) { inp.value = ''; inp.focus(); }
  const results = document.getElementById('searchPanelResults');
  if (results) results.innerHTML = '<div class="muted" style="padding:20px 16px;text-align:center;font-size:12px;">TYPE TO SEARCH ADDRESSES AND CALLS</div>';
}

function closeSearchPanel() {
  const el = document.getElementById('searchBack');
  if (el) el.style.display = 'none';
  autoFocusCmd();
}

function _searchPanelDebounce() {
  if (_searchPanelTimer) clearTimeout(_searchPanelTimer);
  _searchPanelTimer = setTimeout(_doSearchPanel, 350);
}

function _searchPanelKeydown(e) {
  if (e.key === 'Escape') { e.preventDefault(); closeSearchPanel(); }
}

function _doSearchPanel() {
  const inp = document.getElementById('searchPanelInput');
  if (!inp) return;
  const q = inp.value.trim().toUpperCase();
  const results = document.getElementById('searchPanelResults');
  if (!results) return;
  if (q.length < 2) {
    results.innerHTML = '<div class="muted" style="padding:20px 16px;text-align:center;font-size:12px;">TYPE TO SEARCH ADDRESSES AND CALLS</div>';
    return;
  }

  let html = '';

  // ── Address directory (client-side, instant) ──
  const addrs = AddressLookup.search ? AddressLookup.search(q) : [];
  if (addrs.length) {
    html += '<div style="padding:6px 16px 4px;font-size:10px;font-weight:900;letter-spacing:.08em;color:var(--muted);border-bottom:1px solid var(--line);">ADDRESS DIRECTORY (' + addrs.length + ')</div>';
    addrs.slice(0, 8).forEach(a => {
      const addrJ = JSON.stringify(a.addr || a.address || a.name || '');
      html += '<div class="search-result-row">' +
        '<div class="search-result-label">' +
          '<span style="font-weight:900;">' + esc(a.code || '') + '</span>' +
          '<span style="color:var(--muted);font-size:11px;margin-left:6px;">' + esc(a.addr || a.address || a.name || '') + '</span>' +
        '</div>' +
        '<div class="search-result-actions">' +
          '<button class="btn-sm" onclick="searchPanelUse(' + addrJ + ')">USE</button>' +
          '<button class="btn-sm" onclick="searchPanelNc(' + addrJ + ')">NC→</button>' +
        '</div>' +
      '</div>';
    });
  }

  // ── Incident search (from current STATE) ──
  const incs = ((STATE && STATE.incidents) ? STATE.incidents : [])
    .filter(inc => {
      const id = (inc.incident_id || '').toUpperCase();
      const addr = (inc.scene_address || '').toUpperCase();
      const dest = (inc.destination || '').toUpperCase();
      return id.includes(q) || addr.includes(q) || dest.includes(q);
    })
    .sort((a, b) => new Date(_normalizeTs(b.created_at)) - new Date(_normalizeTs(a.created_at)))
    .slice(0, 8);

  if (incs.length) {
    html += '<div style="padding:6px 16px 4px;font-size:10px;font-weight:900;letter-spacing:.08em;color:var(--muted);border-bottom:1px solid var(--line);margin-top:4px;">INCIDENTS (' + incs.length + ')</div>';
    incs.forEach(inc => {
      const addrText = inc.scene_address || inc.destination || '—';
      const type = inc.incident_type || '';
      const status = inc.status || '';
      const statusColor = status === 'ACTIVE' ? 'var(--green)' : status === 'QUEUED' ? 'var(--yellow)' : 'var(--muted)';
      const incIdJ = JSON.stringify(inc.incident_id);
      const addrJ = JSON.stringify(inc.scene_address || '');
      html += '<div class="search-result-row">' +
        '<div class="search-result-label">' +
          '<span style="font-weight:900;color:var(--yellow);">' + esc(inc.incident_id) + '</span>' +
          (type ? '<span style="font-size:10px;color:var(--muted);margin-left:6px;">' + esc(type) + '</span>' : '') +
          ' <span style="font-size:10px;font-weight:900;color:' + statusColor + ';">' + esc(status) + '</span>' +
          '<br><span style="font-size:11px;">' + esc(addrText) + '</span>' +
        '</div>' +
        '<div class="search-result-actions">' +
          '<button class="btn-sm" onclick="openIncident(' + incIdJ + ');closeSearchPanel()">OPEN</button>' +
          (status !== 'ACTIVE' && inc.scene_address ? '<button class="btn-sm" onclick="searchPanelNc(' + addrJ + ')">NC→</button>' : '') +
        '</div>' +
      '</div>';
    });
  }

  if (!html) {
    html = '<div class="muted" style="padding:20px 16px;text-align:center;font-size:12px;">NO RESULTS FOR "' + esc(q) + '"</div>';
  }
  results.innerHTML = html;
}

function searchPanelUse(addr) {
  const inp = document.getElementById('cmd');
  if (inp) { inp.value = addr; inp.focus(); }
  closeSearchPanel();
}

function searchPanelNc(addr) {
  closeSearchPanel();
  openNewIncident(addr);
}

// ════════════════════════════════════════════════════════════════════════════
// ── LifeFlight ADS-B Feed ────────────────────────────────────────────────
// ════════════════════════════════════════════════════════════════════════════

// Haversine distance in nautical miles between two lat/lon points
function _lfnDistNm(lat1, lon1, lat2, lon2) {
  const R = 3440.065;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) *
    Math.sin(dLon/2) * Math.sin(dLon/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// Bearing in degrees (0-360) from point 1 to point 2
function _lfnBearing(lat1, lon1, lat2, lon2) {
  const toR = x => x * Math.PI / 180;
  const dLon = toR(lon2 - lon1);
  const y = Math.sin(dLon) * Math.cos(toR(lat2));
  const x = Math.cos(toR(lat1)) * Math.sin(toR(lat2)) -
    Math.sin(toR(lat1)) * Math.cos(toR(lat2)) * Math.cos(dLon);
  return (Math.atan2(y, x) * 180 / Math.PI + 360) % 360;
}

// Classify a single aircraft's status from raw ADS-B fields
function _lfnClassify(tail, ac) {
  if (!ac) return 'NOSIG';
  const seenPos = ac.seen_pos != null ? ac.seen_pos : (ac.seen != null ? ac.seen : 9999);
  if (seenPos > 600) return 'NOSIG';
  const lat = ac.lat != null ? ac.lat : null;
  const lon = ac.lon != null ? ac.lon : null;
  const onGround = ac.ground === '1' || ac.ground === true ||
    ac.alt_baro === 'ground' || ac.on_ground === true;
  if (onGround || lat == null) {
    if (lat != null && _lfnDistNm(lat, lon, LFN_BASE_LAT, LFN_BASE_LON) <= 3) return 'GND';
    return onGround ? 'LDG' : 'GND';
  }
  // Airborne — check INBOUND to nearest SCMC helipad (non-airport) within 15nm
  const heading = ac.track != null ? ac.track : (ac.true_heading != null ? ac.true_heading : null);
  const altBaro = typeof ac.alt_baro === 'number' ? ac.alt_baro : null;
  let nearestScmc = null, nearestDist = Infinity;
  for (const hp of LFN_HELIPADS) {
    if (!hp.scmc) continue;  // INBOUND alert only for SCMC receiving hospitals
    const d = _lfnDistNm(lat, lon, hp.lat, hp.lon);
    if (d < nearestDist) { nearestDist = d; nearestScmc = hp; }
  }
  if (nearestScmc && nearestDist < 15 && heading != null && altBaro != null) {
    const bear = _lfnBearing(lat, lon, nearestScmc.lat, nearestScmc.lon);
    const hdgDelta = Math.abs(((heading - bear + 540) % 360) - 180);
    const descending = altBaro < (_lfnPrevAlt[tail] != null ? _lfnPrevAlt[tail] : altBaro + 9999) - 200;
    if (hdgDelta <= 45 && descending) {
      if (_lfnPrevInbd[tail]) return 'INBD'; // 2-poll debounce confirmed
      _lfnPrevInbd[tail] = true;
      return 'AIR'; // candidate; confirm next poll
    }
  }
  _lfnPrevInbd[tail] = false;
  return 'AIR';
}

// Return the active board unit entry for a given tail number (if mapped and active)
function getLfnBoardUnit(tail) {
  const uid = AIR_FLEET[tail] && AIR_FLEET[tail].unitId;
  if (!uid || !STATE) return null;
  return (STATE.units || []).find(u => u.unit_id === uid && u.active) || null;
}

// Process ADS-B feed — classify aircraft, fire INBOUND alerts
function applyLfnFeed(acList) {
  const newAircraft = [];
  for (const tail of Object.keys(AIR_FLEET)) {
    const fleetInfo = AIR_FLEET[tail];
    const ac = acList.find(a => (a.r || '').toUpperCase() === tail) || null;
    const prevEntry = _lfnAircraft.find(x => x.tail === tail);
    const prevStatus = prevEntry ? prevEntry.status : null;
    const status = _lfnClassify(tail, ac);
    // Update prev altitude before storing (used by _lfnClassify next poll)
    if (ac && typeof ac.alt_baro === 'number') _lfnPrevAlt[tail] = ac.alt_baro;
    const entry = {
      tail,
      callsign: fleetInfo.callsign,
      unitId:   fleetInfo.unitId,
      type:     fleetInfo.type,
      provider: fleetInfo.provider,
      status,
      lat:    ac && ac.lat != null ? ac.lat : null,
      lon:    ac && ac.lon != null ? ac.lon : null,
      alt_ft: ac && typeof ac.alt_baro === 'number' ? Math.round(ac.alt_baro) : null,
      gs_kts: ac && ac.gs != null ? Math.round(ac.gs) : null,
      heading: ac && ac.track != null ? ac.track : (ac && ac.true_heading != null ? ac.true_heading : null),
      seenPos: ac && ac.seen_pos != null ? ac.seen_pos : (ac && ac.seen != null ? ac.seen : null),
    };
    newAircraft.push(entry);
    // Fire INBOUND alert — once per event, suppress repeats
    if (status === 'INBD' && prevStatus !== 'INBD' && !_lfnInboundAlerted[tail]) {
      _lfnInboundAlerted[tail] = true;
      beepAlert();
      // Auto-expand the panel
      const body   = document.getElementById('lfnPanelBody');
      const toggle = document.getElementById('lfnPanelToggle');
      if (body && body.style.display === 'none') {
        body.style.display = 'flex';
        if (toggle) toggle.textContent = '▲';
      }
    }
    // Clear alert suppression when aircraft leaves INBOUND
    if (status !== 'INBD' && prevStatus === 'INBD') {
      _lfnInboundAlerted[tail] = false;
    }
  }
  _lfnAircraft = newAircraft;
  renderLfnPanel();
  if (document.body.classList.contains('board-map-open')) renderBoardMap();
}

// ADS-B proxy URL — served via GitHub Pages from hoscad-source repo (avoids CORS)
const LFN_PROXY_URL = 'https://ckholden.github.io/hoscad-source/adsb_data.json';

// Fetch ADS-B data via proxy
async function fetchLfnFeed() {
  if (_lfnSyncing || !TOKEN) return;
  _lfnSyncing = true;
  try {
    const res = await fetch(LFN_PROXY_URL + '?_t=' + Date.now());
    if (!res.ok) { console.warn('[LFN] ADS-B fetch error:', res.status); return; }
    const data = await res.json();
    const fleet = (data.ac || []).filter(ac => AIR_FLEET[(ac.r || '').toUpperCase()]);
    applyLfnFeed(fleet);
    _lfnLastSync = new Date();
    updateLfnSyncBadge();
  } catch (e) {
    console.warn('[LFN] fetchLfnFeed error:', e);
  } finally {
    _lfnSyncing = false;
  }
}

function startLfnPolling() {
  if (_lfnPollTimer) return;
  // Pre-seed all fleet entries at NOSIG so every known callsign shows from login
  if (_lfnAircraft.length === 0) {
    _lfnAircraft = Object.entries(AIR_FLEET).map(function([tail, info]) {
      return { tail, callsign: info.callsign, unitId: info.unitId, type: info.type,
               provider: info.provider, status: 'NOSIG',
               lat: null, lon: null, alt_ft: null, gs_kts: null, heading: null, seenPos: null };
    });
  }
  renderLfnPanel();
  fetchLfnFeed();
  _lfnPollTimer = setInterval(fetchLfnFeed, LFN_POLL_INTERVAL);
}

function stopLfnPolling() {
  if (_lfnPollTimer) { clearInterval(_lfnPollTimer); _lfnPollTimer = null; }
}

function updateLfnSyncBadge() {
  const el = document.getElementById('lfnSyncBadge');
  if (!el) return;
  if (!_lfnLastSync) { el.textContent = 'ADS-B: --'; el.style.opacity = '.4'; el.style.color = 'var(--muted)'; return; }
  const secs = Math.floor((Date.now() - _lfnLastSync.getTime()) / 1000);
  const age  = secs < 60 ? secs + 'S' : Math.floor(secs / 60) + 'M';
  const inbd = _lfnAircraft.filter(a => a.status === 'INBD').length;
  const air  = _lfnAircraft.filter(a => a.status === 'AIR' || a.status === 'INBD').length;
  el.textContent = 'ADS-B: ' + age + (air > 0 ? ' (' + air + ' AIR)' : '');
  el.style.opacity = secs > 120 ? '.4' : '1';
  el.style.color   = inbd > 0 ? '#f59e0b' : air > 0 ? '#22c55e' : '#4fa3e0';
}

// ════════════════════════════════════════════════════════════════════════════
// ── DC911 CadView Integration ────────────────────────────────────────────
// ════════════════════════════════════════════════════════════════════════════

function _isDc911Enabled() {
  return localStorage.getItem('hoscad_dc911_enabled') !== 'false';
}

function toggleDc911() {
  const next = !_isDc911Enabled();
  localStorage.setItem('hoscad_dc911_enabled', next ? 'true' : 'false');
  _applyDc911Btn(next);
  renderBoardDiff(STATE);
  showToast('DC911 FEED ' + (next ? 'VISIBLE.' : 'HIDDEN.'));
}

function _applyDc911Btn(enabled) {
  const btn = document.getElementById('btnToggleDC911');
  if (!btn) return;
  btn.textContent = enabled ? 'DC911 ON' : 'DC911 OFF';
  btn.style.opacity = enabled ? '1' : '0.5';
}

function updateDc911SyncBadge() {
  const el = document.getElementById('dc911SyncBadge');
  if (!el) return;
  const updatedAt = STATE && STATE.dc911State && STATE.dc911State.updatedAt;
  if (!updatedAt) { el.textContent = 'DC911: --'; el.style.opacity = '.4'; el.style.color = 'var(--muted)'; return; }
  _dc911LastSync = new Date(updatedAt);
  const secs = Math.floor((Date.now() - _dc911LastSync.getTime()) / 1000);
  if (secs < 120) {
    el.textContent = 'DC911: ' + secs + 's';
    el.style.opacity = '1';
    el.style.color = '#4fa3e0';
  } else if (secs < 600) {
    el.textContent = 'DC911: ' + Math.floor(secs / 60) + 'm';
    el.style.opacity = '0.8';
    el.style.color = '#f59e0b';
  } else {
    el.textContent = 'DC911: STALE';
    el.style.opacity = '0.7';
    el.style.color = '#ef4444';
  }
}

// Render the AIR RESOURCES panel aircraft cards
function renderLfnPanel() {
  const panel      = document.getElementById('lfnPanel');
  const body       = document.getElementById('lfnPanelBody');
  const alertBadge = document.getElementById('lfnAlertBadge');
  if (!panel || !body) return;
  if (!TOKEN) { panel.style.display = 'none'; return; }
  panel.style.display = '';
  // INBOUND alert badge in panel header
  const inbdList = _lfnAircraft.filter(a => a.status === 'INBD');
  if (alertBadge) {
    if (inbdList.length > 0) {
      alertBadge.style.display = '';
      alertBadge.textContent   = 'INBOUND: ' + inbdList.map(a => a.callsign).join(', ');
    } else {
      alertBadge.style.display = 'none';
    }
  }
  const statusLabel = { INBD:'INBOUND', AIR:'AIRBORNE', LDG:'LANDED', GND:'BASE', NOSIG:'NO SIG' };
  const statusColor = { INBD:'#f59e0b', AIR:'#22c55e', LDG:'#3b82f6', GND:'#666', NOSIG:'#444' };
  const providerLabel = { LFN:'LIFEFLIGHT NETWORK', AIRLINK:'AIRLINK CCT' };

  function buildCard(ac) {
    const boardUnit = getLfnBoardUnit(ac.tail);
    const lbl   = statusLabel[ac.status] || ac.status;
    const scl   = 'lfn-ac status-' + (ac.status || 'nosig').toLowerCase();
    const scol  = statusColor[ac.status] || '#444';
    let detail  = '';
    if (ac.status === 'AIR' || ac.status === 'INBD') {
      const altS = ac.alt_ft  != null ? ac.alt_ft.toLocaleString() + 'ft' : '\u2014';
      const spdS = ac.gs_kts  != null ? ac.gs_kts + 'kts'                 : '\u2014';
      const hdgS = ac.heading != null ? String(Math.round(ac.heading)).padStart(3,'0') + '\u00b0' : '\u2014';
      detail = altS + ' \u00b7 ' + spdS + ' \u00b7 HDG ' + hdgS;
      if (ac.status === 'INBD' && ac.gs_kts > 0 && ac.lat != null) {
        let nearDist = Infinity;
        LFN_HELIPADS.forEach(function(hp) {
          if (hp.scmc) { const d = _lfnDistNm(ac.lat, ac.lon, hp.lat, hp.lon); if (d < nearDist) nearDist = d; }
        });
        if (nearDist < Infinity) {
          const etaMin = Math.max(1, Math.round(nearDist / ac.gs_kts * 60));
          detail = 'ETA ~' + etaMin + 'MIN \u00b7 ' + detail;
        }
      }
    } else if (ac.status === 'LDG')   { detail = 'LANDED AWAY'; }
    else if (ac.status === 'GND')     { detail = 'AT BASE';     }
    else if (ac.status === 'NOSIG')   { detail = 'NO ADS-B SIGNAL'; }
    const boardHtml = boardUnit
      ? '<span class="lfn-board-link">\u2192 ' + esc(ac.callsign) +
        (boardUnit.incident ? ' | ' + boardUnit.incident : '') +
        '</span>'
      : '';
    return '<div class="' + scl + '">' +
      '<span class="lfn-callsign">' + esc(ac.callsign) + '</span>' +
      '<span class="lfn-status-text" style="color:' + scol + '">' + esc(lbl) + '</span>' +
      '<span class="lfn-detail">' + esc(detail) + '</span>' +
      boardHtml + '</div>';
  }

  // Group aircraft by provider, render each group with a header
  const providers = ['LFN', 'AIRLINK'];
  let html = '';
  for (const prov of providers) {
    const group = _lfnAircraft.filter(function(a) { return a.provider === prov; });
    if (!group.length) continue;
    html += '<div class="lfn-provider-group">' +
      '<div class="lfn-provider-label">' + esc(providerLabel[prov] || prov) + '</div>' +
      '<div class="lfn-provider-cards">' + group.map(buildCard).join('') + '</div>' +
      '</div>';
  }
  body.innerHTML = html ||
    '<span style="font-size:11px;color:#444;padding:4px 0;">NO AIR RESOURCES IN RANGE</span>';
}

// Heading-rotated SVG triangle icon for aircraft map markers
function getLfnMapIcon(heading, status) {
  const colors = { INBD:'#f59e0b', AIR:'#22c55e', LDG:'#3b82f6', GND:'#888', NOSIG:'#444' };
  const color  = colors[status] || '#888';
  const svg    = '<svg width="20" height="20" viewBox="0 0 20 20">' +
    '<polygon points="10,2 16,18 10,14 4,18" fill="' + color + '" stroke="#fff" stroke-width="1"/></svg>';
  return L.divIcon({
    html: '<div style="transform:rotate(' + (heading || 0) + 'deg);width:20px;height:20px;">' + svg + '</div>',
    className: '', iconSize: [20,20], iconAnchor: [10,10]
  });
}

// SVG divIcon for unit markers — shape encodes type, color encodes status
function getUnitMapIcon(unitType, color) {
  const t = (unitType || '').toUpperCase();
  let svg, w = 16, h = 16;
  if (t === 'EMS') {
    // Ambulance body with cross
    svg = '<svg width="16" height="16" viewBox="0 0 16 16">' +
      '<rect x="1" y="4" width="14" height="9" rx="1.5" fill="' + color + '" stroke="#fff" stroke-width="1"/>' +
      '<rect x="7" y="6" width="2" height="5" fill="#fff"/>' +
      '<rect x="5" y="8" width="6" height="2" fill="#fff"/>' +
      '</svg>';
  } else if (t === 'FIRE') {
    // Wide body with stepped cab
    w = 18; h = 14;
    svg = '<svg width="18" height="14" viewBox="0 0 18 14">' +
      '<rect x="1" y="5" width="16" height="7" rx="1" fill="' + color + '" stroke="#fff" stroke-width="1"/>' +
      '<rect x="1" y="2" width="7" height="5" rx="1" fill="' + color + '" stroke="#fff" stroke-width="1"/>' +
      '</svg>';
  } else if (t === 'AIR') {
    // Diamond — rotor wing shape
    svg = '<svg width="16" height="16" viewBox="0 0 16 16">' +
      '<polygon points="8,1 15,8 8,15 1,8" fill="' + color + '" stroke="#fff" stroke-width="1"/>' +
      '</svg>';
  } else if (t === 'LAW') {
    // Five-pointed star
    svg = '<svg width="16" height="16" viewBox="0 0 16 16">' +
      '<polygon points="8,1 10,6 15,6 11,9.5 12.5,15 8,12 3.5,15 5,9.5 1,6 6,6" fill="' + color + '" stroke="#fff" stroke-width="1"/>' +
      '</svg>';
  } else if (t === 'SUPV') {
    // Circle with inner ring
    svg = '<svg width="16" height="16" viewBox="0 0 16 16">' +
      '<circle cx="8" cy="8" r="7" fill="' + color + '" stroke="#fff" stroke-width="1.5"/>' +
      '<circle cx="8" cy="8" r="3.5" fill="none" stroke="#fff" stroke-width="1.5"/>' +
      '</svg>';
  } else {
    // Default fallback: plain circle
    w = 14; h = 14;
    svg = '<svg width="14" height="14" viewBox="0 0 14 14">' +
      '<circle cx="7" cy="7" r="6" fill="' + color + '" stroke="#fff" stroke-width="1"/>' +
      '</svg>';
  }
  return L.divIcon({
    html: '<div style="line-height:0;">' + svg + '</div>',
    className: '', iconSize: [w, h], iconAnchor: [Math.round(w/2), Math.round(h/2)]
  });
}

function toggleLfnPanel() {
  const body   = document.getElementById('lfnPanelBody');
  const toggle = document.getElementById('lfnPanelToggle');
  if (!body) return;
  const isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : 'flex';
  if (toggle) toggle.textContent = isOpen ? '\u25bc' : '\u25b2';
}

// ════════════════════════════════════════════════════════════════════════════
// ── Board Map (inline panel + POPMAP popout) ──────────────────────────────
// ════════════════════════════════════════════════════════════════════════════

// Per-unit home base coords keyed by unit_id — used for map positioning (priority 3)
const BM_UNIT_HOME = {
  // La Pine Fire
  'M171': [43.6679, -121.5036], 'M172': [43.8208, -121.4421], 'C171': [43.6679, -121.5036],
  // Sunriver Fire
  'M271': [43.8758, -121.4363], 'C271': [43.8758, -121.4363],
  // Bend Fire
  'M371': [44.0489, -121.3153], 'M372': [44.0550, -121.2679], 'C371': [44.0618, -121.2995],
  // Redmond Fire
  'M471': [44.2783, -121.1785], 'M472': [44.2590, -121.1617], 'C471': [44.2783, -121.1785],
  // Crook County Fire
  'M571': [44.3050, -120.7271], 'M572': [44.3050, -120.7271], 'C571': [44.3050, -120.7271],
  // Cloverdale
  'M671': [44.2225, -121.4090], 'C671': [44.2225, -121.4090],
  // Sisters-Camp Sherman
  'M771': [44.2879, -121.5490], 'M772': [44.2879, -121.5490], 'C771': [44.2879, -121.5490],
  // Black Butte Ranch
  'M871': [44.3988, -121.6282], 'C871': [44.3988, -121.6282],
  // Alfalfa
  'M971': [44.0084, -121.1091], 'C971': [44.0084, -121.1091],
  // Crescent
  'M1171': [43.4644, -121.7168], 'M1172': [43.4743, -121.6937], 'C1171': [43.4644, -121.7168],
  // Prineville
  'M1271': [44.3050, -120.7271], 'M1272': [44.3050, -120.7271], 'C1271': [44.3050, -120.7271],
  // Three Rivers
  'M1371': [44.3524, -121.1781], 'C1371': [44.3524, -121.1781],
  // Jefferson County
  'M1771': [44.6289, -121.1278], 'M1772': [44.6289, -121.1278], 'C1771': [44.6289, -121.1278],
  // Warm Springs
  'M2271': [44.7640, -121.2660], 'M2272': [44.7640, -121.2660], 'C2271': [44.7640, -121.2660],
  // AirLink CCT
  'AL1': [44.0613, -121.2832], 'AL2': [44.0946, -121.2000], 'AL4': [42.1561, -121.7339], 'ALG1': [44.0946, -121.2000],
  // Life Flight
  'LF11': [44.2570, -121.1599], 'LF45': [45.2385, -122.7548],
  // OHSU PANDA
  'PANDA1': [45.4990, -122.6855],
  // AMR
  'AMRP1': [44.0613, -121.2832],
  // Adventure Medics
  'AMALS1': [44.1024, -121.2799], 'AMBLS1': [44.1024, -121.2799], 'AMSTR1': [44.1024, -121.2799], 'AMWC1': [44.1024, -121.2799],
  // Law Enforcement
  'BPD301': [44.0630, -121.2975], 'BPD320': [44.0630, -121.2975],
  'DCSO214': [44.0584, -121.3615], 'DCSO250': [44.0584, -121.3615],
  'RPD401': [44.2715, -121.1752],
  'OSP315': [44.0882, -121.2870], 'OSP422': [44.0882, -121.2870],
  // ODOT
  'ODOT-R4-1': [44.1045, -121.2987], 'ODOT-R4-2': [44.1045, -121.2987],
};
// Station bases for MAP STA command — unique fire/EMS station locations
const BM_STATION_BASES = {
  'LA PINE STA101':    [43.6679, -121.5036],
  'LA PINE STA102':    [43.8208, -121.4421],
  'SUNRIVER':          [43.8758, -121.4363],
  'BEND STA301':       [44.0489, -121.3153],
  'BEND STA304':       [44.0550, -121.2679],
  'BEND STA306':       [44.0618, -121.2995],
  'REDMOND STA401':    [44.2783, -121.1785],
  'REDMOND STA403':    [44.2590, -121.1617],
  'CROOK STA1201':     [44.3050, -120.7271],
  'CLOVERDALE':        [44.2225, -121.4090],
  'SISTERS STA701':    [44.2879, -121.5490],
  'BLACK BUTTE':       [44.3988, -121.6282],
  'ALFALFA':           [44.0084, -121.1091],
  'CRESCENT STA1':     [43.4644, -121.7168],
  'CRESCENT STA2':     [43.4743, -121.6937],
  'THREE RIVERS':      [44.3524, -121.1781],
  'JEFFCO MADRAS':     [44.6289, -121.1278],
  'WARM SPRINGS':      [44.7640, -121.2660],
  'AIRLINK SCMC':      [44.0613, -121.2832],
  'AIRLINK KBDN':      [44.0946, -121.2000],
  'LIFE FLIGHT RDM':   [44.2570, -121.1599],
  'ADV MEDICS':        [44.1024, -121.2799],
};
const BM_STATUS_COLORS = {
  AV: '#4caf50', D: '#ff9800', DE: '#ff9800',
  OS: '#f44336', T: '#ffd700', AT: '#2196f3', OOS: '#607d8b',
};
const BM_MAP_CENTER  = [44.05, -120.85];
const BM_MAP_ZOOM    = 9;
const BM_MAP_VIEWBOX = '-122.5,43.0,-119.5,45.5';
const BM_TRICOUNTY   = [[43.3, -122.0], [45.0, -119.4]];

// Known address coordinates — pre-seeded into geocache so Nominatim is never
// called for these (Nominatim sometimes returns wrong results for medical facilities).
const BM_KNOWN_COORDS = {
  // St. Charles Bend (2500 NE Neff Rd) — many address variants dispatchers may type
  'ST CHARLES BEND':                         [44.0672, -121.2690],
  'SAINT CHARLES BEND':                      [44.0672, -121.2690],
  'ST CHARLES MEDICAL CENTER':               [44.0672, -121.2690],
  'ST CHARLES MEDICAL CENTER BEND':          [44.0672, -121.2690],
  'SCMC BEND':                               [44.0672, -121.2690],
  'SCMC':                                    [44.0672, -121.2690],
  '2500 NE NEFF RD':                         [44.0672, -121.2690],
  '2500 NE NEFF RD BEND':                    [44.0672, -121.2690],
  '2500 NE NEFF RD, BEND':                   [44.0672, -121.2690],
  '2500 NE NEFF RD BEND OR':                 [44.0672, -121.2690],
  '2500 NE NEFF RD, BEND OR':               [44.0672, -121.2690],
  '2500 NE NEFF RD, BEND, OR':              [44.0672, -121.2690],
  '2500 NE NEFF RD BEND, OR':               [44.0672, -121.2690],
  // St. Charles Redmond
  'ST CHARLES REDMOND':                      [44.2704, -121.1417],
  'SAINT CHARLES REDMOND':                   [44.2704, -121.1417],
  'SCMC REDMOND':                            [44.2704, -121.1417],
  '1253 N CANAL BLVD':                       [44.2704, -121.1417],
  '1253 N CANAL BLVD REDMOND':               [44.2704, -121.1417],
  '1253 N CANAL BLVD, REDMOND':              [44.2704, -121.1417],
  '1253 N CANAL BLVD REDMOND OR':            [44.2704, -121.1417],
  '1253 NORTH CANAL BLVD REDMOND':           [44.2704, -121.1417],
  // St. Charles Prineville
  'ST CHARLES PRINEVILLE':                   [44.2997, -120.8367],
  'SAINT CHARLES PRINEVILLE':                [44.2997, -120.8367],
  'SCMC PRINEVILLE':                         [44.2997, -120.8367],
  '384 SE COMBS FLAT RD':                    [44.2997, -120.8367],
  '384 SE COMBS FLAT RD PRINEVILLE':         [44.2997, -120.8367],
  '384 SE COMBS FLAT RD, PRINEVILLE':        [44.2997, -120.8367],
  '384 SE COMBS FLAT RD PRINEVILLE OR':      [44.2997, -120.8367],
  '384 SOUTHEAST COMBS FLAT RD PRINEVILLE':  [44.2997, -120.8367],
  // St. Charles Madras
  'ST CHARLES MADRAS':                       [44.6329, -121.1298],
  'SAINT CHARLES MADRAS':                    [44.6329, -121.1298],
  'SCMC MADRAS':                             [44.6329, -121.1298],
  '470 NE A ST':                             [44.6329, -121.1298],
  '470 NE A ST MADRAS':                      [44.6329, -121.1298],
  '470 NE A ST, MADRAS':                     [44.6329, -121.1298],
  '470 NE A ST MADRAS OR':                   [44.6329, -121.1298],
  '470 NORTHEAST A ST MADRAS':               [44.6329, -121.1298],
  // Warm Springs IHS
  'WARM SPRINGS IHS':                        [44.7636, -121.2733],
  'WARM SPRINGS HEALTH CENTER':              [44.7636, -121.2733],
  // Sky Lakes Medical Center (Klamath Falls — mutual aid)
  'SKY LAKES MEDICAL CENTER':                [42.2530, -121.7851],
  'SKY LAKES ED':                            [42.2530, -121.7851],
  '2865 DAGGETT AVE KLAMATH FALLS':          [42.2530, -121.7851],
  // Mid-Columbia Medical Center (The Dalles — mutual aid)
  'MID-COLUMBIA MEDICAL CENTER':             [45.5980, -121.1525],
  'MCMC ED':                                 [45.5980, -121.1525],
  '1700 E 19TH ST THE DALLES':               [45.5980, -121.1525],
  // Lake District Hospital (Lakeview — mutual aid)
  'LAKE DISTRICT HOSPITAL':                  [42.1818, -120.3515],
  '700 S J ST LAKEVIEW':                     [42.1818, -120.3515],
  // Airports / Landing Zones
  'ROBERTS FIELD':                           [44.2542, -121.1486],
  'KRDM AIRPORT':                            [44.2542, -121.1486],
  'REDMOND MUNICIPAL AIRPORT':               [44.2542, -121.1486],
  'LZ ROBERTS FIELD':                        [44.2542, -121.1486],
  'BEND MUNICIPAL AIRPORT':                  [44.0946, -121.2002],
  'KBDN AIRPORT':                            [44.0946, -121.2002],
  'LZ BEND AIRPORT':                         [44.0946, -121.2002],
  'SUNRIVER AIRPORT':                        [43.8763, -121.4530],
  'LZ SUNRIVER AIRPORT':                     [43.8763, -121.4530],
  'PRINEVILLE AIRPORT':                      [44.2878, -120.9055],
  'LZ PRINEVILLE AIRPORT':                   [44.2878, -120.9055],
  'MADRAS MUNICIPAL AIRPORT':                [44.6702, -121.1552],
  'LZ MADRAS AIRPORT':                       [44.6702, -121.1552],
  'KLAMATH FALLS AIRPORT':                   [42.1561, -121.7333],
  'LZ KLAMATH FALLS AIRPORT':               [42.1561, -121.7333],
  'THE DALLES AIRPORT':                      [45.6194, -121.1683],
  'LZ THE DALLES AIRPORT':                   [45.6194, -121.1683],
  // Fairgrounds / LZs
  'DESCHUTES FAIRGROUNDS':                   [44.2665, -121.1775],
  'LZ DESCHUTES FAIRGROUNDS':               [44.2665, -121.1775],
  '3800 SW AIRPORT WAY REDMOND':             [44.2665, -121.1775],
  'JEFFERSON COUNTY FAIRGROUNDS':            [44.6196, -121.1380],
  'LZ JEFFERSON FAIRGROUNDS':               [44.6196, -121.1380],
  'CROOK COUNTY FAIRGROUNDS':                [44.2893, -120.8422],
  'LZ CROOK COUNTY FAIRGROUNDS':            [44.2893, -120.8422],
  'KLAMATH COUNTY FAIRGROUNDS':              [42.2084, -121.7443],
  // Key intersections / highway landmarks
  'US97 AT LA PINE':                         [43.6746, -121.5003],
  'US97 AT SUNRIVER JUNCTION':               [43.8784, -121.4344],
  'US97 AT BEND 3RD ST':                     [44.0582, -121.3063],
  'US97 AT COOLEY RD':                       [44.1006, -121.2775],
  'US97 AT REDMOND':                         [44.2726, -121.1739],
  'US97 AT TERREBONNE':                      [44.3529, -121.1778],
  'US97 AT MADRAS':                          [44.6335, -121.1295],
  'US97 AT CHEMULT':                         [43.2165, -121.7828],
  'US20 AT SISTERS':                         [44.2893, -121.5490],
  'US20 AT BEND':                            [44.0977, -121.2815],
  'US20 AT BROTHERS':                        [43.8140, -120.6030],
  // Schools (MCI staging)
  'MOUNTAIN VIEW HIGH SCHOOL':               [44.0773, -121.2647],
  'MVHS BEND':                               [44.0773, -121.2647],
  'SUMMIT HIGH SCHOOL':                      [44.0576, -121.3613],
  'LA PINE HIGH SCHOOL':                     [43.6702, -121.4868],
  'SISTERS HIGH SCHOOL':                     [44.2889, -121.5618],
  'CROOK COUNTY HIGH SCHOOL':                [44.2928, -120.8331],
  'MADRAS HIGH SCHOOL':                      [44.6274, -121.1198],
  // Courthouses
  'DESCHUTES COUNTY COURTHOUSE':             [44.0587, -121.3124],
  'CROOK COUNTY COURTHOUSE':                 [44.3012, -120.8348],
  'JEFFERSON COUNTY COURTHOUSE':             [44.6328, -121.1282],
  'WASCO COUNTY COURTHOUSE':                 [45.5998, -121.1841],
};

// Hospital destination markers — transport destinations for Central Oregon EMS
const BM_HOSPITALS = [
  { code: 'SCMC-BEND',       name: 'SCMC Bend',                     lat: 44.0672, lon: -121.2690 },
  { code: 'SCMC-REDMOND',    name: 'SCMC Redmond',                  lat: 44.2704, lon: -121.1417 },
  { code: 'SCMC-PRINEVILLE', name: 'SCMC Prineville',               lat: 44.2980, lon: -120.8253 },
  { code: 'SCMC-MADRAS',     name: 'SCMC Madras',                   lat: 44.6329, lon: -121.1298 },
  { code: 'SALEM-HOSP',      name: 'Salem Hospital',                lat: 44.9363, lon: -123.0351 },
  { code: 'OHSU',            name: 'OHSU Portland',                 lat: 45.4991, lon: -122.6870 },
  { code: 'SKY-LAKES',       name: 'Sky Lakes MC (Klamath Falls)',  lat: 42.2530, lon: -121.7851 },
  { code: 'MCMC',            name: 'Mid-Columbia MC (The Dalles)',  lat: 45.5980, lon: -121.1525 },
  { code: 'WS-IHS',          name: 'Warm Springs IHS',              lat: 44.7636, lon: -121.2733 },
];

let _bmLoaded         = false;
let _bmLoading        = false;
let _bmMap            = null;
let _bmMarkers        = [];
let _bmGeoCache       = Object.assign({}, BM_KNOWN_COORDS);
let _bmGeoQueue       = [];
let _bmGeoTimer       = null;
let _popoutMapWindow  = null;

function toggleBoardMap() {
  const open = document.body.classList.toggle('board-map-open');
  const btn  = document.getElementById('tbBtnMAP');
  if (btn) btn.classList.toggle('active', open);
  if (open) {
    _loadBoardLeaflet(() => {
      setTimeout(() => {
        if (!_bmMap) _initBoardMap();
        if (_bmMap) _bmMap.invalidateSize();
        renderBoardMap();
        _startBmGeoQueue();
      }, 300);
    });
  } else {
    _stopBmGeoQueue();
  }
}

function _loadBoardLeaflet(cb) {
  if (_bmLoaded)  { cb(); return; }
  if (_bmLoading) { setTimeout(() => _loadBoardLeaflet(cb), 120); return; }
  _bmLoading = true;
  const link = document.createElement('link');
  link.rel = 'stylesheet';
  link.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
  document.head.appendChild(link);
  const script = document.createElement('script');
  script.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
  script.onload  = () => { _bmLoaded = true; _bmLoading = false; cb(); };
  script.onerror = () => { _bmLoading = false; };
  document.body.appendChild(script);
}

function _initBoardMap() {
  const el = document.getElementById('boardMapContainer');
  if (!el || !window.L) return;
  _bmMap = L.map('boardMapContainer', { zoomControl: true, attributionControl: false })
    .setView(BM_MAP_CENTER, BM_MAP_ZOOM);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 18, opacity: 0.7
  }).addTo(_bmMap);
  const dimData = {
    type: 'Feature',
    geometry: {
      type: 'Polygon',
      coordinates: [
        [[-180,-90],[180,-90],[180,90],[-180,90],[-180,-90]],
        [[BM_TRICOUNTY[0][1], BM_TRICOUNTY[0][0]], [BM_TRICOUNTY[0][1], BM_TRICOUNTY[1][0]],
         [BM_TRICOUNTY[1][1], BM_TRICOUNTY[1][0]], [BM_TRICOUNTY[1][1], BM_TRICOUNTY[0][0]],
         [BM_TRICOUNTY[0][1], BM_TRICOUNTY[0][0]]]
      ]
    }
  };
  L.geoJSON(dimData, { style: { fillColor: '#000', fillOpacity: 0.35, stroke: false, color: 'transparent', weight: 0 } }).addTo(_bmMap);
  L.rectangle(BM_TRICOUNTY, { color: 'rgba(91,163,230,0.4)', weight: 2, fill: false, dashArray: '4 4' }).addTo(_bmMap);
  _bmMap.fitBounds(BM_TRICOUNTY, { padding: [20, 20] });
}

function renderBoardMap() {
  if (!document.body.classList.contains('board-map-open')) return;
  if (!_bmMap || !window.L || !STATE) return;
  // Clear all markers EXCEPT the persistent search pin
  _bmMarkers.forEach(m => { if (m !== _bmSearchPin) { try { _bmMap.removeLayer(m); } catch(e) {} } });
  _bmMarkers = _bmSearchPin ? [_bmSearchPin] : [];

  // Only show units that are active AND updated within last 12 hours (filter stale DB rows)
  const _mapStaleMs = 12 * 60 * 60 * 1000;
  const _mapNow = Date.now();
  const units     = (STATE.units || []).filter(u => u.active && u.updated_at && (_mapNow - new Date(u.updated_at).getTime()) < _mapStaleMs);
  const incidents = (STATE.incidents || []).filter(i => i.status === 'ACTIVE' || i.status === 'QUEUED');

  // Incident pins
  incidents.forEach(inc => {
    const addr = (inc.scene_address || '').trim();
    if (!addr) return;
    const geo = _bmGeoCache[addr];
    if (geo) {
      const pri = inc.priority || '';
      const color = pri === 'PRI-1' ? '#f44336' : pri === 'PRI-2' ? '#ff9800' : '#4fa3e0';
      const shortId = String(inc.incident_id).replace(/^\d{2}-/, '');
      const tip = '<b>' + esc(inc.incident_id) + '</b>' + (inc.incident_type ? '<br>' + inc.incident_type : '') + (addr ? '<br>' + addr : '');
      const m = L.circleMarker(geo, { radius: 9, color, weight: 2, fillColor: color, fillOpacity: 0.35 })
        .bindTooltip(tip, { permanent: false, direction: 'top' });
      m.addTo(_bmMap);
      _bmMarkers.push(m);
    } else if (!_bmGeoCache.hasOwnProperty(addr) && !_bmGeoQueue.some(q => q.addr === addr)) {
      _bmGeoQueue.push({ addr, near: null, bounded: 0 });
    }
  });
  // Ensure geocoding queue is processing if there are pending addresses
  if (_bmGeoQueue.length) _startBmGeoQueue();

  // Unit dots
  const posUsage = {};
  units.forEach(u => {
    let pos = null;
    const stCode = String(u.status || '').toUpperCase();
    // Priority 1: incident scene address
    if (u.incident && ['D','DE','OS','T','AT'].includes(stCode)) {
      const inc = (STATE.incidents || []).find(i => i.incident_id === u.incident);
      if (inc && inc.scene_address) {
        const geo = _bmGeoCache[(inc.scene_address || '').trim()];
        if (geo) pos = [geo[0], geo[1]];
      }
    }
    // Priority 2: [LOC:address] tag in note
    if (!pos) {
      const locAddr = _parseLocTag(u.note);
      if (locAddr) {
        // Direct coordinate tag (from GPSUL) — use immediately, no geocoding needed
        const locCoordMatch = locAddr.match(/^(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)$/);
        if (locCoordMatch) {
          pos = [parseFloat(locCoordMatch[1]), parseFloat(locCoordMatch[2])];
        } else {
          const geo = _bmGeoCache[locAddr];
          if (geo) pos = [geo[0], geo[1]];
          else if (!_bmGeoCache.hasOwnProperty(locAddr) && !_bmGeoQueue.some(q => q.addr === locAddr)) {
            const uid2 = String(u.unit_id || '').toUpperCase();
            _bmGeoQueue.push({ addr: locAddr, near: BM_UNIT_HOME[uid2] || null, bounded: 1 });
          }
        }
      }
    }
    // Priority 3: unit home base coords — if still no position, skip (don't cluster at map center)
    if (!pos) {
      const uid = String(u.unit_id || '').toUpperCase();
      pos = BM_UNIT_HOME[uid] ? [...BM_UNIT_HOME[uid]] : null;
    }
    if (!pos) return;
    const posKey = pos[0].toFixed(4) + ',' + pos[1].toFixed(4);
    const n = posUsage[posKey] = (posUsage[posKey] || 0) + 1;
    if (n > 1) {
      const angle = ((n - 1) * 137.5) * (Math.PI / 180);
      pos = [pos[0] + Math.sin(angle) * 0.0012, pos[1] + Math.cos(angle) * 0.0012];
    }
    const dotColor = BM_STATUS_COLORS[stCode] || '#607d8b';
    const uId = String(u.unit_id || '').toUpperCase();
    // Contextual location in tooltip
    let tipLoc = '';
    if (['D','DE','OS'].includes(stCode) && u.incident) {
      const tInc = (STATE.incidents || []).find(i => i.incident_id === u.incident);
      if (tInc && tInc.scene_address) tipLoc = tInc.scene_address;
    }
    if (!tipLoc && u.destination) tipLoc = u.destination;
    if (!tipLoc) { const lt = _parseLocTag(u.note); if (lt) tipLoc = lt; }
    const tip = '<b>' + uId + '</b> — ' + stCode + (tipLoc ? '<br>' + esc(tipLoc) : '');
    const uIcon = getUnitMapIcon(u.type, dotColor);
    const uMark = L.marker(pos, { icon: uIcon, zIndexOffset: 200 })
      .bindTooltip(tip, { permanent: false, direction: 'top' });
    uMark.addTo(_bmMap);
    _bmMarkers.push(uMark);
    const label = L.marker(pos, {
      icon: L.divIcon({ html: '<div class="v-map-label">' + uId + '</div>', className: '', iconSize: [0, 0] }),
      interactive: false
    }).addTo(_bmMap);
    _bmMarkers.push(label);
  });

  // Hospital helipad markers — "H" circles at all regional hospital landing zones
  LFN_HELIPADS.filter(function(h) { return h.type !== 'AIRPORT'; }).forEach(function(h) {
    const hIcon = L.divIcon({ html: '<div class="lfn-helipad">H</div>', className: '', iconSize: [18,18], iconAnchor: [9,9] });
    const hm = L.marker([h.lat, h.lon], { icon: hIcon, zIndexOffset: 500 })
      .addTo(_bmMap).bindTooltip(h.name, { direction: 'top' });
    _bmMarkers.push(hm);
  });

  // LFN aircraft markers — heading-rotated triangles for known airborne aircraft
  _lfnAircraft.filter(function(ac) { return ac.status !== 'NOSIG' && ac.lat != null; }).forEach(function(ac) {
    const icon = getLfnMapIcon(ac.heading || 0, ac.status);
    const tipParts = [ac.callsign + ' \u2014 ' + ac.status];
    if (ac.alt_ft  != null) tipParts.push(ac.alt_ft.toLocaleString() + 'ft');
    if (ac.gs_kts  != null) tipParts.push(ac.gs_kts + 'kts');
    const tip = tipParts.join(' | ');
    const am = L.marker([ac.lat, ac.lon], { icon: icon, zIndexOffset: 1000 })
      .addTo(_bmMap).bindTooltip(tip, { direction: 'top' });
    _bmMarkers.push(am);
    const albl = L.marker([ac.lat, ac.lon], {
      icon: L.divIcon({ html: '<div class="v-map-label" style="color:#aac">' + esc(ac.callsign) + '</div>', className: '', iconSize: [0,0] }),
      interactive: false
    }).addTo(_bmMap);
    _bmMarkers.push(albl);
  });
}

function _parseLocTag(note) {
  if (!note) return null;
  const m = String(note).match(/\[LOC:([^\]]+)\]/);
  return m ? m[1].trim() : null;
}

function _getUnitMapPos(unitId) {
  if (!STATE) return null;
  const u = (STATE.units || []).find(x => String(x.unit_id).toUpperCase() === unitId.toUpperCase());
  if (!u) return null;
  const stCode = String(u.status || '').toUpperCase();
  // Priority 1: incident scene address (if actively dispatched)
  if (u.incident && ['D','DE','OS','T','AT'].includes(stCode)) {
    const inc = (STATE.incidents || []).find(i => i.incident_id === u.incident);
    if (inc && inc.scene_address) {
      const geo = _bmGeoCache[(inc.scene_address || '').trim()];
      if (geo) return { pos: [geo[0], geo[1]], unit: u, source: 'incident' };
    }
  }
  // Priority 2: [LOC:address] tag in note
  const locAddr = _parseLocTag(u.note);
  if (locAddr) {
    // If LOC tag is raw coordinates (from GPSUL), resolve instantly without Nominatim
    const coordMatch = locAddr.match(/^(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)$/);
    if (coordMatch) return { pos: [parseFloat(coordMatch[1]), parseFloat(coordMatch[2])], unit: u, source: 'loc' };
    const geo = _bmGeoCache[locAddr];
    if (geo) return { pos: [geo[0], geo[1]], unit: u, source: 'loc' };
    // Queue for geocoding if not yet known (bounded=1 forces Central OR; use home coords as tiebreaker)
    if (!_bmGeoCache.hasOwnProperty(locAddr) && !_bmGeoQueue.some(q => q.addr === locAddr)) {
      const uid = String(u.unit_id || '').toUpperCase();
      _bmGeoQueue.push({ addr: locAddr, near: BM_UNIT_HOME[uid] || null, bounded: 1 });
      if (_bmGeoQueue.length) _startBmGeoQueue();
    }
  }
  // Priority 2.5: PP note address — note format is "[PP] 1234 NW MAIN ST · CALL TYPE"
  // Only apply when unit is actively dispatched (not AV/OOS)
  if ((u.source || '').startsWith('PP:') && stCode !== 'AV' && stCode !== 'OOS') {
    const ppNoteMatch = (u.note || '').match(/^\[PP\]\s+([^·\n]+)/i);
    if (ppNoteMatch) {
      const ppAddr = ppNoteMatch[1].trim();
      if (ppAddr && ppAddr.length > 5) {
        const geo = _bmGeoCache[ppAddr.toUpperCase()];
        if (geo) return { pos: [geo[0], geo[1]], unit: u, source: 'pp' };
        if (!_bmGeoCache.hasOwnProperty(ppAddr.toUpperCase()) && !_bmGeoQueue.some(q => q.addr === ppAddr.toUpperCase())) {
          _bmGeoQueue.push({ addr: ppAddr.toUpperCase(), near: null, bounded: 1 });
          if (_bmGeoQueue.length) _startBmGeoQueue();
        }
      }
    }
  }
  // Priority 3: unit home base coordinates
  const uid = String(u.unit_id || '').toUpperCase();
  if (BM_UNIT_HOME[uid]) return { pos: [...BM_UNIT_HOME[uid]], unit: u, source: 'station' };
  return { pos: [...BM_MAP_CENTER], unit: u, source: 'default' };
}

function _getIncidentMapPos(incArg) {
  if (!STATE) return null;
  const id = incArg.replace(/^INC/i, '');
  const inc = (STATE.incidents || []).find(i => {
    const iid = String(i.incident_id || '');
    return iid === incArg || iid.endsWith('-' + id) || iid === id;
  });
  if (!inc) return null;
  const addr = (inc.scene_address || '').trim();
  if (!addr) return { inc, pos: null, source: 'no-address' };
  const geo = _bmGeoCache[addr];
  if (geo) return { inc, pos: [geo[0], geo[1]], source: 'scene' };
  // Queue for geocoding (bounded=0 — incident scenes may be outside Central OR for interfacility transfers)
  if (!_bmGeoCache.hasOwnProperty(addr) && !_bmGeoQueue.some(q => q.addr === addr)) {
    _bmGeoQueue.push({ addr, near: null, bounded: 0 });
    if (_bmGeoQueue.length) _startBmGeoQueue();
  }
  return { inc, pos: null, source: 'geocoding' };
}

function focusUnitOnMap(unitId) {
  if (!STATE) { showToast('NO DATA — WAIT FOR FIRST POLL.'); return; }
  const info = _getUnitMapPos(unitId);
  if (!info) { showToast('UNIT NOT FOUND: ' + unitId); return; }

  // If a better position is pending geocoding, wait and retry rather than focusing on station
  const u = info.unit;
  const stCode = String(u.status || '').toUpperCase();
  if (info.source === 'station' || info.source === 'default') {
    let pendingAddr = null;
    if (u.incident && ['D','DE','OS','T','AT'].includes(stCode)) {
      const inc = (STATE.incidents || []).find(i => i.incident_id === u.incident);
      if (inc && inc.scene_address) {
        const a = (inc.scene_address || '').trim();
        if (!_bmGeoCache.hasOwnProperty(a)) pendingAddr = a;
      }
    }
    if (!pendingAddr) {
      const locAddr = _parseLocTag(u.note);
      if (locAddr && !_bmGeoCache.hasOwnProperty(locAddr)) pendingAddr = locAddr;
    }
    if (pendingAddr) {
      showToast('GEOCODING ' + unitId.toUpperCase() + ' — FOCUSING IN 3s...');
      _ensureMapOpen(() => { renderBoardMap(); });
      setTimeout(() => focusUnitOnMap(unitId), 3000);
      return;
    }
    // No known position and nothing pending — don't zoom to map center (useless)
    if (info.source === 'default') {
      showToast('NO MAP POSITION FOR ' + unitId.toUpperCase() + ' — USE: LOC ' + unitId.toUpperCase() + ' <ADDR>');
      return;
    }
  }

  function doFocus() {
    if (!_bmMap || !window.L) { showToast('MAP NOT READY.'); return; }
    renderBoardMap();
    _bmMap.setView(info.pos, 14, { animate: true });

    // Highlight pulse ring
    const ring = L.circleMarker(info.pos, {
      radius: 20, color: '#ffffff', weight: 3, fillColor: '#4fa3e0',
      fillOpacity: 0.25, dashArray: '6 4', className: 'map-focus-ring'
    }).addTo(_bmMap);
    _bmMarkers.push(ring);

    // Open tooltip on the unit's marker
    for (const m of _bmMarkers) {
      if (m.getTooltip && m.getTooltip()) {
        const tip = m.getTooltip().getContent();
        if (tip && tip.includes('<b>' + unitId.toUpperCase() + '</b>')) {
          m.openTooltip();
          break;
        }
      }
    }

    // Animate pulse ring away after 3s
    setTimeout(() => {
      try { _bmMap.removeLayer(ring); } catch(e) {}
      const idx = _bmMarkers.indexOf(ring);
      if (idx !== -1) _bmMarkers.splice(idx, 1);
    }, 3000);

    const loc = info.source === 'incident' ? 'SCENE' : info.source === 'loc' ? 'LOC' : info.source === 'station' ? info.unit.station : 'DEFAULT';
    showToast('FOCUSED: ' + unitId.toUpperCase() + ' (' + loc + ')');
  }

  _ensureMapOpen(doFocus);
}

function focusIncidentOnMap(incArg) {
  if (!STATE) { showToast('NO DATA — WAIT FOR FIRST POLL.'); return; }
  const info = _getIncidentMapPos(incArg);
  if (!info) { showToast('INCIDENT NOT FOUND: ' + incArg); return; }
  if (!info.pos && info.source === 'no-address') { showToast('NO SCENE ADDRESS FOR ' + info.inc.incident_id); return; }
  if (!info.pos && info.source === 'geocoding') { showToast('GEOCODING SCENE ADDRESS — TRY AGAIN IN A FEW SECONDS.'); _ensureMapOpen(() => { renderBoardMap(); }); return; }

  function doFocus() {
    if (!_bmMap || !window.L) { showToast('MAP NOT READY.'); return; }
    renderBoardMap();
    _bmMap.setView(info.pos, 15, { animate: true });

    const ring = L.circleMarker(info.pos, {
      radius: 22, color: '#ff5555', weight: 3, fillColor: '#ff5555',
      fillOpacity: 0.2, dashArray: '6 4', className: 'map-focus-ring'
    }).addTo(_bmMap);
    _bmMarkers.push(ring);

    // Open tooltip on the incident marker
    for (const m of _bmMarkers) {
      if (m.getTooltip && m.getTooltip()) {
        const tip = m.getTooltip().getContent();
        if (tip && tip.includes('INC') && tip.includes(String(info.inc.incident_id).replace(/^\d{2}-/, ''))) {
          m.openTooltip();
          break;
        }
      }
    }

    setTimeout(() => {
      try { _bmMap.removeLayer(ring); } catch(e) {}
      const idx = _bmMarkers.indexOf(ring);
      if (idx !== -1) _bmMarkers.splice(idx, 1);
    }, 3000);

    showToast('FOCUSED: ' + info.inc.incident_id + (info.inc.scene_address ? ' @ ' + info.inc.scene_address : ''));
  }

  _ensureMapOpen(doFocus);
}

function _ensureMapOpen(cb) {
  if (!document.body.classList.contains('board-map-open')) {
    document.body.classList.add('board-map-open');
    const btn = document.getElementById('tbBtnMAP');
    if (btn) btn.classList.add('active');
    _loadBoardLeaflet(() => {
      setTimeout(() => {
        if (!_bmMap) _initBoardMap();
        if (_bmMap) _bmMap.invalidateSize();
        cb();
        _startBmGeoQueue();
      }, 300);
    });
  } else {
    cb();
  }
}

// ── Map zoom/focus commands ─────────────────────────────────

function mapZoomIn() {
  _ensureMapOpen(() => {
    if (!_bmMap) return;
    _bmMap.zoomIn(1, { animate: true });
    showToast('ZOOM: ' + _bmMap.getZoom());
  });
}

function mapZoomOut() {
  _ensureMapOpen(() => {
    if (!_bmMap) return;
    _bmMap.zoomOut(1, { animate: true });
    showToast('ZOOM: ' + _bmMap.getZoom());
  });
}

function mapFitAll() {
  if (!STATE) { showToast('NO DATA — WAIT FOR FIRST POLL.'); return; }
  _ensureMapOpen(() => {
    if (!_bmMap || !window.L) return;
    renderBoardMap();
    const pts = [];
    (STATE.units || []).filter(u => u.active).forEach(u => {
      const info = _getUnitMapPos(u.unit_id);
      if (info && info.pos) pts.push(info.pos);
    });
    (STATE.incidents || []).filter(i => i.status === 'ACTIVE' || i.status === 'QUEUED').forEach(inc => {
      const addr = (inc.scene_address || '').trim();
      if (addr && _bmGeoCache[addr]) pts.push(_bmGeoCache[addr]);
    });
    if (pts.length > 1) {
      _bmMap.fitBounds(pts, { padding: [30, 30], animate: true });
      showToast('FIT: ' + pts.length + ' MARKERS');
    } else if (pts.length === 1) {
      _bmMap.setView(pts[0], 13, { animate: true });
      showToast('FIT: 1 MARKER');
    } else {
      _bmMap.fitBounds(BM_TRICOUNTY, { padding: [20, 20], animate: true });
      showToast('NO MARKERS — SHOWING REGION');
    }
  });
}

function mapShowStations() {
  _ensureMapOpen(() => {
    if (!_bmMap || !window.L) return;
    renderBoardMap();
    const pts = [];
    Object.entries(BM_STATION_BASES).forEach(([name, coords]) => {
      pts.push(coords);
      const m = L.circleMarker(coords, {
        radius: 12, color: '#ffffff', weight: 2, fillColor: '#2196f3', fillOpacity: 0.4,
        dashArray: '4 3', className: 'map-focus-ring'
      }).bindTooltip('<b>' + name + '</b>', { permanent: true, direction: 'top', className: 'v-map-label' })
        .addTo(_bmMap);
      _bmMarkers.push(m);
    });
    if (pts.length) {
      _bmMap.fitBounds(pts, { padding: [40, 40], animate: true });
    }
    showToast('STATIONS: ' + pts.length + ' SHOWN');
  });
}

function mapShowIncidents() {
  if (!STATE) { showToast('NO DATA — WAIT FOR FIRST POLL.'); return; }
  const active = (STATE.incidents || []).filter(i => i.status === 'ACTIVE' || i.status === 'QUEUED');
  if (!active.length) { showToast('NO ACTIVE INCIDENTS.'); return; }
  _ensureMapOpen(() => {
    if (!_bmMap || !window.L) return;
    renderBoardMap();
    const pts = [];
    active.forEach(inc => {
      const addr = (inc.scene_address || '').trim();
      if (addr && _bmGeoCache[addr]) {
        const geo = _bmGeoCache[addr];
        pts.push(geo);
        const ring = L.circleMarker(geo, {
          radius: 18, color: '#ff5555', weight: 3, fillColor: '#ff5555',
          fillOpacity: 0.2, dashArray: '6 4', className: 'map-focus-ring'
        }).addTo(_bmMap);
        _bmMarkers.push(ring);
        setTimeout(() => {
          try { _bmMap.removeLayer(ring); } catch(e) {}
          const idx = _bmMarkers.indexOf(ring);
          if (idx !== -1) _bmMarkers.splice(idx, 1);
        }, 5000);
      }
    });
    if (pts.length > 1) {
      _bmMap.fitBounds(pts, { padding: [30, 30], animate: true });
    } else if (pts.length === 1) {
      _bmMap.setView(pts[0], 14, { animate: true });
    }
    showToast('INCIDENTS: ' + pts.length + '/' + active.length + ' GEOCODED');
  });
}

function mapReset() {
  _clearSearchPin();
  _ensureMapOpen(() => {
    if (!_bmMap) return;
    _bmMap.fitBounds(BM_TRICOUNTY, { padding: [20, 20], animate: true });
    renderBoardMap();
    showToast('MAP RESET TO DEFAULT VIEW');
  });
}

// MAP <address> — geocode an arbitrary address and focus map on it
// Search pin persists until next MAP <addr>, MAP CLR, or MAP RESET
let _bmSearchPin = null;
function _clearSearchPin() {
  if (_bmSearchPin && _bmMap) {
    try { _bmMap.removeLayer(_bmSearchPin); } catch(e) {}
    const idx = _bmMarkers.indexOf(_bmSearchPin);
    if (idx !== -1) _bmMarkers.splice(idx, 1);
    _bmSearchPin = null;
  }
}

// Expand street direction abbreviations before sending to Nominatim.
// NE/NW/SE/SW as isolated words → NORTHEAST/NORTHWEST/SOUTHEAST/SOUTHWEST.
// Single N/S/E/W only expanded when followed by a space+digit or space+word
// (avoids mangling city names like "BEND E" edge cases).
function _expandDirections(addr) {
  if (!addr) return addr;
  return addr
    .replace(/\bNE\b/g, 'NORTHEAST')
    .replace(/\bNW\b/g, 'NORTHWEST')
    .replace(/\bSE\b/g, 'SOUTHEAST')
    .replace(/\bSW\b/g, 'SOUTHWEST');
}

function mapGeoFocus(addr) {
  if (!addr) { showToast('MAP: ENTER AN ADDRESS.'); return; }

  // 1. Detect lat/lon coordinates — skip geocoding entirely
  const coordMatch = addr.match(/^(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)$/);
  if (coordMatch) {
    const lat = parseFloat(coordMatch[1]), lon = parseFloat(coordMatch[2]);
    if (lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180) {
      _ensureMapOpen(() => {
        if (!_bmMap || !window.L) return;
        _clearSearchPin();
        _bmMap.setView([lat, lon], 15, { animate: true });
        _bmSearchPin = L.circleMarker([lat, lon], {
          radius: 12, color: '#FFD700', weight: 3, fillColor: '#FFD700',
          fillOpacity: 0.3, dashArray: '6 4'
        }).addTo(_bmMap);
        _bmSearchPin.bindTooltip('<b>' + lat.toFixed(5) + ', ' + lon.toFixed(5) + '</b>', { permanent: true, direction: 'top' }).openTooltip();
        _bmMarkers.push(_bmSearchPin);
        showToast('MAP: ' + lat.toFixed(5) + ', ' + lon.toFixed(5));
      });
      return;
    }
  }

  // 2. Check AddressLookup shortcodes
  const resolved = AddressLookup.resolve(addr);
  let searchAddr = resolved !== addr ? resolved : addr;

  // 2.5 — Check known geocache before hitting Nominatim (covers SCMC, known destinations, etc.)
  if (_bmGeoCache.hasOwnProperty(searchAddr)) {
    const geo = _bmGeoCache[searchAddr];
    if (geo) {
      _ensureMapOpen(() => {
        if (!_bmMap || !window.L) return;
        _clearSearchPin();
        _bmMap.setView(geo, 15, { animate: true });
        _bmSearchPin = L.circleMarker(geo, {
          radius: 12, color: '#FFD700', weight: 3, fillColor: '#FFD700',
          fillOpacity: 0.3, dashArray: '6 4'
        }).addTo(_bmMap);
        _bmSearchPin.bindTooltip('<b>' + esc(searchAddr) + '</b>', { permanent: true, direction: 'top' }).openTooltip();
        _bmMarkers.push(_bmSearchPin);
        showToast('MAP: ' + searchAddr);
      });
      return;
    }
  }

  // 3. Handle intersections — normalize / separator to & (Nominatim format)
  if (searchAddr.includes('/')) {
    searchAddr = searchAddr.replace(/\s*\/\s*/g, ' & ');
  }

  // 4. Build query — expand direction abbreviations, then bias to tri-county area
  const expandedAddr = _expandDirections(searchAddr);
  const hasCity = /,|\b(BEND|REDMOND|SISTERS|PRINEVILLE|MADRAS|LA PINE|SUNRIVER|TERREBONNE|POWELL BUTTE|TUMALO|CROOKED RIVER|WARM SPRINGS)\b/i.test(expandedAddr);
  let query = hasCity ? expandedAddr : expandedAddr + ', OR';
  const vbox = '&viewbox=' + BM_MAP_VIEWBOX + '&bounded=0';

  showToast('GEOCODING: ' + query);

  _ensureMapOpen(async () => {
    if (!_bmMap || !window.L) return;
    try {
      const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=us' + vbox + '&q=' + encodeURIComponent(query);
      const resp = await fetch(url);
      const results = await resp.json();
      if (!results || !results.length) {
        showToast('NO RESULTS FOR: ' + query);
        return;
      }
      const r = results[0];
      const lat = parseFloat(r.lat);
      const lon = parseFloat(r.lon);

      // Clear previous search pin
      _clearSearchPin();

      _bmMap.setView([lat, lon], 15, { animate: true });
      // Yellow search pin — persists until next search or MAP RESET/CLR
      _bmSearchPin = L.circleMarker([lat, lon], {
        radius: 12, color: '#FFD700', weight: 3, fillColor: '#FFD700',
        fillOpacity: 0.3, dashArray: '6 4'
      }).addTo(_bmMap);
      const displayName = (r.display_name || addr).substring(0, 100);
      _bmSearchPin.bindTooltip('<b>' + displayName + '</b>', { permanent: true, direction: 'top' }).openTooltip();
      _bmMarkers.push(_bmSearchPin);

      const msg = isIntersection
        ? 'INTERSECTION — SHOWING ' + searchAddr.toUpperCase() + '. VERIFY PIN LOCATION.'
        : displayName;
      showToast(msg);
    } catch(e) {
      showToast('GEOCODING FAILED — NETWORK ERROR.');
    }
  });
}

function _startBmGeoQueue() {
  if (_bmGeoTimer) return;
  _bmGeoTimer = setInterval(_processBmGeoQueue, 1500);
}

function _stopBmGeoQueue() {
  if (_bmGeoTimer) { clearInterval(_bmGeoTimer); _bmGeoTimer = null; }
}

// Approximate squared distance between two lat/lng points (no sqrt needed for comparison)
function _geoDistSq(lat1, lng1, lat2, lng2) {
  const dlat = lat2 - lat1;
  const dlng = (lng2 - lng1) * Math.cos(lat1 * Math.PI / 180);
  return dlat * dlat + dlng * dlng;
}

async function _processBmGeoQueue() {
  if (!_bmGeoQueue.length) return;
  const item = _bmGeoQueue.shift();
  const addr = item.addr;
  const near = item.near;       // [lat, lng] or null
  const bounded = item.bounded; // 1 = strict Central OR viewbox (unit LOCs), 0 = loose (incident scenes)
  if (_bmGeoCache.hasOwnProperty(addr)) return;
  // If addr is raw coordinates (from GPSUL [LOC:lat,lon]), resolve instantly
  const _coordCheck = addr.match(/^(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)$/);
  if (_coordCheck) { _bmGeoCache[addr] = [parseFloat(_coordCheck[1]), parseFloat(_coordCheck[2])]; renderBoardMap(); return; }
  try {
    // Strip leading non-numeric facility name prefix before geocoding
    // e.g. "SAINT CHARLES BEND 2500 NE NEFF RD" → "2500 NE NEFF RD"
    const hasStreetNum = /\d/.test(addr);
    let stripped = hasStreetNum ? addr.replace(/^[A-Za-z\s,.'"-]+?(?=\d)/, '').trim() : addr;
    // Normalize intersection separators: "HWY 97 / SW 61ST" → "HWY 97 & SW 61ST" (Nominatim format)
    stripped = stripped.replace(/\s*\/\s*/g, ' & ');
    let geocodeAddr = _expandDirections(stripped);
    // Extract city name if present — restructure as "<street>, <City>, OR" for better Nominatim accuracy
    const KNOWN_CITIES = ['BEND','REDMOND','PRINEVILLE','MADRAS','SISTERS','LA PINE','LAPINE','WARM SPRINGS','CULVER','METOLIUS','POWELL BUTTE','MITCHELL','DAYVILLE','JOHN DAY','BURNS','LAKEVIEW','KLAMATH FALLS'];
    const cityMatch = KNOWN_CITIES.find(c => {
      const re = new RegExp('\\b' + c.replace(' ', '\\s+') + '\\b', 'i');
      return re.test(geocodeAddr);
    });
    if (cityMatch) {
      geocodeAddr = geocodeAddr.replace(new RegExp('\\b' + cityMatch.replace(' ', '\\s+') + '\\b', 'i'), '').replace(/[,\s]+$/, '').trim() + ', ' + cityMatch.charAt(0) + cityMatch.slice(1).toLowerCase() + ', OR';
    }
    const query = /oregon|,\s*or\b/i.test(geocodeAddr) ? geocodeAddr : geocodeAddr + ', Oregon';
    // Fetch multiple results when we have a home coord to pick closest from
    const limit = near ? 5 : 1;
    const url = 'https://nominatim.openstreetmap.org/search?format=json&q=' +
      encodeURIComponent(query) + '&limit=' + limit + '&viewbox=' + BM_MAP_VIEWBOX +
      '&bounded=' + (bounded || 0) + '&countrycodes=us';
    const res = await fetch(url);
    if (!res.ok) return;
    const data = await res.json();
    if (data && data.length) {
      let best = data[0];
      if (near && data.length > 1) {
        let bestDist = Infinity;
        for (const r of data) {
          const d = _geoDistSq(near[0], near[1], parseFloat(r.lat), parseFloat(r.lon));
          if (d < bestDist) { bestDist = d; best = r; }
        }
      }
      _bmGeoCache[addr] = [parseFloat(best.lat), parseFloat(best.lon)];
      renderBoardMap();
    } else {
      _bmGeoCache[addr] = null;
    }
  } catch (e) {}
}

function openPopoutMap() {
  if (_popoutMapWindow && !_popoutMapWindow.closed) {
    _popoutMapWindow.focus();
    showToast('MAP ALREADY OPEN IN POPOUT.');
    return;
  }
  // Collapse inline map if open
  if (document.body.classList.contains('board-map-open')) {
    document.body.classList.remove('board-map-open');
    const btn = document.getElementById('tbBtnMAP');
    if (btn) btn.classList.remove('active');
    _stopBmGeoQueue();
  }
  _popoutMapWindow = window.open('/popout-map/', 'hoscad-map', 'width=1000,height=750,left=0,top=0');
  if (!_popoutMapWindow) {
    showToast('POPUP BLOCKED — ALLOW POPUPS FOR THIS SITE.', 'warn');
    return;
  }
  function _relayToMap() {
    if (_popoutMapWindow && !_popoutMapWindow.closed && TOKEN) {
      _popoutMapWindow.postMessage({ type: 'HOSCAD_RELAY_TOKEN', token: TOKEN }, window.location.origin);
    }
  }
  _popoutMapWindow.addEventListener('load', _relayToMap);
  window.addEventListener('message', function _relayMapHandler(e) {
    if (e.origin !== window.location.origin) return;
    if (e.data && e.data.type === 'HOSCAD_REQUEST_RELAY_TOKEN') {
      window.removeEventListener('message', _relayMapHandler);
      _relayToMap();
    }
  });
  const check = setInterval(() => {
    if (_popoutMapWindow && _popoutMapWindow.closed) {
      clearInterval(check);
      _popoutMapWindow = null;
    }
  }, 3000);
  showToast('MAP OPENED IN POPOUT.');
}

// ============================================================
// Initialization
// ============================================================
function updateClock() {
  const el = document.getElementById('clockPill');
  if (!el) return;
  const now = new Date();
  el.textContent = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
  updatePopoutClock();
  updateDc911SyncBadge();
}

// ─── Resizable Columns ───────────────────────────────────────────────────────
const _COL_NAMES = ['unit', 'status', 'elapsed', 'dest', 'note', 'inc', 'updated'];

function _saveColWidths() {
  const ths = document.querySelectorAll('#boardTable thead th');
  const w = {};
  ths.forEach((th, i) => {
    if (_COL_NAMES[i] !== 'note') w[_COL_NAMES[i]] = th.offsetWidth;
  });
  try { localStorage.setItem('hoscad_col_widths', JSON.stringify(w)); } catch(e) {}
}

function _loadColWidths() {
  try {
    const saved = JSON.parse(localStorage.getItem('hoscad_col_widths') || 'null');
    if (!saved) return;
    const ths = document.querySelectorAll('#boardTable thead th');
    ths.forEach((th, i) => {
      const name = _COL_NAMES[i];
      if (name !== 'note' && saved[name]) th.style.width = saved[name] + 'px';
    });
  } catch(e) {}
}

function initColumnResize() {
  _loadColWidths();
  const ths = document.querySelectorAll('#boardTable thead th');
  ths.forEach((th, i) => {
    const handle = th.querySelector('.col-resize-handle');
    if (!handle) return; // col-note has no handle
    let startX, startW;
    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      e.stopPropagation(); // don't trigger sort
      startX = e.clientX;
      startW = th.offsetWidth;
      handle.classList.add('is-resizing');
      document.body.style.cursor = 'col-resize';
      function onMove(ev) {
        th.style.width = Math.max(36, startW + (ev.clientX - startX)) + 'px';
      }
      function onUp() {
        handle.classList.remove('is-resizing');
        document.body.style.cursor = '';
        _saveColWidths();
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  });
}
// ─────────────────────────────────────────────────────────────────────────────

// ─── Global Hover Tooltip System ────────────────────────────────────────────
let _tipTimer = null;

function _hideTip() {
  clearTimeout(_tipTimer);
  const tip = document.getElementById('globalTip');
  if (tip) tip.style.display = 'none';
}

function _showTip(anchorEl, html, delay) {
  clearTimeout(_tipTimer);
  _tipTimer = setTimeout(() => {
    const tip = document.getElementById('globalTip');
    if (!tip) return;
    tip.innerHTML = html;
    tip.style.display = 'block';
    const rect = anchorEl.getBoundingClientRect();
    const tipW = 270;
    let left = rect.left;
    let top = rect.bottom + 6;
    if (left + tipW > window.innerWidth - 8) left = Math.max(4, window.innerWidth - tipW - 8);
    if (top + 220 > window.innerHeight) top = rect.top - tip.offsetHeight - 6;
    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  }, delay !== undefined ? delay : 350);
}

function _buildUnitTip(unitId) {
  if (!STATE) return null;
  const u = (STATE.units || []).find(x => x.unit_id === unitId);
  if (!u) return null;
  const agencyStr = u.agency_id || '—';
  const typeStr = [u.type ? u.type.toUpperCase() : null, u.level].filter(Boolean).join(' / ') || '—';
  const stationStr = u.station || '—';
  const assistOnly = u.include_in_recommendations === false;
  const mi = minutesSince(u.updated_at);
  const elStr = mi != null ? formatElapsed(mi) + ' ago' : '—';
  const updStr = fmtTime24(u.updated_at) + (u.updated_by ? ' by ' + u.updated_by.toUpperCase() : '');
  const crewParts = u.unit_info ? String(u.unit_info).split('|').filter(p => /^CM\d:/i.test(p)) : [];
  const etaTipMatch = (u.note || '').match(/\[ETA:(\d+)\]/);
  const rows = [];
  if (crewParts.length) rows.push(['Crew', crewParts.map(p => p.replace(/^CM\d:/i, '').trim()).join(', ')]);
  if (etaTipMatch) rows.push(['ETA', etaTipMatch[1] + ' min']);
  rows.push(
    ['Agency', agencyStr],
    ['Type', typeStr],
    ['Station', stationStr],
    ['Recommend', assistOnly ? 'Assist only' : 'Yes'],
    ['Updated', updStr],
    ['Elapsed', elStr]
  );
  const titleLine = esc(u.unit_id.toUpperCase()) +
    (u.display_name && u.display_name.toUpperCase() !== u.unit_id.toUpperCase() ? ' — ' + esc(u.display_name.toUpperCase()) : '');
  return '<div class="tip-title">' + titleLine + '</div>' +
    rows.map(([k, v]) => '<div class="tip-row"><span class="tip-key">' + k + '</span><span class="tip-val">' + esc(String(v)) + '</span></div>').join('');
}

function _buildIncTip(incId) {
  if (!STATE) return null;
  const inc = (STATE.incidents || []).find(x => x.incident_id === incId);
  if (!inc) return null;
  const unitCount = (STATE.units || []).filter(u => u.active && u.incident === incId).length;
  const mi = minutesSince(inc.dispatch_time || inc.created_at);
  const ageStr = mi != null ? formatElapsed(mi) + ' ago' : '—';
  const note = (inc.incident_note || '').replace(/\[[^\]]*\]/g, '').replace(/\s{2,}/g, ' ').trim();
  const rows = [
    ['Type', inc.incident_type || '—'],
    ['Priority', inc.priority || '—'],
    ['Source', 'Manual'],
    ['Units', unitCount + ' assigned'],
    ['Scene', inc.scene_address || '—'],
    ['Dispatched', ageStr]
  ];
  if (note) rows.push(['Note', note.length > 90 ? note.substring(0, 90) + '…' : note]);
  return '<div class="tip-title">' + esc(inc.incident_id) + '</div>' +
    rows.map(([k, v]) => '<div class="tip-row"><span class="tip-key">' + k + '</span><span class="tip-val">' + esc(String(v)) + '</span></div>').join('');
}

function _initTooltipSystem() {
  document.addEventListener('mouseover', e => {
    if (e.target.closest('#globalTip')) return;
    // INC# span — incident details on the board
    const incSpan = e.target.closest('.clickableIncidentNum');
    if (incSpan && incSpan.dataset.inc) {
      const html = _buildIncTip(incSpan.dataset.inc);
      if (html) { _showTip(incSpan, html); return; }
    }
    // Active Calls Bar card
    const acbCard = e.target.closest('.acb-card[data-inc-id]');
    if (acbCard) {
      const html = _buildIncTip(acbCard.dataset.incId);
      if (html) { _showTip(acbCard, html); return; }
    }
    // Incident queue row
    const incTr = e.target.closest('tr[data-inc-id]');
    if (incTr) {
      const html = _buildIncTip(incTr.dataset.incId);
      if (html) { _showTip(incTr, html); return; }
    }
    // Unit board row
    const unitTr = e.target.closest('tr[data-unit-id]');
    if (unitTr) {
      const html = _buildUnitTip(unitTr.dataset.unitId);
      if (html) { _showTip(unitTr, html); return; }
    }
    _hideTip();
  });
  document.addEventListener('scroll', _hideTip, true);
  document.addEventListener('click', _hideTip, true);
}
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// Realtime — WebSocket triggered-poll (Supabase Realtime, Phoenix protocol)
// On any postgres_changes event for units or incidents, triggers refresh().
// refresh() is guarded by _refreshing so concurrent calls are safely dropped.
// ─────────────────────────────────────────────────────────────────────────────
function _rtSend(msg) {
  if (_RT && _RT.readyState === WebSocket.OPEN) {
    try { _RT.send(JSON.stringify(msg)); } catch(e) {}
  }
}

function _rtConnect() {
  if (_RT && (_RT.readyState === WebSocket.OPEN || _RT.readyState === WebSocket.CONNECTING)) return;
  const url = 'wss://vnqiqxffedudfsdoadqg.supabase.co/realtime/v1/websocket?apikey=' + API._apiKey + '&vsn=1.0.0';
  try { _RT = new WebSocket(url); } catch(e) { return; }
  _RT.onopen = function() {
    _rtRef = 0;
    _rtSend({ topic: 'realtime:units', event: 'phx_join', payload: {
      config: { broadcast: { self: false }, presence: { key: '' },
        postgres_changes: [{ event: '*', schema: 'public', table: 'units' }] }
    }, ref: String(++_rtRef) });
    _rtSend({ topic: 'realtime:incidents', event: 'phx_join', payload: {
      config: { broadcast: { self: false }, presence: { key: '' },
        postgres_changes: [{ event: '*', schema: 'public', table: 'incidents' }] }
    }, ref: String(++_rtRef) });
    if (_rtHbTimer) clearInterval(_rtHbTimer);
    _rtHbTimer = setInterval(function() {
      _rtSend({ topic: 'phoenix', event: 'heartbeat', payload: {}, ref: 'hb' });
    }, 25000);
    refresh(); // catch any state changes missed during disconnect gap
  };
  _RT.onmessage = function(e) {
    try {
      const msg = JSON.parse(e.data);
      if (msg.event === 'postgres_changes') {
        refresh();
      }
    } catch(err) {}
  };
  _RT.onclose = function() {
    if (_rtHbTimer) { clearInterval(_rtHbTimer); _rtHbTimer = null; }
    if (!TOKEN) return; // Logged out — do not reconnect
    if (_rtReconTimer) clearTimeout(_rtReconTimer);
    _rtReconTimer = setTimeout(_rtConnect, 5000);
  };
  _RT.onerror = function() {
    try { _RT.close(); } catch(e) {}
  };
}

function _rtDisconnect() {
  if (_rtHbTimer) { clearInterval(_rtHbTimer); _rtHbTimer = null; }
  if (_rtReconTimer) { clearTimeout(_rtReconTimer); _rtReconTimer = null; }
  if (_RT) { try { _RT.close(); } catch(e) {} _RT = null; }
}
// ─────────────────────────────────────────────────────────────────────────────

async function start() {
  // Refresh positions from live DB (fallback already applied at DOMContentLoaded)
  try {
    const initRes = await API.init();
    if (initRes && initRes.ok && Array.isArray(initRes.positions) && initRes.positions.length > 0) {
      POSITIONS_META = initRes.positions;
      _populateLoginRoleDropdown(initRes.positions.filter(p => p.is_dispatcher));
    }
  } catch (_) {}
  loadViewState();
  refresh();
  AddressLookup.load(); // async, non-blocking — autocomplete works once data arrives
  if (POLL) clearInterval(POLL);
  POLL = setInterval(refresh, 10000);
  _rtConnect();
  startLfnPolling();
  _applyDc911Btn(_isDc911Enabled());
  updateClock();
  var _clockInterval = setInterval(updateClock, 1000);
  let _searchDebounce;
  document.getElementById('search').addEventListener('input', () => {
    clearTimeout(_searchDebounce);
    _searchDebounce = setTimeout(renderBoardDiff, 180);
  });
  document.getElementById('showInactive').addEventListener('change', renderBoardDiff);
  setupColumnSort();
  applyViewState();
  loadScratch();

  // Persistent token relay — serves viewer, board popout, and inc queue popout
  window.addEventListener('message', function _globalRelay(e) {
    if (e.origin !== window.location.origin) return;
    if (e.data && e.data.type === 'HOSCAD_REQUEST_RELAY_TOKEN' && TOKEN && e.source) {
      try { e.source.postMessage({ type: 'HOSCAD_RELAY_TOKEN', token: TOKEN }, window.location.origin); } catch(err){}
    }
    if (e.data && e.data.type === 'HOSCAD_REVIEW_INCIDENT' && e.data.incidentId) {
      openIncident(e.data.incidentId);
    }
  });

  // Staleness indicator — if no successful poll in >30s, warn dispatcher
  setInterval(function() {
    if (!_lastPollAt || !TOKEN) return;
    const staleMs = Date.now() - _lastPollAt;
    if (staleMs > 30000) {
      const e = document.getElementById('livePill');
      if (e && !e.className.includes('stale')) {
        e.className = 'pill stale';
        e.textContent = 'STALE (' + Math.round(staleMs / 1000) + 's)';
      } else if (e && e.className.includes('stale')) {
        e.textContent = 'STALE (' + Math.round(staleMs / 1000) + 's)';
      }
    }
  }, 5000);

  // Throttle polling when tab is hidden (60s) vs visible (10s)
  // Also pause/resume clock and flush pending renders
  document.addEventListener('visibilitychange', function() {
    if (POLL) clearInterval(POLL);
    if (document.hidden) {
      POLL = setInterval(refresh, 60000);
      clearInterval(_clockInterval);
    } else {
      POLL = setInterval(refresh, 10000);
      _clockInterval = setInterval(updateClock, 1000);
      updateClock();
      // Flush any pending render from background updates
      if (_pendingRender) {
        _pendingRender = false;
        renderAll();
      }
    }
  });
}

// DOM Ready
window.addEventListener('load', () => {
  // Attach address autocomplete to destination inputs
  AddrAutocomplete.attach(document.getElementById('mDestination'));
  AddrAutocomplete.attach(document.getElementById('incDestEdit'));

  // Attach address autocomplete to scene fields — onSelect fills street address
  AddrAutocomplete.attach(document.getElementById('newIncScene'), {
    onSelect: function(addr) {
      var el = document.getElementById('newIncScene');
      if (el && addr.address) {
        el.value = (addr.address + ', ' + addr.city + ', ' + addr.state + ' ' + addr.zip).toUpperCase();
      }
    }
  });
  AddrAutocomplete.attach(document.getElementById('incSceneAddress'), {
    onSelect: function(addr) {
      var el = document.getElementById('incSceneAddress');
      if (el && addr.address) {
        el.value = (addr.address + ', ' + addr.city + ', ' + addr.state + ' ' + addr.zip).toUpperCase();
      }
    }
  });

  // Incident modal: Ctrl+Enter saves note
  document.getElementById('incNote').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.ctrlKey) {
      e.preventDefault();
      saveIncidentNote();
    }
  });

  // Setup login form
  document.getElementById('loginRole').value = '';
  document.getElementById('loginCadId').value = '';
  document.getElementById('loginPassword').value = '';

  document.getElementById('loginRole').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') document.getElementById('loginCadId').focus();
  });

  document.getElementById('loginCadId').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') document.getElementById('loginPassword').focus();
  });

  document.getElementById('loginPassword').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') login();
  });

  // Setup command input
  const cI = document.getElementById('cmd');
  cI.addEventListener('input', () => {
    showCmdHints(cI.value.trim());
  });
  cI.addEventListener('keydown', (e) => {
    // Cmd hints navigation
    const hintsOpen = document.getElementById('cmdHints') && document.getElementById('cmdHints').classList.contains('open');
    if (e.key === 'Escape' && hintsOpen) { hideCmdHints(); e.preventDefault(); return; }
    if (e.key === 'Enter') {
      if (hintsOpen && CMD_HINT_INDEX >= 0) { selectCmdHint(CMD_HINT_INDEX); e.preventDefault(); return; }
      hideCmdHints();
      e.preventDefault();
      e.stopPropagation();
      runCommand();
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (hintsOpen) { navigateCmdHints(-1); return; }
      if (CMD_INDEX > 0) {
        CMD_INDEX--;
        cI.value = CMD_HISTORY[CMD_INDEX] || '';
      }
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (hintsOpen) { navigateCmdHints(1); return; }
      if (CMD_INDEX < CMD_HISTORY.length - 1) {
        CMD_INDEX++;
        cI.value = CMD_HISTORY[CMD_INDEX] || '';
      } else {
        CMD_INDEX = CMD_HISTORY.length;
        cI.value = '';
      }
    }
  });
  cI.addEventListener('blur', () => {
    setTimeout(hideCmdHints, 150);
  });

  // Global keyboard shortcuts
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      const nib = document.getElementById('newIncBack');
      const uhb = document.getElementById('uhBack');
      const ib = document.getElementById('incBack');
      const mb = document.getElementById('modalBack');
      const msgb = document.getElementById('msgBack');
      const cd = document.getElementById('confirmDialog');
      const ad = document.getElementById('alertDialog');

      const brb = document.getElementById('bugReportBack');
      const srb = document.getElementById('searchBack');
      if (nib && nib.style.display === 'flex') { closeNewIncident(); return; }
      if (brb && brb.style.display === 'flex') { closeBugReport(); return; }
      if (srb && srb.style.display === 'flex') { closeSearchPanel(); return; }
      if (uhb && uhb.style.display === 'flex') { uhb.style.display = 'none'; autoFocusCmd(); return; }
      if (ib && ib.style.display === 'flex') { ib.style.display = 'none'; autoFocusCmd(); return; }
      if (msgb && msgb.style.display === 'flex') { closeMessages(); return; }
      if (mb && mb.style.display === 'flex') { closeModal(); return; }
      if (cd && cd.classList.contains('active')) { hideConfirm(); return; }
      if (ad && ad.classList.contains('active')) { hideAlert(); return; }

      // Escape also deselects
      if (SELECTED_UNIT_ID) {
        SELECTED_UNIT_ID = null;
        document.querySelectorAll('#boardBody tr.selected').forEach(tr => tr.classList.remove('selected'));
      }
    }

    if (e.ctrlKey && e.key === 'k') { e.preventDefault(); cI.focus(); }
    if (e.ctrlKey && e.key === 'l') { e.preventDefault(); openLogon(); }
    if (e.ctrlKey && e.key === 'd') { e.preventDefault(); cycleDensity(); }
    if (e.ctrlKey && e.key === 'f') { e.preventDefault(); openSearchPanel(); }
    if (e.ctrlKey && e.key === 'm') { e.preventDefault(); toggleBoardMap(); }
    if (e.ctrlKey && e.key === 'b') { e.preventDefault(); openBugReport(); }
    if (e.key === 'F1') { e.preventDefault(); cI.focus(); }
    if (e.key === 'F2') { e.preventDefault(); openNewIncident(); }
    if (e.key === 'F3') { e.preventDefault(); openSearchPanel(); }
    if (e.key === 'F4') { e.preventDefault(); openMessages(); }
    if (e.key === 'F5') { e.preventDefault(); toggleBoardMap(); }
    if (e.key === 'F6') { e.preventDefault(); _execCmd('INCQ'); }
    if (e.key === 'F7') { e.preventDefault(); _execCmd('WHO'); }
    if (e.key === 'F8') { e.preventDefault(); _execCmd('UR'); }
    if (e.key === 'F9') {
      e.preventDefault();
      const role = localStorage.getItem('ems_role') || '';
      const isSupervisor = ['SUPV1','SUPV2','MGR1','MGR2','IT'].includes(role.toUpperCase());
      if (isSupervisor) _execCmd('SHIFT REPORT'); else showToast('F9 SHIFT REPORT — SUPV/MGR/IT ONLY.');
    }
  });

  // Confirm dialog handlers
  document.getElementById('confirmOk').addEventListener('click', () => {
    const cb = CONFIRM_CALLBACK;
    hideConfirm();
    if (cb) cb(true);
  });

  document.getElementById('confirmClose').addEventListener('click', () => {
    const cb = CONFIRM_CANCEL_CALLBACK;
    hideConfirm();
    if (cb) cb();
  });

  document.getElementById('alertClose').addEventListener('click', () => {
    hideAlert();
  });

  // Enter key closes dialogs and modals
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      // Don't intercept Enter in textareas or inputs - let their handlers deal with it
      const tag = e.target.tagName;
      const isTextarea = tag === 'TEXTAREA';
      const isInput = tag === 'INPUT' && e.target.type !== 'button';
      if (isTextarea || isInput) return;

      // Alert/Confirm dialogs - close on Enter (only when not in an input)
      const alertDialog = document.getElementById('alertDialog');
      const confirmDialog = document.getElementById('confirmDialog');
      if (alertDialog.classList.contains('active')) {
        e.preventDefault();
        hideAlert();
        return;
      }
      if (confirmDialog.classList.contains('active')) {
        e.preventDefault();
        const cb = CONFIRM_CALLBACK;
        hideConfirm();
        if (cb) cb(true);
        return;
      }

      // Close other modals on Enter (when not in an input field)
      const uhBack = document.getElementById('uhBack');
      const msgBack = document.getElementById('msgBack');
      if (uhBack && uhBack.style.display === 'flex') {
        closeUH();
        return;
      }
      if (msgBack && msgBack.style.display === 'flex') {
        closeMessages();
        return;
      }
    }
  });

  // Performance: Event delegation for board table (instead of per-row handlers)
  const boardBody = document.getElementById('boardBody');
  if (boardBody) {
    // Single click = select row
    boardBody.addEventListener('click', (e) => {
      // Check if clicked on stack badge
      const stackEl = e.target.closest('[data-stack-unit]');
      if (stackEl) {
        e.stopPropagation();
        const uid = stackEl.getAttribute('data-stack-unit');
        if (_expandedStacks.has(uid)) _expandedStacks.delete(uid);
        else _expandedStacks.add(uid);
        renderBoardDiff();
        return;
      }

      // Check if clicked on incident number
      const incEl = e.target.closest('.clickableIncidentNum');
      if (incEl) {
        e.stopPropagation();
        const incId = incEl.dataset.inc;
        if (incId) openIncident(incId);
        return;
      }

      // Otherwise select the row
      const tr = e.target.closest('tr');
      if (tr && tr.dataset.unitId) {
        e.stopPropagation();
        selectUnit(tr.dataset.unitId);
      }
    });

    // Double click = open edit modal
    boardBody.addEventListener('dblclick', (e) => {
      const tr = e.target.closest('tr');
      if (tr && tr.dataset.unitId) {
        e.preventDefault();
        e.stopPropagation();
        const u = (STATE.units || []).find(u => u.unit_id === tr.dataset.unitId);
        if (u) {
          if (u.source && u.source.startsWith('DC911:')) { showToast('DC911 UNIT — READ-ONLY. DATA FROM DESCHUTES 911 CAD FEED.', 'info'); return; }
          if (u.source && u.source.startsWith('PP:')) { openPpUnitActions(u); return; }
          openModal(u);
        }
      }
    });
  }

  // Session cleanup on tab close — immediately end session so position becomes available
  window.addEventListener('beforeunload', () => {
    if (TOKEN) {
      try {
        fetch(API.baseUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
            'apikey': API._apiKey,
            'Authorization': 'Bearer ' + API._apiKey
          },
          body: new URLSearchParams({ action: 'logout', params: JSON.stringify([TOKEN]) }).toString(),
          keepalive: true
        });
      } catch(e) {}
    }
  });

  // Show login screen
  document.getElementById('loginBack').style.display = 'flex';
  document.getElementById('userLabel').textContent = '—';

  _initTooltipSystem();
  initColumnResize();
});
