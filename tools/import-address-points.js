#!/usr/bin/env node
/**
 * HOSCAD GIS Address Points Importer
 * ====================================
 * Fetches authoritative address point data from ArcGIS REST services and
 * populates the Supabase `address_points` table for CAD address typeahead
 * and canonical address matching.
 *
 * USAGE:
 *   node tools/import-address-points.js [--county deschutes] [--source statewide] [--dry-run]
 *
 * REQUIRED ENVIRONMENT VARIABLES:
 *   SUPABASE_URL          https://vnqiqxffedudfsdoadqg.supabase.co
 *   SUPABASE_SERVICE_KEY  Get from: Supabase Dashboard → Project Settings → API → service_role key
 *
 * SETUP FOR NEW AGENCIES (onboarding):
 *   1. Find the county E911 ArcGIS service URL (county GIS department)
 *   2. Or use the Oregon statewide service and filter by county name
 *   3. Add a new SOURCE_CONFIG entry below
 *   4. Set env vars and run
 *
 * REFRESH SCHEDULE: Run monthly (or after major address changes in the county)
 */

const https = require('https');
const http  = require('http');

// ── Configuration ─────────────────────────────────────────────────────────────

const SUPABASE_URL         = process.env.SUPABASE_URL         || 'https://vnqiqxffedudfsdoadqg.supabase.co';
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || '';

if (!SUPABASE_SERVICE_KEY) {
  console.error('ERROR: SUPABASE_SERVICE_KEY is required.');
  console.error('  Get it from: Supabase Dashboard → Project Settings → API → service_role key');
  console.error('  Set it: SUPABASE_SERVICE_KEY=eyJ... node tools/import-address-points.js');
  process.exit(1);
}

const BATCH_SIZE   = 500;   // Records per Supabase upsert
const PAGE_SIZE    = 2000;  // ArcGIS max records per request
const REQUEST_DELAY_MS = 200; // Polite delay between ArcGIS pages

/**
 * ArcGIS data sources.
 * Add entries here for new agency onboarding.
 * Each entry maps to a specific ArcGIS FeatureService layer.
 */
const SOURCE_CONFIG = {
  // ── Deschutes County E911 (primary — has dedicated cad_address field) ──────
  deschutes: {
    label:    'Deschutes County E911',
    url:      'https://maps.deschutes.org/server/rest/services/Hosted/E911_Address_Points/FeatureServer/0/query',
    county:   'DESCHUTES',
    source:   'DESCHUTES_E911',
    outFields: 'cad_address,add_number,st_predir,st_name,st_postyp,st_posdir,postal_community,zipcode,latitude,longitude',
    where:    '1=1',
    returnGeometry: false,  // lat/lon already in attribute fields
    mapRecord: (a) => ({
      full_address: (a.cad_address || '').trim().toUpperCase(),
      house_num:    String(a.add_number || ''),
      pre_dir:      (a.st_predir      || '').toUpperCase().trim(),
      street_name:  (a.st_name        || '').toUpperCase().trim(),
      street_type:  (a.st_postyp      || '').toUpperCase().trim(),
      post_dir:     (a.st_posdir      || '').toUpperCase().trim(),
      city:         (a.postal_community || '').toUpperCase().trim(),
      zip:          String(a.zipcode  || '').trim(),
      lat:          a.latitude,
      lon:          a.longitude,
      county:       'DESCHUTES',
      source:       'DESCHUTES_E911',
    }),
  },

  // ── Oregon Statewide — Crook County ─────────────────────────────────────────
  crook: {
    label:    'Crook County (Oregon Statewide)',
    url:      'https://services8.arcgis.com/8PAo5HGmvRMlF2eU/arcgis/rest/services/Oregon_Address_Points/FeatureServer/0/query',
    county:   'CROOK',
    source:   'OR_STATEWIDE',
    outFields: 'ADDRESS_FULL,Add_Number,St_PreDir,St_Name,St_PosTyp,St_PosDir,Post_Comm,Post_Code,Latitude,Longitude',
    where:    "County='Crook County'",
    mapRecord: (a) => ({
      full_address: (a.ADDRESS_FULL || '').trim().toUpperCase(),
      house_num:    String(a.Add_Number || ''),
      pre_dir:      (a.St_PreDir || '').toUpperCase().trim(),
      street_name:  (a.St_Name   || '').toUpperCase().trim(),
      street_type:  (a.St_PosTyp || '').toUpperCase().trim(),
      post_dir:     (a.St_PosDir || '').toUpperCase().trim(),
      city:         (a.Post_Comm || '').toUpperCase().trim(),
      zip:          String(a.Post_Code || '').trim(),
      lat:          a.Latitude,
      lon:          a.Longitude,
      county:       'CROOK',
      source:       'OR_STATEWIDE',
    }),
  },

  // ── Oregon Statewide — Jefferson County ─────────────────────────────────────
  jefferson: {
    label:    'Jefferson County (Oregon Statewide)',
    url:      'https://services8.arcgis.com/8PAo5HGmvRMlF2eU/arcgis/rest/services/Oregon_Address_Points/FeatureServer/0/query',
    county:   'JEFFERSON',
    source:   'OR_STATEWIDE',
    outFields: 'ADDRESS_FULL,Add_Number,St_PreDir,St_Name,St_PosTyp,St_PosDir,Post_Comm,Post_Code,Latitude,Longitude',
    where:    "County='Jefferson County'",
    mapRecord: (a) => ({
      full_address: (a.ADDRESS_FULL || '').trim().toUpperCase(),
      house_num:    String(a.Add_Number || ''),
      pre_dir:      (a.St_PreDir || '').toUpperCase().trim(),
      street_name:  (a.St_Name   || '').toUpperCase().trim(),
      street_type:  (a.St_PosTyp || '').toUpperCase().trim(),
      post_dir:     (a.St_PosDir || '').toUpperCase().trim(),
      city:         (a.Post_Comm || '').toUpperCase().trim(),
      zip:          String(a.Post_Code || '').trim(),
      lat:          a.Latitude,
      lon:          a.Longitude,
      county:       'JEFFERSON',
      source:       'OR_STATEWIDE',
    }),
  },

  // ── Oregon Statewide — Lake County ──────────────────────────────────────────
  lake: {
    label:    'Lake County (Oregon Statewide)',
    url:      'https://services8.arcgis.com/8PAo5HGmvRMlF2eU/arcgis/rest/services/Oregon_Address_Points/FeatureServer/0/query',
    county:   'LAKE',
    source:   'OR_STATEWIDE',
    outFields: 'ADDRESS_FULL,Add_Number,St_PreDir,St_Name,St_PosTyp,St_PosDir,Post_Comm,Post_Code,Latitude,Longitude',
    where:    "County='Lake County'",
    mapRecord: (a) => ({
      full_address: (a.ADDRESS_FULL || '').trim().toUpperCase(),
      house_num:    String(a.Add_Number || ''),
      pre_dir:      (a.St_PreDir || '').toUpperCase().trim(),
      street_name:  (a.St_Name   || '').toUpperCase().trim(),
      street_type:  (a.St_PosTyp || '').toUpperCase().trim(),
      post_dir:     (a.St_PosDir || '').toUpperCase().trim(),
      city:         (a.Post_Comm || '').toUpperCase().trim(),
      zip:          String(a.Post_Code || '').trim(),
      lat:          a.Latitude,
      lon:          a.Longitude,
      county:       'LAKE',
      source:       'OR_STATEWIDE',
    }),
  },

  // ── TEMPLATE for new agency onboarding ──────────────────────────────────────
  // NEW_COUNTY: {
  //   label:    'New County (Source)',
  //   url:      'https://county-gis-server/rest/services/.../FeatureServer/0/query',
  //   county:   'NEW_COUNTY',
  //   source:   'COUNTY_E911' or 'OR_STATEWIDE',
  //   outFields: '...',
  //   where:    "County='New County'" or '1=1',
  //   mapRecord: (a) => ({ full_address, house_num, pre_dir, ... lat, lon }),
  // },
};

// ── Address normalization (must match incidents.ts normalizeAddress exactly) ──

function normalizeAddress(raw) {
  if (!raw) return '';
  let s = raw.trim().toUpperCase();
  s = s.replace(/\s+/g, ' ');
  s = s.replace(/,\s*[A-Z\s]+,?\s*(?:OR|WA|CA|ID|NV)?\s*\d{0,5}$/i, '').trim();
  s = s.replace(/,/g, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(/\s+(?:APT|UNIT|STE|SUITE|#)\s*\S+$/i, '').trim();
  s = s.replace(/^0+(\d)/, '$1');
  s = s.replace(/\bNORTHEAST\b/g, 'NE').replace(/\bNORTHWEST\b/g, 'NW')
       .replace(/\bSOUTHEAST\b/g, 'SE').replace(/\bSOUTHWEST\b/g, 'SW')
       .replace(/\bNORTH\b/g, 'N').replace(/\bSOUTH\b/g, 'S')
       .replace(/\bEAST\b/g, 'E').replace(/\bWEST\b/g, 'W');
  s = s.replace(/\b(?:STREET|STR)\b/g, 'ST').replace(/\bAVENUE\b/g, 'AVE')
       .replace(/\bBOULEVARD\b/g, 'BLVD').replace(/\bDRIVE\b/g, 'DR')
       .replace(/\bROAD\b/g, 'RD').replace(/\bHIGHWAY\b/g, 'HWY')
       .replace(/\bLANE\b/g, 'LN').replace(/\bCOURT\b/g, 'CT')
       .replace(/\bPLACE\b/g, 'PL');
  return s.trim();
}

// ── HTTP helpers ──────────────────────────────────────────────────────────────

function fetchJson(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith('https') ? https : http;
    lib.get(url, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(body)); }
        catch (e) { reject(new Error('JSON parse failed: ' + e.message)); }
      });
    }).on('error', reject);
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── Supabase REST upsert ──────────────────────────────────────────────────────

function supabaseUpsert(rows) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(rows);
    const parsedUrl = new URL(SUPABASE_URL + '/rest/v1/address_points');
    const options = {
      hostname: parsedUrl.hostname,
      path:     parsedUrl.pathname + '?on_conflict=full_address%2Ccounty',
      method:   'POST',
      headers: {
        'Content-Type':  'application/json',
        'Content-Length': Buffer.byteLength(body),
        'apikey':         SUPABASE_SERVICE_KEY,
        'Authorization':  'Bearer ' + SUPABASE_SERVICE_KEY,
        'Prefer':         'resolution=merge-duplicates',
      },
    };
    const req = https.request(options, (res) => {
      let resp = '';
      res.on('data', chunk => resp += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) resolve();
        else reject(new Error('Supabase upsert failed ' + res.statusCode + ': ' + resp.substring(0, 200)));
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ── ArcGIS paginator ─────────────────────────────────────────────────────────

async function fetchAllRecords(config) {
  const records = [];
  let offset = 0;
  let page   = 0;

  while (true) {
    const params = new URLSearchParams({
      where:             config.where,
      outFields:         config.outFields,
      returnGeometry:    config.returnGeometry !== false ? 'true' : 'false',
      outSR:             '4326',  // WGS84 lat/lon
      resultOffset:      String(offset),
      resultRecordCount: String(PAGE_SIZE),
      f:                 'json',
    });
    const url = config.url + '?' + params.toString();

    let data;
    try {
      data = await fetchJson(url);
    } catch (e) {
      console.error(`    Page ${page} fetch error: ${e.message}. Retrying once...`);
      await sleep(1000);
      data = await fetchJson(url);
    }

    if (!data.features || data.features.length === 0) break;

    for (const feature of data.features) {
      const attrs = feature.attributes || feature;
      const rec   = config.mapRecord(attrs);
      if (!rec.full_address) continue;
      // Try to get lat/lon from geometry if not in attributes
      if ((rec.lat == null || rec.lon == null) && feature.geometry) {
        rec.lon = feature.geometry.x ?? null;
        rec.lat = feature.geometry.y ?? null;
      }
      rec.canonical   = normalizeAddress(rec.full_address);
      rec.updated_at  = new Date().toISOString();
      records.push(rec);
    }

    page++;
    offset += data.features.length;
    process.stdout.write(`\r    Page ${page} — ${records.length} records fetched...`);

    if (data.features.length < PAGE_SIZE) break;
    await sleep(REQUEST_DELAY_MS);
  }

  process.stdout.write('\n');
  return records;
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function importSource(key, dryRun) {
  const config = SOURCE_CONFIG[key];
  if (!config) { console.error(`Unknown source: ${key}`); return; }

  console.log(`\n[${key.toUpperCase()}] ${config.label}`);
  console.log('  Fetching from ArcGIS...');

  const rawRecords = await fetchAllRecords(config);
  console.log(`  ${rawRecords.length} valid records fetched.`);

  // Deduplicate by (full_address, county) — keep last occurrence (most recent data wins)
  const seen = new Map();
  for (const r of rawRecords) {
    seen.set(r.full_address + '|' + r.county, r);
  }
  const records = Array.from(seen.values());
  if (records.length < rawRecords.length) {
    console.log(`  ${rawRecords.length - records.length} duplicates removed → ${records.length} unique records.`);
  }

  if (dryRun) {
    console.log('  DRY RUN — sample (first 3):');
    records.slice(0, 3).forEach(r => console.log('   ', JSON.stringify(r)));
    return;
  }

  console.log(`  Upserting to Supabase in batches of ${BATCH_SIZE}...`);
  let inserted = 0;
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    const batch = records.slice(i, i + BATCH_SIZE);
    await supabaseUpsert(batch);
    inserted += batch.length;
    process.stdout.write(`\r  ${inserted}/${records.length} upserted...`);
  }
  process.stdout.write('\n');
  console.log(`  [DONE] ${inserted} records imported for ${config.label}.`);
}

async function main() {
  const args    = process.argv.slice(2);
  const dryRun  = args.includes('--dry-run');
  const county  = (args.find(a => a.startsWith('--county=')) || '').replace('--county=', '');
  const all     = args.includes('--all') || (!county);

  console.log('HOSCAD GIS Address Points Importer');
  console.log('===================================');
  if (dryRun) console.log('DRY RUN MODE — nothing will be written to Supabase\n');

  const toRun = county
    ? [county]
    : Object.keys(SOURCE_CONFIG);  // default: all sources

  for (const key of toRun) {
    await importSource(key, dryRun);
  }

  console.log('\nAll done.');
  if (!dryRun) {
    console.log('\nNEXT STEP: Run rebackfillCanonicalAddresses from the admin panel (IT role)');
    console.log('to ensure all existing incidents have canonical_address set correctly.');
  }
}

main().catch(e => { console.error('FATAL:', e.message); process.exit(1); });
