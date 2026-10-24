/**
 * WATER MONITOR — Tytheland shared supply, R S Hill master meter
 * Efento NB-IoT pulse counter 282C0242B88D, channel 1
 *
 * Every morning this script:
 *   1. Pulls the logger's readings from Efento Cloud (the same data the daily PDF report is drawn from).
 *   2. Measures the longest continuous period of ZERO flow inside the night window (default 01:00–05:00 UK time).
 *   3. Grades the night: CLEAR / WATCH / LEAK / NO DATA.
 *   4. Logs the result to "Nightly Log", keeps every pulse in "Raw Readings", and redraws the
 *      "Dashboard" tab (status banner, headline figures, hourly and night-by-night charts).
 *   5. Files that morning's Efento "NEW REPORT" email under a Gmail label and marks it read.
 *   6. Emails a summary every day to the daily list (so silence means the system has failed, never
 *      "all is well"), and copies the leak-alert list in only on nights graded LEAK.
 *
 * Install: create a Google Sheet, Extensions > Apps Script, replace Code.gs with this file, save,
 * reload the Sheet, then use the "Water Monitor" menu: Set up credentials > Find measurement point >
 * Diagnose API > Backfill nights > Install daily trigger.
 */

// ============================================================================================
// CONFIGURATION
// ============================================================================================

const CONFIG = {
  // Logger
  DEVICE_SERIAL: '282C0242B88D',
  CHANNEL: 1,

  // 'ACCUMULATING' = running total (what the 23 Oct report shows: 148 rising to 218).
  // 'PER_INTERVAL' = pulses counted in each measurement period.
  // Getting this wrong inverts the result, so the script also sanity-checks it every day.
  COUNTER_MODE: 'ACCUMULATING',

  // Litres per pulse of the meter's pulse output. Leave null until confirmed;
  // results are then reported in pulses only.
  LITRES_PER_PULSE: null,

  // Night window, UK local time. If START is later than END the window crosses midnight.
  TIMEZONE: 'Europe/London',
  NIGHT_START: '01:00',
  NIGHT_END: '05:00',

  // Grading, on the longest continuous zero-flow run inside the night window (real minutes).
  CLEAR_MIN_ZERO_RUN_MINUTES: 120, // at least this = CLEAR
  WATCH_MIN_ZERO_RUN_MINUTES: 60,  // at least this (but under CLEAR) = WATCH; below = LEAK

  // A day on which nine houses use no water at all is implausible: treat as a fault warning.
  FLAG_ZERO_FLOW_DAY: true,

  // Daily run time (UK). The logger transmits hourly, so 05:00 readings are in by ~06:00.
  RUN_AT_HOUR: 6,
  RUN_NEAR_MINUTE: 45,

  // Efento Cloud API
  API_BASE: 'https://cloud.efento.io/api/v2',
  API_TIMES_ARE_UTC: true, // Efento's API takes and returns UTC times
  FETCH_CHUNK_HOURS: 2,    // the logger records every minute, so keep each request to ~120 readings
  MAX_GAP_FACTOR: 3,       // gaps longer than this × normal interval are reported

  // Gmail handling of the Efento report email
  EFENTO_EMAIL_QUERY: 'from:mailer@efento.io subject:"NEW REPORT" newer_than:2d',
  EFENTO_EMAIL_LABEL: 'Water Monitor/Efento reports',
  MARK_EFENTO_EMAIL_READ: true,

  // Sheets
  SHEET_DASH: 'Dashboard',
  SHEET_CHART: 'Chart Data',
  SHEET_LOG: 'Nightly Log',
  SHEET_RAW: 'Raw Readings',
  SHEET_DIAG: 'Diagnostics',

  // Leak streak count shown in the email looks back this many logged nights
  STREAK_LOOKBACK: 60,

  // Nights shown on the Dashboard charts
  DASH_NIGHTS: 30,
};

const THEME = {
  NAVY: '#0F2A43',
  TEAL: '#0E7C86',
  TEAL_LIGHT: '#9FD3D7',
  INK: '#1F2933',
  MUTED: '#6B7785',
  TILE: '#F3F6F9',
  RULE: '#E1E6EB',
  BAND: '#F7F9FB',
  FONT: 'Roboto',
};

const PROP = {
  LOGIN: 'EFENTO_LOGIN',
  PASSWORD: 'EFENTO_PASSWORD',
  RECIPIENTS: 'ALERT_RECIPIENTS',          // daily email, every morning
  LEAK_RECIPIENTS: 'LEAK_ALERT_RECIPIENTS', // added only when the night is graded LEAK
  MP_ID: 'MEASUREMENT_POINT_ID',
};

const STATUS = {
  CLEAR: { code: 'CLEAR', text: 'No leak indicated', colour: '#1E8449' },
  WATCH: { code: 'WATCH', text: 'Borderline — watch', colour: '#D68910' },
  LEAK: { code: 'LEAK', text: 'Continuous overnight flow — leak likely', colour: '#C0392B' },
  NO_DATA: { code: 'NO DATA', text: 'No usable data for the night', colour: '#7F8C8D' },
  CHECK: { code: 'CHECK', text: 'No flow at all for 24 hours — check the logger and pulse cable', colour: '#6C3483' },
};

// Light background / dark text pairs for status chips in the sheets
const STATUS_CHIP = {
  'LEAK': ['#FADBD8', '#922B21'],
  'WATCH': ['#FDEBD0', '#9C640C'],
  'CLEAR': ['#D5F5E3', '#1D6F42'],
  'NO DATA': ['#EAECEE', '#566573'],
  'CHECK': ['#E8DAEF', '#6C3483'],
};

const LOG_HEADERS = [
  'Night (morning date)', 'Window (UK)', 'Status', 'Longest zero-flow run (min)',
  'Night pulses', 'Night litres', 'Night average L/h', 'Previous-day pulses',
  'Previous-day litres', 'Readings in window', 'Reading interval (min)', 'Notes', 'Checked at',
];

const RAW_HEADERS = ['Measured at (UTC)', 'Measured at (UK)', 'Channel value'];

// ============================================================================================
// MENU
// ============================================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Water Monitor')
    .addItem('1. Set up credentials and recipients…', 'setupCredentials')
    .addItem('Change email recipients…', 'setRecipients')
    .addItem('2. Find measurement point', 'findMeasurementPoint')
    .addItem('3. Diagnose API (raw data to Diagnostics sheet)', 'diagnoseApi')
    .addItem('4. Backfill previous nights…', 'backfillNights')
    .addItem('5. Install daily trigger', 'installDailyTrigger')
    .addSeparator()
    .addItem('Run check for last night now', 'runNightlyCheckManual')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addItem('Remove daily trigger', 'removeTriggers')
    .addToUi();
}

// ============================================================================================
// SETUP
// ============================================================================================

function setupCredentials() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const login = ui.prompt('Efento Cloud login', 'Email address you sign in to cloud.efento.io with:', ui.ButtonSet.OK_CANCEL);
  if (login.getSelectedButton() !== ui.Button.OK) return;

  const pwd = ui.prompt('Efento Cloud password', 'Stored in this script\'s private Script Properties only:', ui.ButtonSet.OK_CANCEL);
  if (pwd.getSelectedButton() !== ui.Button.OK) return;

  const defaultTo = Session.getEffectiveUser().getEmail();
  const rec = ui.prompt('Alert recipients', 'Comma-separated email addresses (blank = ' + defaultTo + '):', ui.ButtonSet.OK_CANCEL);
  if (rec.getSelectedButton() !== ui.Button.OK) return;

  props.setProperty(PROP.LOGIN, login.getResponseText().trim());
  props.setProperty(PROP.PASSWORD, pwd.getResponseText());
  props.setProperty(PROP.RECIPIENTS, rec.getResponseText().trim() || defaultTo);
  CacheService.getScriptCache().remove('efento_token');

  getAccessToken_(); // proves the credentials work
  ui.alert('Credentials saved and Efento Cloud login succeeded.');
}

/**
 * Edit both recipient lists without re-entering the Efento password:
 *   daily list  — gets the summary every morning, whatever the result (and any failure alert)
 *   leak list   — added to that morning's email only when the night is graded LEAK
 */
function setRecipients() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const daily = askAddressList_(ui, 'Daily email (every morning)', recipients_(), false);
  if (daily === null) return;
  const leak = askAddressList_(ui, 'Leak alerts only (LEAK nights)', props.getProperty(PROP.LEAK_RECIPIENTS) || '(none)', true);
  if (leak === null) return;

  props.setProperty(PROP.RECIPIENTS, daily.join(','));
  props.setProperty(PROP.LEAK_RECIPIENTS, leak.join(','));
  ui.alert('Saved.\n\nEvery morning:\n' + daily.join('\n') +
    '\n\nLEAK nights only:\n' + (leak.length ? leak.join('\n') : '(nobody)'));
}

/** Prompts for a comma-separated address list. Returns the list, or null if cancelled/invalid. */
function askAddressList_(ui, title, current, allowEmpty) {
  const r = ui.prompt(title, 'Currently: ' + current + '\n\nEnter the full list, comma-separated' +
    (allowEmpty ? ' (leave blank for nobody):' : ':'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const list = r.getResponseText().split(/[,;\s]+/).map(x => x.trim()).filter(Boolean);
  const bad = list.filter(x => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
  if (bad.length || (!allowEmpty && !list.length)) {
    ui.alert('Not saved — check these addresses: ' + (bad.join(', ') || '(none entered)'));
    return null;
  }
  return list;
}

function findMeasurementPoint() {
  const id = discoverMeasurementPoint_();
  SpreadsheetApp.getUi().alert('Logger ' + CONFIG.DEVICE_SERIAL + ' found: measurement point ID ' + id + '. Saved.');
}

function installDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('runNightlyCheck')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.RUN_AT_HOUR)
    .nearMinute(CONFIG.RUN_NEAR_MINUTE)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
  SpreadsheetApp.getUi().alert('Daily check installed for about ' + pad2_(CONFIG.RUN_AT_HOUR) + ':' +
    pad2_(CONFIG.RUN_NEAR_MINUTE) + ' UK time (Google runs it within ±15 minutes).');
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runNightlyCheck')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ============================================================================================
// MAIN RUN
// ============================================================================================

/** Trigger entry point. Never fails silently: any error is emailed. */
function runNightlyCheck() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const today = localDateString_(new Date());
    const result = checkNight_(today, true);
    buildDashboard_();
    sendSummaryEmail_(result);
  } catch (err) {
    sendFailureEmail_(err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function runNightlyCheckManual() {
  runNightlyCheck();
  SpreadsheetApp.getUi().alert('Check complete. See the Dashboard and your inbox.');
}

function refreshDashboard() {
  buildDashboard_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(getSheet_(CONFIG.SHEET_DASH));
}

function backfillNights() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Backfill', 'How many previous nights to analyse (1–30)? No emails are sent.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const n = Math.max(1, Math.min(30, parseInt(r.getResponseText(), 10) || 0));

  let date = localDateString_(new Date());
  const done = [];
  for (let i = 0; i < n; i++) {
    const res = checkNight_(date, false);
    done.push(date + ': ' + res.status.code);
    date = shiftDate_(date, -1);
  }
  buildDashboard_();
  ui.alert('Backfill complete:\n' + done.join('\n'));
}

/**
 * Analyses the night ending on the morning of `dateStr` (yyyy-MM-dd, UK date) and the full
 * previous calendar day, writes the log row, and optionally files the Efento email.
 */
function checkNight_(dateStr, fileEmail) {
  const mpId = getMeasurementPointId_();
  const night = nightWindow_(dateStr);
  const prevDay = dayWindow_(shiftDate_(dateStr, -1));

  const fetchFrom = new Date(prevDay.start.getTime() - 2 * 3600e3);
  const fetchTo = new Date(Math.min(Date.now(), night.end.getTime() + 2 * 3600e3));
  const samples = fetchSamples_(mpId, fetchFrom, fetchTo);

  appendRawReadings_(samples);

  const n = analyseWindow_(samples, night.start.getTime(), night.end.getTime(), CONFIG.COUNTER_MODE, CONFIG.MAX_GAP_FACTOR);
  const d = analyseWindow_(samples, prevDay.start.getTime(), prevDay.end.getTime(), CONFIG.COUNTER_MODE, CONFIG.MAX_GAP_FACTOR);

  const status = finalStatus_(n, d);
  const notes = [].concat(n.notes);

  if (d.ok && d.pulses === 0 && CONFIG.FLAG_ZERO_FLOW_DAY) {
    notes.push('Previous day shows zero flow for 24 hours — implausible for nine houses: check the logger, the pulse cable, and that COUNTER_MODE is right.');
  }
  if (d.ok && CONFIG.COUNTER_MODE === 'ACCUMULATING' && d.resets > 1) {
    notes.push('Counter went backwards ' + d.resets + ' times in the previous day — the channel may be in per-interval mode. Check COUNTER_MODE.');
  }
  if (samples.length) {
    const lastAgeH = (Date.now() - samples[samples.length - 1].t) / 3600e3;
    if (lastAgeH > 3) notes.push('Latest reading is ' + lastAgeH.toFixed(1) + ' h old — the logger may have stopped transmitting.');
  }

  const lpp = CONFIG.LITRES_PER_PULSE;
  const nightHours = (night.end - night.start) / 3600e3;
  const result = {
    dateStr: dateStr,
    windowText: fmtLocal_(night.start, 'EEE d MMM HH:mm') + ' – ' + fmtLocal_(night.end, 'HH:mm'),
    status: status,
    night: n,
    day: d,
    nightLitres: n.ok && lpp ? n.pulses * lpp : null,
    nightLph: n.ok && lpp ? (n.pulses * lpp) / nightHours : null,
    dayLitres: d.ok && lpp ? d.pulses * lpp : null,
    notes: notes,
    hourly: hourlyBuckets_(samples, prevDay.start.getTime(), prevDay.end.getTime()),
    prevDayLabel: fmtLocal_(prevDay.start, 'EEE d MMM'),
    efentoThreadUrl: fileEmail ? fileEfentoEmail_() : null,
  };

  writeLogRow_(result);
  const streak = leakStreak_();
  result.leakStreak = streak.count;
  result.leakSince = streak.since;
  return result;
}

// ============================================================================================
// ANALYSIS (pure functions — no Google services)
// ============================================================================================

/**
 * samples: [{t: epoch ms, v: number}] sorted ascending, de-duplicated.
 * Returns pulses in window, longest continuous zero-flow run (minutes) and data-quality info.
 */
function analyseWindow_(samples, startMs, endMs, mode, maxGapFactor) {
  const out = {
    ok: false, pulses: null, longestZeroMin: null, samplesInWindow: 0,
    intervalMin: null, resets: 0, gaps: 0, notes: [],
  };
  if (!samples || samples.length < 2) {
    out.notes.push('Fewer than two readings available.');
    return out;
  }

  const interval = medianSpacingMs_(samples);
  out.intervalMin = Math.round(interval / 60000);
  const maxGap = interval * maxGapFactor;
  out.samplesInWindow = samples.filter(s => s.t > startMs && s.t <= endMs).length;

  // Build the list of intervals [a, b] each with a pulse count.
  const intervals = [];
  if (mode === 'ACCUMULATING') {
    for (let i = 1; i < samples.length; i++) {
      const a = samples[i - 1], b = samples[i];
      let pulses = b.v - a.v;
      let reset = false;
      if (pulses < 0) { reset = true; pulses = b.v; } // counter restarted from zero
      intervals.push({ a: a.t, b: b.t, pulses: pulses, reset: reset });
    }
  } else if (mode === 'PER_INTERVAL') {
    for (let i = 0; i < samples.length; i++) {
      const b = samples[i];
      const a = i > 0 ? Math.max(samples[i - 1].t, b.t - interval) : b.t - interval;
      intervals.push({ a: a, b: b.t, pulses: b.v, reset: false });
    }
  } else {
    throw new Error('COUNTER_MODE must be ACCUMULATING or PER_INTERVAL.');
  }

  // Coverage: the window must be covered from start to end without an unexplained hole.
  const covering = intervals.filter(iv => iv.b > startMs && iv.a < endMs);
  if (!covering.length) {
    out.notes.push('No readings cover the window.');
    return out;
  }
  const firstA = covering[0].a, lastB = covering[covering.length - 1].b;
  if (firstA > startMs + interval) {
    out.notes.push('Readings start ' + Math.round((firstA - startMs) / 60000) + ' min after the window opens.');
    return out;
  }
  if (lastB < endMs - interval) {
    out.notes.push('Readings stop ' + Math.round((endMs - lastB) / 60000) + ' min before the window closes.');
    return out;
  }
  if (mode === 'PER_INTERVAL') {
    for (let i = 1; i < covering.length; i++) {
      if (covering[i].a - covering[i - 1].b > interval * 0.5) {
        out.notes.push('Missing readings inside the window (per-interval data cannot bridge gaps).');
        return out;
      }
    }
  }

  // Pulses and longest zero run. An interval only partly inside the window counts its full
  // pulses (we cannot tell when inside it they occurred) — this errs towards reporting flow.
  let pulses = 0, run = 0, longest = 0;
  covering.forEach(iv => {
    const segStart = Math.max(iv.a, startMs), segEnd = Math.min(iv.b, endMs);
    const segMs = Math.max(0, segEnd - segStart);
    if (iv.reset) out.resets++;
    if (iv.b - iv.a > maxGap) out.gaps++;
    pulses += iv.pulses;
    if (iv.pulses === 0 && !iv.reset) {
      run += segMs;
      if (run > longest) longest = run;
    } else {
      run = 0;
    }
  });

  out.ok = true;
  out.pulses = pulses;
  out.longestZeroMin = Math.round(longest / 60000);
  if (out.resets) out.notes.push('Counter reset detected ' + out.resets + ' time(s); treated as flow.');
  if (out.gaps && mode === 'ACCUMULATING') {
    out.notes.push(out.gaps + ' transmission gap(s) in window; the running total still spans them, so the result stands.');
  }
  return out;
}

function gradeNight_(n) {
  if (!n.ok) return STATUS.NO_DATA;
  if (n.longestZeroMin >= CONFIG.CLEAR_MIN_ZERO_RUN_MINUTES) return STATUS.CLEAR;
  if (n.longestZeroMin >= CONFIG.WATCH_MIN_ZERO_RUN_MINUTES) return STATUS.WATCH;
  return STATUS.LEAK;
}

/**
 * The night's grade, overridden to CHECK when the whole previous day shows no flow at all —
 * on a supply serving nine houses that means a dead logger or cable, not a perfect night.
 */
function finalStatus_(n, d) {
  const s = gradeNight_(n);
  if (CONFIG.FLAG_ZERO_FLOW_DAY && d && d.ok && d.pulses === 0 && (s === STATUS.CLEAR || s === STATUS.WATCH)) {
    return STATUS.CHECK;
  }
  return s;
}

/** Value of an accumulating counter at time t: the last reading at or before t (null if none). */
function valueAt_(samples, t) {
  let lo = 0, hi = samples.length - 1, ans = null;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (samples[mid].t <= t) { ans = samples[mid].v; lo = mid + 1; } else { hi = mid - 1; }
  }
  return ans;
}

/**
 * Pulses in each real hour from startMs to endMs (a DST day gives 23 or 25 buckets).
 * Returns [{t: hour start ms, pulses: number|null}] — null where readings don't cover the hour.
 */
function hourlyBuckets_(samples, startMs, endMs) {
  const out = [];
  for (let t = startMs; t < endMs; t += 3600e3) {
    const a = valueAt_(samples, t), b = valueAt_(samples, Math.min(t + 3600e3, endMs));
    out.push({ t: t, pulses: a == null || b == null ? null : Math.max(0, b - a) });
  }
  return out;
}

function medianSpacingMs_(samples) {
  const d = [];
  for (let i = 1; i < samples.length; i++) d.push(samples[i].t - samples[i - 1].t);
  d.sort((x, y) => x - y);
  const m = d[Math.floor(d.length / 2)];
  return m > 0 ? m : 15 * 60000;
}

/**
 * Pulls {t, v} pairs for the configured channel out of an Efento API response.
 * Tolerant of structure: finds any object carrying a timestamp plus either a
 * `channels` array or a channel number and value.
 */
function extractSamples_(json, channel) {
  const out = [];
  const tsKeys = ['measuredAt', 'measured_at', 'timestamp', 'time', 'date'];

  const walk = node => {
    if (Array.isArray(node)) { node.forEach(walk); return; }
    if (!node || typeof node !== 'object') return;

    const tsKey = tsKeys.find(k => typeof node[k] === 'string' || typeof node[k] === 'number');
    if (tsKey) {
      const t = parseApiTime_(node[tsKey]);
      if (Array.isArray(node.channels) && !isNaN(t)) {
        node.channels.forEach(ch => {
          const num = ch.number != null ? ch.number : (ch.channel != null ? ch.channel : ch.channelNumber);
          if (Number(num) === channel && isOkStatus_(ch.status) && ch.value !== null && ch.value !== '') {
            const v = Number(ch.value);
            if (isFinite(v)) out.push({ t: t, v: v });
          }
        });
        return;
      }
      const num = node.channel != null ? node.channel : (node.channelNumber != null ? node.channelNumber : node.number);
      if (num != null && node.value != null && node.value !== '' && !isNaN(t) && Number(num) === channel && isOkStatus_(node.status)) {
        const v = Number(node.value);
        if (isFinite(v)) out.push({ t: t, v: v });
        return;
      }
    }
    Object.keys(node).forEach(k => {
      if (node[k] && typeof node[k] === 'object') walk(node[k]);
    });
  };
  walk(json);
  return dedupeSort_(out);
}

/**
 * Efento reports pulse channels as "OK_CALIBRATION_REQUIRED" (no meter factor set in Efento Cloud);
 * the count itself is valid. Accept any status beginning "OK"; reject errors.
 */
function isOkStatus_(s) {
  return s == null || /^OK/.test(String(s).toUpperCase());
}

function dedupeSort_(arr) {
  const map = {};
  arr.forEach(s => { map[s.t] = s; });
  return Object.keys(map).map(k => map[k]).sort((a, b) => a.t - b.t);
}

function parseApiTime_(s) {
  if (typeof s === 'number') return s < 1e12 ? s * 1000 : s;
  const str = String(s).trim();
  if (/[zZ]$|[+-]\d\d:?\d\d$/.test(str)) return new Date(str.replace(' ', 'T')).getTime();
  const m = str.match(/^(\d{4})-(\d\d)-(\d\d)[ T](\d\d):(\d\d)(?::(\d\d))?/);
  if (!m) return NaN;
  const parts = [+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)];
  if (CONFIG.API_TIMES_ARE_UTC) return Date.UTC.apply(null, parts);
  // Not UTC: interpret as UK local time
  return Utilities.parseDate(str.substring(0, 19).replace('T', ' '), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss').getTime();
}

// ============================================================================================
// EFENTO CLOUD API
// ============================================================================================

function getAccessToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('efento_token');
  if (cached) return cached;

  const props = PropertiesService.getScriptProperties();
  const login = props.getProperty(PROP.LOGIN);
  const password = props.getProperty(PROP.PASSWORD);
  if (!login || !password) throw new Error('Efento credentials not set. Use Water Monitor > Set up credentials.');

  const res = UrlFetchApp.fetch(CONFIG.API_BASE + '/login', {
    method: 'post',
    headers: { 'X-USER-LOGIN': login, 'X-USER-PASSWORD': password },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Efento login failed (HTTP ' + code + '): ' + res.getContentText().substring(0, 300));

  const token = JSON.parse(res.getContentText()).accessToken;
  if (!token) throw new Error('Efento login response contained no accessToken.');
  cache.put('efento_token', token, 45 * 60);
  return token;
}

function apiGet_(path) {
  const url = CONFIG.API_BASE + path;
  for (let attempt = 1; attempt <= 3; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + getAccessToken_() },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) return JSON.parse(res.getContentText());
    if (code === 401) { CacheService.getScriptCache().remove('efento_token'); }
    if ((code === 401 || code === 429 || code >= 500) && attempt < 3) { Utilities.sleep(2000 * attempt); continue; }
    throw new Error('Efento API ' + path + ' failed (HTTP ' + code + '): ' + res.getContentText().substring(0, 300));
  }
  throw new Error('Efento API ' + path + ' failed after retries.');
}

function discoverMeasurementPoint_() {
  const orgs = apiGet_('/organizations').organizations || [];
  for (const org of orgs) {
    const locs = apiGet_('/locations?organization-id=' + org.id).locations || [];
    for (const loc of locs) {
      for (let offset = 0; offset < 1000; offset += 100) {
        const page = apiGet_('/measurement-points?location-ids=' + loc.id + '&limit=100&offset=' + offset);
        const mps = page.measurementPoints || [];
        const hit = mps.find(mp => mp.device && String(mp.device.serialNumber).toUpperCase() === CONFIG.DEVICE_SERIAL.toUpperCase());
        if (hit) {
          PropertiesService.getScriptProperties().setProperty(PROP.MP_ID, String(hit.id));
          return hit.id;
        }
        if (mps.length < 100) break;
      }
    }
  }
  throw new Error('Logger ' + CONFIG.DEVICE_SERIAL + ' not found in any Efento Cloud organisation/location visible to this login.');
}

function getMeasurementPointId_() {
  return PropertiesService.getScriptProperties().getProperty(PROP.MP_ID) || discoverMeasurementPoint_();
}

function fetchSamples_(mpId, from, to) {
  let all = [];
  const chunk = CONFIG.FETCH_CHUNK_HOURS * 3600e3;
  for (let s = from.getTime(); s < to.getTime(); s += chunk) {
    const e = Math.min(s + chunk - 1000, to.getTime());
    const json = apiGet_('/measurement-points/' + mpId + '/measurements?from=' +
      encodeURIComponent(fmtUtc_(new Date(s))) + '&to=' + encodeURIComponent(fmtUtc_(new Date(e))));
    all = all.concat(extractSamples_(json, CONFIG.CHANNEL));
  }
  return dedupeSort_(all);
}

/** Writes raw API output for the last 3 hours to the Diagnostics sheet, plus what was parsed. */
function diagnoseApi() {
  const mpId = getMeasurementPointId_();
  const to = new Date(), from = new Date(to.getTime() - 3 * 3600e3);
  const path = '/measurement-points/' + mpId + '/measurements?from=' +
    encodeURIComponent(fmtUtc_(from)) + '&to=' + encodeURIComponent(fmtUtc_(to));
  const json = apiGet_(path);
  const samples = extractSamples_(json, CONFIG.CHANNEL);

  const sh = getSheet_(CONFIG.SHEET_DIAG);
  sh.clear();
  sh.getRange(1, 1, 4, 2).setValues([
    ['Request', CONFIG.API_BASE + path],
    ['Run at (UK)', fmtLocal_(new Date(), 'yyyy-MM-dd HH:mm:ss')],
    ['Channel ' + CONFIG.CHANNEL + ' readings parsed', samples.length],
    ['Raw JSON (first 45,000 characters)', JSON.stringify(json, null, 1).substring(0, 45000)],
  ]);
  if (samples.length) {
    sh.getRange(6, 1, 1, 2).setValues([['Parsed reading (UK time)', 'Value']]).setFontWeight('bold');
    sh.getRange(7, 1, samples.length, 2).setValues(samples.map(s => [fmtLocal_(new Date(s.t), 'yyyy-MM-dd HH:mm:ss'), s.v]));
  }
  sh.setColumnWidth(2, 700);
  SpreadsheetApp.getUi().alert(samples.length
    ? 'Parsed ' + samples.length + ' readings for channel ' + CONFIG.CHANNEL + '. Compare the times and values on the Diagnostics sheet with Efento Cloud before relying on the results.'
    : 'No channel ' + CONFIG.CHANNEL + ' readings parsed. The raw response is on the Diagnostics sheet — send it to Claude to adjust the parser.');
}

// ============================================================================================
// SHEETS
// ============================================================================================

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground(THEME.NAVY).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontFamily(THEME.FONT).setFontSize(10).setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 40);
}

function writeLogRow_(r) {
  const sh = getSheet_(CONFIG.SHEET_LOG);
  ensureHeaders_(sh, LOG_HEADERS);

  const row = [
    r.dateStr, r.windowText, r.status.code,
    r.night.ok ? r.night.longestZeroMin : '',
    r.night.ok ? r.night.pulses : '',
    r.nightLitres != null ? round1_(r.nightLitres) : '',
    r.nightLph != null ? round1_(r.nightLph) : '',
    r.day.ok ? r.day.pulses : '',
    r.dayLitres != null ? round1_(r.dayLitres) : '',
    r.night.samplesInWindow,
    r.night.intervalMin != null ? r.night.intervalMin : '',
    r.notes.join(' '),
    fmtLocal_(new Date(), 'yyyy-MM-dd HH:mm'),
  ];

  const last = sh.getLastRow();
  let target = last + 1;
  if (last > 1) {
    const dates = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(x => x[0]);
    const idx = dates.indexOf(r.dateStr);
    if (idx >= 0) target = idx + 2;
  }
  sh.getRange(target, 1, 1, 1).setNumberFormat('@');
  sh.getRange(target, 1, 1, row.length).setValues([row]);

  if (sh.getLastRow() > 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEADERS.length).sort({ column: 1, ascending: false });
  }
  formatLog_(sh);
}

/** Column widths, banding, alignment and colour-coded status chips for the Nightly Log. */
function formatLog_(sh) {
  const widths = [110, 175, 90, 120, 90, 90, 110, 120, 120, 110, 110, 380, 130];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const rows = Math.max(sh.getLastRow() - 1, 1);
  const body = sh.getRange(2, 1, rows, LOG_HEADERS.length);
  body.setFontFamily(THEME.FONT).setFontSize(10).setFontColor(THEME.INK)
    .setVerticalAlignment('middle').setBackground(null);
  sh.getRange(2, 1, rows, 11).setHorizontalAlignment('center');
  sh.getRange(2, 12, rows, 1).setWrap(true).setHorizontalAlignment('left');
  sh.getRange(2, 13, rows, 1).setHorizontalAlignment('center').setFontColor(THEME.MUTED);
  sh.getRange(2, 3, rows, 1).setFontWeight('bold');

  sh.getBandings().forEach(b => b.remove());
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), LOG_HEADERS.length)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(THEME.NAVY).setFirstRowColor('#FFFFFF').setSecondRowColor(THEME.BAND);

  const statusCol = sh.getRange('C2:C');
  sh.setConditionalFormatRules(Object.keys(STATUS_CHIP).map(code =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(code)
      .setBackground(STATUS_CHIP[code][0]).setFontColor(STATUS_CHIP[code][1])
      .setRanges([statusCol]).build()));
}

function appendRawReadings_(samples) {
  if (!samples.length) return;
  const sh = getSheet_(CONFIG.SHEET_RAW);
  ensureHeaders_(sh, RAW_HEADERS);

  const last = sh.getLastRow();
  const existing = new Set();
  if (last > 1) {
    const n = Math.min(last - 1, 5000);
    sh.getRange(last - n + 1, 1, n, 1).getDisplayValues().forEach(x => existing.add(x[0]));
  }
  // At one reading a minute a full record would be ~1,440 rows a day. Keep every reading where
  // the value changed (i.e. every pulse) plus one on each hour, which preserves the full story.
  const kept = samples.filter((s, i) =>
    i === 0 || s.v !== samples[i - 1].v || fmtUtc_(new Date(s.t)).substring(14, 16) === '00');
  const rows = kept
    .map(s => [fmtUtc_(new Date(s.t)), fmtLocal_(new Date(s.t), 'yyyy-MM-dd HH:mm:ss'), s.v])
    .filter(r => !existing.has(r[0]));
  if (!rows.length) return;

  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, 2).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, 3).setValues(rows);
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, 3).sort({ column: 1, ascending: true });
  [165, 165, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).setFontFamily(THEME.FONT).setFontSize(10).setHorizontalAlignment('center');
}

/** Consecutive most-recent logged nights graded LEAK, and the first night of that run. */
function leakStreak_() {
  const sh = getSheet_(CONFIG.SHEET_LOG);
  const last = sh.getLastRow();
  if (last < 2) return { count: 0, since: null };
  const vals = sh.getRange(2, 1, Math.min(last - 1, CONFIG.STREAK_LOOKBACK), 3).getDisplayValues();
  let count = 0, since = null;
  for (const v of vals) {
    if (v[2] === STATUS.LEAK.code) { count++; since = v[0]; } else break;
  }
  return { count: count, since: since };
}

// ============================================================================================
// DASHBOARD
// ============================================================================================

/** Rebuilds the Dashboard tab entirely from the Nightly Log and Raw Readings sheets (no API calls). */
function buildDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = getSheet_(CONFIG.SHEET_LOG);
  if (log.getLastRow() < 2) return;

  const rows = log.getRange(2, 1, Math.min(log.getLastRow() - 1, CONFIG.DASH_NIGHTS), LOG_HEADERS.length).getDisplayValues();
  const latest = rows[0];
  const dateStr = latest[0];
  const status = Object.keys(STATUS).map(k => STATUS[k]).find(x => x.code === latest[2]) || STATUS.NO_DATA;
  const win = nightWindow_(dateStr);
  const nightHours = (win.end - win.start) / 3600e3;
  const streak = leakStreak_();
  const lpp = CONFIG.LITRES_PER_PULSE;

  // ---- Readings for yesterday's hourly profile
  const raw = getSheet_(CONFIG.SHEET_RAW);
  const samples = raw.getLastRow() > 1
    ? raw.getRange(2, 1, raw.getLastRow() - 1, 3).getDisplayValues()
        .map(x => ({ t: parseApiTime_(x[0]), v: Number(x[2]) }))
        .filter(x => !isNaN(x.t) && isFinite(x.v))
    : [];
  const prevDay = dayWindow_(shiftDate_(dateStr, -1));
  const hourly = hourlyBuckets_(dedupeSort_(samples), prevDay.start.getTime(), prevDay.end.getTime());

  // ---- Chart data (hidden sheet)
  const cd = getSheet_(CONFIG.SHEET_CHART);
  cd.clear();
  const hourRows = hourly.map(h => {
    const hh = fmtLocal_(new Date(h.t), 'HH:mm');
    const night = isNightHour_(hh);
    const v = h.pulses == null ? '' : h.pulses;
    return [hh.substring(0, 2), night ? v : '', night ? '' : v];
  });
  cd.getRange(1, 1, 1, 3).setValues([['Hour', 'Night window', 'Daytime']]);
  cd.getRange(2, 1, hourRows.length, 1).setNumberFormat('@');
  cd.getRange(2, 1, hourRows.length, 3).setValues(hourRows);

  const nights = rows.slice().reverse().map(x => [
    fmtLocal_(Utilities.parseDate(x[0] + ' 12:00', CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'), 'EEE d MMM'),
    x[3] === '' ? '' : Number(x[3]),
    CONFIG.CLEAR_MIN_ZERO_RUN_MINUTES,
    x[4] === '' ? '' : Number(x[4]),
  ]);
  cd.getRange(1, 5, 1, 4).setValues([['Night', 'Longest zero-flow run (min)', 'CLEAR threshold', 'Night pulses']]);
  cd.getRange(2, 5, nights.length, 1).setNumberFormat('@');
  cd.getRange(2, 5, nights.length, 4).setValues(nights);
  // Charts drawn by a script ignore hidden rows by default, so this tab stays visible (at the end).
  if (cd.isSheetHidden()) cd.showSheet();
  cd.setTabColor('#B0BEC5');

  // ---- Dashboard canvas
  const dash = getSheet_(CONFIG.SHEET_DASH);
  dash.getCharts().forEach(c => dash.removeChart(c));
  dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns()).breakApart();
  dash.clear();
  dash.setHiddenGridlines(true);
  dash.setColumnWidth(1, 24);
  for (let c = 2; c <= 9; c++) dash.setColumnWidth(c, 114);
  dash.setColumnWidth(10, 24);
  dash.getRange(1, 1, 60, 10).setFontFamily(THEME.FONT).setFontColor(THEME.INK).setVerticalAlignment('middle');
  dash.setRowHeight(1, 14);

  // Title
  dash.getRange('B2:I2').merge().setValue('Tytheland Water Monitor').setFontSize(22).setFontWeight('bold').setFontColor(THEME.NAVY);
  dash.setRowHeight(2, 40);
  dash.getRange('B3:I3').merge()
    .setValue('R S Hill master meter  ·  Efento ' + CONFIG.DEVICE_SERIAL + ' channel ' + CONFIG.CHANNEL +
      '  ·  updated ' + (latest[12] || fmtLocal_(new Date(), 'yyyy-MM-dd HH:mm')))
    .setFontSize(10).setFontColor(THEME.MUTED);
  dash.setRowHeight(4, 10);

  // Status banner
  dash.getRange('B5:I6').merge()
    .setValue(status.code + '   ·   ' + status.text)
    .setBackground(status.colour).setFontColor('#FFFFFF').setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center');
  dash.setRowHeight(5, 30); dash.setRowHeight(6, 30);
  dash.getRange('B7:I7').merge()
    .setValue('Night ' + latest[1] + (latest[11] ? '   —   ' + latest[11] : ''))
    .setFontSize(10).setFontColor(THEME.MUTED).setHorizontalAlignment('center').setWrap(true);
  dash.setRowHeight(7, 30);
  dash.setRowHeight(8, 12);

  // Headline tiles
  const nightPulses = latest[4] === '' ? null : Number(latest[4]);
  const dayPulses = latest[7] === '' ? null : Number(latest[7]);
  const share = nightPulses != null && dayPulses ? Math.min(100, Math.round((nightPulses / nightHours) * 24 / dayPulses * 100)) : null;
  const tiles = [
    [latest[3] === '' ? '—' : latest[3] + ' min', 'Longest zero-flow run', 'CLEAR needs ' + CONFIG.CLEAR_MIN_ZERO_RUN_MINUTES + ' min', status.colour],
    [nightPulses == null ? '—' : String(nightPulses), 'Night pulses',
      CONFIG.NIGHT_START + '–' + CONFIG.NIGHT_END + (nightPulses == null ? '' : (lpp ? '  ·  ' + round1_(nightPulses * lpp) + ' L' : '  ·  ≈ ' + round1_(nightPulses / nightHours) + ' per hour')), THEME.INK],
    [dayPulses == null ? '—' : String(dayPulses), 'Previous-day pulses',
      share == null ? '' : 'est. ' + share + '% steady background flow', THEME.INK],
    [String(streak.count), 'LEAK nights running',
      streak.since ? 'since ' + fmtLocal_(Utilities.parseDate(streak.since + ' 12:00', CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'), 'EEE d MMM') : '—',
      streak.count ? STATUS.LEAK.colour : THEME.INK],
  ];
  tiles.forEach((t, i) => {
    const col = 2 + i * 2;
    const block = dash.getRange(9, col, 3, 2);
    block.setBackground(THEME.TILE).setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID_THICK);
    dash.getRange(9, col, 1, 2).merge().setValue(t[0]).setFontSize(26).setFontWeight('bold').setFontColor(t[3]);
    dash.getRange(10, col, 1, 2).merge().setValue(t[1]).setFontSize(10).setFontWeight('bold');
    dash.getRange(11, col, 1, 2).merge().setValue(t[2]).setFontSize(9).setFontColor(THEME.MUTED);
  });
  dash.setRowHeight(9, 48); dash.setRowHeight(10, 20); dash.setRowHeight(11, 22);
  dash.setRowHeight(12, 16);

  // Section headings
  dash.getRange('B13:I13').merge().setValue('Hour by hour  ·  ' + fmtLocal_(prevDay.start, 'EEEE d MMMM'))
    .setFontSize(12).setFontWeight('bold').setFontColor(THEME.NAVY);
  dash.getRange('B30:I30').merge().setValue('Night by night  ·  last ' + nights.length + (nights.length === 1 ? ' night' : ' nights'))
    .setFontSize(12).setFontWeight('bold').setFontColor(THEME.NAVY);

  const base = (b) => b.setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
    .setOption('fontName', THEME.FONT).setOption('backgroundColor', '#FFFFFF')
    .setOption('titleTextStyle', { color: THEME.INK, fontSize: 12, bold: true })
    .setOption('legend', { position: 'bottom', textStyle: { color: THEME.MUTED, fontSize: 10 } })
    .setOption('hAxis', { textStyle: { color: THEME.MUTED, fontSize: 10 }, gridlines: { color: 'transparent' } })
    .setOption('vAxis', { textStyle: { color: THEME.MUTED, fontSize: 10 }, gridlines: { color: THEME.RULE }, minValue: 0 });

  // Hourly profile
  dash.insertChart(base(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(cd.getRange(1, 1, hourRows.length + 1, 3))
    .setNumHeaders(1)
    .setOption('isStacked', true)
    .setOption('colors', [THEME.NAVY, THEME.TEAL_LIGHT])
    .setOption('bar', { groupWidth: '80%' })
    .setOption('chartArea', { left: 50, top: 20, width: '90%', height: '70%' })
    .setOption('width', 912).setOption('height', 300)
    .setPosition(14, 2, 0, 4)).build());

  // Longest zero-flow run vs threshold
  dash.insertChart(base(dash.newChart().setChartType(Charts.ChartType.COMBO)
    .addRange(cd.getRange(1, 5, nights.length + 1, 3))
    .setNumHeaders(1)
    .setOption('title', 'Longest zero-flow run (min)')
    .setOption('seriesType', 'bars')
    .setOption('series', { 0: { color: THEME.TEAL }, 1: { type: 'line', color: STATUS.LEAK.colour, lineDashStyle: [6, 4], lineWidth: 2 } })
    .setOption('chartArea', { left: 45, top: 36, width: '82%', height: '62%' })
    .setOption('width', 452).setOption('height', 300)
    .setPosition(31, 2, 0, 4)).build());

  // Night pulses
  dash.insertChart(base(dash.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(cd.getRange(1, 5, nights.length + 1, 1))
    .addRange(cd.getRange(1, 8, nights.length + 1, 1))
    .setNumHeaders(1)
    .setOption('title', 'Pulses in night window')
    .setOption('colors', [THEME.NAVY])
    .setOption('legend', { position: 'none' })
    .setOption('chartArea', { left: 45, top: 36, width: '82%', height: '62%' })
    .setOption('width', 452).setOption('height', 300)
    .setPosition(31, 6, 4, 4)).build());

  // Footer
  dash.getRange('B47:I47').merge()
    .setValue((lpp ? 'Pulse weight: ' + lpp + ' L per pulse.' : 'Pulse weight not yet set — figures are pulses, not litres.') +
      '  Grades: CLEAR ≥ ' + CONFIG.CLEAR_MIN_ZERO_RUN_MINUTES + ' min of zero flow in the night window; WATCH ≥ ' +
      CONFIG.WATCH_MIN_ZERO_RUN_MINUTES + ' min; LEAK below that.')
    .setFontSize(9).setFontColor(THEME.MUTED).setWrap(true);
  dash.setRowHeight(47, 32);

  // Tab order: Dashboard first, Diagnostics last, remove the empty default sheet
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  dash.setTabColor(THEME.NAVY);
  log.setTabColor(THEME.TEAL);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && blank.getCharts().length === 0) ss.deleteSheet(blank);
  [CONFIG.SHEET_CHART, CONFIG.SHEET_DIAG].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(ss.getNumSheets()); }
  });
  ss.setActiveSheet(dash);
}

// ============================================================================================
// GMAIL
// ============================================================================================

/** Labels (and optionally marks read) recent Efento report emails; returns newest thread URL. */
function fileEfentoEmail_() {
  const threads = GmailApp.search(CONFIG.EFENTO_EMAIL_QUERY, 0, 10);
  if (!threads.length) return null;
  const label = GmailApp.getUserLabelByName(CONFIG.EFENTO_EMAIL_LABEL) || GmailApp.createLabel(CONFIG.EFENTO_EMAIL_LABEL);
  threads.forEach(t => {
    t.addLabel(label);
    if (CONFIG.MARK_EFENTO_EMAIL_READ) t.markRead();
  });
  return threads[0].getPermalink();
}

function recipients_() {
  return PropertiesService.getScriptProperties().getProperty(PROP.RECIPIENTS) || Session.getEffectiveUser().getEmail();
}

function sendSummaryEmail_(r) {
  const s = r.status;
  const lpp = CONFIG.LITRES_PER_PULSE;
  const nightEnd = nightWindow_(r.dateStr).end;
  const nightHours = (nightEnd - nightWindow_(r.dateStr).start) / 3600e3;
  const subject = 'Water monitor ' + fmtLocal_(nightEnd, 'EEE d MMM') + ': ' + s.code +
    (r.leakStreak > 1 && s === STATUS.LEAK ? ' (' + r.leakStreak + ' nights running)' : '');

  const kpis = [
    { label: 'Longest zero-flow run', value: r.night.ok ? r.night.longestZeroMin + ' min' : '—',
      note: 'CLEAR needs ' + CONFIG.CLEAR_MIN_ZERO_RUN_MINUTES, colour: s.colour },
    { label: 'Night pulses', value: r.night.ok ? String(r.night.pulses) : '—',
      note: CONFIG.NIGHT_START + '–' + CONFIG.NIGHT_END + (r.nightLitres != null ? ' · ' + round1_(r.nightLitres) + ' L'
        : (r.night.ok ? ' · ≈' + round1_(r.night.pulses / nightHours) + '/h' : '')) },
    { label: 'Previous day', value: r.day.ok ? String(r.day.pulses) : '—',
      note: r.dayLitres != null ? round1_(r.dayLitres) + ' L' : 'pulses' },
    { label: 'LEAK nights running', value: String(r.leakStreak || 0),
      note: r.leakSince ? 'since ' + fmtLocal_(Utilities.parseDate(r.leakSince + ' 12:00', CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'), 'EEE d MMM') : '' },
  ];

  const hourly = (r.hourly || []).map(h => ({
    label: fmtLocal_(new Date(h.t), 'HH'),
    pulses: h.pulses,
    night: isNightHour_(fmtLocal_(new Date(h.t), 'HH:mm')),
  }));

  const details = [['Night window (UK)', escapeHtml_(r.windowText)],
    ['Readings in window', r.night.samplesInWindow + (r.night.intervalMin ? ' at ' + r.night.intervalMin + '-min intervals' : '')]];
  if (!lpp) details.push(['Volume', 'Pulse weight not yet set — results are in pulses']);

  const links = [['Open the Dashboard', SpreadsheetApp.getActiveSpreadsheet().getUrl()]];
  if (r.efentoThreadUrl) links.push(['This morning\'s Efento PDF', r.efentoThreadUrl]);

  const html = buildEmailHtml_({
    heading: 'Night ending ' + fmtLocal_(nightEnd, 'EEEE d MMMM'),
    status: s,
    kpis: kpis,
    hourlyTitle: r.prevDayLabel ? 'Pulses per hour · ' + r.prevDayLabel : '',
    hourly: hourly,
    details: details,
    notes: r.notes.map(escapeHtml_),
    links: links,
  });

  const text = [s.code + ' — ' + s.text, r.windowText]
    .concat(kpis.map(k => k.label + ': ' + k.value + (k.note ? ' (' + k.note + ')' : '')))
    .concat(r.notes).concat(links.map(l => l[0] + ': ' + l[1])).join('\n');
  MailApp.sendEmail({ to: summaryRecipients_(s), subject: subject, body: text, htmlBody: html });
}

/** Daily list always; leak-alert list added only on LEAK nights. Duplicates removed. */
function summaryRecipients_(status) {
  const daily = recipients_().split(',');
  const leak = status === STATUS.LEAK
    ? (PropertiesService.getScriptProperties().getProperty(PROP.LEAK_RECIPIENTS) || '').split(',')
    : [];
  const seen = {};
  return daily.concat(leak).map(x => x.trim()).filter(x => x && !seen[x.toLowerCase()] && (seen[x.toLowerCase()] = true)).join(',');
}

/** True when a local HH:mm falls inside the night window. */
function isNightHour_(hhmm) {
  const a = CONFIG.NIGHT_START, b = CONFIG.NIGHT_END;
  return a < b ? (hhmm >= a && hhmm < b) : (hhmm >= a || hhmm < b);
}

/** Pure HTML builder for the daily email (inline styles and tables, so Gmail renders it). */
function buildEmailHtml_(m) {
  const F = 'font-family:Roboto,Helvetica,Arial,sans-serif;';
  const tiles = m.kpis.map(k =>
    '<td width="25%" style="padding:6px;" valign="top">' +
      '<div style="background:' + THEME.TILE + ';border-radius:10px;padding:14px 10px;text-align:center;height:92px;">' +
        '<div style="' + F + 'font-size:26px;font-weight:700;color:' + (k.colour || THEME.INK) + ';">' + k.value + '</div>' +
        '<div style="' + F + 'font-size:12px;font-weight:600;color:' + THEME.INK + ';margin-top:4px;">' + k.label + '</div>' +
        '<div style="' + F + 'font-size:11px;color:' + THEME.MUTED + ';margin-top:2px;">' + (k.note || '&nbsp;') + '</div>' +
      '</div></td>').join('');

  let chart = '';
  if (m.hourly && m.hourly.length) {
    const max = Math.max(1, ...m.hourly.map(h => h.pulses || 0));
    const bars = m.hourly.map(h => {
      const ht = h.pulses == null ? 0 : Math.max(2, Math.round((h.pulses / max) * 70));
      const col = h.pulses == null ? THEME.RULE : (h.night ? THEME.NAVY : THEME.TEAL_LIGHT);
      return '<td valign="bottom" align="center" style="padding:0 1px;">' +
        '<div style="' + F + 'font-size:9px;color:' + THEME.MUTED + ';">' + (h.pulses == null ? '' : h.pulses) + '</div>' +
        '<div style="height:' + ht + 'px;background:' + col + ';border-radius:2px 2px 0 0;"></div></td>';
    }).join('');
    const labels = m.hourly.map(h => '<td align="center" style="' + F + 'font-size:9px;color:' + THEME.MUTED + ';padding-top:3px;">' + h.label + '</td>').join('');
    chart = '<tr><td style="padding:8px 24px 4px;">' +
      '<div style="' + F + 'font-size:13px;font-weight:600;color:' + THEME.INK + ';margin-bottom:8px;">' + m.hourlyTitle + '</div>' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="table-layout:fixed;"><tr style="height:90px;">' + bars + '</tr><tr>' + labels + '</tr></table>' +
      '<div style="' + F + 'font-size:11px;color:' + THEME.MUTED + ';margin-top:6px;">' +
        '<span style="display:inline-block;width:10px;height:10px;background:' + THEME.NAVY + ';border-radius:2px;"></span> night window &nbsp; ' +
        '<span style="display:inline-block;width:10px;height:10px;background:' + THEME.TEAL_LIGHT + ';border-radius:2px;"></span> daytime</div>' +
      '</td></tr>';
  }

  const details = m.details.map(d =>
    '<tr><td style="' + F + 'font-size:12px;color:' + THEME.MUTED + ';padding:6px 0;border-bottom:1px solid ' + THEME.RULE + ';white-space:nowrap;">' + d[0] + '</td>' +
    '<td style="' + F + 'font-size:12px;color:' + THEME.INK + ';padding:6px 0 6px 16px;border-bottom:1px solid ' + THEME.RULE + ';">' + d[1] + '</td></tr>').join('');

  const notes = m.notes && m.notes.length
    ? '<tr><td style="padding:4px 24px 8px;"><div style="background:#FEF5E7;border-left:4px solid #D68910;border-radius:6px;padding:10px 12px;' + F + 'font-size:12px;color:' + THEME.INK + ';">' + m.notes.join('<br>') + '</div></td></tr>'
    : '';

  const links = m.links.map(l => '<a href="' + l[1] + '" style="display:inline-block;margin:4px 6px;padding:10px 16px;background:' +
    THEME.TEAL + ';color:#FFFFFF;text-decoration:none;border-radius:6px;' + F + 'font-size:13px;font-weight:600;">' + l[0] + '</a>').join('');

  return '<div style="background:#EEF2F5;padding:24px 8px;">' +
    '<table width="600" align="center" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#FFFFFF;border-radius:12px;overflow:hidden;">' +
      '<tr><td style="background:' + THEME.NAVY + ';padding:18px 24px;">' +
        '<div style="' + F + 'font-size:11px;letter-spacing:2px;color:' + THEME.TEAL_LIGHT + ';font-weight:600;">TYTHELAND WATER MONITOR</div>' +
        '<div style="' + F + 'font-size:20px;font-weight:700;color:#FFFFFF;margin-top:4px;">' + m.heading + '</div></td></tr>' +
      '<tr><td style="background:' + m.status.colour + ';padding:14px 24px;">' +
        '<span style="' + F + 'font-size:18px;font-weight:700;color:#FFFFFF;letter-spacing:1px;">' + m.status.code + '</span>' +
        '<span style="' + F + 'font-size:14px;color:#FFFFFF;"> &nbsp;·&nbsp; ' + m.status.text + '</span></td></tr>' +
      '<tr><td style="padding:14px 18px 6px;"><table width="100%" cellpadding="0" cellspacing="0"><tr>' + tiles + '</tr></table></td></tr>' +
      chart + notes +
      '<tr><td style="padding:8px 24px;"><table width="100%" cellpadding="0" cellspacing="0">' + details + '</table></td></tr>' +
      '<tr><td align="center" style="padding:12px 18px 22px;">' + links + '</td></tr>' +
    '</table>' +
    '<div style="' + F + 'font-size:11px;color:' + THEME.MUTED + ';text-align:center;margin-top:10px;">R S Hill master meter · Efento ' + CONFIG.DEVICE_SERIAL + ' channel ' + CONFIG.CHANNEL + '</div>' +
  '</div>';
}

function sendFailureEmail_(err) {
  try {
    MailApp.sendEmail({
      to: recipients_(),
      subject: 'Water monitor FAILED — ' + fmtLocal_(new Date(), 'EEE d MMM HH:mm'),
      body: 'The nightly water check did not complete, so last night has NOT been assessed.\n\n' +
        'Error: ' + (err && err.message ? err.message : err) + '\n\n' +
        'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    });
  } catch (e) {
    console.error('Could not send failure email: ' + e);
  }
}

// ============================================================================================
// DATE / FORMAT HELPERS
// ============================================================================================

/** Night window ending on the morning of dateStr (UK). Real instants, so DST change nights are exact. */
function nightWindow_(dateStr) {
  const crossesMidnight = CONFIG.NIGHT_START > CONFIG.NIGHT_END;
  const startDate = crossesMidnight ? shiftDate_(dateStr, -1) : dateStr;
  const start = Utilities.parseDate(startDate + ' ' + CONFIG.NIGHT_START, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  const end = Utilities.parseDate(dateStr + ' ' + CONFIG.NIGHT_END, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  if (end <= start) throw new Error('Night window is invalid — check NIGHT_START and NIGHT_END.');
  return { start: start, end: end };
}

function dayWindow_(dateStr) {
  return {
    start: Utilities.parseDate(dateStr + ' 00:00', CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'),
    end: Utilities.parseDate(shiftDate_(dateStr, 1) + ' 00:00', CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'),
  };
}

/** Adds whole days to a yyyy-MM-dd string (anchored at noon so DST can never shift the date). */
function shiftDate_(dateStr, days) {
  const noon = Utilities.parseDate(dateStr + ' 12:00', CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  return Utilities.formatDate(new Date(noon.getTime() + days * 86400e3), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function localDateString_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function fmtLocal_(d, f) { return Utilities.formatDate(d, CONFIG.TIMEZONE, f); }
function fmtUtc_(d) { return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd HH:mm:ss'); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function round1_(x) { return Math.round(x * 10) / 10; }
function escapeHtml_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
