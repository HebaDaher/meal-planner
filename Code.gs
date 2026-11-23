/**
 * Meal Planner & Ingredient Checker → Google Sheet   (version 3)
 * Paste this whole file into Extensions → Apps Script of your Google Sheet and click Save.
 * Then Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy  (your link stays the same).
 *
 * Tabs it keeps up to date by itself:
 *  - "Visitors": ONE row per person, updated while they use the website.
 *  - "Scans":    ONE row per product checked (real barcode, label photo, typed, or sample) — used for Goals 2 and 3.
 *  - "Clicks":   every tap each person made.
 *  - "Usability goals": all 7 goals with the task, how it is measured, the target and the live result.
 */
const VISITORS_TAB = 'Visitors', SCANS_TAB = 'Scans', CLICKS_TAB = 'Clicks', GOALS_TAB = 'Usability goals';
const VERSION = 'v3.1';
const HEADERS = [
  'Date', 'Start time', 'Session ID', 'Name', 'Problem', 'Device', 'Status',
  'Total minutes in website', 'Sign-up (min)',
  'Create profile (min) – Goal 1', 'Profile saved',
  'Products checked', 'Fastest scan → result (sec) – Goal 2', 'Slowest scan → result (sec)',
  'False Safe results – Goal 3',
  'Failed scans', 'Slowest recovery (sec) – Goal 4', 'Undo used',
  'Saved to Pantry', 'Plan a meal (min)', 'Sugar swap used', 'Meals planned',
  'Core features completed (of 3) – Goal 5',
  'Issues reported', 'Total taps', 'Empty-area taps', 'Back presses', 'Screens visited', 'Last screen', 'Last update'
];
// You fill these in yourself. The website never overwrites them.
const MANUAL = ['Needed help? (Yes/No) – Goals 1 & 5', 'Goal 6 – average rating 1–5 (from your survey)', 'Goal 7 – would use weekly? (Yes/No)'];
const SCAN_HEADERS = ['Session ID', 'Name', 'Time from start (m:ss)', 'Profile checked', 'Product', 'Barcode', 'How it was checked',
  'Ingredients (as read)', 'Result shown', 'Why', 'Reference allergens (product database)', 'Automatic check – Goal 3', 'Seconds to result – Goal 2'];
const SCAN_MANUAL = ['Your check against the pack (Correct/Wrong)'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) One row per visitor (only the website's columns are written, so your own columns stay)
    const sh = tab_(ss, VISITORS_TAB, HEADERS, MANUAL);
    const row = HEADERS.map(h => (data.row && data.row[h] !== undefined && data.row[h] !== null) ? data.row[h] : '');
    const idCol = HEADERS.indexOf('Session ID') + 1;
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, idCol, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const found = ids.indexOf(String(data.row['Session ID']));
    sh.getRange(found >= 0 ? found + 2 : last + 1, 1, 1, row.length).setValues([row]);

    // 2) Every product checked
    if (data.scans && data.scans.length) {
      const sc = tab_(ss, SCANS_TAB, SCAN_HEADERS, SCAN_MANUAL);
      const rows = data.scans.map(k => [data.row['Session ID'], data.row['Name'], "'" + k.time, k.profile, k.product, k.barcode ? "'" + k.barcode : '', k.source,
        k.ings, k.shown, k.why, k.ref, k.auto, k.sec]);
      sc.getRange(sc.getLastRow() + 1, 1, rows.length, SCAN_HEADERS.length).setValues(rows);
    }

    // 3) New taps since the last update
    if (data.clicks && data.clicks.length) {
      const c = tab_(ss, CLICKS_TAB, ['Session ID', 'Name', 'Time from start (m:ss)', 'Screen', 'Type', 'What they tapped'], []);
      const rows = data.clicks.map(k => [data.row['Session ID'], data.row['Name'], "'" + k.time, k.screen, k.type, k.label]);
      c.getRange(c.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    }

    goals_(ss);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Creates the tab if needed and always keeps its header row up to date.
function tab_(ss, name, headers, manual) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const all = headers.concat(manual);
  const cur = sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, all.length).getValues()[0].map(String) : [];
  if (cur.join('|') !== all.join('|')) {
    sh.setRightToLeft(false);
    sh.getRange(1, 1, 1, all.length).setValues([all]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#DBE6FE');
    if (manual.length) sh.getRange(1, headers.length + 1, 1, manual.length).setFontWeight('bold').setBackground('#FEF3C7');
  }
  return sh;
}

function col_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function goals_(ss) {
  let g = ss.getSheetByName(GOALS_TAB);
  if (g && String(g.getRange('H1').getValue()) === VERSION) return;
  // make sure the tabs the formulas point to exist first (otherwise Google shows #REF!)
  tab_(ss, VISITORS_TAB, HEADERS, MANUAL); tab_(ss, SCANS_TAB, SCAN_HEADERS, SCAN_MANUAL);
  if (!g) g = ss.insertSheet(GOALS_TAB);
  g.clear();
  g.setRightToLeft(false);
  const V = "'" + VISITORS_TAB + "'!", S = "'" + SCANS_TAB + "'!";
  const v = h => { const c = col_(HEADERS.indexOf(h) + 1); return V + c + '2:' + c; };
  const vm = i => { const c = col_(HEADERS.length + 1 + i); return V + c + '2:' + c; };
  const s = h => { const c = col_(SCAN_HEADERS.indexOf(h) + 1); return S + c + '2:' + c; };
  const sMan = (() => { const c = col_(SCAN_HEADERS.length + 1); return S + c + '2:' + c; })();
  const g1 = v('Create profile (min) – Goal 1'), g3 = v('False Safe results – Goal 3'), fail = v('Failed scans'), rec = v('Slowest recovery (sec) – Goal 4'),
    core = v('Core features completed (of 3) – Goal 5'), name = v('Name'), help = vm(0), g6 = vm(1), g7 = vm(2),
    sec = s('Seconds to result – Goal 2'), auto = s('Automatic check – Goal 3'), shown = s('Result shown');
  const pct = (num, den) => 'IFERROR(TEXT((' + num + ')/(' + den + '),"0%"),"No data yet")';
  const met = (num, den, t) => 'IFERROR(IF((' + num + ')/(' + den + ')>=' + t + ',"Met","Not met"),"No data yet")';
  const rows = [
    ['Goal', 'Task for the participant', 'How it is measured', 'Target', 'Result now', 'Met?'],
    ['Goal 1 – Effectiveness: set up an allergy/diabetes profile in under 5 minutes without help',
     'Sign up, then create your profile (allergies with severity, or your sugar limit).',
     'Automatic: minutes from sign-up to “Save profile” (Visitors). Mark anyone who needed help as “Yes” in the yellow “Needed help?” column — they don’t count as a success.',
     '≥ 90% of first-time users',
     '=' + pct('COUNTIFS(' + g1 + ',"<=5",' + help + ',"<>Yes")', 'COUNT(' + g1 + ')'),
     '=' + met('COUNTIFS(' + g1 + ',"<=5",' + help + ',"<>Yes")', 'COUNT(' + g1 + ')', 0.9)],
    ['Goal 2 – Efficiency: scan a product and see its safety result in under 15 seconds',
     'Scan the barcode of a packaged food you have at home (or photograph its ingredients).',
     'Automatic: seconds from opening the scanner to the result screen, for every product checked (Scans).',
     '≥ 90% of scans under 15 s',
     '=IFERROR(TEXT(COUNTIFS(' + sec + ',"<15")/COUNT(' + sec + '),"0%")&" of "&COUNT(' + sec + ')&" scans under 15 s (median "&ROUND(MEDIAN(' + sec + '),1)&" s)","No data yet")',
     '=' + met('COUNTIFS(' + sec + ',"<15")', 'COUNT(' + sec + ')', 0.9)],
    ['Goal 3 – Safety: correctly label unsafe / safe every time',
     'Check 2–3 real products, including at least one that contains your allergen.',
     'Automatic: every result is compared with the allergens listed for that product in the Open Food Facts database (Scans → “Automatic check”). For typed or photographed labels, write Correct/Wrong in the yellow column after checking the pack. A False Safe = an unsafe product shown as “Safe”.',
     '0 False Safe results',
     '=IFERROR((COUNTIF(' + auto + ',"FALSE SAFE")+COUNTIFS(' + shown + ',"Safe",' + sMan + ',"Wrong"))&" False Safe out of "&COUNTA(' + shown + ')&" results","No data yet")',
     '=IF(COUNTA(' + shown + ')=0,"No data yet",IF(COUNTIF(' + auto + ',"FALSE SAFE")+COUNTIFS(' + shown + ',"Safe",' + sMan + ',"Wrong")=0,"Met","Not met"))'],
    ['Goal 4 – Error tolerance: recover from a failed scan in under 30 seconds',
     'When a scan fails (barcode not found, no ingredient list, unreadable photo), get a result another way — retake, photograph the label or type it.',
     'Automatic: seconds from the failed scan to the next real result (Visitors → “Slowest recovery”). The app always shows “Uncertain” when data is missing; Undo is counted too.',
     '≥ 80% of users who had a failed scan',
     '=' + pct('COUNTIFS(' + fail + ',">0",' + rec + ',"<=30")', 'COUNTIF(' + fail + ',">0")'),
     '=' + met('COUNTIFS(' + fail + ',">0",' + rec + ',"<=30")', 'COUNTIF(' + fail + ',">0")', 0.8)],
    ['Goal 5 – Easy to learn: complete the core features on the first attempt without help',
     'Without help: create a profile, check a product, and add a recipe to the weekly meal plan.',
     'Automatic: number of core features completed (0–3) per visitor. Anyone marked “Yes” in “Needed help?” doesn’t count as a success.',
     '≥ 80% of users',
     '=' + pct('COUNTIFS(' + core + ',3,' + help + ',"<>Yes")', 'COUNTA(' + name + ')'),
     '=' + met('COUNTIFS(' + core + ',3,' + help + ',"<>Yes")', 'COUNTA(' + name + ')', 0.8)],
    ['Goal 6 – Satisfaction: 4 out of 5 average rating',
     'After using the website, rate scanning, filtering recipes, the warnings and meal planning from 1 to 5 in your separate survey.',
     'Type each person’s average of the 4 ratings into the yellow Goal 6 column in Visitors.',
     'Average ≥ 4.0 / 5',
     '=IFERROR(ROUND(AVERAGE(' + g6 + '),2)&" / 5","No data yet")',
     '=IFERROR(IF(AVERAGE(' + g6 + ')>=4,"Met","Not met"),"No data yet")'],
    ['Goal 7 – Engaging: would use it for weekly meal planning',
     'After using the website, answer “Would you use this app for your weekly meal planning?” (Yes/No) in your separate survey.',
     'Type Yes or No into the yellow Goal 7 column in Visitors.',
     '≥ 70% Yes',
     '=' + pct('COUNTIF(' + g7 + ',"Yes")', 'COUNTA(' + g7 + ')'),
     '=' + met('COUNTIF(' + g7 + ',"Yes")', 'COUNTA(' + g7 + ')', 0.7)],
    ['', '', '', 'Participants so far', '=COUNTA(' + name + ')', '']
  ];
  g.getRange(1, 1, rows.length, 4).setValues(rows.map(r => r.slice(0, 4)));
  g.getRange(1, 5, 1, 2).setValues([rows[0].slice(4)]);
  g.getRange(2, 5, rows.length - 1, 2).setFormulas(rows.slice(1).map(r => r.slice(4).map(x => x || '')));
  g.getRange('H1').setValue(VERSION).setFontColor('#FFFFFF');
  g.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#DBE6FE');
  g.getRange(1, 1, rows.length, 6).setWrap(true).setVerticalAlignment('top');
  g.setColumnWidth(1, 240); g.setColumnWidth(2, 280); g.setColumnWidth(3, 340); g.setColumnWidth(4, 150); g.setColumnWidth(5, 220); g.setColumnWidth(6, 90);
  g.setFrozenRows(1);
}

// Open the web-app link in a browser to check it works.
function doGet() {
  return ContentService.createTextOutput('Connected (v3.1): your website can now send visits to this sheet.');
}
