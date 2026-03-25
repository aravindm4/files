/* ──────────────────────────────────────────────
   Metabase Integration
   ────────────────────────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase')
    .addItem('Generate Performer Summary', 'generatePerformerSummary')
    .addSeparator()
    .addItem('Reset Credentials', 'resetCredentials')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function getCredentials() {
  const props = PropertiesService.getUserProperties();
  let username = props.getProperty('MB_USERNAME');
  let password = props.getProperty('MB_PASSWORD');

  if (!username || !password) {
    const ui = SpreadsheetApp.getUi();

    const u = ui.prompt('Enter Metabase Username');
    if (u.getSelectedButton() !== ui.Button.OK) return;

    const p = ui.prompt('Enter Metabase Password');
    if (p.getSelectedButton() !== ui.Button.OK) return;

    username = u.getResponseText();
    password = p.getResponseText();

    props.setProperty('MB_USERNAME', username);
    props.setProperty('MB_PASSWORD', password);
  }

  return { username, password };
}

function loginToMetabase({ username, password }) {
  const response = UrlFetchApp.fetch(
    'https://data-public.ssmmhospital.com/api/session',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        username: username,
        password: password
      })
    }
  );

  const json = JSON.parse(response.getContentText());
  return json.id; // session token
}

function fetchQuestionById(sessionToken, questionId) {
  const response = UrlFetchApp.fetch(
    'https://data-public.ssmmhospital.com/api/card/' + questionId + '/query/csv',
    {
      method: 'post',
      headers: {
        'X-Metabase-Session': sessionToken
      }
    }
  );

  return parseCsv(response.getContentText());
}

/**
 * Parse CSV text into a 2D array.
 * Handles quoted fields containing commas, newlines, and escaped quotes.
 */
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < text.length && text[i + 1] === '"') {
          field += '"';
          i++;                   // skip escaped quote
        } else {
          inQuotes = false;      // end of quoted field
        }
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        row.push(field);
        field = '';
      } else if (ch === '\n') {
        row.push(field);
        field = '';
        if (row.length > 1 || row[0] !== '') rows.push(row);
        row = [];
      } else if (ch === '\r') {
        // skip carriage return
      } else {
        field += ch;
      }
    }
  }

  // last field / row
  if (field || row.length) {
    row.push(field);
    if (row.length > 1 || row[0] !== '') rows.push(row);
  }

  return rows;
}

function resetCredentials() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('MB_USERNAME');
  props.deleteProperty('MB_PASSWORD');
  SpreadsheetApp.getUi().alert('Metabase credentials have been cleared. You will be prompted on next run.');
}

/* ──────────────────────────────────────────────
   Performer Summary Generator
   ────────────────────────────────────────────── */

/**
 * Extract date portion from a datetime string.
 * Handles "Feb 28, 2026, 11:54 PM" → "Feb 28, 2026"
 * and ISO format "2026-02-28T23:54:00" → "2026-02-28".
 */
function extractDate(datetimeStr) {
  if (!datetimeStr) return '';

  // "Mon DD, YYYY, HH:MM AM/PM" → "Mon DD, YYYY"
  var match = datetimeStr.match(/^([A-Za-z]+ \d{1,2}, \d{4})/);
  if (match) return match[1];

  // ISO "YYYY-MM-DD..."
  var isoMatch = datetimeStr.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoMatch) return isoMatch[1];

  // Fallback: everything before the last comma
  var parts = datetimeStr.split(',');
  if (parts.length >= 3) return parts.slice(0, 2).join(',').trim();

  return datetimeStr;
}

function generatePerformerSummary() {

  var QUESTION_ID = 213; // Metabase question ID for the CHARGEITEM LIST dataset

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var OUTPUT = "Performer_Summary";

  var creds = getCredentials();
  if (!creds) return;
  var sessionToken = loginToMetabase(creds);
  var data = fetchQuestionById(sessionToken, QUESTION_ID);

  if (!data || data.length < 2) {
    SpreadsheetApp.getUi().alert('No data returned from Metabase.');
    return;
  }

  var headers = data[0];
  var normalized = headers.map(function (h) { return String(h).trim().toUpperCase(); });

  var out = ss.getSheetByName(OUTPUT);
  if (!out) out = ss.insertSheet(OUTPUT);
  out.clear();

  // Column lookup (case-insensitive, tries both spaced and snake_case variants)
  function col(name) {
    var idx = normalized.indexOf(name.toUpperCase());
    if (idx === -1) idx = normalized.indexOf(name.toUpperCase().replace(/ /g, '_'));
    return idx;
  }

  var CATEGORY  = col("CATEGORY");
  var PRICE     = col("TOTAL PRICE");
  var DATETIME  = col("DATETIME");
  var PERFORMER = col("PERFORMER");

  var required = {
    "CATEGORY": CATEGORY,
    "TOTAL PRICE": PRICE,
    "DATETIME": DATETIME,
    "PERFORMER": PERFORMER
  };

  var missing = Object.entries(required).filter(function (e) { return e[1] === -1; }).map(function (e) { return e[0]; });
  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Missing columns in Metabase data:\n' + missing.join(', ') +
      '\n\nActual headers found:\n' + headers.join(', ')
    );
    return;
  }

  var categorySet = new Set();
  var grouped = {};

  for (var i = 1; i < data.length; i++) {

    var r = data[i].map(function (v) { return String(v).trim(); });

    var performer = r[PERFORMER];
    if (!performer) continue;

    var date = extractDate(r[DATETIME]);
    var cat  = r[CATEGORY];
    var price = parseFloat(r[PRICE]) || 0;

    categorySet.add(cat);

    var key = performer + "|" + date;

    if (!grouped[key]) {
      grouped[key] = {
        performer: performer,
        date: date,
        cats: {}
      };
    }

    if (!grouped[key].cats[cat]) grouped[key].cats[cat] = 0;
    grouped[key].cats[cat] += price;
  }

  var categories = Array.from(categorySet).sort();

  var header = [
    "PERFORMER",
    "DATE",
  ].concat(categories).concat(["TOTAL"]);

  var output = [header];

  Object.values(grouped).forEach(function (g) {

    var row = [g.performer, g.date];
    var total = 0;

    categories.forEach(function (c) {
      var v = g.cats[c] || 0;
      row.push(v);
      total += v;
    });

    row.push(total);
    output.push(row);
  });

  out.getRange(1, 1, output.length, output[0].length).setValues(output);
  out.setFrozenRows(1);
  out.autoResizeColumns(1, output[0].length);
}
