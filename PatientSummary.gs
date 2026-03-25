/* ──────────────────────────────────────────────
   Metabase Integration
   ────────────────────────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase')
    .addItem('Sync Charges', 'main')
    .addItem('Generate Patient Summary', 'generatePatientSummary')
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

function fetchQuestion(sessionToken) {
  const response = UrlFetchApp.fetch(
    'https://data-public.ssmmhospital.com/api/card/212/query/csv',
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

function writeToSheet(data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  if (!data || data.length === 0) {
    sheet.getRange('A1').setValue('No data');
    return;
  }

  // Pad rows to uniform length in case CSV produced ragged arrays
  const colCount = data[0].length;
  const uniform = data.map(r => {
    while (r.length < colCount) r.push('');
    return r.slice(0, colCount);
  });

  sheet.getRange(1, 1, uniform.length, colCount).setValues(uniform);
}

function main() {
  const creds = getCredentials();
  if (!creds) return;
  const sessionToken = loginToMetabase(creds);
  const data = fetchQuestion(sessionToken);
  writeToSheet(data);
}

/* ──────────────────────────────────────────────
   Patient Summary Generator
   ────────────────────────────────────────────── */

function generatePatientSummary() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const OUTPUT = "Patient_Summary";

  // Fetch data from Metabase (returns 2D array: headers + rows)
  const creds = getCredentials();
  if (!creds) return;
  const sessionToken = loginToMetabase(creds);
  const data = fetchQuestion(sessionToken);

  if (!data || data.length < 2) {
    SpreadsheetApp.getUi().alert('No data returned from Metabase.');
    return;
  }

  const headers = data[0];

  let out = ss.getSheetByName(OUTPUT);
  if (!out) out = ss.insertSheet(OUTPUT);
  out.clear();

  const col = name => headers.indexOf(name);

  const SSMM = col("SSMM ID");
  const PATIENT = col("PATIENT");
  const START = col("ENCOUNTER START DATE");
  const END = col("ENCOUNTER END DATE");
  const CATEGORY = col("CATEGORY");
  const PRICE = col("TOTAL PRICE");
  const CARETEAM = col("CARE TEAM");

  const categorySet = new Set();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {

    const r = data[i];

    if (!r[PATIENT]) continue;

    const key =
      r[SSMM] + "|" +
      r[START] + "|" +
      r[END];

    const cat = r[CATEGORY];
    const price = parseFloat(r[PRICE]) || 0;

    categorySet.add(cat);

    if (!grouped[key]) {

      grouped[key] = {
        ssmm: r[SSMM],
        patient: r[PATIENT],
        care: new Set(),
        start: r[START],
        end: r[END],
        cats: {}
      };

    }

    if (r[CARETEAM]) grouped[key].care.add(r[CARETEAM]);

    if (!grouped[key].cats[cat]) grouped[key].cats[cat] = 0;

    grouped[key].cats[cat] += price;

  }

  const categories = Array.from(categorySet).sort();

  const header = [
    "SSMM ID",
    "PATIENT",
    "CARE TEAM",
    "ENCOUNTER START DATE",
    "ENCOUNTER END DATE",
    ...categories,
    "TOTAL"
  ];

  const output = [header];

  Object.values(grouped).forEach(g => {

    let row = [
      g.ssmm,
      g.patient,
      Array.from(g.care).join(", "),
      g.start,
      g.end
    ];

    let total = 0;

    categories.forEach(c => {

      let v = g.cats[c] || 0;
      row.push(v);
      total += v;

    });

    row.push(total);

    output.push(row);

  });

  out.getRange(1,1,output.length,output[0].length).setValues(output);

  out.setFrozenRows(1);
  out.autoResizeColumns(1,output[0].length);

}
