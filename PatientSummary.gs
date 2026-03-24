// ─── Menu ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase')
    .addItem('Sync Charges & Generate Summary', 'main')
    .addToUi();
}

// ─── Credentials ─────────────────────────────────────────────────────────────

function getCredentials() {
  const props = PropertiesService.getUserProperties();
  let username = props.getProperty('MB_USERNAME');
  let password = props.getProperty('MB_PASSWORD');

  if (!username || !password) {
    const ui = SpreadsheetApp.getUi();

    const u = ui.prompt('Enter Metabase Username');
    if (u.getSelectedButton() !== ui.Button.OK) return null;

    const p = ui.prompt('Enter Metabase Password');
    if (p.getSelectedButton() !== ui.Button.OK) return null;

    username = u.getResponseText();
    password = p.getResponseText();

    props.setProperty('MB_USERNAME', username);
    props.setProperty('MB_PASSWORD', password);
  }

  return { username, password };
}

// ─── Metabase API ─────────────────────────────────────────────────────────────

function loginToMetabase({ username, password }) {
  const response = UrlFetchApp.fetch(
    'https://metabase.ohc.network/api/session',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ username: username, password: password })
    }
  );

  const json = JSON.parse(response.getContentText());
  return json.id; // session token
}

function fetchQuestion(sessionToken) {
  const response = UrlFetchApp.fetch(
    'https://metabase.ohc.network/api/card/874/query/json',
    {
      method: 'post',
      headers: {
        'X-Metabase-Session': sessionToken
      }
    }
  );

  return JSON.parse(response.getContentText());
}

// ─── Write raw data to Raw_Data sheet ────────────────────────────────────────

function writeRawData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET_NAME = 'Raw_Data';

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clear();

  if (!data || data.length === 0) {
    sheet.getRange('A1').setValue('No data');
    return;
  }

  const headers = Object.keys(data[0]);
  const rows = data.map(row => headers.map(h => row[h] !== undefined ? row[h] : ''));

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ─── Generate Patient Summary ─────────────────────────────────────────────────

function generatePatientSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SOURCE = 'Raw_Data';
  const OUTPUT = 'Patient_Summary';

  const sheet = ss.getSheetByName(SOURCE);
  if (!sheet) throw new Error('Raw_Data sheet not found');

  let out = ss.getSheetByName(OUTPUT);
  if (!out) out = ss.insertSheet(OUTPUT);
  out.clear();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = name => headers.indexOf(name);

  const SSMM     = col('SSMM ID');
  const PATIENT  = col('PATIENT');
  const START    = col('ENCOUNTER START DATE');
  const END      = col('ENCOUNTER END DATE');
  const CATEGORY = col('CATEGORY');
  const PRICE    = col('TOTAL PRICE');
  const CARETEAM = col('CARE TEAM');

  const categorySet = new Set();
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    if (!r[PATIENT]) continue;

    const key = r[SSMM] + '|' + r[START] + '|' + r[END];

    const cat   = r[CATEGORY];
    const price = parseFloat(r[PRICE]) || 0;

    categorySet.add(cat);

    if (!grouped[key]) {
      grouped[key] = {
        ssmm:    r[SSMM],
        patient: r[PATIENT],
        care:    new Set(),
        start:   r[START],
        end:     r[END],
        cats:    {}
      };
    }

    if (r[CARETEAM]) grouped[key].care.add(r[CARETEAM]);

    if (!grouped[key].cats[cat]) grouped[key].cats[cat] = 0;
    grouped[key].cats[cat] += price;
  }

  const categories = Array.from(categorySet).sort();

  const header = [
    'SSMM ID',
    'PATIENT',
    'CARE TEAM',
    'ENCOUNTER START DATE',
    'ENCOUNTER END DATE',
    ...categories,
    'TOTAL'
  ];

  const output = [header];

  Object.values(grouped).forEach(g => {
    const row = [
      g.ssmm,
      g.patient,
      Array.from(g.care).join(', '),
      g.start,
      g.end
    ];

    let total = 0;

    categories.forEach(c => {
      const v = g.cats[c] || 0;
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

// ─── Main entry point ─────────────────────────────────────────────────────────

function main() {
  const creds = getCredentials();
  if (!creds) return; // user cancelled credential prompt

  const sessionToken = loginToMetabase(creds);
  const data = fetchQuestion(sessionToken);

  writeRawData(data);
  generatePatientSummary();

  SpreadsheetApp.getUi().alert('Done! Patient_Summary sheet has been updated.');
}
