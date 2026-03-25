/* ──────────────────────────────────────────────
   Metabase Integration
   ────────────────────────────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase')
    .addItem('Generate Patient Summary', 'generatePatientSummary')
    .addToUi();
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
    'https://data-public.ssmmhospital.com/api/card/212/query/json',
    {
      method: 'post',
      headers: {
        'X-Metabase-Session': sessionToken
      }
    }
  );

  return JSON.parse(response.getContentText());
}

/* ──────────────────────────────────────────────
   Patient Summary Generator
   ────────────────────────────────────────────── */

function generatePatientSummary() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const OUTPUT = "Patient_Summary";

  // Fetch data from Metabase instead of Raw_Data sheet
  const creds = getCredentials();
  if (!creds) return;
  const sessionToken = loginToMetabase(creds);
  const jsonData = fetchQuestion(sessionToken);

  if (!jsonData || jsonData.length === 0) {
    SpreadsheetApp.getUi().alert('No data returned from Metabase.');
    return;
  }

  // Convert JSON objects to 2D array (same format as sheet.getDataRange().getValues())
  const headers = Object.keys(jsonData[0]);
  const data = [headers];
  jsonData.forEach(row => data.push(headers.map(h => row[h])));

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
