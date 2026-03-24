// ─── Constants ───────────────────────────────────────────────────────────────

// Apps Script's UrlFetchApp silently truncates responses above ~50 MB.
// Treat anything over 40 MB as too large to parse safely.
var MAX_SAFE_RESPONSE_BYTES = 40 * 1024 * 1024;

// If a single page returns more than this multiple of the requested page size,
// the server is ignoring the pagination parameter.
var PAGINATION_FAILURE_MULTIPLIER = 2;

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
    'https://data-public.ssmmhospital.com/api/session',
    {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({ username: username, password: password })
    }
  );

  const status = response.getResponseCode();
  const body   = response.getContentText();
  if (status !== 200) {
    throw new Error('Metabase login failed (HTTP ' + status + '): ' + body.slice(0, 300));
  }

  const json = JSON.parse(body);
  if (!json.id) throw new Error('Metabase login did not return a session token: ' + body.slice(0, 300));
  return json.id; // session token
}

function fetchOnePage(sessionToken, page, pageSize) {
  const response = UrlFetchApp.fetch(
    'https://data-public.ssmmhospital.com/api/card/212/query',
    {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: {
        'X-Metabase-Session': sessionToken
      },
      payload: JSON.stringify({
        parameters: [],
        page: { page: page, items: pageSize }
      })
    }
  );

  const status = response.getResponseCode();
  const body   = response.getContentText();

  if (status === 401 || status === 403) {
    return null; // signal caller to re-login
  }

  if (status !== 202 && status !== 200) {
    throw new Error('Metabase query failed on page ' + page + ' (HTTP ' + status + '): ' + body.slice(0, 300));
  }

  let json;
  try {
    json = JSON.parse(body);
  } catch (e) {
    // Apps Script silently truncates URL Fetch responses larger than ~50 MB,
    // which corrupts the JSON. When this happens the response starts with '['
    // (the flat-array export format Metabase uses for this card type) and the
    // Metabase page: {page, items} pagination parameter is being ignored, so
    // the full dataset is returned in one shot.
    //
    // Fix required on the Metabase card: add {{limit}} and {{offset}} template
    // variables to the SQL query and pass them as parameters here, OR add a
    // date-range filter to the card and pass start/end dates as parameters so
    // each request covers only a manageable slice of data.
    if (body.length > MAX_SAFE_RESPONSE_BYTES) {
      throw new Error(
        'Response on page ' + page + ' is ' + Math.round(body.length / 1024 / 1024) + ' MB — ' +
        'too large for Google Apps Script (50 MB limit). ' +
        'The Metabase card is returning all rows without respecting the pagination parameter. ' +
        'Add LIMIT/OFFSET template variables to the card SQL, or add a date-range filter ' +
        'so each request returns a manageable slice of data.'
      );
    }
    throw new Error(
      'JSON parse error on page ' + page + ': ' + e.message +
      ' — first 200 chars of response: ' + body.slice(0, 200)
    );
  }

  // Metabase can return results in two different shapes depending on the card
  // type and server version:
  //
  //   Standard format: { data: { cols: [{name},...], rows: [[v,v,...],...]}, ... }
  //   Flat-array format: [ { col: val, ... }, ... ]
  //
  // Normalise the flat-array format so the rest of the code only has to deal
  // with one shape.
  if (Array.isArray(json)) {
    if (json.length === 0) return { data: { cols: [], rows: [] } };
    const colNames = Object.keys(json[0]);
    return {
      data: {
        cols: colNames.map(name => ({ name: name })),
        rows: json.map(record => colNames.map(name => record[name]))
      }
    };
  }

  return json;
}

function fetchQuestion(sessionToken) {
  const PAGE_SIZE = 2000;
  const allRows = [];
  let page = 1;
  let cols = null;
  let token = sessionToken;

  while (true) {
    let json = fetchOnePage(token, page, PAGE_SIZE);

    // Session expired mid-run — re-authenticate once and retry
    if (json === null) {
      const creds = getCredentials();
      if (!creds) throw new Error('Session expired and credentials were not re-entered.');
      token = loginToMetabase(creds);
      json = fetchOnePage(token, page, PAGE_SIZE);
      if (json === null) throw new Error('Re-authentication failed; please check your credentials.');
    }

    const data = json.data;

    if (!data || !data.rows || data.rows.length === 0) break;

    // If the server returned far more rows than the requested page size it has
    // ignored the pagination parameter and is dumping the whole dataset.  Fail
    // fast with a clear message rather than silently processing a partial or
    // truncated result set.
    if (data.rows.length > PAGE_SIZE * PAGINATION_FAILURE_MULTIPLIER) {
      throw new Error(
        'Page ' + page + ' returned ' + data.rows.length + ' rows but only ' + PAGE_SIZE +
        ' were requested — the Metabase card is ignoring the pagination parameter ' +
        'and returning the full dataset. Add LIMIT/OFFSET template variables to the ' +
        'card SQL, or add a date-range filter to keep each response under 50 MB.'
      );
    }

    if (!cols) {
      cols = data.cols.map(c => c.name);
    }

    data.rows.forEach(row => {
      const obj = {};
      cols.forEach((col, i) => { obj[col] = row[i]; });
      allRows.push(obj);
    });

    if (data.rows.length < PAGE_SIZE) break;
    page++;
  }

  return allRows;
}

// ─── Generate Patient Summary ─────────────────────────────────────────────────

// Accepts the JSON array returned directly by fetchQuestion()
function generatePatientSummary(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const OUTPUT = 'Patient_Summary';

  let out = ss.getSheetByName(OUTPUT);
  if (!out) out = ss.insertSheet(OUTPUT);
  out.clear();

  if (!records || records.length === 0) {
    out.getRange('A1').setValue('No data');
    return;
  }

  const categorySet = new Set();
  const grouped = {};

  records.forEach(r => {
    // Accept whichever field name Metabase returns for the patient name
    if (!r.patient_name) r.patient_name = r['Patient Name'] || r['patient'] || r['name'] || '';
    if (!r.patient_name) return;

    const key = r.ssmm_id + '|' + r.period_start + '|' + r.period_end;

    const cat   = r.category;
    const price = parseFloat(r.total_price) || 0;

    categorySet.add(cat);

    if (!grouped[key]) {
      grouped[key] = {
        ssmm:    r.ssmm_id,
        patient: r.patient_name,
        care:    new Set(),
        start:   r.period_start,
        end:     r.period_end,
        cats:    {}
      };
    }

    if (r.care_team_members) grouped[key].care.add(r.care_team_members);

    if (!grouped[key].cats[cat]) grouped[key].cats[cat] = 0;
    grouped[key].cats[cat] += price;
  });

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

  generatePatientSummary(data);

  SpreadsheetApp.getUi().alert('Done! Patient_Summary sheet has been updated.');
}
