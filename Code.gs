const SHEET_ID       = '14aBkpWKEoLfJ3-_N33HmaltlMdd8gbzuXEeunnABm30';

const TASKS_SHEET    = 'Tasks';
const PERIOD_SHEET   = 'PeriodEntries';
const LINKS_SHEET    = 'Links';

const TASK_HEADERS   = ['id','title','description','category','dueDate','priority','completed','createdAt'];
const PERIOD_HEADERS = ['date'];
const LINKS_HEADERS  = ['id','label','url','createdAt'];

// ── GET ──────────────────────────────────────────────────────
function doGet(e) {
  const type = e.parameter.type;
  try {
    if (type === 'tasks')  return respond({ success: true, data: readSheet(TASKS_SHEET,  TASK_HEADERS)   });
    if (type === 'period') return respond({ success: true, data: readSheet(PERIOD_SHEET, PERIOD_HEADERS) });
    if (type === 'links')  return respond({ success: true, data: readSheet(LINKS_SHEET,  LINKS_HEADERS)  });
    return respond({ success: false, error: 'Unknown type' });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

// ── POST ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const type = body.type;
    const data = body.data;

    if (type === 'tasks')  { writeSheet(TASKS_SHEET,  TASK_HEADERS,   data); return respond({ success: true }); }
    if (type === 'period') { writeSheet(PERIOD_SHEET, PERIOD_HEADERS, data); return respond({ success: true }); }
    if (type === 'links')  { writeSheet(LINKS_SHEET,  LINKS_HEADERS,  data); return respond({ success: true }); }
    return respond({ success: false, error: 'Unknown type' });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

// ── Helpers ──────────────────────────────────────────────────
function getOrCreateSheet(name, headers) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#e2e8f0');
  }
  return sheet;
}

function readSheet(name, headers) {
  const sheet = getOrCreateSheet(name, headers);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const head = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    head.forEach((col, i) => { obj[col] = row[i] === '' ? null : row[i]; });
    if ('completed' in obj) obj.completed = obj.completed === true || obj.completed === 'true';
    if ('priority'  in obj) obj.priority  = obj.priority  === true || obj.priority  === 'true';
    return obj;
  });
}

function writeSheet(name, headers, data) {
  const sheet = getOrCreateSheet(name, headers);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  if (!data || data.length === 0) return;
  const rows = data.map(item => headers.map(h => {
    const v = item[h];
    return v === undefined || v === null ? '' : v;
  }));
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function respond(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
