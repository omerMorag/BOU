/**
 * בואו · קליטת הרשמות מדף הנחיתה לתוך Google Sheets
 *
 * איך מתקינים:
 * 1. פותחים את הגיליון, ובתפריט Extensions → Apps Script
 * 2. מוחקים את מה שיש שם ומדביקים את כל הקובץ הזה
 * 3. Deploy → New deployment → סוג: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. מאשרים את ההרשאות, ומעתיקים את ה-Web app URL (מסתיים ב-/exec)
 * 5. מדביקים אותו ב-SHEET_URL בסקריפט שבסוף index.html
 */

const SHEET_NAME = 'נרשמים';
const HEADERS = ['תאריך הרשמה', 'שם', 'טלפון', 'גיל', 'מין', 'הערות', 'מחיר', 'שילם?'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.name || !d.phone) throw new Error('missing fields');

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      clean_(d.name),
      "'" + clean_(d.phone).replace(/^'/, ''),   // שומר את ה-0 בתחילת המספר
      Number(d.age) || '',
      clean_(d.gender),
      clean_(d.note),
      Number(d.price) || '',
      false
    ]);
    sheet.getRange(sheet.getLastRow(), HEADERS.length).insertCheckboxes();

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* מריצים פעם אחת: נותן שם לגיליון ויוצר את לשונית "נרשמים" */
function setup() {
  SpreadsheetApp.getActiveSpreadsheet().rename('בואו · נרשמים לריטריט 16.10');
  getSheet_();
}

/* בדיקה מהדפדפן: פתיחת כתובת ה-/exec צריכה להציג {"ok":true,"status":"ready"} */
function doGet() {
  getSheet_();
  return json_({ ok: true, status: 'ready' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#FFF6C8');
    sheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  }
  return sheet;
}

/* מקצר טקסט ומונע נוסחאות שמוזרקות דרך הטופס */
function clean_(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
