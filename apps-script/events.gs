/**
 * @OnlyCurrentDoc
 *
 * (The line above limits this script to the one spreadsheet it is attached to.
 *  Without it, Google asks for access to ALL of your spreadsheets.)
 *
 * Events feed for the Office Booking System.
 *
 * Lives inside the shared Google Sheet (Extensions → Apps Script) and hands the
 * booking page ONLY the upcoming Lecture Series and Jour fixe dates. The page
 * never sees the sheet itself, so the sheet can stay private.
 *
 * Deploy: Deploy → New deployment → type "Web app"
 *         Execute as: Me    Who has access: Anyone
 * Then paste the web app URL into EVENTS_URL in index.html.
 * After editing this script: Deploy → Manage deployments → Edit → New version.
 */

const LOOKBACK_DAYS = 60;        // also send events from the last 60 days
const INCLUDE_AGENDA = false;    // true = also send the Jour fixe agenda text

function doGet() {
  // Must be created from inside the sheet (Extensions → Apps Script) by someone who can edit it.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const since = new Date(Date.now() - LOOKBACK_DAYS * 24 * 60 * 60 * 1000);

  // Date cells arrive as Date objects, but "instanceof Date" is unreliable in Apps Script.
  const isDate = v => Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
  // Find a tab by name, ignoring upper/lower case and spaces.
  const norm = s => String(s).toLowerCase().replace(/\s+/g, '');
  const findTab = name => ss.getSheets().filter(s => norm(s.getName()) === norm(name))[0];
  const missing = [];

  // Rows from one tab: [date, ...the chosen columns], only rows with a real date.
  const read = (tabName, cols) => {
    const sheet = findTab(tabName);
    if (!sheet) { missing.push(tabName); return []; }
    return sheet.getDataRange().getValues().slice(1)
      .filter(r => isDate(r[0]) && r[0].getTime() >= since.getTime())
      .map(r => [Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')]
        .concat(cols.map(c => String(r[c] == null ? '' : r[c]).trim())));
  };

  // Lecture Series Schedule: Date | Speaker | Host | Title | Room | Chair
  const lecture = read('Lecture Series Schedule', [1, 2, 3, 4, 5])
    .filter(r => !/^no lecture/i.test(r[1]));                 // "No Lecture Series"

  // JourFixe: Date | Type | Chair | (minutes) | Agenda | (decisions – never sent)
  const jf = read('JourFixe', [1, 2, 4])
    .filter(r => !/^\s*(no jour ?fixe|cancel)/i.test(r[3]))   // "no Jour fixe", "cancelled …"
    .map(r => INCLUDE_AGENDA ? r : r.slice(0, 3));

  return ContentService
    .createTextOutput(JSON.stringify({ lecture: lecture, jf: jf, missingTabs: missing }))
    .setMimeType(ContentService.MimeType.JSON);
}
