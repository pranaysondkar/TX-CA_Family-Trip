/**
 * Family trip bookings API, pasted into the online Google Sheet:
 *   Extensions → Apps Script → replace everything with this file → set FAMILY_PIN → Deploy → New deployment
 *   Type: Web app · Execute as: Me · Who has access: Anyone → Deploy → copy the URL ending in /exec.
 *
 * GET  <url>                         → {ok, rows:[{id, booked, actual, inPlan, paidBy}]}
 * POST <url>  body (text/plain JSON) → {pin, id, booked, actual, paidBy}  updates one row in the Bookings tab.
 * Anyone can read (same as the public website); only people who know FAMILY_PIN can change anything.
 */
const FAMILY_PIN = 'CHANGE-ME';   // pick a short PIN and share it with the family only
const TAB = 'Bookings';

function doGet() {
  return json({ ok: true, rows: readRows_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Bad request' }); }
  if (String(body.pin || '') !== FAMILY_PIN) return json({ ok: false, error: 'Wrong PIN' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(TAB);
    const data = sh.getDataRange().getValues();
    const c = cols_(data[0]);
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][c.id]).trim() !== body.id) continue;
      sh.getRange(r + 1, c.booked + 1).setValue(body.booked ? 'Yes' : 'No');
      const amt = body.actual === null || body.actual === '' || body.actual === undefined ? '' : Number(body.actual);
      sh.getRange(r + 1, c.actual + 1).setValue(isNaN(amt) ? '' : amt);
      if (c.paidBy >= 0 && body.paidBy !== undefined) sh.getRange(r + 1, c.paidBy + 1).setValue(String(body.paidBy).slice(0, 40));
      SpreadsheetApp.flush();
      return json({ ok: true, rows: readRows_() });
    }
    return json({ ok: false, error: 'Item not found: ' + body.id });
  } finally {
    lock.releaseLock();
  }
}

function readRows_() {
  const data = SpreadsheetApp.getActive().getSheetByName(TAB).getDataRange().getValues();
  const c = cols_(data[0]);
  return data.slice(1).filter(r => String(r[c.id]).trim()).map(r => ({
    id: String(r[c.id]).trim(),
    booked: /^y/i.test(String(r[c.booked])),
    actual: r[c.actual] === '' || r[c.actual] === null ? null : Number(r[c.actual]),
    inPlan: c.inPlan < 0 ? null : !/^n/i.test(String(r[c.inPlan])),
    paidBy: c.paidBy < 0 ? '' : String(r[c.paidBy] || ''),
  }));
}

function cols_(h) {
  const find = p => h.findIndex(x => String(x).trim().toLowerCase().startsWith(p));
  return { id: find('id'), booked: find('booked'), actual: find('actual'), inPlan: find('in plan'), paidBy: find('paid by') };
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
