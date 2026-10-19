/**
 * Sargam by Sanskriti — website forms → Google Sheet + email alerts
 *
 * Three forms on the website post here. Each goes to its own tab:
 *   Enquiries       ← contact form + bestseller "Get a quote" pop-up
 *   Custom Designs  ← "Request a custom design" pop-up
 *   Reviews         ← "Send us two lines" review pop-up
 * Every submission also emails ALERT_EMAIL with the details and a link to the sheet.
 */
const ALERT_EMAIL = 'sargambysanskriti@gmail.com';   // set to '' to stop email alerts

const FORMS = {
  enquiry: {
    tab: 'Enquiries',
    headers: ['Timestamp','Name','Phone','Email','Company / College','Items / Collections','Product','Item code','Quantity','Needed by','Branding','Requirement','Status'],
    row: p => [p.name, p.phone, p.email, p.organisation, p.item, p.product, p.item_code, p.quantity, p.needed_by, p.branding, p.requirement, 'New'],
    subject: p => `New enquiry: ${p.name} — ${p.quantity || '?'} × ${p.product || p.item || ''}`
  },
  custom_design: {
    tab: 'Custom Designs',
    headers: ['Timestamp','Name','Phone','Email','Company / College','Design on (products)','Concept','Approx. quantity','Needed by','Status'],
    row: p => [p.name, p.phone, p.email, p.organisation, p.products, p.concept, p.quantity, p.needed_by, 'New'],
    subject: p => `New custom design request: ${p.name} — ${p.products || ''}`
  },
  review: {
    tab: 'Reviews',
    headers: ['Timestamp','Name','Role / Company','Rating (1-5)','Review','OK to publish?','Status'],
    row: p => [p.name, p.organisation, p.rating, p.review, p.publish_ok || 'No', 'New'],
    subject: p => `New review (${p.rating || '?'}★) from ${p.name}`
  }
};

function sheetFor_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(cfg.tab);
  if (!sh) sh = ss.insertSheet(cfg.tab);
  if (sh.getLastRow() === 0) {
    sh.appendRow(cfg.headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cfg.headers.length).setFontWeight('bold').setBackground('#1f3a93').setFontColor('#ffffff');
  }
  return sh;
}

/** Run once from the editor: creates the three tabs and logs the sheet link. */
function setup() {
  Object.values(FORMS).forEach(sheetFor_);
  Logger.log('Form entries will be saved to: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
}

/** Optional: run to test each form end-to-end (adds sample rows + sends emails). */
function testAll() {
  doPost({ parameter: { form_type: 'enquiry', name: 'Test Enquiry', phone: '9999999999', item: 'Joining Kits, Drinkware', quantity: '50' } });
  doPost({ parameter: { form_type: 'custom_design', name: 'Test Design', phone: '9999999999', products: 'T-shirts, Diaries', concept: 'Diwali kit' } });
  doPost({ parameter: { form_type: 'review', name: 'Test Review', rating: '5', review: 'Great!' } });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const cfg = FORMS[p.form_type] || FORMS.enquiry;
    const clean = v => String(v == null ? '' : v).slice(0, 3000).replace(/^[=+\-@]/, "'$&");
    const safe = {}; Object.keys(p).forEach(k => safe[k] = clean(p[k]));
    sheetFor_(cfg).appendRow([new Date()].concat(cfg.row(safe)));
    if (ALERT_EMAIL) {
      const lines = cfg.headers.slice(1, -1).map((h, i) => h + ': ' + (cfg.row(safe)[i] || '—')).join('\n');
      MailApp.sendEmail({
        to: ALERT_EMAIL,
        subject: cfg.subject(safe),
        body: lines + '\n\nOpen the sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
        replyTo: /@/.test(safe.email || '') ? safe.email : ALERT_EMAIL
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}
