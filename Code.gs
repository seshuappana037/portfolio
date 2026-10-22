/**
 * Google Apps Script backend for the Seshu Appana portfolio contact form.
 * Setup: create a Google Sheet, open Extensions > Apps Script, paste this file,
 * set SHEET_ID below, then Deploy > New deployment > Web app.
 * Execute as: Me. Who has access: Anyone. Copy the /exec URL into index.html.
 */
const SHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Recruiter Contacts';
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const required = ['name', 'email', 'subject', 'message'];
    required.forEach(k => { if (!String(data[k] || '').trim()) throw new Error('Missing required field: ' + k); });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(data.email))) throw new Error('Invalid email');
    if (SHEET_ID === 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') throw new Error('Configure SHEET_ID first');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['Submitted At','Recruiter Name','Email','Company','Subject','Message','Source']); sh.setFrozenRows(1); }
    sh.appendRow([new Date(), String(data.name).slice(0,200), String(data.email).slice(0,320), String(data.company || '').slice(0,200), String(data.subject).slice(0,300), String(data.message).slice(0,5000), String(data.source || '').slice(0,1000)]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err.message || err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
function doGet() { return ContentService.createTextOutput('Portfolio contact endpoint is running.'); }
