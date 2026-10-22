
const SHEET_ID = '1Qur3oC1A1FJagFIDAekfmZGzreSDTDtani4x1S2qiuA';
const SHEET_NAME = 'Portfolio Recruiter Contact';

function doGet() {
  return ContentService.createTextOutput(
    'Portfolio recruiter contact service is running.'
  );
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No contact data received.');
    }

    const data = JSON.parse(e.postData.contents);

    const name = String(data.name || '').trim();
    const email = String(data.email || '').trim();
    const company = String(data.company || '').trim();
    const message = String(data.message || '').trim();

    if (!name || !email) {
      throw new Error('Recruiter name and email are required.');
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      throw new Error('Please provide a valid email address.');
    }

    lock.waitLock(10000);

    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Recruiter Name',
        'Email',
        'Company',
        'Message'
      ]);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      name,
      email,
      company,
      message
    ]);

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Contact saved successfully.'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Contact form error: ' + error.message);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
