// ============================================================
//  Zimbabwe Corporate Summit — Check-in Backend
// ============================================================

const SHEET_ID = '1SDqFPeL6HMO_DLxHz0gYKFfpr87Dp41U48JdvfKrMuk';

const COL = {
  FIRST:    'First Name',
  SURNAME:  'Surname',
  EMAIL:    'Email',
  SESSIONS: 'Sessions',
};

const SESSIONS = [
  {id:'summit', name:'Corporate Summit', day:1, date:'Jun\n28'}
];

// ── HTTP handler ─────────────────────────────────────────────

function doGet(e) {
  const params = e.parameter;
  const action = params.action || '';
  let result;

  try {
    if (action === 'checkin') {
      result = handleCheckin(params.code, params.session);
    } else if (action === 'sessions') {
      result = { sessions: SESSIONS };
    } else if (action === 'count') {
      result = handleCount(params.session);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Check-in logic ────────────────────────────────────────────

function handleCheckin(email, sessionName) {
  if (!email || !sessionName) return { status: 'denied', reason: 'Missing email or session' };

  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName('Registrants');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const idx = {
    first:    headers.indexOf(COL.FIRST),
    surname:  headers.indexOf(COL.SURNAME),
    email:    headers.indexOf(COL.EMAIL),
    sessions: headers.indexOf(COL.SESSIONS),
  };

  for (const [key, val] of Object.entries(idx)) {
    if (val === -1) return { status: 'denied', reason: `Sheet column not found: ${key}` };
  }

  const rowIndex = data.findIndex(
    (row, i) => i > 0 && String(row[idx.email]).trim().toLowerCase() === email.trim().toLowerCase()
  );

  if (rowIndex === -1) {
    return { status: 'walkin', name: email, reason: 'Not in registration list' };
  }

  const row       = data[rowIndex];
  const fullName  = `${row[idx.first]} ${row[idx.surname]}`.trim();
  const attendeeEmail = row[idx.email];

  const registeredSessions = String(row[idx.sessions])
    .split(',')
    .map(s => s.trim().toLowerCase());

  if (!registeredSessions.includes(sessionName.trim().toLowerCase())) {
    return { status: 'denied', name: fullName, reason: `Not registered for ${sessionName}` };
  }

  const checkinKey = 'checkin_' + sessionName.replace(/[^a-zA-Z0-9]/g, '_');
  let checkinColIdx = headers.indexOf(checkinKey);

  if (checkinColIdx === -1) {
    checkinColIdx = headers.length;
    sheet.getRange(1, checkinColIdx + 1).setValue(checkinKey);
    headers.push(checkinKey);
  }

  if (row[checkinColIdx]) {
    return { status: 'already', name: fullName, time: row[checkinColIdx] };
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(rowIndex + 1, checkinColIdx + 1).setValue(timestamp);
  logAttendance(fullName, attendeeEmail, sessionName, timestamp);

  return { status: 'approved', name: fullName, detail: sessionName, time: timestamp };
}

// ── Count expected attendees ──────────────────────────────────

function handleCount(sessionName) {
  if (!sessionName) return { total: 0 };

  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName('Registrants');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const sessIdx = headers.indexOf(COL.SESSIONS);

  let total = 0;
  for (let i = 1; i < data.length; i++) {
    const sessions = String(data[i][sessIdx]).split(',').map(s => s.trim().toLowerCase());
    if (sessions.includes(sessionName.trim().toLowerCase())) total++;
  }

  return { total };
}

// ── Attendance log ────────────────────────────────────────────

function logAttendance(name, email, sessionName, timestamp) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let logSheet = ss.getSheetByName('Attendance Log');

  if (!logSheet) {
    logSheet = ss.insertSheet('Attendance Log');
    logSheet.appendRow(['Timestamp', 'Full Name', 'Email', 'Session']);
    logSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }

  logSheet.appendRow([timestamp, name, email, sessionName]);
}

// ── Style sheet headers ───────────────────────────────────────

function styleSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Registrants');
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  header.setBackground('#004165');
  header.setFontColor('#f2df74');
  header.setFontWeight('bold');
  header.setFontFamily('Arial');
  header.setFontSize(11);
  sheet.setFrozenRows(1);

  const log = ss.getSheetByName('Attendance Log');
  if (log) {
    const logHeader = log.getRange(1, 1, 1, log.getLastColumn());
    logHeader.setBackground('#004165');
    logHeader.setFontColor('#f2df74');
    logHeader.setFontWeight('bold');
    log.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('Sheet styled!');
}

// ── Send QR code emails ───────────────────────────────────────
// Uses Resend (resend.com) so emails send from divisionc@toastmasters129.org
// Steps: sign up at resend.com → verify toastmasters129.org domain → create API key → paste below

const MAILJET_API_KEY    = 'TODO_paste_your_mailjet_api_key_here';
const MAILJET_SECRET_KEY = 'TODO_paste_your_mailjet_secret_key_here';
const FROM_EMAIL         = 'divisionc@toastmasters129.org';
const FROM_NAME          = 'Zimbabwe Corporate Summit';

function sendQREmails() {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName('Registrants');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const idx = {
    first:    headers.indexOf(COL.FIRST),
    surname:  headers.indexOf(COL.SURNAME),
    email:    headers.indexOf(COL.EMAIL),
    sessions: headers.indexOf(COL.SESSIONS),
    sent:     headers.indexOf('QR Sent'),
  };

  let sentColIdx = idx.sent;
  if (sentColIdx === -1) {
    sentColIdx = headers.length;
    sheet.getRange(1, sentColIdx + 1).setValue('QR Sent');
  }

  Logger.log(`Headers: ${JSON.stringify(headers)}`);
  Logger.log(`sentColIdx: ${sentColIdx}, total rows: ${data.length - 1}`);

  let sent = 0;

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    const email = String(row[idx.email]).trim();
    Logger.log(`Row ${i}: email="${email}", sentValue="${row[sentColIdx]}"`);
    if (!email) { Logger.log(`Row ${i}: skipped — no email`); continue; }
    if (row[sentColIdx]) { Logger.log(`Row ${i}: skipped — already sent`); continue; }

    const firstName = row[idx.first];
    const fullName  = `${row[idx.first]} ${row[idx.surname]}`.trim();

    const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(email)}&size=250&margin=2`;

    // TODO: replace with public Drive URL for ToastmastersLogo.png
    // Upload to Drive → Share (Anyone with link) → https://drive.google.com/uc?export=view&id=FILE_ID
    const LOGO_URL = 'https://drive.google.com/uc?export=view&id=TODO_FILE_ID';

    const subject  = `Your Summit Pass — Zimbabwe Corporate Summit`;
    const textBody = `Hi ${firstName},\n\nYou're registered for the Toastmasters Zimbabwe Corporate Summit: The Future is Human!\n\nDate: Saturday 28 June 2025, 14:00 – 16:00\n\nYour QR code: ${qrUrl}\n\nShow it at the door to check in.\n\nSee you there!\nToastmasters Zimbabwe`;

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:500px;margin:auto;color:#1a1a1c">

        <div style="background:#ffffff;padding:28px 28px 20px;text-align:center;border-radius:12px 12px 0 0;border-bottom:3px solid #006094">
          <img src="${LOGO_URL}" width="110" alt="Toastmasters Zimbabwe" style="display:block;margin:0 auto 18px" />
          <div style="line-height:1;margin-bottom:8px">
            <div style="color:#004165;font-size:22px;font-weight:600;letter-spacing:1px">The Future</div>
            <div style="color:#004165;line-height:1">
              <span style="font-size:18px;font-weight:400">IS</span><span style="font-size:52px;font-weight:900;letter-spacing:-1px">HUMAN</span>
            </div>
          </div>
          <div style="color:#004165;font-style:italic;font-size:12px;margin-bottom:16px">Mastering The Skills Machines Can't</div>
          <div style="display:inline-block;background:linear-gradient(135deg,#004165,#006094);color:#ffffff;padding:6px 22px;border-radius:4px;font-weight:700;font-size:12px;letter-spacing:3px">SUMMIT PASS</div>
        </div>

        <div style="background:#fff;border:1px solid #e0e8ed;border-top:none;padding:28px;border-radius:0 0 12px 12px">
          <p style="margin:0 0 16px">Hi <strong>${firstName}</strong>,</p>
          <p style="margin:0 0 16px">You're all set for the <strong>Toastmasters Zimbabwe Corporate Summit</strong>.</p>
          <table style="width:100%;border-collapse:collapse;margin:0 0 24px;font-size:13px">
            <tr><td style="padding:6px 0;color:#888;width:80px">Date</td><td style="padding:6px 0"><strong>Saturday 28 June 2025</strong></td></tr>
            <tr><td style="padding:6px 0;color:#888">Time</td><td style="padding:6px 0"><strong>14:00 – 16:00</strong></td></tr>
          </table>
          <p style="margin:0 0 12px">Show this QR code at the door to check in:</p>
          <div style="text-align:center;margin:20px 0;padding:20px;background:#f5f8fa;border-radius:10px;border:1px solid #e0e8ed">
            <img src="${qrUrl}" width="200" height="200" alt="Your QR code" style="display:block;margin:0 auto 10px"/>
            <p style="color:#888;font-size:11px;margin:8px 0 0;font-family:monospace">${email}</p>
          </div>
          <div style="background:#e8f0fe;border-left:3px solid #004165;padding:10px 14px;margin:0 0 20px;font-size:13px;color:#004165">
            📱 Can't see the QR code? <a href="${qrUrl}" style="color:#004165">Click here to open it</a> on your phone.
          </div>
          <hr style="border:none;border-top:1px solid #eee;margin:24px 0"/>
          <p style="color:#888;font-size:12px;margin:0">Toastmasters Zimbabwe</p>
        </div>
      </div>
    `;

    const auth = Utilities.base64Encode(MAILJET_API_KEY + ':' + MAILJET_SECRET_KEY);
    const response = UrlFetchApp.fetch('https://api.mailjet.com/v3.1/send', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Basic ' + auth },
      payload: JSON.stringify({
        Messages: [{
          From:     { Email: FROM_EMAIL, Name: FROM_NAME },
          To:       [{ Email: email, Name: fullName }],
          Subject:  subject,
          TextPart: textBody,
          HTMLPart: htmlBody,
        }]
      }),
      muteHttpExceptions: true,
    });

    const respCode = response.getResponseCode();
    const respBody = response.getContentText();
    Logger.log(`Mailjet response for ${email}: ${respCode} — ${respBody}`);
    if (respCode !== 200) {
      Logger.log(`FAILED for ${email}`);
      continue;
    }

    sheet.getRange(i + 1, sentColIdx + 1).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    );
    sent++;
    Utilities.sleep(300);
  }

  Logger.log(`QR emails sent: ${sent}`);
  SpreadsheetApp.getUi().alert(`Done! Sent ${sent} QR code email(s).`);
}
