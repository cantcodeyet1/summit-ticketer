const https  = require('https');
const fs     = require('fs');
const XLSX   = require('C:/Users/Nyash/Downloads/node_modules/xlsx');

// ── Config — paste FRESH keys from Mailjet ─────────────────────────
const MAILJET_API_KEY    = '1ac422847c9c4ba17c62c054a9c7b65d';
const MAILJET_SECRET_KEY = 'e3d36943a5db572afce6f440bb843351';
const FROM_EMAIL  = 'divisionc@toastmasters129.org';
const FROM_NAME   = 'Lawyers Toastmasters Club';
const LOGO_URL    = 'https://cantcodeyet1.github.io/summit-ticketer/ToastmastersLogo.png';
const FILE        = 'C:/Users/Nyash/Downloads/Lawyers Toastmasters Club RSVP (Responses).xlsx';

// ── Test mode: set to your email to send only to yourself, leave '' to send to all ──
const TEST_EMAIL  = '';

// ── Read sheet ─────────────────────────────────────────────────────
const wb      = XLSX.readFile(FILE);
const ws      = wb.Sheets[wb.SheetNames[0]];
const rows    = XLSX.utils.sheet_to_json(ws, { defval: '' });

const auth = Buffer.from(`${MAILJET_API_KEY}:${MAILJET_SECRET_KEY}`).toString('base64');

function post(payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const req  = https.request({
      hostname: 'api.mailjet.com',
      path:     '/v3.1/send',
      method:   'POST',
      headers:  {
        'Authorization': `Basic ${auth}`,
        'Content-Type':  'application/json',
        'Content-Length': Buffer.byteLength(body),
      }
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

(async () => {
  let sent = 0, skipped = 0, failed = 0;

  for (const row of rows) {
    const email    = String(row['Email address'] || '').trim();
    const fullname = String(row['Full Name'] || '').trim();
    const first    = fullname.split(' ')[0];
    const qrSent   = row['QR Sent'];

    if (!email || !email.includes('@')) continue;
    if (!TEST_EMAIL && qrSent) { skipped++; continue; }

    // In test mode: only process the first row, redirect delivery to TEST_EMAIL
    const toEmail = TEST_EMAIL || email;
    if (TEST_EMAIL && sent + failed > 0) break;

    const qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(email)}&size=250&margin=2`;

    const textBody =
      `Hi ${first},\n\n` +
      `You're registered for the Lawyers Toastmasters Club Launch!\n\n` +
      `Date: Saturday 27 June 2026, 09:00 AM\n\n` +
      `Your QR code: ${qrUrl}\n\n` +
      `Show it at the door to check in.\n\nSee you there!\nToastmasters Zimbabwe`;

    const htmlBody = `
<div style="font-family:Arial,sans-serif;max-width:500px;margin:auto;color:#1a1a1c">
  <div style="background:#ffffff;padding:28px 28px 20px;text-align:center;border-radius:12px 12px 0 0;border-bottom:3px solid #006094">
    <img src="${LOGO_URL}" width="110" alt="Toastmasters Zimbabwe" style="display:block;margin:0 auto 18px"/>
    <div style="color:#004165;font-size:26px;font-weight:700;letter-spacing:1px;margin-bottom:6px">Lawyers</div>
    <div style="color:#004165;font-size:22px;font-weight:400;margin-bottom:6px">Toastmasters Club</div>
    <div style="display:inline-block;background:linear-gradient(135deg,#004165,#006094);color:#ffffff;padding:6px 22px;border-radius:4px;font-weight:700;font-size:12px;letter-spacing:3px;margin-top:10px">LAUNCH PASS</div>
  </div>
  <div style="background:#fff;border:1px solid #e0e8ed;border-top:none;padding:28px;border-radius:0 0 12px 12px">
    <p style="margin:0 0 16px">Hi <strong>${first}</strong>,</p>
    <p style="margin:0 0 16px">You're all set for the <strong>Lawyers Toastmasters Club Launch</strong>.</p>
    <table style="width:100%;border-collapse:collapse;margin:0 0 24px;font-size:13px">
      <tr><td style="padding:6px 0;color:#888;width:80px">Date</td><td style="padding:6px 0"><strong>Saturday 27 June 2026</strong></td></tr>
      <tr><td style="padding:6px 0;color:#888">Time</td><td style="padding:6px 0"><strong>09:00 AM</strong></td></tr>
    </table>
    <p style="margin:0 0 12px">Show this QR code at the door to check in:</p>
    <div style="text-align:center;margin:20px 0;padding:20px;background:#f5f8fa;border-radius:10px;border:1px solid #e0e8ed">
      <img src="${qrUrl}" width="200" height="200" alt="Your QR code" style="display:block;margin:0 auto 10px"/>
      <p style="color:#888;font-size:11px;margin:8px 0 0;font-family:monospace">${email}</p>
    </div>
    <div style="background:#e8f0fe;border-left:3px solid #004165;padding:10px 14px;margin:0 0 20px;font-size:13px;color:#004165">
      Can't see the QR code? <a href="${qrUrl}" style="color:#004165">Click here to open it</a> on your phone.
    </div>
    <hr style="border:none;border-top:1px solid #eee;margin:24px 0"/>
    <p style="color:#888;font-size:12px;margin:0">Toastmasters Zimbabwe</p>
  </div>
</div>`;

    try {
      const res = await post({
        Messages: [{
          From:     { Email: FROM_EMAIL, Name: FROM_NAME },
          To:       [{ Email: toEmail, Name: fullname }],
          Subject:  'Your Launch Pass - Lawyers Toastmasters Club',
          TextPart: textBody,
          HTMLPart: htmlBody,
        }]
      });

      if (res.status === 200) {
        console.log(`  SENT  ${toEmail}${TEST_EMAIL ? ` (test — row email: ${email})` : ''}`);
        row['QR Sent'] = new Date().toISOString().replace('T', ' ').slice(0, 19);
        sent++;
      } else {
        console.log(`  FAIL  ${email}: ${res.status} — ${res.body.slice(0, 120)}`);
        failed++;
      }
    } catch (e) {
      console.log(`  ERROR ${email}: ${e.message}`);
      failed++;
    }

    await sleep(300);
  }

  // Write QR Sent timestamps back to the sheet
  const updated = XLSX.utils.json_to_sheet(rows);
  wb.Sheets[wb.SheetNames[0]] = updated;
  XLSX.writeFile(wb, FILE);

  console.log(`\nDone — sent: ${sent}  skipped: ${skipped}  failed: ${failed}`);
  console.log('Sheet updated with QR Sent timestamps.');
})();
