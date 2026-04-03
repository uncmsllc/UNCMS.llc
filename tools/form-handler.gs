/**
 * UNCMS LLC — Contact Form Handler
 * Deploy this as a Google Apps Script Web App.
 */

// ── CONFIG ────────────────────────────────────────────────────────────────
const NOTIFICATION_EMAIL       = 'info@uncms.com';
const FOLLOWUP_DAYS_OUT        = 1;   // next business day
const FOLLOWUP_HOUR            = 9;   // 9:00 AM
const FOLLOWUP_DURATION_MINUTES = 30;
// ─────────────────────────────────────────────────────────────────────────


/**
 * RUN THIS FIRST from the Apps Script editor (select testSetup → Run).
 * It authorizes Gmail + Calendar and sends a real test email + creates a
 * test calendar event so you can confirm everything is wired up correctly.
 */
function testSetup() {
  try {
    // ── Test Calendar ───────────────────────────────────────────────────
    const calendar  = CalendarApp.getDefaultCalendar();
    const now       = new Date();
    const testStart = new Date(now.getTime() + 60 * 60 * 1000); // 1 hour from now
    const testEnd   = new Date(testStart.getTime() + 30 * 60 * 1000);

    calendar.createEvent(
      'TEST — UNCMS Form Handler (safe to delete)',
      testStart,
      testEnd,
      { description: 'This event was created to verify the UNCMS form handler can write to Google Calendar. Safe to delete.' }
    );
    Logger.log('✅ Calendar: test event created successfully');

    // ── Test Email ──────────────────────────────────────────────────────
    GmailApp.sendEmail(
      NOTIFICATION_EMAIL,
      'TEST — UNCMS Form Handler is connected',
      'This is a test email confirming the UNCMS contact form is successfully linked to Google Workspace.\n\nYou can ignore this message.'
    );
    Logger.log('✅ Email: test message sent to ' + NOTIFICATION_EMAIL);

    Logger.log('✅ Setup complete. Both Calendar and Gmail are authorized and working.');

  } catch (err) {
    Logger.log('❌ Error during testSetup: ' + err.message);
    throw err;
  }
}


/**
 * Handles POST requests from the website contact form.
 */
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    Logger.log('doPost received: ' + raw);

    const submittedAt = new Date();

    const inquiryLabels = {
      self:   'Themselves',
      family: 'A Family Member',
      friend: 'A Friend',
      other:  'General Information',
      '':     'Not specified'
    };
    const inquiryLabel = inquiryLabels[data.inquiry] || data.inquiry || 'Not specified';

    // ── 1. Calendar event ────────────────────────────────────────────────
    const calendar   = CalendarApp.getDefaultCalendar();
    const eventStart = getFollowupDate(submittedAt, FOLLOWUP_DAYS_OUT);
    eventStart.setHours(FOLLOWUP_HOUR, 0, 0, 0);
    const eventEnd = new Date(eventStart.getTime() + FOLLOWUP_DURATION_MINUTES * 60 * 1000);

    const eventTitle = 'Follow Up: ' + data.fname + ' ' + data.lname + ' — Inquiry';
    const eventDescription = [
      'NEW INQUIRY — UNCMS Website',
      '',
      'Name:           ' + data.fname + ' ' + data.lname,
      'Phone:          ' + (data.phone   || 'Not provided'),
      'Email:          ' + (data.email   || 'Not provided'),
      'Inquiring for:  ' + inquiryLabel,
      '',
      'Message:',
      data.message || '(No message provided)',
      '',
      'Submitted: ' + submittedAt.toLocaleString('en-US', { timeZone: 'America/New_York' }) + ' ET'
    ].join('\n');

    calendar.createEvent(eventTitle, eventStart, eventEnd, { description: eventDescription });
    Logger.log('Calendar event created: ' + eventTitle);

    // ── 2. Email notification ────────────────────────────────────────────
    const emailSubject = 'New Inquiry: ' + data.fname + ' ' + data.lname;
    const emailBody = [
      'A new inquiry was submitted through the UNCMS website.',
      '',
      'Name:          ' + data.fname + ' ' + data.lname,
      'Phone:         ' + (data.phone   || 'Not provided'),
      'Email:         ' + (data.email   || 'Not provided'),
      'Inquiring for: ' + inquiryLabel,
      '',
      'Message:',
      data.message || '(No message provided)',
      '',
      'Submitted: ' + submittedAt.toLocaleString('en-US', { timeZone: 'America/New_York' }) + ' ET',
      '',
      '---',
      'A follow-up reminder has been added to your Google Calendar.'
    ].join('\n');

    GmailApp.sendEmail(NOTIFICATION_EMAIL, emailSubject, emailBody);
    Logger.log('Email sent to ' + NOTIFICATION_EMAIL);

    return jsonResponse({ success: true });

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse({ success: false, error: err.message });
  }
}


/**
 * Returns N business days after the given date (skips Sat/Sun).
 */
function getFollowupDate(fromDate, businessDaysOut) {
  const result = new Date(fromDate);
  let added = 0;
  while (added < businessDaysOut) {
    result.setDate(result.getDate() + 1);
    const day = result.getDay();
    if (day !== 0 && day !== 6) added++;
  }
  return result;
}


/**
 * JSON response helper.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Health-check — open the /exec URL in a browser to confirm the script is live.
 */
function doGet(e) {
  return jsonResponse({ status: 'UNCMS form handler is active.' });
}


// ══════════════════════════════════════════════════════════════════════════
//  FIX — HOW TO AUTHORIZE & REDEPLOY
// ══════════════════════════════════════════════════════════════════════════
//
//  The script must be authorized to use Gmail and Calendar before it will
//  work from the live website. Follow these steps once:
//
//  1. Open https://script.google.com → open the "UNCMS Form Handler" project.
//
//  2. Replace ALL the code with the contents of this file and save (Cmd+S).
//
//  3. In the function dropdown (top toolbar), select "testSetup".
//     Click Run (▶). You will be prompted to review and grant permissions —
//     click "Allow". This authorizes Gmail and Calendar.
//
//  4. Check the Execution Log (View → Logs). You should see:
//       ✅ Calendar: test event created successfully
//       ✅ Email: test message sent to info@uncms.com
//       ✅ Setup complete.
//     Also check your Google Calendar and info@uncms.com inbox to confirm.
//
//  5. Deploy → New deployment
//     - Type: Web app
//     - Execute as: Me
//     - Who has access: Anyone
//     Click Deploy. Copy the new /exec URL.
//
//  6. Update APPS_SCRIPT_URL in index.html with the new URL, then push to GitHub.
//
// ══════════════════════════════════════════════════════════════════════════
