/**
 * UNCMS LLC — Contact Form Handler
 * Deploy this as a Google Apps Script Web App.
 *
 * What it does:
 *   1. Receives form submissions from the website
 *   2. Creates a follow-up calendar event on your Google Calendar
 *   3. Sends you an email notification with the inquiry details
 *
 * Setup instructions are at the bottom of this file.
 */

// ── CONFIG ─────────────────────────────────────────────────────────────────
// The email address that receives inquiry notifications.
// Defaults to your Google account email. Change this if you want notifications
// sent to a different address (e.g., an admin inbox).
const NOTIFICATION_EMAIL = Session.getActiveUser().getEmail();

// How many business days out to schedule the follow-up reminder.
// 1 = next business day (recommended for intake inquiries)
const FOLLOWUP_DAYS_OUT = 1;

// What time the follow-up event appears on your calendar (24h format)
const FOLLOWUP_HOUR = 9;   // 9 = 9:00 AM
const FOLLOWUP_DURATION_MINUTES = 30;
// ───────────────────────────────────────────────────────────────────────────


/**
 * Handles POST requests from the website contact form.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const submittedAt = new Date();

    // Build a human-readable label for the inquiry type
    const inquiryLabels = {
      self:   'Themselves',
      family: 'A Family Member',
      friend: 'A Friend',
      other:  'General Information',
      '':     'Not specified'
    };
    const inquiryLabel = inquiryLabels[data.inquiry] || data.inquiry || 'Not specified';

    // ── 1. Create calendar follow-up event ───────────────────────────────
    const calendar = CalendarApp.getDefaultCalendar();
    const eventStart = getFollowupDate(submittedAt, FOLLOWUP_DAYS_OUT);
    eventStart.setHours(FOLLOWUP_HOUR, 0, 0, 0);
    const eventEnd = new Date(eventStart.getTime() + FOLLOWUP_DURATION_MINUTES * 60 * 1000);

    const eventTitle = `Follow Up: ${data.fname} ${data.lname} — Inquiry`;
    const eventDescription = [
      '📋 NEW INQUIRY — UNCMS Website',
      '',
      `Name:           ${data.fname} ${data.lname}`,
      `Phone:          ${data.phone  || 'Not provided'}`,
      `Email:          ${data.email  || 'Not provided'}`,
      `Inquiring for:  ${inquiryLabel}`,
      '',
      'Message:',
      data.message || '(No message provided)',
      '',
      `Submitted: ${submittedAt.toLocaleString('en-US', { timeZone: 'America/New_York' })} ET`
    ].join('\n');

    calendar.createEvent(eventTitle, eventStart, eventEnd, {
      description: eventDescription
    });

    // ── 2. Send email notification ────────────────────────────────────────
    const emailSubject = `New Inquiry: ${data.fname} ${data.lname}`;
    const emailBody = [
      'A new inquiry was submitted through the UNCMS website.',
      '',
      `Name:          ${data.fname} ${data.lname}`,
      `Phone:         ${data.phone  || 'Not provided'}`,
      `Email:         ${data.email  || 'Not provided'}`,
      `Inquiring for: ${inquiryLabel}`,
      '',
      'Message:',
      data.message || '(No message provided)',
      '',
      `Submitted: ${submittedAt.toLocaleString('en-US', { timeZone: 'America/New_York' })} ET`,
      '',
      '---',
      'A follow-up reminder has been added to your Google Calendar.'
    ].join('\n');

    GmailApp.sendEmail(NOTIFICATION_EMAIL, emailSubject, emailBody);

    // ── 3. Return success ─────────────────────────────────────────────────
    return jsonResponse({ success: true });

  } catch (err) {
    console.error('Form handler error:', err);
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
    if (day !== 0 && day !== 6) added++; // skip Sunday (0) and Saturday (6)
  }
  return result;
}


/**
 * Wraps a JSON object in a CORS-enabled text response.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Health-check endpoint — visit the deployed URL in a browser to confirm it's live.
 */
function doGet(e) {
  return jsonResponse({ status: 'UNCMS form handler is active.' });
}


// ══════════════════════════════════════════════════════════════════════════════
//  HOW TO DEPLOY THIS SCRIPT
// ══════════════════════════════════════════════════════════════════════════════
//
//  1. Go to https://script.google.com — sign in with your Google Workspace account.
//
//  2. Click "New project", paste the entire contents of this file, and save it
//     (Ctrl+S / Cmd+S). Name the project "UNCMS Form Handler".
//
//  3. Click "Deploy" → "New deployment".
//     - Type: Web app
//     - Execute as: Me (your Google account)
//     - Who has access: Anyone
//     Click "Deploy" and authorize the permissions when prompted.
//
//  4. Copy the "Web app URL" that appears (it looks like:
//     https://script.google.com/macros/s/XXXXXXXXXXXX/exec)
//
//  5. Open index.html and replace the placeholder:
//        const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
//     with your actual URL.
//
//  6. Test it: submit the contact form on the website. You should receive an
//     email and see a calendar event appear on your Google Calendar.
//
//  IMPORTANT: Each time you edit this script, you must create a NEW deployment
//  (Deploy → New deployment) — editing a deployment in place does not update
//  the live version.
// ══════════════════════════════════════════════════════════════════════════════
