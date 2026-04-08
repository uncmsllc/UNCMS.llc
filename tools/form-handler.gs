/**
 * UNCMS LLC — Contact Form + Review Handler
 * Deploy this as a Google Apps Script Web App.
 *
 * REQUIRED SETUP — run once in the Apps Script editor:
 *   1. Go to Project Settings → Script Properties and add:
 *        GITHUB_TOKEN  →  a GitHub Personal Access Token with "repo" scope
 *                         (create at github.com → Settings → Developer settings
 *                          → Personal access tokens → Tokens (classic))
 *   2. Run testSetup() to authorize MailApp / Calendar and verify everything works.
 *   3. Deploy → New deployment (Web app, Execute as: Me, Access: Anyone).
 *   4. Paste the new /exec URL into index.html as APPS_SCRIPT_URL and push.
 */

// ── CONFIG ────────────────────────────────────────────────────────────────
const NOTIFICATION_EMAIL        = 'info@uncms.org';
const GITHUB_OWNER              = 'uncmsllc';
const GITHUB_REPO               = 'UNCMS.llc';
const GITHUB_FILE               = 'index.html';
const GITHUB_BRANCH             = 'main';
const FOLLOWUP_DAYS_OUT         = 1;    // next business day
const FOLLOWUP_HOUR             = 9;    // 9:00 AM
const FOLLOWUP_DURATION_MINUTES = 30;
// ─────────────────────────────────────────────────────────────────────────


// ══════════════════════════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════════════════════════

/**
 * GET handler — processes review approval / decline links from email.
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.action === 'approve' && params.id) {
    return approveReview(params.id);
  }
  if (params.action === 'decline' && params.id) {
    return declineReview(params.id);
  }

  return jsonResponse({ status: 'UNCMS form handler is active.' });
}


/**
 * POST handler — routes contact form submissions and review submissions.
 */
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    Logger.log('doPost received: ' + raw);

    if (data.type === 'review') {
      return handleReviewSubmission(data);
    }

    return handleContactSubmission(data);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse({ success: false, error: err.message });
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  CONTACT FORM
// ══════════════════════════════════════════════════════════════════════════

function handleContactSubmission(data) {
  const submittedAt = new Date();

  const inquiryLabels = {
    self:   'Themselves',
    family: 'A Family Member',
    friend: 'A Friend',
    other:  'General Information',
    '':     'Not specified'
  };
  const inquiryLabel = inquiryLabels[data.inquiry] || data.inquiry || 'Not specified';

  // ── Calendar follow-up event ─────────────────────────────────────────
  const calendar   = CalendarApp.getDefaultCalendar();
  const eventStart = getFollowupDate(submittedAt, FOLLOWUP_DAYS_OUT);
  eventStart.setHours(FOLLOWUP_HOUR, 0, 0, 0);
  const eventEnd   = new Date(eventStart.getTime() + FOLLOWUP_DURATION_MINUTES * 60 * 1000);

  const eventTitle = 'Follow Up: ' + data.fname + ' ' + data.lname + ' — Inquiry';
  const eventDesc  = [
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

  calendar.createEvent(eventTitle, eventStart, eventEnd, { description: eventDesc });

  // ── Email notification ───────────────────────────────────────────────
  const subject = 'New Inquiry: ' + data.fname + ' ' + data.lname;
  const body    = [
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

  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
  Logger.log('Contact email sent to ' + NOTIFICATION_EMAIL);

  return jsonResponse({ success: true });
}


// ══════════════════════════════════════════════════════════════════════════
//  REVIEW SUBMISSION
// ══════════════════════════════════════════════════════════════════════════

function handleReviewSubmission(data) {
  const submittedAt = new Date();

  // Store review keyed by a unique ID so it can be retrieved on approval
  const reviewId = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(
    'review_' + reviewId,
    JSON.stringify({
      id:          reviewId,
      name:        data.name   || '',
      role:        data.role   || '',
      rating:      data.rating || 5,
      review:      data.review || '',
      submittedAt: submittedAt.toISOString()
    })
  );

  const scriptUrl  = ScriptApp.getService().getUrl();
  const approveUrl = scriptUrl + '?action=approve&id=' + reviewId;
  const declineUrl = scriptUrl + '?action=decline&id=' + reviewId;
  const stars      = '★'.repeat(data.rating || 5) + '☆'.repeat(5 - (data.rating || 5));

  const subject = '⭐ Review Pending Approval — ' + (data.name || 'Anonymous');

  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#222;">',
    '  <div style="background:#c0392b;padding:24px 28px;border-radius:8px 8px 0 0;">',
    '    <h2 style="margin:0;color:#fff;font-size:20px;">New Review Awaiting Approval</h2>',
    '    <p style="margin:6px 0 0;color:rgba(255,255,255,0.8);font-size:13px;">UNCMS Website — review submitted by a resident or family member</p>',
    '  </div>',
    '  <div style="background:#f9f9f9;padding:28px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;">',
    '    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">',
    '      <tr><td style="padding:8px 0;font-size:13px;color:#666;width:100px;"><strong>Name</strong></td><td style="padding:8px 0;font-size:15px;">' + escapeHtml(data.name || 'Not provided') + '</td></tr>',
    '      <tr><td style="padding:8px 0;font-size:13px;color:#666;"><strong>Role</strong></td><td style="padding:8px 0;font-size:15px;">' + escapeHtml(data.role || 'Not provided') + '</td></tr>',
    '      <tr><td style="padding:8px 0;font-size:13px;color:#666;"><strong>Rating</strong></td><td style="padding:8px 0;font-size:18px;color:#c0392b;">' + stars + '</td></tr>',
    '    </table>',
    '    <div style="background:#fff;border-left:4px solid #c0392b;padding:16px 20px;border-radius:4px;margin-bottom:28px;font-style:italic;font-size:15px;line-height:1.7;">',
    '      &ldquo;' + escapeHtml(data.review || '') + '&rdquo;',
    '    </div>',
    '    <p style="font-size:14px;font-weight:700;margin-bottom:14px;">Do you want to publish this review on the UNCMS website?</p>',
    '    <div style="display:flex;gap:12px;">',
    '      <a href="' + approveUrl + '" style="display:inline-block;background:#27ae60;color:#fff;text-decoration:none;padding:13px 28px;border-radius:6px;font-weight:700;font-size:14px;margin-right:12px;">✅ Approve &amp; Publish</a>',
    '      <a href="' + declineUrl + '" style="display:inline-block;background:#e0e0e0;color:#333;text-decoration:none;padding:13px 28px;border-radius:6px;font-weight:700;font-size:14px;">✗ Decline</a>',
    '    </div>',
    '    <p style="font-size:12px;color:#999;margin-top:20px;">Clicking Approve will instantly add this review to the UNCMS website and push it to GitHub for deployment. Clicking Decline will permanently discard it.</p>',
    '    <hr style="border:none;border-top:1px solid #e0e0e0;margin:20px 0;">',
    '    <p style="font-size:11px;color:#bbb;margin:0;">Submitted ' + submittedAt.toLocaleString('en-US', { timeZone: 'America/New_York' }) + ' ET</p>',
    '  </div>',
    '</div>'
  ].join('\n');

  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, '', { htmlBody: htmlBody });
  Logger.log('Review approval email sent to ' + NOTIFICATION_EMAIL + ' (id: ' + reviewId + ')');

  return jsonResponse({ success: true });
}


// ══════════════════════════════════════════════════════════════════════════
//  REVIEW APPROVAL / DECLINE
// ══════════════════════════════════════════════════════════════════════════

function approveReview(reviewId) {
  const props  = PropertiesService.getScriptProperties();
  const stored = props.getProperty('review_' + reviewId);

  if (!stored) {
    return htmlPage('Already Processed',
      '<h2 style="color:#e74c3c;">Review Not Found</h2>' +
      '<p>This review has already been processed or the link has expired.</p>'
    );
  }

  const review = JSON.parse(stored);

  // ── Build testimonial card HTML ──────────────────────────────────────
  const initial  = review.name ? review.name.trim().charAt(0).toUpperCase() : '?';
  const rating   = Math.min(Math.max(parseInt(review.rating) || 5, 1), 5);
  const starsHtml = Array(rating).fill('<span>★</span>').join('') +
                    Array(5 - rating).fill('<span style="opacity:0.3">★</span>').join('');

  const cardHtml =
    '\n        <div class="testimonial-card fade-up">\n' +
    '          <div class="testimonial-stars">\n' +
    '            ' + starsHtml + '\n' +
    '          </div>\n' +
    '          <div class="testimonial-quote-mark">\u201C</div>\n' +
    '          <p class="testimonial-text">\n' +
    '            ' + escapeHtml(review.review) + '\n' +
    '          </p>\n' +
    '          <div class="testimonial-author">\n' +
    '            <div class="testimonial-avatar">' + initial + '</div>\n' +
    '            <div>\n' +
    '              <div class="testimonial-name">' + escapeHtml(review.name) + '</div>\n' +
    '              <div class="testimonial-role">' + escapeHtml(review.role) + '</div>\n' +
    '            </div>\n' +
    '          </div>\n' +
    '        </div>\n';

  // ── GitHub API: fetch current index.html ─────────────────────────────
  const token = props.getProperty('GITHUB_TOKEN');
  if (!token) {
    return htmlPage('Configuration Error',
      '<h2 style="color:#e74c3c;">GITHUB_TOKEN not set</h2>' +
      '<p>Add your GitHub Personal Access Token to Script Properties as <code>GITHUB_TOKEN</code> and try again.</p>'
    );
  }

  const apiUrl = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO +
                 '/contents/' + GITHUB_FILE + '?ref=' + GITHUB_BRANCH;
  const headers = {
    'Authorization': 'token ' + token,
    'Accept': 'application/vnd.github.v3+json',
    'User-Agent': 'UNCMS-FormHandler'
  };

  let fileData, currentContent, fileSha;
  try {
    const getResp = UrlFetchApp.fetch(apiUrl, { headers: headers, muteHttpExceptions: true });
    if (getResp.getResponseCode() !== 200) {
      throw new Error('GitHub GET failed: ' + getResp.getResponseCode() + ' ' + getResp.getContentText());
    }
    fileData       = JSON.parse(getResp.getContentText());
    fileSha        = fileData.sha;
    currentContent = Utilities.newBlob(
      Utilities.base64Decode(fileData.content.replace(/[\r\n]/g, ''))
    ).getDataAsString();
  } catch (err) {
    Logger.log('GitHub fetch error: ' + err.message);
    return htmlPage('GitHub Error',
      '<h2 style="color:#e74c3c;">Could not fetch index.html from GitHub</h2>' +
      '<p>' + err.message + '</p>'
    );
  }

  // ── Inject card before the Review Links Bar marker ───────────────────
  const MARKER = '      <!-- Review Links Bar -->';
  if (currentContent.indexOf(MARKER) === -1) {
    return htmlPage('Insertion Error',
      '<h2 style="color:#e74c3c;">Could not find insertion point in index.html</h2>' +
      '<p>The <code>&lt;!-- Review Links Bar --&gt;</code> comment is missing. Please check the file.</p>'
    );
  }

  const updatedContent = currentContent.replace(MARKER, cardHtml + '\n' + MARKER);

  // ── GitHub API: commit updated file ─────────────────────────────────
  const encodedContent = Utilities.base64Encode(
    Utilities.newBlob(updatedContent, 'text/plain').getBytes()
  );

  const putPayload = JSON.stringify({
    message: 'Add approved review from ' + review.name + ' [auto]',
    content: encodedContent,
    sha:     fileSha,
    branch:  GITHUB_BRANCH
  });

  try {
    const putResp = UrlFetchApp.fetch(
      'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/' + GITHUB_FILE,
      {
        method:           'put',
        headers:          Object.assign({}, headers, { 'Content-Type': 'application/json' }),
        payload:          putPayload,
        muteHttpExceptions: true
      }
    );
    if (putResp.getResponseCode() !== 200 && putResp.getResponseCode() !== 201) {
      throw new Error('GitHub PUT failed: ' + putResp.getResponseCode() + ' ' + putResp.getContentText());
    }
  } catch (err) {
    Logger.log('GitHub commit error: ' + err.message);
    return htmlPage('Commit Error',
      '<h2 style="color:#e74c3c;">Could not push to GitHub</h2>' +
      '<p>' + err.message + '</p>'
    );
  }

  // ── Clean up stored review ───────────────────────────────────────────
  props.deleteProperty('review_' + reviewId);
  Logger.log('Review approved and pushed to GitHub: ' + review.name);

  // ── Notify owner ─────────────────────────────────────────────────────
  MailApp.sendEmail(
    NOTIFICATION_EMAIL,
    '✅ Review Published — ' + review.name,
    'The review from ' + review.name + ' has been approved, committed to GitHub, and is now deploying to the UNCMS website.'
  );

  return htmlPage('Review Published!',
    '<div style="text-align:center;padding:20px 0;">' +
    '<div style="font-size:64px;margin-bottom:16px;">✅</div>' +
    '<h2 style="color:#27ae60;margin-bottom:8px;">Review Published!</h2>' +
    '<p style="color:#555;font-size:15px;">The review from <strong>' + escapeHtml(review.name) + '</strong> has been committed to GitHub and is now deploying to the UNCMS website.</p>' +
    '</div>'
  );
}


function declineReview(reviewId) {
  const props  = PropertiesService.getScriptProperties();
  const stored = props.getProperty('review_' + reviewId);

  if (!stored) {
    return htmlPage('Already Processed',
      '<h2 style="color:#e74c3c;">Review Not Found</h2>' +
      '<p>This review has already been processed or the link has expired.</p>'
    );
  }

  const review = JSON.parse(stored);
  props.deleteProperty('review_' + reviewId);
  Logger.log('Review declined: ' + review.name);

  return htmlPage('Review Declined',
    '<div style="text-align:center;padding:20px 0;">' +
    '<div style="font-size:64px;margin-bottom:16px;">✗</div>' +
    '<h2 style="color:#555;margin-bottom:8px;">Review Declined</h2>' +
    '<p style="color:#777;font-size:15px;">The review from <strong>' + escapeHtml(review.name) + '</strong> has been discarded and will not appear on the website.</p>' +
    '</div>'
  );
}


// ══════════════════════════════════════════════════════════════════════════
//  TEST SETUP
// ══════════════════════════════════════════════════════════════════════════

/**
 * RUN THIS FIRST from the Apps Script editor to authorize MailApp + Calendar
 * and verify the GitHub token is working.
 */
function testSetup() {
  try {
    // ── Calendar ────────────────────────────────────────────────────────
    const calendar  = CalendarApp.getDefaultCalendar();
    const now       = new Date();
    const testStart = new Date(now.getTime() + 60 * 60 * 1000);
    const testEnd   = new Date(testStart.getTime() + 30 * 60 * 1000);
    calendar.createEvent('TEST — UNCMS Form Handler (safe to delete)', testStart, testEnd, {
      description: 'Verification event. Safe to delete.'
    });
    Logger.log('✅ Calendar: test event created');

    // ── Gmail ───────────────────────────────────────────────────────────
    MailApp.sendEmail(
      NOTIFICATION_EMAIL,
      'TEST — UNCMS Form Handler is connected',
      'This confirms the form handler is authorized and working.\nYou can ignore this message.'
    );
    Logger.log('✅ MailApp: test sent to ' + NOTIFICATION_EMAIL);

    // ── GitHub token ────────────────────────────────────────────────────
    const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
    if (!token) {
      Logger.log('⚠️  GITHUB_TOKEN not set in Script Properties. Add it before approving reviews.');
    } else {
      const resp = UrlFetchApp.fetch(
        'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO,
        {
          headers: {
            'Authorization': 'token ' + token,
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'UNCMS-FormHandler'
          },
          muteHttpExceptions: true
        }
      );
      if (resp.getResponseCode() === 200) {
        Logger.log('✅ GitHub token: valid — repo access confirmed');
      } else {
        Logger.log('❌ GitHub token: failed (' + resp.getResponseCode() + '). Check token permissions.');
      }
    }

    Logger.log('✅ Setup complete.');
  } catch (err) {
    Logger.log('❌ Error: ' + err.message);
    throw err;
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════════════

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

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function htmlPage(title, bodyHtml) {
  const html =
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>' + title + ' — UNCMS</title>' +
    '<style>body{font-family:Arial,sans-serif;max-width:520px;margin:60px auto;padding:0 20px;color:#333;}' +
    'code{background:#f0f0f0;padding:2px 6px;border-radius:4px;font-size:13px;}</style>' +
    '</head><body>' + bodyHtml + '</body></html>';
  return HtmlService.createHtmlOutput(html);
}
