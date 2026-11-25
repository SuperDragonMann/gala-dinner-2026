/**
 * Gala Dinner 2026 — Menu Selection API (Google Apps Script, bound to a Google Sheet)
 * Front-end pages (index.html / admin.html) call this web app with fetch().
 *
 * Sheets (created automatically on first use):
 *   Guests  – one row per guest (the master record you manage)
 *   Summary – live counts per main course, who has not chosen, dietary notes, admin link
 *   Log     – every submission / admin change, time-stamped
 */

const CONFIG = {
  // Public guest page (GitHub Pages). Personal links = PAGE_URL + '?g=CODE'
  PAGE_URL: 'https://superdragonmann.github.io/gala-dinner-2026/',
  // Optional last time guests can submit, e.g. '2026-11-30T23:59:00+08:00'. '' = no deadline.
  DEADLINE: '',
  EVENT_TITLE: 'Gala Dinner cum NEXX AI and Innovation Award Ceremony',
  EMAIL_SUBJECT: 'Gala Dinner 3 Dec 2026 — please choose your main course',
  REMINDER_SUBJECT: 'Reminder: choose your main course for the Gala Dinner (3 Dec)',
  SENDER_NAME: 'NEXX Innovation Alliance & HKSAII'
};

const MAINS = {
  A: 'A) French Roasted Rack of Lamb 法式烤羊架',
  B: 'B) U.S. Angus Ribeye Steak 10oz 美國安格斯肉眼',
  C: 'C) Pan-Seared Cod Fillet 香煎銀鱈魚柳',
  D: 'D) Vegetarian Truffle Mushroom Risotto 雜菌意大利飯 (V)'
};

const HEAD = ['Code', 'First Name', 'Surname', 'Email', 'Organisation', 'Main Course', 'Dietary / Allergy Notes',
  'Status', 'Source', 'Submitted At', 'Times Changed', 'Personal Link', 'Email Sent At', 'Last Updated By', 'Confirmation Emailed At'];

// Full dish names for the confirmation email
const DISH = {
  A: { zh: '法式烤羊架伴薯菜配紅酒汁', en: 'French Roasted Rack of Lamb with Vegetables and Red Wine Sauce' },
  B: { zh: '美國安格斯肉眼配紅酒汁 (10oz)', en: 'Grilled U.S. Angus Ribeye Steak with Red Wine Sauce (10oz)' },
  C: { zh: '香煎銀鱈魚柳配焦糖檸蜜汁', en: 'Pan-Seared Baked Cod Fillet with Caramel Lemon Glaze' },
  D: { zh: '普羅旺斯菜伴扒露荀大啡菇配黑松露醬燴雜菌意大利飯', en: 'Ratatouille with Grilled Asparagus and Portobello Mushroom and Black Truffle Mixed Mushroom Risotto (Vegetarian)' }
};
const RUNDOWN = [
  ['6:00 – 6:30pm', 'Reception Open', ''],
  ['6:30 – 6:35pm', 'Opening Remarks', 'Mann Yim, President, Hong Kong Society of Artificial Intelligence and Innovation (HKSAII)'],
  ['6:35 – 6:45pm', 'Keynote', 'Digital Policy Office (DPO) — speaker to be confirmed'],
  ['6:45 – 7:00pm', 'Fireside Chat: Navigating the AI Shift — Talent Development, Data Governance, and the Future of Business Models', 'Moderator: Mann Yim (HKSAII) · Guest Speaker: SK Cheong, Chairman of the Hong Kong Economic Journal and Vice Chairman of HK Television Entertainment (ViuTV)'],
  ['7:00 – 8:20pm', 'Gala Dinner', ''],
  ['8:20 – 8:35pm', '2nd Edition NEXX AI and Innovation Leader Awards', 'Awardees: CIOs / CTOs'],
  ['8:35 – 9:00pm', 'Group Photo and Networking', '']
];
const COL = {}; HEAD.forEach((h, i) => COL[h] = i);

/* ================= Web API ================= */

function doGet(e) { return route_((e && e.parameter) || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  return route_(Object.assign({}, (e && e.parameter) || {}, body));
}

function route_(p) {
  let out;
  try {
    ensureSheets_();
    switch (p.action) {
      case 'guest': out = getGuest_(p.code); break;
      case 'submit': out = submit_(p); break;
      case 'adminData': out = adminData_(p.key); break;
      case 'adminSave': out = adminSave_(p.key, p.guest || {}); break;
      case 'adminDelete': out = adminDelete_(p.key, p.code); break;
      default: out = { service: 'gala-menu', closed: isClosed_() };
    }
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getGuest_(code) {
  code = cleanCode_(code);
  const closed = isClosed_();
  if (!code) return { closed, guest: null };
  const rows = guestSheet_().getDataRange().getValues();
  const r = findRow_(rows, code, '');
  return { closed, guest: r > 0 ? rowToGuest_(rows[r]) : null };
}

function submit_(p) {
  if (isClosed_()) throw new Error('Menu selection has closed. Please contact the organiser.');
  const main = String(p.main || '').toUpperCase();
  if (!MAINS[main]) throw new Error('Please choose a main course (A, B, C or D).');
  const g = cleanGuest_(p);
  if (!g.firstName || !g.surname) throw new Error('Please enter your first name and surname.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(g.email)) throw new Error('Please enter a valid email address.');

  return withLock_(() => {
    const sh = guestSheet_();
    const rows = sh.getDataRange().getValues();
    const r = findRow_(rows, cleanCode_(p.code), g.email);
    const now = new Date();
    let row;
    if (r < 0) {
      row = blankRow_(uniqueCode_(rows));
      row[COL.Source] = 'Self sign-up (not on list)';
    } else {
      row = rows[r].slice();
      if (row[COL['Main Course']]) row[COL['Times Changed']] = (Number(row[COL['Times Changed']]) || 0) + 1;
    }
    row[COL['First Name']] = g.firstName; row[COL.Surname] = g.surname; row[COL.Email] = g.email;
    if (g.org) row[COL.Organisation] = g.org;
    row[COL['Main Course']] = MAINS[main];
    row[COL['Dietary / Allergy Notes']] = g.dietary;
    row[COL.Status] = 'Selected';
    row[COL['Submitted At']] = now;
    row[COL['Last Updated By']] = 'Guest';
    writeRow_(sh, r, row);
    log_('Guest ' + (r < 0 ? 'new' : 'update'), row);
    const rowNum = r < 0 ? sh.getLastRow() : r + 1;
    let emailSent = false;
    try {
      sendConfirmation_(row, r >= 0 && Number(row[COL['Times Changed']]) > 0);
      sh.getRange(rowNum, COL['Confirmation Emailed At'] + 1).setValue(new Date());
      emailSent = true;
    } catch (err) { console.warn('Confirmation email failed: ' + err); }
    return { guest: rowToGuest_(row), emailSent: emailSent };
  });
}

/* ================= Admin API ================= */

function adminData_(key) {
  checkKey_(key);
  const ss = SpreadsheetApp.getActive();
  const rows = guestSheet_().getDataRange().getValues().slice(1).filter(x => x[COL.Code] || x[COL['First Name']] || x[COL.Email]);
  return {
    guests: rows.map(rowToGuest_),
    mains: MAINS,
    sheetUrl: ss.getUrl(),
    deadline: CONFIG.DEADLINE,
    closed: isClosed_(),
    pageUrl: CONFIG.PAGE_URL
  };
}

function adminSave_(key, p) {
  checkKey_(key);
  const g = cleanGuest_(p);
  const main = String(p.main || '').toUpperCase();
  if (main && !MAINS[main]) throw new Error('Main course must be A, B, C, D or blank.');
  if (!g.firstName && !g.surname && !g.email) throw new Error('Enter at least a name or an email.');
  return withLock_(() => {
    const sh = guestSheet_();
    const rows = sh.getDataRange().getValues();
    const code = cleanCode_(p.code);
    let r = code ? findRow_(rows, code, '') : -1;
    if (r < 0 && !code && g.email) {
      const dup = findRow_(rows, '', g.email);
      if (dup > 0) throw new Error('A guest with this email already exists (' + rows[dup][COL.Code] + ').');
    }
    let row;
    if (r < 0) { row = blankRow_(code || uniqueCode_(rows)); row[COL.Source] = 'Added by admin'; }
    else row = rows[r].slice();
    const before = String(row[COL['Main Course']] || '');
    row[COL['First Name']] = g.firstName; row[COL.Surname] = g.surname; row[COL.Email] = g.email;
    row[COL.Organisation] = g.org; row[COL['Dietary / Allergy Notes']] = g.dietary;
    row[COL['Main Course']] = main ? MAINS[main] : '';
    row[COL.Status] = main ? 'Selected' : 'Pending';
    if (main && before !== MAINS[main]) {
      if (before) row[COL['Times Changed']] = (Number(row[COL['Times Changed']]) || 0) + 1;
      row[COL['Submitted At']] = new Date();
    }
    if (!main) row[COL['Submitted At']] = '';
    row[COL['Last Updated By']] = 'Admin ' + Utilities.formatDate(new Date(), 'Asia/Hong_Kong', 'd MMM HH:mm');
    writeRow_(sh, r, row);
    log_('Admin ' + (r < 0 ? 'add' : 'edit'), row);
    return { guest: rowToGuest_(row) };
  });
}

function adminDelete_(key, code) {
  checkKey_(key);
  code = cleanCode_(code);
  if (!code) throw new Error('Missing guest code.');
  return withLock_(() => {
    const sh = guestSheet_();
    const rows = sh.getDataRange().getValues();
    const r = findRow_(rows, code, '');
    if (r < 0) throw new Error('Guest not found.');
    log_('Admin delete', rows[r]);
    sh.deleteRow(r + 1);
    return { deleted: code };
  });
}

/* ================= Sheet menu ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Gala Menu')
    .addItem('Set up / repair sheets', 'setupSheets')
    .addItem('Generate codes & personal links', 'generateLinks')
    .addItem('Show admin page link', 'showAdminLink')
    .addSeparator()
    .addItem('Send menu email to guests not yet emailed', 'sendInvites')
    .addItem('Send reminder to guests who have not chosen', 'sendReminders')
    .addToUi();
}

function setupSheets() { ensureSheets_(); SpreadsheetApp.getActive().toast('Sheets are ready.'); }

function showAdminLink() {
  ensureSheets_();
  SpreadsheetApp.getUi().alert('Admin page (keep private):\n\n' + adminUrl_());
}

function generateLinks() {
  ensureSheets_();
  const sh = guestSheet_();
  const rows = sh.getDataRange().getValues();
  let n = 0;
  for (let r = 1; r < rows.length; r++) {
    if (!rows[r][COL['First Name']] && !rows[r][COL.Email]) continue;
    if (!rows[r][COL.Code]) { rows[r][COL.Code] = uniqueCode_(rows); n++; }
    rows[r][COL['Personal Link']] = linkFor_(rows[r][COL.Code]);
    if (!rows[r][COL.Status]) rows[r][COL.Status] = rows[r][COL['Main Course']] ? 'Selected' : 'Pending';
    if (!rows[r][COL.Source]) rows[r][COL.Source] = 'Invited';
    if (rows[r][COL['Times Changed']] === '') rows[r][COL['Times Changed']] = 0;
  }
  if (rows.length > 1) sh.getRange(1, 1, rows.length, HEAD.length).setValues(rows);
  SpreadsheetApp.getUi().alert(n + ' new code(s) created. Every guest now has a personal link in column L.');
}

function sendInvites() { sendMail_(false); }
function sendReminders() { sendMail_(true); }

function sendMail_(reminder) {
  const ui = SpreadsheetApp.getUi();
  const sh = guestSheet_();
  const rows = sh.getDataRange().getValues();
  const targets = [];
  for (let r = 1; r < rows.length; r++) {
    const x = rows[r];
    if (!x[COL.Email] || !x[COL['Personal Link']]) continue;
    if (reminder ? !x[COL['Main Course']] : !x[COL['Email Sent At']]) targets.push(r);
  }
  if (!targets.length) { ui.alert('Nobody to email. Run "Generate codes & personal links" first.'); return; }
  const quota = MailApp.getRemainingDailyQuota();
  if (ui.alert((reminder ? 'Send reminder' : 'Send menu email') + ' to ' + targets.length + ' guest(s)?\nEmails left today: ' + quota, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  let sent = 0;
  for (const r of targets) {
    if (sent >= quota) break;
    const x = rows[r];
    MailApp.sendEmail({ to: x[COL.Email], name: CONFIG.SENDER_NAME,
      subject: reminder ? CONFIG.REMINDER_SUBJECT : CONFIG.EMAIL_SUBJECT,
      htmlBody: emailHtml_(x[COL['First Name']], x[COL['Personal Link']], reminder) });
    sh.getRange(r + 1, COL['Email Sent At'] + 1).setValue(new Date());
    sent++;
  }
  ui.alert(sent + ' email(s) sent.' + (sent < targets.length ? ' Daily quota reached — run again tomorrow for the rest.' : ''));
}

function emailHtml_(first, link, reminder) {
  const esc = s => String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  return '<div style="font-family:Arial,sans-serif;font-size:15px;color:#2a1d17;max-width:560px">' +
    '<p>Dear ' + esc(first || 'Guest') + ',</p>' +
    (reminder ? '<p>We have not yet received your main-course choice for the Gala Dinner.</p>'
              : '<p>Thank you for joining us at the <b>' + esc(CONFIG.EVENT_TITLE) + '</b>.</p>') +
    '<p><b>Thursday 3 December 2026, 6:00–9:00pm</b><br>9/F, Bank of China (Hong Kong) Complex, City University of Hong Kong, Kowloon Tong</p>' +
    '<p>Please choose your main course with your personal link:</p>' +
    '<p><a href="' + esc(link) + '" style="background:#b2351f;color:#fff;padding:12px 22px;border-radius:6px;text-decoration:none;display:inline-block;font-weight:bold">Choose my menu 選擇主菜</a></p>' +
    '<p style="font-size:13px;color:#6b5a4e">Or copy this link: ' + esc(link) + '<br>The link is personal to you — reopen it any time to change your choice.</p>' +
    '<p>We look forward to seeing you.<br>' + esc(CONFIG.SENDER_NAME) + '</p></div>';
}

/* ================= Confirmation email ================= */

/** Run this once from the editor: it asks for e-mail permission and sends you a sample confirmation. */
function testConfirmationEmail() {
  const me = Session.getEffectiveUser().getEmail();
  const row = blankRow_('SAMPLE');
  row[COL['First Name']] = 'Test'; row[COL.Surname] = 'Guest'; row[COL.Email] = me;
  row[COL['Main Course']] = MAINS.B; row[COL['Dietary / Allergy Notes']] = 'Nut allergy (sample)';
  row[COL['Personal Link']] = linkFor_('SAMPLE');
  sendConfirmation_(row, false);
  Logger.log('Sample confirmation sent to ' + me);
}

function sendConfirmation_(row, updated) {
  const email = String(row[COL.Email] || '');
  if (!email) return;
  const k = String(row[COL['Main Course']] || '').charAt(0);
  MailApp.sendEmail({
    to: email,
    name: CONFIG.SENDER_NAME,
    subject: (updated ? 'Updated: ' : '') + 'Your Gala Dinner menu confirmation — 3 Dec 2026 · 晚宴菜單確認',
    htmlBody: confirmationHtml_(row, k, updated),
    attachments: [Utilities.newBlob(icsFor_(row, k), 'text/calendar', 'Gala-Dinner-3-Dec-2026.ics')]
  });
}

function confirmationHtml_(row, k, updated) {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const base = CONFIG.PAGE_URL.replace(/\/?$/, '/');
  const name = [row[COL['First Name']], row[COL.Surname]].filter(String).join(' ');
  const d = DISH[k] || { zh: '', en: '' };
  const diet = String(row[COL['Dietary / Allergy Notes']] || '');
  const link = String(row[COL['Personal Link']] || base);
  const gcal = 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=' + encodeURIComponent('Gala Dinner cum NEXX AI and Innovation Award Ceremony') +
    '&dates=20261203T100000Z/20261203T130000Z&location=' + encodeURIComponent('9/F, Bank of China (Hong Kong) Complex, City University of Hong Kong, Kowloon Tong') +
    '&details=' + encodeURIComponent('Your main course: ' + k + ') ' + d.en + '\nMenu link: ' + link);
  const red = '#b2351f', ink = '#2a1d17', soft = '#6b5a4e', cream = '#f6ead3';
  const course = (label, zh, en) => '<tr><td style="padding:10px 0;border-bottom:1px solid #eadcc2;text-align:center">' +
    '<div style="font-size:12px;color:#3e5a3b;letter-spacing:.04em">' + label + '</div><div style="font-size:15px;font-weight:bold;color:' + ink + '">' + zh + '</div><div style="font-size:13px;color:' + soft + '">' + en + '</div></td></tr>';
  const rd = RUNDOWN.map(x => '<tr><td style="padding:9px 12px 9px 0;vertical-align:top;white-space:nowrap;color:' + red + ';font-weight:bold;font-size:14px">' + esc(x[0]) +
    '</td><td style="padding:9px 0;vertical-align:top;font-size:14px;color:' + ink + '"><b>' + esc(x[1]) + '</b>' + (x[2] ? '<br><span style="color:' + soft + ';font-size:13px">' + esc(x[2]) + '</span>' : '') + '</td></tr>').join('');
  return '<div style="margin:0;padding:0;background:' + cream + '"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + cream + '"><tr><td align="center" style="padding:20px 10px">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#fffaf0;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;color:' + ink + '">' +
    '<tr><td><img src="' + base + 'assets/email-header.jpg" width="600" alt="Gala Dinner cum NEXX AI and Innovation Award Ceremony" style="display:block;width:100%;height:auto;border:0"></td></tr>' +
    '<tr><td style="background:' + red + ';color:#fff6e6;padding:14px 24px;font-size:14px;text-align:center"><b>Thursday 3 December 2026 · 6:00 – 9:00pm</b><br>9/F, Bank of China (Hong Kong) Complex, City University of Hong Kong, Kowloon Tong</td></tr>' +
    '<tr><td style="padding:28px 28px 8px"><p style="margin:0 0 6px;font-size:16px">Dear ' + esc(name || 'Guest') + ',</p>' +
    '<p style="margin:0;font-size:15px;line-height:1.55">' + (updated ? 'Your menu choice has been <b>updated</b>.' : 'Thank you — your menu choice is confirmed.') + ' Please keep this email for your record.<br><span style="color:' + soft + '">多謝！你的晚宴菜單選擇已確認，請保留此電郵作記錄。</span></p></td></tr>' +
    '<tr><td style="padding:18px 28px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:2px solid ' + red + ';border-radius:12px;background:#fff">' +
    '<tr><td width="170" style="padding:14px;vertical-align:middle"><img src="' + base + 'assets/dish-' + k + '.jpg" width="150" height="150" alt="" style="display:block;border-radius:10px;border:0"></td>' +
    '<td style="padding:14px 16px 14px 0;vertical-align:middle"><div style="font-size:12px;color:' + soft + '">Your main course 你的主菜</div>' +
    '<div style="font-size:30px;font-weight:bold;color:' + red + ';line-height:1.1">' + esc(k) + '</div>' +
    '<div style="font-size:16px;font-weight:bold;margin-top:4px">' + esc(d.zh) + '</div><div style="font-size:14px;color:' + soft + '">' + esc(d.en) + '</div>' +
    (diet ? '<div style="font-size:13px;margin-top:8px"><b>Dietary notes:</b> ' + esc(diet) + '</div>' : '') + '</td></tr></table></td></tr>' +
    '<tr><td style="padding:6px 28px 4px"><div style="font-size:18px;font-weight:bold;color:' + red + ';text-align:center;margin-bottom:4px">Your Gala Dinner Menu 晚宴菜單</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
    course('Salad 沙律', '牛油果蕃茄沙沙田園沙律', 'Avocado and Tomato Salsa Garden Salad') +
    course('Soup 湯', '周打魚湯', 'New England Clam Chowder') +
    course('Bread 麵包', '麵包及牛油', 'Bread &amp; Butter') +
    course('Main course 主菜 — your choice', esc(k + ') ' + d.zh), esc(d.en)) +
    course('Dessert 甜品', '草莓奶凍', 'Strawberry Panna Cotta') + '</table></td></tr>' +
    '<tr><td style="padding:22px 28px 6px"><div style="font-size:18px;font-weight:bold;color:' + red + ';margin-bottom:6px">Event Rundown 活動流程</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rd + '</table></td></tr>' +
    '<tr><td style="padding:22px 28px;text-align:center">' +
    '<a href="' + esc(gcal) + '" style="display:inline-block;margin:4px;background:' + red + ';color:#fff;text-decoration:none;padding:11px 18px;border-radius:8px;font-weight:bold;font-size:14px">Add to Google Calendar</a>' +
    '<a href="' + base + 'assets/gala-dinner-2026.ics" style="display:inline-block;margin:4px;background:#fff;color:' + red + ';border:2px solid ' + red + ';text-decoration:none;padding:9px 16px;border-radius:8px;font-weight:bold;font-size:14px">Apple / Outlook Calendar</a>' +
    '<div style="font-size:12px;color:' + soft + ';margin-top:8px">A calendar file (.ics) is also attached to this email.</div></td></tr>' +
    '<tr><td style="padding:0 28px 26px;text-align:center;font-size:13px;color:' + soft + '">Need to change your choice? Use your personal link:<br><a href="' + esc(link) + '" style="color:' + red + '">' + esc(link) + '</a></td></tr>' +
    '<tr><td style="background:' + red + ';color:#fff6e6;padding:14px 24px;font-size:12px;text-align:center">' + esc(CONFIG.SENDER_NAME) + '</td></tr>' +
    '</table></td></tr></table></div>';
}

function icsFor_(row, k) {
  const esc = s => String(s || '').replace(/\\/g, '\\\\').replace(/([,;])/g, '\\$1').replace(/\r?\n/g, '\\n');
  const d = DISH[k] || { en: '' };
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//NEXX Innovation Alliance & HKSAII//Gala Dinner 2026//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'BEGIN:VEVENT', 'UID:gala-2026-' + row[COL.Code] + '@nexx-hksaii.hk', 'DTSTAMP:' + stamp,
    'DTSTART:20261203T100000Z', 'DTEND:20261203T130000Z',
    'SUMMARY:' + esc('Gala Dinner cum NEXX AI and Innovation Award Ceremony'),
    'LOCATION:' + esc('9/F, Bank of China (Hong Kong) Complex, City University of Hong Kong, Kowloon Tong, Hong Kong'),
    'DESCRIPTION:' + esc('Your main course: ' + k + ') ' + d.en + '\nReception opens 6:00pm.\nYour menu link: ' + row[COL['Personal Link']]),
    'URL:' + row[COL['Personal Link']],
    'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Gala Dinner tomorrow', 'TRIGGER:-P1D', 'END:VALARM',
    'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Gala Dinner in 2 hours', 'TRIGGER:-PT2H', 'END:VALARM',
    'END:VEVENT', 'END:VCALENDAR'];
  return lines.map(foldIcs_).join('\r\n') + '\r\n';
}
function foldIcs_(line) {
  const out = []; let cur = '', len = 0;
  for (const ch of line) {
    const n = Utilities.newBlob(ch).getBytes().length;
    if (len + n > (out.length ? 74 : 75)) { out.push(cur); cur = ''; len = 0; }
    cur += ch; len += n;
  }
  out.push(cur);
  return out.join('\r\n ');
}

/* ================= Helpers ================= */

function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  let g = ss.getSheetByName('Guests');
  if (!g) {
    g = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName('Guests') : ss.insertSheet('Guests', 0);
    g.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#b2351f').setFontColor('#ffffff');
    g.setFrozenRows(1);
    g.setColumnWidths(1, HEAD.length, 140); g.setColumnWidth(6, 320); g.setColumnWidth(7, 240); g.setColumnWidth(12, 330);
  }
  if (g.getRange(1, HEAD.length).getValue() !== HEAD[HEAD.length - 1]) {
    g.getRange(1, HEAD.length).setValue(HEAD[HEAD.length - 1]).setFontWeight('bold').setBackground('#b2351f').setFontColor('#ffffff');
  }
  if (!ss.getSheetByName('Summary')) {
    const s = ss.insertSheet('Summary', 1);
    const data = [
      ['Main course', 'Count'],
      [MAINS.A, '=COUNTIF(Guests!F:F,"A)*")'],
      [MAINS.B, '=COUNTIF(Guests!F:F,"B)*")'],
      [MAINS.C, '=COUNTIF(Guests!F:F,"C)*")'],
      [MAINS.D, '=COUNTIF(Guests!F:F,"D)*")'],
      ['Total chosen', '=SUM(B2:B5)'],
      ['Guests in list', '=COUNTA(Guests!A2:A)'],
      ['Not yet chosen', '=COUNTIFS(Guests!A2:A,"<>",Guests!F2:F,"")'],
      ['', ''],
      ['Guests who have NOT chosen yet', '']
    ];
    s.getRange(1, 1, data.length, 2).setValues(data);
    s.getRange('A1:B1').setFontWeight('bold').setBackground('#b2351f').setFontColor('#ffffff');
    s.getRange('A6:B8').setFontWeight('bold'); s.getRange('A10').setFontWeight('bold');
    s.getRange('A11').setFormula('=IFERROR(FILTER({Guests!B2:B&" "&Guests!C2:C, Guests!D2:D}, Guests!A2:A<>"", Guests!F2:F=""), "Everyone has chosen")');
    s.getRange('E1').setValue('Dietary / allergy notes').setFontWeight('bold');
    s.getRange('E2').setFormula('=IFERROR(FILTER({Guests!B2:B&" "&Guests!C2:C, LEFT(Guests!F2:F,1), Guests!G2:G}, Guests!G2:G<>""), "None")');
    s.setColumnWidth(1, 380); s.setColumnWidth(5, 200); s.setColumnWidth(6, 60); s.setColumnWidth(7, 320);
  }
  if (!ss.getSheetByName('Log')) {
    const l = ss.insertSheet('Log');
    l.appendRow(['Time', 'Action', 'Code', 'First Name', 'Surname', 'Email', 'Main Course', 'Dietary / Allergy']);
    l.getRange('A1:H1').setFontWeight('bold'); l.setFrozenRows(1);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) {
    props.setProperty('ADMIN_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  }
  const s = ss.getSheetByName('Summary');
  if (s.getRange('I1').getValue() !== 'Admin page (private)') {
    s.getRange('I1').setValue('Admin page (private)').setFontWeight('bold');
    s.getRange('I2').setValue(adminUrl_());
  }
}

function adminUrl_() {
  return CONFIG.PAGE_URL.replace(/\/?$/, '/') + 'admin.html?k=' + PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
}
function checkKey_(key) {
  const k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!k || String(key || '') !== k) throw new Error('Admin key not valid.');
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function guestSheet_() { return SpreadsheetApp.getActive().getSheetByName('Guests'); }
function findRow_(rows, code, email) {
  for (let r = 1; r < rows.length; r++) {
    if (code && String(rows[r][COL.Code]).toUpperCase() === code) return r;
  }
  if (email) for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][COL.Email]).trim().toLowerCase() === email) return r;
  }
  return -1;
}
function writeRow_(sh, r, row) {
  if (r < 0) sh.appendRow(row); else sh.getRange(r + 1, 1, 1, HEAD.length).setValues([row]);
}
function blankRow_(code) {
  const row = HEAD.map(() => '');
  row[COL.Code] = code; row[COL['Times Changed']] = 0; row[COL.Status] = 'Pending';
  row[COL['Personal Link']] = linkFor_(code);
  return row;
}
function cleanGuest_(p) {
  const t = (v, n) => String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, n);
  return {
    firstName: t(p.firstName, 60), surname: t(p.surname, 60),
    email: t(p.email, 160).toLowerCase(), org: t(p.org, 160),
    dietary: String(p.dietary || '').trim().slice(0, 500)
  };
}
function log_(action, row) {
  SpreadsheetApp.getActive().getSheetByName('Log').appendRow([new Date(), action, row[COL.Code], row[COL['First Name']],
    row[COL.Surname], row[COL.Email], row[COL['Main Course']], row[COL['Dietary / Allergy Notes']]]);
}
function fmt_(d) { return d ? Utilities.formatDate(new Date(d), 'Asia/Hong_Kong', 'd MMM yyyy, h:mm a') : ''; }
function rowToGuest_(x) {
  const m = String(x[COL['Main Course']] || '');
  return {
    code: String(x[COL.Code] || ''), firstName: String(x[COL['First Name']] || ''), surname: String(x[COL.Surname] || ''),
    email: String(x[COL.Email] || ''), org: String(x[COL.Organisation] || ''), main: m ? m.charAt(0) : '',
    dietary: String(x[COL['Dietary / Allergy Notes']] || ''), status: String(x[COL.Status] || ''),
    source: String(x[COL.Source] || ''), submittedAt: fmt_(x[COL['Submitted At']]),
    changed: Number(x[COL['Times Changed']]) || 0, link: String(x[COL['Personal Link']] || ''),
    emailedAt: fmt_(x[COL['Email Sent At']]), updatedBy: String(x[COL['Last Updated By']] || '')
  };
}
function cleanCode_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 12); }
function uniqueCode_(rows) {
  const used = new Set(rows.map(x => String(x[COL.Code]).toUpperCase()));
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let c;
  do { c = ''; for (let i = 0; i < 6; i++) c += abc[Math.floor(Math.random() * abc.length)]; } while (used.has(c));
  return c;
}
function linkFor_(code) { return CONFIG.PAGE_URL.replace(/\/?$/, '/') + '?g=' + code; }
function isClosed_() { return !!CONFIG.DEADLINE && new Date() > new Date(CONFIG.DEADLINE); }
