/*****************************************************************
 * OPERA XML GRABBER  (v4 - private downloads)
 *
 * Pulls the manager / journal XML report attachments for one month
 * from a Gmail label, keeps the latest run of each report, zips them
 * and lists any day that is missing a file.
 *
 * WEB APP:     the page downloads the zip straight to the browser.
 *              Nothing is saved to or shared from Google Drive.
 * EDITOR USE:  downloadThisMonth()  downloadLastMonth()  downloadPickedMonth()
 *              (saves a PRIVATE zip in your Drive and logs its link)
 * SECURITY:    lockDownOldZips()  makes every file in DRIVE_DIR private
 * DIAGNOSTICS: listMyLabels()  debugSample()
 *****************************************************************/

/* ============ CONFIG - edit these ============ */
const LABEL      = 'Automated XML Reports - Night Audit';
const DRIVE_DIR  = 'Opera XML Exports';      // editor runs only

// Who may use the web app besides you (the owner). Leave empty for
// "only me". Example: ['frontdesk@yourhotel.com', 'gm@yourhotel.com']
const ALLOWED_USERS = [];

// used only by downloadPickedMonth()
const PICK_YEAR  = 2026;
const PICK_MONTH = 8;          // 1-12
/* ============================================= */

const REPORT_RE  = /^(manager|journal)(\d{8})_(\d+)\.xml$/i;
const PAGE       = 100;        // threads read per Gmail call


/*****************************************************************
 * Web app entry points
 *****************************************************************/
function doGet() {
  if (!viewerAllowed_()) {
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif">You do not have access to this tool.</p>')
      .setTitle('Opera XML Grabber');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Opera XML Grabber')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Called from the page. Returns the zip itself (base64), never a link.
function webGrab(y, m) {
  try {
    if (!viewerAllowed_()) return { ok: false, msg: 'You do not have access to this tool.' };
    const r = collect_(Number(y), Number(m));
    if (!r.ok) return r;
    const zip = Utilities.zip(r.blobs, r.fname);
    return {
      ok: true, msg: r.msg, gaps: r.gaps, fname: r.fname,
      data: Utilities.base64Encode(zip.getBytes())
    };
  } catch (err) {
    return { ok: false, msg: 'Error: ' + err.message };
  }
}

// Month list for the dropdown: current month back 11.
function webMonths() {
  if (!viewerAllowed_()) return [];
  const out = [];
  const d = new Date();
  d.setDate(1);
  for (let i = 0; i < 12; i++) {
    out.push({ y: d.getFullYear(), m: d.getMonth() + 1,
               label: monthName_(d.getMonth() + 1) + ' ' + d.getFullYear() });
    d.setMonth(d.getMonth() - 1);
  }
  return out;
}

// The owner always gets in. Anyone else must be signed in and listed in
// ALLOWED_USERS. Google hides the viewer's address from scripts in some
// setups (for example a personal Gmail account viewing someone else's
// web app); an unknown viewer is refused, so a leaked URL is useless.
function viewerAllowed_() {
  const viewer = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner  = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!viewer) return false;
  if (viewer === owner) return true;
  return ALLOWED_USERS.map(function (e) { return e.toLowerCase(); }).indexOf(viewer) !== -1;
}


/*****************************************************************
 * Editor entry points
 *****************************************************************/
function downloadThisMonth() {
  const d = new Date();
  saveAndLog_(d.getFullYear(), d.getMonth() + 1);
}

function downloadLastMonth() {
  const d = new Date();
  d.setDate(1);                // avoid the 31st-rolls-over bug
  d.setMonth(d.getMonth() - 1);
  saveAndLog_(d.getFullYear(), d.getMonth() + 1);
}

function downloadPickedMonth() {
  saveAndLog_(PICK_YEAR, PICK_MONTH);
}

// Editor runs save the zip to your Drive with default (private) sharing:
// only you can open the link.
function saveAndLog_(y, m) {
  const r = collect_(y, m);
  Logger.log('');
  Logger.log(r.msg);
  if (r.gaps && r.gaps.length) {
    Logger.log('Gaps:');
    r.gaps.forEach(function (g) { Logger.log('   ' + g); });
  }
  if (!r.ok) return;

  const dir = folder_(DRIVE_DIR);
  const old = dir.getFilesByName(r.fname);
  while (old.hasNext()) old.next().setTrashed(true);
  const file = dir.createFile(Utilities.zip(r.blobs, r.fname));
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

  Logger.log('');
  Logger.log('DOWNLOAD (only you can open this) >>> ' +
             'https://drive.google.com/uc?export=download&id=' + file.getId());
}

// One-off clean-up for zips that older versions shared by link.
function lockDownOldZips() {
  const it = DriveApp.getFoldersByName(DRIVE_DIR);
  let n = 0;
  while (it.hasNext()) {
    const files = it.next().getFiles();
    while (files.hasNext()) {
      files.next().setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      n++;
    }
  }
  Logger.log(n + ' file(s) in "' + DRIVE_DIR + '" are now private.');
}


/*****************************************************************
 * Core: find the month's reports
 *****************************************************************/
function collect_(y, m) {
  const tag = '' + y + pad_(m);          // e.g. "202608"

  // The 5am email carries the PREVIOUS day's business date, so the
  // email window runs from the 1st into the first days of the next
  // month. The FILENAME date is the authority, not the email date.
  const from = new Date(y, m - 1, 1, 0, 0, 0);
  const to   = new Date(y, m, 3, 23, 59, 59);

  const label = GmailApp.getUserLabelByName(LABEL);
  if (!label) return { ok: false, msg: 'Label not found: ' + LABEL };

  const best = {};                       // "manager|20260811" -> {job, blob}

  // Threads come newest first. Page through them until every thread on a
  // page was last active before the window opened, so older months are
  // found no matter how many emails the label holds.
  for (let start = 0; ; start += PAGE) {
    const threads = label.getThreads(start, PAGE);
    if (!threads.length) break;
    threads.forEach(function (th) {
      if (th.getLastMessageDate() < from) return;
      th.getMessages().forEach(function (msg) {
        const d = msg.getDate();
        if (d < from || d > to) return;  // cheap metadata check, no blob fetch
        msg.getAttachments().forEach(function (att) {
          const mt = REPORT_RE.exec(att.getName());
          if (!mt || mt[2].substring(0, 6) !== tag) return;   // wrong report month
          const key = mt[1].toLowerCase() + '|' + mt[2];
          const job = parseInt(mt[3], 10);
          if (!best[key] || job > best[key].job) {
            // higher job number = later run = the good one
            best[key] = { job: job, blob: att.copyBlob().setName(att.getName()) };
          }
        });
      });
    });
    const oldest = threads[threads.length - 1].getLastMessageDate();
    if (oldest < from || threads.length < PAGE) break;
  }

  const keys = Object.keys(best).sort();
  if (!keys.length) {
    return { ok: false, msg: 'No reports found for ' + monthName_(m) + ' ' + y + '.' };
  }

  /* ---- gap report ---- */
  const now      = new Date();
  const isCurr   = (y === now.getFullYear() && m === now.getMonth() + 1);
  const lastDay  = isCurr ? now.getDate() - 1 : new Date(y, m, 0).getDate();
  const gaps     = [];
  let complete   = 0;

  for (let d = 1; d <= lastDay; d++) {
    const ds = tag + pad_(d);
    const hasM = !!best['manager|' + ds];
    const hasJ = !!best['journal|' + ds];
    if (hasM && hasJ)    complete++;
    else if (!hasM && !hasJ) gaps.push(monthName_(m) + ' ' + d + ' - both files missing');
    else if (!hasM)      gaps.push(monthName_(m) + ' ' + d + ' - manager missing');
    else                 gaps.push(monthName_(m) + ' ' + d + ' - journal missing');
  }

  return {
    ok:    true,
    msg:   keys.length + ' files (' + complete + ' complete day' + (complete === 1 ? '' : 's') +
           ') from ' + monthName_(m) + ' ' + y,
    fname: 'OperaXML_' + y + '-' + pad_(m) + '.zip',
    blobs: keys.map(function (k) { return best[k].blob; }),
    gaps:  gaps
  };
}


/*****************************************************************
 * Helpers
 *****************************************************************/
function pad_(n) {
  return ('0' + n).slice(-2);
}

function monthName_(m) {
  return ['January','February','March','April','May','June','July',
          'August','September','October','November','December'][m - 1];
}

function folder_(path) {
  const parts = path.split('/');
  let cur = DriveApp.getRootFolder();
  parts.forEach(function (p) {
    const it = cur.getFoldersByName(p);
    cur = it.hasNext() ? it.next() : cur.createFolder(p);
  });
  return cur;
}


/*****************************************************************
 * Diagnostics
 *****************************************************************/
function listMyLabels() {
  GmailApp.getUserLabels().forEach(function (l) {
    Logger.log('"' + l.getName() + '"');
  });
}

// Prints the newest 5 threads on the label with their real filenames
// and whether each one matches the expected pattern.
function debugSample() {
  const label = GmailApp.getUserLabelByName(LABEL);
  if (!label) { Logger.log('LABEL NOT FOUND: ' + LABEL); return; }
  const threads = label.getThreads(0, 5);
  Logger.log('Sampling ' + threads.length + ' newest threads');
  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      Logger.log('--- ' + msg.getDate() + '  |  ' + msg.getSubject());
      const atts = msg.getAttachments();
      if (!atts.length) Logger.log('      (no attachments)');
      atts.forEach(function (att) {
        Logger.log('      ' + att.getName() +
                   (REPORT_RE.test(att.getName()) ? '   [MATCH]' : '   [NO MATCH]'));
      });
    });
  });
}
