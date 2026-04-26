/**
 * ============================================================
 * JOB PIPELINE DASHBOARD v3 — APPS SCRIPT WEB APP
 * ============================================================
 * Owner:        Louie Radburnd
 * Last updated: 23 April 2026
 * Spec:         job_pipeline_dashboard_v3_spec.md
 *
 * DEPLOYMENT
 * Deploy as Web App. Execute as: Me. Who has access: Anyone.
 * All POST requests use Content-Type: text/plain to bypass CORS
 * preflight. Requests and responses are JSON strings.
 *
 * ENDPOINTS
 * GET  ?action=ping
 * GET  ?action=getPipeline
 * GET  ?action=getRow&rowId=N
 * GET  ?action=getQueue
 * GET  ?action=getPrompt&type=T&rowId=N
 * POST {action: "addToQueue", content, url}
 * POST {action: "removeFromQueue", queueId}
 * POST {action: "clearQueue"}
 * POST {action: "batchAppendAnalysed", tsv}
 * POST {action: "savePhase1Docs", rowId, coverLetterLink, resumeLink}
 * POST {action: "markAsApplied", rowId, submittedResumeContent, submittedCoverLetterContent, appliedVia, applicationDate}
 * POST {action: "saveInterviewPrep", rowId, interviewPrepJson}
 * POST {action: "updateCardStatus", rowId, newStatus}
 * POST {action: "savePersonalAngle", rowId, personalAngle}
 * POST {action: "saveNote", rowId, noteText}
 * POST {action: "appendAnalysedJob", ...}  (legacy, kept for compat)
 * ============================================================
 */


// ============================================================
// CONFIGURATION
// ============================================================

const SHEET_ID      = '13cSuuItK8YfiNGa9xJDN3Eds4zOHVsUHtWQGyu0h53M';
const PIPELINE_TAB  = 'Job Pipeline v2';
const QUEUE_TAB     = 'Intake Queue';
const TOTAL_COLS    = 38;

const COL = {
  STATUS:                      1,
  COMPANY:                     2,
  ROLE:                        3,
  FIT_SCORE:                   4,
  RECOMMENDATION:              5,
  URL:                         6,
  KEY_NOTES:                   7,
  COMPANY_INTEL:               8,
  SALARY:                      9,
  SALARY_EXPECTATION:         10,
  LOCATION:                   11,
  WORK_ARRANGEMENT:           12,
  KEY_ALIGNMENTS:             13,
  POTENTIAL_CONCERNS:         14,
  APPLICATION_PRIORITY:       15,
  TAILORED_PITCH:             16,
  ANALYSIS_DATE:              17,
  CULTURAL_FIT:               18,
  RESUME_MD:                  19,
  COVER_LETTER_MD:            20,
  SUBMITTED_RESUME_LINK:      21,
  SUBMITTED_COVER_LETTER_LINK:22,
  INTERVIEW_PREP:             23,
  LINKEDIN_CONTACTS_LINK:     24,
  LAST_ACTIVITY:              25,
  YOUR_NOTES:                 26,
  DUPLICATE_CHECK:            27,
  APPLICATION_DATE:           28,
  APPLICATION_METHOD:         29,
  FOLLOW_UP_DATE:             30,
  DATE_ADDED:                 31,
  CONTACT_NAME:               32,
  CONTACT_EMAIL:              33,
  CONTACT_LINKEDIN:           34,
  PERSONAL_ANGLE:             35,
  AD_CONTENT:                 36,
  STAGE_ENTERED_AT:           37,
  STATUS_HISTORY:             38
};

const ACTIVE_STATUSES  = ['Sourced', 'Analysed', 'Ready to Apply', 'Applied'];
const HOLDING_STATUSES = ['On Hold', 'Unsure', 'Rethink'];
const ARCHIVE_STATUSES = ['Rejected', 'Withdrawn'];
const APPLIED_SUB      = ['Applied', 'Interviewing', 'Interviewed', 'Offer', 'Accepted'];


// ============================================================
// ROUTERS
// ============================================================

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';

    // No action = serve the HTML dashboard
    if (!action) {
      const template = HtmlService.createTemplateFromFile('index');
      template.scriptUrl = ScriptApp.getService().getUrl();
      return template.evaluate()
        .setTitle('Job Pipeline Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    switch (action) {
      case 'ping':         return jsonResponse({ ok: true, ts: now() });
      case 'getPipeline':  return jsonResponse(getPipeline());
      case 'getRow':       return jsonResponse(getRow(parseInt(e.parameter.rowId, 10)));
      case 'getQueue':     return jsonResponse(getQueue());
      case 'getPrompt':    return jsonResponse(getPrompt(e.parameter.type, e.parameter));
      default:             return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action || '';
    switch (action) {
      case 'addToQueue':            return jsonResponse(addToQueue(body));
      case 'removeFromQueue':       return jsonResponse(removeFromQueue(body));
      case 'clearQueue':            return jsonResponse(clearQueue());
      case 'batchAppendAnalysed':   return jsonResponse(batchAppendAnalysed(body));
      case 'savePhase1Docs':        return jsonResponse(savePhase1Docs(body));
      case 'markAsApplied':         return jsonResponse(markAsApplied(body));
      case 'saveInterviewPrep':     return jsonResponse(saveInterviewPrep(body));
      case 'updateRoundScore':      return jsonResponse(updateRoundScore(body));
      case 'updateCardStatus':      return jsonResponse(updateCardStatus(body));
      case 'savePersonalAngle':     return jsonResponse(savePersonalAngle(body));
      case 'saveNote':              return jsonResponse(saveNote(body));
      case 'updateNote':            return jsonResponse(updateNote(body));
      case 'deleteNote':            return jsonResponse(deleteNote(body));
      case 'saveAdContent':         return jsonResponse(saveAdContent(body));
      case 'saveStatusHistory':     return jsonResponse(saveStatusHistory(body));
      case 'batchUpdateJourney':    return jsonResponse(batchUpdateJourney(body));
      case 'appendAnalysedJob':     return jsonResponse(appendAnalysedJob(body));
      case 'batchSaveAnalyse':      return jsonResponse(batchSaveAnalyse(body));
      case 'batchSaveGenerateDocs': return jsonResponse(batchSaveGenerateDocs(body));
      case 'batchSaveInterviewPrep':return jsonResponse(batchSaveInterviewPrep(body));
      default:                      return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/**
 * apiRouter — single entrypoint called by the frontend via google.script.run.
 *
 * The HtmlService frontend lives in a googleusercontent.com sandbox which
 * can\'t fetch() the /exec URL on script.google.com due to CORS. Instead,
 * the frontend calls google.script.run.apiRouter({method, params/body}),
 * which Apps Script proxies directly to this function without crossing
 * origins.
 *
 * Returns a JSON STRING (not a plain object). This is intentional —
 * google.script.run\'s object serialiser silently returns null for
 * responses containing Date objects, functions, or certain nested
 * structures. Returning a string bypasses that landmine. Frontend
 * parses via JSON.parse.
 *
 * Mirrors the case switches in doGet and doPost so both behave identically
 * regardless of whether the call came via HTTP fetch or google.script.run.
 */
function apiRouter(request) {
  try {
    request = request || {};
    const method = request.method || 'GET';
    let result;

    if (method === 'GET') {
      const params = request.params || {};
      const action = params.action || '';
      switch (action) {
        case 'ping':         result = { ok: true, ts: now() }; break;
        case 'getPipeline':  result = getPipeline(); break;
        case 'getRow':       result = getRow(parseInt(params.rowId, 10)); break;
        case 'getQueue':     result = getQueue(); break;
        case 'getPrompt':    result = getPrompt(params.type, params); break;
        default:             result = { ok: false, error: 'Unknown GET action: ' + action };
      }
    } else {
      // POST path
      const body = request.body || {};
      const action = body.action || '';
      switch (action) {
        case 'addToQueue':            result = addToQueue(body); break;
        case 'removeFromQueue':       result = removeFromQueue(body); break;
        case 'clearQueue':            result = clearQueue(); break;
        case 'batchAppendAnalysed':   result = batchAppendAnalysed(body); break;
        case 'savePhase1Docs':        result = savePhase1Docs(body); break;
        case 'markAsApplied':         result = markAsApplied(body); break;
        case 'saveInterviewPrep':     result = saveInterviewPrep(body); break;
        case 'updateRoundScore':      result = updateRoundScore(body); break;
        case 'updateCardStatus':      result = updateCardStatus(body); break;
        case 'savePersonalAngle':     result = savePersonalAngle(body); break;
        case 'saveNote':              result = saveNote(body); break;
        case 'updateNote':            result = updateNote(body); break;
        case 'deleteNote':            result = deleteNote(body); break;
        case 'saveAdContent':         result = saveAdContent(body); break;
        case 'saveStatusHistory':     result = saveStatusHistory(body); break;
        case 'batchUpdateJourney':    result = batchUpdateJourney(body); break;
        case 'appendAnalysedJob':     result = appendAnalysedJob(body); break;
        case 'batchSaveAnalyse':      result = batchSaveAnalyse(body); break;
        case 'batchSaveGenerateDocs': result = batchSaveGenerateDocs(body); break;
        case 'batchSaveInterviewPrep':result = batchSaveInterviewPrep(body); break;
        default:                      result = { ok: false, error: 'Unknown POST action: ' + action };
      }
    }

    // Stringify to sidestep google.script.run\'s brittle object serialiser.
    // Date objects, nested undefined values, and certain shapes cause it
    // to silently return null on the client. JSON strings pass through
    // cleanly every time.
    return JSON.stringify(result);
  } catch (err) {
    return JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


// ============================================================
// PIPELINE READS
// ============================================================

function getPipeline() {
  const sheet = openPipeline();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
  const rows = values.map(function (r, i) { return rowToObject(r, i + 2); });
  return { ok: true, rows: rows };
}

function getRow(rowId) {
  if (!rowId || isNaN(rowId)) return { ok: false, error: 'Invalid rowId' };
  const sheet = openPipeline();
  const values = sheet.getRange(rowId, 1, 1, TOTAL_COLS).getValues()[0];
  return { ok: true, row: rowToObject(values, rowId) };
}

function rowToObject(r, rowId) {
  return {
    rowId:                       rowId,
    status:                      r[COL.STATUS - 1],
    company:                     r[COL.COMPANY - 1],
    role:                        r[COL.ROLE - 1],
    fitScore:                    r[COL.FIT_SCORE - 1],
    recommendation:              r[COL.RECOMMENDATION - 1],
    url:                         r[COL.URL - 1],
    keyNotes:                    r[COL.KEY_NOTES - 1],
    companyIntel:                parseJSONSafe(r[COL.COMPANY_INTEL - 1]),
    salary:                      r[COL.SALARY - 1],
    salaryExpectation:           r[COL.SALARY_EXPECTATION - 1],
    location:                    r[COL.LOCATION - 1],
    workArrangement:             r[COL.WORK_ARRANGEMENT - 1],
    keyAlignments:               r[COL.KEY_ALIGNMENTS - 1],
    potentialConcerns:           r[COL.POTENTIAL_CONCERNS - 1],
    applicationPriority:         r[COL.APPLICATION_PRIORITY - 1],
    tailoredPitch:               r[COL.TAILORED_PITCH - 1],
    analysisDate:                r[COL.ANALYSIS_DATE - 1],
    culturalFit:                 r[COL.CULTURAL_FIT - 1],
    resumeMd:                    r[COL.RESUME_MD - 1],
    coverLetterMd:               r[COL.COVER_LETTER_MD - 1],
    submittedResumeLink:         r[COL.SUBMITTED_RESUME_LINK - 1],
    submittedCoverLetterLink:    r[COL.SUBMITTED_COVER_LETTER_LINK - 1],
    interviewPrep:               parseJSONSafe(r[COL.INTERVIEW_PREP - 1]),
    linkedinContactsLink:        r[COL.LINKEDIN_CONTACTS_LINK - 1],
    lastActivity:                r[COL.LAST_ACTIVITY - 1],
    yourNotes:                   r[COL.YOUR_NOTES - 1],
    notes:                       parseNotes(r[COL.YOUR_NOTES - 1]),
    duplicateCheck:              r[COL.DUPLICATE_CHECK - 1],
    applicationDate:             r[COL.APPLICATION_DATE - 1],
    applicationMethod:           r[COL.APPLICATION_METHOD - 1],
    followUpDate:                r[COL.FOLLOW_UP_DATE - 1],
    dateAdded:                   r[COL.DATE_ADDED - 1],
    contactName:                 r[COL.CONTACT_NAME - 1],
    contactEmail:                r[COL.CONTACT_EMAIL - 1],
    contactLinkedin:             r[COL.CONTACT_LINKEDIN - 1],
    personalAngle:               r[COL.PERSONAL_ANGLE - 1],
    adContent:                   r[COL.AD_CONTENT - 1],
    stageEnteredAt:              r[COL.STAGE_ENTERED_AT - 1],
    statusHistory:               parseStatusHistory(r[COL.STATUS_HISTORY - 1])
  };
}

/**
 * Parses the Your Notes (col Z) cell into an array of { ts, text } entries.
 * Shapes supported:
 *   - JSON array of { ts, text }  (current format)
 *   - Plain text                  (legacy — treated as a single untimestamped entry)
 *   - Blank                       (returns [])
 */
function parseNotes(raw) {
  if (!raw && raw !== 0) return [];
  const s = String(raw).trim();
  if (!s) return [];

  // Attempt JSON parse; if it's a valid array, accept it
  try {
    const parsed = JSON.parse(s);
    if (Array.isArray(parsed)) {
      return parsed.filter(function (e) { return e && e.text; });
    }
  } catch (e) { /* fall through */ }

  // Legacy plain-text note — preserve as a single entry with no timestamp
  return [{ ts: null, text: s, legacy: true }];
}

/**
 * Parses STATUS_HISTORY (col 38) into an array of { status, at } entries
 * in chronological order (oldest first). Empty/garbage returns [].
 */
function parseStatusHistory(raw) {
  if (!raw && raw !== 0) return [];
  const s = String(raw).trim();
  if (!s) return [];
  try {
    const parsed = JSON.parse(s);
    if (Array.isArray(parsed)) {
      return parsed
        .filter(function (e) { return e && e.status; })
        .map(function (e) { return { status: e.status, at: e.at || null }; });
    }
  } catch (e) { /* fall through */ }
  return [];
}


// ============================================================
// QUEUE MANAGEMENT
// ============================================================

function getQueue() {
  const sheet = openQueue();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, items: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const items = values
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const url = String(r[1] || '');
      return {
        id:       String(r[0]),
        url:      url,
        content:  String(r[2] || ''),
        addedAt:  r[3],
        source:   detectSource(url)
      };
    });
  return { ok: true, items: items, count: items.length };
}

function addToQueue(body) {
  const content = (body.content || '').trim();
  const url     = (body.url || '').trim();
  if (!content) return { ok: false, error: 'Content is required' };
  if (!url)     return { ok: false, error: 'URL is required' };

  const sheet = openQueue();
  const id = generateId();
  sheet.appendRow([id, url, content, new Date()]);
  return { ok: true, queue: getQueue().items };
}

function removeFromQueue(body) {
  const queueId = body.queueId;
  if (!queueId) return { ok: false, error: 'queueId is required' };

  const sheet = openQueue();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, queue: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(queueId)) {
      sheet.deleteRow(i + 2);
      return { ok: true, queue: getQueue().items };
    }
  }
  return { ok: false, error: 'Queue item not found' };
}

function clearQueue() {
  const sheet = openQueue();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  return { ok: true, queue: [] };
}


// ============================================================
// STATUS TRANSITION HELPER
// ============================================================

function setStatusWithTimestamp(sheet, rowId, newStatus) {
  const now = new Date();
  const existingStatus = sheet.getRange(rowId, COL.STATUS).getValue();
  sheet.getRange(rowId, COL.STATUS).setValue(newStatus);
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(now);
  if (existingStatus !== newStatus) {
    sheet.getRange(rowId, COL.STAGE_ENTERED_AT).setValue(now);
    appendStatusHistory(sheet, rowId, newStatus, now);
  }
}

/**
 * Append a { status, at } entry to STATUS_HISTORY (col 38).
 * Chronological order, oldest first. Idempotent guard: doesn't append
 * if the most recent entry is already this status (avoids duplicates
 * when setStatusWithTimestamp is called repeatedly with the same value).
 */
function appendStatusHistory(sheet, rowId, status, at) {
  const raw = sheet.getRange(rowId, COL.STATUS_HISTORY).getValue();
  const history = parseStatusHistory(raw);

  const last = history.length ? history[history.length - 1] : null;
  if (last && last.status === status) return;

  history.push({ status: status, at: (at instanceof Date ? at.toISOString() : at) });
  sheet.getRange(rowId, COL.STATUS_HISTORY).setValue(JSON.stringify(history));
}

/**
 * Forward-only enforcement: once a card has reached Applied or any later
 * substage, it cannot move back to Analysed or Ready to Apply. Holding
 * statuses (On Hold / Unsure / Rethink) and archive (Rejected / Withdrawn)
 * are still allowed from any stage, including post-Applied — those are
 * lateral moves, not backward.
 *
 * Returns true if the move is allowed, false if it should be blocked.
 */
function isStatusTransitionAllowed(history, currentStatus, newStatus) {
  const PRE_APPLIED = ['Analysed', 'Ready to Apply'];
  const APPLIED_OR_LATER = ['Applied', 'Interviewing', 'Interviewed', 'Offer', 'Accepted'];

  // Has this card EVER reached Applied or later?
  const hasBeenApplied =
    APPLIED_OR_LATER.indexOf(currentStatus) > -1 ||
    history.some(function (h) { return APPLIED_OR_LATER.indexOf(h.status) > -1; });

  // Once Applied, can't move back to Analysed or Ready to Apply
  if (hasBeenApplied && PRE_APPLIED.indexOf(newStatus) > -1) return false;

  return true;
}

function parseApplicationDate(dateStr) {
  if (!dateStr) return new Date();

  const m = String(dateStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return new Date(dateStr);

  const now = new Date();
  const todayStr = now.getFullYear() + '-' +
    String(now.getMonth() + 1).padStart(2, '0') + '-' +
    String(now.getDate()).padStart(2, '0');

  if (dateStr === todayStr) return now;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}


// ============================================================
// BATCH APPEND (from Analyse master prompt TSV output)
// ============================================================

function batchAppendAnalysed(body) {
  const tsv = (body.tsv || '').trim();
  if (!tsv) return { ok: false, error: 'TSV is required' };

  const lines = tsv.split('\n').filter(function (l) { return l.trim(); });
  if (lines.length < 2) return { ok: false, error: 'TSV must include a header row and at least one data row' };

  const dataRows = lines.slice(1);
  const sheet = openPipeline();
  const appended = [];

  dataRows.forEach(function (line) {
    const cells = line.split('\t');
    while (cells.length < TOTAL_COLS) cells.push('');

    cells[COL.STATUS - 1]            = 'Analysed';
    cells[COL.DATE_ADDED - 1]        = new Date();
    cells[COL.LAST_ACTIVITY - 1]     = new Date();
    cells[COL.STAGE_ENTERED_AT - 1]  = new Date();
    if (!cells[COL.ANALYSIS_DATE - 1]) cells[COL.ANALYSIS_DATE - 1] = new Date();

    if (cells[COL.AD_CONTENT - 1]) {
      cells[COL.AD_CONTENT - 1] = String(cells[COL.AD_CONTENT - 1]).replace(/\\n/g, '\n');
    }

    sheet.appendRow(cells);
    appended.push(sheet.getLastRow());
  });

  clearQueue();

  return { ok: true, appendedRowIds: appended, count: appended.length };
}


// ============================================================
// STAGE TRANSITIONS
// ============================================================

/**
 * Phase 1 paste-back: saves Google Doc LINKS only.
 * The textareas for pasting content have been removed in favour of asking
 * Louie to link the Google Docs he creates manually. Doc content stays in
 * the source of truth (the Doc), not mirrored in the sheet.
 */
function savePhase1Docs(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  const sheet = openPipeline();

  if (body.coverLetterLink) sheet.getRange(rowId, COL.SUBMITTED_COVER_LETTER_LINK).setValue(body.coverLetterLink);
  if (body.resumeLink)      sheet.getRange(rowId, COL.SUBMITTED_RESUME_LINK).setValue(body.resumeLink);

  // Back-compat: still accept content if sent (legacy clients), but don't require it
  if (body.coverLetterContent) sheet.getRange(rowId, COL.COVER_LETTER_MD).setValue(body.coverLetterContent);
  if (body.resumeContent)      sheet.getRange(rowId, COL.RESUME_MD).setValue(body.resumeContent);

  setStatusWithTimestamp(sheet, rowId, 'Ready to Apply');

  return { ok: true, row: rowToObject(sheet.getRange(rowId, 1, 1, TOTAL_COLS).getValues()[0], rowId) };
}

function markAsApplied(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  const sheet = openPipeline();

  if (body.submittedResumeContent)      sheet.getRange(rowId, COL.RESUME_MD).setValue(body.submittedResumeContent);
  if (body.submittedCoverLetterContent) sheet.getRange(rowId, COL.COVER_LETTER_MD).setValue(body.submittedCoverLetterContent);

  sheet.getRange(rowId, COL.APPLICATION_METHOD).setValue(body.appliedVia || '');
  sheet.getRange(rowId, COL.APPLICATION_DATE).setValue(parseApplicationDate(body.applicationDate));
  setStatusWithTimestamp(sheet, rowId, 'Applied');

  const existingPrep = sheet.getRange(rowId, COL.INTERVIEW_PREP).getValue();
  if (!existingPrep) {
    const shell = { generatedAt: null, questions: [] };
    sheet.getRange(rowId, COL.INTERVIEW_PREP).setValue(JSON.stringify(shell));
  }

  return { ok: true, row: rowToObject(sheet.getRange(rowId, 1, 1, TOTAL_COLS).getValues()[0], rowId) };
}

function saveInterviewPrep(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  if (!body.interviewPrepJson) return { ok: false, error: 'interviewPrepJson is required' };

  try { JSON.parse(body.interviewPrepJson); }
  catch (e) { return { ok: false, error: 'Invalid JSON: ' + String(e) }; }

  const sheet = openPipeline();
  sheet.getRange(rowId, COL.INTERVIEW_PREP).setValue(body.interviewPrepJson);
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
  return { ok: true };
}

/**
 * Update a single round's score on an interview prep question. Used when the
 * user edits a score inline from the dashboard. Loads the prep JSON, finds
 * the question by id, finds the round by round number, sets score, writes back.
 *
 * Body: { rowId, questionId, roundNumber, score }
 * Score must be integer 0-10. 0 means "no score / cleared".
 */
function updateRoundScore(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  if (!body.questionId) return { ok: false, error: 'questionId is required' };
  const roundNumber = parseInt(body.roundNumber, 10);
  if (isNaN(roundNumber)) return { ok: false, error: 'roundNumber is required' };
  const score = parseInt(body.score, 10);
  if (isNaN(score) || score < 0 || score > 10) return { ok: false, error: 'score must be integer 0-10' };

  const sheet = openPipeline();
  const raw = sheet.getRange(rowId, COL.INTERVIEW_PREP).getValue();
  let prep;
  try { prep = raw ? JSON.parse(raw) : null; }
  catch (e) { return { ok: false, error: 'Stored prep JSON is invalid: ' + String(e) }; }
  if (!prep || !Array.isArray(prep.questions)) return { ok: false, error: 'No interview prep on this row' };

  const q = prep.questions.find(function (x) { return x.id === body.questionId; });
  if (!q) return { ok: false, error: 'Question not found: ' + body.questionId };
  if (!Array.isArray(q.rounds)) return { ok: false, error: 'No rounds on this question' };

  const r = q.rounds.find(function (x) { return x.round === roundNumber; });
  if (!r) return { ok: false, error: 'Round not found: ' + roundNumber };

  r.score = score;

  sheet.getRange(rowId, COL.INTERVIEW_PREP).setValue(JSON.stringify(prep));
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
  return { ok: true };
}

function updateCardStatus(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  if (!body.newStatus) return { ok: false, error: 'newStatus is required' };

  const validStatuses = [].concat(ACTIVE_STATUSES, HOLDING_STATUSES, ARCHIVE_STATUSES, APPLIED_SUB);
  if (validStatuses.indexOf(body.newStatus) === -1) {
    return { ok: false, error: 'Invalid status: ' + body.newStatus };
  }

  const sheet = openPipeline();
  const currentStatus = sheet.getRange(rowId, COL.STATUS).getValue();
  const history = parseStatusHistory(sheet.getRange(rowId, COL.STATUS_HISTORY).getValue());

  if (!isStatusTransitionAllowed(history, currentStatus, body.newStatus)) {
    return {
      ok: false,
      error: 'Cannot move back to ' + body.newStatus + ' once a card has reached Applied. Once submitted, an application can only move forward through interview substages or be archived.'
    };
  }

  setStatusWithTimestamp(sheet, rowId, body.newStatus);
  return { ok: true };
}

function savePersonalAngle(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };

  const sheet = openPipeline();
  sheet.getRange(rowId, COL.PERSONAL_ANGLE).setValue(body.personalAngle || '');
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
  return { ok: true };
}

/**
 * Save ad content to col AD_CONTENT for a row. Used when backfilling
 * older cards that were analysed before the ad-content-capture column
 * existed, and also when the user pastes an updated ad to refresh intel
 * against newer copy. Accepts markdown or plain text.
 */
function saveAdContent(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  const content = (body.adContent || '').trim();

  const sheet = openPipeline();
  sheet.getRange(rowId, COL.AD_CONTENT).setValue(content);
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
  return { ok: true, charCount: content.length };
}

/**
 * Appends a timestamped note to the Your Notes column (col Z).
 * Storage format: JSON array of { ts, text } entries, newest first.
 * Legacy plain-text notes are preserved: they become the oldest entry
 * (flagged `legacy: true`) when the first new JSON note is added.
 */
function saveNote(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };
  const text = (body.noteText || '').trim();
  if (!text) return { ok: false, error: 'noteText is required' };

  const sheet = openPipeline();
  const raw = sheet.getRange(rowId, COL.YOUR_NOTES).getValue();
  const existing = parseNotes(raw);

  const newEntry = { ts: new Date().toISOString(), text: text };
  const updated = [newEntry].concat(existing); // newest first

  sheet.getRange(rowId, COL.YOUR_NOTES).setValue(JSON.stringify(updated));
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());

  return { ok: true, notes: updated };
}

/**
 * Update the text of an existing note entry by index. Preserves the
 * original timestamp — edits don\'t bump the note\'s ts because this is
 * "fix a typo" not "post again." Touches LAST_ACTIVITY though so the
 * card reads as recently-touched.
 */
function updateNote(body) {
  const rowId = parseInt(body.rowId, 10);
  const idx = parseInt(body.idx, 10);
  const text = (body.text || '').trim();

  if (!rowId) return { ok: false, error: 'rowId is required' };
  if (isNaN(idx) || idx < 0) return { ok: false, error: 'idx is required' };
  if (!text) return { ok: false, error: 'text is required (empty = use deleteNote instead)' };

  const sheet = openPipeline();
  const raw = sheet.getRange(rowId, COL.YOUR_NOTES).getValue();
  const existing = parseNotes(raw);

  if (idx >= existing.length) {
    return { ok: false, error: 'idx ' + idx + ' out of range (have ' + existing.length + ' notes)' };
  }

  existing[idx] = {
    ts: existing[idx].ts || new Date().toISOString(),
    text: text,
    edited: new Date().toISOString()
  };
  if (existing[idx].legacy) existing[idx].legacy = true;  // preserve legacy flag if present

  sheet.getRange(rowId, COL.YOUR_NOTES).setValue(JSON.stringify(existing));
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());

  return { ok: true, notes: existing };
}

/**
 * Delete a note entry by index. Destructive — no undo.
 */
function deleteNote(body) {
  const rowId = parseInt(body.rowId, 10);
  const idx = parseInt(body.idx, 10);

  if (!rowId) return { ok: false, error: 'rowId is required' };
  if (isNaN(idx) || idx < 0) return { ok: false, error: 'idx is required' };

  const sheet = openPipeline();
  const raw = sheet.getRange(rowId, COL.YOUR_NOTES).getValue();
  const existing = parseNotes(raw);

  if (idx >= existing.length) {
    return { ok: false, error: 'idx ' + idx + ' out of range' };
  }

  existing.splice(idx, 1);
  sheet.getRange(rowId, COL.YOUR_NOTES).setValue(JSON.stringify(existing));
  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());

  return { ok: true, notes: existing };
}

/**
 * Replace the full STATUS_HISTORY for a row with a user-edited version.
 * Called from the Edit journey modal. Accepts whatever array the user
 * built — no sequencing enforcement, because the user's memory of what
 * actually happened is more reliable than any rule we could impose.
 *
 * Side effect: if the last entry in the edited history differs from the
 * current STATUS cell, the STATUS cell is updated too — so editing the
 * journey to end at a different status moves the card to that status.
 * This is the right behaviour when fixing "I moved it to Offer but it
 * should've been Accepted" mistakes.
 */
function saveStatusHistory(body) {
  const rowId = parseInt(body.rowId, 10);
  if (!rowId) return { ok: false, error: 'rowId is required' };

  let history;
  try {
    history = JSON.parse(body.historyJson || '[]');
  } catch (e) {
    return { ok: false, error: 'Invalid history JSON: ' + String(e) };
  }
  if (!Array.isArray(history)) return { ok: false, error: 'History must be an array' };

  // Validate each entry shape — status required, at optional but should
  // be a valid date string if present
  const validStatuses = [].concat(ACTIVE_STATUSES, HOLDING_STATUSES, ARCHIVE_STATUSES, APPLIED_SUB);
  const clean = [];
  for (var i = 0; i < history.length; i++) {
    const entry = history[i];
    if (!entry || !entry.status) continue;
    if (validStatuses.indexOf(entry.status) === -1) {
      return { ok: false, error: 'Invalid status at entry ' + (i + 1) + ': ' + entry.status };
    }
    clean.push({
      status: entry.status,
      at: entry.at ? toIso(entry.at) : null
    });
  }

  const sheet = openPipeline();
  sheet.getRange(rowId, COL.STATUS_HISTORY).setValue(JSON.stringify(clean));

  // Sync the current STATUS cell with the last entry if different
  if (clean.length) {
    const finalStatus = clean[clean.length - 1].status;
    const currentStatus = sheet.getRange(rowId, COL.STATUS).getValue();
    if (finalStatus !== currentStatus) {
      sheet.getRange(rowId, COL.STATUS).setValue(finalStatus);
      sheet.getRange(rowId, COL.STAGE_ENTERED_AT).setValue(new Date());
    }
  }

  sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());

  return { ok: true, history: clean };
}


// ============================================================
// BATCH JOURNEY UPDATE
// ============================================================

/**
 * Apply a status journey template to multiple cards at once. For each
 * selected card, the template entries are merged into the existing
 * journey using these rules:
 *   - If template lists a status the card doesn\'t have, add it with
 *     the template\'s date
 *   - If template lists a status the card already has, UPDATE the
 *     existing entry\'s date to match the template (backdate fix)
 *   - Existing journey entries not mentioned in the template stay
 *     untouched
 *
 * After merging, the journey is sorted chronologically by date and
 * the card\'s current STATUS cell is synced to the latest entry.
 *
 * No forward-only lock: this path is for backdating and historical
 * correction, NOT for changing current status. Use updateCardStatus
 * for "move card forward" actions.
 *
 * Body: {
 *   rowIds: [123, 456, 789],
 *   template: [
 *     { status: "Analysed", at: "2026-03-01T..." },
 *     { status: "Applied",  at: "2026-04-15T..." }
 *   ]
 * }
 *
 * Returns: { ok, updated: N, skipped: N, results: [{ rowId, status: "updated"|"skipped"|"error", message }] }
 */
function batchUpdateJourney(body) {
  // Defensive parsing: rowIds may arrive as array, as comma-separated string,
  // or wrapped — try to coerce sensibly before failing.
  let rowIds = [];
  if (Array.isArray(body.rowIds)) {
    rowIds = body.rowIds.map(function (n) { return parseInt(n, 10); }).filter(function (n) { return !isNaN(n) && n > 0; });
  } else if (typeof body.rowIds === 'string') {
    rowIds = body.rowIds.split(',').map(function (n) { return parseInt(n, 10); }).filter(function (n) { return !isNaN(n) && n > 0; });
  }
  if (!rowIds.length) {
    return { ok: false, error: 'rowIds is required (got: ' + JSON.stringify(body.rowIds) + ')' };
  }

  const template = Array.isArray(body.template) ? body.template : [];
  if (!template.length) {
    return { ok: false, error: 'template is required (got: ' + JSON.stringify(body.template) + ')' };
  }

  // Validate template up-front so a bad status fails fast for the whole batch
  const validStatuses = [].concat(ACTIVE_STATUSES, HOLDING_STATUSES, ARCHIVE_STATUSES, APPLIED_SUB);
  for (var t = 0; t < template.length; t++) {
    if (!template[t] || !template[t].status) {
      return { ok: false, error: 'Template entry ' + (t + 1) + ' is missing status' };
    }
    if (validStatuses.indexOf(template[t].status) === -1) {
      return { ok: false, error: 'Template entry ' + (t + 1) + ' has invalid status: "' + template[t].status + '" (valid: ' + validStatuses.join(', ') + ')' };
    }
    if (!template[t].at) {
      return { ok: false, error: 'Template entry ' + (t + 1) + ' (' + template[t].status + ') is missing date' };
    }
  }

  let sheet;
  try {
    sheet = openPipeline();
  } catch (e) {
    return { ok: false, error: 'Failed to open Pipeline sheet: ' + String(e) };
  }

  const results = [];
  let updated = 0;
  let skipped = 0;

  for (var i = 0; i < rowIds.length; i++) {
    const rowId = rowIds[i];
    try {
      const existingRaw = sheet.getRange(rowId, COL.STATUS_HISTORY).getValue();
      const existing = parseStatusHistory(existingRaw);

      // Build a status->entry map for O(1) lookup. Template overrides win.
      const merged = {};
      existing.forEach(function (e) { merged[e.status] = { status: e.status, at: e.at }; });
      template.forEach(function (e) {
        const isoVal = toIso(e.at);
        if (!isoVal) {
          throw new Error('Could not parse date "' + e.at + '" for status ' + e.status);
        }
        merged[e.status] = { status: e.status, at: isoVal };
      });

      // Convert back to chronological array. Sort by date ascending; entries
      // with no date go last in their original position (rare).
      const mergedArr = Object.keys(merged).map(function (k) { return merged[k]; });
      mergedArr.sort(function (a, b) {
        if (!a.at && !b.at) return 0;
        if (!a.at) return 1;
        if (!b.at) return -1;
        return new Date(a.at).getTime() - new Date(b.at).getTime();
      });

      // Persist the merged journey
      sheet.getRange(rowId, COL.STATUS_HISTORY).setValue(JSON.stringify(mergedArr));

      // Sync current STATUS to the chronologically-latest entry (which is
      // probably what the user wants — if they backdated Analysed for a
      // card that\'s currently Applied, Applied stays as the current status
      // because Applied has the later date).
      if (mergedArr.length) {
        const finalStatus = mergedArr[mergedArr.length - 1].status;
        const currentStatus = sheet.getRange(rowId, COL.STATUS).getValue();
        if (finalStatus !== currentStatus) {
          sheet.getRange(rowId, COL.STATUS).setValue(finalStatus);
          sheet.getRange(rowId, COL.STAGE_ENTERED_AT).setValue(new Date());
        }
      }

      sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
      updated++;
      results.push({ rowId: rowId, status: 'updated' });
    } catch (err) {
      skipped++;
      results.push({ rowId: rowId, status: 'error', message: String(err && err.message ? err.message : err) });
    }
  }

  // If everything was skipped, surface that as a failure with details.
  // Otherwise return success even if some rows failed (partial success).
  if (updated === 0 && skipped > 0) {
    const errorSummary = results.filter(function (r) { return r.status === 'error'; })
      .map(function (r) { return 'Row ' + r.rowId + ': ' + r.message; })
      .slice(0, 3)  // first 3 for brevity
      .join(' | ');
    return { ok: false, error: 'All ' + skipped + ' rows failed. ' + errorSummary, results: results };
  }

  return { ok: true, updated: updated, skipped: skipped, results: results };
}


// ============================================================
// LEGACY: appendAnalysedJob (kept for backward compat)
// ============================================================

function appendAnalysedJob(body) {
  const force = body.force === true;
  const sheet = openPipeline();

  if (!force && body.url) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const startRow = Math.max(2, lastRow - 499);
      const urls = sheet.getRange(startRow, COL.URL, lastRow - startRow + 1, 1).getValues();
      for (var i = 0; i < urls.length; i++) {
        if (urls[i][0] === body.url) {
          return { ok: false, duplicate: true, message: 'URL already in pipeline. Pass force:true to override.' };
        }
      }
    }
  }

  const row = new Array(TOTAL_COLS).fill('');
  row[COL.STATUS - 1]                = 'Analysed';
  row[COL.COMPANY - 1]               = body.company || '';
  row[COL.ROLE - 1]                  = body.role || '';
  row[COL.FIT_SCORE - 1]             = body.fitScore || '';
  row[COL.RECOMMENDATION - 1]        = body.recommendation || '';
  row[COL.URL - 1]                   = body.url || '';
  row[COL.KEY_NOTES - 1]             = body.keyNotes || '';
  row[COL.COMPANY_INTEL - 1]         = body.companyIntel || '';
  row[COL.SALARY - 1]                = body.salary || '';
  row[COL.LOCATION - 1]              = body.location || '';
  row[COL.WORK_ARRANGEMENT - 1]      = body.workArrangement || '';
  row[COL.KEY_ALIGNMENTS - 1]        = body.keyAlignments || '';
  row[COL.POTENTIAL_CONCERNS - 1]    = body.potentialConcerns || '';
  row[COL.APPLICATION_PRIORITY - 1]  = body.applicationPriority || '';
  row[COL.TAILORED_PITCH - 1]        = body.tailoredPitch || '';
  row[COL.ANALYSIS_DATE - 1]         = new Date();
  row[COL.DATE_ADDED - 1]            = new Date();
  row[COL.LAST_ACTIVITY - 1]         = new Date();
  row[COL.PERSONAL_ANGLE - 1]        = body.personalAngle || '';

  sheet.appendRow(row);
  return { ok: true, rowId: sheet.getLastRow() };
}


// ============================================================
// PROMPT DELIVERY
// ============================================================

function getPrompt(type, params) {
  function parseIds(s) {
    return String(s || '').split(',').map(function (v) { return parseInt(v, 10); }).filter(function (n) { return !isNaN(n); });
  }

  switch (type) {
    case 'analyse':              return { ok: true, prompt: buildAnalysePrompt() };
    case 'generateDocs':         return { ok: true, prompt: buildGenerateDocsPrompt(parseInt(params.rowId, 10)) };
    case 'interviewPrep':        return { ok: true, prompt: buildInterviewPrepPrompt(parseInt(params.rowId, 10)) };
    case 'continuePractice':     return { ok: true, prompt: buildContinuePracticePrompt(parseInt(params.rowId, 10)) };
    case 'voicePractice':        return { ok: true, prompt: buildVoicePracticePrompt(parseInt(params.rowId, 10)) };
    case 'batchAnalyse':         return { ok: true, prompt: buildBatchAnalysePrompt(parseIds(params.rowIds)) };
    case 'batchGenerateDocs':    return { ok: true, prompt: buildBatchGenerateDocsPrompt(parseIds(params.rowIds)) };
    case 'batchInterviewPrep':   return { ok: true, prompt: buildBatchInterviewPrepPrompt(parseIds(params.rowIds)) };
    default: return { ok: false, error: 'Unknown prompt type: ' + type };
  }
}


// ============================================================
// PROMPT BUILDERS
// ============================================================

function buildAnalysePrompt() {
  const queue = getQueue().items;
  if (!queue.length) return 'Queue is empty. Add jobs via the intake modal before generating the analyse prompt.';

  const jobsBlock = queue.map(function (item, i) {
    return 'JOB ' + (i + 1) + '\n' +
           'URL: ' + item.url + '\n' +
           'SOURCE: ' + item.source + '\n' +
           'CONTENT:\n' + item.content + '\n';
  }).join('\n---\n\n');

  return [
    'You are analysing ' + queue.length + ' job ad' + (queue.length > 1 ? 's' : '') + ' for Louie Radburnd\'s job pipeline. Produce a TSV output for direct paste into the Job Pipeline v2 spreadsheet.',
    '',
    'JOBS:',
    '',
    jobsBlock,
    '',
    'FOR EACH JOB, GENERATE:',
    '- Company and Role (extracted from content)',
    '- Fit Score (integer 0-10, whole numbers only, 8+ is strong) based on Louie\'s profile: 18+ years full-stack B2B/association marketing, MTAA 7 years (Marketing & Membership Manager), IndustraCom GM and Head of Growth, Sydney-based',
    '- Recommendation: MUST be one of exactly these three values (uppercase): "APPLY", "CONSIDER", "SKIP". Use APPLY for strong fits (~7+), CONSIDER for borderline, SKIP for weak fits.',
    '- Key Notes: 2-3 sentences on what the role actually is and why it matters',
    '- Company Intel JSON (single-line, escaped) with this structure:',
    '  {"industry":"...","products":"...","scale":"...","newsworthy":"...","marketingInsights":{"positioning":"...","goToMarket":"...","competitors":"...","teamSignals":"...","talkingPoints":"..."}}',
    '- Salary (from ad, or "Not specified")',
    '- Location and Work Arrangement',
    '- Key Alignments: 3-5 items separated by " | "',
    '- Potential Concerns: 2-4 items separated by " | "',
    '- Application Priority: MUST be one of exactly these three values (uppercase): "HIGH", "MEDIUM", "LOW"',
    '- Tailored Pitch: 1-2 sentences on how Louie should position',
    '- Personal Angle: 1 sentence on the human hook (e.g. parent-of-two for early-childhood roles, long-time user for product roles)',
    '- Ad Content Markdown: REQUIRED. A clean markdown representation of the full job ad, so Louie can reference it without revisiting the URL. Include role title, company, salary if stated, full responsibilities, required experience, and any selection criteria. Use # and ## for structure, - for bullets. CRITICAL: within the TSV cell, replace every real newline with the literal string \\n (backslash-n) so the cell remains on one line. Do NOT include tab characters. The dashboard will restore newlines when displaying.',
    '',
    'OUTPUT FORMAT:',
    'Return a single TSV block, tab-separated, newline-ended. First row is the header below. One row per job.',
    'Status\\tCompany\\tRole\\tFit Score\\tRecommendation\\tURL\\tKey Notes\\tCompany Intel\\tSalary\\tSalary Expectation\\tLocation\\tWork Arrangement\\tKey Alignments\\tPotential Concerns\\tApplication Priority\\tTailored Pitch\\tAnalysis Date\\tCultural Fit\\tResume Markdown\\tCover Letter Markdown\\tSubmitted Resume Link\\tSubmitted Cover Letter Link\\tInterview Prep\\tLinkedIn Contacts Link\\tLast Activity\\tYour Notes\\tDuplicate Check\\tApplication Date\\tApplication Method\\tFollow-up Date\\tDate Added\\tContact Name\\tContact Email\\tContact LinkedIn\\tPersonal Angle\\tAd Content Markdown',
    '',
    'ROW RULES:',
    '- Status: "Analysed"',
    '- Company Intel: the JSON as a single escaped string (no real tabs or newlines inside)',
    '- Salary Expectation, Cultural Fit, Resume Markdown through Contact LinkedIn: leave blank',
    '- Analysis Date and Date Added: today',
    '- Personal Angle: the sentence you generated',
    '- Ad Content Markdown: the full job ad as markdown, with newlines replaced by the literal two-character sequence \\n',
    '',
    'IMPORTANT:',
    '- No em dashes anywhere',
    '- Australian English spellings (organise, analyse, colour, etc.)',
    '- TSV must be valid: no stray tabs or real newlines inside cell values (newlines in Ad Content Markdown must be the literal two-character \\n, not real newlines)',
    '- Output the raw TSV only, no surrounding code fences, no commentary before or after',
    '',
    'Wrap the TSV in <tsv> tags so it\'s unambiguous to paste back.'
  ].join('\n');
}

function buildGenerateDocsPrompt(rowId) {
  const row = getRow(rowId).row;
  if (!row) return 'Row ' + rowId + ' not found.';

  const intel = row.companyIntel || {};
  const mi = intel.marketingInsights || {};

  return [
    'Generate a tailored resume and cover letter for Louie Radburnd applying to the role below. Use your master resume (Louie_Radburnd_Master_Resume_FINAL.docx in this project) as the source of truth for career history.',
    '',
    'ROLE: ' + (row.role || ''),
    'COMPANY: ' + (row.company || ''),
    'LOCATION: ' + (row.location || ''),
    'WORK ARRANGEMENT: ' + (row.workArrangement || ''),
    'SALARY: ' + (row.salary || ''),
    'JOB URL: ' + (row.url || ''),
    '',
    'KEY NOTES: ' + (row.keyNotes || ''),
    'KEY ALIGNMENTS: ' + (row.keyAlignments || ''),
    'POTENTIAL CONCERNS: ' + (row.potentialConcerns || ''),
    'TAILORED PITCH: ' + (row.tailoredPitch || ''),
    'PERSONAL ANGLE: ' + (row.personalAngle || ''),
    '',
    'COMPANY INTEL:',
    '- Industry: ' + (intel.industry || ''),
    '- Products: ' + (intel.products || ''),
    '- Scale: ' + (intel.scale || ''),
    '- Positioning: ' + (mi.positioning || ''),
    '- Go-to-market: ' + (mi.goToMarket || ''),
    '- Competitors: ' + (mi.competitors || ''),
    '- Team signals: ' + (mi.teamSignals || ''),
    '',
    'RESUME REQUIREMENTS:',
    '- Tailor content from master resume, selecting most relevant experience and achievements',
    '- Prioritise measurable outcomes over responsibilities',
    '- Minimum 2 bullets per role, minimum 5 for MTAA and IndustraCom',
    '- Include all 7 roles (2008-2025)',
    '- Profile starts with "Most marketers specialise. I\'ve spent 18+ years doing all of it..." (Version 3)',
    '- Closing personal line in profile should be clever, role-specific, one sentence max — make the reader smirk or nod',
    '- No em dashes anywhere',
    '- Australian English',
    '- 2 pages max',
    '',
    'COVER LETTER VOICE (locked):',
    '- Direct, conversational, slightly opinionated',
    '- First 1-2 paragraphs: answer "why this role / why this company" explicitly',
    '- Concrete proof points from MTAA / IndustraCom with real numbers',
    '- Address obvious gaps (sector, credentials) on own terms',
    '- Include AI paragraph citing the n8n-based job pipeline (fit scoring, deep company research, dashboard-generated tailored docs) with marketing-equivalent use cases',
    '- Forward-looking close',
    '- No em dashes',
    '- Australian English',
    '- 250-400 words, one A4 page',
    '',
    'COVER LETTER FORMAT (for output as markdown — Louie will paste into his branded DOCX template):',
    '- No bold colon-format headings inside the body',
    '- Salutation: "Hi Hiring Manager"',
    '- Sign off: "Louie Radburnd" only',
    '',
    'OUTPUT FORMAT:',
    'Return exactly two blocks, clearly delimited:',
    '',
    '=== COVER LETTER ===',
    '[cover letter markdown content here]',
    '=== END COVER LETTER ===',
    '',
    '=== RESUME ===',
    '[resume markdown content here]',
    '=== END RESUME ===',
    '',
    'After both blocks, stop. No commentary, no explanation.'
  ].join('\n');
}

function buildInterviewPrepPrompt(rowId) {
  const row = getRow(rowId).row;
  if (!row) return 'Row ' + rowId + ' not found.';

  const intel = row.companyIntel || {};
  const mi = intel.marketingInsights || {};

  return [
    'Generate an interview prep JSON for Louie Radburnd applying to the role below. Return valid JSON only, no commentary.',
    '',
    'ROLE: ' + (row.role || ''),
    'COMPANY: ' + (row.company || ''),
    'JOB URL: ' + (row.url || ''),
    '',
    'KEY NOTES: ' + (row.keyNotes || ''),
    'KEY ALIGNMENTS: ' + (row.keyAlignments || ''),
    'POTENTIAL CONCERNS: ' + (row.potentialConcerns || ''),
    'TAILORED PITCH: ' + (row.tailoredPitch || ''),
    'PERSONAL ANGLE: ' + (row.personalAngle || ''),
    '',
    'COMPANY INTEL:',
    '- Industry: ' + (intel.industry || ''),
    '- Products: ' + (intel.products || ''),
    '- Positioning: ' + (mi.positioning || ''),
    '- Go-to-market: ' + (mi.goToMarket || ''),
    '- Competitors: ' + (mi.competitors || ''),
    '- Team signals: ' + (mi.teamSignals || ''),
    '- Talking points: ' + (mi.talkingPoints || ''),
    '',
    'GENERATE 6 ANTICIPATED INTERVIEW QUESTIONS:',
    '- 4 marked "Likely" (high probability based on role and company context)',
    '- 2 marked "Possible" (plausible follow-ups, stretch questions)',
    '- Mix strategic (how would you approach X), behavioural (tell me about a time when), and company-specific (what\'s your view on our recent move)',
    '',
    'FOR EACH QUESTION, PROVIDE A PREPARED ANSWER:',
    '- Reference specific Louie stories from his career (MTAA conference scaling, IndustraCom turnaround, RegConnect build, ABC rapid response campaign, membership transformation)',
    '- Use real numbers where applicable',
    '- Prepared answer is a STARTING POINT for Louie to refine in practice sessions, not a final script',
    '- Tone: confident, specific, not over-polished',
    '- preparedAnswerBullets: the same answer compressed to 3-5 short bullet points for quick recall right before the interview. Each bullet 6-12 words, leading with the verb or the number',
    '',
    'OUTPUT FORMAT — return exactly this JSON shape:',
    '',
    '{',
    '  "generatedAt": "<ISO 8601 timestamp>",',
    '  "questions": [',
    '    {',
    '      "id": "q1",',
    '      "text": "<question text>",',
    '      "likelihood": "Likely",',
    '      "status": "not_practiced",',
    '      "preparedAnswer": "<starting-point answer>",',
    '      "preparedAnswerBullets": ["<bullet 1>", "<bullet 2>", "<bullet 3>"],',
    '      "currentRound": 0,',
    '      "rounds": []',
    '    }',
    '    // 5 more questions following the same shape',
    '  ]',
    '}',
    '',
    'Wrap the JSON in <interview_prep> tags so it\'s unambiguous to paste back.',
    '',
    'IMPORTANT:',
    '- No em dashes',
    '- Australian English',
    '- JSON must be valid and parseable',
    '- Do NOT include a STAR responses section or a claims audit section',
    '- Do NOT invent metrics that aren\'t in Louie\'s verified career history'
  ].join('\n');
}

function buildContinuePracticePrompt(rowId) {
  const row = getRow(rowId).row;
  if (!row) return 'Row ' + rowId + ' not found.';

  const prep = row.interviewPrep || { questions: [] };
  const questions = prep.questions || [];

  const summary = questions.map(function (q, i) {
    const status = q.status || 'not_practiced';
    const rounds = q.currentRound || 0;
    return '- Q' + (i + 1) + ' [' + status + ', round ' + rounds + ']: ' + q.text;
  }).join('\n');

  return [
    'Continue the interview practice session with Louie for the role below. Act as the interviewer. Ask one question, wait for his answer, give honest direct feedback, then prompt him to refine or move to the next question.',
    '',
    'ROLE: ' + (row.role || ''),
    'COMPANY: ' + (row.company || ''),
    '',
    'QUESTION STATE:',
    summary || '(no questions yet — generate them first via the interview prep prompt)',
    '',
    'FULL INTERVIEW PREP JSON (current state):',
    JSON.stringify(prep, null, 2),
    '',
    'YOUR JOB:',
    '1. Suggest which question to work on next (prioritise in_progress over not_practiced, lowest round number first)',
    '2. Ask Louie that question out loud — not the prepared answer, the actual question',
    '3. Wait for his response',
    '4. Give feedback that is:',
    '   - Honest and specific (what worked, what didn\'t, why)',
    '   - Concrete (point to phrases, not vague impressions)',
    '   - Actionable (what to change next round)',
    '   - Never flattering without evidence',
    '5. At the end of each round, output an updated JSON blob with the new round appended to that question\'s rounds array. Include:',
    '   - round: <number>',
    '   - userAnswer: <his answer>',
    '   - feedback: <your full feedback>',
    '   - summary: <one-line synthesis for history, e.g. "Better than round 1, owned the decision, still missing specifics">',
    '   - score: <integer 1-10> using this rubric:',
    '       1-3: hesitant, vague, no specifics, dodged the question',
    '       4-5: covered the basics but missed the point or rambled',
    '       6-7: solid, on-topic, used a real story, minor polish needed',
    '       8-9: sharp, specific, well-structured, lands the punchline',
    '       10: nailed it, ready for the actual interview',
    '6. Offer: continue this question for another round, mark it as nailed, or move to the next one',
    '',
    'OUTPUT FORMAT PER ROUND:',
    'Your feedback in prose, followed by an updated JSON blob wrapped in <interview_prep> tags.',
    '',
    'IMPORTANT:',
    '- No em dashes',
    '- Australian English',
    '- When a question is marked nailed, drop the rounds array in favour of a final polished answer',
    '- If Louie\'s answer references a story from his STAR bank, check it aligns with his verified history (MTAA conference $214k to $815k, IndustraCom -10 to +2 per month, RegConnect 6-week build, etc.)'
  ].join('\n');
}

/**
 * Voice-mode practice prompt for ChatGPT. Bundles the role, ad, current
 * prep state, voice-mode rules, and end-of-session JSON spec so Louie can
 * paste it into ChatGPT voice mode and run a verbal mock interview.
 *
 * Assumes the ChatGPT project already has the resume + STAR bank uploaded
 * (Voice and Standards file etc), so per-session prompt stays lean.
 *
 * The output JSON spec at the end matches the existing paste-back parser,
 * so finishing in ChatGPT and pasting back into the dashboard works the
 * same as the text-mode flow.
 */
function buildVoicePracticePrompt(rowId) {
  const row = getRow(rowId).row;
  if (!row) return 'Row ' + rowId + ' not found.';

  const prep = row.interviewPrep || { questions: [] };
  const questions = prep.questions || [];

  const summary = questions.map(function (q, i) {
    const status = q.status || 'not_practiced';
    const rounds = q.currentRound || 0;
    return '- Q' + (i + 1) + ' [' + status + ', round ' + rounds + ']: ' + q.text;
  }).join('\n');

  return [
    'VOICE PRACTICE SESSION — INTERVIEW COACH',
    '',
    'You are running a voice-mode mock interview for Louie Radburnd. He will speak; you will respond conversationally. Treat this like a real coaching session, not a chat with an LLM.',
    '',
    'ROLE: ' + (row.role || ''),
    'COMPANY: ' + (row.company || ''),
    'JOB URL: ' + (row.url || ''),
    '',
    'AD CONTENT:',
    (row.adContent || '(no ad content)'),
    '',
    'CURRENT PRACTICE STATE:',
    summary || '(no questions yet)',
    '',
    'FULL INTERVIEW PREP JSON:',
    JSON.stringify(prep, null, 2),
    '',
    'VOICE-MODE RULES:',
    '1. Speak naturally. One question at a time. No reading lists, no JSON during the session.',
    '2. After Louie answers, give honest spoken feedback in 2-4 sentences. What worked, what didn\'t, what to change next round. No vague praise.',
    '3. Ask if he wants to try the same question again, mark it nailed, or move on.',
    '4. Pick the next question by priority: in_progress (lowest round first), then not_practiced. Skip nailed ones unless he asks for a refresher.',
    '5. Keep your tone direct and Australian. No corporate filler. No "great answer!" without specifics.',
    '6. If he gives a story, check it against his verified history (MTAA conference $214k to $815k over 7 years, IndustraCom -10 to +2 per month inside 6 months, RegConnect built in 6 weeks, MedTech LinkedIn 200 to 16,000+, etc). If a number sounds off, gently flag it.',
    '',
    'END OF SESSION:',
    'When Louie says "wrap up" or "we\'re done", do these three things in order:',
    '1. Give a 2-3 sentence summary of what improved and what still needs work.',
    '2. Output the updated interview prep JSON wrapped in <interview_prep> tags. For each question he practiced this session, append new round(s) to the rounds array. Each new round must include:',
    '   - round: <number>',
    '   - userAnswer: <a short paraphrase of what he said, since this is voice>',
    '   - feedback: <your full feedback for that round>',
    '   - summary: <one-line synthesis>',
    '   - score: <integer 1-10> using this rubric:',
    '       1-3: hesitant, vague, no specifics, dodged the question',
    '       4-5: covered the basics but missed the point or rambled',
    '       6-7: solid, on-topic, used a real story, minor polish needed',
    '       8-9: sharp, specific, well-structured, lands the punchline',
    '       10: nailed it, ready for the actual interview',
    '   For nailed questions, drop the rounds array in favour of a final polished preparedAnswer.',
    '3. Tell him to paste the JSON back into the dashboard via "Save practice JSON".',
    '',
    'IMPORTANT:',
    '- No em dashes in spoken or written output',
    '- Australian English',
    '- The JSON is the durable record. Without it, this session evaporates when the chat closes.'
  ].join('\n');
}


// ============================================================
// BATCH PROMPT BUILDERS
// ============================================================

function buildBatchAnalysePrompt(rowIds) {
  if (!rowIds || !rowIds.length) return 'No row IDs provided.';
  const rows = rowIds.map(function (id) { return getRow(id).row; }).filter(Boolean);
  if (!rows.length) return 'No valid rows found for the given IDs.';

  const jobsBlock = rows.map(function (r) {
    return '<<< JOB rowId=' + r.rowId + ' >>>\n' +
           'URL: ' + (r.url || '') + '\n' +
           'COMPANY: ' + (r.company || '') + '\n' +
           'CURRENT ROLE: ' + (r.role || '') + '\n' +
           'AD CONTENT:\n' + (r.adContent || '(no ad content captured)') + '\n' +
           '<<< END JOB ' + r.rowId + ' >>>';
  }).join('\n\n');

  return [
    'You are re-analysing ' + rows.length + ' job ad' + (rows.length > 1 ? 's' : '') + ' for Louie Radburnd. Produce a fresh analysis for each job, using the current ad content. Return one block per job using the exact delimiter format shown.',
    '',
    'JOBS:',
    '',
    jobsBlock,
    '',
    'FOR EACH JOB:',
    '- Fit Score (integer 0-10, whole numbers only, 8+ is strong)',
    '- Recommendation: MUST be one of exactly these three values (uppercase): "APPLY", "CONSIDER", "SKIP". Use APPLY for strong fits (~7+), CONSIDER for borderline, SKIP for weak fits.',
    '- Key Notes: 2-3 sentences on what the role is and why it matters',
    '- Key Alignments: 3-5 items separated by " | "',
    '- Potential Concerns: 2-4 items separated by " | "',
    '- Application Priority: MUST be one of exactly these three values (uppercase): "HIGH", "MEDIUM", "LOW"',
    '- Tailored Pitch: 1-2 sentences on how Louie should position',
    '- Personal Angle: 1 sentence on the human hook',
    '- Company Intel JSON single-line:',
    '  {"industry":"...","products":"...","scale":"...","newsworthy":"...","marketingInsights":{"positioning":"...","goToMarket":"...","competitors":"...","teamSignals":"...","talkingPoints":"..."}}',
    '',
    'OUTPUT FORMAT: one block per job using exact delimiters:',
    '',
    '<<< RESULT rowId=N >>>',
    'FIT_SCORE: 8',
    'RECOMMENDATION: Apply',
    'KEY_NOTES: ...',
    'KEY_ALIGNMENTS: item 1 | item 2 | item 3',
    'POTENTIAL_CONCERNS: item 1 | item 2',
    'APPLICATION_PRIORITY: High',
    'TAILORED_PITCH: ...',
    'PERSONAL_ANGLE: ...',
    'COMPANY_INTEL: {"industry":"..."}',
    '<<< END RESULT N >>>',
    '',
    'IMPORTANT:',
    '- No em dashes anywhere',
    '- Australian English',
    '- Do not escape quotes inside values (use straight quotes; the parser handles them)',
    '- Each RESULT block must contain ALL fields, even if repeating existing values',
    '- Output only the RESULT blocks, no surrounding prose'
  ].join('\n');
}

function buildBatchGenerateDocsPrompt(rowIds) {
  if (!rowIds || !rowIds.length) return 'No row IDs provided.';
  const rows = rowIds.map(function (id) { return getRow(id).row; }).filter(Boolean);
  if (!rows.length) return 'No valid rows found for the given IDs.';

  const jobsBlock = rows.map(function (r) {
    const intel = r.companyIntel || {};
    const mi = intel.marketingInsights || {};
    return '<<< JOB rowId=' + r.rowId + ' >>>\n' +
           'ROLE: ' + (r.role || '') + '\n' +
           'COMPANY: ' + (r.company || '') + '\n' +
           'LOCATION: ' + (r.location || '') + '\n' +
           'SALARY: ' + (r.salary || '') + '\n' +
           'URL: ' + (r.url || '') + '\n' +
           'KEY NOTES: ' + (r.keyNotes || '') + '\n' +
           'KEY ALIGNMENTS: ' + (r.keyAlignments || '') + '\n' +
           'POTENTIAL CONCERNS: ' + (r.potentialConcerns || '') + '\n' +
           'TAILORED PITCH: ' + (r.tailoredPitch || '') + '\n' +
           'PERSONAL ANGLE: ' + (r.personalAngle || '') + '\n' +
           'COMPANY INTEL: ' + JSON.stringify({ industry: intel.industry, positioning: mi.positioning, goToMarket: mi.goToMarket, competitors: mi.competitors }) + '\n' +
           '<<< END JOB ' + r.rowId + ' >>>';
  }).join('\n\n');

  return [
    'Generate a tailored resume and cover letter for Louie Radburnd for each of the ' + rows.length + ' jobs below. Use Louie_Radburnd_Master_Resume_FINAL.docx in this project as the source of truth.',
    '',
    'JOBS:',
    '',
    jobsBlock,
    '',
    'RESUME REQUIREMENTS:',
    '- Tailor content, selecting most relevant experience and achievements per role',
    '- Prioritise measurable outcomes over responsibilities',
    '- Minimum 2 bullets per role, minimum 5 for MTAA and IndustraCom',
    '- Include all 7 roles (2008-2025)',
    '- Profile starts with "Most marketers specialise. I\'ve spent 18+ years doing all of it..."',
    '- Closing personal line in profile: clever, role-specific, one sentence, make reader smirk or nod',
    '- No em dashes, Australian English, 2 pages max',
    '',
    'COVER LETTER VOICE (locked):',
    '- Direct, conversational, slightly opinionated',
    '- First 1-2 paragraphs answer "why this role / why this company" explicitly',
    '- Concrete proof points from MTAA / IndustraCom with real numbers',
    '- Address obvious gaps on own terms',
    '- Include AI paragraph citing the n8n job pipeline with marketing-equivalent use cases',
    '- Forward-looking close',
    '- No em dashes, Australian English, 250-400 words, one A4 page',
    '- Salutation: "Hi Hiring Manager"',
    '- Sign off: "Louie Radburnd" only',
    '- No bold colon-format headings in the body',
    '',
    'OUTPUT FORMAT: one block per job using exact delimiters:',
    '',
    '<<< RESULT rowId=N >>>',
    '=== COVER LETTER ===',
    '[cover letter markdown]',
    '=== END COVER LETTER ===',
    '=== RESUME ===',
    '[resume markdown]',
    '=== END RESUME ===',
    '<<< END RESULT N >>>',
    '',
    'IMPORTANT:',
    '- Output only the RESULT blocks, no commentary between them',
    '- Each job gets its own complete resume + cover letter',
    '- Do not reference other jobs inside a RESULT block'
  ].join('\n');
}

function buildBatchInterviewPrepPrompt(rowIds) {
  if (!rowIds || !rowIds.length) return 'No row IDs provided.';
  const rows = rowIds.map(function (id) { return getRow(id).row; }).filter(Boolean);
  if (!rows.length) return 'No valid rows found for the given IDs.';

  const jobsBlock = rows.map(function (r) {
    const intel = r.companyIntel || {};
    const mi = intel.marketingInsights || {};
    return '<<< JOB rowId=' + r.rowId + ' >>>\n' +
           'ROLE: ' + (r.role || '') + '\n' +
           'COMPANY: ' + (r.company || '') + '\n' +
           'URL: ' + (r.url || '') + '\n' +
           'KEY NOTES: ' + (r.keyNotes || '') + '\n' +
           'KEY ALIGNMENTS: ' + (r.keyAlignments || '') + '\n' +
           'POTENTIAL CONCERNS: ' + (r.potentialConcerns || '') + '\n' +
           'TAILORED PITCH: ' + (r.tailoredPitch || '') + '\n' +
           'COMPANY INTEL: ' + JSON.stringify({ industry: intel.industry, positioning: mi.positioning, competitors: mi.competitors, teamSignals: mi.teamSignals }) + '\n' +
           '<<< END JOB ' + r.rowId + ' >>>';
  }).join('\n\n');

  return [
    'Generate interview prep for ' + rows.length + ' role' + (rows.length > 1 ? 's' : '') + ' Louie has applied for. Return one JSON object per job using the exact delimiter format.',
    '',
    'JOBS:',
    '',
    jobsBlock,
    '',
    'FOR EACH JOB, produce a JSON object with this shape:',
    '{',
    '  "generatedAt": "<ISO 8601>",',
    '  "questions": [',
    '    {"id":"q1","text":"...","context":"...","preparedAnswer":"...","status":"not_practiced","currentRound":0,"rounds":[]},',
    '    ...8-12 questions total',
    '  ],',
    '  "anticipatedQuestions": ["..."],',
    '  "companyDeepDive": {"culture":"...","recentNews":"...","leadershipSignals":"..."},',
    '  "talkingPoints": ["..."]',
    '}',
    '',
    'QUESTION MIX (aim for 8-12 per job):',
    '- 3-4 role-specific behavioural questions tied to the job\'s key alignments',
    '- 2-3 "why this company" / "why this role" angles',
    '- 1-2 addressing potential concerns openly',
    '- 2-3 classic leadership or competency questions relevant to the seniority',
    '',
    'ANTI-HALLUCINATION RULES:',
    '- Do NOT invent metrics, dates, or story details. Use only what is verified in Louie\'s career (18+ years, MTAA conference $214k to $815k, 200 to 16k LinkedIn, 90% renewal, 42% associate growth, IndustraCom -10 to +2 per month, 15% YoY revenue, 22% SEO lift, RegConnect 6-week build, etc.)',
    '- Leave preparedAnswer as a framework or prompt (e.g. "Use MTAA renewal story: situation/task/action/result"), not a fabricated STAR narrative',
    '- No em dashes, Australian English',
    '',
    'OUTPUT FORMAT: one block per job using exact delimiters:',
    '',
    '<<< RESULT rowId=N >>>',
    '{ valid JSON object }',
    '<<< END RESULT N >>>',
    '',
    'IMPORTANT:',
    '- Output only RESULT blocks',
    '- JSON must be valid and parseable, no trailing commas, no comments inside JSON'
  ].join('\n');
}


// ============================================================
// BATCH SAVE HANDLERS
// ============================================================

/**
 * Parse a batch response into per-rowId blocks.
 *
 * FIX 23-Apr-2026: The previous version used a regex backreference
 * `\1` to match the closing delimiter. On large responses (three jobs,
 * ~8kb each, all in one paste) that backreference caused catastrophic
 * regex backtracking in V8's engine — Apps Script ran past its execution
 * budget and the browser saw it as "Failed to fetch".
 *
 * New approach: find RESULT open delimiters linearly, then for each one
 * search forward for the MATCHING END delimiter by rowId. No backrefs,
 * no lazy-greedy traps. Order-agnostic and tolerant of nested delimiters.
 */
function parseBatchResponse(raw) {
  const result = {};
  if (!raw) return result;
  let text = String(raw);

  // Strip common markdown wrappers Claude sometimes adds — code fences
  // around the whole response, triple backticks on individual blocks, etc.
  // The parser is delimiter-based so the fences would prevent match.
  text = text.replace(/```[a-z]*\n?/gi, '').replace(/```/g, '');

  // Find all open tags with their rowId and position. Regex is tolerant
  // of extra whitespace and the rowId attribute being quoted/unquoted.
  const openRe = /<<<\s*RESULT\s+rowId\s*=\s*"?(\d+)"?\s*>>>/g;
  const opens = [];
  let m;
  while ((m = openRe.exec(text)) !== null) {
    opens.push({ rowId: m[1], start: m.index, contentStart: m.index + m[0].length });
  }

  // For each open, find its matching END RESULT N (tolerant matching)
  opens.forEach(function (o) {
    const endPattern = '<<<\\s*END\\s+RESULT\\s+' + o.rowId + '\\s*>>>';
    const endRe = new RegExp(endPattern);
    const slice = text.substring(o.contentStart);
    const endMatch = slice.match(endRe);
    if (endMatch) {
      const content = slice.substring(0, endMatch.index).trim();
      result[o.rowId] = content;
    }
  });

  return result;
}

/**
 * Coerce whatever recommendation value Claude produced into one of the
 * three values the sheet\'s data validation allows: APPLY, CONSIDER, SKIP.
 * Anything unknown defaults to CONSIDER — safer than failing the write.
 */
function normaliseRecommendation(raw) {
  if (!raw) return '';
  const v = String(raw).trim().toUpperCase();
  if (v === 'APPLY') return 'APPLY';
  if (v === 'CONSIDER') return 'CONSIDER';
  if (v === 'SKIP') return 'SKIP';
  // Map older / friendlier phrasings to the canonical three values
  if (v === 'STRONGLY APPLY' || v === 'STRONG APPLY' || v.indexOf('STRONG') > -1) return 'APPLY';
  if (v === 'PASS' || v === 'REJECT' || v === 'NO') return 'SKIP';
  if (v === 'MAYBE' || v === 'POSSIBLY' || v === 'UNSURE') return 'CONSIDER';
  // Fallback — if Claude sent something weird, default to CONSIDER so the
  // write doesn\'t get rejected by the data validation rule on col E.
  return 'CONSIDER';
}

/**
 * Coerce whatever application priority value Claude produced into one of
 * the three values the sheet\'s data validation allows: HIGH, MEDIUM, LOW.
 * Anything unknown defaults to MEDIUM.
 */
function normalisePriority(raw) {
  if (!raw) return '';
  const v = String(raw).trim().toUpperCase();
  if (v === 'HIGH' || v === 'H') return 'HIGH';
  if (v === 'MEDIUM' || v === 'MED' || v === 'M') return 'MEDIUM';
  if (v === 'LOW' || v === 'L') return 'LOW';
  // Map common variants
  if (v === 'URGENT' || v === 'TOP' || v === 'CRITICAL') return 'HIGH';
  if (v === 'MODERATE' || v === 'STANDARD' || v === 'NORMAL') return 'MEDIUM';
  if (v === 'MINIMAL' || v === 'MINOR' || v === 'OPTIONAL') return 'LOW';
  // Fallback
  return 'MEDIUM';
}

/**
 * Extract a FIELD: value from a batch result block. Values may span
 * multiple lines — the value continues until either the next "FIELD:"
 * line or the end of the block. This is more forgiving than requiring
 * Claude to produce single-line values for every field.
 */
function parseField(block, field) {
  // Match FIELD: at line start, capture everything up to either the next
  // ALLCAPS_WITH_UNDERSCORES: line or the end of the block. The lookahead
  // prevents the regex from eating the next field\'s label.
  const re = new RegExp('^' + field + ':\\s*([\\s\\S]*?)(?=\\n[A-Z_]+:|$)', 'm');
  const m = block.match(re);
  return m ? m[1].trim() : '';
}

function batchSaveAnalyse(body) {
  const parsed = parseBatchResponse(body.response || '');
  const rowIds = Object.keys(parsed);
  if (!rowIds.length) {
    // Return a preview of what came through so the user can see why parsing
    // failed. Common cases: response didn't include the delimiter tags,
    // rowId number didn't match, or Claude wrapped everything in prose.
    const preview = (body.response || '').substring(0, 240).replace(/\n/g, ' ').trim();
    return {
      ok: false,
      error: 'No valid RESULT blocks found. Check the response includes <<< RESULT rowId=N >>> delimiters. First 240 chars: "' + preview + '..."'
    };
  }

  const sheet = openPipeline();
  const saved = [];
  const skipped = [];

  rowIds.forEach(function (rowIdStr) {
    const rowId = parseInt(rowIdStr, 10);
    const block = parsed[rowIdStr];

    try {
      const fitScore = parseField(block, 'FIT_SCORE');
      const recommendation = normaliseRecommendation(parseField(block, 'RECOMMENDATION'));
      const keyNotes = parseField(block, 'KEY_NOTES');
      const keyAlignments = parseField(block, 'KEY_ALIGNMENTS');
      const potentialConcerns = parseField(block, 'POTENTIAL_CONCERNS');
      const applicationPriority = normalisePriority(parseField(block, 'APPLICATION_PRIORITY'));
      const tailoredPitch = parseField(block, 'TAILORED_PITCH');
      const personalAngle = parseField(block, 'PERSONAL_ANGLE');
      const companyIntel = parseField(block, 'COMPANY_INTEL');

      if (fitScore) sheet.getRange(rowId, COL.FIT_SCORE).setValue(fitScore);
      if (recommendation) sheet.getRange(rowId, COL.RECOMMENDATION).setValue(recommendation);
      if (keyNotes) sheet.getRange(rowId, COL.KEY_NOTES).setValue(keyNotes);
      if (keyAlignments) sheet.getRange(rowId, COL.KEY_ALIGNMENTS).setValue(keyAlignments);
      if (potentialConcerns) sheet.getRange(rowId, COL.POTENTIAL_CONCERNS).setValue(potentialConcerns);
      if (applicationPriority) sheet.getRange(rowId, COL.APPLICATION_PRIORITY).setValue(applicationPriority);
      if (tailoredPitch) sheet.getRange(rowId, COL.TAILORED_PITCH).setValue(tailoredPitch);
      if (personalAngle) sheet.getRange(rowId, COL.PERSONAL_ANGLE).setValue(personalAngle);
      if (companyIntel) sheet.getRange(rowId, COL.COMPANY_INTEL).setValue(companyIntel);

      sheet.getRange(rowId, COL.ANALYSIS_DATE).setValue(new Date());
      sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
      saved.push(rowId);
    } catch (e) {
      skipped.push({ rowId: rowId, error: String(e) });
    }
  });

  return { ok: true, saved: saved, skipped: skipped };
}

function batchSaveGenerateDocs(body) {
  const parsed = parseBatchResponse(body.response || '');
  const rowIds = Object.keys(parsed);
  if (!rowIds.length) return { ok: false, error: 'No valid RESULT blocks found in response' };

  const sheet = openPipeline();
  const saved = [];
  const skipped = [];

  rowIds.forEach(function (rowIdStr) {
    const rowId = parseInt(rowIdStr, 10);
    const block = parsed[rowIdStr];

    try {
      const clMatch = block.match(/=== COVER LETTER ===([\s\S]*?)=== END COVER LETTER ===/);
      const resMatch = block.match(/=== RESUME ===([\s\S]*?)=== END RESUME ===/);

      if (!clMatch || !resMatch) {
        skipped.push({ rowId: rowId, error: 'Missing cover letter or resume delimiters' });
        return;
      }

      const coverLetter = clMatch[1].trim();
      const resume = resMatch[1].trim();

      sheet.getRange(rowId, COL.COVER_LETTER_MD).setValue(coverLetter);
      sheet.getRange(rowId, COL.RESUME_MD).setValue(resume);
      setStatusWithTimestamp(sheet, rowId, 'Ready to Apply');
      saved.push(rowId);
    } catch (e) {
      skipped.push({ rowId: rowId, error: String(e) });
    }
  });

  return { ok: true, saved: saved, skipped: skipped };
}

function batchSaveInterviewPrep(body) {
  const parsed = parseBatchResponse(body.response || '');
  const rowIds = Object.keys(parsed);
  if (!rowIds.length) return { ok: false, error: 'No valid RESULT blocks found in response' };

  const sheet = openPipeline();
  const saved = [];
  const skipped = [];

  rowIds.forEach(function (rowIdStr) {
    const rowId = parseInt(rowIdStr, 10);
    const block = parsed[rowIdStr];

    try {
      JSON.parse(block);
      sheet.getRange(rowId, COL.INTERVIEW_PREP).setValue(block);
      sheet.getRange(rowId, COL.LAST_ACTIVITY).setValue(new Date());
      saved.push(rowId);
    } catch (e) {
      skipped.push({ rowId: rowId, error: 'Invalid JSON: ' + String(e) });
    }
  });

  return { ok: true, saved: saved, skipped: skipped };
}


// ============================================================
// UTILITIES
// ============================================================

function openPipeline() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(PIPELINE_TAB);
}

function openQueue() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(QUEUE_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_TAB);
    sheet.getRange(1, 1, 1, 4).setValues([['id', 'url', 'content', 'added_at']]);
  }
  return sheet;
}

function detectSource(url) {
  if (!url) return 'Other';
  const u = String(url).toLowerCase();
  if (u.indexOf('seek.com') > -1)     return 'Seek';
  if (u.indexOf('linkedin.com') > -1) return 'LinkedIn';
  if (u.indexOf('indeed.com') > -1)   return 'Indeed';
  return 'Other';
}

function parseJSONSafe(val) {
  if (!val) return null;
  if (typeof val === 'object') return val;
  try { return JSON.parse(String(val)); }
  catch (e) { return { _raw: String(val), _parseError: true }; }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function now() {
  return new Date().toISOString();
}

function generateId() {
  return 'q_' + Date.now().toString(36) + '_' + Math.random().toString(36).substr(2, 6);
}


// ============================================================
// TEST HARNESS
// ============================================================

function _testPing() {
  Logger.log(getPipeline());
}

function _testGetQueue() {
  Logger.log(getQueue());
}

function _testBuildAnalyse() {
  Logger.log(buildAnalysePrompt());
}

/**
 * Verify the new batch parser handles a multi-job response without hanging.
 * Run from Apps Script editor: function dropdown → _testBatchParser → Run.
 * Should log 3 rowIds and their content lengths in under a second.
 */
function _testBatchParser() {
  const sample =
    '<<< RESULT rowId=15 >>>\nFIT_SCORE: 7.2\nRECOMMENDATION: Apply\n<<< END RESULT 15 >>>\n' +
    '<<< RESULT rowId=30 >>>\nFIT_SCORE: 7.8\nRECOMMENDATION: Apply\n<<< END RESULT 30 >>>\n' +
    '<<< RESULT rowId=60 >>>\nFIT_SCORE: 7.5\nRECOMMENDATION: Apply\n<<< END RESULT 60 >>>';
  const parsed = parseBatchResponse(sample);
  Logger.log('Parsed rowIds: ' + Object.keys(parsed).join(','));
  Object.keys(parsed).forEach(function (k) {
    Logger.log('  ' + k + ' — ' + parsed[k].length + ' chars');
  });
  return parsed;
}


// ============================================================
// ONE-OFF BACKFILL (run once after adding STAGE_ENTERED_AT column)
// ============================================================

function _backfillStageEnteredAt() {
  const sheet = openPipeline();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows'); return; }

  const range = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS);
  const values = range.getValues();
  let updated = 0;
  let skipped = 0;

  values.forEach(function (r, i) {
    const rowNum = i + 2;
    if (r[COL.STAGE_ENTERED_AT - 1]) { skipped++; return; }

    const status = r[COL.STATUS - 1];
    let proxy = null;

    if (['Applied', 'Interviewing', 'Interviewed', 'Offer', 'Accepted'].indexOf(status) > -1) {
      proxy = r[COL.APPLICATION_DATE - 1] || r[COL.LAST_ACTIVITY - 1];
    } else if (status === 'Analysed') {
      proxy = r[COL.ANALYSIS_DATE - 1] || r[COL.DATE_ADDED - 1] || r[COL.LAST_ACTIVITY - 1];
    } else {
      proxy = r[COL.LAST_ACTIVITY - 1] || r[COL.DATE_ADDED - 1];
    }

    if (proxy) {
      sheet.getRange(rowNum, COL.STAGE_ENTERED_AT).setValue(proxy);
      updated++;
    }
  });

  Logger.log('Backfill complete: ' + updated + ' rows updated, ' + skipped + ' already had value');
  return { updated: updated, skipped: skipped };
}


// ============================================================
// ONE-OFF FIT SCORE BACKFILL
// ============================================================

/**
 * One-off: normalise all Fit Score cells to integer 0–10.
 *
 * Historical rows were analysed on a 0–100 scale (72, 48, 78, 82) or
 * stored as decimals (7.8, 8.5). The canonical scale is now INTEGER
 * 0–10 with no decimals. This function:
 *   - Divides any value >10 by 10 (0–100 normalisation)
 *   - Rounds all values to the nearest whole number
 *   - Leaves blanks and non-numerics alone
 *
 * SAFE TO RE-RUN: values already at integer 0–10 round to themselves,
 * so running twice is a no-op.
 *
 * Run from Apps Script editor: function dropdown → _backfillFitScores → Run.
 * Check Logs (View → Logs, or Cmd+Enter) for a summary of what changed.
 */
function _backfillFitScores() {
  const sheet = openPipeline();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows'); return; }

  const range = sheet.getRange(2, COL.FIT_SCORE, lastRow - 1, 1);
  const values = range.getValues();
  const changes = [];
  let updated = 0;
  let skipped = 0;
  let blank = 0;

  const newValues = values.map(function (row, i) {
    const rowNum = i + 2;
    const raw = row[0];

    // Blank cells stay blank
    if (raw === '' || raw === null || raw === undefined) { blank++; return ['']; }

    const n = parseFloat(raw);
    if (isNaN(n)) { skipped++; return [raw]; }

    // Normalise 0–100 scale to 0–10 if needed, then round to integer
    const onTenScale = n > 10 ? n / 10 : n;
    const normalised = Math.round(onTenScale);

    // Already a correct integer — leave alone
    if (normalised === n) { skipped++; return [raw]; }

    changes.push({ row: rowNum, from: n, to: normalised });
    updated++;
    return [normalised];
  });

  range.setValues(newValues);

  Logger.log('Fit score backfill complete');
  Logger.log('  Updated: ' + updated);
  Logger.log('  Skipped (already correct integer, or non-numeric): ' + skipped);
  Logger.log('  Blank: ' + blank);
  if (changes.length) {
    Logger.log('Changes:');
    changes.forEach(function (c) {
      Logger.log('  Row ' + c.row + ': ' + c.from + ' → ' + c.to);
    });
  }
  return { updated: updated, skipped: skipped, blank: blank, changes: changes };
}


// ============================================================
// ONE-OFF STATUS HISTORY BACKFILL
// ============================================================

/**
 * One-off: reconstruct STATUS_HISTORY (col 38) from existing date columns.
 *
 * For each row we infer a plausible journey using the dates we DO have:
 *   - DATE_ADDED      → first event (status: 'Analysed')
 *   - APPLICATION_DATE → 'Applied' event (if present)
 *   - STAGE_ENTERED_AT → current status event (if newer than the others)
 *
 * Only fills cells where STATUS_HISTORY is currently blank. Re-runnable.
 *
 * Run from Apps Script editor: function dropdown → _backfillStatusHistory → Run.
 *
 * Note: this is best-effort reconstruction. We can't recover transitions
 * we have no dates for — e.g. if a card went Analysed → Ready to Apply →
 * Applied with no Ready-to-Apply timestamp, only Analysed and Applied
 * will show. From now on, all transitions are captured live.
 */
function _backfillStatusHistory() {
  const sheet = openPipeline();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('No data rows'); return; }

  const range = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS);
  const values = range.getValues();
  let updated = 0;
  let skipped = 0;

  values.forEach(function (r, i) {
    const rowNum = i + 2;
    const existing = r[COL.STATUS_HISTORY - 1];
    if (existing) { skipped++; return; }

    const status        = r[COL.STATUS - 1];
    const dateAdded     = r[COL.DATE_ADDED - 1];
    const applicationAt = r[COL.APPLICATION_DATE - 1];
    const stageEntered  = r[COL.STAGE_ENTERED_AT - 1];
    const lastActivity  = r[COL.LAST_ACTIVITY - 1];

    const events = [];

    // Genesis event — always Analysed at date added
    if (dateAdded) {
      events.push({ status: 'Analysed', at: toIso(dateAdded) });
    }

    // Applied event — for any status that's been through Applied
    if (applicationAt && ['Applied', 'Interviewing', 'Interviewed', 'Offer', 'Accepted', 'Rejected', 'Withdrawn'].indexOf(status) > -1) {
      events.push({ status: 'Applied', at: toIso(applicationAt) });
    }

    // Current status event — only if it's not already represented
    const finalAt = toIso(stageEntered || lastActivity || new Date());
    if (status && (!events.length || events[events.length - 1].status !== status)) {
      events.push({ status: status, at: finalAt });
    }

    if (events.length) {
      sheet.getRange(rowNum, COL.STATUS_HISTORY).setValue(JSON.stringify(events));
      updated++;
    }
  });

  Logger.log('Status history backfill complete');
  Logger.log('  Updated: ' + updated);
  Logger.log('  Skipped (already had history): ' + skipped);
  return { updated: updated, skipped: skipped };
}

function toIso(d) {
  if (!d) return null;
  if (d instanceof Date) return d.toISOString();
  const parsed = new Date(d);
  return isNaN(parsed.getTime()) ? null : parsed.toISOString();
}


// ============================================================
// ONE-OFF COLUMN FORMATTING
// ============================================================

function _enforceColumnFormats() {
  const sheet = openPipeline();
  const lastRow = Math.max(2, sheet.getLastRow());
  const numRows = lastRow - 1;

  const dateCols = [
    COL.ANALYSIS_DATE,
    COL.LAST_ACTIVITY,
    COL.APPLICATION_DATE,
    COL.FOLLOW_UP_DATE,
    COL.DATE_ADDED,
    COL.STAGE_ENTERED_AT
  ];
  const dateFormat = 'd/MM/yyyy HH:mm';

  dateCols.forEach(function (colIdx) {
    const range = sheet.getRange(2, colIdx, numRows, 1);
    range.setNumberFormat(dateFormat);
  });

  const textCols = [
    COL.KEY_NOTES,
    COL.COMPANY_INTEL,
    COL.KEY_ALIGNMENTS,
    COL.POTENTIAL_CONCERNS,
    COL.TAILORED_PITCH,
    COL.CULTURAL_FIT,
    COL.RESUME_MD,
    COL.COVER_LETTER_MD,
    COL.INTERVIEW_PREP,
    COL.YOUR_NOTES,
    COL.PERSONAL_ANGLE,
    COL.AD_CONTENT,
    COL.STATUS_HISTORY
  ];

  textCols.forEach(function (colIdx) {
    const range = sheet.getRange(2, colIdx, numRows, 1);
    range.setNumberFormat('@');
    range.setWrap(true);
    range.setVerticalAlignment('top');
  });

  Logger.log('Column formats enforced across ' + numRows + ' data rows');
  return { ok: true, rows: numRows };
}