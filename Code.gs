// ================================================================
// Fortify Health -- Compliance Email System v3.1 (Production)
// ================================================================
// INSTALL: Extensions > Apps Script > paste > Save > Reload sheet
// USE: Fortify Health menu in the toolbar
// ================================================================
//
// INDEX LOGIC (documented to prevent off-by-one bugs):
//
//   getDataRange().getValues() returns a 2D array:
//     data[0]  = header row   = sheet row 1
//     data[1]  = first record = sheet row 2
//     data[i]  = record       = sheet row (i + 1)
//
//   So to read a cell from sheet for data[i]:
//     dataSheet.getRange(i + 1, columnIndex + 1)
//                        ^^^^^  ^^^^^^^^^^^^^^^^
//                     sheet row  1-based column
//
// ================================================================

var CONFIG = {
  DATA_SHEET:       "Mill_Compliance_Data",
  CONTACT_SHEET:    "Mill_Contacts",
  MANAGER_SHEET:    "Field_Managers",
  LOG_SHEET:        "Draft_Log",
  SETTINGS_SHEET:   "System_Settings",

  BATCH_SIZE:         20,
  CHECKPOINT_KEY:     "fh_last_processed_idx",
  FSSAI_PPM_STD:      40,
  PREMIX_RATIO:       0.15,
  GMAIL_DAILY_LIMIT:  1400,
  MAX_RETRIES:        3,
  RETRY_DELAY_MS:     2000,
  SEND_JITTER_MS:     1500,
  FOLLOWUP_DAYS:      { HIGH: 3, MEDIUM: 7, LOW: 30 },
  ADMIN_EMAIL:        "admin@fortifyhealth.in",
  ADMIN_BCC:          false,
  AUTO_SEND:          false,
  CC_MANAGER:         true,
};


// ================================================================
//  MENU (no emojis -- clean professional labels)
// ================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Fortify Health")
    .addItem("Draft compliance emails", "draftComplianceEmails")
    .addItem("Auto-send compliance emails", "autoSendComplianceEmails")
    .addItem("Send all pending drafts", "sendAllPendingDrafts")
    .addSeparator()
    .addSubMenu(ui.createMenu("Schedule Triggers")
      .addItem("Custom date and time trigger...", "showCustomTriggerDialog")
      .addSeparator()
      .addItem("Weekly HIGH tier (Mon 09:00)", "setupWeeklyHighTrigger")
      .addItem("Bi-weekly MEDIUM (Wed 14:00)", "setupBiweeklyMediumTrigger")
      .addItem("Monthly LOW auto-send (1st 10:00)", "setupMonthlyLowTrigger")
      .addItem("Daily auto-send (09:00)", "setupDailyAutoSend")
      .addSeparator()
      .addItem("Escalation check (Fri 17:00)", "setupEscalationCheck")
      .addItem("Weekly digest (Fri 18:00)", "setupWeeklyDigestTrigger")
      .addSeparator()
      .addItem("View all active triggers", "viewActiveTriggers")
      .addItem("Remove ALL triggers", "removeAllTriggers"))
    .addSeparator()
    .addSubMenu(ui.createMenu("Reports")
      .addItem("Check escalations now", "checkEscalations")
      .addItem("Send weekly digest now", "sendWeeklyDigest")
      .addItem("Generate mill report", "generateMillReport")
      .addItem("Export log to CSV", "exportLogToCsv"))
    .addSeparator()
    .addSubMenu(ui.createMenu("Data Tools")
      .addItem("Recalculate all drift labels", "recalcAllDrift")
      .addItem("Reset all email flags to REVIEW", "resetAllFlags")
      .addItem("Validate all rows (dry run)", "validateAllRows")
      .addItem("Count mills by tier", "countByTier")
      .addItem("Find missing contacts", "findMissingContacts"))
    .addSeparator()
    .addSubMenu(ui.createMenu("One-Time Send")
      .addItem("Send to specific mill...", "sendToSpecificMill")
      .addItem("Send to specific tier...", "sendToSpecificTier")
      .addItem("Schedule one-time at date/time...", "scheduleOneTimeSend"))
    .addSeparator()
    .addItem("Reset batch checkpoint", "resetCheckpoint")
    .addToUi();
}


// ================================================================
//  ENTRY POINTS
// ================================================================

function draftComplianceEmails() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { _alert("Another run is in progress. Please wait."); return; }
  try { _runBatch(false, null); }
  finally { lock.releaseLock(); }
}

function autoSendComplianceEmails() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Logger.log("Locked out"); return; }
  try { _runBatch(true, null); }
  finally { lock.releaseLock(); }
}

function triggerHighTier()            { _lockedBatch(false, "HIGH"); }
function triggerMediumTier()          { _lockedBatch(false, "MEDIUM"); }
function triggerLowTier()             { _lockedBatch(false, "LOW"); }
function triggerAutoSendHighTier()    { _lockedBatch(true,  "HIGH"); }
function triggerAutoSendMediumTier()  { _lockedBatch(true,  "MEDIUM"); }
function triggerAutoSendLowTier()     { _lockedBatch(true,  "LOW"); }

function _lockedBatch(forceSend, tierFilter) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try { _runBatch(forceSend, tierFilter); }
  finally { lock.releaseLock(); }
}


// ================================================================
//  CORE BATCH PROCESSOR
// ================================================================

function _runBatch(forceSend, tierFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = loadSettings(ss);

  var dataSheet    = ss.getSheetByName(config.DATA_SHEET);
  var contactSheet = ss.getSheetByName(config.CONTACT_SHEET);
  if (!dataSheet)    { _alert("Sheet not found: " + config.DATA_SHEET); return; }
  if (!contactSheet) { _alert("Sheet not found: " + config.CONTACT_SHEET); return; }

  var logSheet = getOrCreateLogSheet(ss);
  var managers = buildLookupMap(ss.getSheetByName(config.MANAGER_SHEET));
  var contacts = buildContactMap(contactSheet);

  var data    = dataSheet.getDataRange().getValues();
  var headers = data[0];

  // Build column name -> 0-based index map
  var colMap = {};
  for (var h = 0; h < headers.length; h++) {
    colMap[String(headers[h]).trim()] = h;
  }

  // Verify required columns exist
  var requiredCols = ["Mill ID", "Mill Name", "State", "Risk Tier", "Email Flag",
                      "M1 Compliance pct", "M2 Compliance pct", "M3 Compliance pct",
                      "Lab Result ppm Iron", "Production Volume MT",
                      "Premix Used kg", "Expected Premix kg", "Premix Deviation pct",
                      "Lab Variance pct", "Days Since Audit", "Reporting Month"];
  for (var rc = 0; rc < requiredCols.length; rc++) {
    if (colMap[requiredCols[rc]] === undefined) {
      _alert("Required column missing: \"" + requiredCols[rc] + "\"\n\nCheck your Mill_Compliance_Data headers.");
      return;
    }
  }

  var col = function(name) { return colMap[name]; };
  var flagColZero = col("Email Flag");  // 0-based index

  // Gmail quota guard
  var quota = MailApp.getRemainingDailyQuota();
  if (quota < 10) {
    _alert("Gmail daily quota too low (" + quota + " remaining). Try again tomorrow.");
    return;
  }

  var props = PropertiesService.getScriptProperties();
  // startIdx = index into data[] array; 1 = first data row
  var startIdx = tierFilter ? 1 : parseInt(props.getProperty(config.CHECKPOINT_KEY) || "1", 10);
  if (startIdx < 1) startIdx = 1;

  var drafted = 0, sent = 0, skipped = 0, errors = [];
  var shouldSend = forceSend || config.AUTO_SEND;

  // ── MAIN LOOP ──
  for (var i = startIdx; i < data.length; i++) {
    var row      = data[i];           // data[i] = the record
    var sheetRow = i + 1;             // sheet row number (1-based)

    var millId = String(row[col("Mill ID")] || "").trim();
    if (!millId) { continue; }  // skip truly empty rows silently

    var millName = String(row[col("Mill Name")] || "").trim();
    var state    = String(row[col("State")] || "").trim();

    // ── Tier filter ──
    var rawTier = normaliseTier(row[col("Risk Tier")]);
    if (tierFilter && rawTier !== tierFilter) { continue; }

    // ── Read flag LIVE from sheet (not from stale data[] cache) ──
    var flagCell  = dataSheet.getRange(sheetRow, flagColZero + 1);
    var emailFlag = String(flagCell.getValue()).trim();

    // SKIP already-processed rows
    if (emailFlag.indexOf("DRAFTED") === 0 ||
        emailFlag.indexOf("SENT") === 0 ||
        emailFlag.indexOf("DATA ERROR") === 0) {
      skipped++;
      continue;
    }

    // All other values (REVIEW, empty, blank, anything) = process this row

    // ── Validation ──
    var validation = validateRow(row, col, millName, config);
    if (!validation.valid) {
      errors.push(millName + ": " + validation.reason);
      flagCell.setValue("DATA ERROR - " + validation.reason);
      flagCell.setBackground("#FCE4D6");
      skipped++;
      continue;
    }

    // ── Contact lookup ──
    var contact = contacts[millId];
    if (!contact) {
      errors.push(millName + " (" + millId + "): no matching contact in Mill_Contacts");
      flagCell.setValue("DATA ERROR - no contact found");
      flagCell.setBackground("#FCE4D6");
      skipped++;
      continue;
    }

    // ── Tier check ──
    if (!rawTier) {
      errors.push(millName + ": unrecognised tier \"" + row[col("Risk Tier")] + "\"");
      flagCell.setValue("DATA ERROR - invalid tier");
      flagCell.setBackground("#FCE4D6");
      skipped++;
      continue;
    }

    var manager = managers[state] || { name: "Fortify Health Field Team", email: null };
    var payload = extractPayload(row, col, rawTier, manager.name);
    var email   = buildEmail(payload, contact.name, config);

    // Per-iteration quota check
    if (MailApp.getRemainingDailyQuota() < 5) {
      errors.push("Gmail quota exhausted -- stopped early");
      if (!tierFilter) props.setProperty(config.CHECKPOINT_KEY, String(i));
      _finish(drafted, sent, skipped, errors, true, sheetRow, data.length);
      return;
    }

    // ── SEND OR DRAFT ──
    try {
      var opts = { htmlBody: email.htmlBody, name: "Fortify Health Programme" };
      if (config.CC_MANAGER && manager.email) { opts.cc = manager.email; }
      if (config.ADMIN_BCC && config.ADMIN_EMAIL) { opts.bcc = config.ADMIN_EMAIL; }

      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

      if (shouldSend) {
        _sendWithRetry(contact.email, email.subject, email.body, opts, config);
        sent++;
        flagCell.setValue("SENT - " + timestamp);
        flagCell.setBackground("#D5F5E3");
      } else {
        GmailApp.createDraft(contact.email, email.subject, email.body, opts);
        drafted++;
        flagCell.setValue("DRAFTED - " + timestamp);
        flagCell.setBackground("#E2EFDA");
      }

      appendToLog(logSheet, {
        timestamp: timestamp, millId: millId, millName: millName, tier: rawTier,
        emailType: emailFlag || "AUTO", subject: email.subject,
        fieldManager: manager.name, contactEmail: contact.email,
        followUpDue: getFollowUpDate(rawTier, config),
        drift: payload.drift, status: shouldSend ? "Sent" : "Drafted"
      });

      // Rate-limit jitter between emails
      if (config.SEND_JITTER_MS > 0) {
        Utilities.sleep(Math.floor(Math.random() * config.SEND_JITTER_MS));
      }

    } catch (e) {
      errors.push(millName + ": " + e.message);
    }

    // Batch checkpoint (pause after N emails)
    if ((drafted + sent) > 0 && (drafted + sent) % config.BATCH_SIZE === 0) {
      if (!tierFilter) props.setProperty(config.CHECKPOINT_KEY, String(i + 1));
      SpreadsheetApp.flush();
      _finish(drafted, sent, skipped, errors, true, sheetRow, data.length);
      return;
    }
  }

  // Done -- clear checkpoint
  if (!tierFilter) props.deleteProperty(config.CHECKPOINT_KEY);
  SpreadsheetApp.flush();
  _finish(drafted, sent, skipped, errors, false, 0, 0);
}


// ================================================================
//  SEND WITH RETRY (exponential backoff)
// ================================================================

function _sendWithRetry(to, subject, body, options, config) {
  var lastError;
  for (var attempt = 1; attempt <= config.MAX_RETRIES; attempt++) {
    try {
      MailApp.sendEmail(to, subject, body, options);
      return;
    } catch (e) {
      lastError = e;
      if (attempt < config.MAX_RETRIES) {
        Utilities.sleep(config.RETRY_DELAY_MS * attempt);
      }
    }
  }
  throw lastError;
}


// ================================================================
//  VALIDATION
// ================================================================

function validateRow(row, col, millName, config) {
  var parseNum = function(field) {
    var raw = row[col(field)];
    if (raw === null || raw === undefined) return NaN;
    // If it is already a number, return it directly
    if (typeof raw === "number") return raw;
    var s = String(raw).replace(/%/g, "").replace(/,/g, "").replace(/\+/g, "").trim();
    if (s === "" || s === "N/A" || s === "-" || s === "--") return NaN;
    var n = parseFloat(s);
    return n;
  };

  var prodVol    = parseNum("Production Volume MT");
  var premixUsed = parseNum("Premix Used kg");
  var premixExp  = parseNum("Expected Premix kg");
  var labPpm     = parseNum("Lab Result ppm Iron");
  var m1 = parseNum("M1 Compliance pct");
  var m2 = parseNum("M2 Compliance pct");
  var m3 = parseNum("M3 Compliance pct");

  if (isNaN(prodVol) || prodVol <= 0)               return { valid: false, reason: "Production volume missing or zero" };
  if (isNaN(premixUsed) || premixUsed < 0)          return { valid: false, reason: "Premix usage invalid" };
  if (isNaN(premixExp) || premixExp <= 0)           return { valid: false, reason: "Expected premix missing" };
  if (isNaN(labPpm) || labPpm < 0 || labPpm > 200) return { valid: false, reason: "Lab ppm out of range (" + labPpm + ")" };
  if (!isNaN(m1) && (m1 < 0 || m1 > 100))          return { valid: false, reason: "M1 out of range (" + m1 + ")" };
  if (!isNaN(m2) && (m2 < 0 || m2 > 100))          return { valid: false, reason: "M2 out of range (" + m2 + ")" };
  if (!isNaN(m3) && (m3 < 0 || m3 > 100))          return { valid: false, reason: "M3 out of range (" + m3 + ")" };

  if (prodVol > 0) {
    var ratio = premixUsed / prodVol;
    if (ratio > config.PREMIX_RATIO * 3)
      return { valid: false, reason: "Premix ratio implausible (" + ratio.toFixed(3) + ")" };
  }

  return { valid: true };
}


// ================================================================
//  TIER, DRIFT, PAYLOAD
// ================================================================

function normaliseTier(raw) {
  var t = String(raw || "").trim().toUpperCase();
  if (t === "HIGH" || t === "MEDIUM" || t === "LOW") return t;
  return null;
}

function getDriftLabel(m1, m2, m3) {
  if (m1 === null || m2 === null || m3 === null) return "Unknown";
  if (isNaN(m1) || isNaN(m2) || isNaN(m3)) return "Unknown";
  var d1 = m2 - m3, d2 = m1 - m2;
  if (d1 < -5 && d2 < -5) return "Declining";
  if (d1 > 3 && d2 > 3) return "Improving";
  if (Math.abs(d1) <= 3 && Math.abs(d2) <= 3) return "Stable";
  return "Inconsistent";
}

function extractPayload(row, col, tier, fieldManager) {
  var parseNum = function(f) {
    var raw = row[col(f)];
    if (typeof raw === "number") return raw;
    var v = parseFloat(String(raw || "").replace(/%/g, "").replace(/,/g, "").replace(/\+/g, ""));
    return isNaN(v) ? null : v;
  };
  var m1 = parseNum("M1 Compliance pct");
  var m2 = parseNum("M2 Compliance pct");
  var m3 = parseNum("M3 Compliance pct");
  return {
    millName:     row[col("Mill Name")],
    state:        row[col("State")],
    month:        row[col("Reporting Month")],
    premixUsed:   row[col("Premix Used kg")],
    premixExp:    row[col("Expected Premix kg")],
    premixDev:    row[col("Premix Deviation pct")],
    labPpm:       row[col("Lab Result ppm Iron")],
    labVar:       row[col("Lab Variance pct")],
    daysSince:    row[col("Days Since Audit")],
    m1: m1, m2: m2, m3: m3,
    drift:        getDriftLabel(m1, m2, m3),
    tier:         tier,
    fieldManager: fieldManager,
  };
}


// ================================================================
//  EMAIL BUILDER (plain text + HTML)
// ================================================================

function buildEmail(p, contactName, config) {
  var subject, body;

  if (p.tier === "HIGH") {
    var auditNote = (p.daysSince && p.daysSince > 120)
      ? "\n\nWe also note your last verified audit was " + p.daysSince + " days ago. We would like to include a field visit as part of this review."
      : "";
    subject = "Urgent: Fortification quality review required - " + p.millName + ", " + p.month;
    body = "Dear " + contactName + ",\n\n" +
      "I am writing because our " + p.month + " data for " + p.millName + " has flagged an issue that needs immediate attention.\n\n" +
      "Your iron premix usage was " + p.premixUsed + " kg against an expected " + p.premixExp + " kg (" + p.premixDev + " shortfall). " +
      "The lab result of " + p.labPpm + " ppm is below the FSSAI standard of " + config.FSSAI_PPM_STD + " ppm. " +
      "Given the three-month trend of " + p.m3 + "% > " + p.m2 + "% > " + p.m1 + "% (" + p.drift.toLowerCase() + "), " +
      "this suggests a calibration issue rather than a one-time anomaly." + auditNote + "\n\n" +
      "We would like to schedule a 30-minute technical call this week. Could you share your availability?\n\n" +
      "Warm regards,\n" + p.fieldManager + "\nFortify Health";

  } else if (p.tier === "MEDIUM") {
    var dn = (p.drift === "Declining")
      ? "Your trend (" + p.m3 + "% > " + p.m2 + "% > " + p.m1 + "%) is moving in the wrong direction."
      : "Your trend (" + p.m3 + "% > " + p.m2 + "% > " + p.m1 + "%) is " + p.drift.toLowerCase() + " and worth monitoring.";
    subject = "Fortification data review - action recommended - " + p.millName + ", " + p.month;
    body = "Dear " + contactName + ",\n\n" +
      "Thank you for your continued participation in the fortification programme. Our " + p.month + " data for " + p.millName + " is worth a brief conversation.\n\n" +
      "Your overall compliance is in an acceptable range, but " + dn + " The lab result of " + p.labPpm + " ppm shows a " + p.labVar + " variance from the " + config.FSSAI_PPM_STD + " ppm standard.\n\n" +
      "I would suggest a 20-minute call to review the numbers together. Would later this week work?\n\n" +
      "Best regards,\n" + p.fieldManager + "\nFortify Health";

  } else if (p.m1 !== null && p.m1 >= 95) {
    subject = "Strong fortification performance - " + p.millName + ", " + p.month;
    body = "Dear " + contactName + ",\n\n" +
      "A quick note to recognise " + p.millName + "'s performance in " + p.month + ".\n\n" +
      "Your compliance rate of " + p.m1 + "% is above programme target, lab results are at " + p.labPpm + " ppm, and premix usage is accurate.\n\n" +
      "No action required. We wanted to make sure good work is acknowledged.\n\n" +
      "Best,\n" + p.fieldManager + "\nFortify Health";

  } else {
    subject = "Monthly check-in - " + p.millName + " fortification data, " + p.month;
    body = "Dear " + contactName + ",\n\n" +
      "A brief note following our review of " + p.month + " data for " + p.millName + ".\n\n" +
      "Your compliance and lab results are in the acceptable range. Three-month trend: " +
      p.m3 + "% > " + p.m2 + "% > " + p.m1 + "% (" + p.drift.toLowerCase() + ").\n\n" +
      "If anything has changed in your production process or premix supplier, please let us know.\n\n" +
      "Best regards,\n" + p.fieldManager + "\nFortify Health";
  }

  return { subject: subject, body: body, htmlBody: _htmlEmail(body, p, config) };
}

function _htmlEmail(plain, p, config) {
  var tc = ({ HIGH: "#DC2626", MEDIUM: "#D97706", LOW: "#059669" })[p.tier] || "#6B7280";
  var pc = (p.labPpm < config.FSSAI_PPM_STD) ? "#DC2626" : "#059669";
  var bh = plain.replace(/\n\n/g, '</p><p style="margin:0 0 14px 0;">').replace(/\n/g, "<br>");
  return '<html><body style="margin:0;padding:0;background:#f8f9fa;font-family:Arial,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;margin:0 auto;background:#fff;">' +
    '<tr><td style="background:linear-gradient(135deg,#1E3A5F,#2563EB);padding:24px 32px;">' +
    '<table width="100%"><tr><td style="color:#fff;font-size:18px;font-weight:700;">Fortify Health</td>' +
    '<td align="right"><span style="padding:4px 12px;border-radius:20px;font-size:11px;font-weight:600;color:#fff;background:' + tc + ';">' + p.tier + '</span></td></tr></table></td></tr>' +
    '<tr><td style="padding:14px 32px;background:#f1f5f9;border-bottom:1px solid #e2e8f0;font-size:12px;color:#64748b;">' +
    '<strong>' + p.millName + '</strong> | ' + p.month + ' | <span style="color:' + pc + ';font-weight:600;">' + p.labPpm + ' ppm</span> | Trend: ' + p.m3 + ' > ' + p.m2 + ' > ' + p.m1 + '% | ' + p.drift + '</td></tr>' +
    '<tr><td style="padding:28px 32px;font-size:14px;line-height:1.7;color:#374151;"><p style="margin:0 0 14px 0;">' + bh + '</p></td></tr>' +
    '<tr><td style="padding:16px 32px;background:#f8f9fa;border-top:1px solid #e2e8f0;font-size:11px;color:#9ca3af;text-align:center;">' +
    'Fortify Health | Wheat Flour Fortification Programme | India | Follow-up: ' + (config.FOLLOWUP_DAYS[p.tier] || 7) + ' days</td></tr></table></body></html>';
}


// ================================================================
//  CUSTOM DATE/TIME TRIGGER DIALOG
// ================================================================

function showCustomTriggerDialog() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial;padding:16px;background:#f8f9fa;}' +
    'h3{color:#1F4E79;margin:0 0 16px 0;}' +
    'label{display:block;margin:12px 0 4px;font-size:12px;font-weight:600;color:#374151;}' +
    'select,input{width:100%;padding:8px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;box-sizing:border-box;background:#fff;}' +
    '.row{display:flex;gap:10px;}.row>div{flex:1;}' +
    'button{margin-top:16px;padding:10px;background:#2563EB;color:#fff;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;width:100%;}' +
    'button:hover{background:#1d4ed8;}' +
    '.note{font-size:11px;color:#6b7280;margin-top:10px;}' +
    '#status{margin-top:10px;padding:8px;border-radius:4px;display:none;font-size:12px;}' +
    '</style>' +
    '<h3>Custom Schedule Trigger</h3>' +
    '<label>Action</label><select id="action">' +
    '<option value="draftComplianceEmails">Draft all emails</option>' +
    '<option value="autoSendComplianceEmails">Auto-send all emails</option>' +
    '<option value="triggerHighTier">Draft HIGH tier only</option>' +
    '<option value="triggerMediumTier">Draft MEDIUM tier only</option>' +
    '<option value="triggerLowTier">Draft LOW tier only</option>' +
    '<option value="triggerAutoSendHighTier">Auto-send HIGH tier</option>' +
    '<option value="triggerAutoSendMediumTier">Auto-send MEDIUM tier</option>' +
    '<option value="triggerAutoSendLowTier">Auto-send LOW tier</option>' +
    '<option value="checkEscalations">Check escalations</option>' +
    '<option value="sendWeeklyDigest">Send weekly digest</option>' +
    '</select>' +
    '<label>Frequency</label><select id="freq" onchange="toggle()">' +
    '<option value="once">One-time (specific date and time)</option>' +
    '<option value="daily">Daily</option>' +
    '<option value="weekly">Weekly</option>' +
    '<option value="monthly">Monthly</option></select>' +
    '<div id="dateDiv"><label>Date</label><input type="date" id="date"/></div>' +
    '<div id="dayDiv" style="display:none"><label>Day of week</label><select id="day">' +
    '<option value="1">Monday</option><option value="2">Tuesday</option><option value="3">Wednesday</option>' +
    '<option value="4">Thursday</option><option value="5">Friday</option><option value="6">Saturday</option>' +
    '<option value="7">Sunday</option></select></div>' +
    '<div id="domDiv" style="display:none"><label>Day of month (1-28)</label>' +
    '<input type="number" id="dom" min="1" max="28" value="1"/></div>' +
    '<div class="row">' +
    '<div><label>Hour (0-23)</label><input type="number" id="hour" min="0" max="23" value="9"/></div>' +
    '<div><label>Minute (0-59)</label><input type="number" id="min" min="0" max="59" value="0"/></div></div>' +
    '<button onclick="go()">Create Trigger</button>' +
    '<div id="status"></div>' +
    '<div class="note">Times use your Apps Script project timezone. Set it in Project Settings.</div>' +
    '<script>' +
    'function toggle(){var f=document.getElementById("freq").value;' +
    'document.getElementById("dateDiv").style.display=f==="once"?"block":"none";' +
    'document.getElementById("dayDiv").style.display=f==="weekly"?"block":"none";' +
    'document.getElementById("domDiv").style.display=f==="monthly"?"block":"none";}' +
    'function go(){' +
    'document.getElementById("status").style.display="block";' +
    'document.getElementById("status").style.background="#DBEAFE";' +
    'document.getElementById("status").innerText="Creating trigger...";' +
    'var o={action:document.getElementById("action").value,' +
    'freq:document.getElementById("freq").value,' +
    'date:document.getElementById("date").value,' +
    'day:parseInt(document.getElementById("day").value),' +
    'dom:parseInt(document.getElementById("dom").value),' +
    'hour:parseInt(document.getElementById("hour").value),' +
    'minute:parseInt(document.getElementById("min").value)};' +
    'google.script.run' +
    '.withSuccessHandler(function(msg){' +
    'document.getElementById("status").style.background="#D1FAE5";' +
    'document.getElementById("status").innerText=msg;' +
    'setTimeout(function(){google.script.host.close();},2000);})' +
    '.withFailureHandler(function(e){' +
    'document.getElementById("status").style.background="#FEE2E2";' +
    'document.getElementById("status").innerText="Error: "+e.message;})' +
    '.createCustomTrigger(o);}' +
    '</script>'
  ).setWidth(400).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "Create Custom Trigger");
}

function createCustomTrigger(opts) {
  // Remove existing triggers for same function
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === opts.action) ScriptApp.deleteTrigger(t);
  });

  var dayMap = { 1: ScriptApp.WeekDay.MONDAY, 2: ScriptApp.WeekDay.TUESDAY,
    3: ScriptApp.WeekDay.WEDNESDAY, 4: ScriptApp.WeekDay.THURSDAY,
    5: ScriptApp.WeekDay.FRIDAY, 6: ScriptApp.WeekDay.SATURDAY,
    7: ScriptApp.WeekDay.SUNDAY };

  if (opts.freq === "once") {
    if (!opts.date) throw new Error("Please select a date.");
    var parts = opts.date.split("-");
    var y = parseInt(parts[0]), mo = parseInt(parts[1]) - 1, d = parseInt(parts[2]);
    var triggerDate = new Date(y, mo, d, opts.hour || 0, opts.minute || 0, 0);

    // Add 2 minutes buffer to handle near-future times
    var now = new Date();
    now.setMinutes(now.getMinutes() - 2);
    if (triggerDate <= now) {
      throw new Error("The selected date and time (" + triggerDate.toLocaleString() + ") is in the past. Please choose a future date.");
    }

    ScriptApp.newTrigger(opts.action).timeBased().at(triggerDate).create();
    return "Trigger created: " + opts.action + " at " + Utilities.formatDate(triggerDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  }

  var b = ScriptApp.newTrigger(opts.action).timeBased();
  if (opts.freq === "daily") {
    b.everyDays(1).atHour(opts.hour || 9);
  } else if (opts.freq === "weekly") {
    b.onWeekDay(dayMap[opts.day] || ScriptApp.WeekDay.MONDAY).atHour(opts.hour || 9);
  } else if (opts.freq === "monthly") {
    b.onMonthDay(opts.dom || 1).atHour(opts.hour || 9);
  }
  b.create();

  return "Trigger created: " + opts.action + " (" + opts.freq + " at " + (opts.hour || 9) + ":00)";
}


// ================================================================
//  PRESET TRIGGERS
// ================================================================

function _mkTrigger(fn, freq, hour, dow, dom) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
  });
  var b = ScriptApp.newTrigger(fn).timeBased();
  var dm = { 1:ScriptApp.WeekDay.MONDAY, 2:ScriptApp.WeekDay.TUESDAY, 3:ScriptApp.WeekDay.WEDNESDAY,
    4:ScriptApp.WeekDay.THURSDAY, 5:ScriptApp.WeekDay.FRIDAY, 6:ScriptApp.WeekDay.SATURDAY, 7:ScriptApp.WeekDay.SUNDAY };
  if (freq === "daily") b.everyDays(1).atHour(hour);
  else if (freq === "weekly") b.onWeekDay(dm[dow]).atHour(hour);
  else if (freq === "monthly") b.onMonthDay(dom || 1).atHour(hour);
  b.create();
}

function setupWeeklyHighTrigger()      { _mkTrigger("triggerHighTier","weekly",9,1);              _alert("Done. Weekly HIGH tier: Monday 09:00."); }
function setupBiweeklyMediumTrigger()  { _mkTrigger("triggerMediumTier","weekly",14,3);           _alert("Done. MEDIUM tier: Wednesday 14:00."); }
function setupMonthlyLowTrigger()      { _mkTrigger("triggerAutoSendLowTier","monthly",10,null,1);_alert("Done. Monthly LOW auto-send: 1st at 10:00."); }
function setupDailyAutoSend()          { _mkTrigger("autoSendComplianceEmails","daily",9);        _alert("Done. Daily auto-send: 09:00."); }
function setupEscalationCheck()        { _mkTrigger("checkEscalations","weekly",17,5);            _alert("Done. Escalation check: Friday 17:00."); }
function setupWeeklyDigestTrigger()    { _mkTrigger("sendWeeklyDigest","weekly",18,5);            _alert("Done. Weekly digest: Friday 18:00."); }

function removeAllTriggers() {
  var count = ScriptApp.getProjectTriggers().length;
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  _alert("Removed " + count + " trigger(s).");
}

function viewActiveTriggers() {
  var tt = ScriptApp.getProjectTriggers();
  if (!tt.length) { _alert("No active triggers."); return; }
  var lines = tt.map(function(t, i) {
    return (i + 1) + ". " + t.getHandlerFunction() + " (" + t.getEventType() + ")";
  });
  _alert("Active triggers (" + tt.length + "):\n\n" + lines.join("\n"));
}


// ================================================================
//  ONE-TIME SEND FUNCTIONS
// ================================================================

function sendToSpecificMill() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Send to Specific Mill", "Enter Mill ID (e.g. FH-MH-001):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var targetId = resp.getResponseText().trim();
  if (!targetId) { _alert("No Mill ID entered."); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = loadSettings(ss);
  var ds = ss.getSheetByName(config.DATA_SHEET);
  var data = ds.getDataRange().getValues();
  var headers = data[0];
  var ci = {}; for (var h = 0; h < headers.length; h++) ci[headers[h]] = h;
  var col = function(n) { return ci[n]; };
  var contacts = buildContactMap(ss.getSheetByName(config.CONTACT_SHEET));
  var managers = buildLookupMap(ss.getSheetByName(config.MANAGER_SHEET));

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col("Mill ID")]).trim() !== targetId) continue;

    var row = data[i];
    var millName = row[col("Mill Name")];
    var state = String(row[col("State")] || "").trim();
    var tier = normaliseTier(row[col("Risk Tier")]);
    var contact = contacts[targetId];
    if (!contact) { _alert("No contact found for " + targetId + " in Mill_Contacts."); return; }
    if (!tier) { _alert("Invalid tier for " + targetId + "."); return; }

    var manager = managers[state] || { name: "Fortify Health Field Team", email: null };
    var payload = extractPayload(row, col, tier, manager.name);
    var email = buildEmail(payload, contact.name, config);

    var mode = ui.alert("Send or Draft?",
      "Mill: " + millName + " (" + tier + ")\nTo: " + contact.email +
      "\n\nYES = Send now\nNO = Create draft",
      ui.ButtonSet.YES_NO);

    var opts = { htmlBody: email.htmlBody, name: "Fortify Health Programme" };
    if (config.CC_MANAGER && manager.email) opts.cc = manager.email;
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    if (mode === ui.Button.YES) {
      MailApp.sendEmail(contact.email, email.subject, email.body, opts);
      ds.getRange(i + 1, col("Email Flag") + 1).setValue("SENT - " + ts).setBackground("#D5F5E3");
      _alert("Sent to " + contact.email + ".");
    } else {
      GmailApp.createDraft(contact.email, email.subject, email.body, opts);
      ds.getRange(i + 1, col("Email Flag") + 1).setValue("DRAFTED - " + ts).setBackground("#E2EFDA");
      _alert("Draft created for " + contact.email + ".");
    }
    return;
  }
  _alert("Mill ID \"" + targetId + "\" not found in " + config.DATA_SHEET + ".");
}

function sendToSpecificTier() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Send to Tier", "Enter tier (HIGH, MEDIUM, or LOW):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var tier = resp.getResponseText().trim().toUpperCase();
  if (tier !== "HIGH" && tier !== "MEDIUM" && tier !== "LOW") { _alert("Invalid tier. Enter HIGH, MEDIUM, or LOW."); return; }
  var mode = ui.alert("Mode", "YES = Send immediately\nNO = Create drafts\nCANCEL = Abort", ui.ButtonSet.YES_NO_CANCEL);
  if (mode === ui.Button.CANCEL) return;
  _lockedBatch(mode === ui.Button.YES, tier);
}

function scheduleOneTimeSend() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt("Schedule One-Time Send",
    "Enter date and time (format: YYYY-MM-DD HH:MM)\nExample: 2026-04-20 09:30",
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var m = resp.getResponseText().trim().match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})$/);
  if (!m) { _alert("Invalid format. Use: YYYY-MM-DD HH:MM"); return; }
  var dt = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]), parseInt(m[4]), parseInt(m[5]), 0);
  if (dt <= new Date()) { _alert("Date must be in the future."); return; }
  var mode = ui.alert("Mode", "YES = Auto-send at that time\nNO = Create drafts at that time", ui.ButtonSet.YES_NO_CANCEL);
  if (mode === ui.Button.CANCEL) return;
  var fn = (mode === ui.Button.YES) ? "autoSendComplianceEmails" : "draftComplianceEmails";
  ScriptApp.newTrigger(fn).timeBased().at(dt).create();
  _alert("Scheduled: " + fn + "\nDate: " + Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
}


// ================================================================
//  SEND ALL PENDING DRAFTS
// ================================================================

function sendAllPendingDrafts() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("Send All Pending Drafts",
    "This will send ALL Fortify Health email drafts currently in your Gmail.\n\nProceed?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var drafts = GmailApp.getDrafts();
  var sc = 0, total = 0;
  drafts.forEach(function(d) {
    var subj = d.getMessage().getSubject() || "";
    if (subj.indexOf("Fortify") !== -1 || subj.indexOf("fortification") !== -1 || subj.indexOf("Fortification") !== -1) {
      total++;
      try { d.send(); sc++; Utilities.sleep(500); }
      catch (e) { Logger.log("Failed to send draft: " + e.message); }
    }
  });
  _alert("Result: Sent " + sc + " of " + total + " Fortify Health drafts.");
}


// ================================================================
//  DATA TOOLS
// ================================================================

function recalcAllDrift() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName(CONFIG.DATA_SHEET);
  var data = ds.getDataRange().getValues();
  var h = data[0]; var ci = {}; for (var j = 0; j < h.length; j++) ci[h[j]] = j;
  var dc = ci["Drift Label"];
  if (dc === undefined) { _alert("Column 'Drift Label' not found in " + CONFIG.DATA_SHEET + "."); return; }
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var pn = function(c) { var v = data[i][ci[c]]; return (typeof v === "number") ? v : parseFloat(String(v || "").replace(/%/g, "")); };
    ds.getRange(i + 1, dc + 1).setValue(getDriftLabel(pn("M1 Compliance pct"), pn("M2 Compliance pct"), pn("M3 Compliance pct")));
    n++;
  }
  _alert("Updated drift labels for " + n + " rows.");
}

function resetAllFlags() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("Reset All Email Flags",
    "This resets ALL Email Flag values to REVIEW.\nExisting DRAFTED/SENT/DATA ERROR status will be lost.\n\nProceed?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName(CONFIG.DATA_SHEET);
  var data = ds.getDataRange().getValues();
  var fc = data[0].indexOf("Email Flag");
  if (fc === -1) { _alert("Column 'Email Flag' not found."); return; }
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      ds.getRange(i + 1, fc + 1).setValue("REVIEW").setBackground("#EDE9FE");
      n++;
    }
  }
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.CHECKPOINT_KEY);
  _alert("Reset " + n + " flags to REVIEW. Checkpoint cleared.");
}

function validateAllRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var config = loadSettings(ss);
  var ds = ss.getSheetByName(config.DATA_SHEET); var data = ds.getDataRange().getValues();
  var h = data[0]; var ci = {}; for (var j = 0; j < h.length; j++) ci[h[j]] = j;
  var col = function(n) { return ci[n]; };
  var ok = 0, bad = 0, errs = [];
  for (var i = 1; i < data.length; i++) {
    var mn = data[i][col("Mill Name")]; if (!mn) continue;
    var r = validateRow(data[i], col, mn, config);
    if (r.valid) ok++; else { bad++; errs.push("Row " + (i + 1) + " " + mn + ": " + r.reason); }
  }
  var msg = "Validation Results:\n\nValid: " + ok + "\nInvalid: " + bad;
  if (errs.length) msg += "\n\nDetails:\n- " + errs.join("\n- ");
  _alert(msg);
}

function countByTier() {
  var ds = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DATA_SHEET);
  var data = ds.getDataRange().getValues(); var tc = data[0].indexOf("Risk Tier");
  var c = { HIGH: 0, MEDIUM: 0, LOW: 0, Other: 0 };
  for (var i = 1; i < data.length; i++) {
    var t = normaliseTier(data[i][tc]); if (t) c[t]++; else if (data[i][0]) c.Other++;
  }
  _alert("Mill Count by Tier:\n\nHIGH:   " + c.HIGH + "\nMEDIUM: " + c.MEDIUM + "\nLOW:    " + c.LOW +
    (c.Other > 0 ? "\nOther:  " + c.Other : ""));
}

function findMissingContacts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ds = ss.getSheetByName(CONFIG.DATA_SHEET); var data = ds.getDataRange().getValues();
  var contacts = buildContactMap(ss.getSheetByName(CONFIG.CONTACT_SHEET));
  var ic = data[0].indexOf("Mill ID"); var miss = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][ic] || "").trim();
    if (id && !contacts[id]) miss.push(id + " - " + (data[i][ic + 1] || "unknown"));
  }
  _alert(miss.length
    ? "Missing contacts (" + miss.length + "):\n\n- " + miss.join("\n- ")
    : "All mills have matching contacts.");
}


// ================================================================
//  REPORTS
// ================================================================

function checkEscalations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var config = loadSettings(ss);
  var ls = ss.getSheetByName(config.LOG_SHEET);
  if (!ls) { _alert("No " + config.LOG_SHEET + " sheet found."); return; }
  var ld = ls.getDataRange().getValues(); var hm = {};
  for (var i = 1; i < ld.length; i++) {
    if (String(ld[i][3]).trim() === "HIGH") { var m = ld[i][2]; hm[m] = (hm[m] || 0) + 1; }
  }
  var esc = [];
  for (var n in hm) { if (hm[n] >= 3) esc.push(n + " (" + hm[n] + " entries)"); }
  if (!esc.length) { _alert("No escalations found. All clear."); return; }
  var body = "ESCALATION ALERT\n\n" + esc.length + " mills have been HIGH tier for 3+ consecutive entries:\n\n- " +
    esc.join("\n- ") + "\n\nThese mills may require direct intervention.\n\n-- Fortify Health System";
  try {
    MailApp.sendEmail(config.ADMIN_EMAIL, "Fortify Health - Escalation Alert", body);
    _alert("Escalation alert sent to " + config.ADMIN_EMAIL + ".\n\n" + esc.length + " mills flagged.");
  } catch (e) { _alert("Escalation mills found but email failed:\n\n- " + esc.join("\n- ")); }
}

function sendWeeklyDigest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var config = loadSettings(ss);
  var ds = ss.getSheetByName(config.DATA_SHEET); var data = ds.getDataRange().getValues();
  var tc = data[0].indexOf("Risk Tier"); var fc = data[0].indexOf("Email Flag");
  var tot = 0, hi = 0, me = 0, lo = 0, dr = 0, se = 0, er = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; tot++;
    var t = normaliseTier(data[i][tc]);
    if (t === "HIGH") hi++; else if (t === "MEDIUM") me++; else if (t === "LOW") lo++;
    var f = String(data[i][fc] || "");
    if (f.indexOf("DRAFTED") === 0) dr++; else if (f.indexOf("SENT") === 0) se++; else if (f.indexOf("DATA ERROR") === 0) er++;
  }
  var body = "FORTIFY HEALTH - WEEKLY DIGEST\n" +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy") +
    "\n\nProgramme Overview:\n  Total mills: " + tot +
    "\n  HIGH:   " + hi + " (" + (tot ? (hi / tot * 100).toFixed(1) : "0") + "%)" +
    "\n  MEDIUM: " + me + " (" + (tot ? (me / tot * 100).toFixed(1) : "0") + "%)" +
    "\n  LOW:    " + lo + " (" + (tot ? (lo / tot * 100).toFixed(1) : "0") + "%)" +
    "\n\nEmail Status:\n  Drafted: " + dr + "\n  Sent: " + se + "\n  Errors: " + er +
    "\n\n-- Fortify Health Automated System";
  try {
    MailApp.sendEmail(config.ADMIN_EMAIL, "Fortify Health - Weekly Digest", body);
    _alert("Digest sent to " + config.ADMIN_EMAIL + ".");
  } catch (e) { _alert("Digest email failed: " + e.message); }
}

function generateMillReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var config = loadSettings(ss);
  var ds = ss.getSheetByName(config.DATA_SHEET); var data = ds.getDataRange().getValues();
  var h = data[0]; var ci = {}; for (var j = 0; j < h.length; j++) ci[h[j]] = j;
  var col = function(n) { return ci[n]; };

  var lines = ["MILL COMPLIANCE REPORT - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"), ""];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    lines.push(data[i][col("Mill ID")] + " | " + data[i][col("Mill Name")] + " | " +
      data[i][col("State")] + " | " + data[i][col("Risk Tier")] + " | M1=" +
      data[i][col("M1 Compliance pct")] + "% | " + data[i][col("Lab Result ppm Iron")] +
      "ppm | " + data[i][col("Email Flag")]);
  }
  var rs = ss.getSheetByName("_Report") || ss.insertSheet("_Report");
  rs.clear(); rs.getRange(1, 1).setValue(lines.join("\n")).setFontFamily("Consolas").setFontSize(10);
  rs.setColumnWidth(1, 900); ss.setActiveSheet(rs);
  _alert("Report generated in the _Report sheet.");
}

function exportLogToCsv() {
  var ls = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.LOG_SHEET);
  if (!ls) { _alert("No " + CONFIG.LOG_SHEET + " sheet found."); return; }
  var data = ls.getDataRange().getValues();
  var csv = data.map(function(r) {
    return r.map(function(c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(",");
  }).join("\n");
  var blob = Utilities.newBlob(csv, "text/csv", "fortify_health_log_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd") + ".csv");
  var file = DriveApp.createFile(blob);
  _alert("CSV exported to Google Drive:\n" + file.getUrl());
}


// ================================================================
//  SETTINGS LOADER
// ================================================================

function loadSettings(ss) {
  var result = {};
  for (var k in CONFIG) result[k] = CONFIG[k];

  var st = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  if (!st) return result;

  var data = st.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if (!key || val === undefined || val === "" || key === "Setting") continue;
    if (val === "true" || val === "TRUE") val = true;
    else if (val === "false" || val === "FALSE") val = false;
    else if (typeof val === "string" && !isNaN(val) && val.trim() !== "") val = parseFloat(val);
    result[key] = val;
  }

  // Map follow-up day overrides
  result.FOLLOWUP_DAYS = {
    HIGH:   result.HIGH_FOLLOWUP_DAYS || CONFIG.FOLLOWUP_DAYS.HIGH,
    MEDIUM: result.MEDIUM_FOLLOWUP_DAYS || CONFIG.FOLLOWUP_DAYS.MEDIUM,
    LOW:    result.LOW_FOLLOWUP_DAYS || CONFIG.FOLLOWUP_DAYS.LOW
  };

  return result;
}


// ================================================================
//  HELPERS
// ================================================================

function buildContactMap(sheet) {
  var map = {};
  if (!sheet) return map;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][0] || "").trim();
    if (id) map[id] = { name: rows[i][1], email: rows[i][2] };
  }
  return map;
}

function buildLookupMap(sheet) {
  var map = {};
  if (!sheet) return map;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var key = String(rows[i][0] || "").trim();
    if (key) map[key] = { name: rows[i][1], email: rows[i][2] };
  }
  return map;
}

function getOrCreateLogSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET);
    var h = ["Timestamp", "Mill ID", "Mill Name", "Risk Tier", "Email Type",
             "Subject (preview)", "Field Manager", "Contact Email",
             "Follow-up Due", "Drift Trend", "Status", "Notes"];
    sheet.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight("bold").setBackground("#1F4E79").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(6, 300);
  }
  return sheet;
}

function appendToLog(sheet, e) {
  sheet.appendRow([
    e.timestamp, e.millId, e.millName, e.tier, e.emailType,
    e.subject.substring(0, 80), e.fieldManager, e.contactEmail,
    e.followUpDue, e.drift || "-", e.status, ""
  ]);
}

function getFollowUpDate(tier, config) {
  var days = (config.FOLLOWUP_DAYS && config.FOLLOWUP_DAYS[tier]) || 7;
  var d = new Date();
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function _alert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) { Logger.log(msg); }
}

function _finish(drafted, sent, skipped, errors, partial, lastRow, totalRows) {
  var msg = partial
    ? "Batch paused at row " + lastRow + " of " + totalRows + ". Run again to continue.\n\n"
    : "Run complete.\n\n";
  msg += "Drafted: " + drafted + "  |  Sent: " + sent + "  |  Skipped: " + skipped;
  if (errors.length) {
    msg += "\n\nIssues (" + errors.length + "):\n- " + errors.join("\n- ");
  }
  _alert(msg);
}

function resetCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.CHECKPOINT_KEY);
  _alert("Checkpoint cleared. Next run starts from the beginning.");
}
