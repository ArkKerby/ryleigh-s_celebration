const SETTINGS = {
  spreadsheetId: "1xzzbHzWASeCMjvG9m6PuXlq5XVXgesqlxW6xRDQZmC8",
  responseSheetName: "RSVP Responses",
  directorySheetName: "Godparent Directory",
  summarySheetName: "Summary",
  maxSubmissionsPerDay: 200,
  maxConfirmationEmailsPerDay: 40,
  recipientCooldownSeconds: 3600
};

const GODPARENT_SEED = [
  ["Ma. Theresa F. Leyva", "Ninang"],
  ["Myrell D. Fuentes", "Ninang"],
  ["Nina Reevina B. Zamora", "Ninang"],
  ["Jacqueline P. Ciriaco", "Ninang"],
  ["Ma. Angeline G. Jabagat", "Ninang"],
  ["Leorey A. Merlan", "Ninang"],
  ["Erica Diane Z. Deluna", "Ninang"],
  ["Rosemarie D. Villanueva", "Ninang"],
  ["Michael D. Fuentes", "Ninong"],
  ["Kylea Arjuna F. Leyva", "Ninong"],
  ["Calvin M. Duran", "Ninong"],
  ["Albert Fortuna", "Ninong"],
  ["Juniel Iway", "Ninong"],
  ["Alfer Nuevo", "Ninong"],
  ["John Mark T. Laguador", "Ninong"],
  ["Prince Edward D. Velasco", "Ninong"]
];

const RESPONSE_HEADERS = [
  "Submitted At",
  "Guest Name (Submitted)",
  "Email",
  "Attendance",
  "Guest Type (Selected)",
  "Matched Godparent Name",
  "Final Role",
  "Name Match Status",
  "Blessing / Message",
  "Confirmation Email Sent",
  "Check-in Code",
  "Check-in Status",
  "Checked In At",
  "Check-in Pass Sent"
];

function setupRsvpSystem() {
  const spreadsheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const responses = getOrCreateSheet_(spreadsheet, SETTINGS.responseSheetName);
  const directory = getOrCreateSheet_(spreadsheet, SETTINGS.directorySheetName);
  const summary = getOrCreateSheet_(spreadsheet, SETTINGS.summarySheetName);

  ensureHeaders_(responses, RESPONSE_HEADERS);
  ensureHeaders_(directory, ["Full Name", "Role"]);
  backfillCheckInCodes_(responses);
  if (directory.getLastRow() <= 1) {
    directory.getRange(2, 1, GODPARENT_SEED.length, 2).setValues(GODPARENT_SEED);
  }
  formatTable_(responses, RESPONSE_HEADERS.length);
  formatTable_(directory, 2);
  setupSummary_(summary);
  Logger.log("RSVP tables are ready: " + spreadsheet.getUrl());
}

function doGet() {
  return HtmlService.createHtmlOutput("Ryleigh's RSVP endpoint is ready.");
}

function doPost(event) {
  const params = event && event.parameter ? event.parameter : {};
  const token = clean_(params.submissionToken, 100);

  try {
    if (clean_(params.website, 200)) {
      return frameReply_(token, true, false, "");
    }

    const guest = validateSubmission_(params);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    let classification;
    let responseRow;
    let checkInCode = "";
    try {
      const spreadsheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
      const responseSheet = spreadsheet.getSheetByName(SETTINGS.responseSheetName);
      const directorySheet = spreadsheet.getSheetByName(SETTINGS.directorySheetName);
      if (!responseSheet || !directorySheet) {
        throw new Error("Run setupRsvpSystem() before accepting RSVPs.");
      }
      ensureHeaders_(responseSheet, RESPONSE_HEADERS);

      const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
      const properties = PropertiesService.getScriptProperties();
      const submissionCountKey = "rsvp-submission-count-" + day;
      const submissionCount = Number(properties.getProperty(submissionCountKey) || 0);
      if (submissionCount >= SETTINGS.maxSubmissionsPerDay) {
        throw new Error("The daily RSVP submission limit has been reached.");
      }

      classification = classifyGuest_(guest.name, guest.guestType, directorySheet);
      checkInCode = /^Yes,/i.test(guest.attendance) ? createCheckInCode_(responseSheet) : "";
      responseRow = responseSheet.getLastRow() + 1;
      responseSheet.appendRow([
        new Date(),
        sheetSafe_(guest.name),
        sheetSafe_(guest.email),
        sheetSafe_(guest.attendance),
        guest.guestType,
        sheetSafe_(classification.matchedName),
        classification.finalRole,
        classification.matchStatus,
        sheetSafe_(guest.message),
        "Pending",
        checkInCode,
        checkInCode ? "Not checked in" : "N/A",
        "",
        checkInCode ? "Not sent" : "N/A"
      ]);
      properties.setProperty(submissionCountKey, String(submissionCount + 1));
      try {
        formatTable_(responseSheet, RESPONSE_HEADERS.length);
      } catch (formatError) {
        console.error("RSVP saved, but table formatting could not be refreshed: " + formatError);
      }
    } finally {
      lock.releaseLock();
    }

    let emailSent = false;
    try {
      const emailGuest = Object.assign({}, guest, { checkInCode: checkInCode });
      emailSent = sendConfirmation_(emailGuest, classification, false);
    } catch (mailError) {
      console.error("RSVP saved but confirmation email failed: " + mailError);
    }
    try {
      markEmailStatus_(responseRow, emailSent, /^Yes,/i.test(guest.attendance));
    } catch (statusError) {
      console.error("RSVP saved, but email status could not be updated: " + statusError);
    }
    return frameReply_(token, true, emailSent, "");
  } catch (error) {
    console.error("RSVP submission failed: " + error);
    return frameReply_(token, false, false, "We could not record your RSVP. Please try again or contact the host.");
  }
}

function validateSubmission_(params) {
  const name = clean_(params.name, 120);
  const email = clean_(params.email, 254);
  const attendance = clean_(params.attendance, 80);
  const guestType = clean_(params.guestType, 40);
  const message = clean_(params.message, 2000);
  const allowedAttendance = ["Yes, I’ll be there!", "Sorry, I can’t make it"];
  const allowedGuestTypes = ["Ninong", "Ninang", "Other guest"];

  if (!name || !email || !message) throw new Error("Required fields are missing.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("The email address is invalid.");
  if (allowedAttendance.indexOf(attendance) === -1) throw new Error("Choose an attendance option.");
  if (allowedGuestTypes.indexOf(guestType) === -1) throw new Error("Choose a guest type.");
  return { name: name, email: email, attendance: attendance, guestType: guestType, message: message };
}

function classifyGuest_(submittedName, selectedType, directorySheet) {
  if (selectedType !== "Other guest") {
    return { matchedName: "", finalRole: selectedType, matchStatus: "Self-selected role" };
  }

  const rows = directorySheet.getDataRange().getValues().slice(1);
  const submittedTokens = nameTokens_(submittedName);
  if (submittedTokens.length < 2) {
    return { matchedName: "", finalRole: "Other guest", matchStatus: "No match: enter at least two name parts" };
  }

  const matches = rows
    .filter(function (row) {
      return row[0] && row[1] && containsNamePhrase_(submittedTokens, nameTokens_(row[0]));
    })
    .map(function (row) { return { name: String(row[0]).trim(), role: String(row[1]).trim() }; });
  const uniqueMatches = matches.filter(function (match, index, all) {
    return all.findIndex(function (other) {
      return other.name === match.name && other.role === match.role;
    }) === index;
  });

  if (uniqueMatches.length === 1) {
    return {
      matchedName: uniqueMatches[0].name,
      finalRole: uniqueMatches[0].role,
      matchStatus: "Matched to godparent directory"
    };
  }
  if (uniqueMatches.length > 1) {
    return { matchedName: "", finalRole: "Review", matchStatus: "Ambiguous match: check manually" };
  }
  return { matchedName: "", finalRole: "Other guest", matchStatus: "No roster match" };
}

function nameTokens_(name) {
  return String(name || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
}

function containsNamePhrase_(phrase, fullName) {
  let matched = 0;
  for (let index = 0; index < fullName.length && matched < phrase.length; index++) {
    if (phrase[matched] === fullName[index]) matched++;
  }
  return matched === phrase.length;
}

function sendConfirmation_(guest, classification, bypassRecipientCooldown) {
  const now = new Date();
  const day = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  const properties = PropertiesService.getScriptProperties();
  const sentTodayKey = "rsvp-mail-count-" + day;
  const cache = CacheService.getScriptCache();
  const recipientKey = "rsvp-mail-" + day + "-" + hash_(guest.email.toLowerCase());
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sentToday = Number(properties.getProperty(sentTodayKey) || 0);
    if (sentToday >= SETTINGS.maxConfirmationEmailsPerDay) return false;
    if (!bypassRecipientCooldown && cache.get(recipientKey)) return false;
    if (MailApp.getRemainingDailyQuota() < 1) return false;

    const hasEntrancePass = /^Yes,/i.test(guest.attendance) && Boolean(guest.checkInCode);
    const subject = hasEntrancePass
      ? "Your Royal RSVP and Entrance Pass | Princess Ryleigh"
      : "RSVP Confirmation | Princess Ryleigh";
    const emailRole = classification.finalRole === "Review" ? "Other guest (name needs review)" : classification.finalRole;
    const roleLine = emailRole + (classification.matchedName ? " (" + classification.matchedName + ")" : "");
    const bodyLines = [
      "PRINCESS RYLEIGH",
      "RSVP CONFIRMATION",
      "For " + guest.name,
      "",
      "Your response has been recorded.",
      "",
      "DATE",
      "Saturday, November 21, 2026",
      "",
      "CHRISTENING | 10:30 in the morning",
      "Parish of the Immaculate Heart of Mary",
      "Daang Bakal Road, Antipolo, Philippines, 1870",
      "",
      "BIRTHDAY CELEBRATION | 4:00 in the afternoon",
      "CPoint Villa",
      "678, 1800 Bagong Silang, Marikina, 1800",
      "",
      "Attendance: " + guest.attendance,
      "Guest type: " + roleLine,
      ""
    ];
    if (hasEntrancePass) {
      bodyLines.push(
        "ENTRANCE CODE: " + guest.checkInCode,
        "Show this code to the event team at the entrance. It admits the named guest only.",
        ""
      );
    }
    bodyLines.push(
      "Strictly no Plus 1. This invitation is for the named guest only.",
      "",
      "Your blessing:",
      guest.message,
      "",
      "With love,",
      "The Merlan Family"
    );

    MailApp.sendEmail({
      to: guest.email,
      subject: subject,
      body: bodyLines.join("\n"),
      htmlBody: buildRsvpEmailHtml_(guest, classification, emailRole),
      name: "Princess Ryleigh's Invitation"
    });
    properties.setProperty(sentTodayKey, String(sentToday + 1));
    cache.put(recipientKey, "1", SETTINGS.recipientCooldownSeconds);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function buildRsvpEmailHtml_(guest, classification, emailRole) {
  const attending = /^Yes,/i.test(guest.attendance);
  const template = HtmlService.createTemplateFromFile("RsvpEmail");
  template.guestName = guest.name || "";
  template.intro = attending
    ? "Your RSVP is confirmed. We look forward to celebrating this special day with you."
    : "Thank you for letting us know. Your response has been recorded.";
  template.hasEntrancePass = attending && Boolean(guest.checkInCode);
  template.attending = attending;
  template.checkInCode = guest.checkInCode || "";
  template.attendance = guest.attendance || "";
  template.emailRole = emailRole || "";
  template.matchedName = classification.matchedName || "";
  template.message = guest.message || "";
  return template.evaluate().getContent();
}
function markEmailStatus_(row, emailSent, hasCheckInCode) {
  const sheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId).getSheetByName(SETTINGS.responseSheetName);
  if (row <= 1) return;
  sheet.getRange(row, headerColumn_("Confirmation Email Sent")).setValue(emailSent ? "Yes" : "Not sent");
  if (hasCheckInCode) {
    sheet.getRange(row, headerColumn_("Check-in Pass Sent")).setValue(emailSent ? "Yes" : "Not sent");
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Ryleigh Event")
    .addItem("Check in by invitation code", "checkInGuestByCode")
    .addItem("Email entrance passes to existing RSVPs", "emailExistingCheckInPasses")
    .addToUi();
}

function checkInGuestByCode() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    "Check in a guest",
    "Enter the entrance code shown in the guest's email:",
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const code = normalizeCheckInCode_(prompt.getResponseText());
  if (!code) {
    ui.alert("No code entered", "Enter the code from the guest's email and try again.", ui.ButtonSet.OK);
    return;
  }

  const sheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId).getSheetByName(SETTINGS.responseSheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No RSVPs found", "The RSVP response sheet has no guest records yet.", ui.ButtonSet.OK);
    return;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let message = "No RSVP matched that code. Check the code and try again.";
  try {
    const codeColumn = headerColumn_("Check-in Code");
    const statusColumn = headerColumn_("Check-in Status");
    const checkedInAtColumn = headerColumn_("Checked In At");
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RESPONSE_HEADERS.length).getValues();
    const codeIndex = codeColumn - 1;
    const statusIndex = statusColumn - 1;
    const attendanceIndex = headerColumn_("Attendance") - 1;
    const nameIndex = headerColumn_("Guest Name (Submitted)") - 1;
    const matchingIndex = rows.findIndex(function (row) {
      return normalizeCheckInCode_(row[codeIndex]) === code;
    });

    if (matchingIndex !== -1) {
      const rowNumber = matchingIndex + 2;
      const record = rows[matchingIndex];
      const guestName = String(record[nameIndex] || "Guest");
      if (!/^Yes,/i.test(String(record[attendanceIndex] || ""))) {
        message = guestName + " is not marked as attending. Do not check in this RSVP.";
      } else if (String(record[statusIndex]) === "Checked in") {
        message = guestName + " has already been checked in.";
      } else if (String(record[statusIndex]) !== "Not checked in") {
        message = guestName + " has an unexpected check-in status. Please ask the event lead to review the RSVP sheet.";
      } else {
        sheet.getRange(rowNumber, statusColumn).setValue("Checked in");
        sheet.getRange(rowNumber, checkedInAtColumn).setValue(new Date());
        message = "Check-in confirmed for " + guestName + ".";
      }
    }
  } finally {
    lock.releaseLock();
  }

  ui.alert("Entrance check", message, ui.ButtonSet.OK);
}

function emailExistingCheckInPasses() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    "Email entrance passes",
    "This sends the updated invitation and entrance code to attending guests who have not received a pass yet. Continue?",
    ui.ButtonSet.YES_NO
  );
  if (confirmation !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId).getSheetByName(SETTINGS.responseSheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No RSVPs found", "The RSVP response sheet has no guest records yet.", ui.ButtonSet.OK);
    return;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, RESPONSE_HEADERS.length).getValues();
  const attendanceIndex = headerColumn_("Attendance") - 1;
  const passSentIndex = headerColumn_("Check-in Pass Sent") - 1;
  const codeIndex = headerColumn_("Check-in Code") - 1;
  let sent = 0;
  let stoppedAtLimit = false;

  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    if (!/^Yes,/i.test(String(row[attendanceIndex] || "")) || !row[codeIndex] || row[passSentIndex] === "Yes") continue;

    const guest = {
      name: String(row[headerColumn_("Guest Name (Submitted)") - 1] || "Guest"),
      email: String(row[headerColumn_("Email") - 1] || ""),
      attendance: String(row[attendanceIndex]),
      guestType: String(row[headerColumn_("Guest Type (Selected)") - 1] || "Other guest"),
      message: String(row[headerColumn_("Blessing / Message") - 1] || ""),
      checkInCode: String(row[codeIndex])
    };
    const classification = {
      matchedName: String(row[headerColumn_("Matched Godparent Name") - 1] || ""),
      finalRole: String(row[headerColumn_("Final Role") - 1] || "Other guest"),
      matchStatus: String(row[headerColumn_("Name Match Status") - 1] || "")
    };
    const emailSent = sendConfirmation_(guest, classification, true);
    if (!emailSent) {
      stoppedAtLimit = true;
      break;
    }
    markEmailStatus_(index + 2, true, true);
    sent++;
  }

  const ending = stoppedAtLimit ? " Some passes remain; you can run this menu item again later." : "";
  ui.alert("Entrance passes", "Sent " + sent + " updated pass email(s)." + ending, ui.ButtonSet.OK);
}

function createCheckInCode_(sheet) {
  const codeColumn = headerColumn_("Check-in Code");
  const lastRow = sheet.getLastRow();
  const existing = new Set(lastRow > 1
    ? sheet.getRange(2, codeColumn, lastRow - 1, 1).getValues().map(function (row) {
      return normalizeCheckInCode_(row[0]);
    }).filter(Boolean)
    : []);

  for (let attempt = 0; attempt < 10; attempt++) {
    const random = Utilities.getUuid().replace(/-/g, "").substring(0, 8).toUpperCase();
    const code = "RYL-" + random.substring(0, 4) + "-" + random.substring(4);
    if (!existing.has(normalizeCheckInCode_(code))) return code;
  }
  throw new Error("Could not generate a unique entrance code. Please try again.");
}

function backfillCheckInCodes_(sheet) {
  if (sheet.getLastRow() < 2) return;
  const codeColumn = headerColumn_("Check-in Code");
  const statusColumn = headerColumn_("Check-in Status");
  const passSentColumn = headerColumn_("Check-in Pass Sent");
  const attendanceColumn = headerColumn_("Attendance");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RESPONSE_HEADERS.length).getValues();

  rows.forEach(function (row, index) {
    const rowNumber = index + 2;
    const attending = /^Yes,/i.test(String(row[attendanceColumn - 1] || ""));
    let code = String(row[codeColumn - 1] || "");
    if (attending && !code) {
      code = createCheckInCode_(sheet);
      sheet.getRange(rowNumber, codeColumn).setValue(code);
    }
    if (attending) {
      if (!row[statusColumn - 1]) sheet.getRange(rowNumber, statusColumn).setValue("Not checked in");
      if (!row[passSentColumn - 1]) sheet.getRange(rowNumber, passSentColumn).setValue("Not sent");
    } else {
      if (!row[statusColumn - 1]) sheet.getRange(rowNumber, statusColumn).setValue("N/A");
      if (!row[passSentColumn - 1]) sheet.getRange(rowNumber, passSentColumn).setValue("N/A");
    }
  });
}

function normalizeCheckInCode_(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function headerColumn_(name) {
  const index = RESPONSE_HEADERS.indexOf(name);
  if (index === -1) throw new Error("Missing response column: " + name);
  return index + 1;
}
function frameReply_(token, ok, emailSent, message) {
  const safeToken = /^[a-zA-Z0-9_-]{8,100}$/.test(token) ? token : "";
  const payload = JSON.stringify({ type: "ryleigh-rsvp-result", token: safeToken, ok: ok, emailSent: emailSent, message: message })
    .replace(/</g, "\\u003c");
  const html = "<!doctype html><html><body><script>(function(){window.top.postMessage(" + payload + ", '*');})();<\/script></body></html>";
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  const isEmpty = sheet.getLastRow() === 0 || sheet.getRange(1, 1).isBlank();
  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const width = Math.max(sheet.getLastColumn(), headers.length);
    const currentHeaders = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (cell) { return cell.trim(); });
    while (currentHeaders.length && !currentHeaders[currentHeaders.length - 1]) currentHeaders.pop();
    const matchesExistingPrefix = currentHeaders.every(function (header, index) { return header === headers[index]; });
    if (!matchesExistingPrefix || currentHeaders.length > headers.length) {
      throw new Error('The tab "' + sheet.getName() + '" already has different headers. Rename it or review it before setup.');
    }
    if (currentHeaders.length < headers.length) {
      sheet.getRange(1, currentHeaders.length + 1, 1, headers.length - currentHeaders.length)
        .setValues([headers.slice(currentHeaders.length)]);
    }
  }
  sheet.setFrozenRows(1);
}

function formatTable_(sheet, columnCount) {
  const headerRange = sheet.getRange(1, 1, 1, columnCount);
  headerRange
    .setBackground("#a64f67")
    .setFontColor("#fffaf2")
    .setFontWeight("bold")
    .setWrap(true);
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), columnCount).setVerticalAlignment("top");
  sheet.autoResizeColumns(1, columnCount);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), columnCount).createFilter();
}

function setupSummary_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange("A1:B1").merge().setValue("Princess Ryleigh RSVP Summary");
  sheet.getRange("A1:B1").setBackground("#a64f67").setFontColor("#fffaf2").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A3:B5").setValues([
    ["Attendance", "Count"],
    ["Yes, I’ll be there!", ""],
    ["Sorry, I can’t make it", ""]
  ]);
  sheet.getRange("B4").setFormula("=COUNTIF('RSVP Responses'!D:D,A4)");
  sheet.getRange("B5").setFormula("=COUNTIF('RSVP Responses'!D:D,A5)");
  sheet.getRange("A7:B11").setValues([
    ["Final Role", "Count"],
    ["Ninong", ""],
    ["Ninang", ""],
    ["Other guest", ""],
    ["Review", ""]
  ]);
  for (let row = 8; row <= 11; row++) {
    sheet.getRange(row, 2).setFormula("=COUNTIF('RSVP Responses'!G:G,A" + row + ")");
  }
  sheet.getRange("A3:B3").setBackground("#ead8d1").setFontWeight("bold");
  sheet.getRange("A7:B7").setBackground("#ead8d1").setFontWeight("bold");
  sheet.autoResizeColumns(1, 2);
  sheet.setFrozenRows(1);
}

function sheetSafe_(value) {
  const text = String(value || "");
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function clean_(value, maxLength) {
  return String(value || "").replace(/\r\n?/g, "\n").trim().slice(0, maxLength);
}

function hash_(value) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value)).replace(/=+$/g, "");
}

function escapeHtml_(value) {
  const replacements = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
  return String(value == null ? "" : value).replace(/[&<>"']/g, function (character) {
    return replacements[character];
  });
}
