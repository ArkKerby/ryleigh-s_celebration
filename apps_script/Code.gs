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
  "Confirmation Email Sent"
];

function setupRsvpSystem() {
  const spreadsheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const responses = getOrCreateSheet_(spreadsheet, SETTINGS.responseSheetName);
  const directory = getOrCreateSheet_(spreadsheet, SETTINGS.directorySheetName);
  const summary = getOrCreateSheet_(spreadsheet, SETTINGS.summarySheetName);

  ensureHeaders_(responses, RESPONSE_HEADERS);
  ensureHeaders_(directory, ["Full Name", "Role"]);
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
    try {
      const spreadsheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
      const responseSheet = spreadsheet.getSheetByName(SETTINGS.responseSheetName);
      const directorySheet = spreadsheet.getSheetByName(SETTINGS.directorySheetName);
      if (!responseSheet || !directorySheet) {
        throw new Error("Run setupRsvpSystem() before accepting RSVPs.");
      }

      const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
      const properties = PropertiesService.getScriptProperties();
      const submissionCountKey = "rsvp-submission-count-" + day;
      const submissionCount = Number(properties.getProperty(submissionCountKey) || 0);
      if (submissionCount >= SETTINGS.maxSubmissionsPerDay) {
        throw new Error("The daily RSVP submission limit has been reached.");
      }

      classification = classifyGuest_(guest.name, guest.guestType, directorySheet);
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
        "Pending"
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
      emailSent = sendConfirmation_(guest, classification);
    } catch (mailError) {
      console.error("RSVP saved but confirmation email failed: " + mailError);
    }
    try {
      markEmailStatus_(responseRow, emailSent);
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

function sendConfirmation_(guest, classification) {
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
    if (cache.get(recipientKey)) return false;
    if (MailApp.getRemainingDailyQuota() < 1) return false;

    const subject = "Your RSVP is received | Princess Ryleigh";
    const emailRole = classification.finalRole === "Review" ? "Other guest (name needs review)" : classification.finalRole;
    const matchedNameLine = classification.matchedName ? "Matched godparent: " + classification.matchedName : "";
    const roleLine = emailRole + (classification.matchedName ? " (" + classification.matchedName + ")" : "");
    const bodyLines = [
      "Hi " + guest.name + ",",
      "",
      "We received your RSVP for Princess Ryleigh’s Christening and First Birthday Celebration.",
      "",
      "Attendance: " + guest.attendance,
      "Guest type: " + roleLine,
      matchedNameLine || null,
      "",
      "Your message:",
      guest.message,
      "",
      "Thank you for being part of Ryleigh’s special day!",
      "With love,",
      "The Merlan Family"
    ];
    const body = bodyLines.filter(function (line) { return line !== null; }).join("\n");
    const htmlBody = [
      '<div style="margin:0;padding:30px 12px;background-color:#f8efec;font-family:Georgia,\'Times New Roman\',serif;color:#593743;">',
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:620px;margin:0 auto;border:1px solid #ead8d1;border-radius:18px;overflow:hidden;background-color:#fffaf4;">',
      '<tr><td style="padding:28px 24px;text-align:center;background-color:#a64f67;color:#fffaf4;">',
      '<p style="margin:0 0 10px;font-family:Arial,sans-serif;font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#f8dfc4;">&#10022; A royal thank you &#10022;</p>',
      '<h1 style="margin:0;font-size:30px;line-height:1.2;font-weight:normal;">Your RSVP is received</h1>',
      '<p style="margin:10px 0 0;font-size:16px;line-height:1.5;color:#fff0e8;">Princess Ryleigh\'s Christening &amp; First Birthday</p>',
      '</td></tr>',
      '<tr><td style="padding:30px 28px 24px;">',
      '<p style="margin:0 0 16px;font-size:19px;">Dear ' + escapeHtml_(guest.name) + ',</p>',
      '<p style="margin:0 0 22px;font-size:15px;line-height:1.7;color:#725f58;">Thank you for letting us know. We have recorded your RSVP and are so happy to celebrate this special day with you.</p>',
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;background-color:#f8efec;border:1px solid #efdfd9;border-radius:12px;">',
      '<tr><td style="padding:14px 16px;border-bottom:1px solid #ead8d1;font-family:Arial,sans-serif;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;color:#987b70;">Attendance</td><td align="right" style="padding:14px 16px;border-bottom:1px solid #ead8d1;font-size:15px;color:#593743;">' + escapeHtml_(guest.attendance) + '</td></tr>',
      '<tr><td style="padding:14px 16px;font-family:Arial,sans-serif;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;color:#987b70;">Guest type</td><td align="right" style="padding:14px 16px;font-size:15px;color:#593743;">' + escapeHtml_(emailRole) + '</td></tr>',
      '</table>',
      (classification.matchedName ? '<p style="margin:12px 2px 0;font-size:13px;color:#806b61;">Godparent match: <strong style="color:#593743;">' + escapeHtml_(classification.matchedName) + '</strong></p>' : ''),
      '<p style="margin:24px 0 8px;font-family:Arial,sans-serif;font-size:11px;letter-spacing:1.5px;text-transform:uppercase;color:#987b70;">Your blessing</p>',
      '<div style="padding:15px 17px;border-left:3px solid #c9a56d;border-radius:0 10px 10px 0;background-color:#fff4ee;font-size:15px;line-height:1.7;color:#654753;white-space:pre-wrap;word-break:break-word;">' + escapeHtml_(guest.message) + '</div>',
      '<p style="margin:25px 0 0;font-size:15px;line-height:1.7;color:#725f58;">Thank you for being part of Ryleigh\'s fairytale. We look forward to celebrating with you!</p>',
      '<p style="margin:18px 0 0;font-size:15px;line-height:1.7;color:#725f58;">With love,<br><strong style="color:#a64f67;">The Merlan Family</strong></p>',
      '</td></tr>',
      '<tr><td style="padding:15px 22px;text-align:center;border-top:1px solid #ead8d1;background-color:#fff5ef;font-family:Arial,sans-serif;font-size:11px;line-height:1.6;color:#9a8278;">A little celebration filled with love, blessings, and storybook magic.</td></tr>',
      '</table></div>'
    ].join("");

    MailApp.sendEmail({
      to: guest.email,
      subject: subject,
      body: body,
      htmlBody: htmlBody,
      name: "Princess Ryleigh’s Invitation"
    });
    properties.setProperty(sentTodayKey, String(sentToday + 1));
    cache.put(recipientKey, "1", SETTINGS.recipientCooldownSeconds);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function markEmailStatus_(row, emailSent) {
  const sheet = SpreadsheetApp.openById(SETTINGS.spreadsheetId).getSheetByName(SETTINGS.responseSheetName);
  if (row > 1) sheet.getRange(row, RESPONSE_HEADERS.length).setValue(emailSent ? "Yes" : "Not sent");
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
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map(function (cell) { return cell.trim(); });
    if (currentHeaders.join("|") !== headers.join("|")) {
      throw new Error('The tab "' + sheet.getName() + '" already has different headers. Rename it or review it before setup.');
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
