// ============================================================
// Meeting Scheduler — Google Apps Script Backend
// ============================================================
// SETUP:
//   1. Open your Meeting Scheduler Google Sheet
//   2. Extensions > Apps Script — paste this code
//   3. Set SENDER_ALIAS and HOST_EMAIL below
//   4. Deploy > New deployment > Web App
//      Execute as: Me | Who has access: Anyone
// ============================================================

const SENDER_ALIAS        = "email@domain.com";   // ← change this
const HOST_EMAIL          = "email@domain.com";   // ← change this
const NETLIFY_URL         = "https://azizsaam.github.io/fluffy-octo-invention/";
const APP_NAME            = "Meeting Scheduler";
const SHEET_POLLS         = "Polls";
const SHEET_RESPONSES     = "Responses";
const SHEET_INVITEES      = "Invitees";
const SHEET_SUMMARY       = "Summary";

// ------------------------------------------------------------
// HTTP Router
// ------------------------------------------------------------
function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === "getPoll")             return json(getPoll(e.parameter.pollId));
    if (action === "getResults")          return json(getResults(e.parameter.pollId));
    if (action === "getInvitee")          return json(getInvitee(e.parameter.token));
    if (action === "createPoll")          return json(createPoll(JSON.parse(e.parameter.data)));
    if (action === "addInvitees")         return json(addInvitees(JSON.parse(e.parameter.data)));
    if (action === "submitResponse")      return json(submitResponse(JSON.parse(e.parameter.data)));
    return json({ error: "Unknown action" });
  } catch(err) {
    return json({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === "createPoll")     return json(createPoll(body));
    if (body.action === "addInvitees")    return json(addInvitees(body));
    if (body.action === "submitResponse") return json(submitResponse(body));
    return json({ error: "Unknown action" });
  } catch(err) {
    return json({ error: err.message });
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp(data, callback) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ------------------------------------------------------------
// Sheet helpers
// ------------------------------------------------------------
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}

function uid() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 12);
}

// ------------------------------------------------------------
// Create poll
// ------------------------------------------------------------
function createPoll(body) {
  const pollId = uid();

  const pollSheet = getOrCreateSheet(SHEET_POLLS,
    ["pollId","title","description","dates","times","hostEmail","createdAt"]);
  pollSheet.appendRow([
    pollId, body.title, body.description || "",
    JSON.stringify(body.dates), JSON.stringify(body.times),
    body.hostEmail || HOST_EMAIL, new Date().toISOString()
  ]);

  const invSheet = getOrCreateSheet(SHEET_INVITEES,
    ["token","pollId","name","email","sentAt"]);

  const inviteeRows = [];
  (body.invitees || []).forEach(inv => {
    const token = uid();
    invSheet.appendRow([token, pollId, inv.name || "", inv.email, new Date().toISOString()]);
    inviteeRows.push({ token, email: inv.email, name: inv.name || "" });
  });

  inviteeRows.forEach(inv => {
    if (!inv.email) return;
    const respondUrl = `${NETLIFY_URL}?token=${inv.token}`;
    sendInviteEmail(inv, body.title, body.description, body.dates, body.times, respondUrl);
  });

  return { success: true, pollId, inviteeCount: inviteeRows.length };
}

// ------------------------------------------------------------
// Add invitees to existing poll
// ------------------------------------------------------------
function addInvitees(body) {
  const poll = getPoll(body.pollId);
  if (poll.error) return poll;

  const invSheet = getOrCreateSheet(SHEET_INVITEES,
    ["token","pollId","name","email","sentAt"]);

  const newRows = [];
  (body.invitees || []).forEach(inv => {
    if (!inv.email) return;
    const token = uid();
    invSheet.appendRow([token, body.pollId, inv.name || "", inv.email, new Date().toISOString()]);
    newRows.push({ token, email: inv.email, name: inv.name || "" });
  });

  newRows.forEach(inv => {
    const respondUrl = `${NETLIFY_URL}?token=${inv.token}`;
    sendInviteEmail(inv, poll.title, poll.description, poll.dates, poll.times, respondUrl);
  });

  return { success: true, inviteeCount: newRows.length };
}

// ------------------------------------------------------------
// Get poll
// ------------------------------------------------------------
function getPoll(pollId) {
  const sheet = getOrCreateSheet(SHEET_POLLS,
    ["pollId","title","description","dates","times","hostEmail","createdAt"]);
  const rows = sheetToObjects(sheet);
  const poll = rows.find(r => r.pollId === pollId);
  if (!poll) return { error: "Poll not found" };
  return {
    pollId: poll.pollId,
    title: poll.title,
    description: poll.description,
    dates: JSON.parse(poll.dates),
    times: JSON.parse(poll.times),
  };
}

// ------------------------------------------------------------
// Get invitee by token
// ------------------------------------------------------------
function getInvitee(token) {
  const sheet = getOrCreateSheet(SHEET_INVITEES,
    ["token","pollId","name","email","sentAt"]);
  const rows = sheetToObjects(sheet);
  const inv = rows.find(r => r.token === token);
  if (!inv) return { error: "Token not found" };
  return { token: inv.token, pollId: inv.pollId, name: inv.name, email: inv.email };
}

// ------------------------------------------------------------
// Submit response
// ------------------------------------------------------------
function submitResponse(body) {
  const respSheet = getOrCreateSheet(SHEET_RESPONSES,
    ["responseId","pollId","token","name","email","selections","submittedAt"]);

  const existing = sheetToObjects(respSheet).find(r => r.token === body.token);
  if (existing) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESPONSES);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === body.token) {
        sheet.getRange(i + 1, 4, 1, 4).setValues([[
          body.name, body.email || existing.email,
          JSON.stringify(body.selections), new Date().toISOString()
        ]]);
        break;
      }
    }
  } else {
    respSheet.appendRow([
      uid(), body.pollId, body.token,
      body.name, body.email || "",
      JSON.stringify(body.selections),
      new Date().toISOString()
    ]);
  }

  const poll = getPoll(body.pollId);
  try { notifyHost(body, poll); } catch(e) { Logger.log('notifyHost error: ' + e.message); }
  try { updateSummarySheet(body.pollId, poll); } catch(e) { Logger.log('updateSummarySheet error: ' + e.message); }

  return { success: true };
}

// ------------------------------------------------------------
// Get results
// ------------------------------------------------------------
function getResults(pollId) {
  const poll = getPoll(pollId);
  if (poll.error) return poll;

  const respSheet = getOrCreateSheet(SHEET_RESPONSES,
    ["responseId","pollId","token","name","email","selections","submittedAt"]);
  const responses = sheetToObjects(respSheet)
    .filter(r => r.pollId === pollId)
    .map(r => ({
      name: r.name, email: r.email, token: r.token,
      selections: JSON.parse(r.selections || "{}"),
      submittedAt: r.submittedAt,
    }));

  const invSheet = getOrCreateSheet(SHEET_INVITEES,
    ["token","pollId","name","email","sentAt"]);
  const invitees = sheetToObjects(invSheet)
    .filter(r => r.pollId === pollId)
    .map(r => ({ name: r.name, email: r.email, token: r.token }));

  return { poll, responses, invitees };
}

// ------------------------------------------------------------
// Update Summary sheet — human-readable grid
// ------------------------------------------------------------
function updateSummarySheet(pollId, poll) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SUMMARY);
  } else {
    // Unmerge all cells first, then clear, to avoid merge conflicts
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clearContents();
    sheet.clearFormats();
  }

  const respSheet = ss.getSheetByName(SHEET_RESPONSES);
  if (!respSheet) return;

  const responses = sheetToObjects(respSheet)
    .filter(r => r.pollId === pollId)
    .map(r => ({ name: r.name, selections: JSON.parse(r.selections || "{}") }));

  const total = responses.length;
  if (total === 0) return;

  // Build header row: blank, then each date+time combo
  const slots = [];
  poll.dates.forEach(d => poll.times.forEach(t => slots.push({ date: d, time: t, key: d + '__' + t })));

  // Score each slot
  const scores = {};
  slots.forEach(s => scores[s.key] = 0);
  responses.forEach(r => {
    Object.keys(r.selections).forEach(k => { if (r.selections[k]) scores[k] = (scores[k] || 0) + 1; });
  });
  const maxScore = Math.max(...Object.values(scores));

  // Title
  sheet.getRange(1, 1).setValue(poll.title + " — Availability Summary");
  sheet.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  sheet.getRange(2, 1).setValue("Updated: " + new Date().toLocaleString());
  sheet.getRange(2, 1).setFontColor("#888888");

  // Group by date — write date headers spanning time slot columns
  let col = 2;
  const dateGroups = {};
  poll.dates.forEach(d => {
    dateGroups[d] = { startCol: col, count: poll.times.length };
    col += poll.times.length;
  });

  // Row 4: date headers
  poll.dates.forEach(d => {
    const g = dateGroups[d];
    const dt = new Date(d + "T12:00:00");
    const label = dt.toLocaleDateString("en-US", { weekday:"short", month:"short", day:"numeric" });
    sheet.getRange(4, g.startCol, 1, g.count).merge()
      .setValue(label)
      .setBackground("#1D9E75").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");
  });

  // Row 5: time headers
  sheet.getRange(5, 1).setValue("Person").setFontWeight("bold").setBackground("#f5f5f2");
  slots.forEach((s, i) => {
    sheet.getRange(5, i + 2).setValue(s.time)
      .setBackground("#f5f5f2").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true);
  });

  // Rows 6+: one row per person
  responses.forEach((r, ri) => {
    sheet.getRange(6 + ri, 1).setValue(r.name);
    slots.forEach((s, si) => {
      const avail = !!r.selections[s.key];
      const cell = sheet.getRange(6 + ri, si + 2);
      cell.setValue(avail ? "✓" : "")
        .setHorizontalAlignment("center")
        .setBackground(avail ? "#e6f9f3" : "#ffffff")
        .setFontColor(avail ? "#0a7a54" : "#cccccc");
    });
  });

  // Summary count row
  const countRow = 6 + responses.length + 1;
  sheet.getRange(countRow, 1).setValue("Available (" + total + " total)").setFontWeight("bold");
  slots.forEach((s, i) => {
    const count = scores[s.key] || 0;
    const isBest = count === maxScore && count > 0;
    const cell = sheet.getRange(countRow, i + 2);
    cell.setValue(count + "/" + total)
      .setFontWeight("bold").setHorizontalAlignment("center")
      .setBackground(isBest ? "#1D9E75" : count > 0 ? "#e6f9f3" : "#f5f5f5")
      .setFontColor(isBest ? "#ffffff" : count > 0 ? "#0a7a54" : "#999999");
  });

  // Best slot label
  const bestSlots = slots.filter(s => scores[s.key] === maxScore && maxScore > 0);
  if (bestSlots.length > 0) {
    const bestRow = countRow + 2;
    const bestLabels = bestSlots.map(s => {
      const dt = new Date(s.date + "T12:00:00");
      return dt.toLocaleDateString("en-US", { weekday:"long", month:"long", day:"numeric" }) + ", " + s.time;
    });
    sheet.getRange(bestRow, 1, 1, slots.length + 1).merge()
      .setValue("★ Best slot" + (bestSlots.length > 1 ? "s" : "") + ": " + bestLabels.join(" | "))
      .setBackground("#fff3cd").setFontColor("#856404").setFontWeight("bold");
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, slots.length + 1);
}

// ------------------------------------------------------------
// Email: invite
// ------------------------------------------------------------
function sendInviteEmail(inv, title, description, dates, times, respondUrl) {
  const dateList = dates.map(d => {
    const dt = new Date(d + "T12:00:00");
    return "• " + dt.toLocaleDateString("en-US", { weekday:"long", month:"long", day:"numeric" });
  }).join("<br>");

  const html = `
    <div style="font-family:sans-serif;max-width:520px;margin:0 auto;padding:20px">
      <h2 style="font-size:20px;font-weight:700;margin-bottom:6px">${title}</h2>
      ${description ? `<p style="color:#555;margin-bottom:16px">${description}</p>` : ""}
      <p style="color:#333;margin-bottom:12px">Please indicate your availability for the following dates:</p>
      <div style="background:#f5f5f2;border-radius:10px;padding:14px 18px;margin:0 0 20px;font-size:14px;line-height:2;color:#333">
        ${dateList}
      </div>
      <a href="${respondUrl}" style="display:block;background:#111;color:#fff;text-decoration:none;padding:14px 22px;border-radius:10px;font-size:15px;font-weight:600;text-align:center;margin-bottom:20px">
        Submit my availability &rarr;
      </a>
      <p style="font-size:12px;color:#aaa">This link is personal to you. You can return to it to update your response.</p>
    </div>`;

  GmailApp.sendEmail(inv.email, `[${APP_NAME}] Your availability needed: ${title}`, "", {
    from: SENDER_ALIAS,
    htmlBody: html,
    name: APP_NAME,
  });
}

// ------------------------------------------------------------
// Email: host notification
// ------------------------------------------------------------
function notifyHost(response, poll) {
  const yesSlots = Object.entries(response.selections)
    .filter(([,v]) => v)
    .map(([k]) => {
      const [date, time] = k.split("__");
      const dt = new Date(date + "T12:00:00");
      return "• " + dt.toLocaleDateString("en-US", { weekday:"short", month:"short", day:"numeric" }) + " — " + time;
    })
    .join("<br>") || "(none selected)";

  const html = `
    <div style="font-family:sans-serif;max-width:520px;padding:20px">
      <h2 style="font-size:18px;font-weight:700;margin-bottom:4px">${response.name} responded</h2>
      <p style="color:#555;margin-bottom:16px">Poll: <strong>${poll.title}</strong></p>
      <div style="background:#f0faf5;border-left:4px solid #1D9E75;padding:12px 16px;font-size:14px;line-height:1.9;border-radius:0 8px 8px 0">
        ${yesSlots}
      </div>
      <p style="font-size:12px;color:#aaa;margin-top:16px">Check the Summary tab in your Google Sheet for the full picture.</p>
    </div>`;

  GmailApp.sendEmail(HOST_EMAIL, `[${APP_NAME}] ${response.name} responded to "${poll.title}"`, "", {
    from: SENDER_ALIAS,
    htmlBody: html,
    name: APP_NAME,
  });
}

// ------------------------------------------------------------
// Manual trigger — run this from Apps Script to rebuild Summary
// for the most recent poll
// ------------------------------------------------------------
function rebuildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pollSheet = ss.getSheetByName(SHEET_POLLS);
  if (!pollSheet) { Logger.log('No Polls sheet found'); return; }
  const rows = sheetToObjects(pollSheet);
  if (!rows.length) { Logger.log('No polls found'); return; }
  const latest = rows[rows.length - 1];
  const poll = {
    pollId: latest.pollId,
    title: latest.title,
    description: latest.description,
    dates: JSON.parse(latest.dates),
    times: JSON.parse(latest.times)
  };
  Logger.log('Rebuilding summary for: ' + poll.title);
  updateSummarySheet(poll.pollId, poll);
  Logger.log('Done');
}
