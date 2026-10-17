/**
 * PROSPER MUSIC ACADEMY / PROSPER WORK
 * Google Apps Script 2-Way Live Sync Web App API
 * Synchronizes Schedule Center, Rescheduled Classes Hub, Parents & Multi-Admin Records
 */

const CONFIG = {
  SPREADSHEET_ID: "1TfKJieoXpdeua4u15csOq-gXuyuFy5pCyqFREDVO9BY",
  SHEET_SCHEDULES: "Schedule Center",
  SHEET_RESCHEDULES: "Rescheduled",
  SHEET_PARENTS: "Parents Database",
  SHEET_HISTORY: "Payment History",
  SHEET_AUDIT: "Audit Sync Log",
  SHEET_QUESTIONS: "Dynamic Questions"
};

function getSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet(ss, name, headers) {
  ss = ss || getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Dynamically locate the main schedule sheet tab (prioritizes Schedule Center)
 */
function findScheduleSheet(ss) {
  ss = ss || getSpreadsheet_();
  const priorityNames = [
    CONFIG.SHEET_SCHEDULES, // "Schedule Center"
    "Master Schedule Roster",
    "Lesson Schedule & Attendance Dashboard 1",
    "ProsperWork Student Lesson Dashboard 1",
    "ProsperWork Student Dashboard",
    "Sheet1"
  ];
  for (let p = 0; p < priorityNames.length; p++) {
    const s = ss.getSheetByName(priorityNames[p]);
    if (s && s.getLastRow() >= 1) return s;
  }

  // Search through all tabs for any sheet containing "schedule", "roster", "student", or "class"
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName().toLowerCase();
    if (name.includes("schedule") || name.includes("roster") || name.includes("student") || name.includes("class")) {
      return allSheets[i];
    }
  }
  return allSheets[0];
}

/**
 * Dynamically locate the Rescheduled classes tab (prioritizes Rescheduled)
 */
function findRescheduleSheet(ss) {
  ss = ss || getSpreadsheet_();
  const priorityNames = [
    CONFIG.SHEET_RESCHEDULES, // "Rescheduled"
    "Rescheduled Classes",
    "Reschedule Hub",
    "Reschedules"
  ];
  for (let p = 0; p < priorityNames.length; p++) {
    const s = ss.getSheetByName(priorityNames[p]);
    if (s) return s;
  }

  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName().toLowerCase();
    if (name.includes("resched")) return allSheets[i];
  }

  return getOrCreateSheet(ss, CONFIG.SHEET_RESCHEDULES, [
    "Student Name", "Parents Name", "Reschedule /Trial Date", "Scheduled Date", "Teacher", "Assigned Admin", "Status", "Notes"
  ]);
}

function headerKey(val) {
  return String(val || "").trim().toLowerCase();
}

/**
 * Map columns from the actual Schedule Center header row dynamically.
 * Live headers: Child Name, Parent Name, WhatsApp, Teacher, Mon..Sat,
 * Local Time, Status, LESSON START, Admin Assigned, Google Meet link, Zoom Link
 */
function getScheduleColMap(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const keys = headers.map(headerKey);

  function col(matchers, fallback, exactOnly) {
    for (let i = 0; i < keys.length; i++) {
      const h = keys[i];
      if (!h) continue;
      for (let m = 0; m < matchers.length; m++) {
        const needle = matchers[m];
        if (h === needle || (!exactOnly && h.indexOf(needle) !== -1)) return i + 1;
      }
    }
    return fallback;
  }

  const idCol = col(["schedule id"], 0);
  return {
    id: idCol,
    child: col(["child name", "student name", "child", "student"], 1),
    parent: col(["parent name", "parent"], 2),
    whatsapp: col(["whatsapp", "phone"], 3),
    teacher: col(["teacher", "instructor"], 4),
    mon: col(["mon", "monday"], 5, true),
    tue: col(["tue", "tuesday"], 6, true),
    wed: col(["wed", "wednesday"], 7, true),
    thu: col(["thu", "thursday"], 8, true),
    fri: col(["fri", "friday"], 9, true),
    sat: col(["sat", "saturday"], 10, true),
    time: col(["local time"], 11) || col(["time"], 11),
    status: col(["status"], 12),
    lessonStart: col(["lesson start"], 13),
    admin: col(["admin assigned"], 14) || col(["admin"], 14),
    meet: col(["google meet", "meet"], 15),
    zoom: col(["zoom"], 16),
    notes: col(["notes"], 0),
    parentId: col(["parent id"], 0),
    paymentAmount: col(["payment amount"], 0),
    paymentConvert: col(["payment convert"], 0),
    paymentMade: col(["payment made"], 0),
    paymentDate: col(["payment date"], 0),
    billing: col(["billing"], 0),
    customDate: col(["custom date"], 0),
    lastCol: lastCol
  };
}

/**
 * Map columns from the Rescheduled header row dynamically.
 * Live headers: Student Name, Parents Name, Reschedule /Trial Date, Scheduled Date, Teacher, Assigned Admin
 */
function getRescheduleColMap(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const keys = headers.map(headerKey);

  function col(matchers, fallback) {
    for (let i = 0; i < keys.length; i++) {
      const h = keys[i];
      if (!h) continue;
      for (let m = 0; m < matchers.length; m++) {
        if (h.indexOf(matchers[m]) !== -1) return i + 1;
      }
    }
    return fallback;
  }

  return {
    student: col(["student name", "student", "child name", "child"], 1),
    parent: col(["parents name", "parent name", "parent"], 2),
    rescheduleDate: col(["reschedule /trial date", "reschedule date", "rescheduled date", "trial date", "reschedule", "slot"], 3),
    scheduledDate: col(["scheduled date", "original date", "old date", "current date"], 4),
    teacher: col(["teacher not yet assigen", "teacher", "instructor"], 5),
    admin: col(["assigned admin", "admin assigned", "admin"], 6),
    status: col(["status"], 7),
    notes: col(["notes", "comments", "reason"], 8),
    lastCol: Math.max(lastCol, 8)
  };
}

function cell(row, col1based) {
  if (!col1based || col1based < 1) return "";
  return row[col1based - 1];
}

function findScheduleRow_(sheet, map, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const schedId = String((data && (data.schedule_id || data.id)) || "").trim().toUpperCase();
  const childName = String((data && (data.student_name || data.child_name)) || "").trim().toLowerCase();
  const teacher = String((data && (data.teacher || data.teacher_name)) || "").trim().toLowerCase();
  const timeVal = String((data && (data.previous_time || data.old_time || data.local_time || data.time || data.time_start)) || "").trim().toLowerCase();

  const width = map.lastCol;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  if (map.id && schedId && schedId.indexOf("SCH-") === 0) {
    for (let r = 0; r < values.length; r++) {
      const rowId = String(values[r][map.id - 1] || "").trim().toUpperCase();
      if (rowId === schedId) return r + 2;
    }
  }

  let nameMatches = [];
  for (let r = 0; r < values.length; r++) {
    const rowName = String(values[r][map.child - 1] || "").trim().toLowerCase();
    const rowA = String(values[r][0] || "").trim().toLowerCase();
    if (childName && (rowName === childName || rowA === childName)) nameMatches.push(r);
  }

  if (nameMatches.length === 1) return nameMatches[0] + 2;

  if (nameMatches.length > 1) {
    for (let i = 0; i < nameMatches.length; i++) {
      const r = nameMatches[i];
      const rowTeacher = String(values[r][map.teacher - 1] || "").trim().toLowerCase();
      const rowTime = String(values[r][map.time - 1] || "").trim().toLowerCase();
      if (teacher && timeVal && rowTeacher === teacher && rowTime === timeVal) return r + 2;
    }
    for (let i = 0; i < nameMatches.length; i++) {
      const r = nameMatches[i];
      const rowTeacher = String(values[r][map.teacher - 1] || "").trim().toLowerCase();
      if (teacher && rowTeacher === teacher) return r + 2;
    }
    for (let i = 0; i < nameMatches.length; i++) {
      const r = nameMatches[i];
      const rowTime = String(values[r][map.time - 1] || "").trim().toLowerCase();
      if (timeVal && rowTime === timeVal) return r + 2;
    }
    return nameMatches[0] + 2;
  }

  return -1;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🎼 Academy Sync")
    .addItem("🔄 Refresh Calculations", "refreshCalculations")
    .addItem("🌐 Web App Deployment Info", "showDeploymentGuide")
    .addToUi();
}

function showDeploymentGuide() {
  const html = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, sans-serif; padding: 18px; font-size: 13px; line-height: 1.6; color: #1E293B;">
      <h3 style="color: #4F46E5; margin-top: 0; margin-bottom: 12px; font-size: 16px;">🎼 Web App Deployment Instructions</h3>
      <ol style="padding-left: 20px; margin: 0 0 16px 0;">
        <li style="margin-bottom: 8px;">Click <b>Deploy &gt; New deployment</b> in the top right.</li>
        <li style="margin-bottom: 8px;">Next to <i>"Select type"</i> on the left, click the <b>Gear icon ⚙️</b> and select <b>Web app</b>. <br><i style="color:#64748B;">(Note: "Execute as" only appears after clicking the ⚙️ gear icon!)</i></li>
        <li style="margin-bottom: 8px;">Set <b>Execute as:</b> <code>Me</code>.</li>
        <li style="margin-bottom: 8px;">Set <b>Who has access:</b> <code>Anyone</code> <i>(Important: Must be "Anyone", NOT "Only myself")</i>.</li>
        <li style="margin-bottom: 8px;">Click <b>Deploy</b> and authorize if prompted.</li>
        <li>Copy the <b>Web app URL</b> (ends in <code>/exec</code>) and save it in your website settings!</li>
      </ol>
      <div style="background: #F1F5F9; border-radius: 8px; padding: 10px; font-size: 12px; color: #475569;">
        💡 <b>Note for Existing Deployments:</b> If editing an existing deployment in <i>Manage deployments</i>, Google locks "Execute as" to the original setting and only shows the <b>Version</b> dropdown. Choose <b>New version</b> and click <b>Deploy</b>.
      </div>
    </div>
  `;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(500).setHeight(360), "Deployment Guide");
}

function refreshCalculations() {
  SpreadsheetApp.flush();
  getSpreadsheet_().toast("Calculations and sync status refreshed.", "Refreshed", 3);
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

var API_VERSION = "2-way-live-v4";

function applyIncomingAction_(ss, action, data) {
  action = String(action || "").trim();
  if (action === "save_reschedule" || action === "reschedule" || action === "new_reschedule" || action === "edit_reschedule") {
    return recordRescheduleNote(ss, data);
  } else if (action === "delete_reschedule" || action === "remove_reschedule") {
    return deleteRescheduleRow(ss, data);
  } else if (action === "deleteSchedule" || action === "delete_schedule" || action === "delete_student" || action === "delete_class") {
    return deleteScheduleRow(ss, data);
  } else if (action === "saveSchedule" || action === "new_student" || action === "edit_student") {
    return saveScheduleRow(ss, data);
  } else if (action === "assign_admin") {
    return assignAdminToStudent(ss, data);
  } else if (action === "new_admin") {
    logAudit(ss, "admin", data.name || "Admin", "CREATE_ADMIN", "Website", "New Admin registered: " + (data.name || ""));
    SpreadsheetApp.flush();
    return "Admin " + (data.name || "") + " recorded.";
  } else if (action === "completeSchedule") {
    return markScheduleCompleted(ss, data);
  } else if (action === "archiveSchedule") {
    return archiveScheduleRow(ss, data);
  } else if (action === "recordPayment" || action === "savePayment" || action === "updatePayment") {
    return recordPaymentRow(ss, data);
  } else if (action === "deletePayment" || action === "removePayment") {
    return deletePaymentRow(ss, data);
  } else if (action === "saveParent" || action === "save_parent" || action === "new_parent" || action === "edit_parent") {
    return saveParentRow(ss, data);
  } else if (action === "deleteParent" || action === "delete_parent" || action === "remove_parent") {
    return deleteParentRow(ss, data);
  } else if (action === "syncAll") {
    return syncBulkSchedules(ss, data.schedules || (Array.isArray(data) ? data : []));
  }
  return saveScheduleRow(ss, data);
}

/**
 * Handles GET requests from Netlify / Web App
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || "").trim();

    if (action === "ping") {
      return createJsonResponse({ status: "success", api_version: API_VERSION });
    }

    if (action && action !== "get_all") {
      let data = params;
      if (params.data) {
        try { data = JSON.parse(params.data); } catch (err) { data = params; }
      }
      const ssWrite = getSpreadsheet_();
      const msg = applyIncomingAction_(ssWrite, action, data);
      return createJsonResponse({ status: "success", api_version: API_VERSION, message: msg });
    }

    const ss = getSpreadsheet_();
    const schedSheet = findScheduleSheet(ss);
    const reschSheet = findRescheduleSheet(ss);

    // If opened directly in a browser tab without parameters, show friendly HTML status dashboard
    if (!action && (!e || !e.parameter || Object.keys(e.parameter).length === 0)) {
      const html = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <title>Prosper Music Academy — Sync API</title>
          <style>
            body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0F172A; color: #F8FAFC; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }
            .card { background: #1E293B; border: 1px solid #334155; border-radius: 16px; padding: 32px; max-width: 480px; width: 100%; box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.5); text-align: center; }
            .badge { display: inline-flex; align-items: center; background: rgba(16, 185, 129, 0.15); color: #34D399; padding: 6px 14px; border-radius: 9999px; font-weight: 700; font-size: 13px; margin-bottom: 20px; }
            .dot { width: 8px; height: 8px; border-radius: 50%; background: #10B981; margin-right: 8px; }
            h1 { font-size: 20px; margin: 0 0 10px 0; color: #FFFFFF; font-weight: 800; }
            p { font-size: 13px; color: #94A3B8; margin: 0 0 24px 0; line-height: 1.5; }
            .stat-box { background: #0F172A; border: 1px solid #334155; border-radius: 12px; padding: 14px; text-align: left; font-size: 13px; margin-bottom: 20px; }
            .stat-row { display: flex; justify-content: space-between; margin-bottom: 8px; }
            .stat-row:last-child { margin-bottom: 0; }
            .stat-label { color: #94A3B8; }
            .stat-val { color: #F8FAFC; font-weight: 600; font-family: monospace; }
          </style>
        </head>
        <body>
          <div class="card">
            <div class="badge"><div class="dot"></div>Live 2-Way API Connected</div>
            <h1>🎼 Prosper Music Academy Sync API</h1>
            <p>The Google Sheets synchronization web app is active, verified, and accepting requests from your website.</p>
            <div class="stat-box">
              <div class="stat-row"><span class="stat-label">Spreadsheet:</span><span class="stat-val">${ss.getName()}</span></div>
              <div class="stat-row"><span class="stat-label">Schedule Center:</span><span class="stat-val">${schedSheet ? schedSheet.getLastRow() + ' rows' : 'Connected'}</span></div>
              <div class="stat-row"><span class="stat-label">Rescheduled Hub:</span><span class="stat-val">${reschSheet ? reschSheet.getLastRow() + ' rows' : 'Connected'}</span></div>
              <div class="stat-row"><span class="stat-label">API Version:</span><span class="stat-val">${API_VERSION}</span></div>
            </div>
            <p style="font-size: 12px; color: #64748B; margin-bottom: 0;">You can close this tab and return to your Prosper Music Academy website.</p>
          </div>
        </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(html).setTitle("Prosper Music Academy — Sync API");
    }

    const schedules = [];

    if (schedSheet && schedSheet.getLastRow() >= 2) {
      const lastRow = schedSheet.getLastRow();
      const map = getScheduleColMap(schedSheet);
      const data = schedSheet.getRange(2, 1, lastRow - 1, map.lastCol).getValues();

      function checkDay(val) {
        if (val === true || val === 1) return true;
        const s = String(val || "").trim().toLowerCase();
        return s === "true" || s === "yes" || s === "y" || s === "x" || s === "1" || s === "✓";
      }

      for (let r = 0; r < data.length; r++) {
        const row = data[r];
        const hasData = row.some(function(c) { return c !== "" && c !== null && c !== undefined; });
        if (!hasData) continue;

        const studentName = String(cell(row, map.child) || "").trim();
        let id = map.id ? String(cell(row, map.id) || "").trim() : "";
        if (!id && studentName) {
          id = "SCH-" + String(100 + r);
        }
        if (!studentName && !id) continue;

        const parentName = String(cell(row, map.parent) || "").trim();
        const whatsapp = String(cell(row, map.whatsapp) || "").trim();
        const teacher = String(cell(row, map.teacher) || "").trim();

        let isMon = checkDay(cell(row, map.mon));
        let isTue = checkDay(cell(row, map.tue));
        let isWed = checkDay(cell(row, map.wed));
        let isThu = checkDay(cell(row, map.thu));
        let isFri = checkDay(cell(row, map.fri));
        let isSat = checkDay(cell(row, map.sat));

        const localTime = formatTimeOrDate(cell(row, map.time));
        const status = String(cell(row, map.status) || "ACTIVE").trim().toUpperCase();
        const lessonStart = formatDate(cell(row, map.lessonStart));
        const parentId = map.parentId ? String(cell(row, map.parentId) || "").trim() : "";
        const paymentAmount = map.paymentAmount ? String(cell(row, map.paymentAmount) || "").trim() : "";
        const paymentConvert = map.paymentConvert ? Number(cell(row, map.paymentConvert) || 0) : 0;
        const paymentMadeAndBal = map.paymentMade ? String(cell(row, map.paymentMade) || "").trim() : "";
        const paymentDate = map.paymentDate ? formatDate(cell(row, map.paymentDate)) : "";
        const billingTerms = map.billing ? String(cell(row, map.billing) || "Default (Next Month)").trim() : "Default (Next Month)";
        const customDate = map.customDate ? formatDate(cell(row, map.customDate)) : "";
        const notes = map.notes ? String(cell(row, map.notes) || "").trim() : "";
        const adminAssigned = String(cell(row, map.admin) || "").trim() || "Not Yet Assigned";
        const meetLink = map.meet ? String(cell(row, map.meet) || "").trim() : "";
        const zoomLink = map.zoom ? String(cell(row, map.zoom) || "").trim() : "";

        if (!isMon && !isTue && !isWed && !isThu && !isFri && !isSat) {
          const combinedText = (localTime + " " + notes).toLowerCase();
          if (combinedText.indexOf("mon") !== -1) isMon = true;
          if (combinedText.indexOf("tue") !== -1) isTue = true;
          if (combinedText.indexOf("wed") !== -1) isWed = true;
          if (combinedText.indexOf("thu") !== -1) isThu = true;
          if (combinedText.indexOf("fri") !== -1) isFri = true;
          if (combinedText.indexOf("sat") !== -1) isSat = true;
        }

        const days = [];
        if (isMon) days.push("Monday");
        if (isTue) days.push("Tuesday");
        if (isWed) days.push("Wednesday");
        if (isThu) days.push("Thursday");
        if (isFri) days.push("Friday");
        if (isSat) days.push("Saturday");

        let paid = 0, bal = paymentConvert;
        if (paymentMadeAndBal) {
          const m = paymentMadeAndBal.match(/Paid:\s*\$?([\d.]+)/i);
          const b = paymentMadeAndBal.match(/Bal:\s*\$?([\d.]+)/i);
          if (m) paid = parseFloat(m[1]) || 0;
          if (b) bal = parseFloat(b[1]) || 0;
        }

        schedules.push({
          schedule_id: id,
          id: id,
          student_name: studentName,
          child_name: studentName,
          parent_name: parentName,
          whatsapp: whatsapp,
          teacher: teacher,
          teacher_name: teacher,
          days_of_week: days.join(", "),
          days: days.join(", "),
          mon: isMon ? 1 : 0,
          tue: isTue ? 1 : 0,
          wed: isWed ? 1 : 0,
          thu: isThu ? 1 : 0,
          fri: isFri ? 1 : 0,
          sat: isSat ? 1 : 0,
          local_time: localTime,
          time: localTime,
          time_start: localTime.split("-")[0] ? localTime.split("-")[0].trim() : localTime,
          time_end: localTime.split("-")[1] ? localTime.split("-")[1].trim() : "",
          status: status,
          lesson_start: lessonStart,
          parent_id: parentId,
          payment_amount: paymentAmount,
          total_fee: paymentConvert,
          payment_convert: paymentConvert,
          paid_amount: paid,
          balance: bal,
          payment_status: bal <= 0 && paymentConvert > 0 ? "PAID" : (paid > 0 ? "PARTIAL" : "UNPAID"),
          payment_made_and_balance: paymentMadeAndBal,
          payment_date: paymentDate,
          billing_terms: billingTerms,
          billing_frequency: billingTerms,
          custom_date: customDate,
          notes: notes,
          admin_notes: adminAssigned,
          admin: adminAssigned,
          meet_link: meetLink,
          zoom_link: zoomLink,
          last_synced: new Date().toISOString(),
          sync_status: "SYNCED"
        });
      }
    }

    // 2. Read Reschedules from Rescheduled tab
    const reschedules = readRescheduleHub_(ss);

    // 3. Read Parents from Parents Database tab
    const parentSheet = ss.getSheetByName(CONFIG.SHEET_PARENTS);
    const parents = [];
    if (parentSheet && parentSheet.getLastRow() >= 2) {
      const pData = parentSheet.getRange(2, 1, parentSheet.getLastRow() - 1, Math.min(parentSheet.getLastColumn(), 10)).getValues();
      for (let r = 0; r < pData.length; r++) {
        const row = pData[r];
        if (row[0] || row[1]) {
          const pid = String(row[0] || ("PAR-" + String(100 + r))).trim();
          const pname = String(row[1] || "").trim();
          const phone = String(row[2] || "").trim();
          const email = String(row[3] || "").trim();
          const addr = String(row[4] || "").trim();
          const rel = String(row[5] || "Mother").trim();
          const studs = String(row[6] || "").split(",").map(x => x.trim()).filter(Boolean);
          const st = String(row[7] || "ACTIVE").trim().toUpperCase();
          const pnotes = String(row[8] || "").trim();
          const updDate = formatDate(row[9]);

          parents.push({
            id: pid,
            parent_id: pid,
            name: pname,
            parent_name: pname,
            phone: phone,
            whatsapp: phone,
            email: email,
            address: addr,
            relationship: rel,
            associated_students: studs,
            status: st === "INACTIVE" ? "INACTIVE" : "ACTIVE",
            notes: pnotes,
            currency: "GBP",
            billing_terms: "Default (Next Month)",
            updated_date: updDate
          });
        }
      }
    }

    // 4. Read Payment History
    const history = readPaymentHistory_(ss);

    // 5. Collect all unique admins from schedules + reschedules + defaults
    const sheetAdmins = new Set(["Admin Deborah", "Admin Gloria", "Admin Ojo", "Admin Prosper", "Office"]);
    schedules.forEach(s => {
      if (s.admin && s.admin !== "Not Yet Assigned" && String(s.admin).trim()) {
        sheetAdmins.add(String(s.admin).trim());
      }
    });
    reschedules.forEach(tr => {
      if (tr.admin && tr.admin !== "Not Yet Assigned" && String(tr.admin).trim()) {
        sheetAdmins.add(String(tr.admin).trim());
      }
    });

    return createJsonResponse({
      status: "success",
      api_version: API_VERSION,
      sheet_name: schedSheet ? schedSheet.getName() : "Unknown",
      timestamp: new Date().toISOString(),
      schedules: schedules,
      reschedules: reschedules,
      parents: parents,
      history: history,
      admins: Array.from(sheetAdmins)
    });

  } catch (err) {
    return createJsonResponse({ status: "error", message: err.message });
  }
}

/**
 * Handles POST requests from Netlify / Web App
 */
function doPost(e) {
  try {
    const rawData = e.postData ? e.postData.contents : "{}";
    const body = JSON.parse(rawData);
    const action = body.action || body.original_action || "saveSchedule";
    const data = body.schedule || body.reschedule || body.student || body.parent || body.data || body;
    const ss = getSpreadsheet_();
    const msg = applyIncomingAction_(ss, action, data);
    return createJsonResponse({ status: "success", api_version: API_VERSION, message: msg });
  } catch (err) {
    return createJsonResponse({ status: "error", api_version: API_VERSION, message: err.message });
  }
}

function setMapped_(sheet, row, col, value) {
  if (col && col >= 1) sheet.getRange(row, col).setValue(value);
}

function findFirstEmptyScheduleRow_(sheet, map) {
  const colIdx = (map && map.child && map.child >= 1) ? map.child : 1;
  const numRows = Math.min(sheet.getMaxRows(), 1000);
  if (numRows < 2) return 2;
  const values = sheet.getRange(2, colIdx, numRows - 1, 1).getValues();
  for (let r = 0; r < values.length; r++) {
    const val = String(values[r][0] || "").trim();
    if (!val || val === "") {
      return r + 2; // 1-based row index directly below last populated row
    }
  }
  return values.length + 2;
}

/**
 * Save or update a student schedule row in Google Sheets
 */
function saveScheduleRow(ss, s) {
  if (!s) return "No schedule payload provided.";

  const sheet = findScheduleSheet(ss);
  const map = getScheduleColMap(sheet);
  const childName = String(s.student_name || s.child_name || "").trim();
  const schedId = String(s.schedule_id || s.id || "").trim();

  if (!schedId && !childName) {
    return "No student name or schedule ID provided.";
  }

  let targetRow = findScheduleRow_(sheet, map, s);
  const creating = targetRow === -1;
  if (creating) {
    targetRow = findFirstEmptyScheduleRow_(sheet, map);
  }

  const teacher = s.teacher || s.teacher_name || "";
  const daysStr = String(s.days_of_week || s.days || "").toLowerCase();
  const hasDays = daysStr.length > 0 || s.mon === 1 || s.tue === 1 || s.wed === 1 || s.thu === 1 || s.fri === 1 || s.sat === 1;

  const isMon = (s.mon === 1 || s.mon === true || daysStr.indexOf("mon") !== -1);
  const isTue = (s.tue === 1 || s.tue === true || daysStr.indexOf("tue") !== -1);
  const isWed = (s.wed === 1 || s.wed === true || daysStr.indexOf("wed") !== -1);
  const isThu = (s.thu === 1 || s.thu === true || daysStr.indexOf("thu") !== -1);
  const isFri = (s.fri === 1 || s.fri === true || daysStr.indexOf("fri") !== -1);
  const isSat = (s.sat === 1 || s.sat === true || daysStr.indexOf("sat") !== -1);

  const localTime = s.local_time || s.time || (s.time_start ? (s.time_start + (s.time_end ? "-" + s.time_end : "")) : "");
  const rawStatus = String(s.status || "").toUpperCase();
  const status = /^(ACTIVE|PAUSE|PAUSED|DONE|COMPLETED|ARCHIVED)$/.test(rawStatus) ? rawStatus : "";
  const adminVal = s.admin || s.admin_assigned || s.admin_notes || "";

  if (map.id && schedId.indexOf("SCH-") === 0) setMapped_(sheet, targetRow, map.id, schedId);
  if (childName) setMapped_(sheet, targetRow, map.child, childName);
  if (creating || s.parent_name) setMapped_(sheet, targetRow, map.parent, s.parent_name || "");
  if (creating || s.whatsapp) setMapped_(sheet, targetRow, map.whatsapp, s.whatsapp || "");
  if (teacher) setMapped_(sheet, targetRow, map.teacher, teacher);
  if (hasDays) {
    setMapped_(sheet, targetRow, map.mon, isMon ? true : false);
    setMapped_(sheet, targetRow, map.tue, isTue ? true : false);
    setMapped_(sheet, targetRow, map.wed, isWed ? true : false);
    setMapped_(sheet, targetRow, map.thu, isThu ? true : false);
    setMapped_(sheet, targetRow, map.fri, isFri ? true : false);
    setMapped_(sheet, targetRow, map.sat, isSat ? true : false);
  }
  if (localTime) setMapped_(sheet, targetRow, map.time, localTime);
  if (status) setMapped_(sheet, targetRow, map.status, status);
  if (s.lesson_start) setMapped_(sheet, targetRow, map.lessonStart, s.lesson_start);
  if (s.parent_id) setMapped_(sheet, targetRow, map.parentId, s.parent_id);
  if (Number(s.total_fee || s.payment_convert || s.fee || 0) > 0) {
    const totalFee = Number(s.total_fee || s.payment_convert || s.fee || 0);
    const paid = Number(s.paid_amount || s.paid || 0);
    const bal = Number(s.balance !== undefined ? s.balance : Math.max(0, totalFee - paid));
    setMapped_(sheet, targetRow, map.paymentAmount, s.payment_amount || ("$" + totalFee.toFixed(2)));
    setMapped_(sheet, targetRow, map.paymentConvert, totalFee);
    setMapped_(sheet, targetRow, map.paymentMade, s.payment_made_and_balance || ("Paid: $" + paid.toFixed(2) + ", Bal: $" + bal.toFixed(2)));
  }
  if (s.payment_date) setMapped_(sheet, targetRow, map.paymentDate, s.payment_date);
  if (s.billing_terms || s.billing_frequency) setMapped_(sheet, targetRow, map.billing, s.billing_terms || s.billing_frequency);
  if (s.custom_date) setMapped_(sheet, targetRow, map.customDate, s.custom_date);
  if (map.notes && s.notes) setMapped_(sheet, targetRow, map.notes, s.notes);
  if (adminVal) setMapped_(sheet, targetRow, map.admin, adminVal);
  if (s.meet_link) setMapped_(sheet, targetRow, map.meet, s.meet_link);
  if (s.zoom_link) setMapped_(sheet, targetRow, map.zoom, s.zoom_link);

  logAudit(ss, "schedule", schedId || childName, "SAVE", "Website", "Saved " + childName + " in row " + targetRow);
  SpreadsheetApp.flush();
  return "Schedule " + (schedId || childName) + " saved in spreadsheet row " + targetRow;
}

function assignAdminToStudent(ss, data) {
  const sheet = findScheduleSheet(ss);
  const map = getScheduleColMap(sheet);
  const newAdmin = String(data.admin || data.admin_assigned || "Not Yet Assigned").trim();
  const targetRow = findScheduleRow_(sheet, map, data);
  if (targetRow === -1) {
    return "Student " + (data.id || data.child_name || data.student_name || "") + " not found to assign admin.";
  }
  setMapped_(sheet, targetRow, map.admin, newAdmin);
  logAudit(ss, "schedule", data.id || data.child_name || data.student_name, "ASSIGN_ADMIN", "Website", "Assigned admin: " + newAdmin);
  SpreadsheetApp.flush();
  return "Admin " + newAdmin + " assigned to row " + targetRow;
}

function extractDayFromSlot_(slot) {
  const s = String(slot || "").toLowerCase();
  if (s.indexOf("mon") !== -1) return "Monday";
  if (s.indexOf("tue") !== -1) return "Tuesday";
  if (s.indexOf("wed") !== -1) return "Wednesday";
  if (s.indexOf("thu") !== -1) return "Thursday";
  if (s.indexOf("fri") !== -1) return "Friday";
  if (s.indexOf("sat") !== -1) return "Saturday";
  return "";
}

function extractTimeFromSlot_(slot) {
  const s = String(slot || "");
  const m = s.match(/(\d{1,2}(:\d{2})?\s*[:.]?\s*(am|pm)?)/i);
  return m ? String(m[0]).trim() : "";
}

/**
 * Reads all rows from the live Rescheduled tab
 */
function readRescheduleHub_(ss) {
  const out = [];
  try {
    const sheet = findRescheduleSheet(ss);
    if (!sheet || sheet.getLastRow() < 2) return out;
    const map = getRescheduleColMap(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, map.lastCol).getValues();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const studentName = String(cell(row, map.student) || "").trim();
      if (!studentName) continue;
      out.push({
        id: "resch_row_" + (i + 2),
        student_name: studentName,
        child_name: studentName,
        parent_name: String(cell(row, map.parent) || "").trim(),
        reschedule_date: String(cell(row, map.rescheduleDate) || "").trim(),
        scheduled_date: formatDate(cell(row, map.scheduledDate)),
        teacher: String(cell(row, map.teacher) || "Unassigned").trim(),
        admin: String(cell(row, map.admin) || "Office").trim() || "Office",
        status: String(cell(row, map.status) || "CONFIRMED").trim() || "CONFIRMED",
        notes: String(cell(row, map.notes) || "").trim()
      });
    }
  } catch (e) {
    Logger.log("Error reading reschedules: " + e);
  }
  return out;
}

/**
 * Saves a reschedule directly into the Rescheduled sheet tab
 * AND updates the student's lesson slot in Schedule Center
 */
function recordRescheduleNote(ss, data) {
  const slot = String((data && (data.reschedule_date || data.new_slot)) || "").trim();
  const parsedTime = data.local_time || data.time || extractTimeFromSlot_(slot);
  const parsedDays = data.days_of_week || data.days || extractDayFromSlot_(slot);

  // 1. Save in the Rescheduled tab
  const reschResult = upsertRescheduleHub_(ss, data, parsedDays, parsedTime);

  // 2. Also update Schedule Center if this student is enrolled
  try {
    const schedSheet = findScheduleSheet(ss);
    const map = getScheduleColMap(schedSheet);
    const targetRow = findScheduleRow_(schedSheet, map, data);
    if (targetRow !== -1) {
      if (parsedDays) {
        const dLower = parsedDays.toLowerCase();
        setMapped_(schedSheet, targetRow, map.mon, dLower.includes("mon"));
        setMapped_(schedSheet, targetRow, map.tue, dLower.includes("tue"));
        setMapped_(schedSheet, targetRow, map.wed, dLower.includes("wed"));
        setMapped_(schedSheet, targetRow, map.thu, dLower.includes("thu"));
        setMapped_(schedSheet, targetRow, map.fri, dLower.includes("fri"));
        setMapped_(schedSheet, targetRow, map.sat, dLower.includes("sat"));
      }
      if (parsedTime) setMapped_(schedSheet, targetRow, map.time, parsedTime);
      if (data.teacher || data.teacher_name) setMapped_(schedSheet, targetRow, map.teacher, data.teacher || data.teacher_name);
      if (data.admin) setMapped_(schedSheet, targetRow, map.admin, data.admin);

      if (map.notes) {
        const prevNotes = schedSheet.getRange(targetRow, map.notes).getValue() || "";
        schedSheet.getRange(targetRow, map.notes).setValue(prevNotes + " | Rescheduled to: " + (slot || (parsedDays + " " + parsedTime)));
      }
    }
  } catch(e) {
    Logger.log("Could not update Schedule Center row: " + e);
  }

  SpreadsheetApp.flush();
  return reschResult;
}

/**
 * Upsert into the Rescheduled sheet tab matching its exact columns
 */
function upsertRescheduleHub_(ss, data, parsedDays, parsedTime) {
  const sheet = findRescheduleSheet(ss);
  const map = getRescheduleColMap(sheet);
  const studentName = String((data && (data.student_name || data.child_name)) || "").trim();
  if (!studentName) return "No student name provided for reschedule.";

  const slot = String((data && (data.reschedule_date || data.new_slot)) || ((parsedDays || "") + " " + (parsedTime || ""))).trim();
  const orig = String((data && (data.scheduled_date || data.previous_days || data.original_slot)) || "").trim();
  const teacher = String((data && (data.teacher || data.teacher_name)) || "Unassigned").trim();
  const admin = String((data && (data.admin || data.admin_assigned)) || "Office").trim();
  const status = String((data && data.status) || "CONFIRMED").trim();
  const notes = String((data && data.notes) || "").trim();
  const parentName = String((data && data.parent_name) || "").trim();

  const last = sheet.getLastRow();
  let target = last + 1;
  if (last >= 2) {
    const rows = sheet.getRange(2, 1, last - 1, map.lastCol).getValues();
    for (let i = 0; i < rows.length; i++) {
      const n = String(cell(rows[i], map.student) || "").trim().toLowerCase();
      const s = String(cell(rows[i], map.rescheduleDate) || "").trim().toLowerCase();
      if (n === studentName.toLowerCase() && (s === slot.toLowerCase() || !s)) {
        target = i + 2;
        break;
      }
    }
  }

  setMapped_(sheet, target, map.student, studentName);
  if (parentName) setMapped_(sheet, target, map.parent, parentName);
  if (slot) setMapped_(sheet, target, map.rescheduleDate, slot);
  if (orig) setMapped_(sheet, target, map.scheduledDate, orig);
  if (teacher) setMapped_(sheet, target, map.teacher, teacher);
  if (admin) setMapped_(sheet, target, map.admin, admin);
  if (map.status) setMapped_(sheet, target, map.status, status);
  if (map.notes && notes) setMapped_(sheet, target, map.notes, notes);

  logAudit(ss, "reschedule", studentName, "SAVE_RESCHEDULE", "Website", "Saved reschedule: " + slot + " (Row " + target + ")");
  SpreadsheetApp.flush();
  return "Rescheduled session for " + studentName + " saved in row " + target;
}

/**
 * Delete a regular class/student row from the Schedule Center sheet tab
 */
function deleteScheduleRow(ss, data) {
  const sheet = findScheduleSheet(ss);
  const map = getScheduleColMap(sheet);
  const schedId = String((data && (data.schedule_id || data.id)) || "").trim().toUpperCase();
  const childName = String((data && (data.child_name || data.student_name)) || "").trim().toLowerCase();

  let targetRow = findScheduleRow_(sheet, map, data);

  // Fallback scan if findScheduleRow_ didn't find exact match
  if (targetRow === -1 && (schedId || childName) && sheet.getLastRow() >= 2) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, map.lastCol).getValues();
    for (let r = 0; r < rows.length; r++) {
      const rowId = String(cell(rows[r], map.id) || "").trim().toUpperCase();
      const rowName = String(cell(rows[r], map.child) || cell(rows[r], 1) || "").trim().toLowerCase();
      if ((schedId && rowId === schedId) || (childName && rowName === childName)) {
        targetRow = r + 2;
        break;
      }
    }
  }

  if (targetRow === -1) {
    return "Schedule not found to delete.";
  }

  sheet.deleteRow(targetRow);
  logAudit(ss, "schedule", schedId || childName, "DELETE_SCHEDULE", "Website", "Deleted schedule row " + targetRow);
  SpreadsheetApp.flush();
  return "Schedule class for " + (childName || schedId) + " deleted from Google Sheet.";
}

/**
 * Delete a rescheduled class row from the Rescheduled sheet tab
 */
function deleteRescheduleRow(ss, data) {
  const sheet = findRescheduleSheet(ss);
  const map = getRescheduleColMap(sheet);
  const studentName = String((data && (data.student_name || data.child_name)) || "").trim().toLowerCase();
  const slot = String((data && (data.reschedule_date || data.slot)) || "").trim().toLowerCase();
  if (!studentName) return "No student name provided.";

  const last = sheet.getLastRow();
  if (last < 2) return "No rescheduled classes found.";

  const rows = sheet.getRange(2, 1, last - 1, map.lastCol).getValues();
  // 1. Try match by name AND slot
  for (let i = 0; i < rows.length; i++) {
    const n = String(cell(rows[i], map.student) || "").trim().toLowerCase();
    const s = String(cell(rows[i], map.rescheduleDate) || "").trim().toLowerCase();
    if (n === studentName && (!slot || s === slot || s.indexOf(slot) !== -1 || slot.indexOf(s) !== -1)) {
      sheet.deleteRow(i + 2);
      logAudit(ss, "reschedule", studentName, "DELETE_RESCHEDULE", "Website", "Deleted reschedule row " + (i + 2));
      SpreadsheetApp.flush();
      return "Rescheduled class for " + studentName + " removed from Google Sheet.";
    }
  }

  // 2. Fallback: match by student name alone
  for (let i = 0; i < rows.length; i++) {
    const n = String(cell(rows[i], map.student) || "").trim().toLowerCase();
    if (n === studentName) {
      sheet.deleteRow(i + 2);
      logAudit(ss, "reschedule", studentName, "DELETE_RESCHEDULE", "Website", "Deleted reschedule row " + (i + 2));
      SpreadsheetApp.flush();
      return "Rescheduled class for " + studentName + " removed from Google Sheet.";
    }
  }

  return "Rescheduled class not found to delete.";
}

function recordAttendanceNote(ss, data) {
  const sheet = findScheduleSheet(ss);
  const map = getScheduleColMap(sheet);
  const targetRow = findScheduleRow_(sheet, map, {
    id: data.student_id || data.id,
    schedule_id: data.student_id || data.schedule_id,
    child_name: data.student_name || data.child_name,
    teacher: data.teacher,
    time: data.time
  });
  if (targetRow === -1) return "Schedule not found for attendance.";

  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  if (map.notes) {
    const prevNotes = sheet.getRange(targetRow, map.notes).getValue() || "";
    sheet.getRange(targetRow, map.notes).setValue(prevNotes + " | [" + nowStr + "] Attendance: " + (data.status || "DONE"));
  }
  logAudit(ss, "schedule", data.student_id || data.student_name, "ATTENDANCE", "Website", data.notes || data.status || "");
  SpreadsheetApp.flush();
  return "Attendance recorded for " + (data.student_id || data.student_name || "");
}

function markScheduleCompleted(ss, data) {
  const sheet = findScheduleSheet(ss);
  const map = getScheduleColMap(sheet);
  const targetRow = findScheduleRow_(sheet, map, data);
  if (targetRow === -1) return "Schedule " + (data.schedule_id || data.id || "") + " not found.";
  setMapped_(sheet, targetRow, map.status, "COMPLETED");
  if (map.notes) {
    const prevNotes = sheet.getRange(targetRow, map.notes).getValue() || "";
    sheet.getRange(targetRow, map.notes).setValue(prevNotes + " | Completed: " + (data.period || ""));
  }
  logAudit(ss, "schedule", data.schedule_id || data.id, "COMPLETE", "Website", "Marked completed for period: " + data.period);
  SpreadsheetApp.flush();
  return "Schedule marked completed.";
}

function archiveScheduleRow(ss, data) {
  const sheet = findScheduleSheet(ss);
  const map = getScheduleColMap(sheet);
  const targetRow = findScheduleRow_(sheet, map, data);
  if (targetRow === -1) return "Schedule " + (data.schedule_id || data.id || "") + " not found.";
  setMapped_(sheet, targetRow, map.status, "ARCHIVED");
  logAudit(ss, "schedule", data.schedule_id || data.id, "ARCHIVE", "Website", "Archived safely");
  SpreadsheetApp.flush();
  return "Schedule archived.";
}

function saveParentRow(ss, data) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_PARENTS);
  if (!sheet) {
    sheet = getOrCreateSheet(ss, CONFIG.SHEET_PARENTS, [
      "Parent ID", "Parent Name", "Phone", "Email", "Address", "Relationship", "Associated Students", "Status", "Notes", "Updated Date"
    ]);
  }
  const pid = String(data.parent_id || data.id || ("PAR-" + String(Date.now()).slice(-4))).trim();
  const pname = String(data.parent_name || data.name || "").trim();
  if (!pname && !pid) return "No parent name or ID provided.";

  const lastRow = sheet.getLastRow();
  let targetRow = lastRow + 1;
  if (lastRow >= 2) {
    const pData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let r = 0; r < pData.length; r++) {
      const rowPid = String(pData[r][0] || "").trim();
      const rowPname = String(pData[r][1] || "").trim().toLowerCase();
      if ((pid && rowPid === pid) || (pname && rowPname === pname.toLowerCase())) {
        targetRow = r + 2;
        break;
      }
    }
  }

  const studentsStr = Array.isArray(data.associated_students) ? data.associated_students.join(", ") : String(data.associated_students || data.students || "").trim();
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

  sheet.getRange(targetRow, 1, 1, 10).setValues([[
    pid,
    pname,
    String(data.whatsapp || data.phone || "").trim(),
    String(data.email || "").trim(),
    String(data.address || "").trim(),
    String(data.relationship || "Mother").trim(),
    studentsStr,
    String(data.status || "ACTIVE").trim().toUpperCase(),
    String(data.notes || "").trim(),
    nowStr
  ]]);

  logAudit(ss, "parent", pid, "SAVE_PARENT", "Website", `Saved parent profile for ${pname} (${pid}) in row ${targetRow}`);
  SpreadsheetApp.flush();
  return `Parent ${pname || pid} saved in row ${targetRow}.`;
}

function deleteParentRow(ss, data) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_PARENTS);
  if (!sheet || sheet.getLastRow() < 2) return "Parents database empty.";
  const pid = String(data.parent_id || data.id || "").trim();
  const pname = String(data.parent_name || data.name || "").trim().toLowerCase();
  if (!pid && !pname) return "No parent ID or name provided to delete.";

  const lastRow = sheet.getLastRow();
  const pData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let r = 0; r < pData.length; r++) {
    const rowPid = String(pData[r][0] || "").trim();
    const rowPname = String(pData[r][1] || "").trim().toLowerCase();
    if ((pid && rowPid === pid) || (pname && rowPname === pname)) {
      sheet.deleteRow(r + 2);
      logAudit(ss, "parent", pid || pname, "DELETE_PARENT", "Website", `Deleted parent profile ${pname} (${pid})`);
      SpreadsheetApp.flush();
      return `Parent ${pname || pid} deleted.`;
    }
  }
  return `Parent ${pid || pname} not found to delete.`;
}

function readPaymentHistory_(ss) {
  const out = [];
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEET_HISTORY);
    if (!sheet || sheet.getLastRow() < 2) return out;
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 7)).getValues();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row[0] && row[1]) {
        out.push({
          id: String(row[0]).trim(),
          parentId: String(row[1]).trim(),
          amount: Number(row[2]) || 0,
          date: formatDate(row[3]),
          method: String(row[4] || "Bank Transfer").trim(),
          ref: String(row[5] || "").trim(),
          notes: String(row[6] || "").trim()
        });
      }
    }
  } catch(e) {}
  return out;
}

function recordPaymentRow(ss, data) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  if (!sheet) {
    sheet = getOrCreateSheet(ss, CONFIG.SHEET_HISTORY, ["Transaction ID", "Parent ID", "Amount", "Payment Date", "Payment Method", "Reference", "Notes"]);
  }
  const txnId = data.id || ("TXN-" + String(Date.now()).slice(-6));
  const lastRow = sheet.getLastRow();
  let targetRow = lastRow + 1;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let r = 0; r < ids.length; r++) {
      if (String(ids[r][0]).trim() === txnId) {
        targetRow = r + 2;
        break;
      }
    }
  }
  sheet.getRange(targetRow, 1, 1, 7).setValues([[
    txnId,
    data.parentId || data.parent_id || "",
    Number(data.amount) || 0,
    data.date || formatDate(new Date()),
    data.method || "Bank Transfer",
    data.ref || "",
    data.notes || ""
  ]]);
  logAudit(ss, "payment", txnId, "RECORD_PAYMENT", "Website", `Recorded payment of ${data.amount} for parent ${data.parentId}`);
  SpreadsheetApp.flush();
  return `Payment ${txnId} saved.`;
}

function deletePaymentRow(ss, data) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return "Payment history sheet empty.";
  const txnId = String(data.id || data.txn_id || "").trim();
  if (!txnId) return "No transaction ID provided.";
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let r = 0; r < ids.length; r++) {
    if (String(ids[r][0]).trim() === txnId) {
      sheet.deleteRow(r + 2);
      logAudit(ss, "payment", txnId, "DELETE_PAYMENT", "Website", `Deleted payment transaction ${txnId}`);
      SpreadsheetApp.flush();
      return `Payment ${txnId} deleted.`;
    }
  }
  return `Payment ${txnId} not found.`;
}

function syncBulkSchedules(ss, schedules) {
  if (!Array.isArray(schedules) || schedules.length === 0) return "No schedules provided.";
  let count = 0;
  for (let i = 0; i < schedules.length; i++) {
    saveScheduleRow(ss, schedules[i]);
    count++;
  }
  SpreadsheetApp.flush();
  return `Bulk synced ${count} schedules.`;
}

function logAudit(ss, entity, recordId, action, source, details) {
  try {
    let auditSheet = ss.getSheetByName(CONFIG.SHEET_AUDIT);
    if (!auditSheet) {
      auditSheet = getOrCreateSheet(ss, CONFIG.SHEET_AUDIT, ["Timestamp", "Entity", "Record ID", "Action", "Source", "Details"]);
    }
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    auditSheet.appendRow([nowStr, entity, recordId, action, source, details]);
  } catch (e) {}
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

function formatTimeOrDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const hours = val.getHours();
    const minutes = val.getMinutes();
    const period = hours >= 12 ? "PM" : "AM";
    const displayHr = hours % 12 === 0 ? 12 : hours % 12;
    const mnStr = minutes < 10 ? "0" + minutes : minutes;
    return (displayHr < 10 ? "0" + displayHr : displayHr) + ":" + mnStr + " " + period;
  }
  return String(val);
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testConnection() {
  const ss = getSpreadsheet_();
  const schedSheet = findScheduleSheet(ss);
  const reschSheet = findRescheduleSheet(ss);
  const msg = "Connected to: " + ss.getName() + 
    "\nSchedule Tab: " + (schedSheet ? schedSheet.getName() : "(none)") + " (Rows: " + (schedSheet ? schedSheet.getLastRow() : 0) + ")" +
    "\nRescheduled Tab: " + (reschSheet ? reschSheet.getName() : "(none)") + " (Rows: " + (reschSheet ? reschSheet.getLastRow() : 0) + ")";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}
