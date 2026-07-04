/**********************************************************************
 * AUTOSERVE — VEHICLE STATUS TRACKER — GOOGLE APPS SCRIPT BACKEND
 * ---------------------------------------------------------------
 * Sheet structure required (create these 3 tabs in one Spreadsheet):
 *
 * TAB 1: "Users"
 *   Username | PIN | FullName | Role
 *   ravi     | 1234 | Ravi Kumar | Mechanic
 *   admin    | 9999 | Owner      | Manager
 *
 * TAB 2: "Lists"   (dropdown config, sheet-driven like AutoServe)
 *   Mechanics
 *   Ravi
 *   Kumar
 *   Suresh
 *
 * TAB 3: "JobCards"
 *   JobID | VehicleNo | CustomerName | Phone | IntakeDate | Status |
 *   IntakeTime | StartTime | FinishTime | DeliveredTime | Mechanic | CreatedBy
 *
 * DEPLOY: Extensions > Apps Script > paste this code > Deploy > 
 *         New deployment > Web app > Execute as: Me > Who has access: Anyone
 *         Copy the /exec URL into GAS_URL in the HTML file.
 * IMPORTANT: Redeploy (new version) every time you edit this script.
 **********************************************************************/

var SS = SpreadsheetApp.getActiveSpreadsheet();
var SHEET_USERS = "Users";
var SHEET_LISTS = "Lists";
var SHEET_JOBS = "JobCards";

function doGet(e) {
  var action = e.parameter.action;
  var result;
  try {
    if (action === "login") result = login(e.parameter.username, e.parameter.pin);
    else if (action === "getLists") result = getLists();
    else if (action === "getJobs") result = getJobs();
    else if (action === "addJob") result = addJob(e.parameter);
    else if (action === "updateStatus") result = updateStatus(e.parameter);
    else if (action === "assignMechanic") result = assignMechanic(e.parameter);
    else result = { success: false, message: "Unknown action" };
  } catch (err) {
    result = { success: false, message: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- LOGIN ---------------- */
function login(username, pin) {
  var sheet = SS.getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colUser = headers.indexOf("Username");
  var colPin = headers.indexOf("PIN");
  var colName = headers.indexOf("FullName");
  var colRole = headers.indexOf("Role");

  for (var i = 1; i < data.length; i++) {
    var rowUser = String(data[i][colUser]).trim().toLowerCase();
    var rowPin = String(data[i][colPin]).trim();
    if (rowUser === String(username).trim().toLowerCase() && rowPin === String(pin).trim()) {
      return {
        success: true,
        username: data[i][colUser],
        fullName: data[i][colName],
        role: data[i][colRole]
      };
    }
  }
  return { success: false, message: "Invalid username or PIN" };
}

/* ---------------- LISTS (dropdown config) ---------------- */
function getLists() {
  var sheet = SS.getSheetByName(SHEET_LISTS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var lists = {};
  headers.forEach(function (h, colIdx) {
    var values = [];
    for (var i = 1; i < data.length; i++) {
      var val = data[i][colIdx];
      if (val !== "" && val !== null) values.push(val);
    }
    lists[h] = values;
  });
  return { success: true, lists: lists };
}

/* ---------------- GET ALL JOBS ---------------- */
function getJobs() {
  var sheet = SS.getSheetByName(SHEET_JOBS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var jobs = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][headers.indexOf("JobID")] === "") continue;
    var job = {};
    headers.forEach(function (h, colIdx) {
      var val = data[i][colIdx];
      if (val instanceof Date) {
        // Format dates/times as strings for JSON
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      }
      job[h] = val;
    });
    jobs.push(job);
  }
  return { success: true, jobs: jobs };
}

/* ---------------- ADD NEW JOB (INTAKE) ---------------- */
function addJob(p) {
  var sheet = SS.getSheetByName(SHEET_JOBS);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var jobId = "JC" + new Date().getTime();
  var now = new Date();
  var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");

  // intakeDate comes from the date picker (yyyy-MM-dd); default to today if missing
  var intakeDate = p.intakeDate || Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  var rowData = {
    JobID: jobId,
    VehicleNo: p.vehicleNo || "",
    CustomerName: p.customerName || "",
    Phone: p.phone || "",
    IntakeDate: intakeDate,
    Status: "intake",
    IntakeTime: nowStr,
    StartTime: "",
    FinishTime: "",
    DeliveredTime: "",
    Mechanic: p.mechanic || "",
    CreatedBy: p.username || ""
  };

  var newRow = headers.map(function (h) { return rowData[h] !== undefined ? rowData[h] : ""; });
  sheet.appendRow(newRow);

  return { success: true, jobId: jobId };
}

/* ---------------- UPDATE STATUS (Intake -> Start -> Finished -> Delivered) ---------------- */
function updateStatus(p) {
  var sheet = SS.getSheetByName(SHEET_JOBS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colJobId = headers.indexOf("JobID");
  var colStatus = headers.indexOf("Status");

  var timeColMap = {
    intake: "IntakeTime",
    start: "StartTime",
    finished: "FinishTime",
    delivered: "DeliveredTime"
  };
  var timeCol = headers.indexOf(timeColMap[p.status]);
  var now = new Date();
  var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colJobId]).trim() === String(p.jobId).trim()) {
      sheet.getRange(i + 1, colStatus + 1).setValue(p.status);
      if (timeCol > -1) {
        sheet.getRange(i + 1, timeCol + 1).setValue(nowStr);
      }
      return { success: true, jobId: p.jobId, status: p.status, time: nowStr };
    }
  }
  return { success: false, message: "JobID not found" };
}

/* ---------------- ASSIGN / REASSIGN MECHANIC ---------------- */
function assignMechanic(p) {
  var sheet = SS.getSheetByName(SHEET_JOBS);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colJobId = headers.indexOf("JobID");
  var colMech = headers.indexOf("Mechanic");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colJobId]).trim() === String(p.jobId).trim()) {
      sheet.getRange(i + 1, colMech + 1).setValue(p.mechanic);
      return { success: true, jobId: p.jobId, mechanic: p.mechanic };
    }
  }
  return { success: false, message: "JobID not found" };
}
