/**
 * Cafe Tan90 — Attendance & Tasks
 * Google Apps Script backend
 *
 * How it works:
 * - One standalone Apps Script project, deployed as a Web App.
 * - Every log-in / log-out / task-done tap writes straight into a Google Sheet —
 *   nothing to "export".
 * - The first time anyone uses the Staff page in a new calendar month, the script
 *   checks whether this month already has a spreadsheet. If not, it creates a
 *   brand-new spreadsheet named "Cafe Tan90 Attendance — <Month Year>" (e.g.
 *   "Cafe Tan90 Attendance — August 2026") with one tab per day of that month
 *   (28-31 tabs), all pre-built up front. No manual trigger setup needed — the
 *   first use of the month may take a few extra seconds while it builds every
 *   day's tab, then it's instant after that.
 * - "Today" is always just today's date (Asia/Bangkok) — there's no manual "close
 *   the day" step. The moment the calendar date rolls over, every action starts
 *   writing to the next day's tab automatically, because getTodaySheet() always
 *   looks up whichever tab matches the current date.
 *
 * Employees are stored as up to MAX_EMPLOYEES "slots" (Script Properties, not
 * hardcoded), so the admin can add/rename/remove staff without editing code.
 * Each slot has a fixed column position on every day tab's task table (and a
 * fixed row on the attendance table) — adding someone just activates a free
 * slot, removing someone deactivates theirs. Nothing is ever deleted from the
 * Sheet, so past attendance history is never touched.
 *
 * Tasks, on the other hand, are stored as ROWS, not columns — one row per
 * task, growing downward with no fixed end. That's what makes the number of
 * tasks on a given day effectively unlimited instead of capped at a fixed
 * column count. (There's still one internal soft ceiling — see
 * TASK_ROW_SCAN_WINDOW below — but it's set far past anything a small cafe
 * would ever hit; it's an implementation detail, not a user-facing limit.)
 */

// ---- Configuration ----------------------------------------------------

var TIMEZONE = 'Asia/Bangkok';
var DRIVE_FOLDER_NAME = 'Cafe Tan90 Attendance Sheets';
var MAX_EMPLOYEES = 8;     // employee row-slots (Table 1) / column-slots (Table 2) per day tab
var EMPLOYEES_PROP_KEY = 'EMPLOYEES_V1';
var STANDING_TASKS_PROP_KEY = 'STANDING_TASKS_V1';
var WAGE_RATES_PROP_KEY = 'WAGE_RATES_V1';
var DEFAULT_WAGE_RATE = 40; // baht/hour, used until an admin sets a rate for someone

// How far down column A the script will look for task rows on a single day
// tab. Generous on purpose — a normal day has a handful of tasks — so this
// is a bookkeeping detail, not a task-count limit. If a day ever genuinely
// needs more than this many tasks (very unlikely for a small team), raise
// this number.
var TASK_ROW_SCAN_WINDOW = 400;

// Change this once you deploy (or set an Admin PIN via
// Project Settings > Script Properties > ADMIN_PIN so you don't have to edit code).
var DEFAULT_ADMIN_PIN = '2468';

// ---- Sheet layout -------------------------------------------------------
//
// Table 1 (Attendance): one row per employee slot (fixed, MAX_EMPLOYEES).
//   Columns: A Name, B Log In Time, C Log Out Time, D Hours Worked, E Present Today.
//
// Table 2 (Tasks): one row per TASK (unbounded — nothing else lives below it
//   on the sheet). Two header rows: row 1 has the task-name label plus each
//   employee's name (merged across a pair of columns); row 2 sub-labels each
//   half of that pair "Assigned" / "Done at". Then task rows start and just
//   keep going as needed. Column A is the task name; each employee slot i
//   gets columns taskAssignedCol_(i) (checkbox) and taskDoneCol_(i) (time).

var TABLE1_HEADER_ROW = 3;
var TABLE1_START_ROW = 4;
var TABLE1_END_ROW = TABLE1_START_ROW + MAX_EMPLOYEES - 1;
var TABLE1_COLS = 7; // Name, Log In, Log Out, Hours, Present Today, Login Edited, Logout Edited
// The last two columns are booleans, set true by setAttendanceTimes_ (an
// admin manually correcting a missed or wrong punch) and left blank/false
// otherwise. A day tab built before this feature existed simply has nothing
// in those columns, which reads back as blank/falsy, i.e. "not edited" -
// no migration needed for already-existing tabs.

var TABLE2_HEADER_ROW = TABLE1_END_ROW + 2; // one blank row between tables
var TABLE2_SUBHEADER_ROW = TABLE2_HEADER_ROW + 1;
var TABLE2_TASKS_START_ROW = TABLE2_SUBHEADER_ROW + 1;
var TABLE2_NAME_COL = 1;
var TABLE2_FIRST_EMP_COL = 2; // each employee slot uses a pair of columns starting here
var TABLE2_NUM_COLS = TABLE2_FIRST_EMP_COL - 1 + MAX_EMPLOYEES * 2; // name col + 2 per employee

function taskAssignedCol_(slotIdx) { return TABLE2_FIRST_EMP_COL + slotIdx * 2; }
function taskDoneCol_(slotIdx) { return taskAssignedCol_(slotIdx) + 1; }

// ---- Web app entry point ----------------------------------------------

// Always serves the same page — Staff.html is now a single page with
// Staff/Admin as tabs (Admin.html is pulled in as a body-only partial via
// include(), not served on its own). The optional ?page=admin query string
// still works as a deep link that pre-selects the Admin tab on load (see
// Staff.html's inline script) — it's just no longer a different template.
function doGet(e) {
  var tmpl = HtmlService.createTemplateFromFile('Staff');
  return tmpl.evaluate()
    .setTitle('Cafe Tan90 Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---- Employee roster (slot-based) --------------------------------------

function defaultEmployeeSlots_() {
  var defaults = ['Miw', 'Nureen', 'Richa'];
  var slots = [];
  for (var i = 0; i < MAX_EMPLOYEES; i++) {
    slots.push({ id: 'EMP' + (i + 1), name: defaults[i] || '', active: !!defaults[i], pin: '' });
  }
  return slots;
}

// Always returns exactly MAX_EMPLOYEES slots, in a fixed order that never
// changes — slot position is what ties a person to their row/columns on the
// sheet, so slots are only ever edited in place (name/active), never
// reordered.
function getAllEmployeeSlots_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(EMPLOYEES_PROP_KEY);
  if (raw) {
    try {
      var parsed = JSON.parse(raw);
      if (parsed && parsed.length === MAX_EMPLOYEES) return parsed;
    } catch (err) { /* fall through to reset below */ }
  }
  var def = defaultEmployeeSlots_();
  props.setProperty(EMPLOYEES_PROP_KEY, JSON.stringify(def));
  return def;
}

function saveEmployeeSlots_(slots) {
  PropertiesService.getScriptProperties().setProperty(EMPLOYEES_PROP_KEY, JSON.stringify(slots));
}

// Active employees only, each carrying its slot index (needed to compute
// which physical row/columns it lives on).
function getActiveEmployeesWithIndex_() {
  var slots = getAllEmployeeSlots_();
  var result = [];
  slots.forEach(function (s, i) {
    if (s.active) result.push({ id: s.id, name: s.name, idx: i, pin: s.pin || '' });
  });
  return result;
}

// Checks suppliedPin against one employee's own PIN, OR the admin PIN as a
// master override (covers a forgotten PIN, or admin acting on someone's
// behalf) - used by every employee-facing write/read that needs to know
// *which* employee is asking, not just that they know a valid PIN.
function assertEmployeePin_(employeeId, suppliedPin) {
  var slots = getAllEmployeeSlots_();
  var slot = null;
  for (var i = 0; i < slots.length; i++) {
    if (slots[i].id === employeeId) { slot = slots[i]; break; }
  }
  if (!slot || !slot.active) throw new Error('Unknown employee.');
  if (suppliedPin && suppliedPin === getAdminPin_()) return slot;
  if (!slot.pin) throw new Error('No PIN set for this employee yet - ask admin to set one.');
  if (suppliedPin !== slot.pin) throw new Error('Incorrect PIN.');
  return slot;
}

// Verifies a PIN identifies this employee and returns just their public info
// - what the Staff page calls to "unlock" someone's own row for the rest of
// that page session (nothing is persisted server-side; the client re-sends
// this same pin on every subsequent action for that person).
function identifyEmployee(employeeId, pin) {
  var slot = assertEmployeePin_(employeeId, pin);
  return { id: slot.id, name: slot.name };
}

// ---- Standing (recurring) tasks -----------------------------------------
// A standing task is a task that repeats automatically on chosen days of the
// week — e.g. "Make base milk" every Mon-Fri — with no admin action needed
// once it's set up. { id, name, days: [0=Sun..6=Sat, ...], assignments:
// { employeeId: bool }, active }. Stored the same way as the employee roster
// (Script Properties), so it survives redeploys. No cap on how many standing
// tasks can exist — that used to be limited because standing and one-off
// tasks shared a fixed set of day columns; now that tasks are rows, there's
// nothing to share.

function getAllStandingTasks_() {
  var raw = PropertiesService.getScriptProperties().getProperty(STANDING_TASKS_PROP_KEY);
  if (!raw) return [];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function saveStandingTasksList_(list) {
  PropertiesService.getScriptProperties().setProperty(STANDING_TASKS_PROP_KEY, JSON.stringify(list));
}

function getActiveStandingTasksForWeekday_(weekday) {
  return getAllStandingTasks_().filter(function (t) {
    return t.active && t.days && t.days.indexOf(weekday) !== -1;
  });
}

// Idempotent: appends a row for any standing task that applies to this
// date's weekday and isn't already on the sheet by name. Never overwrites a
// task that's already there (whether it was seeded earlier or hand-edited
// for this one day), so a per-day tweak never gets clobbered by a later
// resync. No column limit to run out of anymore — it just adds a row.
function applyStandingTasksToSheet_(sheet, date) {
  var standing = getActiveStandingTasksForWeekday_(date.getDay());
  if (!standing.length) return;

  var slots = getAllEmployeeSlots_();
  var taskRows = readTaskRows_(sheet);
  var existingNames = taskRows.names.map(function (n) { return String(n).trim().toLowerCase(); });
  var nextRow = TABLE2_TASKS_START_ROW + taskRows.count;

  standing.forEach(function (st) {
    var key = st.name.trim().toLowerCase();
    if (existingNames.indexOf(key) !== -1) return; // already on this day, don't duplicate

    sheet.getRange(nextRow, TABLE2_NAME_COL).setValue(st.name);
    slots.forEach(function (slot, i) {
      sheet.getRange(nextRow, taskAssignedCol_(i)).setValue(!!(st.assignments && st.assignments[slot.id]));
    });
    existingNames.push(key);
    nextRow++;
  });

  applyTaskAssignedCheckboxes_(sheet);
}

// After a standing task is added or edited, push it onto today's tab and
// every remaining tab this month that doesn't already have it — same
// "rest of the month, not the past" scope as syncEmployeeNamesForRestOfMonth_.
// Future months pick it up automatically the first time they're built,
// since ensureDaySheet() seeds standing tasks into every brand-new tab.
function resyncStandingTasksForRestOfMonth_() {
  var now = new Date();
  var monthStart = getMonthStart(now);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  var total = daysInMonth(monthStart);
  var todayDay = now.getDate();

  for (var day = todayDay; day <= total; day++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + (day - 1));
    var sheet = ensureDaySheet(ss, d);
    applyStandingTasksToSheet_(sheet, d);
  }
}

// ---- Date / spreadsheet helpers ---------------------------------------

function getMonthStart(date) {
  var d = new Date(date);
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function daysInMonth(monthStart) {
  return new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();
}

function getAttendanceFolder() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function monthSpreadsheetName(monthStart) {
  return 'Cafe Tan90 Attendance — ' + Utilities.formatDate(monthStart, TIMEZONE, 'MMMM yyyy');
}

function dayTabName(date) {
  return Utilities.formatDate(date, TIMEZONE, 'EEE M/d');
}

// Automatically creates this month's spreadsheet the first time it's needed —
// nobody has to do anything, it just appears the first time the Staff page is
// opened that month, with every day of the month already built in as its own
// tab.
function getOrCreateMonthSpreadsheet(monthStart) {
  var name = monthSpreadsheetName(monthStart);
  var folder = getAttendanceFolder();
  var files = folder.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  var ss = SpreadsheetApp.create(name);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  var root = DriveApp.getRootFolder();
  try { root.removeFile(file); } catch (err) { /* already out of root on some accounts */ }

  var total = daysInMonth(monthStart);
  for (var i = 0; i < total; i++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + i);
    ensureDaySheet(ss, d);
  }
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  return ss;
}

// Table 1's Present Today / Login Edited / Logout Edited columns are plain
// booleans - switching them to Sheets' native checkboxes just changes how
// they're rendered (an actual checkbox instead of the text "TRUE"/"FALSE"),
// not the underlying value, so it's safe to call repeatedly on a sheet
// that's already had real attendance written into it. Idempotent and cheap
// (8 rows x 3 cols), so ensureDaySheet calls this every time it's asked for
// a day's tab - including already-built tabs from before this existed -
// rather than needing a separate one-time migration step.
//
// Only rows that have ever had real content get a checkbox - a spare,
// never-used employee slot (blank name, no log in/out) is left fully blank
// instead of showing an empty checkbox nobody can use. This is judged per
// row from the row's own Name/Log In/Log Out cells, not from whether that
// slot is a *currently* active employee - an employee removed mid-shift
// still has real log-in/out data sitting in their row even after their name
// goes blank (see SETUP.md), and that row must keep its checkboxes so that
// data stays visible, not get silently cleared just because the name is gone.
function applyAttendanceCheckboxes_(sheet) {
  var rows = sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 3).getValues(); // name, log in, log out
  for (var i = 0; i < MAX_EMPLOYEES; i++) {
    var hasContent = !!(String(rows[i][0] || '').trim() || rows[i][1] || rows[i][2]);
    var range = sheet.getRange(TABLE1_START_ROW + i, 5, 1, 3);
    if (hasContent) {
      range.insertCheckboxes();
    } else {
      range.removeCheckboxes();
      range.clearContent();
    }
  }
}

// Same idea as applyAttendanceCheckboxes_, for Table 2's per-employee
// column pair. A slot's "Assigned" / "Done at" sub-header labels and
// Assigned checkboxes show if that employee slot is currently active
// (header non-blank) OR - same historical-preservation reasoning as above -
// if it's inactive but still holds real assigned/done data from before that
// employee was removed. A slot that's simply never been used (inactive, and
// never assigned anything) has both the sub-headers and the column left
// fully blank, instead of showing "Assigned" / "Done at" labels over empty,
// unusable columns.
function applyTaskAssignedCheckboxes_(sheet) {
  var rowCount = getTaskRowCount_(sheet);
  var headers = sheet.getRange(TABLE2_HEADER_ROW, TABLE2_FIRST_EMP_COL, 1, MAX_EMPLOYEES * 2).getValues()[0];
  var grid = rowCount ? sheet.getRange(TABLE2_TASKS_START_ROW, TABLE2_FIRST_EMP_COL, rowCount, MAX_EMPLOYEES * 2).getValues() : [];
  for (var i = 0; i < MAX_EMPLOYEES; i++) {
    var hasHeader = !!String(headers[i * 2] || '').trim();
    var hasContent = hasHeader;
    if (!hasContent) {
      for (var r = 0; r < rowCount; r++) {
        if (grid[r][i * 2] || grid[r][i * 2 + 1]) { hasContent = true; break; }
      }
    }

    var subheaderRange = sheet.getRange(TABLE2_SUBHEADER_ROW, taskAssignedCol_(i), 1, 2);
    if (hasContent) {
      subheaderRange.setValues([['Assigned', 'Done at']]).setFontWeight('bold').setFontColor('#5C6370');
    } else {
      subheaderRange.clearContent();
    }

    if (rowCount) {
      var range = sheet.getRange(TABLE2_TASKS_START_ROW, taskAssignedCol_(i), rowCount, 1);
      if (hasContent) {
        range.insertCheckboxes();
      } else {
        range.removeCheckboxes();
        range.clearContent();
      }
    }
  }
}

function ensureDaySheet(ss, date) {
  var name = dayTabName(date);
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    applyAttendanceCheckboxes_(sheet);
    applyTaskAssignedCheckboxes_(sheet);
    return sheet;
  }

  sheet = ss.insertSheet(name);
  var title = 'Attendance — ' + Utilities.formatDate(date, TIMEZONE, 'EEEE, MMMM d, yyyy');
  sheet.getRange('A1').setValue(title).setFontWeight('bold').setFontSize(13).setFontColor('#1F3864');

  var slots = getAllEmployeeSlots_();
  var names = slots.map(function (s) { return s.active ? s.name : ''; });

  // Table 1: attendance log + Present Today (moved here from the old Table
  // 2, since Table 2 no longer has employee rows to hold it — Present is a
  // per-employee attribute, so it belongs alongside the rest of Table 1).
  sheet.getRange(TABLE1_HEADER_ROW, 1, 1, TABLE1_COLS).setValues([['Name', 'Log In Time', 'Log Out Time', 'Hours Worked', 'Present Today', 'Login Edited', 'Logout Edited']]).setFontWeight('bold');
  sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names.map(function (n) { return [n]; }));
  sheet.getRange(TABLE1_START_ROW, 4, MAX_EMPLOYEES, 1).setFormulaR1C1('=IF(AND(RC[-2]<>"",RC[-1]<>""),(RC[-1]-RC[-2])*24,"")');
  sheet.getRange(TABLE1_START_ROW, 2, MAX_EMPLOYEES, 2).setNumberFormat('hh:mm:ss AM/PM');
  sheet.getRange(TABLE1_START_ROW, 4, MAX_EMPLOYEES, 1).setNumberFormat('0.00" hrs"');
  sheet.getRange(TABLE1_START_ROW, 5, MAX_EMPLOYEES, 3).setValue(false); // Present Today, Login Edited, Logout Edited
  applyAttendanceCheckboxes_(sheet);

  // Table 2: tasks, one row per task (unlimited). Header row: task-name
  // label + each employee's name merged across a pair of columns.
  // Sub-header row: "Assigned" / "Done at" under each half of that pair.
  sheet.getRange(TABLE2_HEADER_ROW, TABLE2_NAME_COL, 2, 1).merge().setValue('Task').setFontWeight('bold').setVerticalAlignment('middle');
  slots.forEach(function (slot, i) {
    var col = taskAssignedCol_(i);
    sheet.getRange(TABLE2_HEADER_ROW, col, 1, 2).merge().setValue(slot.active ? slot.name : '').setFontWeight('bold');
    // Only label an active slot's columns "Assigned" / "Done at" here - a
    // spare, unused slot is left blank rather than showing those labels
    // over columns nobody can fill in. applyTaskAssignedCheckboxes_ is what
    // keeps this correct on every later visit too (a slot that becomes
    // active or inactive after this day tab was already built).
    if (slot.active) {
      sheet.getRange(TABLE2_SUBHEADER_ROW, col).setValue('Assigned').setFontWeight('bold').setFontColor('#5C6370');
      sheet.getRange(TABLE2_SUBHEADER_ROW, col + 1).setValue('Done at').setFontWeight('bold').setFontColor('#5C6370');
    }
  });
  sheet.setColumnWidth(TABLE2_NAME_COL, 160);
  for (var c = 0; c < MAX_EMPLOYEES; c++) {
    sheet.setColumnWidth(taskAssignedCol_(c), 90);
    sheet.setColumnWidth(taskDoneCol_(c), 110);
  }

  // Seed any standing (recurring) tasks that apply to this day of the week —
  // this is what makes "Make base milk every weekday" show up on a brand-new
  // day tab with zero admin action.
  applyStandingTasksToSheet_(sheet, date);

  return sheet;
}

function getTodaySheet() {
  var now = new Date();
  var monthStart = getMonthStart(now);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  return ensureDaySheet(ss, now);
}

function getSheetForDate(date) {
  var monthStart = getMonthStart(date);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  return ensureDaySheet(ss, date);
}

// After an employee is added, renamed, or removed, refresh the Name columns
// on TODAY's tab and every remaining tab this month, so the roster shown
// going forward matches reality — including each employee's header on
// Table 2, since tasks are now laid out with employees as columns there too.
// Tabs that already happened (earlier this month) are deliberately left
// untouched — history stays historically accurate.
function syncEmployeeNamesForRestOfMonth_() {
  var now = new Date();
  var monthStart = getMonthStart(now);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  var slots = getAllEmployeeSlots_();
  var names = slots.map(function (s) { return s.active ? s.name : ''; });
  var total = daysInMonth(monthStart);
  var todayDay = now.getDate();

  for (var day = todayDay; day <= total; day++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + (day - 1));
    var sheet = ensureDaySheet(ss, d);
    sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names.map(function (n) { return [n]; }));
    slots.forEach(function (slot, i) {
      sheet.getRange(TABLE2_HEADER_ROW, taskAssignedCol_(i)).setValue(slot.active ? slot.name : '');
    });
  }
}

// ---- Task-row helpers (Table 2) ----------------------------------------

// How many task rows currently exist on this day tab, by scanning column A
// from the top of the task list until the first blank cell. Relies on the
// invariant — kept by writeTaskRows_/applyStandingTasksToSheet_/planTask —
// that task rows are always contiguous with no gaps, and anything past the
// last real task is cleared blank.
function getTaskRowCount_(sheet) {
  var names = sheet.getRange(TABLE2_TASKS_START_ROW, TABLE2_NAME_COL, TASK_ROW_SCAN_WINDOW, 1).getValues();
  var count = 0;
  for (var i = 0; i < names.length; i++) {
    if (!names[i][0]) break;
    count++;
  }
  return count;
}

// Reads the current task rows as parallel arrays: names[r], assigned[r][slotIdx]
// (bool), done[r][slotIdx] (Date or null). One or two range reads total,
// regardless of task count.
function readTaskRows_(sheet) {
  var rowCount = getTaskRowCount_(sheet);
  if (!rowCount) return { count: 0, names: [], assigned: [], done: [] };

  var names = sheet.getRange(TABLE2_TASKS_START_ROW, TABLE2_NAME_COL, rowCount, 1).getValues().map(function (r) { return r[0]; });
  var fullGrid = sheet.getRange(TABLE2_TASKS_START_ROW, TABLE2_FIRST_EMP_COL, rowCount, MAX_EMPLOYEES * 2).getValues();
  var assigned = fullGrid.map(function (row) {
    var out = [];
    for (var i = 0; i < MAX_EMPLOYEES; i++) out.push(!!row[i * 2]);
    return out;
  });
  var done = fullGrid.map(function (row) {
    var out = [];
    for (var i = 0; i < MAX_EMPLOYEES; i++) out.push(row[i * 2 + 1] || null);
    return out;
  });
  return { count: rowCount, names: names, assigned: assigned, done: done };
}

// Replaces the day's whole task list (names + who's assigned) in one go —
// what "Save tasks" calls. Deliberately leaves the "Done at" column alone
// for rows that still exist after the save, so editing who's assigned to a
// task doesn't erase whether it was already marked done today. Rows beyond
// the new (shorter) list are fully cleared — name, assigned, AND done — so a
// stale completion time can never resurface if that row position gets
// reused by a different task later.
function writeTaskRows_(sheet, tasks, slots) {
  var prevCount = getTaskRowCount_(sheet);
  var newCount = tasks.length;

  if (newCount > 0) {
    var nameValues = tasks.map(function (t) { return [t.name]; });
    sheet.getRange(TABLE2_TASKS_START_ROW, TABLE2_NAME_COL, newCount, 1).setValues(nameValues);

    var assignedGrid = tasks.map(function (t) {
      return slots.map(function (slot) { return !!(t.assignments && t.assignments[slot.id]); });
    });
    slots.forEach(function (slot, i) {
      var colValues = assignedGrid.map(function (row) { return [row[i]]; });
      sheet.getRange(TABLE2_TASKS_START_ROW, taskAssignedCol_(i), newCount, 1).setValues(colValues);
    });
  }

  if (prevCount > newCount) {
    var extra = prevCount - newCount;
    sheet.getRange(TABLE2_TASKS_START_ROW + newCount, TABLE2_NAME_COL, extra, TABLE2_NUM_COLS).clearContent();
  }

  applyTaskAssignedCheckboxes_(sheet);
}

// ---- Staff page API -------------------------------------------------------

function getStaffData() {
  var sheet = getTodaySheet();
  var now = new Date();
  var activeEmployees = getActiveEmployeesWithIndex_();

  var attendance = sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, TABLE1_COLS).getValues(); // name, in, out, hours, present
  var taskRows = readTaskRows_(sheet);

  var employees = activeEmployees.map(function (emp) {
    var i = emp.idx;
    var inTime = attendance[i][1];
    var outTime = attendance[i][2];
    var present = !!attendance[i][4];
    var tasks = [];
    for (var t = 0; t < taskRows.count; t++) {
      var doneVal = taskRows.done[t][i];
      tasks.push({
        col: t, // opaque index into this day's task list (identifies a task row, same idea as the old column index)
        name: taskRows.names[t],
        assigned: !!taskRows.assigned[t][i],
        completedAt: doneVal ? Utilities.formatDate(doneVal, TIMEZONE, 'HH:mm:ss') : null
      });
    }
    return {
      id: emp.id,
      name: emp.name,
      loginTime: inTime ? Utilities.formatDate(inTime, TIMEZONE, 'HH:mm:ss') : null,
      logoutTime: outTime ? Utilities.formatDate(outTime, TIMEZONE, 'HH:mm:ss') : null,
      present: present,
      tasks: tasks
    };
  });

  return {
    day: Utilities.formatDate(now, TIMEZONE, 'EEEE'),
    date: Utilities.formatDate(now, TIMEZONE, 'MMMM d, yyyy'),
    employees: employees
  };
}

function logAction(employeeId, action, pin) {
  assertEmployeePin_(employeeId, pin);
  var sheet = getTodaySheet();
  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId && s.active; });
  if (idx === -1) throw new Error('Unknown or inactive employee: ' + employeeId);

  var row = TABLE1_START_ROW + idx;
  var now = new Date();

  if (action === 'in') {
    var existingIn = sheet.getRange(row, 2).getValue();
    if (!existingIn) {
      sheet.getRange(row, 2).setValue(now);
      sheet.getRange(row, 5).setValue(true); // Present Today
    }
  } else if (action === 'out') {
    var hasIn = sheet.getRange(row, 2).getValue();
    var existingOut = sheet.getRange(row, 3).getValue();
    if (hasIn && !existingOut) {
      sheet.getRange(row, 3).setValue(now);
    }
  } else {
    throw new Error('Unknown action: ' + action);
  }

  return getStaffData();
}

// Marks one employee's one task as done "now", logging the time to the
// second. Idempotent (tapping twice doesn't overwrite the first time) and
// guarded server-side: a task can only be marked done if it was assigned to
// that person and they've already logged in today. `col` is the task's
// index within today's task list (see getStaffData) — it now maps to a
// sheet row instead of a sheet column, but the client-facing meaning
// ("which task in the list") is unchanged.
function completeTask(employeeId, col, pin) {
  assertEmployeePin_(employeeId, pin);
  var sheet = getTodaySheet();
  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId && s.active; });
  if (idx === -1) throw new Error('Unknown or inactive employee: ' + employeeId);

  var rowCount = getTaskRowCount_(sheet);
  if (col < 0 || col >= rowCount) throw new Error('Invalid task.');

  var present = sheet.getRange(TABLE1_START_ROW + idx, 5).getValue();
  if (!present) throw new Error('Log in before marking a task done.');

  var taskRow = TABLE2_TASKS_START_ROW + col;
  var assigned = sheet.getRange(taskRow, taskAssignedCol_(idx)).getValue();
  if (!assigned) throw new Error('This task is not assigned to this person today.');

  var doneCell = sheet.getRange(taskRow, taskDoneCol_(idx));
  if (!doneCell.getValue()) {
    doneCell.setValue(new Date());
  }

  return getStaffData();
}

// Lets an employee undo their OWN accidental tap - today only, and only a
// plain clear (blank), never a specific corrected time. That distinction
// matters: this is just reverting your own immediate mistake through the
// normal one-shot mechanism, so it deliberately does NOT set the "edited"
// flag used by admin corrections - only an admin typing in a specific
// replacement time counts as a correction worth flagging. which: 'in'|'out'.
function undoMyPunch(employeeId, pin, dateStr, which) {
  assertEmployeePin_(employeeId, pin);
  var todayIso = ymdFromDate_(new Date());
  if (dateStr !== todayIso) throw new Error('Can only undo today\'s own entry.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId && s.active; });
  if (idx === -1) throw new Error('Unknown or inactive employee: ' + employeeId);

  var sheet = getSheetForDate(parseDateStr_(dateStr));
  var row = TABLE1_START_ROW + idx;

  if (which === 'in') {
    sheet.getRange(row, 2).setValue('');
  } else if (which === 'out') {
    sheet.getRange(row, 3).setValue('');
  } else {
    throw new Error('Unknown field.');
  }
  var stillIn = sheet.getRange(row, 2).getValue();
  sheet.getRange(row, 5).setValue(!!stillIn); // Present Today follows login, same as a normal tap

  return getStaffData();
}

// Same idea as undoMyPunch but for a "Mark done" tap - today only, plain
// clear, no edited flag (task completions don't carry an edited marker at
// all; only attendance times do, since only those feed payroll).
function undoMyTaskDone(employeeId, pin, dateStr, col) {
  assertEmployeePin_(employeeId, pin);
  var todayIso = ymdFromDate_(new Date());
  if (dateStr !== todayIso) throw new Error('Can only undo today\'s own entry.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId && s.active; });
  if (idx === -1) throw new Error('Unknown or inactive employee: ' + employeeId);

  var sheet = getSheetForDate(parseDateStr_(dateStr));
  var rowCount = getTaskRowCount_(sheet);
  if (col < 0 || col >= rowCount) throw new Error('Invalid task.');

  var taskRow = TABLE2_TASKS_START_ROW + col;
  sheet.getRange(taskRow, taskDoneCol_(idx)).setValue('');

  return getStaffData();
}

// ---- Admin API ------------------------------------------------------------

function getAdminPin_() {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  return stored || DEFAULT_ADMIN_PIN;
}

function checkPin(pin) {
  return pin === getAdminPin_();
}

function assertPin_(pin) {
  if (!checkPin(pin)) throw new Error('Incorrect admin PIN.');
}

// Returns a link straight to *today's* tab in this month's spreadsheet, so
// the admin never has to go hunting through Drive or click through 30 tabs.
function getSheetUrl(pin) {
  assertPin_(pin);
  var sheet = getTodaySheet();
  var ss = sheet.getParent();
  return ss.getUrl() + '#gid=' + sheet.getSheetId();
}

// dateStr: optional 'yyyy-MM-dd' — defaults to today. This is what powers
// the "Tasks for [date]" navigator on the admin page: the same editor, just
// pointed at whichever day you're checking or planning.
function getAdminData(pin, dateStr) {
  assertPin_(pin);
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var sheet = getSheetForDate(date);
  var activeEmployees = getActiveEmployeesWithIndex_();
  var taskRows = readTaskRows_(sheet);

  // Match each of this day's tasks against the active standing-task list by
  // name (same case-insensitive match applyStandingTasksToSheet_ uses to
  // seed idempotently), so the admin UI can show a single unified task list
  // with a "Repeats" chip instead of a separate standing-tasks card. A task
  // only counts as "standing" here if it repeats on this date's own
  // weekday — if it's linked to a standing rule that doesn't include today,
  // it's shown as one-time-for-this-day (the recurring copy lives on its
  // own days).
  var weekday = date.getDay();
  var standingByName = {};
  getAllStandingTasks_().forEach(function (t) {
    if (t.active && t.days && t.days.indexOf(weekday) !== -1) {
      standingByName[t.name.trim().toLowerCase()] = t;
    }
  });

  var tasks = [];
  for (var r = 0; r < taskRows.count; r++) {
    var assignments = {};
    var done = {};
    activeEmployees.forEach(function (emp) {
      assignments[emp.id] = !!taskRows.assigned[r][emp.idx];
      var doneVal = taskRows.done[r][emp.idx];
      done[emp.id] = doneVal ? Utilities.formatDate(doneVal, TIMEZONE, 'HH:mm') : null;
    });
    var match = standingByName[String(taskRows.names[r]).trim().toLowerCase()];
    tasks.push({
      name: taskRows.names[r],
      assignments: assignments,
      done: done,
      standing: match ? { id: match.id, days: match.days } : null
    });
  }
  var todayIso = ymdFromDate_(new Date());
  var thisIso = ymdFromDate_(date);
  return {
    date: thisIso,
    dateLabel: Utilities.formatDate(date, TIMEZONE, 'EEEE, MMMM d, yyyy'),
    isToday: thisIso === todayIso,
    tasks: tasks,
    // hasPin included so the Employees section's Set/Not set column stays
    // correct across every admin-data reload (page load, date navigation,
    // Save tasks), not just right after actually changing a PIN - the
    // client's EMPLOYEES array is refreshed from this same payload every
    // time, and previously this omitted hasPin, silently reverting every
    // row to "Not set" on the very next reload even though the PIN was
    // still saved server-side.
    employees: activeEmployees.map(function (e) { return { id: e.id, name: e.name, hasPin: !!e.pin }; })
  };
}

// Lightweight month overview for the Plan-ahead calendar: which dates in
// `monthIso` ('yyyy-MM') already have at least one task, and how many —
// just enough to draw a dot per date without pulling full task detail for
// every day. Building/reading a month you haven't touched yet takes a beat
// (same one-time cost as any first use of a month — see getOrCreateMonthSpreadsheet).
function getTasksOverviewForMonth(pin, monthIso) {
  assertPin_(pin);
  var parts = String(monthIso).split('-');
  var monthStart = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  var total = daysInMonth(monthStart);

  var byDate = {};
  for (var day = 1; day <= total; day++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + (day - 1));
    var sheet = ensureDaySheet(ss, d);
    var count = getTaskRowCount_(sheet);
    if (count > 0) byDate[ymdFromDate_(d)] = count;
  }
  return byDate;
}

// tasks: [{ name: string, assignments: { employeeId: bool, ... } }, ...] — no limit.
// dateStr: optional 'yyyy-MM-dd' — defaults to today.
function saveTasks(pin, dateStr, tasks) {
  assertPin_(pin);
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var sheet = getSheetForDate(date);
  var slots = getAllEmployeeSlots_();

  writeTaskRows_(sheet, tasks, slots);

  return getAdminData(pin, dateStr);
}

// Applies one task (with its assignments) to every date in `dates`, so the
// admin can plan a task across a week or a month in one go instead of
// re-entering it on each day's tab. Each date is handled independently:
// - If that day already has a task with the same name (from a previous plan,
//   or "Copy previous day's tasks", or manual entry), the assignments are
//   merged in (OR'd together) onto the existing row instead of creating a
//   duplicate.
// - Otherwise the task is appended as a new row — there's no fixed number of
//   task rows to run out of, so nothing gets skipped for being "full".
// dates: array of 'yyyy-MM-dd' strings (interpreted in TIMEZONE).
// assignments: { employeeId: bool, ... }
function planTask(pin, name, assignments, dates) {
  assertPin_(pin);
  name = (name || '').trim();
  if (!name) throw new Error('Enter a task name.');
  if (!dates || !dates.length) throw new Error('Pick at least one date.');
  if (dates.length > 62) throw new Error('That\'s a lot of dates at once. Try a smaller range.');

  var slots = getAllEmployeeSlots_();
  var applied = [];

  dates.forEach(function (dateStr) {
    var date = parseDateStr_(dateStr);
    var sheet = getSheetForDate(date);
    var taskRows = readTaskRows_(sheet);

    var rowOffset = -1;
    for (var r = 0; r < taskRows.count; r++) {
      if (String(taskRows.names[r]).trim().toLowerCase() === name.toLowerCase()) { rowOffset = r; break; }
    }

    if (rowOffset === -1) {
      rowOffset = taskRows.count;
      var newRow = TABLE2_TASKS_START_ROW + rowOffset;
      sheet.getRange(newRow, TABLE2_NAME_COL).setValue(name);
      slots.forEach(function (slot, i) {
        sheet.getRange(newRow, taskAssignedCol_(i)).setValue(!!(assignments && assignments[slot.id]));
      });
    } else {
      var existingRow = TABLE2_TASKS_START_ROW + rowOffset;
      slots.forEach(function (slot, i) {
        var wasChecked = !!taskRows.assigned[rowOffset][i];
        var nowChecked = !!(assignments && assignments[slot.id]);
        if (nowChecked && !wasChecked) {
          sheet.getRange(existingRow, taskAssignedCol_(i)).setValue(true);
        }
      });
    }

    applyTaskAssignedCheckboxes_(sheet);
    applied.push(dateStr);
  });

  return { applied: applied, skipped: [] };
}

function parseDateStr_(s) {
  // s = 'yyyy-MM-dd' — build the Date the same way the sheet lookups expect
  // (a local date at midnight), not via new Date(string) which can shift by
  // a day depending on how the browser/server parses the timezone.
  var parts = String(s).split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function ymdFromDate_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

// Reads the task list + assignments from the day before dateStr (default:
// today) — it does NOT write anything. This only stages those tasks into
// the "Tasks for [date]" editor on the client; nothing is saved to dateStr's
// own tab until the admin clicks "Save tasks", same as any other edit made
// there. (Completion times are never part of this — a fresh day always
// starts with nothing marked done, even when the same tasks are reused.)
function copyPreviousDayTasks(pin, dateStr) {
  assertPin_(pin);
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var prev = new Date(date);
  prev.setDate(prev.getDate() - 1);
  var prevSheet = getSheetForDate(prev);

  var taskRows = readTaskRows_(prevSheet);
  var activeEmployees = getActiveEmployeesWithIndex_();

  var tasks = [];
  for (var r = 0; r < taskRows.count; r++) {
    var assignments = {};
    activeEmployees.forEach(function (emp) { assignments[emp.id] = !!taskRows.assigned[r][emp.idx]; });
    tasks.push({ name: taskRows.names[r], assignments: assignments, standing: null });
  }
  return { tasks: tasks };
}

// ---- Admin: employee roster management -------------------------------------

function listEmployees(pin) {
  assertPin_(pin);
  return getActiveEmployeesWithIndex_().map(function (e) { return { id: e.id, name: e.name, hasPin: !!e.pin }; });
}

function addEmployee(pin, name, employeePin) {
  assertPin_(pin);
  name = (name || '').trim();
  if (!name) throw new Error('Enter a name.');
  employeePin = (employeePin || '').trim();
  if (employeePin && employeePin.length < 4) throw new Error('PIN should be at least 4 characters.');

  var slots = getAllEmployeeSlots_();
  var freeIdx = slots.findIndex(function (s) { return !s.active; });
  if (freeIdx === -1) throw new Error('Max ' + MAX_EMPLOYEES + ' employees. Remove someone first.');

  slots[freeIdx] = { id: slots[freeIdx].id, name: name, active: true, pin: employeePin };
  saveEmployeeSlots_(slots);
  syncEmployeeNamesForRestOfMonth_();

  return { id: slots[freeIdx].id, name: name, employees: listEmployees(pin) };
}

// Admin sets or resets an employee's personal PIN (used both to give a new
// hire their first PIN and to reset one someone's forgotten). '' clears it,
// which locks that employee out of self-service until a PIN is set again —
// the admin PIN still works for them via assertEmployeePin_'s master override
// either way, so removing a PIN never locks admin out of helping them.
function setEmployeePin(pin, employeeId, newPin) {
  assertPin_(pin);
  newPin = (newPin || '').trim();
  if (newPin && newPin.length < 4) throw new Error('PIN should be at least 4 characters.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId; });
  if (idx === -1 || !slots[idx].active) throw new Error('Unknown employee.');

  slots[idx].pin = newPin;
  saveEmployeeSlots_(slots);
  return listEmployees(pin);
}

// Self-service PIN change: currentPin can be the employee's own PIN OR the
// admin PIN (so admin can also walk someone through setting a fresh PIN from
// this same flow without needing the separate Employees-section control).
function changeMyPin(employeeId, currentPin, newPin) {
  assertEmployeePin_(employeeId, currentPin);
  newPin = (newPin || '').trim();
  if (!newPin || newPin.length < 4) throw new Error('PIN should be at least 4 characters.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId; });
  if (idx === -1 || !slots[idx].active) throw new Error('Unknown employee.');

  slots[idx].pin = newPin;
  saveEmployeeSlots_(slots);
  return { ok: true };
}

function renameEmployee(pin, id, newName) {
  assertPin_(pin);
  newName = (newName || '').trim();
  if (!newName) throw new Error('Enter a name.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === id; });
  if (idx === -1 || !slots[idx].active) throw new Error('Unknown employee.');

  slots[idx].name = newName;
  saveEmployeeSlots_(slots);
  syncEmployeeNamesForRestOfMonth_();

  return { employees: listEmployees(pin) };
}

function removeEmployee(pin, id) {
  assertPin_(pin);
  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === id; });
  if (idx === -1) throw new Error('Unknown employee.');

  slots[idx].active = false;
  saveEmployeeSlots_(slots);
  syncEmployeeNamesForRestOfMonth_();

  return { id: id, employees: listEmployees(pin) };
}

// ---- Admin: standing (recurring) tasks -------------------------------------

function listStandingTasks(pin) {
  assertPin_(pin);
  return getAllStandingTasks_();
}

// task: { id?: string, name: string, days: [0-6, ...], assignments: { employeeId: bool } }
// Omit id to create a new one; pass an existing id to edit it in place.
function saveStandingTask(pin, task) {
  assertPin_(pin);
  var name = ((task && task.name) || '').trim();
  if (!name) throw new Error('Enter a task name.');
  var days = (task && task.days) || [];
  if (!days.length) throw new Error('Pick at least one day of the week.');

  var list = getAllStandingTasks_();
  if (task.id) {
    var idx = list.findIndex(function (t) { return t.id === task.id; });
    if (idx === -1) throw new Error('Unknown standing task.');
    list[idx] = { id: task.id, name: name, days: days, assignments: task.assignments || {}, active: true };
  } else {
    list.push({ id: Utilities.getUuid(), name: name, days: days, assignments: task.assignments || {}, active: true });
  }
  saveStandingTasksList_(list);
  resyncStandingTasksForRestOfMonth_();

  return getAllStandingTasks_();
}

// Stops a standing task from being seeded into brand-new days going
// forward. Days that already have it — including future days already
// pre-built this month — are left exactly as they are, same as removing an
// employee doesn't erase their past attendance. Remove it from an individual
// day via that day's own task editor if you need it gone from a specific day.
function removeStandingTask(pin, id) {
  assertPin_(pin);
  var list = getAllStandingTasks_().filter(function (t) { return t.id !== id; });
  saveStandingTasksList_(list);
  return list;
}

// ---- Admin: Payroll -----------------------------------------------------
//
// Pay is computed from the exact log-in/log-out duration (down to the
// second — never rounded to a quarter hour) times an hourly rate. Each
// employee's rate is stored as a small history — { effectiveFrom: 'yyyy-MM-dd',
// rate } entries — rather than a single number, so a raise can be scheduled
// for a specific date (today, a future date, or backdated) without
// silently rewriting what past days were "worth". Pay is totaled by
// employee SLOT, not by name text, so a mid-month rename or removal never
// splits or drops someone's earnings — see readTaskRows_-style slot
// reasoning used elsewhere in this file for the same idea applied to tasks.

function getAllWageRates_() {
  var raw = PropertiesService.getScriptProperties().getProperty(WAGE_RATES_PROP_KEY);
  if (!raw) return {};
  try {
    var parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (err) {
    return {};
  }
}

function saveWageRates_(map) {
  PropertiesService.getScriptProperties().setProperty(WAGE_RATES_PROP_KEY, JSON.stringify(map));
}

// The rate in effect for a given employee on a given date: the entry with
// the latest effectiveFrom that is on or before dateStr, falling back to
// DEFAULT_WAGE_RATE if nothing qualifies yet (e.g. before their first rate
// was ever set).
function getRateForEmployeeOnDate_(employeeId, dateStr) {
  var history = getAllWageRates_()[employeeId] || [];
  var applicable = DEFAULT_WAGE_RATE;
  var best = null;
  history.forEach(function (entry) {
    if (entry.effectiveFrom <= dateStr && (!best || entry.effectiveFrom > best.effectiveFrom)) best = entry;
  });
  if (best) applicable = best.rate;
  return applicable;
}

function round2_(n) { return Math.round(n * 100) / 100; }

function hoursBetween_(inTime, outTime) {
  if (!inTime || !outTime) return 0;
  var ms = outTime.getTime() - inTime.getTime();
  return ms > 0 ? ms / 3600000 : 0;
}

function formatHoursMins_(hoursDecimal) {
  var totalMin = Math.round(hoursDecimal * 60);
  var hh = Math.floor(totalMin / 60);
  var mm = totalMin % 60;
  return hh + 'h ' + (mm < 10 ? '0' + mm : String(mm)) + 'm';
}

// One row per employee SLOT (not just active ones) for a single day tab —
// hours worked, the rate that applied that day, and pay. Blank name = that
// slot wasn't in use that day. Callers filter/label as appropriate.
function computePayrollDay_(sheet, dateStr, slots) {
  var namesCol = sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 1).getValues();
  var inCol = sheet.getRange(TABLE1_START_ROW, 2, MAX_EMPLOYEES, 1).getValues();
  var outCol = sheet.getRange(TABLE1_START_ROW, 3, MAX_EMPLOYEES, 1).getValues();
  var editedCol = sheet.getRange(TABLE1_START_ROW, 6, MAX_EMPLOYEES, 2).getValues(); // [Login Edited, Logout Edited]
  return slots.map(function (s, i) {
    var name = namesCol[i][0] || '';
    var inTime = inCol[i][0] || null;
    var outTime = outCol[i][0] || null;
    var hours = hoursBetween_(inTime, outTime);
    var rate = getRateForEmployeeOnDate_(s.id, dateStr);
    return {
      id: s.id, name: name, hours: hours, rate: rate, pay: hours * rate,
      inTime: inTime, outTime: outTime,
      loginEdited: !!editedCol[i][0], logoutEdited: !!editedCol[i][1]
    };
  });
}

// dateStr: optional 'yyyy-MM-dd' — defaults to today. Shows every currently
// active employee (even with 0 hours so far) plus anyone inactive who still
// has a name on this specific day's tab (so a person removed today still
// shows their pay for today).
function getPayrollDay(pin, dateStr) {
  assertPin_(pin);
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var sheet = getSheetForDate(date);
  var slots = getAllEmployeeSlots_();
  var thisIso = ymdFromDate_(date);
  var dayRows = computePayrollDay_(sheet, thisIso, slots);

  var rows = [];
  dayRows.forEach(function (r, i) {
    if (!slots[i].active && !r.name) return; // slot never used, nothing to show
    rows.push({
      id: r.id,
      name: r.name || slots[i].name || '(removed)',
      hours: round2_(r.hours),
      hoursLabel: formatHoursMins_(r.hours),
      rate: r.rate,
      pay: round2_(r.pay),
      loginTime: r.inTime ? Utilities.formatDate(r.inTime, TIMEZONE, 'HH:mm') : '',
      logoutTime: r.outTime ? Utilities.formatDate(r.outTime, TIMEZONE, 'HH:mm') : '',
      loginEdited: r.loginEdited,
      logoutEdited: r.logoutEdited
    });
  });

  var todayIso = ymdFromDate_(new Date());
  return {
    date: thisIso,
    dateLabel: Utilities.formatDate(date, TIMEZONE, 'EEEE, MMMM d, yyyy'),
    isToday: thisIso === todayIso,
    rows: rows
  };
}

// monthIso: 'yyyy-MM'. Totals hours/pay per employee SLOT across every day
// in that month (days beyond today, if this is the current month, simply
// contribute 0 since nobody's logged in yet — no special-casing needed).
// "Month so far" and "end of month" are the same call: the number is
// whatever's actually on the days that have happened, and it stops growing
// once the month is over.
function getPayrollForMonth(pin, monthIso) {
  assertPin_(pin);
  var parts = String(monthIso).split('-');
  var monthStart = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  var total = daysInMonth(monthStart);
  var slots = getAllEmployeeSlots_();

  var totals = slots.map(function (s) { return { id: s.id, name: s.active ? s.name : '', hours: 0, pay: 0 }; });

  for (var day = 1; day <= total; day++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + (day - 1));
    var dateStr = ymdFromDate_(d);
    var sheet = ensureDaySheet(ss, d);
    var dayRows = computePayrollDay_(sheet, dateStr, slots);
    dayRows.forEach(function (r, i) {
      if (r.name) totals[i].name = r.name; // most recent name seen this month for this slot, so a mid-month rename doesn't split the total
      totals[i].hours += r.hours;
      totals[i].pay += r.pay;
    });
  }

  return totals
    .filter(function (t) { return t.name; })
    .map(function (t) { return { id: t.id, name: t.name, hours: round2_(t.hours), hoursLabel: formatHoursMins_(t.hours), pay: round2_(t.pay) }; });
}

// ---- Employee self-service: personal payroll view --------------------------
// Same shape/computation as getPayrollDay / getPayrollForMonth, just
// authenticated with the employee's own PIN (or the admin PIN) instead of
// the admin PIN alone, and filtered down to that one employee's row/total -
// this is what powers each employee's private "my hours & pay" view, which
// per design nobody else (besides admin) can see.

function getMyPayrollDay(employeeId, pin, dateStr) {
  var slot = assertEmployeePin_(employeeId, pin);
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var sheet = getSheetForDate(date);
  var slots = getAllEmployeeSlots_();
  var thisIso = ymdFromDate_(date);
  var dayRows = computePayrollDay_(sheet, thisIso, slots);
  var idx = slots.findIndex(function (s) { return s.id === employeeId; });
  var r = dayRows[idx];

  var todayIso = ymdFromDate_(new Date());
  return {
    date: thisIso,
    dateLabel: Utilities.formatDate(date, TIMEZONE, 'EEEE, MMMM d, yyyy'),
    isToday: thisIso === todayIso,
    rows: [{
      id: slot.id,
      name: r.name || slot.name,
      hours: round2_(r.hours),
      hoursLabel: formatHoursMins_(r.hours),
      rate: r.rate,
      pay: round2_(r.pay),
      loginTime: r.inTime ? Utilities.formatDate(r.inTime, TIMEZONE, 'HH:mm') : '',
      logoutTime: r.outTime ? Utilities.formatDate(r.outTime, TIMEZONE, 'HH:mm') : '',
      loginEdited: r.loginEdited,
      logoutEdited: r.logoutEdited
    }]
  };
}

function getMyPayrollForMonth(employeeId, pin, monthIso) {
  var slot = assertEmployeePin_(employeeId, pin);
  var parts = String(monthIso).split('-');
  var monthStart = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  var total = daysInMonth(monthStart);
  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId; });

  var hours = 0, pay = 0;
  for (var day = 1; day <= total; day++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + (day - 1));
    var dateStr = ymdFromDate_(d);
    var sheet = ensureDaySheet(ss, d);
    var r = computePayrollDay_(sheet, dateStr, slots)[idx];
    hours += r.hours;
    pay += r.pay;
  }

  return [{ id: slot.id, name: slot.name, hours: round2_(hours), hoursLabel: formatHoursMins_(hours), pay: round2_(pay) }];
}

// Current rate (as of today) + soonest upcoming scheduled change, if any,
// for every active employee — what the rate-editor card shows.
function getWageRatesInfo(pin) {
  assertPin_(pin);
  var employees = getActiveEmployeesWithIndex_();
  var todayIso = ymdFromDate_(new Date());
  var allRates = getAllWageRates_();
  return employees.map(function (e) {
    var history = (allRates[e.id] || []).slice().sort(function (a, b) { return a.effectiveFrom < b.effectiveFrom ? -1 : 1; });
    var upcoming = history.filter(function (h) { return h.effectiveFrom > todayIso; });
    return {
      id: e.id,
      name: e.name,
      currentRate: getRateForEmployeeOnDate_(e.id, todayIso),
      upcoming: upcoming.length ? upcoming[0] : null
    };
  });
}

// effectiveFromStr: 'yyyy-MM-dd' — can be today, a past date (backdated), or
// a future date (scheduled). Re-saving the same effective date replaces
// that entry rather than piling up duplicates.
function setWageRate(pin, employeeId, rate, effectiveFromStr) {
  assertPin_(pin);
  rate = Number(rate);
  if (!(rate >= 0)) throw new Error('Enter a valid hourly rate.');
  if (!effectiveFromStr) throw new Error('Pick an effective date.');

  var map = getAllWageRates_();
  var list = map[employeeId] || [];
  var idx = list.findIndex(function (e) { return e.effectiveFrom === effectiveFromStr; });
  if (idx !== -1) list[idx] = { effectiveFrom: effectiveFromStr, rate: rate };
  else list.push({ effectiveFrom: effectiveFromStr, rate: rate });
  map[employeeId] = list;
  saveWageRates_(map);

  return getWageRatesInfo(pin);
}

// ---- Admin: attendance correction ----------------------------------------
//
// Staff sometimes forget to tap log in/out, or tap the wrong one - and since
// pay is now computed straight from those timestamps, a missed punch isn't
// just a display glitch, it under- or over-pays that day. This lets an admin
// fix the actual log in/out time for any employee on any day (not just
// today), same as editing the Sheet cell by hand would, but validated and
// without needing to know the sheet layout. Editing (or clearing) a field
// only flips its "edited" flag if the value actually changed, so a save that
// touches one field doesn't falsely mark the other as admin-corrected too.
//
// loginTimeStr / logoutTimeStr: 'HH:mm', or '' / falsy to clear that field.
function setAttendanceTimes(pin, employeeId, dateStr, loginTimeStr, logoutTimeStr) {
  assertPin_(pin);
  if (!dateStr) throw new Error('Pick a date.');
  var todayIso = ymdFromDate_(new Date());
  if (dateStr > todayIso) throw new Error('Can\'t set attendance for a future date.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId; });
  if (idx === -1) throw new Error('Unknown employee.');

  var date = parseDateStr_(dateStr);
  function toDateOrNull(timeStr) {
    if (!timeStr) return null;
    var parts = String(timeStr).split(':');
    var h = Number(parts[0]), m = Number(parts[1]);
    if (isNaN(h) || isNaN(m)) throw new Error('Invalid time.');
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), h, m, 0);
  }
  var newLogin = toDateOrNull(loginTimeStr);
  var newLogout = toDateOrNull(logoutTimeStr);
  if (newLogin && newLogout && newLogout.getTime() <= newLogin.getTime()) {
    throw new Error('Log out time must be after log in time.');
  }

  var sheet = getSheetForDate(date);
  var row = TABLE1_START_ROW + idx;

  var existingIn = sheet.getRange(row, 2).getValue();
  var existingOut = sheet.getRange(row, 3).getValue();
  var existingInStr = existingIn ? Utilities.formatDate(existingIn, TIMEZONE, 'HH:mm') : '';
  var existingOutStr = existingOut ? Utilities.formatDate(existingOut, TIMEZONE, 'HH:mm') : '';
  var loginChanged = (loginTimeStr || '') !== existingInStr;
  var logoutChanged = (logoutTimeStr || '') !== existingOutStr;

  sheet.getRange(row, 2).setValue(newLogin || '');
  sheet.getRange(row, 3).setValue(newLogout || '');
  if (loginChanged) sheet.getRange(row, 6).setValue(true);
  if (logoutChanged) sheet.getRange(row, 7).setValue(true);
  // Present Today follows the (possibly just-corrected) log-in, same as a
  // normal tap on the Staff page would - so filling in a missed log-in also
  // makes that day's task-marking work, and clearing one un-marks present.
  sheet.getRange(row, 5).setValue(!!newLogin);

  return getPayrollDay(pin, dateStr);
}

// Admin-only correction for a task's "Done at" time - same idea as
// setAttendanceTimes, but for a task-completion tap instead of a punch.
// Lets admin fix a task that got marked done at the wrong time, or clear one
// that was tapped by accident, for any day up to today. doneTimeStr: 'HH:mm',
// or '' / falsy to clear it. No "edited" flag on tasks (see undoMyTaskDone) -
// this is a lighter-weight correction than attendance since it doesn't feed
// payroll.
function setTaskDoneTime(pin, dateStr, col, employeeId, doneTimeStr) {
  assertPin_(pin);
  if (!dateStr) throw new Error('Pick a date.');
  var todayIso = ymdFromDate_(new Date());
  if (dateStr > todayIso) throw new Error('Can\'t set a task time for a future date.');

  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId; });
  if (idx === -1) throw new Error('Unknown employee.');

  var date = parseDateStr_(dateStr);
  var sheet = getSheetForDate(date);
  var rowCount = getTaskRowCount_(sheet);
  if (col < 0 || col >= rowCount) throw new Error('Invalid task.');

  var newDone = null;
  if (doneTimeStr) {
    var parts = String(doneTimeStr).split(':');
    var h = Number(parts[0]), m = Number(parts[1]);
    if (isNaN(h) || isNaN(m)) throw new Error('Invalid time.');
    newDone = new Date(date.getFullYear(), date.getMonth(), date.getDate(), h, m, 0);
  }

  var taskRow = TABLE2_TASKS_START_ROW + col;
  sheet.getRange(taskRow, taskDoneCol_(idx)).setValue(newDone || '');

  return getAdminData(pin, dateStr);
}

/**
 * Optional one-time helper: run this once from the script editor (select
 * "setAdminPin" in the function dropdown, click Run) to set your own PIN
 * without editing code. Change '2468' below before running, then you can
 * delete the call or leave it — it's harmless to re-run.
 */
function setAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '2468'); // <-- change this
}
