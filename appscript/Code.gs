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
 * Each slot has a fixed row position on every day tab — adding someone just
 * activates a free slot, removing someone deactivates theirs. Nothing is ever
 * deleted from the Sheet, so past attendance history is never touched.
 */

// ---- Configuration ----------------------------------------------------

var TIMEZONE = 'Asia/Bangkok';
var DRIVE_FOLDER_NAME = 'Cafe Tan90 Attendance Sheets';
var MAX_TASK_COLUMNS = 8;  // task columns C..J on the day tab
var MAX_EMPLOYEES = 8;     // employee row-slots per table on the day tab
var EMPLOYEES_PROP_KEY = 'EMPLOYEES_V1';

// Change this once you deploy (or set an Admin PIN via
// Project Settings > Script Properties > ADMIN_PIN so you don't have to edit code).
var DEFAULT_ADMIN_PIN = '2468';

// ---- Sheet row layout (computed from MAX_EMPLOYEES / MAX_TASK_COLUMNS) ----

var TABLE1_HEADER_ROW = 3;
var TABLE1_START_ROW = 4;
var TABLE1_END_ROW = TABLE1_START_ROW + MAX_EMPLOYEES - 1;

var TABLE2_HEADER_ROW = TABLE1_END_ROW + 2; // one blank row between tables
var TABLE2_START_ROW = TABLE2_HEADER_ROW + 1;
var TABLE2_END_ROW = TABLE2_START_ROW + MAX_EMPLOYEES - 1;

var TABLE3_LABEL_ROW = TABLE2_END_ROW + 2;
var TABLE3_HEADER_ROW = TABLE3_LABEL_ROW + 1;
var TABLE3_START_ROW = TABLE3_HEADER_ROW + 1;
var TABLE3_END_ROW = TABLE3_START_ROW + MAX_EMPLOYEES - 1;

// ---- Web app entry point ----------------------------------------------

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page === 'admin') ? 'Admin' : 'Staff';
  var tmpl = HtmlService.createTemplateFromFile(page);
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
    slots.push({ id: 'EMP' + (i + 1), name: defaults[i] || '', active: !!defaults[i] });
  }
  return slots;
}

// Always returns exactly MAX_EMPLOYEES slots, in a fixed order that never
// changes — slot position is what ties a person to their row on the sheet,
// so slots are only ever edited in place (name/active), never reordered.
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
// which physical row it lives on).
function getActiveEmployeesWithIndex_() {
  var slots = getAllEmployeeSlots_();
  var result = [];
  slots.forEach(function (s, i) {
    if (s.active) result.push({ id: s.id, name: s.name, idx: i });
  });
  return result;
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

function ensureDaySheet(ss, date) {
  var name = dayTabName(date);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);
  var title = 'Attendance — ' + Utilities.formatDate(date, TIMEZONE, 'EEEE, MMMM d, yyyy');
  sheet.getRange('A1').setValue(title).setFontWeight('bold').setFontSize(13).setFontColor('#1F3864');

  var slots = getAllEmployeeSlots_();
  var names = slots.map(function (s) { return [s.active ? s.name : '']; });

  // Table 1: attendance log
  sheet.getRange(TABLE1_HEADER_ROW, 1, 1, 4).setValues([['Name', 'Log In Time', 'Log Out Time', 'Hours Worked']]).setFontWeight('bold');
  sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names);
  sheet.getRange(TABLE1_START_ROW, 4, MAX_EMPLOYEES, 1).setFormulaR1C1('=IF(AND(RC[-2]<>"",RC[-1]<>""),(RC[-1]-RC[-2])*24,"")');
  sheet.getRange(TABLE1_START_ROW, 2, MAX_EMPLOYEES, 2).setNumberFormat('hh:mm:ss AM/PM');
  sheet.getRange(TABLE1_START_ROW, 4, MAX_EMPLOYEES, 1).setNumberFormat('0.00" hrs"');

  // Table 2: present-today + task assignments, set by admin
  sheet.getRange(TABLE2_HEADER_ROW, 1, 1, 2).setValues([['Name', 'Present Today']]).setFontWeight('bold');
  sheet.getRange(TABLE2_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names);
  sheet.getRange(TABLE2_START_ROW, 2, MAX_EMPLOYEES, 1).setValue(false);
  sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setFontWeight('bold');

  // Table 3: when each assigned task was actually marked done
  sheet.getRange(TABLE3_LABEL_ROW, 1).setValue('Task completed at').setFontWeight('bold').setFontColor('#1F3864');
  sheet.getRange(TABLE3_HEADER_ROW, 1, 1, 2).setValues([['Name', '']]).setFontWeight('bold');
  sheet.getRange(TABLE3_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setFontWeight('bold');
  sheet.getRange(TABLE3_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names);
  sheet.getRange(TABLE3_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).setNumberFormat('hh:mm:ss AM/PM');

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidths(2, 3, 110);
  for (var c = 3; c <= 2 + MAX_TASK_COLUMNS; c++) sheet.setColumnWidth(c, 130);

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
// going forward matches reality. Tabs that already happened (earlier this
// month) are deliberately left untouched — history stays historically
// accurate.
function syncEmployeeNamesForRestOfMonth_() {
  var now = new Date();
  var monthStart = getMonthStart(now);
  var ss = getOrCreateMonthSpreadsheet(monthStart);
  var slots = getAllEmployeeSlots_();
  var names = slots.map(function (s) { return [s.active ? s.name : '']; });
  var total = daysInMonth(monthStart);
  var todayDay = now.getDate();

  for (var day = todayDay; day <= total; day++) {
    var d = new Date(monthStart);
    d.setDate(d.getDate() + (day - 1));
    var sheet = ensureDaySheet(ss, d);
    sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names);
    sheet.getRange(TABLE2_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names);
    sheet.getRange(TABLE3_START_ROW, 1, MAX_EMPLOYEES, 1).setValues(names);
  }
}

// ---- Staff page API -------------------------------------------------------

function getStaffData() {
  var sheet = getTodaySheet();
  var now = new Date();
  var activeEmployees = getActiveEmployeesWithIndex_();

  var attendance = sheet.getRange(TABLE1_START_ROW, 1, MAX_EMPLOYEES, 3).getValues(); // name, in, out
  var presentCol = sheet.getRange(TABLE2_START_ROW, 2, MAX_EMPLOYEES, 1).getValues(); // present
  var taskHeaders = sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
  var assignedGrid = sheet.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).getValues();
  var completedGrid = sheet.getRange(TABLE3_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).getValues();

  var employees = activeEmployees.map(function (emp) {
    var i = emp.idx;
    var inTime = attendance[i][1];
    var outTime = attendance[i][2];
    var tasks = [];
    for (var c = 0; c < MAX_TASK_COLUMNS; c++) {
      if (!taskHeaders[c]) continue;
      var doneVal = completedGrid[i][c];
      tasks.push({
        col: c,
        name: taskHeaders[c],
        assigned: !!assignedGrid[i][c],
        completedAt: doneVal ? Utilities.formatDate(doneVal, TIMEZONE, 'HH:mm:ss') : null
      });
    }
    return {
      id: emp.id,
      name: emp.name,
      loginTime: inTime ? Utilities.formatDate(inTime, TIMEZONE, 'HH:mm:ss') : null,
      logoutTime: outTime ? Utilities.formatDate(outTime, TIMEZONE, 'HH:mm:ss') : null,
      present: !!presentCol[i][0],
      tasks: tasks
    };
  });

  return {
    day: Utilities.formatDate(now, TIMEZONE, 'EEEE'),
    date: Utilities.formatDate(now, TIMEZONE, 'MMMM d, yyyy'),
    employees: employees
  };
}

function logAction(employeeId, action) {
  var sheet = getTodaySheet();
  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId && s.active; });
  if (idx === -1) throw new Error('Unknown or inactive employee: ' + employeeId);

  var row = TABLE1_START_ROW + idx;
  var presentRow = TABLE2_START_ROW + idx;
  var now = new Date();

  if (action === 'in') {
    var existingIn = sheet.getRange(row, 2).getValue();
    if (!existingIn) {
      sheet.getRange(row, 2).setValue(now);
      sheet.getRange(presentRow, 2).setValue(true);
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
// that person and they've already logged in today.
function completeTask(employeeId, col) {
  var sheet = getTodaySheet();
  var slots = getAllEmployeeSlots_();
  var idx = slots.findIndex(function (s) { return s.id === employeeId && s.active; });
  if (idx === -1) throw new Error('Unknown or inactive employee: ' + employeeId);
  if (col < 0 || col >= MAX_TASK_COLUMNS) throw new Error('Invalid task column.');

  var present = sheet.getRange(TABLE2_START_ROW + idx, 2).getValue();
  if (!present) throw new Error('Log in before marking a task done.');

  var assigned = sheet.getRange(TABLE2_START_ROW + idx, 3 + col).getValue();
  if (!assigned) throw new Error('This task is not assigned to this person today.');

  var doneCell = sheet.getRange(TABLE3_START_ROW + idx, 3 + col);
  if (!doneCell.getValue()) {
    doneCell.setValue(new Date());
  }

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

function getAdminData(pin) {
  assertPin_(pin);
  var sheet = getTodaySheet();
  var activeEmployees = getActiveEmployeesWithIndex_();
  var taskHeaders = sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
  var taskGrid = sheet.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).getValues();

  var tasks = [];
  for (var c = 0; c < MAX_TASK_COLUMNS; c++) {
    if (!taskHeaders[c]) continue;
    var assignments = {};
    activeEmployees.forEach(function (emp) { assignments[emp.id] = !!taskGrid[emp.idx][c]; });
    tasks.push({ name: taskHeaders[c], assignments: assignments });
  }
  return {
    tasks: tasks,
    employees: activeEmployees.map(function (e) { return { id: e.id, name: e.name }; })
  };
}

// tasks: [{ name: string, assignments: { employeeId: bool, ... } }, ...] (max MAX_TASK_COLUMNS)
function saveTasks(pin, tasks) {
  assertPin_(pin);
  if (tasks.length > MAX_TASK_COLUMNS) {
    throw new Error('Max ' + MAX_TASK_COLUMNS + ' tasks per day.');
  }
  var sheet = getTodaySheet();
  var slots = getAllEmployeeSlots_();

  var headerRow = [];
  for (var c = 0; c < MAX_TASK_COLUMNS; c++) {
    headerRow.push(c < tasks.length ? tasks[c].name : '');
  }
  sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]);
  sheet.getRange(TABLE3_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]); // mirror onto the completion table's header too

  var gridValues = [];
  for (var r = 0; r < MAX_EMPLOYEES; r++) {
    var rowVals = [];
    var slotId = slots[r].id;
    for (var c2 = 0; c2 < MAX_TASK_COLUMNS; c2++) {
      var val = false;
      if (c2 < tasks.length && tasks[c2].assignments && tasks[c2].assignments[slotId]) val = true;
      rowVals.push(val);
    }
    gridValues.push(rowVals);
  }
  sheet.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).setValues(gridValues);

  return getAdminData(pin);
}

function copyYesterdayTasks(pin) {
  assertPin_(pin);
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var ySheet = getSheetForDate(yesterday);

  var headerRow = ySheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
  var grid = ySheet.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).getValues();

  var today = getTodaySheet();
  today.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]);
  today.getRange(TABLE3_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]);
  today.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).setValues(grid);
  // Note: completion times are deliberately NOT copied — a fresh day starts
  // with nothing marked done, even if the same tasks are reused.

  return getAdminData(pin);
}

// ---- Admin: employee roster management -------------------------------------

function listEmployees(pin) {
  assertPin_(pin);
  return getActiveEmployeesWithIndex_().map(function (e) { return { id: e.id, name: e.name }; });
}

function addEmployee(pin, name) {
  assertPin_(pin);
  name = (name || '').trim();
  if (!name) throw new Error('Enter a name.');

  var slots = getAllEmployeeSlots_();
  var freeIdx = slots.findIndex(function (s) { return !s.active; });
  if (freeIdx === -1) throw new Error('Max ' + MAX_EMPLOYEES + ' employees — remove someone first.');

  slots[freeIdx] = { id: slots[freeIdx].id, name: name, active: true };
  saveEmployeeSlots_(slots);
  syncEmployeeNamesForRestOfMonth_();

  return { id: slots[freeIdx].id, name: name, employees: listEmployees(pin) };
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

/**
 * Optional one-time helper: run this once from the script editor (select
 * "setAdminPin" in the function dropdown, click Run) to set your own PIN
 * without editing code. Change '2468' below before running, then you can
 * delete the call or leave it — it's harmless to re-run.
 */
function setAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '2468'); // <-- change this
}
