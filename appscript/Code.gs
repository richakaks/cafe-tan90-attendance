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
var STANDING_TASKS_PROP_KEY = 'STANDING_TASKS_V1';
var MAX_STANDING_TASKS = 6; // recurring tasks share the same 8 task columns as one-off/planned tasks

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

// ---- Standing (recurring) tasks -----------------------------------------
// A standing task is a task that repeats automatically on chosen days of the
// week — e.g. "Make base milk" every Mon-Fri — with no admin action needed
// once it's set up. { id, name, days: [0=Sun..6=Sat, ...], assignments:
// { employeeId: bool }, active }. Stored the same way as the employee roster
// (Script Properties), so it survives redeploys.

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

// Idempotent: fills in any standing task that applies to this date's weekday
// and isn't already on the sheet by name. Never overwrites a task that's
// already there (whether it was seeded earlier or hand-edited for this one
// day), so a per-day tweak never gets clobbered by a later resync.
function applyStandingTasksToSheet_(sheet, date) {
  var standing = getActiveStandingTasksForWeekday_(date.getDay());
  if (!standing.length) return;

  var slots = getAllEmployeeSlots_();
  var headerRow = sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
  var changed = false;

  standing.forEach(function (st) {
    var exists = headerRow.some(function (h) { return h && String(h).trim().toLowerCase() === st.name.toLowerCase(); });
    if (exists) return;

    var col = -1;
    for (var c = 0; c < MAX_TASK_COLUMNS; c++) { if (!headerRow[c]) { col = c; break; } }
    if (col === -1) return; // day already has 8 tasks — best-effort, silently skipped

    headerRow[col] = st.name;
    changed = true;
    var colValues = [];
    for (var r = 0; r < MAX_EMPLOYEES; r++) {
      colValues.push([!!(st.assignments && st.assignments[slots[r].id])]);
    }
    sheet.getRange(TABLE2_START_ROW, 3 + col, MAX_EMPLOYEES, 1).setValues(colValues);
  });

  if (changed) {
    sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]);
    sheet.getRange(TABLE3_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]);
  }
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

// dateStr: optional 'yyyy-MM-dd' — defaults to today. This is what powers
// the "Tasks for [date]" navigator on the admin page: the same editor, just
// pointed at whichever day you're checking or planning.
function getAdminData(pin, dateStr) {
  assertPin_(pin);
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var sheet = getSheetForDate(date);
  var activeEmployees = getActiveEmployeesWithIndex_();
  var taskHeaders = sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
  var taskGrid = sheet.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).getValues();

  // Match each of this day's task columns against the active standing-task
  // list by name (same case-insensitive match applyStandingTasksToSheet_
  // uses to seed idempotently), so the admin UI can show a single unified
  // task list with a "Repeats" chip instead of a separate standing-tasks
  // card. A task only counts as "standing" here if it repeats on this
  // date's own weekday — if it's linked to a standing rule that doesn't
  // include today, it's shown as one-time-for-this-day (the recurring copy
  // lives on its own days).
  var weekday = date.getDay();
  var standingByName = {};
  getAllStandingTasks_().forEach(function (t) {
    if (t.active && t.days && t.days.indexOf(weekday) !== -1) {
      standingByName[t.name.trim().toLowerCase()] = t;
    }
  });

  var tasks = [];
  for (var c = 0; c < MAX_TASK_COLUMNS; c++) {
    if (!taskHeaders[c]) continue;
    var assignments = {};
    activeEmployees.forEach(function (emp) { assignments[emp.id] = !!taskGrid[emp.idx][c]; });
    var match = standingByName[String(taskHeaders[c]).trim().toLowerCase()];
    tasks.push({
      name: taskHeaders[c],
      assignments: assignments,
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
    employees: activeEmployees.map(function (e) { return { id: e.id, name: e.name }; })
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
    var headerRow = sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
    var count = headerRow.filter(function (h) { return !!h; }).length;
    if (count > 0) byDate[ymdFromDate_(d)] = count;
  }
  return byDate;
}

// tasks: [{ name: string, assignments: { employeeId: bool, ... } }, ...] (max MAX_TASK_COLUMNS)
// dateStr: optional 'yyyy-MM-dd' — defaults to today.
function saveTasks(pin, dateStr, tasks) {
  assertPin_(pin);
  if (tasks.length > MAX_TASK_COLUMNS) {
    throw new Error('Max ' + MAX_TASK_COLUMNS + ' tasks per day.');
  }
  var date = dateStr ? parseDateStr_(dateStr) : new Date();
  var sheet = getSheetForDate(date);
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

  return getAdminData(pin, dateStr);
}

// Applies one task (with its assignments) to every date in `dates`, so the
// admin can plan a task across a week or a month in one go instead of
// re-entering it on each day's tab. Each date is handled independently:
// - If that day already has a task with the same name (from a previous plan,
//   or "Copy yesterday's tasks", or manual entry), the assignments are
//   merged in (OR'd together) onto the existing column instead of creating a
//   duplicate.
// - Otherwise the task is added to that day's first free task column.
// - A day that already has MAX_TASK_COLUMNS tasks and no matching name is
//   skipped (reported back so the admin can see it happened).
// dates: array of 'yyyy-MM-dd' strings (interpreted in TIMEZONE).
// assignments: { employeeId: bool, ... }
function planTask(pin, name, assignments, dates) {
  assertPin_(pin);
  name = (name || '').trim();
  if (!name) throw new Error('Enter a task name.');
  if (!dates || !dates.length) throw new Error('Pick at least one date.');
  if (dates.length > 62) throw new Error('That\'s a lot of dates at once — try a smaller range.');

  var slots = getAllEmployeeSlots_();
  var applied = [];
  var skipped = [];

  dates.forEach(function (dateStr) {
    var date = parseDateStr_(dateStr);
    var sheet = getSheetForDate(date);

    var headerRow = sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
    var col = -1;
    for (var c = 0; c < MAX_TASK_COLUMNS; c++) {
      if (headerRow[c] && String(headerRow[c]).trim().toLowerCase() === name.toLowerCase()) { col = c; break; }
    }
    if (col === -1) {
      for (var c2 = 0; c2 < MAX_TASK_COLUMNS; c2++) {
        if (!headerRow[c2]) { col = c2; break; }
      }
    }
    if (col === -1) { skipped.push(dateStr); return; }

    headerRow[col] = name;
    sheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]);
    sheet.getRange(TABLE3_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).setValues([headerRow]); // mirror header onto completion table

    var existingCol = sheet.getRange(TABLE2_START_ROW, 3 + col, MAX_EMPLOYEES, 1).getValues();
    var colValues = [];
    for (var r = 0; r < MAX_EMPLOYEES; r++) {
      var slotId = slots[r].id;
      var wasChecked = !!existingCol[r][0];
      var nowChecked = !!(assignments && assignments[slotId]);
      colValues.push([wasChecked || nowChecked]);
    }
    sheet.getRange(TABLE2_START_ROW, 3 + col, MAX_EMPLOYEES, 1).setValues(colValues);

    applied.push(dateStr);
  });

  return { applied: applied, skipped: skipped };
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

  var headerRow = prevSheet.getRange(TABLE2_HEADER_ROW, 3, 1, MAX_TASK_COLUMNS).getValues()[0];
  var grid = prevSheet.getRange(TABLE2_START_ROW, 3, MAX_EMPLOYEES, MAX_TASK_COLUMNS).getValues();
  var activeEmployees = getActiveEmployeesWithIndex_();

  var tasks = [];
  for (var c = 0; c < MAX_TASK_COLUMNS; c++) {
    if (!headerRow[c]) continue;
    var assignments = {};
    activeEmployees.forEach(function (emp) { assignments[emp.id] = !!grid[emp.idx][c]; });
    tasks.push({ name: headerRow[c], assignments: assignments, standing: null });
  }
  return { tasks: tasks };
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
    if (list.filter(function (t) { return t.active; }).length >= MAX_STANDING_TASKS) {
      throw new Error('Max ' + MAX_STANDING_TASKS + ' standing tasks — remove one first.');
    }
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

/**
 * Optional one-time helper: run this once from the script editor (select
 * "setAdminPin" in the function dropdown, click Run) to set your own PIN
 * without editing code. Change '2468' below before running, then you can
 * delete the call or leave it — it's harmless to re-run.
 */
function setAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '2468'); // <-- change this
}
