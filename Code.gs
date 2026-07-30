/**
 * Cross-platform attendance check-in system for Google Sheets.
 *
 * This project is intended to be bound to a Google Sheet. Run
 * initializeAttendanceSystem() once, then deploy as a web app.
 */

const APP = Object.freeze({
  PEOPLE_SHEET: 'People',
  ATTENDANCE_SHEET: 'Attendance',
  SETTINGS_SHEET: 'Settings',
  PEOPLE_HEADERS: [
    'Person ID',
    'Name',
    'Phone',
    'Token',
    'Active',
    'Personal Link',
    'Created At',
  ],
  ATTENDANCE_HEADERS: [
    'Attendance ID',
    'Person ID',
    'Name',
    'Class Date',
    'Check-in Time',
    'Source',
  ],
  SETTINGS_HEADERS: ['Setting', 'Value'],
  DEFAULT_SETTINGS: [
    ['Organization Name', 'Attendance Check-In'],
    ['Timezone', 'America/Chicago'],
    ['Web App URL', ''],
  ],
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance System')
    .addItem('Initialize / repair sheets', 'initializeAttendanceSystem')
    .addSeparator()
    .addItem('Set up new people', 'setupNewPeople')
    .addItem('Refresh personal links', 'refreshPersonalLinks')
    .addSeparator()
    .addItem('Rebuild annual reports', 'rebuildAllAnnualReports')
    .addToUi();
}

function initializeAttendanceSystem() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty(
    'SPREADSHEET_ID',
    spreadsheet.getId()
  );

  const people = ensureSheet_(spreadsheet, APP.PEOPLE_SHEET, APP.PEOPLE_HEADERS);
  const attendance = ensureSheet_(
    spreadsheet,
    APP.ATTENDANCE_SHEET,
    APP.ATTENDANCE_HEADERS
  );
  const settings = ensureSheet_(
    spreadsheet,
    APP.SETTINGS_SHEET,
    APP.SETTINGS_HEADERS
  );

  initializeSettings_(settings);
  formatPeopleSheet_(people);
  formatAttendanceSheet_(attendance);
  formatSettingsSheet_(settings);

  spreadsheet.setSpreadsheetTimeZone(getSetting_('Timezone', 'America/Chicago'));
  SpreadsheetApp.flush();
  toast_('Attendance system is ready.');
}

function setupNewPeople() {
  initializeAttendanceSystem();

  const sheet = getSpreadsheet_().getSheetByName(APP.PEOPLE_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    toast_('Add names to the People sheet first.');
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, APP.PEOPLE_HEADERS.length).getValues();
  let nextNumber = getNextPersonNumber_(values);
  const now = new Date();
  const webAppUrl = getWebAppUrl_();
  let created = 0;

  values.forEach(function (row) {
    const name = String(row[1] || '').trim();
    if (!name || row[0]) return;

    row[0] = 'P' + String(nextNumber++).padStart(4, '0');
    row[1] = name;
    row[2] = normalizePhone_(row[2]);
    row[3] = createToken_();
    row[4] = true;
    row[5] = webAppUrl ? buildPersonalLink_(webAppUrl, row[3]) : '';
    row[6] = now;
    created++;
  });

  sheet.getRange(2, 1, values.length, APP.PEOPLE_HEADERS.length).setValues(values);
  formatPeopleSheet_(sheet);

  if (!created) {
    toast_('No new people needed setup.');
  } else if (!webAppUrl) {
    toast_(
      created +
        ' person(s) set up. Deploy the web app, then use “Refresh personal links.”'
    );
  } else {
    toast_(created + ' person(s) set up with personal links.');
  }
}

function refreshPersonalLinks() {
  const webAppUrl = getWebAppUrl_();
  if (!webAppUrl) {
    throw new Error('Deploy this project as a web app before refreshing links.');
  }

  const sheet = getSpreadsheet_().getSheetByName(APP.PEOPLE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    toast_('There are no people to update.');
    return;
  }

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, APP.PEOPLE_HEADERS.length);
  const values = range.getValues();
  let updated = 0;

  values.forEach(function (row) {
    if (!row[0] || !row[3]) return;
    row[5] = buildPersonalLink_(webAppUrl, row[3]);
    updated++;
  });

  range.setValues(values);
  toast_(updated + ' personal link(s) refreshed.');
}

function doGet(event) {
  const token = String((event && event.parameter && event.parameter.t) || '').trim();
  const person = token ? findPersonByToken_(token) : null;
  const template = HtmlService.createTemplateFromFile('Index');
  const timezone = getSetting_('Timezone', 'America/Chicago');

  template.organizationName = getSetting_(
    'Organization Name',
    'Attendance Check-In'
  );
  template.token = token;
  template.valid = Boolean(person && person.active);
  template.personName = person ? person.name : '';
  template.todayLabel = Utilities.formatDate(
    new Date(),
    timezone,
    'EEEE, MMMM d, yyyy'
  );
  template.errorMessage = person
    ? 'This check-in link is inactive. Please contact an administrator.'
    : 'This check-in link is invalid. Please contact an administrator.';

  return template
    .evaluate()
    .setTitle(template.organizationName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function checkIn(token) {
  token = String(token || '').trim();
  if (!token) return failure_('This check-in link is invalid.');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const person = findPersonByToken_(token);
    if (!person) return failure_('This check-in link is invalid.');
    if (!person.active) return failure_('This check-in link is inactive.');

    const timezone = getSetting_('Timezone', 'America/Chicago');
    const now = new Date();
    const dateKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
    const attendanceSheet = getSpreadsheet_().getSheetByName(APP.ATTENDANCE_SHEET);

    if (hasAttendanceForDate_(attendanceSheet, person.id, dateKey, timezone)) {
      return {
        ok: true,
        duplicate: true,
        message: 'You already checked in today.',
        date: Utilities.formatDate(now, timezone, 'EEEE, MMMM d, yyyy'),
        time: '',
      };
    }

    const classDate = dateAtNoon_(dateKey);
    attendanceSheet.appendRow([
      createAttendanceId_(),
      person.id,
      person.name,
      classDate,
      now,
      'Web App',
    ]);
    formatAttendanceSheet_(attendanceSheet);
    updateYearReport_(dateKey.substring(0, 4));

    return {
      ok: true,
      duplicate: false,
      message: 'You’re checked in!',
      date: Utilities.formatDate(now, timezone, 'EEEE, MMMM d, yyyy'),
      time: Utilities.formatDate(now, timezone, 'h:mm a'),
    };
  } catch (error) {
    console.error(error);
    return failure_('We could not record your check-in. Please try again.');
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function rebuildAllAnnualReports() {
  const spreadsheet = getSpreadsheet_();
  const attendance = spreadsheet.getSheetByName(APP.ATTENDANCE_SHEET);
  const timezone = getSetting_('Timezone', 'America/Chicago');
  const years = {};

  if (attendance && attendance.getLastRow() > 1) {
    attendance
      .getRange(2, 4, attendance.getLastRow() - 1, 1)
      .getValues()
      .forEach(function (row) {
        if (row[0] instanceof Date && !isNaN(row[0])) {
          years[Utilities.formatDate(row[0], timezone, 'yyyy')] = true;
        }
      });
  }

  Object.keys(years)
    .sort()
    .forEach(updateYearReport_);
  toast_(Object.keys(years).length + ' annual report(s) rebuilt.');
}

function updateYearReport_(year) {
  const spreadsheet = getSpreadsheet_();
  const attendance = spreadsheet.getSheetByName(APP.ATTENDANCE_SHEET);
  const timezone = getSetting_('Timezone', 'America/Chicago');
  const grouped = {};

  if (attendance && attendance.getLastRow() > 1) {
    const rows = attendance
      .getRange(2, 1, attendance.getLastRow() - 1, APP.ATTENDANCE_HEADERS.length)
      .getValues();

    rows.forEach(function (row) {
      const classDate = row[3];
      if (!(classDate instanceof Date) || isNaN(classDate)) return;
      const dateKey = Utilities.formatDate(classDate, timezone, 'yyyy-MM-dd');
      if (dateKey.substring(0, 4) !== String(year)) return;

      if (!grouped[dateKey]) grouped[dateKey] = [];
      grouped[dateKey].push({
        name: String(row[2] || ''),
        time: row[4] instanceof Date
          ? Utilities.formatDate(row[4], timezone, 'h:mm a')
          : '',
      });
    });
  }

  let report = spreadsheet.getSheetByName(String(year));
  if (!report) report = spreadsheet.insertSheet(String(year));
  report.clear();

  const output = [['Date / Attendee', 'Check-in Time']];
  const headerRows = [1];
  Object.keys(grouped)
    .sort()
    .forEach(function (dateKey) {
      const displayDate = Utilities.formatDate(
        dateAtNoon_(dateKey),
        timezone,
        'EEEE, MMMM d, yyyy'
      );
      output.push([displayDate, grouped[dateKey].length + ' attendee(s)']);
      headerRows.push(output.length);
      grouped[dateKey]
        .sort(function (a, b) {
          return a.name.localeCompare(b.name);
        })
        .forEach(function (entry) {
          output.push([entry.name, entry.time]);
        });
      output.push(['', '']);
    });

  report.getRange(1, 1, output.length, 2).setValues(output);
  report.setFrozenRows(1);
  report.setColumnWidth(1, 280);
  report.setColumnWidth(2, 140);
  report.getRange(1, 1, output.length, 2).setFontFamily('Arial');
  headerRows.forEach(function (rowNumber) {
    report
      .getRange(rowNumber, 1, 1, 2)
      .setFontWeight('bold')
      .setBackground(rowNumber === 1 ? '#1f4e78' : '#d9eaf7')
      .setFontColor(rowNumber === 1 ? '#ffffff' : '#1f1f1f');
  });
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  headers.forEach(function (header, index) {
    if (!existing[index]) sheet.getRange(1, index + 1).setValue(header);
  });
  styleHeader_(sheet, headers.length);
  return sheet;
}

function initializeSettings_(sheet) {
  const existing = {};
  if (sheet.getLastRow() > 1) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 2)
      .getValues()
      .forEach(function (row) {
        if (row[0]) existing[String(row[0])] = true;
      });
  }

  APP.DEFAULT_SETTINGS.forEach(function (setting) {
    if (!existing[setting[0]]) sheet.appendRow(setting);
  });
}

function getSetting_(key, fallback) {
  const sheet = getSpreadsheet_().getSheetByName(APP.SETTINGS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return fallback;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key && rows[i][1]) return rows[i][1];
  }
  return fallback;
}

function getWebAppUrl_() {
  const configuredUrl = String(getSetting_('Web App URL', '') || '').trim();
  return configuredUrl || ScriptApp.getService().getUrl() || '';
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error(
      'Spreadsheet not configured. Open the bound Sheet and run initializeAttendanceSystem().'
    );
  }
  return active;
}

function findPersonByToken_(token) {
  const sheet = getSpreadsheet_().getSheetByName(APP.PEOPLE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, APP.PEOPLE_HEADERS.length)
    .getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][3]) === token) {
      return {
        id: String(values[i][0]),
        name: String(values[i][1]),
        active: toBoolean_(values[i][4]),
      };
    }
  }
  return null;
}

function hasAttendanceForDate_(sheet, personId, dateKey, timezone) {
  if (!sheet || sheet.getLastRow() < 2) return false;
  const rows = sheet
    .getRange(2, 2, sheet.getLastRow() - 1, 3)
    .getValues();

  return rows.some(function (row) {
    return (
      String(row[0]) === personId &&
      row[2] instanceof Date &&
      Utilities.formatDate(row[2], timezone, 'yyyy-MM-dd') === dateKey
    );
  });
}

function getNextPersonNumber_(values) {
  let maximum = 0;
  values.forEach(function (row) {
    const match = String(row[0] || '').match(/^P(\d+)$/);
    if (match) maximum = Math.max(maximum, Number(match[1]));
  });
  return maximum + 1;
}

function createToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function createAttendanceId_() {
  return 'A-' + Utilities.getUuid();
}

function buildPersonalLink_(webAppUrl, token) {
  return webAppUrl + '?t=' + encodeURIComponent(token);
}

function dateAtNoon_(dateKey) {
  const parts = dateKey.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2], 12, 0, 0);
}

function normalizePhone_(value) {
  return String(value || '').trim();
}

function toBoolean_(value) {
  if (typeof value === 'boolean') return value;
  return ['true', 'yes', '1', 'active'].indexOf(String(value).toLowerCase()) !== -1;
}

function failure_(message) {
  return { ok: false, duplicate: false, message: message, date: '', time: '' };
}

function styleHeader_(sheet, columns) {
  sheet
    .getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontFamily('Arial');
  sheet.setFrozenRows(1);
}

function formatPeopleSheet_(sheet) {
  styleHeader_(sheet, APP.PEOPLE_HEADERS.length);
  sheet.setColumnWidths(1, APP.PEOPLE_HEADERS.length, 140);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(4, 320);
  sheet.setColumnWidth(6, 420);
  sheet.getRange('C:C').setNumberFormat('@');
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('G:G').setNumberFormat('yyyy-mm-dd h:mm AM/PM');
}

function formatAttendanceSheet_(sheet) {
  styleHeader_(sheet, APP.ATTENDANCE_HEADERS.length);
  sheet.setColumnWidths(1, APP.ATTENDANCE_HEADERS.length, 160);
  sheet.setColumnWidth(3, 220);
  sheet.getRange('D:D').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('E:E').setNumberFormat('yyyy-mm-dd h:mm AM/PM');
}

function formatSettingsSheet_(sheet) {
  styleHeader_(sheet, APP.SETTINGS_HEADERS.length);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 300);
}

function toast_(message) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) spreadsheet.toast(message, 'Attendance System', 6);
}
