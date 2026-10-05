/**
 * JOYFUL BIRTH SERVICES — STAFF DASHBOARD BACKEND
 * ------------------------------------------------
 * Separate Apps Script project from the booking backend.
 * Reads/writes the same Bookings sheet, plus two new tabs:
 *   - "Staff"        (login PINs + access scope)
 *   - "CBE Progress" (childbirth education session tracking)
 *
 * SETUP:
 * IMPORTANT — this must live in its OWN Apps Script project, separate
 * from your booking app's script. Do NOT paste this into the same
 * project as code.gs — both files declare things like `doGet` and
 * top-level constants, and Apps Script merges all files in a project
 * into one shared scope, so the two scripts will silently break each
 * other (whichever loads last "wins").
 *
 * 1. Go to script.google.com → New project (a fresh, separate project —
 *    NOT the one opened via your booking sheet's Extensions menu).
 * 2. Paste this whole file in, replacing any starter code.
 * 3. Update DASH_SHEET_ID below to your spreadsheet's ID (the long
 *    string in your sheet's URL — same spreadsheet your booking app
 *    already uses, just a different script project pointing at it).
 * 4. Run `setupDashboardSheets` once from the editor (it will ask for
 *    permission — approve it). This creates the Staff and CBE Progress
 *    tabs and adds a Staff column to Bookings if missing.
 * 5. Deploy → New deployment → Web app → Execute as: Me → Who has
 *    access: Anyone → Deploy. Copy the Web App URL.
 * 6. Paste that URL into dashboard.html where indicated.
 */

const DASH_SHEET_ID = '11fYlJrbIm82zZYV1EtUBMZ7aWo6NbrnFNdeRcVRaRTg'; // same spreadsheet as the booking app

// Calendar that dashboard-created appointments are added to. Kept in sync
// with the booking app's CALENDAR_ID. (Note: the event-creation logic now
// lives in two places — here and the booking app — so if you change how
// events are formatted, update both.)
const DASH_CALENDAR_ID = 'joyfulbirthingca@gmail.com';

const CBE_SESSIONS = [
  { id: 1, name: 'Session 1: 3rd Trimester Prep \u2014 Body & Labor' },
  { id: 2, name: 'Session 2: Labor Toolbox' },
  { id: 3, name: 'Session 3: Medical Interventions & Birth Plan' },
  { id: 4, name: 'Session 4: Breastfeeding & Postpartum Prep' }
];

// Fixed package list — matches the booking app's services. Each package
// lists its trackable components (what shows as a checklist/counter on
// the client's profile) and its price. "components" with type "sessions"
// render as numbered checkboxes (like CBE); type "count" renders as a
// simple tally (e.g. "3 of 6 exercise classes used").
const PACKAGES = [
  {
    id: 'breastfeeding',
    name: 'Breastfeeding Consultation',
    price: 70,
    components: [
      { id: 'consult', label: 'Consultation', type: 'count', target: 1 }
    ]
  },
  {
    id: 'cbe',
    name: 'Childbirth Education (CBE)',
    price: 300,
    components: [
      { id: 'cbe', label: 'CBE Sessions', type: 'sessions', sessions: CBE_SESSIONS }
    ]
  },
  {
    id: 'postpartum',
    name: 'Postpartum Support',
    price: 100,
    components: [
      { id: 'pp', label: 'Postpartum Hours', type: 'count', target: 1 }
    ]
  },
  {
    id: 'pump',
    name: 'Pump Rental',
    price: 70,
    components: [
      { id: 'pump', label: 'Rental Period', type: 'count', target: 1 }
    ]
  },
  {
    id: 'exercise',
    name: 'Prenatal & Postpartum Exercises',
    price: 30,
    components: [
      { id: 'exercise', label: 'Exercise Sessions', type: 'count', target: 1 }
    ]
  },
  {
    id: 'jbm',
    name: 'The Joyful Birth Method',
    price: 850,
    components: [
      { id: 'cbe', label: 'CBE Sessions', type: 'sessions', sessions: CBE_SESSIONS },
      { id: 'doula', label: 'Birth Doula Attendance', type: 'count', target: 1 },
      { id: 'pp_visit', label: 'Postpartum Visit', type: 'count', target: 1 },
      { id: 'exercise_addon', label: 'Exercise Add-on Sessions', type: 'count', target: 6, optional: true }
    ]
  }
];

// ───────────────────────── ONE-TIME SETUP ─────────────────────────

function setupDashboardSheets() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);

  // Staff sheet
  let staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('Staff');
    staffSheet.appendRow(['Name', 'PIN', 'Role', 'Location', 'Active']);
    staffSheet.appendRow(['Joy', '0000', 'owner', 'All', 'TRUE']);
    staffSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  // CBE Progress sheet (legacy — kept for backward compatibility with
  // clients tracked before the full Clients sheet existed)
  let cbeSheet = ss.getSheetByName('CBE Progress');
  if (!cbeSheet) {
    cbeSheet = ss.insertSheet('CBE Progress');
    const headers = ['Client Name', 'Phone', 'Staff', 'Location'].concat(
      CBE_SESSIONS.map(s => s.name)
    ).concat(['Notes']);
    cbeSheet.appendRow(headers);
    cbeSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  // Clients sheet — one row per client, one client profile
  let clientsSheet = ss.getSheetByName('Clients');
  if (!clientsSheet) {
    clientsSheet = ss.insertSheet('Clients');
    clientsSheet.appendRow([
      'Client ID', 'Client Name', 'Phone', 'Email', 'Location', 'Staff',
      'Package ID', 'Package Price (TND)', 'Custom Price (TND)',
      'Progress JSON', 'Status', 'Created At', 'Notes',
      'Assigned Staff', 'Baby Name', 'Baby Birth Date', 'Baby Weight (kg)', 'Baby Birth Details'
    ]);
    clientsSheet.getRange(1, 1, 1, 18).setFontWeight('bold');
  } else {
    // Migration: add any newly-introduced columns to an existing Clients
    // sheet without disturbing the columns/data already there.
    const existingHeaders = clientsSheet.getRange(1, 1, 1, clientsSheet.getLastColumn()).getValues()[0];
    const needed = ['Assigned Staff', 'Baby Name', 'Baby Birth Date', 'Baby Weight (kg)', 'Baby Birth Details'];
    needed.forEach(function (col) {
      if (existingHeaders.indexOf(col) === -1) {
        clientsSheet.getRange(1, clientsSheet.getLastColumn() + 1).setValue(col);
      }
    });
  }

  // Payments sheet — one row per payment received
  let paymentsSheet = ss.getSheetByName('Payments');
  if (!paymentsSheet) {
    paymentsSheet = ss.insertSheet('Payments');
    paymentsSheet.appendRow([
      'Payment ID', 'Client ID', 'Client Name', 'Date', 'Amount (TND)',
      'Method', 'Notes', 'Recorded By'
    ]);
    paymentsSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }

  // Add Staff column to Bookings if not present
  const bookingsSheet = ss.getSheetByName('Bookings');
  if (bookingsSheet) {
    const headerRow = bookingsSheet.getRange(1, 1, 1, bookingsSheet.getLastColumn()).getValues()[0];
    if (headerRow.indexOf('Staff') === -1) {
      bookingsSheet.getRange(1, headerRow.length + 1).setValue('Staff');
    }
  }

  Logger.log('Setup complete. Edit the Staff tab to add real staff members and PINs.');
}

// ───────────────────────── ONE-TIME DIAGNOSTIC ─────────────────────────
// Run this once from the editor (select inspectFicheSheets from the
// function dropdown, click Run) to discover the real tab names and
// column headers of the two fiche response spreadsheets. Check
// View → Logs (or the Execution log panel) after running — copy the
// logged output back to Claude so the import tool can be built against
// the real columns instead of guesses.

const FICHE_BREASTFEEDING_ID = '1-ZpEsi9l-Eg1YwrReFCJYFjeakuNa8jRqhABWVGW6ww';
const FICHE_GENERAL_ID = '1FEC16ZmVkyFTRwFo7YU89EsRYrY-Vw69snl-fPcSjCY';

function inspectFicheSheets() {
  [
    { label: 'BREASTFEEDING FICHE', id: FICHE_BREASTFEEDING_ID },
    { label: 'GENERAL FICHE', id: FICHE_GENERAL_ID }
  ].forEach(function (cfg) {
    Logger.log('───────────────────────────────────────');
    Logger.log(cfg.label + ' (' + cfg.id + ')');
    try {
      const ss = SpreadsheetApp.openById(cfg.id);
      const sheets = ss.getSheets();
      sheets.forEach(function (sheet) {
        const name = sheet.getName();
        const lastCol = sheet.getLastColumn();
        const lastRow = sheet.getLastRow();
        Logger.log('  Tab: "' + name + '" — ' + lastRow + ' rows, ' + lastCol + ' columns');
        if (lastCol > 0 && lastRow > 0) {
          const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
          Logger.log('  Headers: ' + JSON.stringify(headers));
          if (lastRow > 1) {
            const sampleRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
            Logger.log('  Sample row 2: ' + JSON.stringify(sampleRow));
          }
        }
      });
    } catch (err) {
      Logger.log('  ERROR opening this sheet: ' + err.message);
      Logger.log('  (This usually means the script\'s Google account does not have access — open the sheet and share it with the same account this script runs as.)');
    }
  });
  Logger.log('───────────────────────────────────────');
  Logger.log('Done. Copy everything above and send it back so the import tool can be built correctly.');
}

// Combined diagnostic: dumps the EXACT column headers (and one sample row)
// for the Bookings tab + both fiche sheets. Run this from the editor
// (select dumpHeadersForClaude in the function dropdown, click Run, then
// open Execution log) and paste the whole log back if an import ever comes
// back unexpectedly empty.
function dumpHeadersForClaude() {
  const MAIN_ID = '11fYlJrbIm82zZYV1EtUBMZ7aWo6NbrnFNdeRcVRaRTg';
  const targets = [
    { label: 'MAIN SPREADSHEET — Bookings tab', id: MAIN_ID, tab: 'Bookings' },
    { label: 'BREASTFEEDING FICHE', id: '1-ZpEsi9l-Eg1YwrReFCJYFjeakuNa8jRqhABWVGW6ww', tab: null },
    { label: 'GENERAL FICHE', id: '1FEC16ZmVkyFTRwFo7YU89EsRYrY-Vw69snl-fPcSjCY', tab: null }
  ];

  Logger.log('========== HEADER DUMP FOR CLAUDE ==========');
  targets.forEach(function (t) {
    Logger.log('────────────────────────────────────────');
    Logger.log(t.label + '  (' + t.id + ')');
    try {
      const ss = SpreadsheetApp.openById(t.id);
      if (t.tab) {
        Logger.log('  All tabs in this file: ' + JSON.stringify(ss.getSheets().map(function (s) { return s.getName(); })));
      }
      const sheet = t.tab ? ss.getSheetByName(t.tab) : ss.getSheets()[0];
      if (!sheet) {
        Logger.log('  ⚠ Tab "' + t.tab + '" NOT FOUND — check the exact tab name above.');
        return;
      }
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();
      Logger.log('  Reading tab "' + sheet.getName() + '" — ' + lastRow + ' rows, ' + lastCol + ' columns');
      if (lastCol > 0 && lastRow > 0) {
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        Logger.log('  HEADERS: ' + JSON.stringify(headers));
        if (lastRow > 1) {
          Logger.log('  SAMPLE ROW 2: ' + JSON.stringify(sheet.getRange(2, 1, 1, lastCol).getValues()[0]));
        } else {
          Logger.log('  (no data rows yet — only the header row exists)');
        }
      } else {
        Logger.log('  (this sheet is empty)');
      }
    } catch (err) {
      Logger.log('  ERROR opening this sheet: ' + err.message);
    }
  });
  Logger.log('========== END — copy everything above ==========');
}

// ───────────────────────── WEB APP ENTRY ─────────────────────────

function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if (action === 'login') {
      result = handleLogin(e.parameter.pin);
    } else if (action === 'getDashboard') {
      result = getDashboardData(e.parameter.pin);
    } else if (action === 'createAppointment') {
      result = createAppointment(e.parameter);
    } else if (action === 'getCalendarAppointments') {
      result = getCalendarAppointments(e.parameter.pin, e.parameter.month, e.parameter.year);
    } else if (action === 'getMessageQueue') {
      result = getMessageQueue(e.parameter.pin);
    } else if (action === 'getClientBookings') {
      result = getClientBookings(e.parameter.pin, e.parameter.clientId);
    } else if (action === 'getClientProgress') {
      result = getClientProgress(e.parameter.pin, e.parameter.clientName);
    } else if (action === 'updateProgress') {
      result = updateProgress(e.parameter);
    } else if (action === 'getMonthlyStats') {
      result = getMonthlyStats(e.parameter.pin, e.parameter.month, e.parameter.year);
    } else if (action === 'getPackages') {
      result = { success: true, packages: PACKAGES };
    } else if (action === 'getClients') {
      result = getClients(e.parameter.pin);
    } else if (action === 'createClient') {
      result = createClient(e.parameter);
    } else if (action === 'updateClient') {
      result = updateClient(e.parameter);
    } else if (action === 'deleteClient') {
      result = deleteClient(e.parameter);
    } else if (action === 'updateClientComponent') {
      result = updateClientComponent(e.parameter);
    } else if (action === 'addPayment') {
      result = addPayment(e.parameter);
    } else if (action === 'updatePayment') {
      result = updatePayment(e.parameter);
    } else if (action === 'deletePayment') {
      result = deletePayment(e.parameter);
    } else if (action === 'addCharge') {
      result = addCharge(e.parameter);
    } else if (action === 'updateCharge') {
      result = updateCharge(e.parameter);
    } else if (action === 'deleteCharge') {
      result = deleteCharge(e.parameter);
    } else if (action === 'getAppointments') {
      result = getAppointments(e.parameter.pin, e.parameter.clientId);
    } else if (action === 'addAppointment') {
      result = addAppointment(e.parameter);
    } else if (action === 'createGroupClass') {
      result = createGroupClass(e.parameter);
    } else if (action === 'getGroupClasses') {
      result = getGroupClasses(e.parameter.pin);
    } else if (action === 'deleteGroupClass') {
      result = deleteGroupClass(e.parameter);
    } else if (action === 'updateGroupClass') {
      result = updateGroupClass(e.parameter);
    } else if (action === 'bulkAddAppointments') {
      result = bulkAddAppointments(e.parameter);
    } else if (action === 'findDuplicateClients') {
      result = findDuplicateClients(e.parameter.pin);
    } else if (action === 'mergeClients') {
      result = mergeClients(e.parameter);
    } else if (action === 'getPaymentPlan') {
      result = getPaymentPlan(e.parameter.pin, e.parameter.clientId);
    } else if (action === 'addPlanInstallment') {
      result = addPlanInstallment(e.parameter);
    } else if (action === 'markInstallmentPaid') {
      result = markInstallmentPaid(e.parameter);
    } else if (action === 'updatePlanInstallment') {
      result = updatePlanInstallment(e.parameter);
    } else if (action === 'deletePlanInstallment') {
      result = deletePlanInstallment(e.parameter);
    } else if (action === 'updateAppointment') {
      result = updateAppointment(e.parameter);
    } else if (action === 'deleteAppointment') {
      result = deleteAppointment(e.parameter);
    } else if (action === 'addExpense') {
      result = addExpense(e.parameter);
    } else if (action === 'deleteExpense') {
      result = deleteExpense(e.parameter);
    } else if (action === 'getAccountingReport') {
      result = getAccountingReport(e.parameter.pin, e.parameter.periodType, e.parameter.anchor);
    } else if (action === 'getFlexibleReport') {
      result = getFlexibleReport(e.parameter.pin, e.parameter.periodType, e.parameter.anchor, e.parameter.startStr, e.parameter.endStr);
    } else if (action === 'getClientLedger') {
      result = getClientLedger(e.parameter.pin, e.parameter.clientId);
    } else if (action === 'getWeeklyDiversSummary') {
      result = getWeeklyDiversSummary(e.parameter.pin, e.parameter.weekStart, e.parameter.endStr);
    } else if (action === 'generateInvoiceData') {
      let otherServices = [];
      try { otherServices = JSON.parse(e.parameter.otherServices || '[]'); } catch (err) { otherServices = []; }
      let selectedItems = [];
      try { selectedItems = JSON.parse(e.parameter.selectedItems || '[]'); } catch (err) { selectedItems = []; }
      result = generateInvoiceData(e.parameter.pin, e.parameter.periodStart, e.parameter.periodEnd, e.parameter.clientLabel, otherServices, selectedItems);
    } else if (action === 'generateInvoiceNumber') {
      result = generateInvoiceNumber(e.parameter.pin, e.parameter.periodStart, e.parameter.periodEnd, e.parameter.clientLabel, e.parameter.totalTtc);
    } else if (action === 'getImportCandidates') {
      result = getImportCandidates(e.parameter.pin);
    } else if (action === 'lookupFicheByPhone') {
      result = lookupFicheByPhone(e.parameter.pin, e.parameter.phone);
    } else if (action === 'importApprovedClients') {
      let approvedList = [];
      try { approvedList = JSON.parse(e.parameter.approvedList || '[]'); } catch (err) { approvedList = []; }
      result = importApprovedClients(e.parameter.pin, approvedList);
    } else if (action === 'generateClientInvoiceData') {
      result = generateClientInvoiceData(e.parameter.pin, e.parameter.clientId);
    } else if (action === 'generateClientInvoiceNumber') {
      result = generateClientInvoiceNumber(e.parameter.pin, e.parameter.clientId, e.parameter.totalTtc);
    } else if (action === 'generateManualInvoiceData') {
      let services = [];
      try { services = JSON.parse(e.parameter.services || '[]'); } catch (err) { services = []; }
      result = generateManualInvoiceData(e.parameter.pin, e.parameter.clientLabel, e.parameter.objet, services);
    } else if (action === 'getSuggestedInvoiceNumber') {
      result = getSuggestedInvoiceNumber(e.parameter.pin);
    } else if (action === 'recordManualInvoice') {
      result = recordManualInvoice(e.parameter.pin, e.parameter.invoiceNumber, e.parameter.clientLabel, e.parameter.totalTtc, e.parameter.objet);
    } else if (action === 'exportInvoiceToSheet') {
      result = exportInvoiceToSheet(e.parameter);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── AUTH ─────────────────────────

function getStaffByPin(pin) {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  const sheet = ss.getSheetByName('Staff');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[headers.indexOf('PIN')]).trim() === String(pin).trim()) {
      const active = row[headers.indexOf('Active')];
      if (String(active).toUpperCase() !== 'TRUE') continue;
      return {
        name: row[headers.indexOf('Name')],
        role: row[headers.indexOf('Role')],
        location: row[headers.indexOf('Location')]
      };
    }
  }
  return null;
}

function handleLogin(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'PIN not recognized' };
  return { success: true, staff: staff };
}

// ───────────────────────── HELPERS ─────────────────────────

function getBookingsData() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  const sheet = ss.getSheetByName('Bookings');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[headers.indexOf('Booking ID')]) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    data.push(obj);
  }
  return data;
}

function filterByScope(bookings, staff) {
  if (staff.role === 'owner' || staff.location === 'All') return bookings;
  return bookings.filter(b =>
    (b['Staff'] && String(b['Staff']).trim() === staff.name) ||
    (b['Location'] && String(b['Location']).trim() === staff.location)
  );
}

// Some old Bookings rows have a generated client ID (like "C1718..." or
// "DASH-1718...") sitting in the Client Name column instead of a real name.
// This returns a readable label instead of showing the raw code on screen.
function displayName(raw) {
  const s = String(raw || '').trim();
  if (!s) return 'Unknown';
  if (/^(C|DASH-)\d{6,}$/.test(s)) return 'Appointment'; // looks like an internal ID
  return s;
}

function normDate(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  let s = String(val).replace(/^'/, '').trim();
  if (!s) return '';
  // Already yyyy-MM-dd (possibly with a time/T suffix) → keep the date part
  let iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  // dd/MM/yyyy or dd-MM-yyyy (the format the booking app / fiches may use)
  let dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) {
    const dd = ('0' + dmy[1]).slice(-2);
    const mm = ('0' + dmy[2]).slice(-2);
    return dmy[3] + '-' + mm + '-' + dd;
  }
  return s;
}

// Normalizes a Time cell to a plain "HH:MM" string. Sheets sometimes stores
// a bare time as an 1899-12-30 datetime serial, which reads back as a full
// ISO timestamp — this extracts just the clock time from those, and strips
// any leading apostrophe from text-stored times.
function cleanTime(val) {
  if (val === null || val === undefined || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
  }
  let s = String(val).replace(/^'/, '').trim();
  // Mangled ISO like "1899-12-30T09:46:25.000Z" → pull the HH:MM
  const iso = s.match(/T(\d{2}:\d{2})/);
  if (iso) return iso[1];
  // Already-clean "9:46" or "09:46:25" → keep HH:MM
  const hm = s.match(/^(\d{1,2}:\d{2})/);
  if (hm) return hm[1];
  return s;
}

// ───────────────────────── DASHBOARD DATA ─────────────────────────

function getDashboardData(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const upcoming = [];
  const clientsMap = {};

  bookings.forEach(b => {
    const dateStr = normDate(b['Appointment Date']);
    const bDate = dateStr ? new Date(dateStr) : null;
    const clientName = b['Client Name'] || 'Unknown';

    if (bDate && bDate >= today && String(b['Status']).toLowerCase() !== 'cancelled') {
      upcoming.push({
        date: dateStr,
        time: cleanTime(b['Time']),
        clientName: clientName,
        service: b['Service (EN)'],
        location: b['Location'] || '',
        staff: b['Staff'] || '',
        status: b['Status']
      });
    }

    if (!clientsMap[clientName]) {
      clientsMap[clientName] = {
        name: clientName,
        phone: b['Phone'] || '',
        location: b['Location'] || '',
        services: [],
        lastDate: dateStr,
        isCBE: false
      };
    }
    const c = clientsMap[clientName];
    if (b['Service (EN)']) c.services.push(b['Service (EN)']);
    if (dateStr && dateStr > c.lastDate) c.lastDate = dateStr;
    if (String(b['Service (EN)'] || '').toLowerCase().indexOf('childbirth') !== -1) {
      c.isCBE = true;
    }
  });

  upcoming.sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));

  const clients = Object.values(clientsMap).map(c => {
    c.services = Array.from(new Set(c.services));
    return c;
  });

  return {
    success: true,
    staff: staff,
    upcoming: upcoming,
    clients: clients,
    cbeSessions: CBE_SESSIONS
  };
}

// Returns every booked appointment for a client (matched by name) from the
// Bookings sheet — past, today, and upcoming — so their profile can show a
// full appointment history alongside the manual notes log.
function getClientBookings(pin, clientId) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  // Look up the client's name from their profile (bookings are matched by name)
  const cSheet = getClientsSheet();
  const cRows = cSheet.getDataRange().getValues();
  const cHeaders = cRows[0];
  let clientName = '';
  for (let i = 1; i < cRows.length; i++) {
    if (String(cRows[i][cHeaders.indexOf('Client ID')]) === String(clientId)) {
      clientName = String(cRows[i][cHeaders.indexOf('Client Name')] || '').trim();
      break;
    }
  }
  if (!clientName) return { success: true, bookings: [] };

  const bookings = getBookingsData();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const matched = [];
  bookings.forEach(b => {
    const bName = String(b['Client Name'] || '').trim();
    if (bName.toLowerCase() !== clientName.toLowerCase()) return;
    const dateStr = normDate(b['Appointment Date']);
    let when = 'unknown';
    if (dateStr) {
      const d = new Date(dateStr);
      d.setHours(0, 0, 0, 0);
      when = d > today ? 'upcoming' : (d.getTime() === today.getTime() ? 'today' : 'past');
    }
    matched.push({
      date: dateStr,
      time: cleanTime(b['Time']),
      service: b['Service (EN)'] || '',
      location: b['Location'] || '',
      staff: b['Staff'] || '',
      status: b['Status'] || '',
      when: when
    });
  });

  // Most recent first
  matched.sort((a, b) => (b.date + (b.time || '')).localeCompare(a.date + (a.time || '')));

  return { success: true, bookings: matched };
}

// ───────────────────────── CALENDAR FEED ─────────────────────────
// Returns every (non-cancelled) appointment whose date falls in the given
// month, scope-filtered like everything else. Powers the calendar view —
// unlike getDashboardData's "upcoming", this includes past dates in the
// month too, so you can look back as well as ahead.
// ───────────────────────── MESSAGE FLOW QUEUE ─────────────────────────
// Finds appointments that need a WhatsApp touch: ones happening TOMORROW
// (day-before reminder) and ones that happened in the last few days
// (post-visit review/follow-up). Returns enough to build a one-tap wa.me
// message per client. Scope-filtered like everything else.
function getMessageQueue(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today.getTime() + 86400000);
  const followupWindowStart = new Date(today.getTime() - 4 * 86400000); // last 4 days

  const reminders = [];
  const followups = [];

  bookings.forEach(b => {
    if (String(b['Status']).toLowerCase() === 'cancelled') return;
    const dateStr = normDate(b['Appointment Date']);
    if (!dateStr) return;
    const parts = dateStr.split('-').map(Number);
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    d.setHours(0, 0, 0, 0);

    const entry = {
      clientName: displayName(b['Client Name']),
      phone: b['Phone'] || '',
      service: b['Service (EN)'] || '',
      date: dateStr,
      time: cleanTime(b['Time']),
      location: b['Location'] || ''
    };

    if (d.getTime() === tomorrow.getTime()) {
      reminders.push(entry);
    } else if (d.getTime() >= followupWindowStart.getTime() && d.getTime() < today.getTime()) {
      followups.push(entry);
    }
  });

  reminders.sort((a, b) => (a.time || '').localeCompare(b.time || ''));
  followups.sort((a, b) => (b.date + (b.time || '')).localeCompare(a.date + (a.time || '')));

  return { success: true, reminders: reminders, followups: followups };
}

function getCalendarAppointments(pin, month, year) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);

  const m = parseInt(month, 10); // 1-12
  const y = parseInt(year, 10);

  const appointments = [];
  bookings.forEach(b => {
    const dateStr = normDate(b['Appointment Date']);
    if (!dateStr) return;
    // Parse the y/m/d straight from the string (yyyy-MM-dd) rather than via a
    // Date object, whose UTC-vs-local interpretation can shift a booking into
    // the wrong month at month boundaries and drop it from view.
    const parts = dateStr.split('-');
    const by = parseInt(parts[0], 10);
    const bm = parseInt(parts[1], 10);
    if (bm !== m || by !== y) return;
    if (String(b['Status']).toLowerCase() === 'cancelled') return;
    appointments.push({
      date: dateStr,
      time: cleanTime(b['Time']),
      clientName: displayName(b['Client Name']),
      service: b['Service (EN)'] || '',
      location: b['Location'] || '',
      staff: b['Staff'] || '',
      status: b['Status'] || ''
    });
  });

  // Also surface group-class sessions on the calendar. These live in the
  // Group Classes sheet (not Bookings), so read them separately and add each
  // session date that falls in this month.
  try {
    const gSheet = getGroupClassesSheet();
    const gRows = gSheet.getDataRange().getValues();
    if (gRows.length > 1) {
      const gh = gRows[0];
      const nameCol = gh.indexOf('Name');
      const namesCol = gh.indexOf('Client Names (JSON)');
      const datesCol = gh.indexOf('Dates (JSON)');
      const timeCol = gh.indexOf('Time');
      const locCol = gh.indexOf('Location');
      for (let i = 1; i < gRows.length; i++) {
        let dates = [];
        let names = [];
        try { dates = JSON.parse(gRows[i][datesCol] || '[]'); } catch (e) {}
        try { names = JSON.parse(gRows[i][namesCol] || '[]'); } catch (e) {}
        dates.forEach(function (ds) {
          const nd = normDate(ds);
          if (!nd) return;
          const dp = nd.split('-').map(Number);
          if (dp[1] !== m || dp[0] !== y) return;
          appointments.push({
            date: nd,
            time: cleanTime(gRows[i][timeCol]),
            clientName: gRows[i][nameCol] + ' (group · ' + names.length + ')',
            service: 'Group class',
            location: gRows[i][locCol] || '',
            staff: '',
            status: 'Group',
            isGroup: true
          });
        });
      }
    }
  } catch (e) { /* group classes optional */ }

  appointments.sort((a, b) => (a.date + (a.time || '')).localeCompare(b.date + (b.time || '')));

  return { success: true, month: m, year: y, appointments: appointments };
}

// ───────────────────────── CREATE APPOINTMENT (dashboard → Bookings + Calendar) ─────────────────────────
// Writes a new appointment row to the same Bookings sheet the booking app
// uses, aligned by header name (so it can't land in the wrong column even
// if the sheet's column layout changes), and creates a matching Google
// Calendar event. This is the dashboard's own write path — separate from
// the booking app, but targeting the same sheet + calendar.

// ───────────────────── CONSOLIDATE / MERGE CLIENTS ─────────────────────
// Older clients often ended up with several profiles (one per service, or one
// from an import). These find likely duplicates and merge them onto a single
// profile, carrying over EVERY appointment, charge, payment and plan row.
function normName(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function normPhone(s) {
  return String(s || '').replace(/[^\d]/g, '').slice(-8); // last 8 digits (TN local)
}

function findDuplicateClients(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const h = rows[0];
  const idC = h.indexOf('Client ID');
  const nameC = h.indexOf('Client Name');
  const phoneC = h.indexOf('Phone');
  const pkgC = h.indexOf('Package ID');
  const createdC = h.indexOf('Created At');

  const clients = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][idC]) continue;
    clients.push({
      id: String(rows[i][idC]),
      name: String(rows[i][nameC] || ''),
      phone: String(rows[i][phoneC] || ''),
      packageId: String(rows[i][pkgC] || ''),
      createdAt: String(rows[i][createdC] || '')
    });
  }

  // Group by phone when present, otherwise by normalized name.
  const groups = {};
  clients.forEach(function (c) {
    const p = normPhone(c.phone);
    const key = p ? 'p:' + p : 'n:' + normName(c.name);
    if (!key || key === 'n:') return;
    if (!groups[key]) groups[key] = [];
    groups[key].push(c);
  });

  const dupes = [];
  Object.keys(groups).forEach(function (k) {
    if (groups[k].length > 1) {
      // Oldest first — usually the best "primary" to keep
      groups[k].sort(function (a, b) { return String(a.createdAt).localeCompare(String(b.createdAt)); });
      dupes.push({ key: k, matchedOn: k.indexOf('p:') === 0 ? 'phone' : 'name', clients: groups[k] });
    }
  });

  return { success: true, groups: dupes };
}

// Merge one or more duplicate profiles into a primary profile. Every related
// row (appointments, charges, payments, payment plan) is re-pointed to the
// primary's Client ID, blank fields on the primary are filled from the
// duplicates, then the duplicate client rows are removed.
function mergeClients(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (staff.role !== 'owner') return { success: false, error: 'Only the owner can merge clients.' };

  const primaryId = String(params.primaryId || '');
  let mergeIds = [];
  try { mergeIds = JSON.parse(params.mergeIds || '[]').map(String); } catch (e) { mergeIds = []; }
  mergeIds = mergeIds.filter(function (id) { return id && id !== primaryId; });
  if (!primaryId || !mergeIds.length) return { success: false, error: 'Pick a primary profile and at least one to merge into it.' };

  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  const moved = { appointments: 0, charges: 0, payments: 0, plan: 0 };

  // Re-point Client ID on every related sheet
  const repoint = function (sheetName, counterKey) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getDataRange().getValues();
    const hh = data[0];
    const cidx = hh.indexOf('Client ID');
    if (cidx === -1) return;
    for (let i = 1; i < data.length; i++) {
      if (mergeIds.indexOf(String(data[i][cidx])) !== -1) {
        sh.getRange(i + 1, cidx + 1).setValue(primaryId);
        moved[counterKey] += 1;
      }
    }
  };
  repoint('Appointments', 'appointments');
  repoint('Charges', 'charges');
  repoint('Payments', 'payments');
  repoint('Payment Plan', 'plan');

  // Fill any blank fields on the primary from the duplicates, so nothing is lost
  const cSheet = getClientsSheet();
  const cRows = cSheet.getDataRange().getValues();
  const ch = cRows[0];
  const idC = ch.indexOf('Client ID');
  let primaryRow = -1;
  const dupRows = [];
  for (let i = 1; i < cRows.length; i++) {
    const id = String(cRows[i][idC]);
    if (id === primaryId) primaryRow = i;
    else if (mergeIds.indexOf(id) !== -1) dupRows.push(i);
  }
  if (primaryRow === -1) return { success: false, error: 'Primary client not found.' };

  const fillable = ['Phone', 'Email', 'Location', 'Staff', 'Notes', 'Assigned Staff',
                    'Baby Name', 'Baby Birth Date', 'Baby Weight (kg)', 'Baby Birth Details'];
  fillable.forEach(function (col) {
    const c = ch.indexOf(col);
    if (c === -1) return;
    let val = cRows[primaryRow][c];
    if (val !== '' && val !== null && val !== undefined) return; // already set
    for (let d = 0; d < dupRows.length; d++) {
      const dv = cRows[dupRows[d]][c];
      if (dv !== '' && dv !== null && dv !== undefined) {
        cSheet.getRange(primaryRow + 1, c + 1).setValue(dv);
        break;
      }
    }
  });

  // Note on the primary which services the duplicates represented, so their
  // history is visible rather than silently dropped.
  const pkgC = ch.indexOf('Package ID');
  const notesC = ch.indexOf('Notes');
  if (pkgC !== -1 && notesC !== -1) {
    const otherPkgs = dupRows.map(function (r) { return String(cRows[r][pkgC] || ''); })
      .filter(function (p) { return p && p !== String(cRows[primaryRow][pkgC] || ''); });
    if (otherPkgs.length) {
      const existing = String(cSheet.getRange(primaryRow + 1, notesC + 1).getValue() || '');
      const add = 'Merged profiles also had: ' + otherPkgs.join(', ');
      cSheet.getRange(primaryRow + 1, notesC + 1).setValue(existing ? existing + ' | ' + add : add);
    }
  }

  // Delete duplicate rows (bottom-up so indices stay valid)
  dupRows.sort(function (a, b) { return b - a; });
  dupRows.forEach(function (r) { cSheet.deleteRow(r + 1); });

  return { success: true, merged: mergeIds.length, moved: moved };
}

// ─────────────────── PAYMENT PLAN (JBM / doula clients) ───────────────────
// For package clients (esp. Joyful Birth Method) whose work spans months:
// schedule what's due when, mark each installment paid, and see what's been
// paid but NOT yet invoiced. Marking an installment paid writes a real row to
// the Payments ledger, so the client's balance stays the single source of
// truth — the plan only tracks the schedule and the invoice status.
function getPaymentPlanSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  let sheet = ss.getSheetByName('Payment Plan');
  if (!sheet) {
    sheet = ss.insertSheet('Payment Plan');
    sheet.appendRow(['Plan ID', 'Client ID', 'Due Date', 'Amount (TND)', 'Label', 'Status', 'Paid Date', 'Payment ID', 'Invoiced', 'Created At']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  return sheet;
}

function getPaymentPlan(pin, clientId) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getPaymentPlanSheet();
  const rows = sheet.getDataRange().getValues();
  const h = rows[0];
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][h.indexOf('Client ID')]) !== String(clientId)) continue;
    items.push({
      id: String(rows[i][h.indexOf('Plan ID')]),
      dueDate: normDate(rows[i][h.indexOf('Due Date')]),
      amount: parseFloat(rows[i][h.indexOf('Amount (TND)')]) || 0,
      label: String(rows[i][h.indexOf('Label')] || ''),
      status: String(rows[i][h.indexOf('Status')] || 'planned').toLowerCase(),
      paidDate: normDate(rows[i][h.indexOf('Paid Date')]),
      invoiced: String(rows[i][h.indexOf('Invoiced')] || '').toLowerCase() === 'yes'
    });
  }
  items.sort(function (a, b) { return String(a.dueDate).localeCompare(String(b.dueDate)); });

  const planned = items.reduce(function (s, it) { return s + it.amount; }, 0);
  const paid = items.filter(function (it) { return it.status === 'paid'; }).reduce(function (s, it) { return s + it.amount; }, 0);
  // "Ready to invoice" = paid but not yet invoiced
  const toInvoice = items.filter(function (it) { return it.status === 'paid' && !it.invoiced; });
  const toInvoiceTotal = toInvoice.reduce(function (s, it) { return s + it.amount; }, 0);
  const nextDue = items.filter(function (it) { return it.status !== 'paid'; })[0] || null;

  return {
    success: true,
    items: items,
    plannedTotal: Math.round(planned * 1000) / 1000,
    paidTotal: Math.round(paid * 1000) / 1000,
    outstanding: Math.round((planned - paid) * 1000) / 1000,
    toInvoiceCount: toInvoice.length,
    toInvoiceTotal: Math.round(toInvoiceTotal * 1000) / 1000,
    nextDue: nextDue
  };
}

function addPlanInstallment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (!params.clientId) return { success: false, error: 'Missing client.' };
  const amount = Number(params.amount);
  if (!amount || isNaN(amount)) return { success: false, error: 'Enter an amount.' };
  const sheet = getPaymentPlanSheet();
  const id = 'PL' + Date.now();
  sheet.appendRow([
    id,
    params.clientId,
    params.dueDate ? "'" + params.dueDate : '',
    amount,
    params.label || '',
    'planned',
    '',
    '',
    'No',
    new Date().toISOString()
  ]);
  return { success: true, planId: id };
}

// Mark an installment paid: records a real payment in the ledger AND flags the
// plan row, so the balance and the schedule stay in step.
function markInstallmentPaid(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getPaymentPlanSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowNum = findRowByValue(sheet, 'Plan ID', params.planId);
  if (rowNum === -1) return { success: false, error: 'Installment not found.' };

  const amount = Number(sheet.getRange(rowNum, headers.indexOf('Amount (TND)') + 1).getValue()) || 0;
  const label = String(sheet.getRange(rowNum, headers.indexOf('Label') + 1).getValue() || '');
  const paidDate = params.paidDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Record in the real payment ledger
  const pay = addPayment({
    pin: params.pin,
    clientId: params.clientId,
    clientName: params.clientName || '',
    date: paidDate,
    amount: amount,
    method: params.method || '',
    notes: 'Payment plan: ' + (label || 'installment')
  });

  const set = function (h, v) { const c = headers.indexOf(h); if (c !== -1) sheet.getRange(rowNum, c + 1).setValue(v); };
  set('Status', 'paid');
  set('Paid Date', "'" + paidDate);
  if (pay && pay.paymentId) set('Payment ID', pay.paymentId);

  return { success: true, paymentId: pay ? pay.paymentId : '' };
}

function updatePlanInstallment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getPaymentPlanSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowNum = findRowByValue(sheet, 'Plan ID', params.planId);
  if (rowNum === -1) return { success: false, error: 'Installment not found.' };
  const set = function (h, v) { const c = headers.indexOf(h); if (c !== -1) sheet.getRange(rowNum, c + 1).setValue(v); };
  if (params.dueDate !== undefined) set('Due Date', params.dueDate ? "'" + params.dueDate : '');
  if (params.amount !== undefined) set('Amount (TND)', Number(params.amount) || 0);
  if (params.label !== undefined) set('Label', params.label);
  if (params.invoiced !== undefined) set('Invoiced', params.invoiced === 'yes' || params.invoiced === true ? 'Yes' : 'No');
  return { success: true };
}

function deletePlanInstallment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getPaymentPlanSheet();
  const rowNum = findRowByValue(sheet, 'Plan ID', params.planId);
  if (rowNum === -1) return { success: false, error: 'Installment not found.' };
  sheet.deleteRow(rowNum);
  return { success: true };
}

// Finds an existing client by phone (or name), and creates a profile if there
// isn't one. Used by the backfill so past appointments also produce clients in
// the client list — bookings and profiles are separate records, and entering a
// booking alone doesn't create a client.
function ensureClientProfile(name, phone, location, staffName, serviceName) {
  const sheet = getClientsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getDataRange().getValues();
  const idC = headers.indexOf('Client ID');
  const nameC = headers.indexOf('Client Name');
  const phoneC = headers.indexOf('Phone');

  const wantPhone = normPhone(phone);
  const wantName = normName(name);

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][idC]) continue;
    const rowPhone = normPhone(rows[i][phoneC]);
    const rowName = normName(rows[i][nameC]);
    if (wantPhone && rowPhone && wantPhone === rowPhone) return { id: String(rows[i][idC]), created: false };
    if (!wantPhone && wantName && rowName === wantName) return { id: String(rows[i][idC]), created: false };
  }

  // Match the service to one of the six packages where possible
  let pkg = null;
  const sLow = String(serviceName || '').toLowerCase();
  for (let p = 0; p < PACKAGES.length; p++) {
    if (sLow && (sLow.indexOf(PACKAGES[p].name.toLowerCase()) !== -1 ||
                 PACKAGES[p].name.toLowerCase().indexOf(sLow) !== -1)) {
      pkg = PACKAGES[p];
      break;
    }
  }

  const progress = {};
  if (pkg) {
    pkg.components.forEach(function (comp) {
      if (comp.type === 'sessions') progress[comp.id] = comp.sessions.map(function (s) { return { id: s.id, completed: false }; });
      else progress[comp.id] = { done: 0, target: comp.target };
    });
  }

  const id = 'C' + Date.now() + Math.floor(Math.random() * 1000);
  const rowMap = {
    'Client ID': id,
    'Client Name': name || '',
    'Phone': phone || '',
    'Email': '',
    'Location': location || '',
    'Staff': staffName || '',
    'Package ID': pkg ? pkg.id : '',
    'Package Price (TND)': pkg ? pkg.price : '',
    'Custom Price (TND)': '',
    'Progress JSON': JSON.stringify(progress),
    'Status': 'active',
    'Created At': new Date().toISOString(),
    'Notes': pkg ? '' : ('Service: ' + (serviceName || 'unspecified')),
    'Assigned Staff': staffName || '',
    'Baby Name': '', 'Baby Birth Date': '', 'Baby Weight (kg)': '', 'Baby Birth Details': ''
  };
  const row = headers.map(function (h) { return (h in rowMap) ? rowMap[h] : ''; });
  sheet.appendRow(row);
  return { id: id, created: true };
}

// ───────────────────────── BULK BACKFILL ─────────────────────────
// Adds many past appointments at once (e.g. January/February history).
// Writes through the same header-NAME-mapped path as createAppointment so
// rows can never misalign, and coerces every numeric value to a real number
// so a NaN/undefined can't collapse a cell and shift the row.
function bulkAddAppointments(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let rows = [];
  try { rows = JSON.parse(params.rows || '[]'); } catch (e) { rows = []; }
  if (!rows.length) return { success: false, error: 'No appointments to add.' };

  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  const sheet = ss.getSheetByName('Bookings');
  if (!sheet) return { success: false, error: 'Bookings sheet not found' };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const num = function (v) { const n = Number(v); return isNaN(n) ? 0 : n; };
  const out = [];
  const errors = [];
  const createProfiles = String(params.createProfiles || 'yes') === 'yes';
  let profilesCreated = 0;

  rows.forEach(function (r, i) {
    if (!r.date || !r.clientName) {
      errors.push('Row ' + (i + 1) + ': needs at least a date and a client name.');
      return;
    }

    // Make sure this person exists in the client list — a booking alone
    // doesn't create a profile, so backfilled clients would otherwise be
    // invisible in the Clients page.
    if (createProfiles) {
      try {
        const res = ensureClientProfile(r.clientName, r.phone, r.location, r.staff || staff.name, r.service);
        if (res && res.created) profilesCreated += 1;
      } catch (e) { /* keep going — the booking row still gets written */ }
    }

    const basePrice = num(r.basePrice);
    const rowMap = {
      'Booking ID': 'BACKFILL-' + Date.now() + '-' + i,
      'Booked At': new Date(),
      'Status': 'Confirmed',
      'Language': 'EN',
      'Service (EN)': r.service || 'Appointment',
      'Service (FR)': '',
      'Package / Option': r.packageOption || '',
      'Client Name': r.clientName,
      'Phone': r.phone || '',
      'Email': '',
      'Appointment Date': "'" + r.date,
      'Time': r.time ? "'" + r.time : '',
      'Duration (min)': num(r.duration) || 60,
      'Location': r.location || '',
      'Transport Zone': '',
      'Transport Fee (TND)': 0,
      'Pump Model': '',
      'Rental Period': '',
      'Base Price (TND)': basePrice,
      'Total (TND)': basePrice, // no transport on backfilled rows
      'Notes': r.notes || '',
      'Staff': r.staff || staff.name
    };
    out.push(headers.map(function (h) { return (h in rowMap) ? rowMap[h] : ''; }));
  });

  if (!out.length) return { success: false, error: errors.join(' ') || 'Nothing valid to add.' };

  // Write all rows in one go
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, out.length, headers.length).setValues(out);

  return { success: true, added: out.length, profilesCreated: profilesCreated, errors: errors };
}

function createAppointment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  const sheet = ss.getSheetByName('Bookings');
  if (!sheet) return { success: false, error: 'Bookings sheet not found' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const bookingId = 'DASH-' + Date.now();
  const clientName = params.clientName || '';
  const date = params.date || ''; // yyyy-mm-dd
  const time = params.time || ''; // HH:mm
  const duration = parseInt(params.duration, 10) || 60;
  const location = params.location || staff.location || '';
  const assignedStaff = params.staff || staff.name;
  const notes = params.notes || '';

  // If a real package was chosen, use its name + price; otherwise fall back to
  // the free-text title (one-off appointments still work as before).
  let serviceName = params.title || 'Appointment';
  let packageOption = '';
  let basePrice = params.basePrice || '';
  if (params.packageId) {
    const pkg = PACKAGES.filter(function (p) { return p.id === params.packageId; })[0];
    if (pkg) {
      serviceName = pkg.name;
      packageOption = pkg.name;
      // Use the explicit basePrice if one was passed (e.g. a custom price),
      // otherwise default to the package's standard price.
      if (basePrice === '' || basePrice === undefined || basePrice === null) {
        basePrice = pkg.price;
      }
    }
  }
  const title = serviceName;

  // Build the row keyed by header name; anything the dashboard doesn't set
  // (transport, pump) is left blank.
  const rowMap = {
    'Booking ID': bookingId,
    'Booked At': new Date(),
    'Status': 'Confirmed',
    'Language': 'EN',
    'Service (EN)': serviceName,
    'Service (FR)': '',
    'Package / Option': packageOption,
    'Client Name': clientName,
    'Phone': params.phone || '',
    'Email': params.email || '',
    'Appointment Date': "'" + date,
    'Time': time ? "'" + time : '',
    'Duration (min)': duration,
    'Location': location,
    'Transport Zone': '',
    'Transport Fee (TND)': '',
    'Pump Model': '',
    'Rental Period': '',
    'Base Price (TND)': basePrice,
    'Total (TND)': basePrice,
    'Notes': notes,
    'Staff': assignedStaff
  };

  const row = headers.map(function (h) { return (h in rowMap) ? rowMap[h] : ''; });
  sheet.appendRow(row);

  // Create the Google Calendar event
  let calendarOk = false;
  try {
    if (date && time) {
      const parts = date.split('-').map(Number);
      const timeParts = time.split(':').map(Number);
      const start = new Date(parts[0], parts[1] - 1, parts[2], timeParts[0], timeParts[1] || 0, 0);
      const end = new Date(start.getTime() + duration * 60000);
      const calendar = CalendarApp.getCalendarById(DASH_CALENDAR_ID);
      if (calendar) {
        const evtTitle = title + (clientName ? ' | ' + clientName : '');
        const descLines = [];
        if (clientName) descLines.push('Client: ' + clientName);
        if (params.phone) descLines.push('Phone: ' + params.phone);
        if (location) descLines.push('Location: ' + location);
        if (assignedStaff) descLines.push('Staff: ' + assignedStaff);
        if (notes) descLines.push('Notes: ' + notes);
        descLines.push('', 'Created from dashboard · ID: ' + bookingId);
        calendar.createEvent(evtTitle, start, end, { description: descLines.join('\n') });
        calendarOk = true;
      }
    }
  } catch (err) {
    // Row is already saved; report calendar failure but don't fail the whole op
    return { success: true, bookingId: bookingId, calendarOk: false, calendarError: err.message };
  }

  return { success: true, bookingId: bookingId, calendarOk: calendarOk };
}

// ───────────────────────── CBE PROGRESS ─────────────────────────

function getCbeSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  return ss.getSheetByName('CBE Progress');
}

function getClientProgress(pin, clientName) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getCbeSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[headers.indexOf('Client Name')]).trim() === String(clientName).trim()) {
      const progress = CBE_SESSIONS.map(s => ({
        id: s.id,
        name: s.name,
        completed: String(row[headers.indexOf(s.name)]).toUpperCase() === 'TRUE'
      }));
      return { success: true, clientName: clientName, progress: progress, found: true };
    }
  }

  // Not found yet — return blank progress
  return {
    success: true,
    clientName: clientName,
    progress: CBE_SESSIONS.map(s => ({ id: s.id, name: s.name, completed: false })),
    found: false
  };
}

function updateProgress(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getCbeSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const clientName = params.clientName;
  const sessionId = parseInt(params.sessionId, 10);
  const completed = params.completed === 'true';
  const session = CBE_SESSIONS.find(s => s.id === sessionId);
  if (!session) return { success: false, error: 'Invalid session ID' };

  const colIndex = headers.indexOf(session.name);
  let rowIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][headers.indexOf('Client Name')]).trim() === String(clientName).trim()) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    // Create new row
    const newRow = new Array(headers.length).fill('');
    newRow[headers.indexOf('Client Name')] = clientName;
    newRow[headers.indexOf('Phone')] = params.phone || '';
    newRow[headers.indexOf('Staff')] = staff.name;
    newRow[headers.indexOf('Location')] = params.location || staff.location;
    newRow[colIndex] = completed ? 'TRUE' : 'FALSE';
    sheet.appendRow(newRow);
  } else {
    sheet.getRange(rowIndex + 1, colIndex + 1).setValue(completed ? 'TRUE' : 'FALSE');
  }

  return { success: true };
}

// ───────────────────────── MONTHLY STATS ─────────────────────────

function getMonthlyStats(pin, month, year) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);

  const m = parseInt(month, 10); // 1-12
  const y = parseInt(year, 10);

  const filtered = bookings.filter(b => {
    const dateStr = normDate(b['Appointment Date']);
    if (!dateStr) return false;
    const d = new Date(dateStr);
    return d.getMonth() + 1 === m && d.getFullYear() === y &&
      String(b['Status']).toLowerCase() !== 'cancelled';
  });

  const byService = {};
  let totalRevenue = 0;

  filtered.forEach(b => {
    const service = b['Service (EN)'] || 'Other';
    if (!byService[service]) byService[service] = { count: 0, revenue: 0 };
    byService[service].count += 1;
    // Use Base Price, not Total — Total includes the transport/delivery fee,
    // which isn't taxable revenue (personal vehicle, no mileage declared).
    const basePrice = parseFloat(b['Base Price (TND)']) || 0;
    byService[service].revenue += basePrice;
    totalRevenue += basePrice;
  });

  return {
    success: true,
    month: m,
    year: y,
    totalBookings: filtered.length,
    totalRevenue: totalRevenue,
    byService: byService,
    bookings: filtered
  };
}

// ───────────────────────── CLIENTS (package-based profiles) ─────────────────────────

function getClientsSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  return ss.getSheetByName('Clients');
}

function getPaymentsSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  return ss.getSheetByName('Payments');
}

function rowToClient(row, headers) {
  const get = (h) => row[headers.indexOf(h)];
  let progress = {};
  try { progress = JSON.parse(get('Progress JSON') || '{}'); } catch (e) { progress = {}; }
  const pkg = PACKAGES.find(p => p.id === get('Package ID')) || null;
  // Assigned Staff is stored as a comma-separated list (multiple staff
  // can share a client, e.g. lead doula + backup).
  const assignedRaw = String(get('Assigned Staff') || '').trim();
  const assignedStaff = assignedRaw ? assignedRaw.split(',').map(s => s.trim()).filter(Boolean) : [];
  return {
    id: String(get('Client ID') || ''),
    name: String(get('Client Name') || ''),
    phone: String(get('Phone') || ''),
    email: String(get('Email') || ''),
    location: String(get('Location') || ''),
    staff: String(get('Staff') || ''),
    assignedStaff: assignedStaff,
    packageId: String(get('Package ID') || ''),
    packageName: pkg ? pkg.name : String(get('Package ID') || ''),
    packagePrice: parseFloat(get('Package Price (TND)')) || 0,
    customPrice: get('Custom Price (TND)') !== '' ? parseFloat(get('Custom Price (TND)')) : null,
    progress: progress,
    status: String(get('Status') || 'active'),
    createdAt: String(get('Created At') || ''),
    notes: String(get('Notes') || ''),
    babyName: String(get('Baby Name') || ''),
    babyBirthDate: get('Baby Birth Date') ? normDate(get('Baby Birth Date')) : '',
    babyWeight: String(get('Baby Weight (kg)') || ''),
    babyBirthDetails: String(get('Baby Birth Details') || '')
  };
}

function getClients(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let clients = [];

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][headers.indexOf('Client ID')]) continue;
    clients.push(rowToClient(rows[i], headers));
  }

  if (staff.role !== 'owner' && staff.location !== 'All') {
    // Staff see clients that are either at their location, created by them,
    // or where they're in the assigned-staff list (e.g. a backup doula
    // assigned to a client based in the other city still sees them).
    clients = clients.filter(c =>
      c.staff === staff.name ||
      c.location === staff.location ||
      (c.assignedStaff && c.assignedStaff.indexOf(staff.name) !== -1)
    );
  }

  // Attach payment totals to each client
  const payments = getAllPaymentsRaw();
  clients.forEach(c => {
    const clientPayments = payments.filter(p => p.clientId === c.id);
    c.totalPaid = clientPayments.reduce((sum, p) => sum + p.amount, 0);
    const price = c.customPrice !== null ? c.customPrice : c.packagePrice;
    c.totalOwed = price;
    c.balance = price - c.totalPaid;
  });

  return { success: true, clients: clients, packages: PACKAGES, staffList: getStaffList() };
}

// Returns the list of active staff names + locations, for assignment and
// location dropdowns in the dashboard. Owner-level info (PINs) is never
// included — only what the UI needs.
function getStaffList() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  const sheet = ss.getSheetByName('Staff');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const active = rows[i][headers.indexOf('Active')];
    if (String(active).toUpperCase() !== 'TRUE') continue;
    list.push({
      name: String(rows[i][headers.indexOf('Name')] || ''),
      role: String(rows[i][headers.indexOf('Role')] || ''),
      location: String(rows[i][headers.indexOf('Location')] || '')
    });
  }
  return list;
}

function updateClient(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('Client ID');

  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(params.clientId)) { rowIndex = i; break; }
  }
  if (rowIndex === -1) return { success: false, error: 'Client not found' };

  // Only update fields that were actually passed in — leaves everything
  // else (progress, payments, created date) untouched.
  const setField = (header, value) => {
    const col = headers.indexOf(header);
    if (col !== -1 && value !== undefined) {
      sheet.getRange(rowIndex + 1, col + 1).setValue(value);
    }
  };

  if (params.clientName !== undefined) setField('Client Name', params.clientName);
  if (params.phone !== undefined) setField('Phone', params.phone);
  if (params.email !== undefined) setField('Email', params.email);
  if (params.location !== undefined) setField('Location', params.location);
  if (params.assignedStaff !== undefined) setField('Assigned Staff', params.assignedStaff);
  if (params.customPrice !== undefined) setField('Custom Price (TND)', params.customPrice);
  if (params.status !== undefined) setField('Status', params.status);
  if (params.notes !== undefined) setField('Notes', params.notes);
  if (params.babyName !== undefined) setField('Baby Name', params.babyName);
  if (params.babyBirthDate !== undefined) setField('Baby Birth Date', params.babyBirthDate ? "'" + params.babyBirthDate : '');
  if (params.babyWeight !== undefined) setField('Baby Weight (kg)', params.babyWeight);
  if (params.babyBirthDetails !== undefined) setField('Baby Birth Details', params.babyBirthDetails);

  // If the package changed, update package id + price and re-init progress
  // structure for the new package (only if it actually changed, so we
  // don't wipe progress on an unrelated edit).
  if (params.packageId !== undefined) {
    const currentPkgId = String(rows[rowIndex][headers.indexOf('Package ID')]);
    if (params.packageId !== currentPkgId) {
      const pkg = PACKAGES.find(p => p.id === params.packageId);
      if (pkg) {
        setField('Package ID', pkg.id);
        setField('Package Price (TND)', pkg.price);
        const progress = {};
        pkg.components.forEach(comp => {
          if (comp.type === 'sessions') {
            progress[comp.id] = comp.sessions.map(s => ({ id: s.id, completed: false }));
          } else {
            progress[comp.id] = { done: 0, target: comp.target };
          }
        });
        setField('Progress JSON', JSON.stringify(progress));
      }
    }
  }

  return { success: true };
}

// Permanently deletes a client and all their associated financial/log
// records (payments, add-on charges, appointment-notes). Owner-only, since
// this erases records that may have fed into past accounting periods.
// Returns counts of what was removed so the UI can confirm.
function deleteClient(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (staff.role !== 'owner') return { success: false, error: 'Only the owner can delete clients.' };

  const clientId = params.clientId;
  if (!clientId) return { success: false, error: 'No client specified.' };

  // Remove the client's own row from the Clients sheet
  const clientsSheet = getClientsSheet();
  const cRows = clientsSheet.getDataRange().getValues();
  const cIdCol = cRows[0].indexOf('Client ID');
  let clientName = '';
  let clientRemoved = false;
  for (let i = cRows.length - 1; i >= 1; i--) {
    if (String(cRows[i][cIdCol]) === String(clientId)) {
      clientName = cRows[i][cRows[0].indexOf('Client Name')];
      clientsSheet.deleteRow(i + 1);
      clientRemoved = true;
    }
  }
  if (!clientRemoved) return { success: false, error: 'Client not found.' };

  // Helper: delete every row in a sheet whose Client ID matches. Iterates
  // bottom-up so row indices stay valid as rows are removed.
  const purgeByClientId = function (sheet) {
    if (!sheet) return 0;
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 2) return 0;
    const idCol = rows[0].indexOf('Client ID');
    if (idCol === -1) return 0;
    let removed = 0;
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][idCol]) === String(clientId)) {
        sheet.deleteRow(i + 1);
        removed++;
      }
    }
    return removed;
  };

  const paymentsRemoved = purgeByClientId(getPaymentsSheet());
  const chargesRemoved = purgeByClientId(getChargesSheet());
  const appointmentsRemoved = purgeByClientId(getAppointmentsSheet());

  return {
    success: true,
    clientName: clientName,
    paymentsRemoved: paymentsRemoved,
    chargesRemoved: chargesRemoved,
    appointmentsRemoved: appointmentsRemoved
  };
}

function createClient(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const pkg = PACKAGES.find(p => p.id === params.packageId);
  if (!pkg) return { success: false, error: 'Unknown package: ' + params.packageId };

  const sheet = getClientsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const id = 'C' + Date.now();

  // Initialize progress structure based on package components
  const progress = {};
  pkg.components.forEach(comp => {
    if (comp.type === 'sessions') {
      progress[comp.id] = comp.sessions.map(s => ({ id: s.id, completed: false }));
    } else {
      progress[comp.id] = { done: 0, target: comp.target };
    }
  });

  // Build a value-per-column map keyed by header name, then emit in the
  // sheet's real header order — so adding columns later can't shift data
  // into the wrong column (the bug that hit the Bookings sheet before).
  const rowMap = {
    'Client ID': id,
    'Client Name': params.clientName || '',
    'Phone': params.phone || '',
    'Email': params.email || '',
    'Location': params.location || staff.location,
    'Staff': params.staff || staff.name,
    'Package ID': pkg.id,
    'Package Price (TND)': pkg.price,
    'Custom Price (TND)': params.customPrice || '',
    'Progress JSON': JSON.stringify(progress),
    'Status': 'active',
    'Created At': new Date().toISOString(),
    'Notes': params.notes || '',
    'Assigned Staff': params.assignedStaff || '',
    'Baby Name': params.babyName || '',
    'Baby Birth Date': params.babyBirthDate ? "'" + params.babyBirthDate : '',
    'Baby Weight (kg)': params.babyWeight || '',
    'Baby Birth Details': params.babyBirthDetails || ''
  };

  const row = headers.map(function (h) { return (h in rowMap) ? rowMap[h] : ''; });
  sheet.appendRow(row);

  return { success: true, clientId: id };
}

function updateClientComponent(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('Client ID');
  const progCol = headers.indexOf('Progress JSON');

  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(params.clientId)) { rowIndex = i; break; }
  }
  if (rowIndex === -1) return { success: false, error: 'Client not found' };

  let progress = {};
  try { progress = JSON.parse(rows[rowIndex][progCol] || '{}'); } catch (e) { progress = {}; }

  const componentId = params.componentId;

  if (params.sessionId !== undefined && params.sessionId !== '') {
    // Sessions-type component: toggle a specific session
    const sessionId = parseInt(params.sessionId, 10);
    const completed = params.completed === 'true';
    if (!progress[componentId]) progress[componentId] = [];
    const existing = progress[componentId].find(s => s.id === sessionId);
    if (existing) {
      existing.completed = completed;
    } else {
      progress[componentId].push({ id: sessionId, completed: completed });
    }
  } else if (params.done !== undefined) {
    // Count-type component: set a tally value
    const done = parseInt(params.done, 10);
    if (!progress[componentId]) progress[componentId] = { done: 0, target: 1 };
    progress[componentId].done = done;
  }

  sheet.getRange(rowIndex + 1, progCol + 1).setValue(JSON.stringify(progress));
  return { success: true, progress: progress };
}

// ───────────────────────── PAYMENTS ─────────────────────────

function getAllPaymentsRaw() {
  const sheet = getPaymentsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const payments = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[headers.indexOf('Payment ID')]) continue;
    const dateVal = row[headers.indexOf('Date')];
    payments.push({
      id: String(row[headers.indexOf('Payment ID')]),
      clientId: String(row[headers.indexOf('Client ID')]),
      clientName: String(row[headers.indexOf('Client Name')]),
      date: normDate(dateVal),
      amount: parseFloat(row[headers.indexOf('Amount (TND)')]) || 0,
      method: String(row[headers.indexOf('Method')] || ''),
      notes: String(row[headers.indexOf('Notes')] || ''),
      recordedBy: String(row[headers.indexOf('Recorded By')] || '')
    });
  }
  return payments;
}

function addPayment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getPaymentsSheet();
  const id = 'P' + Date.now();

  sheet.appendRow([
    id,
    params.clientId || '',
    params.clientName || '',
    params.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    parseFloat(params.amount) || 0,
    params.method || '',
    params.notes || '',
    staff.name
  ]);

  return { success: true, paymentId: id };
}

function getClientLedger(pin, clientId) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let client = null;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][headers.indexOf('Client ID')]) === String(clientId)) {
      client = rowToClient(rows[i], headers);
      break;
    }
  }
  if (!client) return { success: false, error: 'Client not found' };

  const payments = getAllPaymentsRaw().filter(p => p.clientId === clientId);
  const totalPaid = payments.reduce((sum, p) => sum + p.amount, 0);
  const packagePrice = client.customPrice !== null ? client.customPrice : client.packagePrice;

  // Add-on services (extra charges beyond the package) increase what's owed.
  const charges = getChargesForClient(clientId);
  const addOnTotal = charges.reduce((sum, c) => sum + c.amount, 0);
  const totalOwed = Math.round((packagePrice + addOnTotal) * 1000) / 1000;

  return {
    success: true,
    client: client,
    payments: payments,
    charges: charges,
    packagePrice: packagePrice,
    addOnTotal: addOnTotal,
    totalOwed: totalOwed,
    totalPaid: totalPaid,
    balance: Math.round((totalOwed - totalPaid) * 1000) / 1000
  };
}

// ───────────────────────── ADD-ON SERVICES (extra charges) ─────────────────────────
// Lets a client be charged for services beyond their original package
// (an extra session, a one-off add-on, etc.). Each charge increases the
// amount owed in that client's ledger. Stored in its own "Charges" sheet,
// created on first use so no setup re-run is needed.
function getChargesSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  let sheet = ss.getSheetByName('Charges');
  if (!sheet) {
    sheet = ss.insertSheet('Charges');
    sheet.appendRow(['Charge ID', 'Client ID', 'Date', 'Description', 'Amount (TND)', 'Created At', 'Staff']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

function getChargesForClient(clientId) {
  const sheet = getChargesSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][headers.indexOf('Client ID')]) !== String(clientId)) continue;
    out.push({
      id: String(rows[i][headers.indexOf('Charge ID')]),
      date: normDate(rows[i][headers.indexOf('Date')]),
      description: String(rows[i][headers.indexOf('Description')] || ''),
      amount: parseFloat(rows[i][headers.indexOf('Amount (TND)')]) || 0
    });
  }
  return out;
}

function addCharge(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const amount = parseFloat(params.amount) || 0;
  if (!params.description || amount <= 0) return { success: false, error: 'Enter a description and a price.' };

  const sheet = getChargesSheet();
  const id = 'CH' + Date.now();
  sheet.appendRow([
    id,
    params.clientId || '',
    params.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    params.description,
    amount,
    new Date().toISOString(),
    staff.name
  ]);
  return { success: true, chargeId: id };
}

// ───────────────────────── APPOINTMENTS LOG (with notes) ─────────────────────────
// A per-client log of appointments and the notes taken at each. Separate
// from Bookings (the public booking flow) — this is the staff's private
// working record for a client. Stored in its own "Appointments" sheet,
// created on first use.
function getAppointmentsSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  let sheet = ss.getSheetByName('Appointments');
  if (!sheet) {
    sheet = ss.insertSheet('Appointments');
    sheet.appendRow(['Appointment ID', 'Client ID', 'Date', 'Title', 'Notes', 'Created At', 'Staff']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

function getAppointments(pin, clientId) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getAppointmentsSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, appointments: [] };
  const headers = rows[0];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][headers.indexOf('Client ID')]) !== String(clientId)) continue;
    out.push({
      id: String(rows[i][headers.indexOf('Appointment ID')]),
      date: normDate(rows[i][headers.indexOf('Date')]),
      title: String(rows[i][headers.indexOf('Title')] || ''),
      notes: String(rows[i][headers.indexOf('Notes')] || ''),
      staff: String(rows[i][headers.indexOf('Staff')] || '')
    });
  }
  out.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); }); // most recent first
  return { success: true, appointments: out };
}

// ───────────────────────── GROUP CLASSES ─────────────────────────
// A group class is one session-set (one or more dates) shared by several
// clients. It creates ONE calendar event per date (not one per client), drops
// a note on each attached client's profile, and records the roster. Each
// client still pays their own fee on their own profile — this only handles
// scheduling, not billing.
function getGroupClassesSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  let sheet = ss.getSheetByName('Group Classes');
  if (!sheet) {
    sheet = ss.insertSheet('Group Classes');
    sheet.appendRow(['Group ID', 'Name', 'Client IDs (JSON)', 'Client Names (JSON)', 'Dates (JSON)', 'Time', 'Location', 'Notes', 'Created At', 'Staff']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  return sheet;
}

function createGroupClass(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const name = (params.name || '').trim();
  if (!name) return { success: false, error: 'Give the group class a name.' };

  let clientIds = [];
  let dates = [];
  try { clientIds = JSON.parse(params.clientIds || '[]'); } catch (e) { clientIds = []; }
  try { dates = JSON.parse(params.dates || '[]'); } catch (e) { dates = []; }

  if (!clientIds.length) return { success: false, error: 'Add at least one client.' };
  if (!dates.length) return { success: false, error: 'Add at least one session date.' };

  const time = params.time || '';
  const location = params.location || staff.location || '';
  const notes = params.notes || '';

  // Resolve client names from their profiles (for the calendar event + roster)
  const cSheet = getClientsSheet();
  const cRows = cSheet.getDataRange().getValues();
  const cH = cRows[0];
  const idCol = cH.indexOf('Client ID');
  const nameCol = cH.indexOf('Client Name');
  const namesById = {};
  for (let i = 1; i < cRows.length; i++) {
    namesById[String(cRows[i][idCol])] = String(cRows[i][nameCol] || '');
  }
  const clientNames = clientIds.map(function (id) { return namesById[String(id)] || 'Client'; });

  const groupId = 'G' + Date.now();

  // Record the group class row
  const gSheet = getGroupClassesSheet();
  gSheet.appendRow([
    groupId,
    name,
    JSON.stringify(clientIds),
    JSON.stringify(clientNames),
    JSON.stringify(dates),
    time ? "'" + time : '',
    location,
    notes,
    new Date().toISOString(),
    staff.name
  ]);

  // One calendar event per date (shared by the whole group)
  let calendarOk = true;
  try {
    const calendar = CalendarApp.getCalendarById(DASH_CALENDAR_ID);
    if (calendar) {
      dates.forEach(function (d) {
        if (!d) return;
        const parts = d.split('-').map(Number);
        const tp = (time || '09:00').split(':').map(Number);
        const start = new Date(parts[0], parts[1] - 1, parts[2], tp[0], tp[1] || 0, 0);
        const end = new Date(start.getTime() + 90 * 60000); // default 90-min class
        const title = name + ' (' + clientIds.length + ' clients)';
        const desc = 'Group class: ' + name + '\nClients: ' + clientNames.join(', ') + (notes ? '\n' + notes : '');
        calendar.createEvent(title, start, end, { description: desc });
      });
    } else {
      calendarOk = false;
    }
  } catch (err) {
    calendarOk = false;
  }

  // A note on each attached client's profile, so it shows in their history
  const aSheet = getAppointmentsSheet();
  const dateLabel = dates.join(', ');
  clientIds.forEach(function (id) {
    aSheet.appendRow([
      'A' + Date.now() + '-' + id,
      id,
      dates[0] || '',
      'Group class: ' + name,
      'Part of group class "' + name + '" on ' + dateLabel + (time ? ' at ' + time : '') + (location ? ' · ' + location : ''),
      new Date().toISOString(),
      staff.name
    ]);
  });

  return { success: true, groupId: groupId, calendarOk: calendarOk, sessions: dates.length, clients: clientIds.length };
}

function getGroupClasses(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getGroupClassesSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, classes: [] };
  const h = rows[0];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    let clientNames = [];
    let clientIds = [];
    let dates = [];
    try { clientNames = JSON.parse(rows[i][h.indexOf('Client Names (JSON)')] || '[]'); } catch (e) {}
    try { clientIds = JSON.parse(rows[i][h.indexOf('Client IDs (JSON)')] || '[]'); } catch (e) {}
    try { dates = JSON.parse(rows[i][h.indexOf('Dates (JSON)')] || '[]'); } catch (e) {}
    out.push({
      id: String(rows[i][h.indexOf('Group ID')]),
      name: String(rows[i][h.indexOf('Name')] || ''),
      clientNames: clientNames,
      clientIds: clientIds,
      dates: dates,
      time: cleanTime(rows[i][h.indexOf('Time')]),
      location: String(rows[i][h.indexOf('Location')] || ''),
      notes: String(rows[i][h.indexOf('Notes')] || '')
    });
  }
  out.reverse(); // newest first
  return { success: true, classes: out };
}

function deleteGroupClass(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (staff.role !== 'owner') return { success: false, error: 'Only the owner can delete group classes.' };
  const sheet = getGroupClassesSheet();
  const rowNum = findRowByValue(sheet, 'Group ID', params.groupId);
  if (rowNum === -1) return { success: false, error: 'Group class not found.' };
  sheet.deleteRow(rowNum);
  return { success: true, deleted: true };
}

// Edit an existing group class: change its name, dates, time, location, notes,
// or roster. Re-resolves client names from the roster IDs. Does not touch the
// original calendar events (those were created at booking time) — the dashboard
// calendar reads the Group Classes sheet directly, so it reflects edits.
function updateGroupClass(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getGroupClassesSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowNum = findRowByValue(sheet, 'Group ID', params.groupId);
  if (rowNum === -1) return { success: false, error: 'Group class not found.' };

  const set = function (h, v) { const c = headers.indexOf(h); if (c !== -1) sheet.getRange(rowNum, c + 1).setValue(v); };

  if (params.name !== undefined) set('Name', params.name);
  if (params.time !== undefined) set('Time', params.time ? "'" + params.time : '');
  if (params.location !== undefined) set('Location', params.location);
  if (params.notes !== undefined) set('Notes', params.notes);

  if (params.dates !== undefined) {
    let dates = [];
    try { dates = JSON.parse(params.dates || '[]'); } catch (e) { dates = []; }
    set('Dates (JSON)', JSON.stringify(dates));
  }

  if (params.clientIds !== undefined) {
    let clientIds = [];
    try { clientIds = JSON.parse(params.clientIds || '[]'); } catch (e) { clientIds = []; }
    // Re-resolve names
    const cSheet = getClientsSheet();
    const cRows = cSheet.getDataRange().getValues();
    const cH = cRows[0];
    const idCol = cH.indexOf('Client ID');
    const nameCol = cH.indexOf('Client Name');
    const namesById = {};
    for (let i = 1; i < cRows.length; i++) namesById[String(cRows[i][idCol])] = String(cRows[i][nameCol] || '');
    const clientNames = clientIds.map(function (id) { return namesById[String(id)] || 'Client'; });
    set('Client IDs (JSON)', JSON.stringify(clientIds));
    set('Client Names (JSON)', JSON.stringify(clientNames));
  }

  return { success: true };
}

function addAppointment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (!params.date && !params.title && !params.notes) {
    return { success: false, error: 'Add a date, a title, or some notes.' };
  }
  const sheet = getAppointmentsSheet();
  const id = 'A' + Date.now();
  sheet.appendRow([
    id,
    params.clientId || '',
    params.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    params.title || '',
    params.notes || '',
    new Date().toISOString(),
    staff.name
  ]);
  return { success: true, appointmentId: id };
}

// ───────────────────────── EDIT / DELETE individual records ─────────────────────────
// Shared helper: finds a row in a sheet by an ID column + value, returns its
// 1-based sheet row number (or -1).
function findRowByValue(sheet, idColName, idValue) {
  const rows = sheet.getDataRange().getValues();
  const idCol = rows[0].indexOf(idColName);
  if (idCol === -1) return -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(idValue)) return i + 1; // 1-based for getRange
  }
  return -1;
}

// PAYMENTS — edit amount/method/date, or delete the entry
function updatePayment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getPaymentsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowNum = findRowByValue(sheet, 'Payment ID', params.paymentId);
  if (rowNum === -1) return { success: false, error: 'Payment not found.' };
  const set = (h, v) => { const c = headers.indexOf(h); if (c !== -1 && v !== undefined) sheet.getRange(rowNum, c + 1).setValue(v); };
  if (params.amount !== undefined) set('Amount (TND)', parseFloat(params.amount) || 0);
  if (params.method !== undefined) set('Method', params.method);
  if (params.date !== undefined && params.date) set('Date', "'" + params.date);
  return { success: true };
}

function deletePayment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getPaymentsSheet();
  const rowNum = findRowByValue(sheet, 'Payment ID', params.paymentId);
  if (rowNum === -1) return { success: false, error: 'Payment not found.' };
  sheet.deleteRow(rowNum);
  return { success: true, deleted: true };
}

// CHARGES (add-on services) — edit description/amount, or delete
function updateCharge(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getChargesSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowNum = findRowByValue(sheet, 'Charge ID', params.chargeId);
  if (rowNum === -1) return { success: false, error: 'Charge not found.' };
  const set = (h, v) => { const c = headers.indexOf(h); if (c !== -1 && v !== undefined) sheet.getRange(rowNum, c + 1).setValue(v); };
  if (params.description !== undefined) set('Description', params.description);
  if (params.amount !== undefined) set('Amount (TND)', parseFloat(params.amount) || 0);
  if (params.date !== undefined && params.date) set('Date', "'" + params.date);
  return { success: true };
}

function deleteCharge(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getChargesSheet();
  const rowNum = findRowByValue(sheet, 'Charge ID', params.chargeId);
  if (rowNum === -1) return { success: false, error: 'Charge not found.' };
  sheet.deleteRow(rowNum);
  return { success: true, deleted: true };
}

// APPOINTMENT-LOG entries — edit date/title/notes, or delete
function updateAppointment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getAppointmentsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowNum = findRowByValue(sheet, 'Appointment ID', params.appointmentId);
  if (rowNum === -1) return { success: false, error: 'Appointment not found.' };
  const set = (h, v) => { const c = headers.indexOf(h); if (c !== -1 && v !== undefined) sheet.getRange(rowNum, c + 1).setValue(v); };
  if (params.date !== undefined && params.date) set('Date', "'" + params.date);
  if (params.title !== undefined) set('Title', params.title);
  if (params.notes !== undefined) set('Notes', params.notes);
  return { success: true };
}

function deleteAppointment(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  const sheet = getAppointmentsSheet();
  const rowNum = findRowByValue(sheet, 'Appointment ID', params.appointmentId);
  if (rowNum === -1) return { success: false, error: 'Appointment not found.' };
  sheet.deleteRow(rowNum);
  return { success: true, deleted: true };
}

// ───────────────────────── EXPENSES + ACCOUNTING REPORTS ─────────────────────────
function r3(n) { return Math.round((parseFloat(n) || 0) * 1000) / 1000; }

function getExpensesSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  let sheet = ss.getSheetByName('Expenses');
  if (!sheet) {
    sheet = ss.insertSheet('Expenses');
    sheet.appendRow(['Expense ID', 'Date', 'Category', 'Description', 'Supplier', 'Supplier MF',
      'Amount HT (TND)', 'TVA (TND)', 'Amount TTC (TND)', 'Receipt No', 'Payment Method', 'Created At', 'Staff']);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
  }
  return sheet;
}

// Adds an expense. Accepts any combination of HT / TVA / TTC and fills in the
// rest: enter just TTC and it splits at 19%; enter HT alone and TVA is added.
function addExpense(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const has = function (v) { return v !== undefined && v !== null && String(v).trim() !== '' && !isNaN(parseFloat(v)); };
  let HT = has(params.amountHt) ? r3(params.amountHt) : null;
  let TVA = has(params.tva) ? r3(params.tva) : null;
  let TTC = has(params.amountTtc) ? r3(params.amountTtc) : null;

  if (HT === null && TTC !== null && TVA !== null) HT = r3(TTC - TVA);
  else if (TVA === null && TTC !== null && HT !== null) TVA = r3(TTC - HT);
  else if (TTC === null && HT !== null && TVA !== null) TTC = r3(HT + TVA);
  else if (HT === null && TTC !== null && TVA === null) { HT = r3(TTC / (1 + TVA_RATE)); TVA = r3(TTC - HT); }
  else if (TTC === null && HT !== null && TVA === null) { TVA = r3(HT * TVA_RATE); TTC = r3(HT + TVA); }

  if (HT === null && TVA === null && TTC === null) return { success: false, error: 'Enter at least an amount (HT or TTC).' };
  HT = HT || 0; TVA = TVA || 0; TTC = TTC || r3(HT + TVA);

  const sheet = getExpensesSheet();
  const id = 'E' + Date.now();
  sheet.appendRow([
    id,
    params.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    params.category || '', params.description || '', params.supplier || '', params.supplierMf || '',
    HT, TVA, TTC, params.receiptNo || '', params.method || '',
    new Date().toISOString(), staff.name
  ]);
  return { success: true, expenseId: id };
}

function deleteExpense(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (staff.role !== 'owner') return { success: false, error: 'Only the owner can remove expenses.' };
  const sheet = getExpensesSheet();
  const rows = sheet.getDataRange().getValues();
  const idCol = rows[0].indexOf('Expense ID');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(params.expenseId)) {
      sheet.deleteRow(i + 1);
      return { success: true, deleted: true };
    }
  }
  return { success: false, error: 'Expense not found' };
}

function getExpensesInRange(start, end) {
  const sheet = getExpensesSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const h = rows[0];
  const idx = function (n) { return h.indexOf(n); };
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const ds = normDate(rows[i][idx('Date')]);
    if (!ds) continue;
    const d = new Date(ds);
    if (d < start || d > end) continue;
    out.push({
      id: String(rows[i][idx('Expense ID')]),
      date: ds,
      category: String(rows[i][idx('Category')] || ''),
      description: String(rows[i][idx('Description')] || ''),
      supplier: String(rows[i][idx('Supplier')] || ''),
      supplierMf: String(rows[i][idx('Supplier MF')] || ''),
      ht: parseFloat(rows[i][idx('Amount HT (TND)')]) || 0,
      tva: parseFloat(rows[i][idx('TVA (TND)')]) || 0,
      ttc: parseFloat(rows[i][idx('Amount TTC (TND)')]) || 0,
      receiptNo: String(rows[i][idx('Receipt No')] || ''),
      method: String(rows[i][idx('Payment Method')] || '')
    });
  }
  out.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  return out;
}

// Resolves a period type ('week' | 'month' | 'year') + an anchor date into a
// concrete start/end range and a human label.
function computePeriodRange(periodType, anchor) {
  const a = anchor ? new Date(anchor) : new Date();
  let start, end, label;
  if (periodType === 'week') {
    const day = a.getDay();
    const diff = a.getDate() - day + (day === 0 ? -6 : 1); // Monday
    start = new Date(a.getFullYear(), a.getMonth(), diff);
    end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6, 23, 59, 59, 999);
    label = 'Week of ' + Utilities.formatDate(start, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else if (periodType === 'year') {
    start = new Date(a.getFullYear(), 0, 1);
    end = new Date(a.getFullYear(), 11, 31, 23, 59, 59, 999);
    label = String(a.getFullYear());
  } else {
    start = new Date(a.getFullYear(), a.getMonth(), 1);
    end = new Date(a.getFullYear(), a.getMonth() + 1, 0, 23, 59, 59, 999);
    label = Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM yyyy');
  }
  start.setHours(0, 0, 0, 0);
  return { start: start, end: end, label: label };
}

// Full accounting report for a period: income (taxable CA HT, TVA collectée,
// TTC, plus non-taxable transport), service breakdown, expenses, the net TVA
// position, and the HT result. Income reads Base Price (taxable, treated as
// TTC) per the business rule; transport is reported separately as untaxed.
// ───────────────────────── FLEXIBLE REPORTS ─────────────────────────
// A single report over a chosen period: volume counts (sessions + unique
// clients) by service type, revenue by service and by city, a dedicated
// home-vs-office breakdown, and client stats. Uses Base Price for money.
function serviceTypeLabel(serviceName, packageLabel) {
  const s = String(serviceName || '').toLowerCase();
  const p = String(packageLabel || '').toLowerCase();
  if (s.indexOf('pump') !== -1 || s.indexOf('rental') !== -1) return 'Pump rentals';
  if (s.indexOf('breastfeeding') !== -1 || s.indexOf('lactation') !== -1) return 'Breastfeeding consultations';
  if (s.indexOf('postpartum') !== -1 || s.indexOf('post-partum') !== -1) return 'Postpartum support';
  if (s.indexOf('doula') !== -1 || s.indexOf('birth') !== -1 || p.indexOf('doula') !== -1 || s.indexOf('joyful birth method') !== -1) return 'Doula / births';
  if (s.indexOf('childbirth') !== -1 || s.indexOf('cbe') !== -1) return 'Childbirth education';
  if (s.indexOf('exercise') !== -1) return 'Exercises';
  return serviceName ? serviceName : 'Other';
}

// Classify a Location value into a visit type. Handles the mixed field:
// "At our office" / "Pick-up at office" → office; "Home visit (...)" → home
// (keeping the distance band); "Sousse"/"Tunis"/"All" → city (not a visit type).
function classifyLocation(loc) {
  const l = String(loc || '').trim();
  const low = l.toLowerCase();
  if (low.indexOf('home visit') !== -1) return { kind: 'home', label: l, city: '' };
  if (low.indexOf('office') !== -1) return { kind: 'office', label: 'At our office', city: '' };
  if (l === 'Sousse' || l === 'Tunis' || l === 'All') return { kind: 'city', label: l, city: l };
  if (!l) return { kind: 'unknown', label: 'Unspecified', city: '' };
  return { kind: 'other', label: l, city: '' };
}

function getFlexibleReport(pin, periodType, anchor, startStr, endStr) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  // Period: either a named period (month/year) via computePeriodRange, or a
  // custom start/end when periodType === 'custom'.
  let start, end, label;
  if (periodType === 'custom' && startStr && endStr) {
    start = new Date(startStr); start.setHours(0, 0, 0, 0);
    end = new Date(endStr); end.setHours(23, 59, 59, 999);
    label = startStr + ' to ' + endStr;
  } else {
    const range = computePeriodRange(periodType, anchor);
    start = range.start; end = range.end; label = range.label;
  }

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);
  const inRange = bookings.filter(function (b) {
    const ds = normDate(b['Appointment Date']);
    if (!ds) return false;
    const parts = ds.split('-').map(Number);
    const d = new Date(parts[0], parts[1] - 1, parts[2]);
    return d >= start && d <= end && String(b['Status']).toLowerCase() !== 'cancelled';
  });

  const round3 = function (n) { return Math.round(n * 1000) / 1000; };

  // Volume + revenue by service type
  const byService = {}; // label -> { sessions, clients:Set, revenue }
  // Home vs office split
  const visitSplit = { office: 0, home: 0, city: 0, other: 0, unknown: 0 };
  const homeBands = {}; // "Home visit (0 - 5 km)" -> count
  // Revenue by city (for the geographic rows)
  const byCity = {}; // city -> revenue
  let totalRevenue = 0;

  inRange.forEach(function (b) {
    const type = serviceTypeLabel(b['Service (EN)'], b['Package / Option']);
    const base = parseFloat(b['Base Price (TND)']) || 0;
    const name = String(b['Client Name'] || '').trim().toLowerCase();

    if (!byService[type]) byService[type] = { sessions: 0, clients: {}, revenue: 0 };
    byService[type].sessions += 1;
    if (name) byService[type].clients[name] = true;
    byService[type].revenue += base;
    totalRevenue += base;

    const loc = classifyLocation(b['Location']);
    visitSplit[loc.kind] = (visitSplit[loc.kind] || 0) + 1;
    if (loc.kind === 'home') homeBands[loc.label] = (homeBands[loc.label] || 0) + 1;
    if (loc.kind === 'city' && loc.city) byCity[loc.city] = (byCity[loc.city] || 0) + base;
  });

  const services = Object.keys(byService).map(function (k) {
    return {
      service: k,
      sessions: byService[k].sessions,
      uniqueClients: Object.keys(byService[k].clients).length,
      revenue: round3(byService[k].revenue)
    };
  }).sort(function (a, b) { return b.revenue - a.revenue; });

  const homeBandList = Object.keys(homeBands).map(function (k) {
    return { band: k, count: homeBands[k] };
  }).sort(function (a, b) { return a.band.localeCompare(b.band); });

  const cityList = Object.keys(byCity).map(function (k) {
    return { city: k, revenue: round3(byCity[k]) };
  }).sort(function (a, b) { return b.revenue - a.revenue; });

  // Client stats (from the Clients sheet, not bookings)
  const cSheet = getClientsSheet();
  const cRows = cSheet.getDataRange().getValues();
  const cH = cRows[0];
  const statusCol = cH.indexOf('Status');
  const createdCol = cH.indexOf('Created At');
  const locCol = cH.indexOf('Location');
  let totalClients = 0, activeClients = 0, completedClients = 0, newInPeriod = 0;
  const clientsByCity = {};
  for (let i = 1; i < cRows.length; i++) {
    if (!cRows[i][cH.indexOf('Client ID')]) continue;
    totalClients += 1;
    const st = String(cRows[i][statusCol] || 'active').toLowerCase();
    if (st === 'completed' || st === 'past') completedClients += 1; else activeClients += 1;
    const city = String(cRows[i][locCol] || '').trim();
    if (city) clientsByCity[city] = (clientsByCity[city] || 0) + 1;
    const created = cRows[i][createdCol];
    if (created) {
      const cd = new Date(created);
      if (cd >= start && cd <= end) newInPeriod += 1;
    }
  }

  const clientCityList = Object.keys(clientsByCity).map(function (k) {
    return { city: k, count: clientsByCity[k] };
  }).sort(function (a, b) { return b.count - a.count; });

  return {
    success: true,
    label: label,
    totalBookings: inRange.length,
    totalRevenue: round3(totalRevenue),
    services: services,
    visitSplit: visitSplit,
    homeBands: homeBandList,
    revenueByCity: cityList,
    clients: {
      total: totalClients,
      active: activeClients,
      completed: completedClients,
      newInPeriod: newInPeriod,
      byCity: clientCityList
    }
  };
}

function getAccountingReport(pin, periodType, anchor) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const range = computePeriodRange(periodType, anchor);
  const start = range.start, end = range.end;

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);
  const inRange = bookings.filter(function (b) {
    const ds = normDate(b['Appointment Date']);
    if (!ds) return false;
    const d = new Date(ds);
    return d >= start && d <= end && String(b['Status']).toLowerCase() !== 'cancelled';
  });

  let caTtc = 0, transport = 0;
  const byService = {};
  inRange.forEach(function (b) {
    const base = parseFloat(b['Base Price (TND)']) || 0;
    const tr = parseFloat(b['Transport Fee (TND)']) || 0;
    caTtc += base; transport += tr;
    const svc = b['Service (EN)'] || 'Other';
    if (!byService[svc]) byService[svc] = { count: 0, ttc: 0 };
    byService[svc].count += 1;
    byService[svc].ttc += base;
  });
  const caHt = r3(caTtc / (1 + TVA_RATE));
  const tvaCollected = r3(caTtc - caHt);

  const serviceBreakdown = Object.keys(byService).map(function (s) {
    const ttc = r3(byService[s].ttc);
    const ht = r3(ttc / (1 + TVA_RATE));
    return { service: s, count: byService[s].count, ttc: ttc, ht: ht, tva: r3(ttc - ht) };
  }).sort(function (a, b) { return b.ttc - a.ttc; });

  const expenses = getExpensesInRange(start, end);
  let expHt = 0, expTva = 0, expTtc = 0;
  const expByCat = {};
  expenses.forEach(function (e) {
    expHt += e.ht; expTva += e.tva; expTtc += e.ttc;
    const c = e.category || 'Autre';
    expByCat[c] = r3((expByCat[c] || 0) + e.ttc);
  });

  return {
    success: true,
    periodType: periodType,
    label: range.label,
    income: {
      caHt: caHt, tvaCollected: tvaCollected, caTtc: r3(caTtc),
      transport: r3(transport), totalReceived: r3(caTtc + transport), bookings: inRange.length
    },
    serviceBreakdown: serviceBreakdown,
    expenses: expenses,
    expenseTotals: { ht: r3(expHt), tva: r3(expTva), ttc: r3(expTtc), byCategory: expByCat },
    tva: { collected: tvaCollected, deductible: r3(expTva), toDeclare: r3(tvaCollected - expTva) },
    resultHt: r3(caHt - expHt)
  };
}

// ───────────────────────── CLIENT IMPORT (from Bookings + Fiches) ─────────────────────────
// Builds a list of "candidate" clients from three sources — existing
// Bookings history, the breastfeeding fiche responses, and the general
// fiche responses — merging entries that share the same phone number so
// the same person doesn't show up three times. Each candidate carries a
// SUGGESTED package, which is intentionally not auto-saved: the dashboard
// shows these as a review list so a wrong guess can be corrected before
// any client profile is actually created.

// NOTE: normPhone lives near the top of this file (with findDuplicateClients) —
// it strips non-digits AND keeps only the last 8, so "+216 12 345 678" and
// "12 345 678" (country code omitted) both normalize to the same value.
// A second, weaker definition used to live here (kept the whole digit string,
// so numbers with vs. without +216 didn't match) — that silently overrode the
// real one for every caller in this file, including duplicate-client detection
// and fiche lookup. Removed; do not redefine normPhone anywhere else.

function getFicheRows(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheets()[0]; // form response sheets only ever have one tab
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { headers: [], rows: [] };
    return { headers: data[0], rows: data.slice(1) };
  } catch (err) {
    return { headers: [], rows: [], error: err.message };
  }
}

function findHeaderIndex(headers, mustInclude) {
  // Fiche headers are long trilingual strings like
  // "Full Name  |  Nom complet  |  ..." — match by substring rather
  // than exact text so minor spacing/punctuation differences don't break it.
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().indexOf(mustInclude.toLowerCase()) !== -1) return i;
  }
  return -1;
}

// Suggests a package ID from the general fiche's "I am registering for"
// checkbox answers (a comma-separated string of checked options).
function suggestPackageFromGeneralFiche(registeringFor) {
  const s = String(registeringFor || '').toLowerCase();
  const hasCbe = s.indexOf('childbirth education') !== -1;
  const hasDoula = s.indexOf('doula support') !== -1;
  const hasPrenatalEx = s.indexOf('prenatal exercise') !== -1;
  const hasPostpartumEx = s.indexOf('postpartum exercise') !== -1;

  if (hasCbe && hasDoula) return { packageId: 'jbm', confidence: 'likely', reason: 'CBE + Doula checked together' };
  if (hasCbe) return { packageId: 'cbe', confidence: 'likely', reason: 'Childbirth Education checked' };
  if (hasDoula) return { packageId: null, confidence: 'needs review', reason: 'Doula support checked alone \u2014 no standalone doula package exists; usually part of Joyful Birth Method' };
  if (hasPrenatalEx || hasPostpartumEx) return { packageId: 'exercise', confidence: 'likely', reason: 'Exercise option checked' };
  return { packageId: null, confidence: 'needs review', reason: 'Could not match checked options to a package' };
}

function getImportCandidates(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const candidatesByPhone = {}; // key -> candidate (key = phone, or 'name:xxx' when no phone)

  // 1. From Bookings — anyone already booked but not yet a Clients profile
  const existingClientPhones = {};
  const existingClientNames = {};
  const clientsSheet = getClientsSheet();
  const clientRows = clientsSheet.getDataRange().getValues();
  const clientHeaders = clientRows[0];
  for (let i = 1; i < clientRows.length; i++) {
    const phone = normPhone(clientRows[i][clientHeaders.indexOf('Phone')]);
    if (phone) existingClientPhones[phone] = true;
    const nm = normName(clientRows[i][clientHeaders.indexOf('Client Name')]);
    if (nm) existingClientNames[nm] = true;
  }

  const bookings = getBookingsData();
  bookings.forEach(b => {
    const phone = normPhone(b['Phone']);
    const nm = normName(b['Client Name']);
    // Backfilled and manually-added bookings often have no phone. Fall back to
    // matching on name so those clients can still be imported — otherwise they
    // would be silently skipped and never appear in the client list.
    const key = phone ? phone : (nm ? 'name:' + nm : '');
    if (!key) return;
    if (phone && existingClientPhones[phone]) return;
    if (!phone && nm && existingClientNames[nm]) return;

    if (!candidatesByPhone[key]) {
      candidatesByPhone[key] = {
        phone: phone,
        rawPhone: b['Phone'] || '',
        name: b['Client Name'] || '',
        email: b['Email'] || '',
        sources: [],
        suggestedPackageId: null,
        suggestedReason: '',
        confidence: 'needs review'
      };
    }
    candidatesByPhone[key].sources.push('Bookings: ' + (b['Service (EN)'] || ''));
    // A booking's own service is a strong signal — map it the same way
    // invoices do, but to a package ID instead of an invoice category.
    const svc = String(b['Service (EN)'] || '').toLowerCase();
    if (!candidatesByPhone[key].suggestedPackageId) {
      if (svc.indexOf('joyful birth method') !== -1) {
        candidatesByPhone[key].suggestedPackageId = 'jbm';
        candidatesByPhone[key].confidence = 'likely';
      } else if (svc.indexOf('childbirth') !== -1) {
        candidatesByPhone[key].suggestedPackageId = 'cbe';
        candidatesByPhone[key].confidence = 'likely';
      } else if (svc.indexOf('breastfeeding') !== -1) {
        candidatesByPhone[key].suggestedPackageId = 'breastfeeding';
        candidatesByPhone[key].confidence = 'likely';
      } else if (svc.indexOf('postpartum') !== -1) {
        candidatesByPhone[key].suggestedPackageId = 'postpartum';
        candidatesByPhone[key].confidence = 'likely';
      } else if (svc.indexOf('pump') !== -1) {
        candidatesByPhone[key].suggestedPackageId = 'pump';
        candidatesByPhone[key].confidence = 'likely';
      } else if (svc.indexOf('exercise') !== -1) {
        candidatesByPhone[key].suggestedPackageId = 'exercise';
        candidatesByPhone[key].confidence = 'likely';
      }
    }
  });

  // 2. From the breastfeeding fiche — every row is unambiguously a
  // Breastfeeding Consultation client.
  const bf = getFicheRows(FICHE_BREASTFEEDING_ID);
  if (bf.headers.length) {
    const nameIdx = findHeaderIndex(bf.headers, 'Full Name');
    const phoneIdx = findHeaderIndex(bf.headers, 'Phone');
    const emailIdx = findHeaderIndex(bf.headers, 'Email');
    bf.rows.forEach(row => {
      const phone = normPhone(row[phoneIdx]);
      if (!phone || existingClientPhones[phone]) return;
      if (!candidatesByPhone[phone]) {
        candidatesByPhone[phone] = {
          phone: phone,
          rawPhone: row[phoneIdx],
          name: row[nameIdx] || '',
          email: row[emailIdx] || '',
          sources: [],
          suggestedPackageId: 'breastfeeding',
          suggestedReason: 'Breastfeeding fiche on file',
          confidence: 'likely'
        };
      }
      candidatesByPhone[phone].sources.push('Breastfeeding fiche');
      if (!candidatesByPhone[phone].name && row[nameIdx]) candidatesByPhone[phone].name = row[nameIdx];
    });
  }

  // 3. From the general fiche — package suggested from checkbox answers
  const gf = getFicheRows(FICHE_GENERAL_ID);
  if (gf.headers.length) {
    const nameIdx = findHeaderIndex(gf.headers, 'Full Name');
    const phoneIdx = findHeaderIndex(gf.headers, 'Phone');
    const emailIdx = findHeaderIndex(gf.headers, 'Email');
    const registeringIdx = findHeaderIndex(gf.headers, 'registering for');
    gf.rows.forEach(row => {
      const phone = normPhone(row[phoneIdx]);
      if (!phone || existingClientPhones[phone]) return;
      const suggestion = suggestPackageFromGeneralFiche(row[registeringIdx]);
      if (!candidatesByPhone[phone]) {
        candidatesByPhone[phone] = {
          phone: phone,
          rawPhone: row[phoneIdx],
          name: row[nameIdx] || '',
          email: row[emailIdx] || '',
          sources: [],
          suggestedPackageId: suggestion.packageId,
          suggestedReason: suggestion.reason,
          confidence: suggestion.confidence
        };
      } else if (!candidatesByPhone[phone].suggestedPackageId && suggestion.packageId) {
        candidatesByPhone[phone].suggestedPackageId = suggestion.packageId;
        candidatesByPhone[phone].suggestedReason = suggestion.reason;
        candidatesByPhone[phone].confidence = suggestion.confidence;
      }
      candidatesByPhone[phone].sources.push('General fiche: ' + (row[registeringIdx] || ''));
      if (!candidatesByPhone[phone].name && row[nameIdx]) candidatesByPhone[phone].name = row[nameIdx];
    });
  }

  // Give every candidate a stable unique key. Phone-less candidates (backfilled
  // or manually-added bookings) would otherwise all share an empty phone and
  // collapse into a single entry in the review UI.
  const candidates = Object.keys(candidatesByPhone)
    .map(function (k) {
      const c = candidatesByPhone[k];
      c.key = k;
      return c;
    })
    .filter(function (c) { return c.name; }); // drop blank rows

  return {
    success: true,
    candidates: candidates,
    packages: PACKAGES,
    ficheErrors: [bf.error, gf.error].filter(Boolean),
    debug: {
      bookingsScanned: bookings.length,
      breastfeedingFicheRows: (bf.rows || []).length,
      generalFicheRows: (gf.rows || []).length,
      existingClients: Object.keys(existingClientPhones).length,
      candidatesFound: candidates.length
    }
  };
}

// Creates client profiles for a batch of reviewed/approved candidates.
// `approvedList` is an array of { name, phone, email, packageId } —
// only what the dashboard's review screen actually confirmed, never the
// raw suggestion data, so a skipped or corrected row can't slip through.
function importApprovedClients(pin, approvedList) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let created = 0;
  const sheet = getClientsSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  approvedList.forEach(c => {
    if (!c.packageId) return; // skip anything left unassigned
    const pkg = PACKAGES.find(p => p.id === c.packageId);
    if (!pkg) return;

    const id = 'C' + Date.now() + Math.floor(Math.random() * 1000);

    const progress = {};
    pkg.components.forEach(comp => {
      if (comp.type === 'sessions') {
        progress[comp.id] = comp.sessions.map(s => ({ id: s.id, completed: false }));
      } else {
        progress[comp.id] = { done: 0, target: comp.target };
      }
    });

    const rowMap = {
      'Client ID': id,
      'Client Name': c.name || '',
      'Phone': c.phone || '',
      'Email': c.email || '',
      'Location': staff.location,
      'Staff': staff.name,
      'Package ID': pkg.id,
      'Package Price (TND)': pkg.price,
      'Custom Price (TND)': '',
      'Progress JSON': JSON.stringify(progress),
      'Status': 'active',
      'Created At': new Date().toISOString(),
      'Notes': 'Imported from Bookings/fiches',
      'Assigned Staff': '',
      'Baby Name': '',
      'Baby Birth Date': '',
      'Baby Weight (kg)': '',
      'Baby Birth Details': ''
    };
    const row = headers.map(function (h) { return (h in rowMap) ? rowMap[h] : ''; });
    sheet.appendRow(row);
    created++;
  });

  return { success: true, created: created };
}

// ───────────────────────── ONBOARDING: single fiche lookup ─────────────────────────
// Searches both fiche sheets for one phone number and returns that person's
// details (name, email, and a suggested package) so the onboarding form can
// pre-fill from a fiche the client already submitted. Unlike the bulk import,
// this is a targeted one-person lookup.
function lookupFicheByPhone(pin, phone) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const target = normPhone(phone);
  if (!target) return { success: true, found: false };

  // Breastfeeding fiche first — any match there is unambiguously a
  // breastfeeding consultation client.
  const bf = getFicheRows(FICHE_BREASTFEEDING_ID);
  if (bf.headers && bf.headers.length) {
    const nameIdx = findHeaderIndex(bf.headers, 'Full Name');
    const phoneIdx = findHeaderIndex(bf.headers, 'Phone');
    const emailIdx = findHeaderIndex(bf.headers, 'Email');
    const babyNameIdx = findHeaderIndex(bf.headers, "Baby's First Name");
    const babyDobIdx = findHeaderIndex(bf.headers, "Baby's Date of Birth");
    for (let r = 0; r < bf.rows.length; r++) {
      const row = bf.rows[r];
      if (normPhone(row[phoneIdx]) !== target) continue;
      return {
        success: true,
        found: true,
        source: 'Breastfeeding fiche',
        name: String(row[nameIdx] || ''),
        email: emailIdx !== -1 ? String(row[emailIdx] || '') : '',
        babyName: babyNameIdx !== -1 ? String(row[babyNameIdx] || '') : '',
        babyBirthDate: babyDobIdx !== -1 ? normDate(row[babyDobIdx]) : '',
        suggestedPackageId: 'breastfeeding'
      };
    }
  }

  // General fiche — suggest a package from the "registering for" checkboxes.
  const gf = getFicheRows(FICHE_GENERAL_ID);
  if (gf.headers && gf.headers.length) {
    const nameIdx = findHeaderIndex(gf.headers, 'Full Name');
    const phoneIdx = findHeaderIndex(gf.headers, 'Phone');
    const emailIdx = findHeaderIndex(gf.headers, 'Email');
    const registeringIdx = findHeaderIndex(gf.headers, 'registering for');
    for (let r = 0; r < gf.rows.length; r++) {
      const row = gf.rows[r];
      if (normPhone(row[phoneIdx]) !== target) continue;
      const suggestion = suggestPackageFromGeneralFiche(registeringIdx !== -1 ? row[registeringIdx] : '');
      return {
        success: true,
        found: true,
        source: 'General fiche',
        name: String(row[nameIdx] || ''),
        email: emailIdx !== -1 ? String(row[emailIdx] || '') : '',
        babyName: '',
        babyBirthDate: '',
        suggestedPackageId: suggestion.packageId,
        suggestionReason: suggestion.reason
      };
    }
  }

  return { success: true, found: false };
}

// ───────────────────────── PER-CLIENT INVOICE ─────────────────────────
// Builds a one-line invoice directly from a client's profile (package +
// custom price), rather than aggregating bookings by date range like the
// Clients Divers invoice does. Used for the occasional personalized
// invoice a specific client needs (insurance, bank deposit, etc).
function generateClientInvoiceData(pin, clientId) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let client = null;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][headers.indexOf('Client ID')]) === String(clientId)) {
      client = rowToClient(rows[i], headers);
      break;
    }
  }
  if (!client) return { success: false, error: 'Client not found' };

  const round3 = (n) => Math.round(n * 1000) / 1000;
  const ttc = client.customPrice !== null ? client.customPrice : client.packagePrice;
  const totalHt = round3(ttc / (1 + TVA_RATE));
  const totalTva = round3(totalHt * TVA_RATE);
  const netAPayer = round3(totalHt + totalTva + TIMBRE_FISCALE);

  const lineItems = [{
    description: client.packageName,
    qty: 1,
    tvaPercent: 19,
    puht: totalHt,
    totalHt: totalHt,
    fullAmount: ttc
  }];

  return {
    success: true,
    periodStart: Utilities.formatDate(new Date(client.createdAt || new Date()), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    periodEnd: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    clientLabel: client.name,
    lineItems: lineItems,
    totalHt: totalHt,
    totalTva: totalTva,
    timbreFiscale: TIMBRE_FISCALE,
    netAPayer: netAPayer,
    fullTotal: ttc,
    amountInWords: amountToFrenchWords(netAPayer)
  };
}

function generateClientInvoiceNumber(pin, clientId, totalTtc) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const sheet = getClientsSheet();
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let clientName = '';
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][headers.indexOf('Client ID')]) === String(clientId)) {
      clientName = rows[i][headers.indexOf('Client Name')];
      break;
    }
  }

  const year = new Date().getFullYear();
  const num = getNextInvoiceNumber();
  const invSheet = getInvoicesSheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  invSheet.appendRow([num, year, clientName, today, today, totalTtc || '', new Date().toISOString(), staff.name]);

  return { success: true, invoiceNumber: num + '/' + year };
}

// ───────────────────────── INVOICE GENERATION (Clients Divers) ─────────────────────────
// Real fiscal invoice matching the business's actual French/Tunisian format:
// HT (pre-tax) backed out from the TTC price charged, 19% TVA, 1.000 TND
// fiscal stamp, sequential numbering that never resets, amount in French
// words. Categories are broader than individual services — see CATEGORY_MAP.

const TVA_RATE = 0.19;
const TIMBRE_FISCALE = 1.000;

// Maps a booking's Service (EN) text to one of the four invoice category
// labels actually used on real invoices. Matching is case-insensitive
// substring matching against the service name.
function mapServiceToCategory(serviceName, packageLabel) {
  const s = String(serviceName || '').toLowerCase();
  const p = String(packageLabel || '').toLowerCase();

  if (s.indexOf('pump') !== -1 || s.indexOf('rental') !== -1) {
    return 'Location de tire-lait';
  }
  if (s.indexOf('breastfeeding') !== -1 || s.indexOf('lactation') !== -1) {
    return 'Consultation a l\u2019allaitement';
  }
  if (s.indexOf('postpartum') !== -1 || s.indexOf('post-partum') !== -1) {
    return 'Services Post-partum';
  }
  if (s.indexOf('childbirth') !== -1 || s.indexOf('joyful birth method') !== -1) {
    return 'Preparations a l\u2019accouchement';
  }
  if (s.indexOf('exercise') !== -1) {
    // Prenatal exercises -> Preparations; postpartum exercises -> Post-partum.
    // Exercise sessions don't carry a prenatal/postpartum flag in the
    // booking data today, so default to Preparations a l'accouchement
    // unless the package label hints at postpartum.
    if (p.indexOf('postpartum') !== -1 || p.indexOf('post-partum') !== -1) {
      return 'Services Post-partum';
    }
    return 'Preparations a l\u2019accouchement';
  }
  // Fallback — surfaced as its own line so nothing silently disappears
  return serviceName || 'Autre';
}

// Gets/creates the Invoices sheet used to store sequential invoice numbers
// and a record of every invoice generated (so numbering never repeats or
// resets, even across sessions).
function getInvoicesSheet() {
  const ss = SpreadsheetApp.openById(DASH_SHEET_ID);
  let sheet = ss.getSheetByName('Invoices');
  if (!sheet) {
    sheet = ss.insertSheet('Invoices');
    sheet.appendRow(['Invoice Number', 'Year', 'Client Label', 'Period Start', 'Period End', 'Total TTC (TND)', 'Generated At', 'Generated By']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sheet;
}

function getNextInvoiceNumber() {
  const sheet = getInvoicesSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const numbers = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(r => parseInt(r[0], 10)).filter(n => !isNaN(n));
  if (numbers.length === 0) return 1;
  return Math.max.apply(null, numbers) + 1;
}

// French amount-in-words for Tunisian dinar invoices (whole dinars +
// millimes). Handles the range realistically needed for this business
// (up to low millions); uses standard French number-word rules including
// "quatre-vingt" and the "et" before "un"/"onze" conventions.
function numberToFrenchWords(n) {
  const units = ['', 'un', 'deux', 'trois', 'quatre', 'cinq', 'six', 'sept', 'huit', 'neuf',
    'dix', 'onze', 'douze', 'treize', 'quatorze', 'quinze', 'seize', 'dix-sept', 'dix-huit', 'dix-neuf'];
  const tens = ['', '', 'vingt', 'trente', 'quarante', 'cinquante', 'soixante', 'soixante-dix', 'quatre-vingt', 'quatre-vingt-dix'];

  function under100(num) {
    if (num < 20) return units[num];
    const t = Math.floor(num / 10);
    const u = num % 10;
    if (t === 7 || t === 9) {
      // soixante-dix / quatre-vingt-dix family: tens word already covers 70/90 base
      if (u === 1 && t !== 8) return tens[t] + ' et ' + units[11 - 10 + u]; // soixante et onze etc — handled via units[11] below
      return tens[t] + (u === 0 ? '' : '-' + units[10 + u]);
    }
    if (u === 0) return tens[t];
    if (u === 1 && t !== 8) return tens[t] + ' et un';
    return tens[t] + '-' + units[u];
  }

  function under1000(num) {
    const h = Math.floor(num / 100);
    const rest = num % 100;
    let str = '';
    if (h === 1) str = 'cent';
    else if (h > 1) str = units[h] + ' cent' + (h > 1 && rest === 0 ? 's' : '');
    if (rest > 0) str += (str ? ' ' : '') + under100(rest);
    return str;
  }

  function chunk(num, singular, plural) {
    if (num === 0) return '';
    if (num === 1) return singular;
    return under1000(num) + ' ' + plural;
  }

  if (n === 0) return 'zero';

  const millions = Math.floor(n / 1000000);
  const thousands = Math.floor((n % 1000000) / 1000);
  const remainder = n % 1000;

  let parts = [];
  if (millions > 0) parts.push(chunk(millions, 'un million', 'millions'));
  if (thousands > 0) parts.push(thousands === 1 ? 'mille' : under1000(thousands) + ' mille');
  if (remainder > 0) parts.push(under1000(remainder));

  return parts.join(' ').trim();
}

function amountToFrenchWords(amountTnd) {
  const dinars = Math.floor(amountTnd);
  const millimes = Math.round((amountTnd - dinars) * 1000);
  let words = numberToFrenchWords(dinars) + ' dinar' + (dinars !== 1 ? 's' : '');
  if (millimes > 0) {
    words += ' et ' + numberToFrenchWords(millimes) + ' millime' + (millimes !== 1 ? 's' : '');
  }
  return words;
}

// Builds the full invoice data structure: groups bookings in the date
// range into the four real categories, calculates HT/TVA backward from
// the TTC total charged, and returns everything the printable invoice
// view needs (no formatting/HTML here — that's the frontend's job).
// `otherServices` is optional: an array of { description, amount } for
// one-off / miscellaneous items entered manually rather than pulled from
// bookings (e.g. a custom arrangement, a small extra fee). The amount
// entered is treated as TTC, same as booking totals, so HT/TVA back out
// the same way.
function generateInvoiceData(pin, periodStart, periodEnd, clientLabel, otherServices, selectedItems) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const round3 = (n) => Math.round(n * 1000) / 1000;
  const categoryTotals = {}; // { categoryLabel: { qty, ttc } }

  // Period range — always defined so the return's date formatting works
  // whether we build from selected items or scan bookings.
  const start = new Date(periodStart);
  start.setHours(0, 0, 0, 0);
  const end = new Date(periodEnd);
  end.setHours(23, 59, 59, 999);

  // If the frontend passed an explicit set of selected line items (from the
  // editable weekly view), build ONLY from those. Otherwise fall back to
  // scanning all bookings in the range (legacy behavior).
  if (selectedItems && selectedItems.length) {
    selectedItems.forEach(it => {
      const ttc = parseFloat(it.base) || 0;
      if (ttc <= 0) return; // never invoice a zero-charge line
      const cat = mapServiceToCategory(it.service, it.packageLabel);
      if (!categoryTotals[cat]) categoryTotals[cat] = { qty: 0, ttc: 0 };
      categoryTotals[cat].qty += 1;
      categoryTotals[cat].ttc += ttc;
    });
  } else {
    let bookings = getBookingsData();
    bookings = filterByScope(bookings, staff);

    const filtered = bookings.filter(b => {
      const dateStr = normDate(b['Appointment Date']);
      if (!dateStr) return false;
      const d = new Date(dateStr);
      return d >= start && d <= end && String(b['Status']).toLowerCase() !== 'cancelled';
    });

    // Group into categories, summing the taxable amount (base price only —
    // transport/delivery fees are excluded) and counting quantity. Skip any
    // booking with no charge so free/included visits don't clutter the invoice.
    filtered.forEach(b => {
      const ttc = parseFloat(b['Base Price (TND)']) || 0;
      if (ttc <= 0) return;
      const cat = mapServiceToCategory(b['Service (EN)'], b['Package / Option']);
      if (!categoryTotals[cat]) categoryTotals[cat] = { qty: 0, ttc: 0 };
      categoryTotals[cat].qty += 1;
      categoryTotals[cat].ttc += ttc;
    });
  }

  // Build line items with HT backed out from TTC, rounded to 3 decimals (millimes)
  const lineItems = Object.keys(categoryTotals).map(cat => {
    const { qty, ttc } = categoryTotals[cat];
    const totalHt = round3(ttc / (1 + TVA_RATE));
    const puht = round3(totalHt / qty);
    return {
      description: cat,
      qty: qty,
      tvaPercent: 19,
      puht: puht,
      totalHt: totalHt
    };
  });

  // Manually-entered "Autres Services" lines — each is its own line item,
  // qty always 1, HT backed out from the entered TTC amount just like
  // every other line so the totals stay consistent.
  (otherServices || []).forEach(item => {
    const ttc = parseFloat(item.amount) || 0;
    if (ttc <= 0) return;
    const totalHt = round3(ttc / (1 + TVA_RATE));
    lineItems.push({
      description: item.description || 'Autres Services',
      qty: 1,
      tvaPercent: 19,
      puht: totalHt,
      totalHt: totalHt
    });
  });

  const totalHt = round3(lineItems.reduce((sum, li) => sum + li.totalHt, 0));
  const totalTva = round3(totalHt * TVA_RATE);
  const netAPayer = round3(totalHt + totalTva + TIMBRE_FISCALE);

  return {
    success: true,
    periodStart: Utilities.formatDate(start, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    periodEnd: Utilities.formatDate(end, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
    clientLabel: clientLabel || 'Clients Divers',
    lineItems: lineItems,
    totalHt: totalHt,
    totalTva: totalTva,
    timbreFiscale: TIMBRE_FISCALE,
    netAPayer: netAPayer,
    amountInWords: amountToFrenchWords(netAPayer)
  };
}

function generateInvoiceNumber(pin, periodStart, periodEnd, clientLabel, totalTtc) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const year = new Date(periodEnd).getFullYear();
  const num = getNextInvoiceNumber();
  const sheet = getInvoicesSheet();

  sheet.appendRow([
    num, year, clientLabel || 'Clients Divers', periodStart, periodEnd,
    totalTtc || '', new Date().toISOString(), staff.name
  ]);

  return { success: true, invoiceNumber: num + '/' + year };
}

// ───────────────────────── MANUAL INVOICE (Generate invoice tab) ─────────────────────────
// A clean, fully manual invoice: the user types a client name and one or
// more service lines (description + price). No bookings are read. The price
// entered is treated as TTC (the all-in price the client pays), exactly like
// the "Autres Services" lines elsewhere, so HT/TVA back out the same way and
// the 1.000 TND timbre is added once. `objet` is the invoice's "Object" line
// (e.g. "Prestation de service mois de juin"). The invoice number is NOT
// auto-assigned here — the user sets it manually and it's recorded via
// recordManualInvoice when they save.
function generateManualInvoiceData(pin, clientLabel, objet, services) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  const round3 = function (n) { return Math.round(n * 1000) / 1000; };
  const lineItems = [];
  (services || []).forEach(function (item) {
    const ttc = parseFloat(item.amount) || 0;
    if (ttc <= 0) return;
    const totalHt = round3(ttc / (1 + TVA_RATE));
    lineItems.push({
      description: item.description || 'Prestation de service',
      qty: 1,
      tvaPercent: 19,
      puht: totalHt,
      totalHt: totalHt
    });
  });

  if (lineItems.length === 0) {
    return { success: false, error: 'Add at least one service with a price.' };
  }

  const totalHt = round3(lineItems.reduce(function (s, li) { return s + li.totalHt; }, 0));
  const totalTva = round3(totalHt * TVA_RATE);
  const netAPayer = round3(totalHt + totalTva + TIMBRE_FISCALE);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  return {
    success: true,
    invoiceDate: today,
    objet: objet || 'Prestation de service',
    clientLabel: clientLabel || '',
    lineItems: lineItems,
    totalHt: totalHt,
    totalTva: totalTva,
    timbreFiscale: TIMBRE_FISCALE,
    netAPayer: netAPayer,
    amountInWords: amountToFrenchWords(netAPayer)
  };
}

// Suggests the next sequential number (read-only — does NOT consume it) so
// the manual invoice field can pre-fill a sensible default the user can edit.
function getSuggestedInvoiceNumber(pin) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  return { success: true, suggested: getNextInvoiceNumber(), year: new Date().getFullYear() };
}

// Records a manual invoice into the Invoices log using the number the user
// typed (not an auto-generated one), keeping the audit trail complete. Flags
// (but does not block) a number that was already used, so a genuine
// duplicate is visible rather than silently slipping through.
function recordManualInvoice(pin, invoiceNumber, clientLabel, totalTtc, objet) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };
  if (!invoiceNumber) return { success: false, error: 'No invoice number provided.' };

  const sheet = getInvoicesSheet();
  const year = new Date().getFullYear();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  let duplicate = false;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); });
    if (existing.indexOf(String(invoiceNumber).trim()) !== -1) duplicate = true;
  }

  sheet.appendRow([
    invoiceNumber, year, clientLabel || '', today, today,
    totalTtc || '', new Date().toISOString(), staff.name + (objet ? ' \u2014 ' + objet : '')
  ]);

  return { success: true, recorded: true, duplicate: duplicate };
}

// The external invoice-log spreadsheet Joy already maintains. Columns (order
// read live from the sheet's header row): Month · Services · Amount ·
// Facture No. · Date range · Invoiced.
const EXPORT_SHEET_ID = '1qsjudkobNhtTHwyeknV1izYN1jpU7vqtbPVaBhpQ2qk';

// Appends one summary row to the external invoice-log sheet. Reads the header
// row first and maps by column NAME (not position), so it lands correctly even
// if the column order differs from what we expect and won't scramble the sheet.
function exportInvoiceToSheet(params) {
  const staff = getStaffByPin(params.pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let ss;
  try {
    ss = SpreadsheetApp.openById(EXPORT_SHEET_ID);
  } catch (e) {
    return { success: false, error: 'Could not open the export spreadsheet. Check that the dashboard has access to it.' };
  }
  const sheet = ss.getSheets()[0]; // first (leftmost) tab
  if (!sheet) return { success: false, error: 'No tab found in the export spreadsheet.' };

  const lastCol = sheet.getLastColumn() || 6;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });

  // Match a value to a column by fuzzy header name.
  const findCol = function (candidates) {
    for (let i = 0; i < headers.length; i++) {
      for (let j = 0; j < candidates.length; j++) {
        if (headers[i].indexOf(candidates[j]) !== -1) return i;
      }
    }
    return -1;
  };

  const colMonth = findCol(['month', 'mois']);
  const colServices = findCol(['service']);
  const colAmount = findCol(['amount', 'montant', 'total']);
  const colFacture = findCol(['facture', 'invoice', 'no']);
  const colRange = findCol(['date range', 'range', 'période', 'periode', 'date']);
  const colInvoiced = findCol(['invoiced', 'facturé', 'facture?', 'paid']);

  const row = new Array(headers.length).fill('');
  const put = function (c, v) { if (c !== -1 && c < row.length) row[c] = v; };
  put(colMonth, params.month || '');
  put(colServices, params.services || '');
  put(colAmount, params.amount || '');
  put(colFacture, params.factureNo || '');
  put(colRange, params.dateRange || '');
  put(colInvoiced, 'Yes');

  sheet.appendRow(row);
  return { success: true, exported: true };
}

// for the generic weekly invoice, grouped by service. Individual clients
// who need their own personalized invoice are handled separately via
// getClientLedger (use that client's data to build a one-off invoice).

function getWeeklyDiversSummary(pin, weekStartStr, endStr) {
  const staff = getStaffByPin(pin);
  if (!staff) return { success: false, error: 'Invalid PIN' };

  let bookings = getBookingsData();
  bookings = filterByScope(bookings, staff);

  const weekStart = new Date(weekStartStr);
  weekStart.setHours(0, 0, 0, 0);
  // Default to a 7-day window, but honor an explicit end date (inclusive) so
  // you can combine two weeks or any custom range into one invoice.
  let weekEnd;
  if (endStr) {
    weekEnd = new Date(endStr);
    weekEnd.setHours(0, 0, 0, 0);
    weekEnd = new Date(weekEnd.getTime() + 24 * 60 * 60 * 1000); // make inclusive
  } else {
    weekEnd = new Date(weekStart.getTime() + 7 * 24 * 60 * 60 * 1000);
  }

  const filtered = bookings.filter(b => {
    const dateStr = normDate(b['Appointment Date']);
    if (!dateStr) return false;
    const d = new Date(dateStr);
    return d >= weekStart && d < weekEnd &&
      String(b['Status']).toLowerCase() !== 'cancelled';
  });

  const lineItems = filtered.map((b, idx) => {
    const base = parseFloat(b['Base Price (TND)']) || 0;
    const total = parseFloat(b['Total (TND)']) || 0;
    return {
      id: 'wk-' + idx,
      date: normDate(b['Appointment Date']),
      clientName: displayName(b['Client Name']),
      service: b['Service (EN)'] || '',
      packageLabel: b['Package / Option'] || '',
      base: base,        // taxable amount (excludes transport)
      total: total,      // client-facing (includes transport)
      hasCharge: base > 0
    };
  }).sort((a, b) => a.date.localeCompare(b.date));

  const totalRevenue = lineItems.reduce((sum, li) => sum + li.total, 0);
  const chargeableCount = lineItems.filter(li => li.hasCharge).length;

  return {
    success: true,
    weekStart: Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    weekEnd: Utilities.formatDate(new Date(weekEnd.getTime() - 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    lineItems: lineItems,
    totalRevenue: totalRevenue,
    count: lineItems.length,
    chargeableCount: chargeableCount
  };
}
