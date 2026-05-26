// ============================================================
// Employee Advance & Salary Tracker — Google Apps Script
// ============================================================

const SHEET_ID = '1Pek-5eHD8TyoonvBtRoEpUUGzAbVHz3z_CKjnyOOWzg'; // 🔴 Replace with your Sheet ID

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function getAllRows(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function appendRow(sheetName, obj) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sheet.appendRow(row);
}

function generateID(prefix) {
  return prefix + '_' + new Date().getTime();
}

// ============================================================
// MAIN ENTRY POINT
// ============================================================

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result;

    switch (action) {
      case 'login':            result = login(payload); break;
      case 'getEmployees':     result = getEmployees(payload); break;
      case 'addEmployee':      result = addEmployee(payload); break;
      case 'addAdvance':       result = addAdvance(payload); break;
      case 'addExpense':       result = addExpense(payload); break;
      case 'addSalaryPayment': result = addSalaryPayment(payload); break;
      case 'getEmployeeDues':  result = getEmployeeDues(payload); break;
      case 'getDashboard':     result = getDashboard(payload); break;
      default: result = { success: false, error: 'Unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// AUTH
// ============================================================

function login({ username, password }) {
  const users = getAllRows('Users');
  const user = users.find(u => u.Username === username && u.Password === password);
  if (!user) return { success: false, error: 'Invalid credentials / بيانات غير صحيحة' };
  return { success: true, role: user.Role, site: user.Site, username: user.Username };
}

// ============================================================
// EMPLOYEES
// IqamaNo is the unique identifier — used as primary key across all sheets
// ============================================================

function getEmployees({ role, site }) {
  let employees = getAllRows('Employees').filter(e => e.Status === 'Active');
  if (role === 'site_head') {
    employees = employees.filter(e => e.Site === site);
  }
  return { success: true, data: employees };
}

function addEmployee({ name, iqamaNo, site, salary, addedBy }) {
  if (!iqamaNo) return { success: false, error: 'IqamaNo is required' };

  const employees = getAllRows('Employees');
  const duplicate = employees.find(e => String(e.IqamaNo).trim() === String(iqamaNo).trim());
  if (duplicate) {
    return { success: false, error: 'رقم الإقامة مسجل مسبقاً لـ: ' + duplicate.Name + ' / Iqama already exists for: ' + duplicate.Name };
  }

  appendRow('Employees', {
    IqamaNo: iqamaNo,
    Name: name,
    Site: site,
    Salary: parseFloat(salary),
    Status: 'Active'
  });

  return { success: true, iqamaNo };
}

// ============================================================
// TRANSACTIONS — all keyed by IqamaNo
// ============================================================

function addAdvance({ iqamaNo, amount, notes, addedBy }) {
  appendRow('Advances', {
    ID: generateID('ADV'),
    IqamaNo: iqamaNo,
    Amount: parseFloat(amount),
    Date: new Date().toISOString(),
    AddedBy: addedBy,
    Notes: notes || ''
  });
  return { success: true };
}

function addExpense({ iqamaNo, amount, notes, addedBy }) {
  appendRow('Expenses', {
    ID: generateID('EXP'),
    IqamaNo: iqamaNo,
    Amount: parseFloat(amount),
    Date: new Date().toISOString(),
    AddedBy: addedBy,
    Notes: notes || ''
  });
  return { success: true };
}

function addSalaryPayment({ iqamaNo, amountPaid, month, year, processedBy }) {
  appendRow('SalaryPayments', {
    ID: generateID('SAL'),
    IqamaNo: iqamaNo,
    AmountPaid: parseFloat(amountPaid),
    Month: month,
    Year: year,
    Date: new Date().toISOString(),
    ProcessedBy: processedBy
  });
  return { success: true };
}

// ============================================================
// DUES CALCULATION — keyed by IqamaNo
// Formula: CarryForward + Advances + Expenses - SalaryPaid
// ============================================================

function calculateDues(iqamaNo, month, year) {
  const now = new Date();
  const m = parseInt(month) || (now.getMonth() + 1);
  const y = parseInt(year) || now.getFullYear();

  // Carry forward from previous month
  let prevMonth = m - 1, prevYear = y;
  if (prevMonth === 0) { prevMonth = 12; prevYear = y - 1; }

  const carries = getAllRows('CarryForward');
  const carry = carries.find(c =>
    String(c.IqamaNo).trim() === String(iqamaNo).trim() &&
    parseInt(c.Month) === prevMonth &&
    parseInt(c.Year) === prevYear
  );
  const carryAmount = carry ? parseFloat(carry.CarryAmount) : 0;

  function inMonth(dateStr) {
    const d = new Date(dateStr);
    return d.getMonth() + 1 === m && d.getFullYear() === y;
  }

  const advances = getAllRows('Advances')
    .filter(r => String(r.IqamaNo).trim() === String(iqamaNo).trim() && inMonth(r.Date))
    .reduce((s, r) => s + parseFloat(r.Amount || 0), 0);

  const expenses = getAllRows('Expenses')
    .filter(r => String(r.IqamaNo).trim() === String(iqamaNo).trim() && inMonth(r.Date))
    .reduce((s, r) => s + parseFloat(r.Amount || 0), 0);

  const salaryPaid = getAllRows('SalaryPayments')
    .filter(r =>
      String(r.IqamaNo).trim() === String(iqamaNo).trim() &&
      parseInt(r.Month) === m &&
      parseInt(r.Year) === y
    )
    .reduce((s, r) => s + parseFloat(r.AmountPaid || 0), 0);

  const totalDues = carryAmount + advances + expenses - salaryPaid;

  return { carryForward: carryAmount, advances, expenses, salaryPaid, totalDues, month: m, year: y };
}

function getEmployeeDues({ iqamaNo, month, year }) {
  const employees = getAllRows('Employees');
  const emp = employees.find(e => String(e.IqamaNo).trim() === String(iqamaNo).trim());
  if (!emp) return { success: false, error: 'Employee not found / الموظف غير موجود' };
  const dues = calculateDues(iqamaNo, month, year);
  return { success: true, employee: emp, dues };
}

// ============================================================
// DASHBOARD
// ============================================================

function getDashboard({ month, year, site }) {
  const now = new Date();
  const m = parseInt(month) || (now.getMonth() + 1);
  const y = parseInt(year) || now.getFullYear();

  let employees = getAllRows('Employees').filter(e => e.Status === 'Active');
  if (site && site !== 'ALL') employees = employees.filter(e => e.Site === site);

  const summary = employees.map(emp => {
    const dues = calculateDues(emp.IqamaNo, m, y);
    return { iqamaNo: emp.IqamaNo, name: emp.Name, site: emp.Site, salary: emp.Salary, ...dues };
  });

  const totalOutstanding = summary.reduce((s, e) => s + (e.totalDues > 0 ? e.totalDues : 0), 0);
  const withDues = summary.filter(e => e.totalDues > 0).length;

  return {
    success: true, month: m, year: y, summary,
    stats: { totalOutstanding, totalEmployees: summary.length, withDues }
  };
}

// ============================================================
// MONTH-END CLOSE — triggered automatically on 1st of each month
// Only employees with remaining dues get a CarryForward row
// ============================================================

function closeMonth() {
  const now = new Date();
  let closeMonth = now.getMonth(); // previous month (1-based)
  let closeYear = now.getFullYear();
  if (closeMonth === 0) { closeMonth = 12; closeYear--; }

  const employees = getAllRows('Employees').filter(e => e.Status === 'Active');
  const carries = getAllRows('CarryForward');
  let carried = 0, settled = 0;

  employees.forEach(emp => {
    const dues = calculateDues(emp.IqamaNo, closeMonth, closeYear);
    const exists = carries.find(c =>
      String(c.IqamaNo).trim() === String(emp.IqamaNo).trim() &&
      parseInt(c.Month) === closeMonth &&
      parseInt(c.Year) === closeYear
    );
    if (!exists && dues.totalDues > 0) {
      appendRow('CarryForward', {
        IqamaNo: emp.IqamaNo,
        Name: emp.Name,
        Month: closeMonth,
        Year: closeYear,
        CarryAmount: dues.totalDues,
        CreatedOn: new Date().toISOString()
      });
      carried++;
    } else {
      settled++;
    }
  });

  return { success: true, message: `Closed ${closeMonth}/${closeYear} — Carried: ${carried}, Settled: ${settled}` };
}

// Run ONCE manually from Apps Script editor to install the monthly trigger
function setupMonthlyTrigger() {
  ScriptApp.newTrigger('autoCloseMonth')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();
}

function autoCloseMonth() { closeMonth(); }
