/**
 * S.S PLUS PHARMACY — FINAL GOOGLE SHEETS APPS SCRIPT
 *
 * Sheets:
 * Dashboard | Entry | Customers | Transactions | Messages | Due Reminders | Settings
 *
 * IMPORTANT:
 * - Customer data stays in Google Sheets.
 * - WhatsApp/SMS links prepare a message; they do not silently send it.
 * - Run setupPharmacy() once after pasting this code.
 */

const CFG = {
  ENTRY: 'Entry',
  DASH: 'Dashboard',
  CUSTOMERS: 'Customers',
  TX: 'Transactions',
  MSG: 'Messages',
  REM: 'Due Reminders',
  SETTINGS: 'Settings',
  HISTORY_START: 40,
  HISTORY_ROWS: 100,
  TZ: 'Asia/Kolkata'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('S.S PLUS PHARMACY')
    .addItem('Setup / Repair System', 'setupPharmacy')
    .addItem('Search Customer', 'searchCustomer')
    .addItem('Save Transaction', 'saveEntry')
    .addItem('Add New Customer', 'addNewCustomer')
    .addItem('Refresh Dashboard', 'refreshAll_')
    .addToUi();
}

function setupPharmacy() {
  const ss = SpreadsheetApp.getActive();
  const required = [
    CFG.DASH, CFG.ENTRY, CFG.CUSTOMERS,
    CFG.TX, CFG.MSG, CFG.REM, CFG.SETTINGS
  ];

  required.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const c = ss.getSheetByName(CFG.CUSTOMERS);
  const t = ss.getSheetByName(CFG.TX);
  const m = ss.getSheetByName(CFG.MSG);
  const r = ss.getSheetByName(CFG.REM);
  const s = ss.getSheetByName(CFG.SETTINGS);
  const e = ss.getSheetByName(CFG.ENTRY);

  ensureHeaders_(c, [
    'Customer ID','Customer Name','Phone','Current Due (₹)',
    'Total Credit (₹)','Total Payment (₹)','Status','WhatsApp'
  ]);

  ensureHeaders_(t, [
    'Transaction ID','Date & Time','Customer ID','Type','Amount (₹)',
    'Notes','Phone','Customer Name','Balance / Due (₹)'
  ]);

  ensureHeaders_(m, [
    'Date & Time','Customer ID','Customer Name','Phone',
    'Message Type','Amount','Due After','WhatsApp','SMS','Message'
  ]);

  ensureHeaders_(r, [
    'Customer ID','Customer Name','Phone','Current Due (₹)',
    'Status','WhatsApp Reminder','Last Transaction'
  ]);

  ensureHeaders_(s, ['SETTING','VALUE']);

  if (!s.getRange('A2').getValue()) {
    s.getRange('A2:B4').setValues([
      ['Pharmacy Name','S.S PLUS PHARMACY'],
      ['WhatsApp Enabled',true],
      ['SMS Enabled',true]
    ]);
  }

  // Safe defaults. Do not delete customer/transaction records.
  e.getRange('B5').setValue('Credit');
  e.getRange('B7').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  e.getRange('E9').setValue('Save first');
  e.getRange('E10').setValue('Save first');
  e.getRange('E13').setValue('Automatic');
  e.getRange('E14').setValue('Ready');

  setDropdown_(e.getRange('B5'), ['Credit','Payment']);
  setDropdown_(e.getRange('A10'), ['SAVE']);
  setDropdown_(e.getRange('A16'), ['ADD']);

  refreshAll_();
  ss.toast('System setup / repair completed.', 'S.S PLUS PHARMACY', 5);
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.ENTRY) return;

  const cell = e.range.getA1Notation();
  const value = String(e.value || '').trim();

  if (['B4','B5','B6','B8'].includes(cell)) {
    if (String(sh.getRange('B4').getDisplayValue()).trim()) {
      sh.getRange('B7')
        .setValue(new Date())
        .setNumberFormat('dd/MM/yyyy HH:mm:ss');
    }
  }

  if (cell === 'B4') {
    searchCustomer();
    return;
  }

  if (cell === 'B13' || cell === 'B14') {
    duplicateCheck_();
    return;
  }

  if (cell === 'A10' && value === 'SAVE') {
    saveEntry();
    sh.getRange('A10').clearContent();
    return;
  }

  if (cell === 'A16' && value === 'ADD') {
    addNewCustomer();
    sh.getRange('A16').clearContent();
  }
}

function saveEntry() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    toast_('Another save is already running. Please wait.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActive();
    const e = ss.getSheetByName(CFG.ENTRY);
    const t = ss.getSheetByName(CFG.TX);
    const m = ss.getSheetByName(CFG.MSG);

    const query = String(e.getRange('B4').getDisplayValue()).trim();
    const type = String(e.getRange('B5').getDisplayValue()).trim();
    const amount = Number(e.getRange('B6').getValue());
    const notes = String(e.getRange('B8').getDisplayValue()).trim();

    if (!query) return toast_('Customer Name বা Phone দিন.');
    if (!['Credit','Payment'].includes(type))
      return toast_('Credit অথবা Payment নির্বাচন করুন.');
    if (!(amount > 0))
      return toast_('Amount 0-এর বেশি দিন.');

    const customer = findCustomer_(query);

    if (!customer)
      return toast_('Customer পাওয়া যায়নি. আগে ADD NEW CUSTOMER করুন.');

    const now = new Date();

    // Recalculate balance from the ledger.
    let balance = 0;
    const rows = t.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]) === String(customer.id)) {
        const a = Number(rows[i][4]) || 0;
        balance += rows[i][3] === 'Credit' ? a : -a;
      }
    }

    balance += type === 'Credit' ? amount : -amount;

    const txId = 'TX-' + Utilities.getUuid().slice(0, 8).toUpperCase();

    t.appendRow([
      txId,
      now,
      customer.id,
      type,
      amount,
      notes,
      customer.phone,
      customer.name,
      balance
    ]);

    updateCustomer_(customer.id);

    const updated = findCustomer_(customer.id);
    const finalDue = updated ? updated.due : balance;

    loadCustomer_(customer.id);

    const pharmacy =
      String(getSetting_('Pharmacy Name') || 'S.S PLUS PHARMACY');

    const msg = type === 'Credit'
      ? `Dear ${customer.name}, your credit of ₹${amount.toFixed(2)} has been recorded. Current due: ₹${finalDue.toFixed(2)}. - ${pharmacy}`
      : `Dear ${customer.name}, your payment of ₹${amount.toFixed(2)} has been received. Current due: ₹${finalDue.toFixed(2)}. - ${pharmacy}`;

    const phone = normalizePhone_(customer.phone);

    const wa = phone
      ? 'https://wa.me/' + phone + '?text=' + encodeURIComponent(msg)
      : '';

    const sms = phone
      ? 'sms:' + phone + '?body=' + encodeURIComponent(msg)
      : '';

    if (wa) {
      e.getRange('E9').setFormula(
        '=HYPERLINK("' + wa + '","WhatsApp")'
      );
    } else {
      e.getRange('E9').setValue('No phone');
    }

    if (sms) {
      e.getRange('E10').setFormula(
        '=HYPERLINK("' + sms + '","SMS")'
      );
    } else {
      e.getRange('E10').setValue('No phone');
    }

    m.appendRow([
      now,
      customer.id,
      customer.name,
      customer.phone,
      type,
      amount,
      finalDue,
      wa,
      sms,
      msg
    ]);

    // Keep customer selected; clear only transaction inputs.
    e.getRange('B6').clearContent();
    e.getRange('B8').clearContent();

    refreshAll_();

    ss.toast(
      'Transaction saved. Customer balance and history updated.',
      'S.S PLUS PHARMACY',
      5
    );

  } finally {
    lock.releaseLock();
  }
}

function addNewCustomer() {
  const ss = SpreadsheetApp.getActive();
  const e = ss.getSheetByName(CFG.ENTRY);
  const c = ss.getSheetByName(CFG.CUSTOMERS);

  const name = String(e.getRange('B13').getDisplayValue()).trim();
  const phone = normalizePhone_(e.getRange('B14').getDisplayValue());

  if (!name) return toast_('Customer Name দিন.');
  if (!phone) return toast_('Phone Number দিন.');

  if (findCustomer_(phone) || findCustomer_(name)) {
    e.getRange('E13:E14').setValues([
      ['Duplicate'],
      ['Already exists']
    ]);
    return toast_('এই customer আগে থেকেই আছে.');
  }

  const id =
    'CUS-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  c.appendRow([
    id,
    name,
    phone,
    0,
    0,
    0,
    'CLEAR',
    'https://wa.me/' + phone
  ]);

  e.getRange('B4').setValue(name);
  e.getRange('B13:B14').clearContent();

  e.getRange('E13:E14').setValues([
    ['Automatic'],
    ['Added']
  ]);

  searchCustomer();
  refreshAll_();

  ss.toast(
    'New customer added successfully.',
    'S.S PLUS PHARMACY',
    5
  );
}

function searchCustomer() {
  const e = SpreadsheetApp.getActive().getSheetByName(CFG.ENTRY);
  const q = String(e.getRange('B4').getDisplayValue()).trim();

  if (!q) {
    clearCustomer_();
    return;
  }

  const customer = findCustomer_(q);

  if (!customer) {
    clearCustomer_();
    e.getRange('E14').setValue('Not found');
    return;
  }

  loadCustomer_(customer.id);
}

function loadCustomer_(id) {
  const e = SpreadsheetApp.getActive().getSheetByName(CFG.ENTRY);
  const customer = findCustomer_(id);

  if (!customer) return;

  e.getRange('E4:E6').setValues([
    [customer.id],
    [customer.phone],
    [customer.due]
  ]);

  e.getRange('E6').setNumberFormat('₹#,##0.00');

  if (!e.getRange('B7').getValue()) {
    e.getRange('B7')
      .setValue(new Date())
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  e.getRange(
    CFG.HISTORY_START,
    1,
    CFG.HISTORY_ROWS,
    7
  ).clearContent();

  const t = SpreadsheetApp.getActive().getSheetByName(CFG.TX);
  const rows = t.getDataRange().getValues();
  const history = [];

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) === String(customer.id)) {
      const bal = Number(rows[i][8]) || 0;

      history.push([
        rows[i][1],
        rows[i][3],
        rows[i][2],
        rows[i][4],
        bal,
        rows[i][5],
        bal > 0 ? 'DUE' : bal < 0 ? 'ADVANCE' : 'CLEAR'
      ]);
    }
  }

  history.sort((a,b) => new Date(a[0]) - new Date(b[0]));

  const recent = history.slice(-CFG.HISTORY_ROWS);

  if (recent.length) {
    e.getRange(
      CFG.HISTORY_START,
      1,
      recent.length,
      7
    ).setValues(recent);

    e.getRange(
      CFG.HISTORY_START,
      1,
      recent.length,
      1
    ).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    e.getRange(
      CFG.HISTORY_START,
      4,
      recent.length,
      2
    ).setNumberFormat('₹#,##0.00');
  }
}

function findCustomer_(query) {
  const qRaw = String(query || '').trim();
  if (!qRaw) return null;

  const qPhone = normalizePhone_(qRaw).toLowerCase();

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CFG.CUSTOMERS);

  const rows = sh.getDataRange().getValues();

  // Exact match first.
  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '');
    const name = String(rows[i][1] || '');
    const phone = normalizePhone_(rows[i][2] || '');

    if (
      id === qRaw ||
      phone === qPhone ||
      name.toLowerCase() === qRaw.toLowerCase()
    ) {
      return customerObj_(rows, i);
    }
  }

  // Partial name / phone search.
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][1] || '');
    const phone = normalizePhone_(rows[i][2] || '');

    if (
      name.toLowerCase().includes(qRaw.toLowerCase()) ||
      (qPhone && phone.includes(qPhone))
    ) {
      return customerObj_(rows, i);
    }
  }

  return null;
}

function customerObj_(rows, i) {
  return {
    row: i + 1,
    id: String(rows[i][0] || ''),
    name: String(rows[i][1] || ''),
    phone: String(rows[i][2] || ''),
    due: Number(rows[i][3]) || 0
  };
}

function updateCustomer_(id) {
  const ss = SpreadsheetApp.getActive();
  const c = ss.getSheetByName(CFG.CUSTOMERS);
  const t = ss.getSheetByName(CFG.TX);

  const tx = t.getDataRange().getValues();

  let credit = 0;
  let payment = 0;
  let phone = '';

  for (let i = 1; i < tx.length; i++) {
    if (String(tx[i][2]) === String(id)) {
      const amount = Number(tx[i][4]) || 0;

      if (String(tx[i][3]) === 'Credit') credit += amount;
      if (String(tx[i][3]) === 'Payment') payment += amount;

      phone = String(tx[i][6] || phone);
    }
  }

  const due = credit - payment;
  const status =
    due > 0 ? 'DUE' :
    due < 0 ? 'ADVANCE' :
    'CLEAR';

  const rows = c.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const p = phone || String(rows[i][2] || '');

      c.getRange(i + 1, 3, 1, 6).setValues([[
        p,
        due,
        credit,
        payment,
        status,
        p ? 'https://wa.me/' + normalizePhone_(p) : ''
      ]]);

      c.getRange(i + 1, 4, 1, 3)
        .setNumberFormat('₹#,##0.00');

      return;
    }
  }
}

function refreshAll_() {
  updateAllCustomers_();
  updateDashboard_();
  updateReminders_();
}

function updateAllCustomers_() {
  const c = SpreadsheetApp
    .getActive()
    .getSheetByName(CFG.CUSTOMERS);

  const rows = c.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const id = String(rows[i][0] || '').trim();
    if (id) updateCustomer_(id);
  }
}

function updateDashboard_() {
  const ss = SpreadsheetApp.getActive();
  const d = ss.getSheetByName(CFG.DASH);
  const c = ss.getSheetByName(CFG.CUSTOMERS);
  const t = ss.getSheetByName(CFG.TX);

  if (!d || !c || !t) return;

  const customers = c.getDataRange().getValues();
  const transactions = t.getDataRange().getValues();

  let customerCount = 0;
  let totalCredit = 0;
  let totalPayment = 0;
  let currentDue = 0;
  let dueCustomers = 0;
  let clearCustomers = 0;
  let advanceCustomers = 0;
  let transactionCount = 0;

  for (let i = 1; i < customers.length; i++) {
    if (String(customers[i][1] || '').trim()) {
      customerCount++;

      const due = Number(customers[i][3]) || 0;
      currentDue += due;

      if (due > 0) dueCustomers++;
      else if (due < 0) advanceCustomers++;
      else clearCustomers++;
    }
  }

  for (let i = 1; i < transactions.length; i++) {
    if (String(transactions[i][0] || '').trim()) {
      transactionCount++;

      const amount = Number(transactions[i][4]) || 0;

      if (String(transactions[i][3]) === 'Credit')
        totalCredit += amount;

      if (String(transactions[i][3]) === 'Payment')
        totalPayment += amount;
    }
  }

  d.getRange('A4').setValue(customerCount);
  d.getRange('C4').setValue(totalCredit);
  d.getRange('E4').setValue(totalPayment);
  d.getRange('G4').setValue(currentDue);

  d.getRange('A9').setValue(dueCustomers);
  d.getRange('C9').setValue(clearCustomers);
  d.getRange('E9').setValue(advanceCustomers);
  d.getRange('G9').setValue(transactionCount);

  d.getRangeList(['C4','E4','G4'])
    .setNumberFormat('₹#,##0.00');
}

function updateReminders_() {
  const ss = SpreadsheetApp.getActive();
  const c = ss.getSheetByName(CFG.CUSTOMERS);
  const r = ss.getSheetByName(CFG.REM);
  const t = ss.getSheetByName(CFG.TX);

  const customers = c.getDataRange().getValues();
  const transactions = t.getDataRange().getValues();

  if (r.getMaxRows() > 1) {
    r.getRange(
      2,
      1,
      r.getMaxRows() - 1,
      7
    ).clearContent();
  }

  const output = [];

  for (let i = 1; i < customers.length; i++) {
    const due = Number(customers[i][3]) || 0;

    if (due > 0) {
      let last = '';

      for (let j = 1; j < transactions.length; j++) {
        if (String(transactions[j][2]) === String(customers[i][0])) {
          last = transactions[j][1];
        }
      }

      const phone = normalizePhone_(customers[i][2]);

      const msg =
        `Dear ${customers[i][1]}, your current due is ₹${due.toFixed(2)}. ` +
        `Please make the payment. - ` +
        `${getSetting_('Pharmacy Name') || 'S.S PLUS PHARMACY'}`;

      const wa = phone
        ? 'https://wa.me/' + phone + '?text=' + encodeURIComponent(msg)
        : '';

      output.push([
        customers[i][0],
        customers[i][1],
        customers[i][2],
        due,
        'DUE',
        wa,
        last
      ]);
    }
  }

  if (output.length) {
    r.getRange(
      2,
      1,
      output.length,
      7
    ).setValues(output);

    r.getRange(
      2,
      4,
      output.length,
      1
    ).setNumberFormat('₹#,##0.00');
  }
}

function duplicateCheck_() {
  const e = SpreadsheetApp
    .getActive()
    .getSheetByName(CFG.ENTRY);

  const name = String(
    e.getRange('B13').getDisplayValue()
  ).trim();

  const phone = normalizePhone_(
    e.getRange('B14').getDisplayValue()
  );

  if (!name && !phone) {
    e.getRange('E13:E14').setValues([
      ['Automatic'],
      ['Ready']
    ]);
    return;
  }

  const found =
    (phone && findCustomer_(phone)) ||
    (name && findCustomer_(name));

  e.getRange('E13:E14').setValues([
    ['Automatic'],
    [found ? 'Duplicate found' : 'Available']
  ]);
}

function clearCustomer_() {
  const e = SpreadsheetApp
    .getActive()
    .getSheetByName(CFG.ENTRY);

  e.getRange('E4:E6').clearContent();

  e.getRange(
    CFG.HISTORY_START,
    1,
    CFG.HISTORY_ROWS,
    7
  ).clearContent();

  e.getRange('E9:E10').setValues([
    ['Save first'],
    ['Save first']
  ]);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
    return;
  }

  const existing = sheet
    .getRange(1, 1, 1, headers.length)
    .getValues()[0];

  const blank = existing.every(v => String(v).trim() === '');

  if (blank) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }
}

function setDropdown_(cell, values) {
  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();

  cell.setDataValidation(rule);
}

function getSetting_(key) {
  const s = SpreadsheetApp
    .getActive()
    .getSheetByName(CFG.SETTINGS);

  const rows = s.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(key))
      return rows[i][1];
  }

  return '';
}

function normalizePhone_(phone) {
  let s = String(phone || '').replace(/\D/g, '');

  // India: 10-digit number -> 91 + number.
  if (s.length === 10) s = '91' + s;

  return s;
}

function toast_(message) {
  SpreadsheetApp
    .getActive()
    .toast(message, 'S.S PLUS PHARMACY', 4);
}
