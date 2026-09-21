/**
 * ══════════════════════════════════════════════════════
 * JEWELLERY GST PRO — Google Apps Script Backend
 * ══════════════════════════════════════════════════════
 * 
 * SETUP:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete the default code and paste this entire file
 * 4. Click Deploy → New Deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL
 * 6. Copy your Sheet ID from the URL: docs.google.com/spreadsheets/d/SHEET_ID/edit
 * 7. Paste both into the Jewellery GST Pro Settings page
 */

// ── Sheet Names ──
const SHEETS = {
  companyProfile: 'company profile',
  customers: 'Customers',
  products: 'Products',
  invoices: 'Invoices',
  dcs: 'DeliveryChallans',
  lineOfSales: 'LineOfSales',
  expenses: 'Expenses',
  receipts: 'Receipts',
  payments: 'Payments',
  users: 'Users'
};

// ── Column Headers for each sheet ──
const HEADERS = {
  companyProfile: ['name','pan','gstin','address1','address2','city','state','stateCode','emailId','mobileNo','bankName','ifscCode','accountNo'],
  customers: ['id','name','type','gstin','address','city','state','stateCode','phone','email','createdAt'],
  products: ['id','name','hsn','defaultPurity','makingType','makingCharge','defaultGst','description','createdAt'],
  invoices: ['id','invoiceNumber','invoiceDate','dueDate','invoiceType','customerName','customerGstin','customerAddress','customerState','customerStateCode','customerPhone','taxableAmount','cgstAmount','sgstAmount','igstAmount','grandTotal','itemsJson','notes','losRef','createdBy','createdAt'],
  dcs: ['id','dcNumber','date','dueDate','jobworkerName','jobworkerAddress','purpose','status','totalWeight','itemsJson','notes','createdAt'],
  lineOfSales: ['id','losNumber','date','retailerName','retailerGstin','retailerAddress','retailerState','retailerStateCode','retailerPhone','city','status','totalWeight','totalValue','itemsJson','notes','createdAt'],
  expenses: ['id','voucherNo','date','category','vendor','amount','gstRate','gstAmount','total','paymentMethod','description','createdBy','createdAt'],
  receipts: ['id','receiptNo','date','receivedFrom','amount','mode','towards','description','createdBy','createdAt'],
  payments: ['id','paymentNo','date','paidTo','amount','mode','towards','description','createdBy','createdAt'],
  users: ['id','name','email','username','password','role','createdAt']
};

// ══════════════════════════════════════
// ENTRY POINTS
// ══════════════════════════════════════
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, data, sheetId } = body;
    const ss = SpreadsheetApp.openById(sheetId);
    const result = handleAction(ss, action, data || {});
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, error: err.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ 
    status: 'Jewellery GST Pro API v1.0 — Running' 
  })).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════
// ACTION ROUTER
// ══════════════════════════════════════
function handleAction(ss, action, data) {
  switch (action) {
    // ── Setup ──
    case 'setupSheets':     return setupSheets(ss);
    case 'getAll':          return getAll(ss);

    // ── Company Profile ──
    case 'getCompanyProfile':  return getCompanyProfile(ss);
    case 'saveCompanyProfile': return saveCompanyProfile(ss, data);

    // ── Customers ──
    case 'syncCustomer':    return upsertRow(ss, SHEETS.customers, HEADERS.customers, data);
    case 'deleteCustomer':  return deleteRow(ss, SHEETS.customers, data.id);

    // ── Products ──
    case 'syncProduct':     return upsertRow(ss, SHEETS.products, HEADERS.products, data);
    case 'deleteProduct':   return deleteRow(ss, SHEETS.products, data.id);

    // ── Invoices ──
    case 'syncInvoice':     return upsertRow(ss, SHEETS.invoices, HEADERS.invoices, prepInvoice(data));
    case 'deleteInvoice':   return deleteRow(ss, SHEETS.invoices, data.id);

    // ── Delivery Challans ──
    case 'syncDC':          return upsertRow(ss, SHEETS.dcs, HEADERS.dcs, prepDC(data));
    case 'deleteDC':        return deleteRow(ss, SHEETS.dcs, data.id);

    // ── Line of Sales ──
    case 'syncLOS':         return upsertRow(ss, SHEETS.lineOfSales, HEADERS.lineOfSales, prepLOS(data));
    case 'deleteLOS':       return deleteRow(ss, SHEETS.lineOfSales, data.id);

    // ── Expenses ──
    case 'syncExpense':     return upsertRow(ss, SHEETS.expenses, HEADERS.expenses, data);
    case 'deleteExpense':   return deleteRow(ss, SHEETS.expenses, data.id);

    // ── Receipts ──
    case 'syncReceipt':     return upsertRow(ss, SHEETS.receipts, HEADERS.receipts, data);
    case 'deleteReceipt':   return deleteRow(ss, SHEETS.receipts, data.id);

    // ── Payments ──
    case 'syncPayment':     return upsertRow(ss, SHEETS.payments, HEADERS.payments, data);
    case 'deletePayment':   return deleteRow(ss, SHEETS.payments, data.id);

    // ── Auth ──
    case 'loginUser':       return loginUser(ss, data);
    case 'registerUser':    return registerUser(ss, data);

    // ── Backup (push all from frontend) ──
    case 'backup':          return backupAll(ss, data);

    // ── Invoice PDF Email ──
    case 'sendInvoicePdf':  return sendInvoicePdf(data);

    default:
      return { success: false, error: 'Unknown action: ' + action };
  }
}

// ══════════════════════════════════════
// SETUP — Create all sheets with headers
// ══════════════════════════════════════
function setupSheets(ss) {
  const created = [];
  Object.keys(SHEETS).forEach(key => {
    const name = SHEETS[key];
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      created.push(name);
    }
    // Ensure headers exist
    const headers = HEADERS[key];
    if (headers) {
      const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      const hasHeaders = firstRow[0] === headers[0];
      if (!hasHeaders) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
    }
  });
  return { success: true, created, message: 'All sheets ready (' + created.length + ' new)' };
}

// ══════════════════════════════════════
// GET ALL — Pull entire dataset
// ══════════════════════════════════════
function getAll(ss) {
  const data = {};
  
  // ── Company Profile (single row) ──
  const cpSheet = ss.getSheetByName(SHEETS.companyProfile);
  if (cpSheet && cpSheet.getLastRow() > 1) {
    const rows = cpSheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
    if (rows.length >= 2) {
      const profile = {};
      headers.forEach((h, i) => { profile[h] = rows[1][i]; });
      data.companyProfile = profile;
    }
  }
  
  // ── All other sheets ──
  Object.keys(SHEETS).forEach(key => {
    if (key === 'users' || key === 'companyProfile') return;
    const name = SHEETS[key];
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) {
      data[key] = sheetToObjects(sheet);
    }
  });
  return { success: true, data };
}

// ══════════════════════════════════════
// GENERIC CRUD
// ══════════════════════════════════════

/** Convert sheet data to array of objects using header row as keys */
function sheetToObjects(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }).filter(obj => obj.id); // Skip empty rows
}

/** Upsert: update existing row by ID, or append new row */
function upsertRow(ss, sheetName, headers, data) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet "' + sheetName + '" not found' };
  
  const id = data.id;
  if (!id) return { success: false, error: 'Missing id' };

  // Build row from headers
  const rowData = headers.map(h => {
    // Handle case-insensitive field matching
    const val = data[h] !== undefined ? data[h] : (data[h.toLowerCase()] !== undefined ? data[h.toLowerCase()] : '');
    return val;
  });

  // Find existing row
  const allData = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      found = true;
      break;
    }
  }

  // If not found, append
  if (!found) {
    sheet.appendRow(rowData);
  }

  return { success: true, action: found ? 'updated' : 'created', id };
}

/** Delete row by ID (column A) */
function deleteRow(ss, sheetName, id) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet not found' };
  if (!id) return { success: false, error: 'Missing id' };

  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, deleted: id };
    }
  }
  return { success: true, message: 'ID not found (may already be deleted)' };
}

// ══════════════════════════════════════
// DATA PREP — Serialize items arrays to JSON
// ══════════════════════════════════════
function prepInvoice(data) {
  return {
    ...data,
    itemsJson: data.items ? JSON.stringify(data.items) : (data.itemsJson || '[]')
  };
}

function prepDC(data) {
  return {
    ...data,
    itemsJson: data.items ? JSON.stringify(data.items) : (data.itemsJson || '[]')
  };
}

function prepLOS(data) {
  return {
    ...data,
    itemsJson: data.items ? JSON.stringify(data.items) : (data.itemsJson || '[]')
  };
}

// ══════════════════════════════════════
// COMPANY PROFILE — Read & Write
// ══════════════════════════════════════

/** Read company profile (row 2 of 'company profile' sheet) */
function getCompanyProfile(ss) {
  const sheetName = SHEETS.companyProfile;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Sheet "' + sheetName + '" not found' };
  
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { success: true, profile: null, message: 'No company profile data yet' };
  
  // Read headers from row 1, data from row 2
  const headers = rows[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
  const profile = {};
  headers.forEach((h, i) => { profile[h] = rows[1][i] || ''; });
  
  return { success: true, profile };
}

/** Save company profile (always writes to row 2) */
function saveCompanyProfile(ss, data) {
  const sheetName = SHEETS.companyProfile;
  let sheet = ss.getSheetByName(sheetName);
  
  // Create sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = HEADERS.companyProfile;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  // Read existing headers from the sheet (user may have custom columns)
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerKeys = existingHeaders.map(h => String(h).toLowerCase().replace(/\s+/g, ''));
  
  // Build row matching sheet's own headers
  const rowData = headerKeys.map(h => {
    // Map various header names to our data keys
    const keyMap = {
      'name': data.companyName || data.name || '',
      'companyname': data.companyName || data.name || '',
      'pan': data.pan || '',
      'gstin': data.gstin || '',
      'address1': data.address || data.address1 || '',
      'address': data.address || data.address1 || '',
      'address2': data.address2 || '',
      'city': data.city || '',
      'state': data.state || '',
      'statecode': data.stateCode || '',
      'emailid': data.email || data.emailId || '',
      'email': data.email || data.emailId || '',
      'mobileno': data.phone || data.mobileNo || '',
      'phone': data.phone || data.mobileNo || '',
      'mobile': data.phone || data.mobileNo || '',
      'mobileno': data.phone || data.mobileNo || '',
      'bankname': data.bankName || '',
      'ifsccode': data.ifsc || data.ifscCode || '',
      'ifsc': data.ifsc || data.ifscCode || '',
      'accountno': data.accountNo || '',
      'accountnumber': data.accountNo || '',
    };
    return keyMap[h] !== undefined ? keyMap[h] : (data[h] || '');
  });
  
  // Write to row 2 (always overwrite)
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
  }
  
  return { success: true, message: 'Company profile saved to sheet' };
}

// ══════════════════════════════════════
// AUTH
// ══════════════════════════════════════
function loginUser(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.users);
  if (!sheet || sheet.getLastRow() < 2) return { success: false, error: 'No users registered' };
  
  const users = sheetToObjects(sheet);
  const user = users.find(u => 
    (String(u.email || '').toLowerCase() === String(data.email || '').toLowerCase() || 
     String(u.username || '').toLowerCase() === String(data.email || '').toLowerCase()) && 
    String(u.password) === String(data.password)
  );
  
  if (user) {
    return { success: true, user: { name: user.name, email: user.email, role: user.role || 'user', username: user.username } };
  }
  return { success: false, error: 'Invalid credentials' };
}

function registerUser(ss, data) {
  const sheet = ss.getSheetByName(SHEETS.users);
  if (!sheet) return { success: false, error: 'Users sheet not found' };
  
  const id = 'U' + new Date().getTime();
  const row = [id, data.name || '', data.email || '', data.email || '', data.password || '', 'user', new Date().toISOString()];
  sheet.appendRow(row);
  return { success: true, user: { name: data.name, email: data.email, role: 'user' } };
}

// ══════════════════════════════════════
// BACKUP — Push all data from frontend
// ══════════════════════════════════════
function backupAll(ss, data) {
  const stats = {};
  
  // Process each data type
  const types = [
    { key: 'customers', sheet: SHEETS.customers, headers: HEADERS.customers },
    { key: 'products', sheet: SHEETS.products, headers: HEADERS.products },
    { key: 'invoices', sheet: SHEETS.invoices, headers: HEADERS.invoices, prep: prepInvoice },
    { key: 'dcs', sheet: SHEETS.dcs, headers: HEADERS.dcs, prep: prepDC },
    { key: 'lineOfSales', sheet: SHEETS.lineOfSales, headers: HEADERS.lineOfSales, prep: prepLOS },
    { key: 'expenses', sheet: SHEETS.expenses, headers: HEADERS.expenses },
    { key: 'receipts', sheet: SHEETS.receipts, headers: HEADERS.receipts },
    { key: 'payments', sheet: SHEETS.payments, headers: HEADERS.payments },
  ];

  types.forEach(({ key, sheet: sheetName, headers, prep }) => {
    const items = data[key];
    if (!items || !items.length) { stats[key] = 0; return; }
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { stats[key] = 'sheet not found'; return; }
    
    // Clear existing data (keep headers)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    // Write all rows
    const rows = items.map(item => {
      const processed = prep ? prep(item) : item;
      return headers.map(h => processed[h] !== undefined ? processed[h] : '');
    });
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    stats[key] = rows.length;
  });

  return { success: true, stats, message: 'Full backup complete' };
}

// ══════════════════════════════════════
// SEND INVOICE PDF VIA EMAIL
// ══════════════════════════════════════
function sendInvoicePdf(data) {
  try {
    const htmlContent = data.invoiceHtml || '';
    const emailTo = data.emailTo || '';
    const invNum = data.invoiceNumber || 'Invoice';
    const custName = data.customerName || '';
    const companyName = data.companyName || '';
    const grandTotal = data.grandTotal || 0;

    if (!emailTo || !emailTo.includes('@')) {
      return { success: false, error: 'Invalid email address' };
    }

    // Create PDF blob from HTML
    const blob = HtmlService.createHtmlOutput(htmlContent)
      .getBlob()
      .setName(invNum + '.pdf')
      .getAs('application/pdf');

    // Send email
    const subject = invNum + ' from ' + companyName + ' — ₹' + Number(grandTotal).toLocaleString('en-IN');
    const body = 'Dear ' + custName + ',\n\nPlease find attached invoice ' + invNum + ' for ₹' + Number(grandTotal).toLocaleString('en-IN') + '.\n\nRegards,\n' + companyName;

    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      body: body,
      attachments: [blob],
      name: companyName
    });

    return { success: true, message: 'Invoice PDF sent to ' + emailTo };
  } catch (err) {
    return { success: false, error: 'Email error: ' + err.message };
  }
}
