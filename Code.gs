/**
 * Excel / Google Sheets Inventory Management System - Backend Engine
 * Designed according to complete worksheet specifications.
 */

const SHEETS = {
  MAIN_MENU: "Main Menu",
  INVENTORY: "Inventory Database",
  ADD_STOCK: "Add Stock",
  ADJUST_STOCK: "Adjust Stock",
  REMOVE_STOCK: "Remove Stock",
  SUPPLIERS: "Suppliers",
  TRANSACTIONS: "Stock Transactions",
  LOW_STOCK: "Low Stock Report",
  LOW_STOCK_SUPPLIER: "Low Stock by Supplier",
  LOW_STOCK_DATE: "Low Stock by Date",
  SUPPLIER_VIEW: "Supplier View",
  SETTINGS: "Settings"
};

/**
 * Serves the HTML Web App interface
 */
function doGet() {
  initializeSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inventory Control System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Initializes all 12 workbook worksheets with correct structure if they don't exist
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetDefinitions = [
    { name: SHEETS.MAIN_MENU, headers: [] },
    { 
      name: SHEETS.INVENTORY, 
      headers: ["Product ID", "SKU / Item Code", "Product Name", "Category", "Supplier ID", "Supplier Name", "Unit", "Opening Stock", "Total Stock Added", "Total Stock Removed", "Net Adjustments", "Current Stock", "Minimum Stock Level", "Stock Status", "Unit Cost", "Total Stock Value", "Last Updated"] 
    },
    { 
      name: SHEETS.ADD_STOCK, 
      headers: ["Transaction ID", "Date", "Product SKU", "Product Name", "Supplier", "Quantity Added", "Unit Cost", "Total Cost", "Invoice Number", "Destination Location", "Notes", "Entered By"] 
    },
    { 
      name: SHEETS.ADJUST_STOCK, 
      headers: ["Adjustment ID", "Date", "Product SKU", "Product Name", "Location", "System Quantity", "Physical Quantity", "Difference", "Reason", "Approved By", "Status", "Notes"] 
    },
    { 
      name: SHEETS.REMOVE_STOCK, 
      headers: ["Removal ID", "Date", "Product SKU", "Product Name", "Quantity Removed", "Remaining Balance", "Reason", "Department / Customer", "Reference Number", "Notes", "Recorded By"] 
    },
    { 
      name: SHEETS.SUPPLIERS, 
      headers: ["Supplier ID", "Supplier Name", "Contact Person", "Phone Number", "Email Address", "Business Address", "Payment Terms", "Number of Products", "Low-Stock Product Count", "Supplier Status", "Notes"] 
    },
    { 
      name: SHEETS.TRANSACTIONS, 
      headers: ["Transaction ID", "Transaction Date", "Transaction Type", "Product SKU", "Product Name", "Supplier", "Location", "Quantity In", "Quantity Out", "Net Quantity Change", "Reference ID", "Recorded By", "Notes"] 
    },
    { name: SHEETS.LOW_STOCK, headers: [] },
    { name: SHEETS.LOW_STOCK_SUPPLIER, headers: [] },
    { name: SHEETS.LOW_STOCK_DATE, headers: [] },
    { name: SHEETS.SUPPLIER_VIEW, headers: [] },
    { 
      name: SHEETS.SETTINGS, 
      headers: ["Setting Key", "Setting Value", "Description"] 
    }
  ];

  sheetDefinitions.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      if (def.headers.length > 0) {
        sheet.appendRow(def.headers);
        sheet.getRange(1, 1, 1, def.headers.length)
             .setBackground("#1E2D2A")
             .setFontColor("#FFFFFF")
             .setFontWeight("bold");
        sheet.setFrozenRows(1);
      }
    }
  });

  // Ensure default settings exist
  initDefaultSettings();
}

/**
 * Sets default operational settings if missing
 */
function initDefaultSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    const defaultSettings = [
      ["Business Name", "Apex Logistics Ltd.", "Name of company"],
      ["Contact Email", "support@apexlogistics.com", "Main notification email"],
      ["Default Min Stock Level", "10", "Default low stock trigger threshold"],
      ["Allow Negative Stock", "No", "Whether removals can exceed current stock"],
      ["Require Adjustment Reason", "Yes", "Enforce reason field in adjustments"],
      ["Currency Format", "$", "Currency symbol for reports"],
      ["Report Title", "Official Inventory Replenishment Report", "Default header for PDF exports"]
    ];
    defaultSettings.forEach(row => sheet.appendRow(row));
  }
}

/**
 * Helper to parse a sheet into an array of objects
 */
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      let key = h.toString().trim();
      obj[key] = row[i];
    });
    return obj;
  });
}

/**
 * Get Settings map
 */
function getSettings() {
  const data = getSheetData(SHEETS.SETTINGS);
  let settings = {};
  data.forEach(row => {
    if (row["Setting Key"]) {
      settings[row["Setting Key"]] = row["Setting Value"];
    }
  });
  return settings;
}


/** Utility helpers: validation and collision-resistant transaction IDs. */
function makeId(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}
function findInventoryProduct_(sku) {
  const normalized = String(sku || '').trim();
  if (!normalized) return null;
  return getSheetData(SHEETS.INVENTORY).find(p => String(p['SKU / Item Code'] || '').trim() === normalized) || null;
}
function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength || 500);
}

/**
 * DASHBOARD METRICS API
 */
function getDashboardMetrics() {
  const inventory = getSheetData(SHEETS.INVENTORY);
  const suppliers = getSheetData(SHEETS.SUPPLIERS);
  
  let totalProducts = inventory.length;
  let totalUnits = 0;
  let lowStockCount = 0;
  let outOfStockCount = 0;
  let totalValuation = 0;

  inventory.forEach(item => {
    let currentStock = Number(item["Current Stock"]) || 0;
    let minStock = Number(item["Minimum Stock Level"]) || 0;
    let unitCost = Number(item["Unit Cost"]) || 0;

    totalUnits += currentStock;
    totalValuation += (currentStock * unitCost);

    if (currentStock <= 0) {
      outOfStockCount++;
    } else if (currentStock <= minStock) {
      lowStockCount++;
    }
  });

  return {
    totalProducts,
    totalUnits,
    lowStockCount,
    outOfStockCount,
    totalSuppliers: suppliers.length,
    totalValuation: totalValuation.toFixed(2),
    currentDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
  };
}

/**
 * FETCH INVENTORY DATABASE API
 */
function getInventoryDatabase() {
  recalculateAllStock();
  return getSheetData(SHEETS.INVENTORY);
}

/**
 * FETCH SUPPLIERS LIST
 */
function getSuppliersList() {
  return getSheetData(SHEETS.SUPPLIERS);
}

/**
 * RECALCULATE ALL STOCK BALANCES
 * Core Rule: Current Stock = Opening Stock + Added - Removed + Net Adjustments
 */
function recalculateAllStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
  const invData = invSheet.getDataRange().getValues();
  if (invData.length <= 1) return;

  const txData = getSheetData(SHEETS.TRANSACTIONS);

  // Build movement map by SKU
  let movementMap = {};

  txData.forEach(tx => {
    let sku = tx["Product SKU"];
    if (!sku) return;
    if (!movementMap[sku]) {
      movementMap[sku] = { added: 0, removed: 0, adjustments: 0 };
    }

    let qtyIn = Number(tx["Quantity In"]) || 0;
    let qtyOut = Number(tx["Quantity Out"]) || 0;
    let net = Number(tx["Net Quantity Change"]) || 0;
    let type = tx["Transaction Type"];

    if (type === "ADD") {
      movementMap[sku].added += qtyIn;
    } else if (type === "REMOVE") {
      movementMap[sku].removed += qtyOut;
    } else if (type === "ADJUSTMENT") {
      movementMap[sku].adjustments += net;
    }
  });

  // Update Inventory Sheet
  for (let i = 1; i < invData.length; i++) {
    let sku = invData[i][1]; // Column B: SKU
    let opening = Number(invData[i][7]) || 0; // Column H: Opening Stock
    let minStock = Number(invData[i][12]) || 0; // Column M: Min Stock
    let unitCost = Number(invData[i][14]) || 0; // Column O: Unit Cost

    let added = movementMap[sku] ? movementMap[sku].added : 0;
    let removed = movementMap[sku] ? movementMap[sku].removed : 0;
    let adjustments = movementMap[sku] ? movementMap[sku].adjustments : 0;

    let currentStock = opening + added - removed + adjustments;
    let totalValue = currentStock * unitCost;

    let status = "In Stock";
    if (currentStock <= 0) {
      status = "Out of Stock";
    } else if (currentStock <= minStock) {
      status = "Low Stock";
    }

    let now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    // Update cells (1-indexed)
    invSheet.getRange(i + 1, 9).setValue(added);          // Col I: Added
    invSheet.getRange(i + 1, 10).setValue(removed);       // Col J: Removed
    invSheet.getRange(i + 1, 11).setValue(adjustments);   // Col K: Adjustments
    invSheet.getRange(i + 1, 12).setValue(currentStock);  // Col L: Current Stock
    invSheet.getRange(i + 1, 14).setValue(status);        // Col N: Stock Status
    invSheet.getRange(i + 1, 16).setValue(totalValue);    // Col P: Total Value
    invSheet.getRange(i + 1, 17).setValue(now);           // Col Q: Last Updated
  }
}

/**
 * ADD STOCK FUNCTION
 */
function recordAddStock(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Another stock operation is in progress. Please try again.' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    const addSheet = ss.getSheetByName(SHEETS.ADD_STOCK);
    const product = findInventoryProduct_(formData && formData.sku);
    if (!product) return { success: false, message: 'Select a valid product SKU.' };
    const qty = Number(formData.quantity), cost = Number(formData.unitCost);
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isFinite(cost) || cost < 0)
      return { success: false, message: 'Quantity must be greater than zero and unit cost must be zero or higher.' };
    const txId = makeId('ADD');
    const date = cleanText_(formData.date, 20) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const totalCost = qty * cost;
    const user = Session.getActiveUser().getEmail() || 'System User';
    const name = product['Product Name'];
    const supplier = product['Supplier Name'];
    const location = cleanText_(formData.location, 150);
    const notes = cleanText_(formData.notes, 1000);
    addSheet.appendRow([txId, date, product['SKU / Item Code'], name, supplier, qty, cost, totalCost, cleanText_(formData.invoiceNum, 100), location, notes, user]);
    txSheet.appendRow([txId, date, 'ADD', product['SKU / Item Code'], name, supplier, location, qty, 0, qty, txId, user, notes]);
    recalculateAllStock();
    return { success: true, message: 'Stock added successfully. Reference: ' + txId };
  } catch (err) {
    console.error(err);
    return { success: false, message: 'Stock could not be saved. Please check the sheet setup and try again.' };
  } finally { lock.releaseLock(); }
}

/**
 * ADJUST STOCK FUNCTION
 */
function recordAdjustStock(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Another stock operation is in progress. Please try again.' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    const adjSheet = ss.getSheetByName(SHEETS.ADJUST_STOCK);
    const product = findInventoryProduct_(formData && formData.sku);
    if (!product) return { success: false, message: 'Select a valid product SKU.' };
    const physicalQty = Number(formData.physicalQty);
    if (!Number.isFinite(physicalQty) || physicalQty < 0) return { success: false, message: 'Physical quantity must be zero or greater.' };
    const systemQty = Number(product['Current Stock']) || 0;
    const diff = physicalQty - systemQty;
    const settings = getSettings();
    const reason = cleanText_(formData.reason, 250);
    if (String(settings['Require Adjustment Reason'] || 'Yes') === 'Yes' && !reason)
      return { success: false, message: 'Adjustment reason is required by system settings.' };
    const adjId = makeId('ADJ');
    const date = cleanText_(formData.date, 20) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const user = Session.getActiveUser().getEmail() || 'System User';
    const location = cleanText_(formData.location, 150);
    const notes = cleanText_(formData.notes, 1000);
    adjSheet.appendRow([adjId, date, product['SKU / Item Code'], product['Product Name'], location, systemQty, physicalQty, diff, reason, cleanText_(formData.approvedBy, 150) || user, 'Completed', notes]);
    txSheet.appendRow([adjId, date, 'ADJUSTMENT', product['SKU / Item Code'], product['Product Name'], product['Supplier Name'] || '-', location, Math.max(diff, 0), Math.max(-diff, 0), diff, adjId, user, notes]);
    recalculateAllStock();
    return { success: true, message: 'Stock adjustment saved. Reference: ' + adjId };
  } catch (err) {
    console.error(err);
    return { success: false, message: 'Adjustment could not be saved. Please try again.' };
  } finally { lock.releaseLock(); }
}

/**
 * REMOVE STOCK FUNCTION
 */
function recordRemoveStock(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Another stock operation is in progress. Please try again.' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
    const remSheet = ss.getSheetByName(SHEETS.REMOVE_STOCK);
    const product = findInventoryProduct_(formData && formData.sku);
    if (!product) return { success: false, message: 'Select a valid product SKU.' };
    const removeQty = Number(formData.quantity);
    if (!Number.isFinite(removeQty) || removeQty <= 0) return { success: false, message: 'Quantity to remove must be greater than zero.' };
    // Never trust the browser's displayed balance; calculate it from the ledger while locked.
    recalculateAllStock();
    const freshProduct = findInventoryProduct_(formData.sku);
    const currentStock = Number(freshProduct['Current Stock']) || 0;
    const settings = getSettings();
    if (String(settings['Allow Negative Stock'] || 'No') !== 'Yes' && removeQty > currentStock)
      return { success: false, message: 'Insufficient stock. Available: ' + currentStock + '. Negative stock is disabled.' };
    const remId = makeId('REM');
    const date = cleanText_(formData.date, 20) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const remaining = currentStock - removeQty;
    const user = Session.getActiveUser().getEmail() || 'System User';
    const reason = cleanText_(formData.reason, 250);
    if (!reason) return { success: false, message: 'Removal reason is required.' };
    const dept = cleanText_(formData.deptCustomer, 200), refNum = cleanText_(formData.refNum, 100), notes = cleanText_(formData.notes, 1000);
    remSheet.appendRow([remId, date, freshProduct['SKU / Item Code'], freshProduct['Product Name'], removeQty, remaining, reason, dept, refNum, notes, user]);
    txSheet.appendRow([remId, date, 'REMOVE', freshProduct['SKU / Item Code'], freshProduct['Product Name'], freshProduct['Supplier Name'] || '-', '-', 0, removeQty, -removeQty, remId, user, notes]);
    recalculateAllStock();
    return { success: true, message: 'Stock removed successfully. Reference: ' + remId };
  } catch (err) {
    console.error(err);
    return { success: false, message: 'Removal could not be saved. Please try again.' };
  } finally { lock.releaseLock(); }
}

/**
 * LOW STOCK REPORT ENGINE
 */
function getLowStockReport(filterCategory, filterSupplier) {
  recalculateAllStock();
  const inventory = getSheetData(SHEETS.INVENTORY);

  let report = [];
  inventory.forEach(item => {
    let current = Number(item["Current Stock"]) || 0;
    let min = Number(item["Minimum Stock Level"]) || 0;
    let unitCost = Number(item["Unit Cost"]) || 0;

    if (current <= min) {
      if (filterCategory && filterCategory !== "ALL" && item["Category"] !== filterCategory) return;
      if (filterSupplier && filterSupplier !== "ALL" && item["Supplier Name"] !== filterSupplier) return;

      let shortage = min > current ? (min - current) : 1; // Default reorder at least 1 or shortage amount
      let suggestedOrder = shortage + 5; // Suggested reorder cushion
      let estCost = suggestedOrder * unitCost;

      report.push({
        sku: item["SKU / Item Code"],
        name: item["Product Name"],
        category: item["Category"],
        supplier: item["Supplier Name"],
        currentStock: current,
        minStock: min,
        shortage: shortage,
        suggestedOrder: suggestedOrder,
        unitCost: unitCost.toFixed(2),
        estCost: estCost.toFixed(2),
        status: item["Stock Status"]
      });
    }
  });

  return report;
}

/**
 * LOW STOCK BY DATE REPORT ENGINE
 * Reconstructs opening/closing stock for a specific date range
 */
function getLowStockByDateReport(startDateStr, endDateStr) {
  const transactions = getSheetData(SHEETS.TRANSACTIONS);
  const inventory = getSheetData(SHEETS.INVENTORY);

  if (!startDateStr || !endDateStr) return [];
  let start = new Date(startDateStr + 'T00:00:00');
  let end = new Date(endDateStr + 'T23:59:59');
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) return [];

  let reportMap = {};

  inventory.forEach(item => {
    let sku = item["SKU / Item Code"];
    reportMap[sku] = {
      sku: sku,
      name: item["Product Name"],
      supplier: item["Supplier Name"],
      opening: Number(item["Opening Stock"]) || 0,
      added: 0,
      removed: 0,
      adjustments: 0,
      minStock: Number(item["Minimum Stock Level"]) || 0
    };
  });

  transactions.forEach(tx => {
    let sku = tx["Product SKU"];
    if (!reportMap[sku]) return;

    let txDate = new Date(tx["Transaction Date"]);
    let qtyIn = Number(tx["Quantity In"]) || 0;
    let qtyOut = Number(tx["Quantity Out"]) || 0;
    let net = Number(tx["Net Quantity Change"]) || 0;

    if (txDate < start) {
      // Impact opening stock before period
      reportMap[sku].opening += net;
    } else if (txDate >= start && txDate <= end) {
      if (tx["Transaction Type"] === "ADD") reportMap[sku].added += qtyIn;
      if (tx["Transaction Type"] === "REMOVE") reportMap[sku].removed += qtyOut;
      if (tx["Transaction Type"] === "ADJUSTMENT") reportMap[sku].adjustments += net;
    }
  });

  let results = [];
  Object.keys(reportMap).forEach(sku => {
    let item = reportMap[sku];
    let closing = item.opening + item.added - item.removed + item.adjustments;
    let status = closing <= 0 ? "Out of Stock" : (closing <= item.minStock ? "Low Stock" : "In Stock");

    if (closing <= item.minStock) {
      results.push({
        sku: item.sku,
        name: item.name,
        supplier: item.supplier,
        opening: item.opening,
        added: item.added,
        removed: item.removed,
        adjustments: item.adjustments,
        closing: closing,
        minStock: item.minStock,
        status: status
      });
    }
  });

  return results;
}

/**
 * LOAD SAMPLE DEMO DATA
 */
function loadSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Suppliers
  const supSheet = ss.getSheetByName(SHEETS.SUPPLIERS);
  supSheet.getRange(2, 1, supSheet.getLastRow() + 1, 11).clearContent();
  const sampleSuppliers = [
    ["SUP-001", "Global Tech Parts", "John Doe", "555-0192", "johndoe@globaltech.com", "123 Tech Way, CA", "Net 30", 3, 1, "Active", "Primary electronics vendor"],
    ["SUP-002", "Office Depot Co.", "Sarah Smith", "555-0143", "sarah@officedepot.com", "456 Commerce St, NY", "Net 15", 2, 1, "Active", "Stationery supplier"],
    ["SUP-003", "Industrial Steel Inc.", "Mike Johnson", "555-0188", "mike@indsteel.com", "789 Factory Rd, TX", "Net 45", 2, 0, "Active", "Raw materials"]
  ];
  sampleSuppliers.forEach(row => supSheet.appendRow(row));

  // 2. Inventory Database
  const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
  invSheet.getRange(2, 1, invSheet.getLastRow() + 1, 17).clearContent();
  const sampleProducts = [
    ["PRD-001", "SKU-1001", "Wireless Mouse", "Electronics", "SUP-001", "Global Tech Parts", "PCS", 50, 0, 0, 0, 50, 15, "In Stock", 15.00, 750.00, "2026-10-09"],
    ["PRD-002", "SKU-1002", "Mechanical Keyboard", "Electronics", "SUP-001", "Global Tech Parts", "PCS", 8, 0, 0, 0, 8, 10, "Low Stock", 45.00, 360.00, "2026-10-09"],
    ["PRD-003", "SKU-1003", "USB-C Hub", "Electronics", "SUP-001", "Global Tech Parts", "PCS", 0, 0, 0, 0, 0, 5, "Out of Stock", 22.50, 0.00, "2026-10-09"],
    ["PRD-004", "SKU-2001", "Ergonomic Desk Chair", "Furniture", "SUP-002", "Office Depot Co.", "PCS", 12, 0, 0, 0, 12, 5, "In Stock", 120.00, 1440.00, "2026-10-09"],
    ["PRD-005", "SKU-2002", "A4 Copy Paper Box", "Supplies", "SUP-002", "Office Depot Co.", "BOX", 4, 0, 0, 0, 4, 10, "Low Stock", 28.00, 112.00, "2026-10-09"]
  ];
  sampleProducts.forEach(row => invSheet.appendRow(row));

  // Clear Transactions
  const txSheet = ss.getSheetByName(SHEETS.TRANSACTIONS);
  txSheet.getRange(2, 1, txSheet.getLastRow() + 1, 13).clearContent();

  recalculateAllStock();
  return { success: true, message: "Sample demonstration records successfully loaded!" };
}

/**
 * SAVE SETTINGS API
 */
function saveSystemSettings(settingsObject) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
    if (!sheet) throw new Error('Settings sheet missing. Run initializeSheets first.');
    const values = sheet.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < values.length; i++) if (values[i][0]) rowByKey[String(values[i][0])] = i + 1;
    Object.keys(settingsObject || {}).forEach(key => {
      const value = cleanText_(settingsObject[key], 500);
      if (rowByKey[key]) sheet.getRange(rowByKey[key], 2).setValue(value);
      else sheet.appendRow([key, value, 'User updated setting']);
    });
    return { success: true, message: 'System settings saved successfully.' };
  } catch (err) {
    console.error(err);
    return { success: false, message: 'Settings could not be saved. Please try again.' };
  }
}
