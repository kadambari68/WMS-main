/**
 * COMPLETE WAREHOUSE MANAGEMENT SYSTEM - CODE.GS
 * 
 * V2 ARCHITECTURE - CONSOLIDATED MASTER FILE
 * 
 * Includes:
 * 1. Configuration & Schema
 * 2. Security & Auth Middleware
 * 3. Caching & Helpers
 * 4. Transaction Engine (Inward, Dispatch, Consumption, Returns, TRANSFER)
 * 5. Production Ledger Logic
 * 6. Dashboard Data APIs
 * 7. Web App Routing (doGet)
 */

// !!! EMERGENCY BYPASS - SET TO FALSE FOR PROD !!!
const GLOBAL_DEV_BYPASS = false;

// =====================================================
// 1. CONFIGURATION & SCHEMA - V2
// =====================================================

const CONFIG = {
  SHEETS: {
    MOVEMENT: 'Movement_Audit_Log',
    INVENTORY: 'Inventory_Balance',
    UOM_MASTER: 'UOM_Master',
    BIN: 'Bin_Master',
    RACK: 'Rack_Master',
    ITEM: 'Item_Master',
    BATCH: 'Batch_Master',
    ITEM_GROUP: 'Item_Group_Master',
    INWARD_FORM: 'Form_Inward_Receipt',
    TRANSFER_FORM: 'Form_Inter_Unit_Transfer',
    DISPATCH_FORM: 'Form_Dispatch',
    // V2 New Sheets
    ACCESS_CONTROL: 'Access_Control_List',
    PRODUCTION_LEDGER: 'Production_Order_Ledger',
    PRODUCTION_TRANSFERS: 'production_transfers',
    INVENTORY_MOVEMENT_ARCHIVE: 'inventory_movement_archive',
    QA_EVENTS: 'QA_Events',
    DEAD_STOCK: 'Dead_Stock_Log',
    LOCATION_MASTER: 'Location_master',
    AI_KNOWLEDGE: 'WMS_AI_Knowledge',
    AI_QUERY_LOG: 'AI_Query_Log',
    PACKING_VERSION_MASTER: 'Packing_Version_Master'
  },

  ITEM_COLS: {
    // Current Item_Master schema:
    // item_id, item_code, item_name, item_group_code, uom_code,
    // min_stock_level, is_versioning_enabled, default_quality_days, status
    ITEM_ID: 0, ITEM_CODE: 1, ITEM_NAME: 2,
    ITEM_DESCRIPTION: 2, ITEM_TYPE: 3, ITEM_CATEGORY_CODE: 3, ITEM_GROUP_CODE: 3,
    UOM_CODE: 4, WEIGHT_UOM_CODE: 4, COSTING_METHOD: 9,
    STANDARD_MATERIAL_COST: 10, STATUS: 8, MIN_STOCK_LEVEL: 5
  },

  BATCH_COLS: {
    BATCH_ID: 0, ITEM_CODE: 1, BATCH_NUMBER: 2, BATCH_EXP_DATE: 3,
    LOT_NUMBER: 4, GIN_NO: 5, QUALITY_DATE: 6, QUALITY_STATUS: 7,
    VERSION: 8, VERSION_PARENT_ID: 9, IS_VERSION: 10
  },

  // V2 INVENTORY SCHEMA (The Source of Truth)
  INVENTORY_COLS: {
    INVENTORY_ID: 0,
    ITEM_ID: 1,
    ITEM_CODE_CACHE: 2, // Derived
    BATCH_ID: 3,
    GIN_NO: 4,          // New
    VERSION: 5,         // New
    QUALITY_STATUS: 6,  // New
    QUALITY_DATE: 7,    // New
    BIN_ID: 8,
    SITE: 9,            // Derived
    LOCATION: 10,       // Derived
    TOTAL_QUANTITY: 11, // Truth
    UOM: 12,
    LOT_NO: 13,
    EXPIRY_DATE: 14,
    LAST_UPDATED: 15,
    INWARD_DATE: 16,
    LAST_TRANSFER_DATE: 17
  },

  // V2 MOVEMENT LOG SCHEMA
  MOVEMENT_COLS: {
    MOVEMENT_ID: 0,
    TIMESTAMP: 1,
    MOVEMENT_TYPE: 2,
    ITEM_ID: 3,
    BATCH_ID: 4,
    VERSION: 5,         // New
    GIN_NO: 6,          // New
    PROD_ORDER_REF: 7,  // New
    FROM_BIN_ID: 8,
    TO_BIN_ID: 9,
    QUANTITY: 10,
    UOM: 11,
    QUALITY_STATUS: 12, // New
    USER_EMAIL: 13,
    REMARKS: 14
  },

  // QA EVENTS LOG SCHEMA (audit only - no stock math)
  QA_EVENTS_COLS: {
    EVENT_ID: 0,
    TIMESTAMP: 1,
    INVENTORY_ID: 2,
    ITEM_CODE: 3,
    BATCH_ID: 4,
    BIN_ID: 5,
    PREV_STATUS: 6,
    NEW_STATUS: 7,
    ACTION: 8,
    OVERRIDE_REASON: 9,
    OVERRIDDEN_BY: 10,
    OVERRIDDEN_AT: 11,
    MOVEMENT_ID: 12,
    REMARKS: 13,
    QUANTITY: 14
  },

  // V2 PRODUCTION LEDGER SCHEMA
  PRODUCTION_LEDGER_COLS: {
    ORDER_ID: 0,
    ITEM_ID: 1,
    BATCH_ID: 2,
    VERSION: 3,
    QTY_REQUESTED: 4, // NEW
    QTY_ISSUED: 5,
    QTY_RETURNED: 6,
    NET_OUTSTANDING: 7,
    QTY_REJECTED: 8,
    STATUS: 9,
    LAST_UPDATED: 10
  },

  // Production transfers (audit bridge for production flows)
  PRODUCTION_TRANSFERS_COLS: {
    TRANSFER_ID: 0,
    TRANSFER_DATE: 1,
    TRANSFER_TYPE: 2,
    PRODUCTION_ORDER_NO: 3,
    PRODUCTION_AREA: 4,
    ITEM_CODE: 5,
    ITEM_ID: 6,
    BATCH_NUMBER: 7,
    LOT_NO: 8,
    QUANTITY: 9,
    UOM: 10,
    FROM_SITE: 11,
    FROM_LOCATION: 12,
    FROM_BIN_ID: 13,
    FROM_BIN_CODE: 14,
    TO_SITE: 15,
    TO_LOCATION: 16,
    TO_BIN_ID: 17,
    TO_BIN_CODE: 18,
    RETURNED_BY_NAME: 19,
    CREATED_BY: 20,
    CREATED_AT: 21,
    REMARKS: 22,
    STATUS: 23
  },

  INVENTORY_MOVEMENT_ARCHIVE_COLS: {
    MOVEMENT_ID: 0,
    MOVEMENT_DATE: 1,
    ITEM_ID: 2,
    BATCH_ID: 3,
    FROM_SITE: 4,
    FROM_LOCATION: 5,
    FROM_RACK: 6,
    FROM_BIN: 7,
    TO_SITE: 8,
    TO_LOCATION: 9,
    TO_RACK: 10,
    TO_BIN: 11,
    QUANTITY: 12,
    UOM: 13,
    MOVEMENT_TYPE: 14,
    REMARKS: 15,
    CREATED_BY: 16
  },

  DEAD_STOCK_COLS: {
    DEAD_STOCK_ID: 0,
    TIMESTAMP: 1,
    PROD_ORDER_REF: 2,
    ITEM_ID: 3,
    ITEM_CODE: 4,
    BATCH_ID: 5,
    QUANTITY: 6,
    UOM: 7,
    REASON: 8,
    REMARKS: 9,
    USER_EMAIL: 10
  },

  PACKING_VERSION_COLS: {
    MAPPING_ID: 0,
    ITEM_CODE: 1,
    SYSTEM_VERSION: 2,
    BIN_ID: 3,
    IS_LATEST: 4,
    REASON: 5,
    UPDATED_BY: 6,
    UPDATED_AT: 7,
    LATEST_NAME: 8
  },

  MOVEMENT_TYPES: {
    INWARD: 'INWARD_RECEIPT',
    INWARD_REVERSAL: 'INWARD_REVERSAL',
    DISPATCH: 'DISPATCH',
    CONSUMPTION: 'PROD_CONSUMPTION',
    RETURN_PROD: 'PROD_RETURN',
    RETURN_DEAD: 'PROD_RETURN_DEAD',
    TRANSFER: 'TRANSFER',
    TRANSFER_QUARANTINE: 'TRANSFER_QUARANTINE',
    QA_HOLD: 'QA_HOLD',
    QA_APPROVED: 'QA_APPROVED'
  },
  QA: {
    ALLOW_TRANSFER_BEFORE_QA: false
  }
};

// =====================================================
// SIMPLE STOCK LOOKUP — used by TransferForm picker
// Reads Inventory_Balance directly by item_code_cache.
// Returns ONLY plain strings/numbers — no Date objects,
// no rowData. Safe for google.script.run serialization.
// =====================================================

/**
 * getStockByItemCode
 * Called by TransferForm when user enters an item code.
 * Scans Inventory_Balance col 2 (item_code_cache) for exact match.
 * Also reads Bin_Master col 0/2 to resolve bin_code from bin_id.
 *
 * @param  {string} itemCode  e.g. "ST27"
 * @returns {Array}  [{batchId, binId, binCode, site, location, qty, uom, qaStatus, version}]
 */
/**
 * getStockByItemCode
 * Lean inventory lookup for TransferForm source picker.
 * Returns ONLY plain primitives — no Date objects, no raw arrays.
 * Safe for google.script.run serialization.
 * Uses hardcoded column indices that match INVENTORY_COLS exactly.
 *
 * @param  {string} itemCode — e.g. "ST27"
 * @returns {Array<{batchId,binId,binCode,site,location,qty,uom,qaStatus,version}>}
 */
function getStockByItemCode(itemCode) {
  return protect(function () {
    var code = String(itemCode || '').trim().toUpperCase();
    if (!code) return [];

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    if (!invSheet) return [];

    var invData = invSheet.getDataRange().getValues();

    // Build bin_id → bin_code lookup from Bin_Master (col 0=bin_id, col 2=bin_code)
    var binCodeMap = {};
    var binSheet = ss.getSheetByName(CONFIG.SHEETS.BIN);
    if (binSheet) {
      var binData = binSheet.getDataRange().getValues();
      for (var b = 1; b < binData.length; b++) {
        var bid = String(binData[b][0] || '').trim();
        var bcod = String(binData[b][2] || '').trim();
        if (bid) binCodeMap[bid] = bcod || bid;
      }
    }

    var results = [];
    var C = CONFIG.INVENTORY_COLS;          // shorthand

    for (var i = 1; i < invData.length; i++) {
      var row = invData[i];

      // Match on item_code_cache (col 2)
      if (String(row[C.ITEM_CODE_CACHE] || '').trim().toUpperCase() !== code) continue;

      var qty = Number(row[C.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      var binId = String(row[C.BIN_ID] || '').trim();
      var batchInfo = _resolveBatchReference(code, row[C.BATCH_ID]);
      var batchId = String((batchInfo && batchInfo.batchId) || row[C.BATCH_ID] || '').trim();
      var site = String(row[C.SITE] || '').trim();
      var loc = String(row[C.LOCATION] || '').trim();
      var uom = String(row[C.UOM] || 'NOS').trim() || 'NOS';
      var qa = String(row[C.QUALITY_STATUS] || 'PENDING').trim().toUpperCase();
      var ver = String(row[C.VERSION] || '').trim();

      // All primitives — no Date, no nested arrays — safe for GAS serialization
      results.push({
        batchId: batchId,
        batchNumber: String((batchInfo && batchInfo.batchNumber) || batchId),
        binId: binId,
        binCode: binCodeMap[binId] || binId,
        site: site,
        location: loc,
        qty: qty,
        uom: uom,
        qaStatus: qa,
        version: ver
      });
    }

    // Sort: APPROVED/OVERRIDDEN first, then highest qty
    results.sort(function (a, b) {
      var okA = (a.qaStatus === 'APPROVED' || a.qaStatus === 'OVERRIDDEN') ? 1 : 0;
      var okB = (b.qaStatus === 'APPROVED' || b.qaStatus === 'OVERRIDDEN') ? 1 : 0;
      return okB !== okA ? okB - okA : b.qty - a.qty;
    });

    return results;
  });
}

// =====================================================
// 2. SECURITY & AUTH MIDDLEWARE (Consolidated)
// =====================================================

const SECURITY = {
  ROLES: {
    MANAGER: 'MANAGER',
    WORKER: 'WORKER',
    QUALITY_MANAGER: 'QUALITY_MANAGER'
  },
  CACHE_DURATION: 300
};

const ROLE_RANK = {
  MANAGER: 3,
  QUALITY_MANAGER: 2,
  WORKER: 1
};

/**
 * Gets the effective user.
 * STRICT GLOBAL_DEV_BYPASS ENFORCEMENT.
 */
function getEffectiveUser() {
  if (GLOBAL_DEV_BYPASS === true) {
    return {
      email: 'kadambari.purohit@cultivator.in',
      role: SECURITY.ROLES.MANAGER,
      status: 'ACTIVE',
      fullName: 'System Owner (Bypass)'
    };
  }

  // STANDARD AUTH PATH
  try {
    let email = Session.getActiveUser().getEmail();
    const ownerEmail = Session.getEffectiveUser().getEmail();

    if (!email) throw new Error('Identity unavailable');

    // Owner Escape Hatch for Prod
    if (email.toLowerCase() === ownerEmail.toLowerCase()) {
      return {
        email: email,
        role: SECURITY.ROLES.MANAGER,
        status: 'ACTIVE',
        fullName: 'System Owner'
      };
    }

    return _checkACL(email);
  } catch (e) {
    throw new Error('Auth Failed: ' + e.message);
  }
}

/**
 * ACL Check Helper
 */
function _checkACL(email) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'USER_AUTH_V5_' + email;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.ACCESS_CONTROL);

  if (!sheet) {
    throw new Error('CRITICAL: Access_Control_List Sheet Missing.');
  }

  const data = _getSheetValuesCached(sheet.getName());
  // Assuming Row 0 is header
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      const user = {
        email: email,
        role: String(data[i][1]).toUpperCase(),
        fullName: data[i][2],
        status: String(data[i][3]).toUpperCase()
      };

      if (user.status !== 'ACTIVE') {
        throw new Error('Account Suspended');
      }

      cache.put(cacheKey, JSON.stringify(user), 300);
      return user;
    }
  }
  throw new Error('User not in Access Control List');
}

function requireRole(requiredRole) {
  const user = getEffectiveUser();
  const requiredRank = ROLE_RANK[String(requiredRole || '').toUpperCase()] || 0;
  const userRank = ROLE_RANK[String(user.role || '').toUpperCase()] || 0;
  if (userRank < requiredRank) {
    throw new Error(`Permission Denied: Requires ${requiredRole}`);
  }
  return user;
}

function protect(callback) {
  _beginRequestCache();
  try {
    const user = getEffectiveUser();
    return callback(user);
  } catch (error) {
    Logger.log('Security Violation: ' + error.toString());
    throw error;
  } finally {
    _endRequestCache();
  }
}

function isQualityManager(user) {
  const userRank = ROLE_RANK[String(user && user.role || '').toUpperCase()] || 0;
  const qaRank = ROLE_RANK[SECURITY.ROLES.QUALITY_MANAGER];
  return userRank >= qaRank;
}

function assertQualityManagerAccess() {
  const user = getEffectiveUser();
  if (!isQualityManager(user)) {
    throw new Error('Access denied: Quality Manager required');
  }
  return user;
}

function requireOperationalUser() {
  const user = getEffectiveUser();
  const userRank = ROLE_RANK[String(user && user.role || '').toUpperCase()] || 0;
  const workerRank = ROLE_RANK[SECURITY.ROLES.WORKER];
  if (userRank < workerRank) {
    throw new Error('Permission Denied: Requires WORKER');
  }
  if (user.role === SECURITY.ROLES.QUALITY_MANAGER) {
    throw new Error('Permission Denied: Operational role required');
  }
  return user;
}

// =====================================================
// 3. CACHING & HELPERS
// =====================================================

// Request-scoped cache (one sheet read per execution).
let _REQ_CACHE = { sheets: {}, values: {}, headers: {}, _depth: 0 };

function _ensureRequestCache(reset) {
  if (reset === true || !_REQ_CACHE) {
    _REQ_CACHE = { sheets: {}, values: {}, headers: {}, _depth: 0 };
  }
  return _REQ_CACHE;
}

// Begin/end request cache scope to avoid cross-invocation leakage.
function _beginRequestCache() {
  if (!_REQ_CACHE || !_REQ_CACHE._depth) {
    _ensureRequestCache(true);
  }
  _REQ_CACHE._depth = (_REQ_CACHE._depth || 0) + 1;
}

function _endRequestCache() {
  if (_REQ_CACHE && _REQ_CACHE._depth > 0) {
    _REQ_CACHE._depth -= 1;
  }
}

function _withRequestCache(callback) {
  _beginRequestCache();
  try {
    return callback();
  } finally {
    _endRequestCache();
  }
}

function _getScriptCacheJson(cacheKey) {
  try {
    const raw = CacheService.getScriptCache().get(cacheKey);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function _putScriptCacheJson(cacheKey, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(value), ttlSeconds || 60);
  } catch (e) {
    // Cache failures should never block operational flows.
  }
}

function _ensureProductionLedgerSchema(sheet) {
  if (!sheet) return;
  const target = String(CONFIG.SHEETS.PRODUCTION_LEDGER || '').trim().toUpperCase();
  const actual = String(sheet.getName() || '').trim().toUpperCase();
  if (!target || actual !== target) return;
  _ensureRequestCache();
  if (_REQ_CACHE.productionLedgerSchemaReady) return;

  const lastCol = Math.max(Number(sheet.getLastColumn()) || 1, 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const normalized = header.map(function (h) { return String(h || '').trim().toUpperCase(); });
  const hasRejected = normalized.indexOf('QTY_REJECTED') !== -1;

  if (!hasRejected) {
    // Insert rejected qty immediately after NET_OUTSTANDING to keep ledger clear.
    const insertAfter = CONFIG.PRODUCTION_LEDGER_COLS.NET_OUTSTANDING + 1;
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue('qty_rejected');
    delete _REQ_CACHE.values[CONFIG.SHEETS.PRODUCTION_LEDGER];
    delete _REQ_CACHE.headers[CONFIG.SHEETS.PRODUCTION_LEDGER];
  }

  _REQ_CACHE.productionLedgerSchemaReady = true;
}

function _ensureInventoryDateColumnsSchema(sheet) {
  if (!sheet) return;
  const target = String(CONFIG.SHEETS.INVENTORY || '').trim().toUpperCase();
  const actual = String(sheet.getName() || '').trim().toUpperCase();
  if (!target || actual !== target) return;
  _ensureRequestCache();
  if (_REQ_CACHE.inventoryDateSchemaReady) return;

  const refreshHeader = function () {
    const lastCol = Math.max(Number(sheet.getLastColumn()) || 1, 1);
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    return header.map(function (h) {
      return String(h || '').trim().toUpperCase().replace(/\s+/g, '_');
    });
  };

  let normalized = refreshHeader();
  let inwardIdx = normalized.indexOf('INWARD_DATE');
  let transferIdx = normalized.indexOf('LAST_TRANSFER_DATE');

  if (inwardIdx === -1) {
    const insertAfter = CONFIG.INVENTORY_COLS.LAST_UPDATED + 1;
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue('inward_date');
    normalized = refreshHeader();
    inwardIdx = normalized.indexOf('INWARD_DATE');
    delete _REQ_CACHE.values[CONFIG.SHEETS.INVENTORY];
    delete _REQ_CACHE.headers[CONFIG.SHEETS.INVENTORY];
  }

  if (transferIdx === -1) {
    const afterCol = (inwardIdx >= 0 ? inwardIdx + 1 : (CONFIG.INVENTORY_COLS.INWARD_DATE + 1));
    sheet.insertColumnAfter(afterCol);
    sheet.getRange(1, afterCol + 1).setValue('last_transfer_date');
    delete _REQ_CACHE.values[CONFIG.SHEETS.INVENTORY];
    delete _REQ_CACHE.headers[CONFIG.SHEETS.INVENTORY];
  }

  _REQ_CACHE.inventoryDateSchemaReady = true;
}

function _getSheetCached(sheetName) {
  _ensureRequestCache();
  if (_REQ_CACHE.sheets[sheetName]) return _REQ_CACHE.sheets[sheetName];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, sheetName);
  if (sheetName === CONFIG.SHEETS.PRODUCTION_LEDGER) {
    _ensureProductionLedgerSchema(sheet);
  }
  if (sheetName === CONFIG.SHEETS.INVENTORY) {
    _ensureInventoryDateColumnsSchema(sheet);
  }
  _REQ_CACHE.sheets[sheetName] = sheet;
  return sheet;
}

function _getSheetValuesCached(sheetName) {
  _ensureRequestCache();
  if (_REQ_CACHE.values[sheetName]) return _REQ_CACHE.values[sheetName];
  const sheet = _getSheetCached(sheetName);
  const values = sheet.getDataRange().getValues();
  _REQ_CACHE.values[sheetName] = values;
  return values;
}

function _getProductionLedgerOrderRows(orderRef) {
  _ensureRequestCache();
  const ref = String(orderRef || '').trim();
  if (!ref) return [];
  const cache = _REQ_CACHE.productionLedgerRowsByOrder || (_REQ_CACHE.productionLedgerRowsByOrder = {});
  if (cache[ref]) return cache[ref];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
  const data = _getSheetValuesCached(sheet.getName());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.PRODUCTION_LEDGER_COLS.ORDER_ID] || '').trim() !== ref) continue;
    rows.push({ rowIndex: i + 1, row: data[i] });
  }
  cache[ref] = rows;
  return rows;
}

// One-time/manual safeguard if production ledger schema drifted in old sheets.
function ensureProductionLedgerRejectedColumn() {
  return withScriptLock(function () {
    protect(function () { return requireRole(SECURITY.ROLES.MANAGER); });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
    _ensureProductionLedgerSchema(sheet);
    return { success: true, sheet: sheet.getName() };
  });
}

function migrateInventoryDateColumns() {
  return withScriptLock(function () {
    protect(function () { return requireRole(SECURITY.ROLES.MANAGER); });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    _ensureInventoryDateColumnsSchema(invSheet);

    const invData = _getSheetValuesCached(invSheet.getName());
    const movementSheet = ss.getSheetByName(CONFIG.SHEETS.MOVEMENT);
    const movData = movementSheet ? _getSheetValuesCached(movementSheet.getName()) : [];
    const transferMap = {};

    for (let i = 1; i < movData.length; i++) {
      const row = movData[i];
      const type = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').trim().toUpperCase();
      if (type !== CONFIG.MOVEMENT_TYPES.TRANSFER && type !== CONFIG.MOVEMENT_TYPES.TRANSFER_QUARANTINE) continue;
      const tsRaw = row[CONFIG.MOVEMENT_COLS.TIMESTAMP];
      const ts = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
      if (!(ts instanceof Date) || isNaN(ts.getTime())) continue;
      const itemRaw = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim().toUpperCase();
      const batchRaw = String(row[CONFIG.MOVEMENT_COLS.BATCH_ID] || '').trim().toUpperCase();
      const fromBin = String(row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID] || '').trim().toUpperCase();
      const toBin = String(row[CONFIG.MOVEMENT_COLS.TO_BIN_ID] || '').trim().toUpperCase();
      if (!itemRaw || !batchRaw) continue;
      const keys = [];
      if (fromBin) keys.push(itemRaw + '||' + batchRaw + '||' + fromBin);
      if (toBin) keys.push(itemRaw + '||' + batchRaw + '||' + toBin);
      keys.forEach(function (k) {
        if (!transferMap[k] || ts.getTime() > transferMap[k].getTime()) transferMap[k] = ts;
      });
    }

    let backfilledInward = 0;
    let backfilledTransfer = 0;
    const rowsToWrite = [];
    const rowIndexes = [];

    for (let i = 1; i < invData.length; i++) {
      const row = invData[i].slice();
      let changed = false;

      const inwardVal = row[CONFIG.INVENTORY_COLS.INWARD_DATE];
      if (!inwardVal) {
        let backfillInward = null;
        const invId = String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || '').trim();
        const m = invId.match(/INV-(\d+)-/);
        if (m && m[1]) {
          const tsNum = Number(m[1]);
          if (isFinite(tsNum) && tsNum > 0) backfillInward = new Date(tsNum);
        }
        if (!backfillInward) {
          const lu = row[CONFIG.INVENTORY_COLS.LAST_UPDATED];
          if (lu instanceof Date) backfillInward = lu;
          else if (String(lu || '').trim()) {
            const parsed = new Date(lu);
            if (!isNaN(parsed.getTime())) backfillInward = parsed;
          }
        }
        if (backfillInward) {
          row[CONFIG.INVENTORY_COLS.INWARD_DATE] = backfillInward;
          backfilledInward++;
          changed = true;
        }
      }

      const transferVal = row[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE];
      if (!transferVal) {
        const itemIdRaw = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();
        const itemCodeRaw = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
        const batchRaw = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim().toUpperCase();
        const binRaw = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim().toUpperCase();
        const byItemId = itemIdRaw ? transferMap[itemIdRaw + '||' + batchRaw + '||' + binRaw] : null;
        const byCode = itemCodeRaw ? transferMap[itemCodeRaw + '||' + batchRaw + '||' + binRaw] : null;
        const backfillTransfer = byItemId || byCode || null;
        if (backfillTransfer) {
          row[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = backfillTransfer;
          backfilledTransfer++;
          changed = true;
        }
      }

      if (changed) {
        rowsToWrite.push(row);
        rowIndexes.push(i + 1);
      }
    }

    for (let i = 0; i < rowsToWrite.length; i++) {
      const rowIndex = rowIndexes[i];
      const row = rowsToWrite[i];
      invSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, rowIndex, CONFIG.INVENTORY_COLS.INWARD_DATE + 1, row[CONFIG.INVENTORY_COLS.INWARD_DATE]);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, rowIndex, CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE + 1, row[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE]);
    }

    return {
      success: true,
      rowsTouched: rowsToWrite.length,
      inwardBackfilled: backfilledInward,
      transferBackfilled: backfilledTransfer
    };
  });
}


function _getSheetHeaderMapCached(sheetName) {
  _ensureRequestCache();
  if (_REQ_CACHE.headers[sheetName]) return _REQ_CACHE.headers[sheetName];
  const values = _getSheetValuesCached(sheetName);
  const header = values[0] || [];
  const map = {};
  header.forEach((h, idx) => {
    const key = String(h || '').trim().toUpperCase();
    if (key) map[key] = idx;
  });
  _REQ_CACHE.headers[sheetName] = map;
  return map;
}

function _pickColumnIndex(headerMap, candidates, fallbackIdx) {
  const list = Array.isArray(candidates) ? candidates : [];
  for (let i = 0; i < list.length; i++) {
    const key = String(list[i] || '').trim().toUpperCase();
    if (!key) continue;
    if (typeof headerMap[key] === 'number') return headerMap[key];
  }
  return (typeof fallbackIdx === 'number') ? fallbackIdx : null;
}

function _getItemSheetColumnMap(sheet) {
  if (!sheet) return null;
  _ensureRequestCache();
  _REQ_CACHE.itemColMaps = _REQ_CACHE.itemColMaps || {};
  const cacheKey = sheet.getName();
  if (_REQ_CACHE.itemColMaps[cacheKey]) return _REQ_CACHE.itemColMaps[cacheKey];

  const headerMap = _getSheetHeaderMap(sheet);
  const map = {
    ITEM_ID: _pickColumnIndex(headerMap, ['ITEM_ID', 'ID'], CONFIG.ITEM_COLS.ITEM_ID),
    ITEM_CODE: _pickColumnIndex(headerMap, ['ITEM_CODE', 'CODE', 'ITEM_CODE_CACHE'], CONFIG.ITEM_COLS.ITEM_CODE),
    ITEM_NAME: _pickColumnIndex(headerMap, ['ITEM_NAME', 'NAME'], CONFIG.ITEM_COLS.ITEM_NAME),
    ITEM_GROUP_CODE: _pickColumnIndex(headerMap, ['ITEM_GROUP_CODE', 'ITEM_GROUP', 'GROUP_CODE'], CONFIG.ITEM_COLS.ITEM_GROUP_CODE),
    UOM_CODE: _pickColumnIndex(headerMap, ['UOM_CODE', 'UOM', 'BASE_UOM'], CONFIG.ITEM_COLS.UOM_CODE),
    STATUS: _pickColumnIndex(headerMap, ['STATUS', 'ITEM_STATUS'], CONFIG.ITEM_COLS.STATUS),
    MIN_STOCK_LEVEL: _pickColumnIndex(headerMap, ['MIN_STOCK_LEVEL', 'MIN_STOCK', 'MINIMUM_STOCK_LEVEL'], CONFIG.ITEM_COLS.MIN_STOCK_LEVEL),
    DEFAULT_QUALITY_DAYS: _pickColumnIndex(headerMap, ['DEFAULT_QUALITY_DAYS', 'QUALITY_DAYS', 'QA_DAYS'], null),
    ALT_UOM: _pickColumnIndex(headerMap, ['ALT_UOM'], null),
    CONVERSION_FACTOR_TO_KG: _pickColumnIndex(headerMap, ['CONVERSION_FACTOR_TO_KG', 'CONVERSION_TO_KG', 'FACTOR_TO_KG'], null)
  };

  _REQ_CACHE.itemColMaps[cacheKey] = map;
  return map;
}

function _getInventorySheetColumnMap(sheet) {
  if (!sheet) return null;
  _ensureRequestCache();
  _REQ_CACHE.inventoryColMaps = _REQ_CACHE.inventoryColMaps || {};
  const cacheKey = sheet.getName();
  if (_REQ_CACHE.inventoryColMaps[cacheKey]) return _REQ_CACHE.inventoryColMaps[cacheKey];

  const headerMap = _getSheetHeaderMap(sheet);
  const map = {
    INVENTORY_ID: _pickColumnIndex(headerMap, ['INVENTORY_ID', 'INV_ID'], CONFIG.INVENTORY_COLS.INVENTORY_ID),
    ITEM_ID: _pickColumnIndex(headerMap, ['ITEM_ID', 'ID'], CONFIG.INVENTORY_COLS.ITEM_ID),
    ITEM_CODE_CACHE: _pickColumnIndex(headerMap, ['ITEM_CODE_CACHE', 'ITEM_CODE', 'CODE'], CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE),
    BATCH_ID: _pickColumnIndex(headerMap, ['BATCH_ID', 'BATCH', 'BATCH_NO'], CONFIG.INVENTORY_COLS.BATCH_ID),
    GIN_NO: _pickColumnIndex(headerMap, ['GIN_NO', 'BILL_NO', 'REF_NO'], CONFIG.INVENTORY_COLS.GIN_NO),
    VERSION: _pickColumnIndex(headerMap, ['VERSION'], CONFIG.INVENTORY_COLS.VERSION),
    QUALITY_STATUS: _pickColumnIndex(headerMap, ['QUALITY_STATUS', 'QA_STATUS', 'STATUS'], CONFIG.INVENTORY_COLS.QUALITY_STATUS),
    QUALITY_DATE: _pickColumnIndex(headerMap, ['QUALITY_DATE', 'QA_DATE'], CONFIG.INVENTORY_COLS.QUALITY_DATE),
    BIN_ID: _pickColumnIndex(headerMap, ['BIN_ID', 'BIN', 'BIN_CODE'], CONFIG.INVENTORY_COLS.BIN_ID),
    SITE: _pickColumnIndex(headerMap, ['SITE'], CONFIG.INVENTORY_COLS.SITE),
    LOCATION: _pickColumnIndex(headerMap, ['LOCATION'], CONFIG.INVENTORY_COLS.LOCATION),
    TOTAL_QUANTITY: _pickColumnIndex(headerMap, ['TOTAL_QUANTITY', 'QUANTITY', 'QTY'], CONFIG.INVENTORY_COLS.TOTAL_QUANTITY),
    UOM: _pickColumnIndex(headerMap, ['UOM'], CONFIG.INVENTORY_COLS.UOM),
    LOT_NO: _pickColumnIndex(headerMap, ['LOT_NO', 'LOT'], CONFIG.INVENTORY_COLS.LOT_NO),
    EXPIRY_DATE: _pickColumnIndex(headerMap, ['EXPIRY_DATE', 'EXP_DATE'], CONFIG.INVENTORY_COLS.EXPIRY_DATE),
    LAST_UPDATED: _pickColumnIndex(headerMap, ['LAST_UPDATED', 'UPDATED_AT'], CONFIG.INVENTORY_COLS.LAST_UPDATED),
    INWARD_DATE: _pickColumnIndex(headerMap, ['INWARD_DATE', 'RECEIPT_DATE'], CONFIG.INVENTORY_COLS.INWARD_DATE),
    LAST_TRANSFER_DATE: _pickColumnIndex(headerMap, ['LAST_TRANSFER_DATE', 'TRANSFER_DATE'], CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE)
  };

  _REQ_CACHE.inventoryColMaps[cacheKey] = map;
  Logger.log('[INV_COL_MAP] sheet=' + cacheKey + ' INWARD_DATE_idx=' + map.INWARD_DATE + ' LAST_TRANSFER_DATE_idx=' + map.LAST_TRANSFER_DATE);
  return map;
}

function _getItemMasterMaps() {
  _ensureRequestCache();
  if (_REQ_CACHE.itemMasterMaps) return _REQ_CACHE.itemMasterMaps;

  const empty = {
    idToCode: {},
    codeToId: {},
    codeToName: {},
    codeToUom: {},
    codeToGroup: {},
    codeToMinStock: {},   // keyed by code — last-write-wins (legacy, kept for dashboards)
    idToMinStock: {},     // keyed by item_id — correct for per-item alerts
    idToItemInfo: {},     // keyed by item_id — { code, name, uom, minStock }
    codeToStatus: {},
    codeDisplayByNorm: {}
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);
  if (!sheet) {
    _REQ_CACHE.itemMasterMaps = empty;
    return empty;
  }

  const data = _getSheetValuesCached(sheet.getName());
  const cols = _getItemSheetColumnMap(sheet);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[cols.ITEM_ID] || '').trim();
    const code = String(row[cols.ITEM_CODE] || '').trim();
    if (!id || !code) continue;

    const idNorm = id.toUpperCase();
    const codeNorm = code.toUpperCase();
    const name = String(row[cols.ITEM_NAME] || '').trim();
    const group = String(row[cols.ITEM_GROUP_CODE] || '').trim();
    const uom = String(row[cols.UOM_CODE] || 'KG').trim().toUpperCase() || 'KG';
    const statusRaw = (typeof cols.STATUS === 'number') ? String(row[cols.STATUS] || '') : '';
    const status = statusRaw.trim() ? statusRaw.trim().toUpperCase() : 'ACTIVE';
    const minStock = (typeof cols.MIN_STOCK_LEVEL === 'number') ? Number(row[cols.MIN_STOCK_LEVEL]) : 0;

    empty.idToCode[idNorm] = code;
    empty.codeToId[codeNorm] = id;
    empty.codeToName[codeNorm] = name || code;
    empty.codeToUom[codeNorm] = uom;
    empty.codeToGroup[codeNorm] = group;
    empty.codeToStatus[codeNorm] = status;
    empty.codeDisplayByNorm[codeNorm] = code;
    if (isFinite(minStock) && minStock > 0) {
      empty.codeToMinStock[codeNorm] = minStock;
      // Per item_id: correct for duplicate-code items (e.g. 3x ST27 with different ids)
      empty.idToMinStock[idNorm] = minStock;
      empty.idToItemInfo[idNorm] = {
        id: id, code: code, codeNorm: codeNorm,
        name: name || code, uom: uom, minStock: minStock
      };
    }
  }

  _REQ_CACHE.itemMasterMaps = empty;
  return empty;
}

function _updateSheetCacheCell(sheetName, rowIndex, colIndex, value) {
  _ensureRequestCache();
  const name = String(sheetName || '').trim();
  if (!name || !_REQ_CACHE.values[name]) return;
  const r = Number(rowIndex) - 1;
  const c = Number(colIndex) - 1;
  if (!isFinite(r) || !isFinite(c)) return;
  if (_REQ_CACHE.values[name][r]) {
    _REQ_CACHE.values[name][r][c] = value;
  }
}

function _appendSheetCacheRow(sheetName, row) {
  _ensureRequestCache();
  const name = String(sheetName || '').trim();
  if (!name || !_REQ_CACHE.values[name]) return;
  _REQ_CACHE.values[name].push(row);
}

function _setSheetCacheRow(sheetName, rowIndex, row) {
  _ensureRequestCache();
  const name = String(sheetName || '').trim();
  const idx = Number(rowIndex) - 1;
  if (!name || !_REQ_CACHE.values[name] || !isFinite(idx) || idx < 0) return;
  _REQ_CACHE.values[name][idx] = row;
}

function _getSheetByNameLoose(ss, sheetName) {
  if (!ss || !sheetName) return null;
  const direct = ss.getSheetByName(sheetName);
  if (direct) return direct;
  const target = String(sheetName || '').trim().toUpperCase();
  if (!target) return null;
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const nm = String(sheets[i].getName() || '').trim().toUpperCase();
    if (nm === target) return sheets[i];
  }
  return null;
}

function _deleteSheetCacheRow(sheetName, rowIndex) {
  _ensureRequestCache();
  const name = String(sheetName || '').trim();
  if (!name || !_REQ_CACHE.values[name]) return;
  const idx = Number(rowIndex) - 1;
  if (!isFinite(idx) || idx < 0 || idx >= _REQ_CACHE.values[name].length) return;
  _REQ_CACHE.values[name].splice(idx, 1);
  if (name === CONFIG.SHEETS.INVENTORY) {
    delete _REQ_CACHE.inventoryUsageSummaryByBin;
    delete _REQ_CACHE.inventoryUsageSummaryByBinSet;
  }
}

function _clearSheetCache(sheetName) {
  _ensureRequestCache();
  const name = String(sheetName || '').trim();
  if (!name) return;
  delete _REQ_CACHE.values[name];
  delete _REQ_CACHE.headers[name];
  if (name === CONFIG.SHEETS.BATCH) {
    delete _REQ_CACHE.batchMasterMaps;
  }
  if (name === CONFIG.SHEETS.BIN) {
    delete _REQ_CACHE.binMasterMetaContext;
    delete _REQ_CACHE.inventoryUsageSummaryByBin;
    delete _REQ_CACHE.inventoryUsageSummaryByBinSet;
  }
  if (name === CONFIG.SHEETS.INVENTORY) {
    delete _REQ_CACHE.inventoryUsageSummaryByBin;
    delete _REQ_CACHE.inventoryUsageSummaryByBinSet;
  }
  if (name === CONFIG.SHEETS.PRODUCTION_LEDGER) {
    delete _REQ_CACHE.productionLedgerRowsByOrder;
  }
  _clearScriptCachesForSheet(name);
}

function _clearScriptCaches(keys) {
  if (!keys || keys.length === 0) return;
  try {
    CacheService.getScriptCache().removeAll(keys);
  } catch (e) {
    // Persistent cache is an optimization only; operational writes must continue.
  }
}

function _clearScriptCachesForSheet(sheetName) {
  const name = String(sheetName || '').trim();
  const keys = [];

  if (name === CONFIG.SHEETS.BIN || name === CONFIG.SHEETS.RACK || name === CONFIG.SHEETS.LOCATION_MASTER) {
    keys.push('BINS_FAST_V3', 'BINS_V6', 'LOCATIONS_MASTER_V1');
  }

  if (name === CONFIG.SHEETS.PRODUCTION_LEDGER) {
    keys.push('ACTIVE_PRODUCTION_ORDERS_V1');
  }

  if (name === CONFIG.SHEETS.INVENTORY) {
    keys.push('ACTIVE_PRODUCTION_ORDERS_V1');
  }

  if (name === CONFIG.SHEETS.PACKING_VERSION_MASTER || name === CONFIG.SHEETS.INVENTORY) {
    keys.push('PACKING_MATERIAL_DASHBOARD_V1', 'PACKING_MATERIAL_DASHBOARD_V2');
  }

  _clearScriptCaches(keys);
}

function _clearInwardLookupCaches(itemCodes) {
  const keys = [];
  (itemCodes || []).forEach(function (code) {
    const norm = String(code || '').trim().toUpperCase();
    if (!norm) return;
    keys.push('INWARD_NEXT_VERSION_V4_' + norm);
  });
  _clearScriptCaches(keys);
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function getItemByCodeCached(itemCode) {
  const cache = CacheService.getScriptCache();
  const normCode = String(itemCode || '').trim().toUpperCase();
  if (!normCode) return null;
  const cacheKey = 'item_' + normCode;
  const cachedItem = cache.get(cacheKey);
  if (cachedItem) return JSON.parse(cachedItem);

  const maps = _getItemMasterMaps();
  const id = maps.codeToId[normCode];
  if (!id) return null;

  const item = {
    id: id,
    code: maps.codeDisplayByNorm[normCode] || normCode,
    name: maps.codeToName[normCode] || (maps.codeDisplayByNorm[normCode] || normCode)
  };
  cache.put(cacheKey, JSON.stringify(item), 300);
  return item;
}

function getItemIdByCode(itemCode) {
  const item = getItemByCodeCached(itemCode);
  return item ? item.id : null;
}

function _findHeaderColumnLoose(headers, candidates) {
  const wanted = (candidates || []).map(function (c) {
    return String(c || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  });
  for (let i = 0; i < headers.length; i++) {
    const key = String(headers[i] || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (wanted.indexOf(key) !== -1) return i;
  }
  return -1;
}

function _findMinStockSourceHeader_(sourceData) {
  const codeHeaders = ['ITEM CODE', 'ITEM_CODE', 'CODE', 'ITEMCODE', 'ITEM NO', 'ITEM'];
  const minHeaders = [
    'MINIMUM',
    'MINUMUM',
    'MINIMUM STOCK',
    'MINUMUM STOCK',
    'MINIMUM STOCK LEVEL',
    'MINUMUM STOCK LEVEL',
    'MIN_STOCK',
    'MIN STOCK',
    'MIN_STOCK_LEVEL',
    'MIN LEVEL',
    'MINIMUM LEVEL',
    'REORDER LEVEL'
  ];
  const maxRows = Math.min(sourceData.length, 10);
  for (let r = 0; r < maxRows; r++) {
    const headers = sourceData[r] || [];
    const codeCol = _findHeaderColumnLoose(headers, codeHeaders);
    const minCol = _findHeaderColumnLoose(headers, minHeaders);
    if (codeCol >= 0 && minCol >= 0) {
      return { headerRow: r, codeCol: codeCol, minCol: minCol };
    }
  }
  return null;
}

function previewMinStockUpdateFromSheet(sheetName) {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);
    return _buildMinStockUpdatePlan_(sheetName || 'Min_Stock_Update');
  });
}

function bulkUpdateMinStockFromSheet(sheetName) {
  return withScriptLock(function () {
    protect(function () { return requireRole(SECURITY.ROLES.MANAGER); });
    const plan = _buildMinStockUpdatePlan_(sheetName || 'Min_Stock_Update');
    if (!plan.updates || plan.updates.length === 0) {
      return {
        success: false,
        updated: 0,
        missing: plan.missing || [],
        skipped: plan.skipped || [],
        message: 'No matching min-stock rows found to update.'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ITEM);
    const minCol = _getItemSheetColumnMap(itemSheet).MIN_STOCK_LEVEL + 1;
    plan.updates.forEach(function (u) {
      itemSheet.getRange(u.itemMasterRow, minCol).setValue(u.newMinStock);
      _updateSheetCacheCell(CONFIG.SHEETS.ITEM, u.itemMasterRow, minCol, u.newMinStock);
    });
    _clearSheetCache(CONFIG.SHEETS.ITEM);

    return {
      success: true,
      updated: plan.updates.length,
      missing: plan.missing,
      skipped: plan.skipped,
      message: 'Updated min stock for ' + plan.updates.length + ' item(s).'
    };
  });
}

function _buildMinStockUpdatePlan_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = _getSheetOrThrow(ss, sheetName || 'Min_Stock_Update');
  const itemSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ITEM);
  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) {
    return { updates: [], missing: [], skipped: ['Source sheet has no data rows.'] };
  }

  const header = _findMinStockSourceHeader_(sourceData);
  if (!header) {
    const sampleHeaders = sourceData.slice(0, Math.min(sourceData.length, 3)).map(function (row, idx) {
      return 'Row ' + (idx + 1) + ': ' + (row || []).filter(String).join(' | ');
    }).join(' ; ');
    throw new Error('Source sheet must have Item Code and Minimum/Min Stock columns. Found: ' + sampleHeaders);
  }
  const codeCol = header.codeCol;
  const minCol = header.minCol;

  const itemCols = _getItemSheetColumnMap(itemSheet);
  const itemData = _getSheetValuesCached(itemSheet.getName());
  const rowByCode = {};
  for (let i = 1; i < itemData.length; i++) {
    const code = String(itemData[i][itemCols.ITEM_CODE] || '').trim().toUpperCase();
    if (!code) continue;
    rowByCode[code] = i + 1;
  }

  const updates = [];
  const missing = [];
  const skipped = [];
  const seen = {};
  for (let r = header.headerRow + 1; r < sourceData.length; r++) {
    const code = String(sourceData[r][codeCol] || '').trim().toUpperCase();
    const min = Number(sourceData[r][minCol]);
    if (!code) continue;
    if (!isFinite(min) || min < 0) {
      skipped.push('Row ' + (r + 1) + ' ' + code + ': invalid minimum stock');
      continue;
    }
    if (seen[code]) {
      skipped.push('Row ' + (r + 1) + ' ' + code + ': duplicate in update sheet');
      continue;
    }
    seen[code] = true;

    const itemMasterRow = rowByCode[code];
    if (!itemMasterRow) {
      missing.push(code);
      continue;
    }
    updates.push({ itemCode: code, itemMasterRow: itemMasterRow, newMinStock: min });
  }

  return { updates: updates, missing: missing, skipped: skipped };
}

function _getItemCodeById(itemId) {
  if (!itemId) return null;
  const targetId = String(itemId || '').trim().toUpperCase();
  const maps = _getItemMasterMaps();
  return maps.idToCode[targetId] || null;
}

function _getCanonicalItemCode(itemCode, itemId) {
  const rawCode = String(itemCode || '').trim();
  const rawId = String(itemId || '').trim();
  const maps = _getItemMasterMaps();

  if (rawId) {
    const byId = maps.idToCode[rawId.toUpperCase()];
    if (byId) return String(byId || '').trim();
  }

  if (rawCode) {
    const byCode = maps.codeDisplayByNorm[rawCode.toUpperCase()];
    if (byCode) return String(byCode || '').trim();
  }

  return rawCode ? rawCode.toUpperCase() : '';
}


/**
 * Get all racks/locations/bins from Rack_Master for dropdown population
 * Schema: rack_id, rack_code, site, location, rack_capacity, rack_status
 */
function getRacksData() {
  return _withRequestCache(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
    if (!sheet) return [];

    const data = _getSheetValuesCached(sheet.getName());
    const racks = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // Skip empty rows
        racks.push({
          rackId: row[0],
          rackCode: row[1],
          site: row[2],
          location: row[3],
          capacity: row[4] || 0,
          status: row[5] || 'Active'
        });
      }
    }
    return racks;
  });
}

/**
 * Get all locations from Location_master sheet
 */
function getLocationsFromMaster() {
  return _withRequestCache(function () {
    const cacheKey = 'LOCATIONS_MASTER_V1';
    const cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.LOCATION_MASTER);
    if (!sheet) return [];
    const data = _getSheetValuesCached(sheet.getName());
    const locs = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        locs.push({
          site: String(data[i][0]).trim(),
          location: String(data[i][1]).trim()
        });
      }
    }
    _putScriptCacheJson(cacheKey, locs, 120);
    return locs;
  });
}

/**
 * Get all bins with hierarchy (Site/Location from Rack)
 */
function getBinsFast() {
  return _withRequestCache(function () {
    const cacheKey = 'BINS_FAST_V3';
    const cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
    const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
    const headerMap = _getSheetHeaderMap(binSheet);

    const rackMap = {};
    if (rackSheet) {
      const rackData = _getSheetValuesCached(rackSheet.getName());
      for (let i = 1; i < rackData.length; i++) {
        const rackId = String(rackData[i][0] || '').trim();
        if (!rackId) continue;
        rackMap[rackId] = {
          site: String(rackData[i][2] || '').trim(),
          location: String(rackData[i][3] || '').trim()
        };
      }
    }

    const data = _getSheetValuesCached(binSheet.getName());
    const binMetaMap = _buildBinMasterMetaMap(binSheet, data, headerMap);
    const bins = [];

    for (let i = 1; i < data.length; i++) {
      const binId = String(data[i][0] || '').trim();
      if (!binId) continue;
      const rackRef = String(data[i][1] || '').trim();
      const rackMeta = rackMap[rackRef] || { site: 'Main', location: 'Main' };
      const meta = binMetaMap[binId] || {};
      const capacityUom = _normalizeCapacityUom(meta.declaredCapacityUom || meta.capacityUom || 'KG');
      const binCode = String(data[i][2] || binId).trim();

      bins.push({
        binId: binId,
        binCode: binCode,
        site: rackMeta.site,
        location: rackMeta.location,
        capacityUom: capacityUom,
        displayLabel: binCode + ' (' + capacityUom + ')'
      });
    }

    _putScriptCacheJson(cacheKey, bins, 300);
    return bins;
  });
}

function getBins() {
  return _withRequestCache(function () {
    const cacheKey = 'BINS_V6';
    const cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
    // Optional: check Rack sheet for hierarchy if Bins link to Racks
    const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);

    // Build Rack Map if Rack Sheet exists
    const rackMap = {};
    if (rackSheet) {
      const rackData = _getSheetValuesCached(rackSheet.getName());
      for (let i = 1; i < rackData.length; i++) {
        const rackId = String(rackData[i][0]).trim(); // Rack ID (Col A)
        rackMap[rackId] = {
          site: String(rackData[i][2] || '').trim(), // Site (Col C)
          location: String(rackData[i][3] || '').trim() // Location (Col D)
        };
      }
    }

    const data = _getSheetValuesCached(binSheet.getName());
    const headerMap = _getSheetHeaderMap(binSheet);
    const binMetaMap = _buildBinMasterMetaMap(binSheet, data, headerMap);
    const usageSummary = _getInventoryUsageSummaryByBin(binMetaMap);
    const bins = [];

    for (let i = 1; i < data.length; i++) {
      const binId = String(data[i][0] || '').trim();
      if (!binId) continue;
      const rackRef = String(data[i][1]).trim(); // Rack Ref (Col B)
      const metaRow = binMetaMap[binId] || {};
      const usage = usageSummary.byBin[binId] || { byUom: {}, hasInventory: false };
      const metrics = _getBinCapacityMetrics(metaRow, usage);
      const binCode = data[i][2]; // Bin Code (Col C)
      const maxCapacity = Number(metrics.maxCapacity || 0);
      const currentUsage = Number(metrics.currentUsage || 0);
      const availableCapacity = Number(metrics.availableCapacity || 0);
      const capacityUom = _normalizeCapacityUom(metaRow.declaredCapacityUom || metaRow.capacityUom || metrics.capacityUom || 'KG');

      // Resolve Site/Location from Rack Map or Defaults
      const rackMeta = rackMap[rackRef] || { site: 'Main', location: 'Main' };

      bins.push({
        binId: binId,
        binCode: binCode,
        site: rackMeta.site,
        location: rackMeta.location,
        maxCapacity: maxCapacity,
        currentUsage: currentUsage,
        availableCapacity: availableCapacity,
        capacityUom: capacityUom,
        displayLabel: String(binCode || binId).trim() + ' (' + capacityUom + ')',
        capacities: metaRow.capacities || {},
        capacityLines: metrics.capacityLines || [],
        utilizationPct: Number(metrics.utilizationPct || 0),
        capacityWarning: String(metrics.capacityWarning || '')
      });
    }
    _putScriptCacheJson(cacheKey, bins, 90);
    return bins;
  });
}

// Legacy getBatchByLotNo prototypes removed; keep composite/batch-id resolvers below.

function _formatDateForUi(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof value === 'string') {
    return value.split('T')[0];
  }
  return '';
}

// Resolve Batch_Master row using composite key (item_code + batch_number + lot_number).
function _resolveBatchCompositeRow(itemCode, batchNumber, lotNumber) {
  const code = String(itemCode || '').trim();
  const batchNo = String(batchNumber || '').trim();
  const lotNo = String(lotNumber || '').trim();
  if (!code || !batchNo || !lotNo) {
    return { status: 'INVALID', message: 'Item Code, Batch No, and Lot No are required' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BATCH);
  const data = _getSheetValuesCached(sheet.getName());
  const matches = [];
  const codeU = code.toUpperCase();
  const batchU = batchNo.toUpperCase();
  const lotU = lotNo.toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowCode = String(row[CONFIG.BATCH_COLS.ITEM_CODE] || '').trim().toUpperCase();
    const rowBatch = String(row[CONFIG.BATCH_COLS.BATCH_NUMBER] || '').trim().toUpperCase();
    const rowLot = String(row[CONFIG.BATCH_COLS.LOT_NUMBER] || '').trim().toUpperCase();
    if (rowCode === codeU && rowBatch === batchU && rowLot === lotU) {
      matches.push({ row: row, rowIndex: i + 1 });
    }
  }

  if (matches.length === 0) {
    return { status: 'NOT_FOUND', message: 'Batch not found for Item + Batch + Lot' };
  }
  if (matches.length > 1) {
    Logger.log(`Batch Master configuration error: multiple matches for ${code}/${batchNo}/${lotNo}`);
    return { status: 'MULTIPLE', message: 'Multiple batch matches found for Item + Batch + Lot' };
  }
  return { status: 'OK', row: matches[0].row, rowIndex: matches[0].rowIndex };
}

function _getItemDefaultQualityDays(itemCode) {
  const code = String(itemCode || '').trim();
  if (!code) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);
  if (!sheet) return 0;
  const data = _getSheetValuesCached(sheet.getName());
  if (data.length < 2) return 0;
  const cols = _getItemSheetColumnMap(sheet);
  const idx = cols.DEFAULT_QUALITY_DAYS;
  if (typeof idx !== 'number') return 0;
  const itemCodeIdx = cols.ITEM_CODE;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][itemCodeIdx] || '').trim().toUpperCase() === code.toUpperCase()) {
      const days = Number(data[i][idx]);
      if (isFinite(days) && days > 0) return days;
      break;
    }
  }
  return 0;
}

function _deriveQualityDate(itemCode, batchQualityDate) {
  if (batchQualityDate) return batchQualityDate;
  const days = _getItemDefaultQualityDays(itemCode);
  if (!days) return '';
  const dt = new Date();
  dt.setDate(dt.getDate() + days);
  return dt;
}

function _buildBatchMetaFromRow(row) {
  const expDate = row[CONFIG.BATCH_COLS.BATCH_EXP_DATE];
  const qualityDate = row[CONFIG.BATCH_COLS.QUALITY_DATE];
  return {
    batchId: String(row[CONFIG.BATCH_COLS.BATCH_ID] || '').trim(),
    batchNumber: String(row[CONFIG.BATCH_COLS.BATCH_NUMBER] || '').trim(),
    itemCode: String(row[CONFIG.BATCH_COLS.ITEM_CODE] || '').trim(),
    lotNumber: String(row[CONFIG.BATCH_COLS.LOT_NUMBER] || '').trim(),
    ginNo: String(row[CONFIG.BATCH_COLS.GIN_NO] || '').trim(),
    batchExpDate: expDate,
    qualityDate: qualityDate,
    qualityStatus: String(row[CONFIG.BATCH_COLS.QUALITY_STATUS] || '').trim().toUpperCase(),
    version: String(row[CONFIG.BATCH_COLS.VERSION] || '').trim(),
    versionParentId: String(row[CONFIG.BATCH_COLS.VERSION_PARENT_ID] || '').trim(),
    isVersion: row[CONFIG.BATCH_COLS.IS_VERSION]
  };
}

// Public resolver used by UI (returns formatted dates).
function resolveBatchComposite(itemCode, batchNumber, lotNumber) {
  return protect(function () {
    const res = _resolveBatchCompositeRow(itemCode, batchNumber, lotNumber);
    if (res.status === 'INVALID') throw new Error(res.message);
    if (res.status === 'NOT_FOUND') throw new Error('Batch not found for Item + Batch + Lot. Please verify Batch Master.');
    if (res.status === 'MULTIPLE') throw new Error('Multiple batch matches found for Item + Batch + Lot. Please contact admin.');
    const meta = _buildBatchMetaFromRow(res.row);
    return {
      batchId: meta.batchId,
      batchNumber: meta.batchNumber,
      itemCode: meta.itemCode,
      lotNumber: meta.lotNumber,
      ginNo: meta.ginNo || '',
      batchExpDate: _formatDateForUi(meta.batchExpDate),
      qualityDate: _formatDateForUi(meta.qualityDate),
      qualityStatus: meta.qualityStatus || 'PENDING',
      version: meta.version || ''
    };
  });
}

function _getBatchById(batchId, itemCode) {
  const info = _resolveBatchReference(itemCode, batchId);
  return (info && info.meta) ? info.meta : null;
}

function _getNextBatchMasterId(data) {
  let maxId = 0;
  const rows = Array.isArray(data) ? data : [];
  for (let i = 1; i < rows.length; i++) {
    const raw = String(rows[i][CONFIG.BATCH_COLS.BATCH_ID] || '').trim();
    if (!/^\d+$/.test(raw)) continue;
    const n = Number(raw);
    if (isFinite(n) && n > maxId) maxId = n;
  }
  return String(maxId > 0 ? (maxId + 1) : new Date().getTime());
}

function _getBatchMasterMaps() {
  _ensureRequestCache();
  if (_REQ_CACHE.batchMasterMaps) return _REQ_CACHE.batchMasterMaps;

  const maps = {
    byId: {},
    byItemAndId: {},
    byItemAndNumber: {}
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.BATCH);
  if (!sheet) {
    _REQ_CACHE.batchMasterMaps = maps;
    return maps;
  }

  const data = _getSheetValuesCached(sheet.getName());
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const itemCode = _getCanonicalItemCode(row[CONFIG.BATCH_COLS.ITEM_CODE]);
    const batchId = String(row[CONFIG.BATCH_COLS.BATCH_ID] || '').trim();
    const batchNumber = String(row[CONFIG.BATCH_COLS.BATCH_NUMBER] || batchId || '').trim();
    if (!itemCode || (!batchId && !batchNumber)) continue;

    const meta = _buildBatchMetaFromRow(row);
    meta.itemCode = itemCode;
    meta.batchId = batchId || batchNumber;
    meta.batchNumber = batchNumber || batchId;
    meta.rowIndex = i + 1;

    const codeU = itemCode.toUpperCase();
    const idU = String(meta.batchId || '').trim().toUpperCase();
    const noU = String(meta.batchNumber || '').trim().toUpperCase();

    if (idU) {
      maps.byId[idU] = meta;
      maps.byItemAndId[codeU + '||' + idU] = meta;
    }
    if (noU) {
      maps.byItemAndNumber[codeU + '||' + noU] = meta;
    }
  }

  _REQ_CACHE.batchMasterMaps = maps;
  return maps;
}

function _resolveBatchReference(itemCode, batchRef) {
  const code = _getCanonicalItemCode(itemCode);
  const ref = String(batchRef || '').trim();
  const acceptedKeys = {};
  if (ref) acceptedKeys[ref.toUpperCase()] = true;
  if (!code || !ref) {
    return {
      itemCode: code,
      batchId: ref,
      batchNumber: ref,
      acceptedKeys: acceptedKeys,
      meta: null
    };
  }

  const maps = _getBatchMasterMaps();
  const codeU = code.toUpperCase();
  const refU = ref.toUpperCase();
  const meta = maps.byItemAndId[codeU + '||' + refU] ||
    maps.byItemAndNumber[codeU + '||' + refU] ||
    null;

  if (meta) {
    if (meta.batchId) acceptedKeys[String(meta.batchId).trim().toUpperCase()] = true;
    if (meta.batchNumber) acceptedKeys[String(meta.batchNumber).trim().toUpperCase()] = true;
    return {
      itemCode: code,
      batchId: String(meta.batchId || ref).trim(),
      batchNumber: String(meta.batchNumber || ref).trim(),
      acceptedKeys: acceptedKeys,
      meta: meta
    };
  }

  return {
    itemCode: code,
    batchId: ref,
    batchNumber: ref,
    acceptedKeys: acceptedKeys,
    meta: null
  };
}

function _getBatchDisplayNumber(itemCode, batchRef) {
  const info = _resolveBatchReference(itemCode, batchRef);
  return String((info && info.batchNumber) || batchRef || '').trim();
}

function _getDefaultInwardExpiryDate() {
  return new Date(2030, 0, 31);
}

function _normalizeInwardExpiryDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    if (value.getTime() === 0) return _getDefaultInwardExpiryDate();
    return value;
  }
  if (typeof value === 'number') {
    if (!isFinite(value) || value <= 0) return _getDefaultInwardExpiryDate();
    const parsedNumberDate = new Date(value);
    if (!isNaN(parsedNumberDate.getTime())) {
      if (parsedNumberDate.getTime() === 0) return _getDefaultInwardExpiryDate();
      return parsedNumberDate;
    }
  }
  if (String(value || '').trim()) {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) {
      if (parsed.getTime() === 0) return _getDefaultInwardExpiryDate();
      return parsed;
    }
  }
  return _getDefaultInwardExpiryDate();
}

function _isSameCalendarDate(a, b) {
  if (!(a instanceof Date) || isNaN(a.getTime())) return false;
  if (!(b instanceof Date) || isNaN(b.getTime())) return false;
  return a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

function _isBatchMasterRowEmpty(row) {
  const width = CONFIG.BATCH_COLS.IS_VERSION + 1;
  const cells = Array.isArray(row) ? row : [];
  for (let i = 0; i < width; i++) {
    if (String(cells[i] || '').trim()) return false;
  }
  return true;
}

function _getNextBatchMasterWriteRow(data) {
  const rows = Array.isArray(data) ? data : [];
  for (let i = 1; i < rows.length; i++) {
    if (_isBatchMasterRowEmpty(rows[i])) return i + 1;
  }
  return rows.length + 1;
}

function _upsertBatchMasterFromInward(itemCode, batchNumber, fields) {
  const code = _getCanonicalItemCode(itemCode);
  const batchNo = String(batchNumber || '').trim();
  if (!code || !batchNo) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BATCH);
  const data = _getSheetValuesCached(sheet.getName());
  const codeU = code.toUpperCase();
  const batchU = batchNo.toUpperCase();
  const rawExpiry = fields && fields.rawExpiryDate;
  const hasExplicitExpiry = (rawExpiry instanceof Date && !isNaN(rawExpiry.getTime())) || !!String(rawExpiry || '').trim();
  const expiryDate = _normalizeInwardExpiryDate((fields && fields.expiryDate) || rawExpiry);
  const defaultExpiry = _getDefaultInwardExpiryDate();
  const lotNo = String((fields && fields.lotNo) || '').trim();
  const ginNo = String((fields && fields.ginNo) || '').trim();
  const version = String((fields && fields.version) || '').trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const rowCode = String(data[i][CONFIG.BATCH_COLS.ITEM_CODE] || '').trim().toUpperCase();
    const rowBatch = String(data[i][CONFIG.BATCH_COLS.BATCH_NUMBER] || data[i][CONFIG.BATCH_COLS.BATCH_ID] || '').trim().toUpperCase();
    if (rowCode !== codeU || rowBatch !== batchU) continue;

    const rowData = data[i].slice();
    let changed = false;
    let resolvedBatchId = String(rowData[CONFIG.BATCH_COLS.BATCH_ID] || '').trim();
    if (!resolvedBatchId || resolvedBatchId.toUpperCase() === batchU) {
      resolvedBatchId = _getNextBatchMasterId(data);
      rowData[CONFIG.BATCH_COLS.BATCH_ID] = resolvedBatchId;
      changed = true;
    }
    if (!String(rowData[CONFIG.BATCH_COLS.ITEM_CODE] || '').trim()) {
      rowData[CONFIG.BATCH_COLS.ITEM_CODE] = code;
      changed = true;
    }
    if (!String(rowData[CONFIG.BATCH_COLS.BATCH_NUMBER] || '').trim()) {
      rowData[CONFIG.BATCH_COLS.BATCH_NUMBER] = batchNo;
      changed = true;
    }
    const existingExpiry = rowData[CONFIG.BATCH_COLS.BATCH_EXP_DATE];
    const parsedExistingExpiry = existingExpiry ? new Date(existingExpiry) : null;
    if (!existingExpiry) {
      rowData[CONFIG.BATCH_COLS.BATCH_EXP_DATE] = expiryDate;
      changed = true;
    } else if (hasExplicitExpiry && parsedExistingExpiry instanceof Date && !isNaN(parsedExistingExpiry.getTime()) &&
      _isSameCalendarDate(parsedExistingExpiry, defaultExpiry) &&
      !_isSameCalendarDate(parsedExistingExpiry, expiryDate)) {
      rowData[CONFIG.BATCH_COLS.BATCH_EXP_DATE] = expiryDate;
      changed = true;
    }
    if (lotNo && !String(rowData[CONFIG.BATCH_COLS.LOT_NUMBER] || '').trim()) {
      rowData[CONFIG.BATCH_COLS.LOT_NUMBER] = lotNo;
      changed = true;
    }
    if (ginNo && !String(rowData[CONFIG.BATCH_COLS.GIN_NO] || '').trim()) {
      rowData[CONFIG.BATCH_COLS.GIN_NO] = ginNo;
      changed = true;
    }
    if (version && !String(rowData[CONFIG.BATCH_COLS.VERSION] || '').trim()) {
      rowData[CONFIG.BATCH_COLS.VERSION] = version;
      changed = true;
    }
    if (!String(rowData[CONFIG.BATCH_COLS.QUALITY_STATUS] || '').trim()) {
      rowData[CONFIG.BATCH_COLS.QUALITY_STATUS] = 'PENDING';
      changed = true;
    }

    if (changed) {
      const width = Math.max(rowData.length, CONFIG.BATCH_COLS.IS_VERSION + 1);
      while (rowData.length < width) rowData.push('');
      sheet.getRange(i + 1, 1, 1, width).setValues([rowData.slice(0, width)]);
      _setSheetCacheRow(CONFIG.SHEETS.BATCH, i + 1, rowData.slice(0, width));
      delete _REQ_CACHE.batchMasterMaps;
      _clearScriptCachesForSheet(CONFIG.SHEETS.BATCH);
    }
    return _buildBatchMetaFromRow(rowData);
  }

  const newBatchId = _getNextBatchMasterId(data);
  const newRow = [];
  newRow[CONFIG.BATCH_COLS.BATCH_ID] = newBatchId;
  newRow[CONFIG.BATCH_COLS.ITEM_CODE] = code;
  newRow[CONFIG.BATCH_COLS.BATCH_NUMBER] = batchNo;
  newRow[CONFIG.BATCH_COLS.BATCH_EXP_DATE] = expiryDate;
  newRow[CONFIG.BATCH_COLS.LOT_NUMBER] = lotNo;
  newRow[CONFIG.BATCH_COLS.GIN_NO] = ginNo;
  newRow[CONFIG.BATCH_COLS.QUALITY_DATE] = '';
  newRow[CONFIG.BATCH_COLS.QUALITY_STATUS] = 'PENDING';
  newRow[CONFIG.BATCH_COLS.VERSION] = version || '';
  newRow[CONFIG.BATCH_COLS.VERSION_PARENT_ID] = '';
  newRow[CONFIG.BATCH_COLS.IS_VERSION] = '';
  const width = CONFIG.BATCH_COLS.IS_VERSION + 1;
  while (newRow.length < width) newRow.push('');
  const writeRow = _getNextBatchMasterWriteRow(data);
  sheet.getRange(writeRow, 1, 1, width).setValues([newRow.slice(0, width)]);
  _setSheetCacheRow(CONFIG.SHEETS.BATCH, writeRow, newRow.slice(0, width));
  delete _REQ_CACHE.batchMasterMaps;
  _clearScriptCachesForSheet(CONFIG.SHEETS.BATCH);
  return _buildBatchMetaFromRow(newRow);
}

function getItems() {
  return _withRequestCache(function () {
    const maps = _getItemMasterMaps();
    const items = [];
    Object.keys(maps.codeToId).forEach(function (codeNorm) {
      const status = String(maps.codeToStatus[codeNorm] || 'ACTIVE').trim().toUpperCase();
      if (status && status !== 'ACTIVE') return;
      items.push({
        id: String(maps.codeToId[codeNorm] || ''),
        code: String(maps.codeDisplayByNorm[codeNorm] || codeNorm),
        name: String(maps.codeToName[codeNorm] || maps.codeDisplayByNorm[codeNorm] || codeNorm),
        group: String(maps.codeToGroup[codeNorm] || '')
      });
    });
    items.sort(function (a, b) { return String(a.code).localeCompare(String(b.code)); });
    return items;
  });
}

function getBatches() {
  return _withRequestCache(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.BATCH);
    if (!sheet) return [];

    const data = _getSheetValuesCached(sheet.getName());
    const batches = [];

    // Skip header
    for (let i = 1; i < data.length; i++) {
      batches.push({
        id: data[i][CONFIG.BATCH_COLS.BATCH_ID],
        itemCode: data[i][CONFIG.BATCH_COLS.ITEM_CODE],
        number: data[i][CONFIG.BATCH_COLS.BATCH_NUMBER],
        expDate: data[i][CONFIG.BATCH_COLS.BATCH_EXP_DATE]
      });
    }
    return batches;
  });
}

function getBatchesForItem(itemCode) {
  return protect(function () {
    const code = String(itemCode || '').trim().toUpperCase();
    if (!code) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.BATCH);
    if (!sheet) return [];

    const data = _getSheetValuesCached(sheet.getName());
    const seen = {};
    const batches = [];

    for (let i = 1; i < data.length; i++) {
      const rowCode = String(data[i][CONFIG.BATCH_COLS.ITEM_CODE] || '').trim().toUpperCase();
      if (rowCode !== code) continue;

      const batchNo = String(
        data[i][CONFIG.BATCH_COLS.BATCH_NUMBER] ||
        data[i][CONFIG.BATCH_COLS.BATCH_ID] ||
        ''
      ).trim();
      if (!batchNo || seen[batchNo]) continue;

      seen[batchNo] = true;
      batches.push(batchNo);
    }

    batches.sort(function (a, b) { return String(a).localeCompare(String(b)); });
    return batches;
  });
}

function getRacks() {
  return _withRequestCache(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
    const binSheet = ss.getSheetByName(CONFIG.SHEETS.BIN);

    if (!rackSheet || !binSheet) return [];

    // Load all data
    const rackData = _getSheetValuesCached(rackSheet.getName());
    const binData = _getSheetValuesCached(binSheet.getName());
    const binColMap = _getSheetHeaderMap(binSheet);
    const binMetaMap = _buildBinMasterMetaMap(binSheet, binData, binColMap);
    const usageSummary = _getInventoryUsageSummaryByBin(binMetaMap);

    // Build bin map: bin_id -> { max_capacity, items[] }
    const binMap = {};
    for (let i = 1; i < binData.length; i++) {
      const binId = String(binData[i][0] || '').trim();
      if (!binId) continue;
      const meta = binMetaMap[binId] || {};
      const binCode = meta.binCode || binData[i][2] || '';
      const usage = usageSummary.byBin[binId] || { byUom: {}, hasInventory: false };
      const metrics = _getBinCapacityMetrics(meta, usage);
      const maxCapacity = Number(metrics.maxCapacity || 0);
      binMap[binId] = {
        binId: binId,
        code: binCode,  // Frontend expects 'code' not 'binCode'
        binCode: binCode,
        maxCapacity: maxCapacity,
        currentUsage: Number(metrics.currentUsage || 0),
        capacityUom: _normalizeCapacityUom(meta.declaredCapacityUom || meta.capacityUom || metrics.capacityUom || 'KG'),
        capacities: meta.capacities || {},
        capacityLines: metrics.capacityLines || [],
        utilizationPct: Number(metrics.utilizationPct || 0),
        capacityWarning: String(metrics.capacityWarning || ''),
        items: []
      };
    }

    // Build rack-to-bins map: rack_id -> [bin_ids]
    const rackToBins = {};
    for (let binId in binMap) {
      const bin = binMap[binId];
      // Find rack that contains this bin (simplified: assume bin code starts with rack code or lookup)
      // For now, we'll populate this from bin data column 1 (rack reference)
    }

    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    const invData = invSheet ? _getSheetValuesCached(invSheet.getName()) : [];
    if (invData.length > 0) {
      for (let i = 1; i < invData.length; i++) {
        const row = invData[i];
        const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
        const itemCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
        const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        const qaStatus = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').trim();

        if (binId && binMap[binId] && itemCode && qty > 0) {
          binMap[binId].items.push({
            itemCode: itemCode,
            itemName: itemCode, // Use code as fallback name
            qty: qty,
            qaStatus: qaStatus
          });
        }
      }
    }

    // Build racks array
    const racks = [];
    for (let i = 1; i < rackData.length; i++) {
      const rackId = String(rackData[i][0] || '').trim();
      if (!rackId) continue;
      const rackCode = rackData[i][1] || '';
      const site = String(rackData[i][2] || 'Unit 1').trim();
      const location = String(rackData[i][3] || 'Main').trim();
      const maxCapacity = Number(rackData[i][4]) || 0;

      // Find all bins for this rack
      const rackBins = [];
      for (let binId in binMap) {
        // Simple approach: if bin's rack reference matches rack ID
        const binRow = binData.find((row, idx) => {
          if (idx === 0) return false; // Skip header
          return String(row[0] || '').trim() === binId && String(row[1] || '').trim() === rackId;
        });
        if (binRow) {
          const binObj = binMap[binId];
          // Add rack info to bin for modal display
          binObj.rackCode = rackCode;
          binObj.site = site;
          binObj.location = location;
          binObj.available = Number((((binObj.capacityLines || [])[0] || {}).availableCapacity) || binObj.availableCapacity || Math.max(0, Number(binObj.maxCapacity || 0) - Number(binObj.currentUsage || 0)));
          rackBins.push(binObj);
        }
      }

      racks.push({
        rackCode: rackCode,
        rackId: rackId,
        site: site,
        location: location,
        maxCapacity: maxCapacity,
        currentUsage: rackBins.reduce(function (sum, b) {
          return sum + Number(usageSummary.kgByBin[b.binId] || 0);
        }, 0),
        bins: rackBins
      });
    }

    return racks;
  });
}

function getAvailableStock(itemCode, batchId, binId) {
  const code = String(itemCode || '').trim();
  const allRows = getInventoryReadView({ itemCode: code, batchId: batchId, binId: binId });
  if (!allRows || allRows.length === 0) {
    _logInventoryVisibility('Available stock', code, batchId, binId, [], [], 'No matching inventory rows');
    return 0;
  }
  return allRows.reduce((acc, row) => acc + (Number(row.quantity) || 0), 0);
}

// =====================================================
// 4. TRANSACTION ENGINE (CORE)
// =====================================================

function withScriptLock(callback) {
  _beginRequestCache();
  const lock = LockService.getScriptLock();
  let hasLock = false;
  try {
    lock.waitLock(30000);
    hasLock = true;
    return callback();
  } catch (e) {
    Logger.log('Transaction Error: ' + e.toString());
    if (!hasLock) {
      throw new Error('System is busy. Please try again.');
    }
    throw e;
  } finally {
    if (hasLock) lock.releaseLock();
    _endRequestCache();
  }
}

function _getSheetOrThrow(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Try trim and fallback name if common variant
    const sheets = ss.getSheets();
    const cleanTarget = sheetName.trim().toUpperCase();
    const found = sheets.find(s => s.getName().trim().toUpperCase() === cleanTarget);
    if (found) {
      if (sheetName === CONFIG.SHEETS.PRODUCTION_LEDGER) {
        _ensureProductionLedgerSchema(found);
      }
      if (sheetName === CONFIG.SHEETS.INVENTORY) {
        _ensureInventoryDateColumnsSchema(found);
      }
      return found;
    }

    throw new Error("CRITICAL ERROR: Sheet '" + sheetName + "' is missing. Please check your Google Sheet names.");
  }
  if (sheetName === CONFIG.SHEETS.PRODUCTION_LEDGER) {
    _ensureProductionLedgerSchema(sheet);
  }
  if (sheetName === CONFIG.SHEETS.INVENTORY) {
    _ensureInventoryDateColumnsSchema(sheet);
  }
  return sheet;
}

/** ==== CHANGE =====
 * CRITICAL FIX: Inventory identity is item_code_cache + batch_id + bin_id.
 * _readInventoryState MUST use item_code_cache only (item_id is legacy).
 * EXECUTION ONLY - FIFO + QA enforced here.
 */
function _readInventoryState(itemCode, batchId, binId, version) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(sheet.getName());
  const results = [];

  const rawTarget = String(itemCode || '').trim();
  const rawTargetU = rawTarget.toUpperCase();
  let targetCode = rawTargetU;
  let targetId = '';
  try {
    const byCode = getItemByCodeCached(rawTarget);
    if (byCode) {
      targetCode = String(byCode.code || rawTarget).trim().toUpperCase();
      targetId = String(byCode.id || '').trim().toUpperCase();
    } else {
      const byId = _getItemByIdCached(rawTarget);
      if (byId) {
        targetCode = String(byId.code || rawTarget).trim().toUpperCase();
        targetId = String(byId.id || '').trim().toUpperCase();
      }
    }
  } catch (e) {
    // Keep raw fallback keys
  }
  const itemIdToCode = {};
  const searchBatch = String(batchId || '').trim().toUpperCase();
  const batchInfo = searchBatch ? _resolveBatchReference(targetCode || rawTarget, batchId) : null;
  const acceptedBatchKeys = batchInfo ? (batchInfo.acceptedKeys || {}) : null;

  // Resolve bin identifier to both ID and Code (handles legacy rows storing bin code)
  let searchBinRaw = String(binId || '').trim();
  let searchBinId = searchBinRaw;
  let searchBinCode = searchBinRaw;
  if (searchBinRaw) {
    try {
      const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
      const binData = _getSheetValuesCached(binSheet.getName());
      for (let i = 1; i < binData.length; i++) {
        const id = String(binData[i][0] || '').trim();
        const code = String(binData[i][2] || '').trim();
        if (id && (id === searchBinRaw || code === searchBinRaw)) {
          searchBinId = id || searchBinRaw;
          searchBinCode = code || searchBinRaw;
          break;
        }
      }
    } catch (e) {
      // Ignore and fall back to raw bin identifier
    }
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const rowResolvedCode = String(
      _resolveInventoryItemCode(
        row,
        row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE],
        row[CONFIG.INVENTORY_COLS.ITEM_ID],
        itemIdToCode
      ) || ''
    ).trim().toUpperCase();
    const rowRawCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
    const rowRawId = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();

    const matchesCode = !!targetCode && (rowResolvedCode === targetCode || rowRawCode === targetCode);
    const matchesId = !!targetId && (rowRawId === targetId || rowRawCode === targetId || rowResolvedCode === targetId);
    const matchesRaw = !!rawTargetU && (rowResolvedCode === rawTargetU || rowRawCode === rawTargetU || rowRawId === rawTargetU);
    const itemMatch = matchesCode || matchesId || matchesRaw;

    const rowBatch = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim().toUpperCase();
    const batchMatch = searchBatch ? !!acceptedBatchKeys[rowBatch] : true;

    const rowBin = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim().toUpperCase();
    const searchBinIdU = String(searchBinId || '').trim().toUpperCase();
    const searchBinCodeU = String(searchBinCode || '').trim().toUpperCase();
    const binMatch = (rowBin === searchBinIdU) || (rowBin === searchBinCodeU);

    if (itemMatch && batchMatch && binMatch) {
      if (version && String(row[CONFIG.INVENTORY_COLS.VERSION]) !== String(version)) {
        continue;
      }
      results.push({
        rowIndex: i + 1,
        currentQty: Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0,
        rowData: row
      });
    }
  }

  return results;
}

function _toTimestamp(value, fallback) {
  if (value instanceof Date) return value.getTime();
  if (typeof value === 'number') return value;
  const parsed = Date.parse(value);
  if (!isNaN(parsed)) return parsed;
  return fallback;
}

function _getInventoryCreatedTs(rowData) {
  const invId = String(rowData[CONFIG.INVENTORY_COLS.INVENTORY_ID] || '');
  const match = invId.match(/INV-(\d+)-/);
  if (match && match[1]) return Number(match[1]);
  return Number.NaN;
}

// ======== ADDED ==========

function _assertFifoAnchors(rows, context) {
  if (!rows || rows.length === 0) return;
  const unanchored = rows.filter(r => {
    const qd = r.rowData[CONFIG.INVENTORY_COLS.QUALITY_DATE];
    const gin = String(r.rowData[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim();
    const invId = String(r.rowData[CONFIG.INVENTORY_COLS.INVENTORY_ID] || '').trim();
    return !qd && !gin && !invId;
  });
  if (unanchored.length > 0) {
    Logger.log(`[FIFO_WARN] ${context}: ${unanchored.length} row(s) have no FIFO anchor (no quality_date, no GIN_NO, no INV_ID). Sort order is best-effort for these rows.`);
  }
}

function _sortInventoryRowsFifo(rows) {
  _assertFifoAnchors(rows, 'FIFO');
  const mapped = (rows || []).map((r, idx) => {
    const qd = r.rowData[CONFIG.INVENTORY_COLS.QUALITY_DATE];
    const qdTs = _toTimestamp(qd, NaN);
    const gin = String(r.rowData[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim();
    return { r: r, idx: idx, qdTs: qdTs, gin: gin, invTs: _getInventoryCreatedTs(r.rowData) };
  });
  mapped.sort((a, b) => {
    const aHasDate = isFinite(a.qdTs);
    const bHasDate = isFinite(b.qdTs);
    if (aHasDate && bHasDate && a.qdTs !== b.qdTs) return a.qdTs - b.qdTs;
    if (aHasDate !== bHasDate) return aHasDate ? -1 : 1;
    if (a.gin && b.gin && a.gin !== b.gin) return a.gin.localeCompare(b.gin);
    if (!!a.gin !== !!b.gin) return a.gin ? -1 : 1;
    if (isFinite(a.invTs) && isFinite(b.invTs) && a.invTs !== b.invTs) return a.invTs - b.invTs;
    if (a.idx !== b.idx) return a.idx - b.idx;
    return 0;
  });
  return mapped.map(m => m.r);
}


/**
 * Returns the FIFO-primary version from ordered inventory rows.
 * If rows span multiple versions (valid FIFO cross-version split), returns the
 * version of the FIRST (oldest/FIFO) consumed row as primary, and logs the split.
 * NEVER throws for multi-version FIFO - that is a valid business scenario.
 * Only throws when a row has no version value at all (data integrity error).
 *
 * Bug fix: old code threw 'Multiple versions detected' when FIFO qty needed spanned
 * more than one inventory version row. Example: Required 1100, V1=1000, V2=500 in
 * the same bin -> FIFO picks rows from both -> previously threw, now succeeds.
 *
 * @param {Array}  rows    - FIFO-ordered inventory rows (from _validateQualityEligibility)
 * @param {string} context - label used for logging only
 * @returns {string}       - primary version string (e.g. 'V1')
 */
function _getSingleInventoryVersion(rows, context) {
  if (!rows || rows.length === 0) {
    throw new Error(context + ': No inventory rows provided for version resolution');
  }
  var versionList = [];
  for (var i = 0; i < rows.length; i++) {
    var v = String(rows[i].rowData[CONFIG.INVENTORY_COLS.VERSION] || '').trim();
    if (!v) {
      throw new Error(context + ': Inventory row at FIFO position ' + (i + 1) +
        ' is missing VERSION field - data integrity issue. Please contact admin.');
    }
    versionList.push(v.toUpperCase());
  }
  var uniqueVersions = [];
  var seen = {};
  for (var j = 0; j < versionList.length; j++) {
    if (!seen[versionList[j]]) { seen[versionList[j]] = true; uniqueVersions.push(versionList[j]); }
  }
  if (uniqueVersions.length > 1) {
    // Multi-version FIFO - VALID. FIFO order is already enforced by the caller.
    // Log for audit trail. Primary = oldest version (first FIFO row consumed).
    Logger.log('[VERSION_SPLIT] ' + context +
      ' - FIFO consumption spans ' + uniqueVersions.length + ' versions: [' +
      uniqueVersions.join(', ') + ']. Primary (oldest FIFO): ' + versionList[0]);
  }
  // Always return the version of the first (oldest FIFO) row - used as primary for ledger.
  return versionList[0];
}

/**
 * Identity enforcement: Item Code + Batch ID are required together.
 */
function _assertItemCodeBatch(itemCode, batchId, context) {
  const code = String(itemCode || '').trim();
  const batch = String(batchId || '').trim();
  if (!code || !batch) {
    throw new Error(`${context} requires Item Code and Batch No`);
  }
  return { itemCode: code, batchId: batch };
}

// Normalize QA status to the enforced lifecycle.
function _normalizeQaStatus(value) {
  if (!value) return 'PENDING';

  const raw = String(value).trim().toUpperCase();

  if (
    raw === 'PENDING' ||
    raw === 'HOLD' ||
    raw === 'REJECTED' ||
    raw === 'APPROVED' ||
    raw === 'OVERRIDDEN' ||
    raw === 'CLOSED'
  ) {
    return raw;
  }

  return 'PENDING';
}


function _logInventoryVisibility(context, itemCode, batchId, binId, allRows, rowsWithQty, reason) {
  const code = String(itemCode || '').trim();
  const batch = String(batchId || '').trim();
  const bin = String(binId || '').trim();
  const totalRows = Array.isArray(allRows) ? allRows.length : 0;
  const keptRows = Array.isArray(rowsWithQty) ? rowsWithQty.length : 0;
  const removed = [];

  if (Array.isArray(allRows)) {
    allRows.forEach(r => {
      const qty = Number(r.currentQty) || 0;
      if (qty > 0) return;
      removed.push({ rowIndex: r.rowIndex, qty: qty, reason: 'qty<=0' });
    });
  }

  Logger.log(`[INV_VIS] ${context} item_code=${code} batch_id=${batch} bin_id=${bin} total_rows=${totalRows} kept_rows=${keptRows} reason=${String(reason || '').trim()}`);
  if (removed.length > 0) {
    Logger.log(`[INV_VIS_REMOVED] ${context} ${JSON.stringify(removed)}`);
  }
}

function _resolveBatchIdFromInventory(itemCode, batchId, binId, context) {
  const code = String(itemCode || '').trim();
  if (!code) throw new Error(`${context} requires Item Code`);
  const requested = String(batchId || '').trim();
  const requestedInfo = requested ? _resolveBatchReference(code, requested) : null;
  const requestedKeys = requestedInfo ? (requestedInfo.acceptedKeys || {}) : {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(sheet.getName());

  // Resolve bin identifier to both ID and Code (handles legacy rows storing bin code)
  let searchBinRaw = String(binId || '').trim();
  let searchBinId = searchBinRaw;
  let searchBinCode = searchBinRaw;
  if (searchBinRaw) {
    try {
      const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
      const binData = _getSheetValuesCached(binSheet.getName());
      for (let i = 1; i < binData.length; i++) {
        const id = String(binData[i][0] || '').trim();
        const binCode = String(binData[i][2] || '').trim();
        if (id && (id === searchBinRaw || binCode === searchBinRaw)) {
          searchBinId = id || searchBinRaw;
          searchBinCode = binCode || searchBinRaw;
          break;
        }
      }
    } catch (e) { }
  }

  const codeU = code.toUpperCase();
  const batchMap = {};
  let matchedRequested = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
    if (rowCode !== codeU) continue;

    const rowBin = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim().toUpperCase();
    if (searchBinRaw) {
      const binMatch = (rowBin === String(searchBinId).trim().toUpperCase()) ||
        (rowBin === String(searchBinCode).trim().toUpperCase());
      if (!binMatch) continue;
    }

    const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (qty <= 0) continue;

    const rowBatchRaw = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
    const rowBatchU = rowBatchRaw.toUpperCase();
    if (!rowBatchRaw) continue;

    const resolvedRow = _resolveBatchReference(code, rowBatchRaw);
    const canonicalBatchId = String((resolvedRow && resolvedRow.batchId) || rowBatchRaw).trim();
    const canonicalBatchU = canonicalBatchId.toUpperCase();

    if (requested && (requestedKeys[rowBatchU] || requestedKeys[canonicalBatchU])) matchedRequested = true;
    if (!batchMap[canonicalBatchU]) batchMap[canonicalBatchU] = canonicalBatchId;
  }

  if (requested && matchedRequested) {
    return String((requestedInfo && requestedInfo.batchId) || requested).trim();
  }
  const keys = Object.keys(batchMap);
  if (keys.length === 1) return batchMap[keys[0]];
  if (keys.length > 1) {
    const binLabel = searchBinRaw ? searchBinRaw : 'ANY';
    throw new Error(`Multiple batch IDs found for ${code} in Bin ${binLabel}. Use canonical batch_id.`);
  }
  return requested;
}

function getOperationPolicy(operationType) {
  const op = String(operationType || '').trim().toUpperCase();
  if (op === 'TRANSFER') {
    return { allowPending: false, allowHold: false, allowRejected: false, allowOverridden: false, fifoRequired: true };
  }
  if (op === 'CONSUMPTION') {
    return { allowPending: false, allowHold: false, allowRejected: false, allowOverridden: true, fifoRequired: true };
  }
  if (op === 'DISPATCH') {
    return { allowPending: false, allowHold: false, allowRejected: false, allowOverridden: true, fifoRequired: true };
  }
  return { allowPending: false, allowHold: false, allowRejected: false, allowOverridden: false, fifoRequired: true };
}

/**
 * Shared FIFO eligibility evaluator.
 * Used by Dispatch, Consumption, Transfer, and QA Summary.
 */
function _evaluateFifoEligibility(itemCode, batchId, rows, requiredQty, policy) {
  _assertItemCodeBatch(itemCode, batchId, 'Eligibility check');
  const ordered = _sortInventoryRowsFifo(rows || []);
  const policyObj = policy || getOperationPolicy('DISPATCH');
  const allowPending = policyObj.allowPending === true;
  const allowHold = policyObj.allowHold === true;
  const allowRejected = policyObj.allowRejected === true;
  const allowOverridden = policyObj.allowOverridden === true;
  const need = Number(requiredQty);
  const hasRequired = isFinite(need) && need > 0;

  let total = 0;
  let approved = 0;
  let pending = 0;
  let rejected = 0;
  let overridden = 0;
  let closed = 0;
  let hold = 0;
  let fifoEligible = 0;
  let firstBlockStatus = '';

  for (const row of ordered) {
    const qty = Number(row.currentQty) || 0;
    if (qty <= 0) continue;
    total += qty;
    let status = _normalizeQaStatus(row.rowData[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');

    if (status === 'APPROVED') approved += qty;
    else if (status === 'HOLD') hold += qty;
    else if (status === 'REJECTED') rejected += qty;
    else if (status === 'OVERRIDDEN') overridden += qty;
    else if (status === 'CLOSED') closed += qty;
    else pending += qty;

    if (!firstBlockStatus) {
      const allowed = (status === 'APPROVED') ||
        (status === 'OVERRIDDEN' && allowOverridden) ||
        (status === 'PENDING' && allowPending) ||
        (status === 'HOLD' && allowHold) ||
        (status === 'REJECTED' && allowRejected);
      if (allowed) {
        fifoEligible += qty;
      } else {
        firstBlockStatus = status;
      }
    }
  }

  let blockedByRejected = false;
  let blockedByPending = false;
  let blockedByHold = false;
  let blockedByClosed = false;
  let blockedByOverridden = false;
  if (hasRequired) {
    if (fifoEligible < need) {
      blockedByRejected = firstBlockStatus === 'REJECTED';
      blockedByPending = firstBlockStatus === 'PENDING';
      blockedByHold = firstBlockStatus === 'HOLD';
      blockedByClosed = firstBlockStatus === 'CLOSED';
      blockedByOverridden = firstBlockStatus === 'OVERRIDDEN';
    }
  } else {
    blockedByRejected = firstBlockStatus === 'REJECTED';
    blockedByPending = firstBlockStatus === 'PENDING';
    blockedByHold = firstBlockStatus === 'HOLD';
    blockedByClosed = firstBlockStatus === 'CLOSED';
    blockedByOverridden = firstBlockStatus === 'OVERRIDDEN';
  }

  const fulfillable = hasRequired ? fifoEligible >= need : fifoEligible > 0;
  let failureReason = '';

  if (!fulfillable) {
    if (total <= 0) {
      failureReason = 'No stock exists';
    } else if (blockedByRejected || blockedByHold || blockedByPending) {
      failureReason = `Stock is under QA HOLD / Pending approval (Batch: ${batchId})`;
    } else if (blockedByClosed) {
      failureReason = `Stock closed (Batch: ${batchId})`;
    } else if (blockedByOverridden) {
      failureReason = `Stock overridden - not eligible for this operation (Batch: ${batchId})`;
    } else {
      failureReason = 'FIFO quantity insufficient';
    }
  }

  return {
    totalQty: total,
    fifoEligibleQty: fifoEligible,
    approvedQty: approved,
    pendingQty: pending,
    rejectedQty: rejected,
    holdQty: hold,
    overriddenQty: overridden,
    closedQty: closed,
    blockedByPending: blockedByPending,
    blockedByHold: blockedByHold,
    blockedByRejected: blockedByRejected,
    blockedByClosed: blockedByClosed,
    blockedByOverridden: blockedByOverridden,
    fulfillable: fulfillable,
    failureReason: failureReason,
    orderedRows: ordered
  };
}

/**
 * Dispatch / Consumption eligibility (QA must be APPROVED for FIFO traversal).
 */
function _validateQualityEligibility(itemCode, batchId, rows, requiredQty, operationType, options) {
  const opts = options || {};
  const allowUnapproved = opts.allowUnapproved === true;
  const policy = getOperationPolicy(operationType || 'DISPATCH');
  if (allowUnapproved) {
    policy.allowPending = true;
    policy.allowHold = true;
    policy.allowRejected = true;
    policy.allowOverridden = true;
  }
  const evalRes = _evaluateFifoEligibility(itemCode, batchId, rows, requiredQty, policy);
  if (!evalRes.fulfillable) {
    throw new Error(evalRes.failureReason);
  }
  return evalRes;
}

/**
 * Transfer eligibility (QA_PENDING/HOLD blocked).
 */
function _validateTransferEligibility(itemCode, batchId, rows, requiredQty) {
  const policy = getOperationPolicy('TRANSFER');
  const evalRes = _evaluateFifoEligibility(itemCode, batchId, rows, requiredQty, policy);
  if (!evalRes.fulfillable) {
    throw new Error(evalRes.failureReason);
  }
  return evalRes;
}

function _buildQaOverrideMeta(user, movementType, reason, enabled) {
  if (!enabled) return null;
  return {
    qa_override: true,
    qa_override_by: user && user.email ? user.email : '',
    qa_override_reason: String(reason || '').trim(),
    qa_override_timestamp: new Date().toISOString(),
    qa_override_movement_type: movementType
  };
}

function _formatRemarksWithOverride(remarks, qaMeta) {
  if (!qaMeta) return remarks;
  const base = String(remarks || '').trim();
  const metaStr = JSON.stringify(qaMeta);
  return base ? `${base} | QA_OVERRIDE ${metaStr}` : `QA_OVERRIDE ${metaStr}`;
}

// Centralized QA override handler (state + audit only, never movement).
function _applyQaOverrideForRow(row, context, qaStatusChanges, qaEvents) {
  if (!row || !context) return;
  const currentStatus = _normalizeQaStatus(row.rowData[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');
  if (currentStatus === 'APPROVED' || currentStatus === 'OVERRIDDEN' || currentStatus === 'CLOSED') return;
  const now = new Date();
  qaStatusChanges.push({
    // FIX: inventory_id is source of truth
    // RowIndex-only overrides are unsafe when batch/lot resolution skips rows; keep inventory_id for exact updates.
    inventoryId: String(row.rowData[CONFIG.INVENTORY_COLS.INVENTORY_ID] || ''),
    rowIndex: row.rowIndex,
    prevStatus: currentStatus || 'PENDING',
    prevDate: row.rowData[CONFIG.INVENTORY_COLS.QUALITY_DATE],
    newStatus: 'OVERRIDDEN',
    newDate: now
  });
  qaEvents.push({
    inventoryId: String(row.rowData[CONFIG.INVENTORY_COLS.INVENTORY_ID] || ''),
    itemCode: context.itemCode || '',
    batchId: context.batchId || '',
    binId: context.binId || '',
    prevStatus: currentStatus || 'PENDING',
    newStatus: 'OVERRIDDEN',
    action: 'OVERRIDE',
    overrideReason: context.overrideReason || '',
    overriddenBy: context.userEmail || '',
    overriddenAt: now,
    remarks: context.remarks || ''
  });
}

function _rollbackInventoryChanges(changes, sheet) {
  if (!changes || changes.length === 0) return;
  changes.forEach(ch => {
    if (ch.action === 'create') {
      try { sheet.deleteRow(ch.rowIndex); } catch (e) { }
      _deleteSheetCacheRow(CONFIG.SHEETS.INVENTORY, ch.rowIndex);
      return;
    }
    sheet.getRange(ch.rowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1).setValue(ch.prevQty);
    if (typeof ch.prevUpdated !== 'undefined') {
      sheet.getRange(ch.rowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1).setValue(ch.prevUpdated);
    }
    _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, ch.rowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1, ch.prevQty);
    if (typeof ch.prevUpdated !== 'undefined') {
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, ch.rowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1, ch.prevUpdated);
    }
  });
}

function _rollbackProductionLedger(change) {
  if (!change) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
  if (change.action === 'create') {
    try { sheet.deleteRow(change.rowIndex); } catch (e) { }
    _deleteSheetCacheRow(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex);
    return;
  }
  if (change.action === 'update') {
    const cols = CONFIG.PRODUCTION_LEDGER_COLS;
    const data = _getSheetValuesCached(sheet.getName());
    const rowData = (data[change.rowIndex - 1] || []);
    const statusIdx = _detectProductionLedgerStatusIdx(rowData, cols.STATUS);
    const updatedIdx = _detectProductionLedgerUpdatedIdx(rowData, statusIdx, cols.LAST_UPDATED);
    sheet.getRange(change.rowIndex, cols.QTY_ISSUED + 1).setValue(change.prevIssued);
    sheet.getRange(change.rowIndex, cols.QTY_RETURNED + 1).setValue(change.prevReturned);
    sheet.getRange(change.rowIndex, cols.QTY_REJECTED + 1).setValue(change.prevRejected || 0);
    sheet.getRange(change.rowIndex, cols.NET_OUTSTANDING + 1).setValue(change.prevNet);
    sheet.getRange(change.rowIndex, statusIdx + 1).setValue(change.prevStatus);
    sheet.getRange(change.rowIndex, updatedIdx + 1).setValue(change.prevUpdated);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, cols.QTY_ISSUED + 1, change.prevIssued);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, cols.QTY_RETURNED + 1, change.prevReturned);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, cols.QTY_REJECTED + 1, change.prevRejected || 0);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, cols.NET_OUTSTANDING + 1, change.prevNet);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, statusIdx + 1, change.prevStatus);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, updatedIdx + 1, change.prevUpdated);
    if (typeof change.prevBatch !== 'undefined') {
      sheet.getRange(change.rowIndex, cols.BATCH_ID + 1).setValue(change.prevBatch);
      _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, cols.BATCH_ID + 1, change.prevBatch);
    }
    if (typeof change.prevVersion !== 'undefined') {
      sheet.getRange(change.rowIndex, cols.VERSION + 1).setValue(change.prevVersion);
      _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, change.rowIndex, cols.VERSION + 1, change.prevVersion);
    }
  }
}

function _applyQaStatusChanges(sheet, changes) {
  if (!changes || changes.length === 0) return;
  changes.forEach(ch => {
    // FIX: inventory_id is source of truth
    // RowIndex-only updates are unsafe when batch/lot resolution skips rows; prefer inventory_id when present.
    let targetRow = ch.rowIndex;
    if (ch.inventoryId) {
      const byId = _findInventoryRowIndexById(ch.inventoryId);
      if (byId > 0) targetRow = byId;
    }
    sheet.getRange(targetRow, CONFIG.INVENTORY_COLS.QUALITY_STATUS + 1).setValue(ch.newStatus);
    if (typeof ch.newDate !== 'undefined') {
      sheet.getRange(targetRow, CONFIG.INVENTORY_COLS.QUALITY_DATE + 1).setValue(ch.newDate);
    }
    _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, targetRow, CONFIG.INVENTORY_COLS.QUALITY_STATUS + 1, ch.newStatus);
    if (typeof ch.newDate !== 'undefined') {
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, targetRow, CONFIG.INVENTORY_COLS.QUALITY_DATE + 1, ch.newDate);
    }
  });
}

function _rollbackQaStatusChanges(sheet, changes) {
  if (!changes || changes.length === 0) return;
  changes.forEach(ch => {
    // FIX: inventory_id is source of truth
    // Rollback must target inventory_id to avoid desync when batch/lot resolution skips rows.
    let targetRow = ch.rowIndex;
    if (ch.inventoryId) {
      const byId = _findInventoryRowIndexById(ch.inventoryId);
      if (byId > 0) targetRow = byId;
    }
    sheet.getRange(targetRow, CONFIG.INVENTORY_COLS.QUALITY_STATUS + 1).setValue(ch.prevStatus);
    if (typeof ch.prevDate !== 'undefined') {
      sheet.getRange(targetRow, CONFIG.INVENTORY_COLS.QUALITY_DATE + 1).setValue(ch.prevDate);
    }
    _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, targetRow, CONFIG.INVENTORY_COLS.QUALITY_STATUS + 1, ch.prevStatus);
    if (typeof ch.prevDate !== 'undefined') {
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, targetRow, CONFIG.INVENTORY_COLS.QUALITY_DATE + 1, ch.prevDate);
    }
  });
}

function _appendQaEventsBatch(events) {
  // QA_Events is audit-only. Never use it for stock quantities or FIFO eligibility.
  if (!events || events.length === 0) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.QA_EVENTS);
  const rows = events.map(e => {
    const row = [];
    row[CONFIG.QA_EVENTS_COLS.EVENT_ID] = 'QAE-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    row[CONFIG.QA_EVENTS_COLS.TIMESTAMP] = new Date();
    row[CONFIG.QA_EVENTS_COLS.INVENTORY_ID] = e.inventoryId || '';
    row[CONFIG.QA_EVENTS_COLS.ITEM_CODE] = e.itemCode || '';
    row[CONFIG.QA_EVENTS_COLS.BATCH_ID] = e.batchId || '';
    row[CONFIG.QA_EVENTS_COLS.BIN_ID] = e.binId || '';
    row[CONFIG.QA_EVENTS_COLS.PREV_STATUS] = e.prevStatus || '';
    row[CONFIG.QA_EVENTS_COLS.NEW_STATUS] = e.newStatus || '';
    row[CONFIG.QA_EVENTS_COLS.ACTION] = e.action || '';
    row[CONFIG.QA_EVENTS_COLS.OVERRIDE_REASON] = e.overrideReason || '';
    row[CONFIG.QA_EVENTS_COLS.OVERRIDDEN_BY] = e.overriddenBy || '';
    row[CONFIG.QA_EVENTS_COLS.OVERRIDDEN_AT] = e.overriddenAt || '';
    row[CONFIG.QA_EVENTS_COLS.MOVEMENT_ID] = e.movementId || '';
    row[CONFIG.QA_EVENTS_COLS.REMARKS] = e.remarks || '';
    row[CONFIG.QA_EVENTS_COLS.QUANTITY] = e.quantity || 0;
    return row;
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, CONFIG.QA_EVENTS_COLS.QUANTITY + 1).setValues(rows);
  rows.forEach(r => _appendSheetCacheRow(CONFIG.SHEETS.QA_EVENTS, r));
  return { rowIndex: startRow, numRows: rows.length };
}

function _rollbackQaEvents(change) {
  if (!change || !change.numRows) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.QA_EVENTS);
  for (let i = 0; i < change.numRows; i++) {
    try { sheet.deleteRow(change.rowIndex); } catch (e) { }
    _deleteSheetCacheRow(CONFIG.SHEETS.QA_EVENTS, change.rowIndex);
  }
}

function _resolveMovementLogUom(data, itemMaps) {
  const direct = String((data && data.uom) || '').trim().toUpperCase();
  if (direct) return direct;

  const maps = itemMaps || _getItemMasterMaps();
  const itemIdNorm = String((data && data.itemId) || '').trim().toUpperCase();
  let code = String((data && data.itemCode) || '').trim();
  if (!code && itemIdNorm) code = maps.idToCode[itemIdNorm] || '';
  if (!code) return 'KG';

  const codeNorm = code.toUpperCase();
  return String(maps.codeToUom[codeNorm] || _getItemUomCode(code) || 'KG').trim().toUpperCase() || 'KG';
}

function _appendMovementLogsBatch(logs) {
  if (!logs || logs.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
  const itemMaps = _getItemMasterMaps();
  const rows = logs.map(data => {
    const row = [];
    row[CONFIG.MOVEMENT_COLS.MOVEMENT_ID] = 'MOV-' + new Date().getTime() + '-' + Math.random().toString(36).slice(2, 7).toUpperCase();
    row[CONFIG.MOVEMENT_COLS.TIMESTAMP] = new Date();
    row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] = data.type;
    row[CONFIG.MOVEMENT_COLS.ITEM_ID] = data.itemId;
    row[CONFIG.MOVEMENT_COLS.BATCH_ID] = data.batchId;
    // Do not auto-default version; missing stays UNSPECIFIED in read models.
    row[CONFIG.MOVEMENT_COLS.VERSION] = data.version || '';
    row[CONFIG.MOVEMENT_COLS.GIN_NO] = data.ginNo || '';
    row[CONFIG.MOVEMENT_COLS.PROD_ORDER_REF] = data.prodOrderRef || '';
    row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID] = data.fromBinId || '';
    row[CONFIG.MOVEMENT_COLS.TO_BIN_ID] = data.toBinId || '';
    row[CONFIG.MOVEMENT_COLS.QUANTITY] = data.quantity;
    row[CONFIG.MOVEMENT_COLS.UOM] = _resolveMovementLogUom(data, itemMaps);
    row[CONFIG.MOVEMENT_COLS.QUALITY_STATUS] = data.qualityStatus || '';
    row[CONFIG.MOVEMENT_COLS.USER_EMAIL] = Session.getActiveUser().getEmail();
    row[CONFIG.MOVEMENT_COLS.REMARKS] = data.remarks || '';
    return row;
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, CONFIG.MOVEMENT_COLS.REMARKS + 1).setValues(rows);
  rows.forEach(r => _appendSheetCacheRow(CONFIG.SHEETS.MOVEMENT, r));
}

function _getBinMetaById(binId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const binSheet = ss.getSheetByName(CONFIG.SHEETS.BIN);
  if (!binSheet) return { binCode: String(binId || '').trim() };
  const binData = _getSheetValuesCached(binSheet.getName());
  for (let i = 1; i < binData.length; i++) {
    if (String(binData[i][0] || '').trim() === String(binId || '').trim()) {
      return {
        binCode: String(binData[i][2] || binId).trim()
      };
    }
  }
  return { binCode: String(binId || '').trim() };
}

function _appendProductionTransfersBatch(rows) {
  if (!rows || rows.length === 0) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_TRANSFERS);
  const cols = CONFIG.PRODUCTION_TRANSFERS_COLS;
  const itemMaps = _getItemMasterMaps();
  const out = rows.map(r => {
    const row = [];
    const resolvedItemCode = String(r.itemCode || '').trim();
    const fallbackUom = _resolveMovementLogUom({ itemCode: resolvedItemCode, itemId: r.itemId, uom: r.uom }, itemMaps);
    row[cols.TRANSFER_ID] = r.transferId || ('PT-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000));
    row[cols.TRANSFER_DATE] = r.transferDate || new Date();
    row[cols.TRANSFER_TYPE] = r.transferType || '';
    row[cols.PRODUCTION_ORDER_NO] = r.productionOrderNo || '';
    row[cols.PRODUCTION_AREA] = r.productionArea || '';
    row[cols.ITEM_CODE] = resolvedItemCode;
    row[cols.ITEM_ID] = r.itemId || '';
    row[cols.BATCH_NUMBER] = r.batchNumber || r.batchId || '';
    if (typeof cols.BATCH_ID === 'number' && cols.BATCH_ID !== cols.BATCH_NUMBER) {
      row[cols.BATCH_ID] = r.batchId || r.batchNumber || '';
    }
    if (typeof cols.LOT_NO === 'number') {
      row[cols.LOT_NO] = r.lotNumber || r.lotNo || '';
    }
    row[cols.QUANTITY] = r.quantity || 0;
    if (typeof cols.UOM === 'number') {
      row[cols.UOM] = fallbackUom;
    }
    row[cols.FROM_SITE] = r.fromSite || '';
    row[cols.FROM_LOCATION] = r.fromLocation || '';
    row[cols.FROM_BIN_ID] = r.fromBinId || '';
    row[cols.FROM_BIN_CODE] = r.fromBinCode || '';
    row[cols.TO_SITE] = r.toSite || '';
    row[cols.TO_LOCATION] = r.toLocation || '';
    row[cols.TO_BIN_ID] = r.toBinId || '';
    row[cols.TO_BIN_CODE] = r.toBinCode || '';
    row[cols.RETURNED_BY_NAME] = r.returnedByName || '';
    row[cols.CREATED_BY] = r.createdBy || '';
    row[cols.CREATED_AT] = r.createdAt || new Date();
    row[cols.REMARKS] = r.remarks || '';
    row[cols.STATUS] = r.status || 'COMPLETED';
    return row;
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, out.length, cols.STATUS + 1).setValues(out);
  out.forEach(r => _appendSheetCacheRow(CONFIG.SHEETS.PRODUCTION_TRANSFERS, r));
  return { rowIndex: startRow, numRows: out.length };
}

function _rollbackProductionTransfers(change) {
  if (!change || !change.numRows) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_TRANSFERS);
  for (let i = 0; i < change.numRows; i++) {
    try { sheet.deleteRow(change.rowIndex); } catch (e) { }
    _deleteSheetCacheRow(CONFIG.SHEETS.PRODUCTION_TRANSFERS, change.rowIndex);
  }
}

function _getDeadStockSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.DEAD_STOCK);
  if (!sheet) {
    const target = String(CONFIG.SHEETS.DEAD_STOCK || '').trim().toUpperCase();
    const sheets = ss.getSheets();
    sheet = sheets.find(function (s) {
      return String(s.getName() || '').trim().toUpperCase() === target;
    }) || null;
  }
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.DEAD_STOCK);
  }
  // Bug 1 fix: also write header if sheet has rows but row 1 is NOT a proper header.
  // This handles the case where dead_stock_log was populated without a header row,
  // causing _getDeadStockSummary to mismap column indexes and return 0 records.
  var needsHeader = (sheet.getLastRow() === 0);
  if (!needsHeader && sheet.getLastRow() > 0) {
    var firstRowCheck = sheet.getRange(1, 1, 1, CONFIG.DEAD_STOCK_COLS.USER_EMAIL + 1).getValues()[0];
    var col0Val = String(firstRowCheck[0] || '').trim().toUpperCase();
    // If column A of row 1 is not 'DEAD_STOCK_ID', the header row is missing
    if (col0Val !== 'DEAD_STOCK_ID') {
      Logger.log('[DEAD_STOCK] Header row missing from ' + CONFIG.SHEETS.DEAD_STOCK +
        '. Row 1 col A = "' + col0Val + '". Inserting header row now.');
      sheet.insertRowBefore(1);
      needsHeader = true;
      // Invalidate any cached values for this sheet
      _ensureRequestCache();
      delete _REQ_CACHE.values[CONFIG.SHEETS.DEAD_STOCK];
      delete _REQ_CACHE.headers[CONFIG.SHEETS.DEAD_STOCK];
    }
  }
  if (needsHeader) {
    var cols = CONFIG.DEAD_STOCK_COLS;
    var hdr = [];
    hdr[cols.DEAD_STOCK_ID] = 'dead_stock_id';
    hdr[cols.TIMESTAMP] = 'timestamp';
    hdr[cols.PROD_ORDER_REF] = 'prod_order_ref';
    hdr[cols.ITEM_ID] = 'item_id';
    hdr[cols.ITEM_CODE] = 'item_code';
    hdr[cols.BATCH_ID] = 'batch_id';
    hdr[cols.QUANTITY] = 'quantity';
    hdr[cols.UOM] = 'uom';
    hdr[cols.REASON] = 'reason';
    hdr[cols.REMARKS] = 'remarks';
    hdr[cols.USER_EMAIL] = 'user_email';
    sheet.getRange(1, 1, 1, cols.USER_EMAIL + 1).setValues([hdr]);
  }
  return sheet;
}

function _appendDeadStockRows(rows) {
  if (!rows || rows.length === 0) return null;
  const sheet = _getDeadStockSheet();
  const sheetName = sheet.getName();
  const cols = CONFIG.DEAD_STOCK_COLS;
  const out = rows.map(function (r) {
    const row = [];
    row[cols.DEAD_STOCK_ID] = 'DS-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    row[cols.TIMESTAMP] = r.timestamp || new Date();
    row[cols.PROD_ORDER_REF] = r.prodOrderRef || '';
    row[cols.ITEM_ID] = r.itemId || '';
    row[cols.ITEM_CODE] = r.itemCode || '';
    row[cols.BATCH_ID] = r.batchId || '';
    row[cols.QUANTITY] = Number(r.quantity) || 0;
    row[cols.UOM] = String(r.uom || 'KG').trim().toUpperCase() || 'KG';
    row[cols.REASON] = r.reason || '';
    row[cols.REMARKS] = r.remarks || r.reason || '';
    row[cols.USER_EMAIL] = r.userEmail || Session.getActiveUser().getEmail();
    return row;
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, out.length, cols.USER_EMAIL + 1).setValues(out);
  out.forEach(function (r) { _appendSheetCacheRow(sheetName, r); });
  return { rowIndex: startRow, numRows: out.length, sheetName: sheetName };
}

function _rollbackDeadStock(change) {
  if (!change || !change.numRows) return;
  const sheet = _getDeadStockSheet();
  const sheetName = String(change.sheetName || sheet.getName() || '');
  for (let i = 0; i < change.numRows; i++) {
    try { sheet.deleteRow(change.rowIndex); } catch (e) { }
    _deleteSheetCacheRow(sheetName, change.rowIndex);
  }
}

function _findInventoryRowIndexById(inventoryId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const finder = sheet.createTextFinder(String(inventoryId)).matchEntireCell(true);
  const match = finder.findNext();
  if (!match) return -1;
  return match.getRow();
}

/**
 * Internal helper: Get Item by ID (cached)
 */
function _getItemByIdCached(itemId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'item_id_' + itemId;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);
  if (!sheet) return null;

  const data = _getSheetValuesCached(sheet.getName());
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(itemId)) {
      const item = { id: data[i][0], code: data[i][1], name: data[i][2] };
      cache.put(cacheKey, JSON.stringify(item), 300);
      return item;
    }
  }
  return null;
}

// ==== END OF CHANGES ====== 


function _persistInventoryStatesBatch(changes, sheet) {
  if (!changes || changes.length === 0) return;
  const targetSheet = sheet || _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.INVENTORY);
  const startCol = CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1;
  const endCol = CONFIG.INVENTORY_COLS.LAST_UPDATED + 1;
  const width = endCol - startCol + 1;
  const now = new Date();

  const byRow = {};
  changes.forEach(function (ch) {
    if (!ch || !isFinite(Number(ch.rowIndex))) return;
    if (typeof ch.newQty === 'undefined') return;
    byRow[Number(ch.rowIndex)] = Number(ch.newQty);
  });
  const rowIndexes = Object.keys(byRow).map(function (r) { return Number(r); }).filter(function (r) { return r > 1; });
  if (rowIndexes.length === 0) return;
  rowIndexes.sort(function (a, b) { return a - b; });

  const data = _getSheetValuesCached(targetSheet.getName());
  let groupStart = rowIndexes[0];
  let group = [rowIndexes[0]];

  function flushGroup() {
    if (group.length === 0) return;
    const values = group.map(function (rowIndex) {
      const rowData = (data[rowIndex - 1] || []).slice(startCol - 1, endCol);
      while (rowData.length < width) rowData.push('');
      rowData[0] = byRow[rowIndex];
      rowData[width - 1] = now;
      return rowData;
    });
    targetSheet.getRange(groupStart, startCol, group.length, width).setValues(values);
    group.forEach(function (rowIndex) {
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, rowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1, byRow[rowIndex]);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, rowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1, now);
    });
  }

  for (let i = 1; i < rowIndexes.length; i++) {
    const idx = rowIndexes[i];
    const prev = rowIndexes[i - 1];
    if (idx === prev + 1) {
      group.push(idx);
      continue;
    }
    flushGroup();
    groupStart = idx;
    group = [idx];
  }
  flushGroup();
  delete _REQ_CACHE.inventoryUsageSummaryByBin;
  delete _REQ_CACHE.inventoryUsageSummaryByBinSet;
}

function _persistInventoryState(rowIndex, newQty, sheet) {
  _persistInventoryStatesBatch([{ rowIndex: rowIndex, newQty: newQty }], sheet);
}

function _setInventoryLastTransferDateForRows(rowIndexes, transferDate, sheet) {
  if (!Array.isArray(rowIndexes) || rowIndexes.length === 0) return;
  if (typeof CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE !== 'number') return;
  const targetSheet = sheet || _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.INVENTORY);
  const col = CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE + 1;
  const when = transferDate || new Date();
  const seen = {};
  rowIndexes.forEach(function (r) {
    const row = Number(r);
    if (!isFinite(row) || row <= 1) return;
    if (seen[row]) return;
    seen[row] = true;
    targetSheet.getRange(row, col).setValue(when);
    _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, row, col, when);
  });
}

function _cleanupZeroQtyInventoryRows(changes, sheet) {
  if (!changes || changes.length === 0) return 0;
  const targetSheet = sheet || _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.INVENTORY);
  const rows = {};

  changes.forEach(function (ch) {
    if (!ch || ch.action === 'create') return;
    const rowIndex = Number(ch.rowIndex);
    if (!isFinite(rowIndex) || rowIndex <= 1) return;
    const newQty = Number(ch.newQty);
    if (!isFinite(newQty) || newQty > 0) return;
    rows[rowIndex] = true;
  });

  const rowIndexes = Object.keys(rows).map(function (r) { return Number(r); })
    .filter(function (r) { return isFinite(r) && r > 1; })
    .sort(function (a, b) { return b - a; });
  if (rowIndexes.length === 0) return 0;

  rowIndexes.forEach(function (rowIndex) {
    try {
      targetSheet.deleteRow(rowIndex);
      _deleteSheetCacheRow(CONFIG.SHEETS.INVENTORY, rowIndex);
    } catch (e) {
      Logger.log('Zero-qty cleanup skipped for row ' + rowIndex + ': ' + e.message);
    }
  });
  return rowIndexes.length;
}

function _createInventoryRow(params) {
  // CRITICAL: Validate bin exists before creating inventory
  _validateBinExists(params.binId);
  if (!params.version) throw new Error('Inventory version required');
  if (!params.ginNo) throw new Error('Bill No. required for inventory');
  if (!params.itemCode || !params.batchId) throw new Error('Item Code and Batch No required for inventory');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const id = 'INV-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  const rowIndex = sheet.getLastRow() + 1;

  const row = [];
  row[CONFIG.INVENTORY_COLS.INVENTORY_ID] = id;
  row[CONFIG.INVENTORY_COLS.ITEM_ID] = params.itemId || params.itemCode;
  row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] = params.itemCode || params.itemId;
  row[CONFIG.INVENTORY_COLS.BATCH_ID] = params.batchId;
  row[CONFIG.INVENTORY_COLS.GIN_NO] = params.ginNo;
  row[CONFIG.INVENTORY_COLS.VERSION] = params.version || '';
  row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = params.qualityStatus || 'PENDING';
  row[CONFIG.INVENTORY_COLS.QUALITY_DATE] = params.qualityDate || '';
  row[CONFIG.INVENTORY_COLS.BIN_ID] = params.binId;
  row[CONFIG.INVENTORY_COLS.SITE] = params.site || 'Main';
  row[CONFIG.INVENTORY_COLS.LOCATION] = params.location || 'Main';
  row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = params.quantity;
  if (CONFIG.INVENTORY_COLS.UOM !== undefined) {
    row[CONFIG.INVENTORY_COLS.UOM] = String(params.uom || _getItemUomCode(params.itemCode || params.itemId || '') || 'KG').trim().toUpperCase();
  }
  if (CONFIG.INVENTORY_COLS.LOT_NO !== undefined) {
    row[CONFIG.INVENTORY_COLS.LOT_NO] = params.lotNo || '';
  }
  if (CONFIG.INVENTORY_COLS.EXPIRY_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = _normalizeInwardExpiryDate(params.expiryDate);
  }
  row[CONFIG.INVENTORY_COLS.LAST_UPDATED] = new Date();
  if (CONFIG.INVENTORY_COLS.INWARD_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.INWARD_DATE] = params.inwardDate || new Date();
  }
  if (CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = params.lastTransferDate || '';
  }

  sheet.appendRow(row);
  _appendSheetCacheRow(CONFIG.SHEETS.INVENTORY, row);
  if (params.returnRowIndex === true) {
    return { id: id, rowIndex: rowIndex };
  }
  return id;
}

function _appendInventoryRowsBatch(paramsList, sheet) {
  if (!paramsList || paramsList.length === 0) return [];
  const targetSheet = sheet || _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.INVENTORY);
  const width = CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE + 1;
  const startRow = targetSheet.getLastRow() + 1;
  const now = new Date();
  const rows = [];
  const out = [];

  paramsList.forEach(function (params, idx) {
    _validateBinExists(params.binId);
    if (!params.version) throw new Error('Inventory version required');
    if (!params.ginNo) throw new Error('Bill No. required for inventory');
    if (!params.itemCode || !params.batchId) throw new Error('Item Code and Batch No required for inventory');

    const row = [];
    const id = 'INV-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000) + '-' + idx;
    row[CONFIG.INVENTORY_COLS.INVENTORY_ID] = id;
    row[CONFIG.INVENTORY_COLS.ITEM_ID] = params.itemId || params.itemCode;
    row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] = params.itemCode || params.itemId;
    row[CONFIG.INVENTORY_COLS.BATCH_ID] = params.batchId;
    row[CONFIG.INVENTORY_COLS.GIN_NO] = params.ginNo;
    row[CONFIG.INVENTORY_COLS.VERSION] = params.version || '';
    row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = params.qualityStatus || 'PENDING';
    row[CONFIG.INVENTORY_COLS.QUALITY_DATE] = params.qualityDate || '';
    row[CONFIG.INVENTORY_COLS.BIN_ID] = params.binId;
    row[CONFIG.INVENTORY_COLS.SITE] = params.site || 'Main';
    row[CONFIG.INVENTORY_COLS.LOCATION] = params.location || 'Main';
    row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = params.quantity;
    row[CONFIG.INVENTORY_COLS.UOM] = String(params.uom || _getItemUomCode(params.itemCode || params.itemId || '') || 'KG').trim().toUpperCase();
    row[CONFIG.INVENTORY_COLS.LOT_NO] = params.lotNo || '';
    row[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = _normalizeInwardExpiryDate(params.expiryDate);
    row[CONFIG.INVENTORY_COLS.LAST_UPDATED] = params.overrideDate ? new Date(params.overrideDate) : now;
    row[CONFIG.INVENTORY_COLS.INWARD_DATE] = params.inwardDate ? new Date(params.inwardDate) : now;
    row[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = params.lastTransferDate || '';
    while (row.length < width) row.push('');
    rows.push(row.slice(0, width));
    out.push({ id: id, rowIndex: startRow + idx, row: row.slice(0, width) });
  });

  targetSheet.getRange(startRow, 1, rows.length, width).setValues(rows);
  rows.forEach(function (row) {
    _appendSheetCacheRow(CONFIG.SHEETS.INVENTORY, row);
  });
  delete _REQ_CACHE.inventoryUsageSummaryByBin;
  delete _REQ_CACHE.inventoryUsageSummaryByBinSet;
  return out;
}

function _appendMovementLog(data) {
  _appendMovementLogsBatch([data]);
}

// =====================================================
// 5. PRODUCTION LEDGER LOGIC
// =====================================================

/**
 * CRITICAL SECURITY: Server-authoritative item lookup
 * NEVER trust client-provided itemId - always resolve from itemCode
 * @param {string} itemCode - Item code from client
 * @returns {string} Validated item ID from Item_Master
 * @throws {Error} If item code not found
 */
function _getValidatedItemId(itemCode) {
  if (!itemCode) throw new Error('Item Code required');

  const maps = _getItemMasterMaps();
  const itemId = maps.codeToId[String(itemCode).trim().toUpperCase()];
  if (itemId) return String(itemId);

  throw new Error(`Invalid Item Code: ${itemCode}. Not found in Item_Master.`);
}

/**
 * CRITICAL DATA INTEGRITY: Validate bin exists before creating inventory
 * @param {string} binId - Bin ID from client
 * @returns {boolean} True if bin exists
 * @throws {Error} If bin not found in Bin_Master
 */
function _validateBinExists(binId) {
  if (!binId) throw new Error('Bin ID required');

  const ctx = _getBinMasterMetaContext();
  if ((ctx.metaMap || {})[String(binId).trim()]) return true;

  throw new Error(`Invalid Bin ID: ${binId}. Bin does not exist in Bin_Master.`);
}

// ===== UOM & CAPACITY HELPERS (item-base UOM internal units) =====

function _getSheetHeaderMap(sheet) {
  if (!sheet) return {};
  return _getSheetHeaderMapCached(sheet.getName());
}

function _parseUomActiveFlag(raw) {
  const txt = String(raw || '').trim().toUpperCase();
  if (!txt) return true;
  if (txt === 'FALSE' || txt === '0' || txt === 'NO' || txt === 'INACTIVE') return false;
  return true;
}

function _getUomMasterMeta() {
  _ensureRequestCache();
  if (_REQ_CACHE.uomMasterMeta) return _REQ_CACHE.uomMasterMeta;

  const meta = { byCode: {}, byBase: {}, allCodes: [] };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.UOM_MASTER || 'UOM_Master');
  if (!sheet) {
    _REQ_CACHE.uomMasterMeta = meta;
    return meta;
  }

  const data = _getSheetValuesCached(sheet.getName());
  const headerMap = _getSheetHeaderMap(sheet);
  const codeIdx = (typeof headerMap['UOM_CODE'] === 'number') ? headerMap['UOM_CODE'] : 0;
  const baseIdx = (typeof headerMap['BASE_UOM'] === 'number')
    ? headerMap['BASE_UOM']
    : ((typeof headerMap['BASE_UOM_CODE'] === 'number') ? headerMap['BASE_UOM_CODE'] : 1);
  const factorIdx = (typeof headerMap['FACTOR_TO_BASE'] === 'number')
    ? headerMap['FACTOR_TO_BASE']
    : ((typeof headerMap['CONVERSION_TO_BASE'] === 'number')
      ? headerMap['CONVERSION_TO_BASE']
      : ((typeof headerMap['FACTOR'] === 'number') ? headerMap['FACTOR'] : 2));
  const activeIdx = (typeof headerMap['IS_ACTIVE'] === 'number')
    ? headerMap['IS_ACTIVE']
    : ((typeof headerMap['STATUS'] === 'number') ? headerMap['STATUS'] : null);
  const orderIdx = (typeof headerMap['DISPLAY_ORDER'] === 'number')
    ? headerMap['DISPLAY_ORDER']
    : ((typeof headerMap['SORT_ORDER'] === 'number') ? headerMap['SORT_ORDER'] : null);
  const ordered = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = String(row[codeIdx] || '').trim().toUpperCase();
    if (!code) continue;
    const base = String(row[baseIdx] || code).trim().toUpperCase() || code;
    const factorRaw = Number(row[factorIdx]);
    let factorToBase = (isFinite(factorRaw) && factorRaw > 0) ? factorRaw : 1;
    if (code === base) factorToBase = 1;
    const isActive = activeIdx === null ? true : _parseUomActiveFlag(row[activeIdx]);
    if (!isActive) continue;

    meta.byCode[code] = { code: code, base: base, factorToBase: factorToBase };
    if (!meta.byBase[base]) meta.byBase[base] = [];
    meta.byBase[base].push(code);
    const ordRaw = (typeof orderIdx === 'number') ? Number(row[orderIdx]) : NaN;
    const ord = (isFinite(ordRaw) && ordRaw > 0) ? ordRaw : 999999;
    ordered.push({ code: code, order: ord });
  }

  Object.keys(meta.byBase).forEach(function (base) {
    meta.byBase[base] = Array.from(new Set(meta.byBase[base]));
    if (meta.byBase[base].indexOf(base) === -1) meta.byBase[base].unshift(base);
  });

  ordered.sort(function (a, b) {
    if (a.order !== b.order) return a.order - b.order;
    return String(a.code).localeCompare(String(b.code));
  });
  meta.allCodes = ordered.map(function (x) { return x.code; })
    .filter(function (c, idx, arr) { return arr.indexOf(c) === idx; });
  if (meta.allCodes.length === 0) meta.allCodes = ['KG'];

  _REQ_CACHE.uomMasterMeta = meta;
  return meta;
}

function _getUomFactorToBase(inputUom, baseUom) {
  const inU = String(inputUom || '').trim().toUpperCase();
  const base = String(baseUom || '').trim().toUpperCase();
  if (!inU || !base) return NaN;
  if (inU === base) return 1;

  const meta = _getUomMasterMeta();
  const entry = meta.byCode[inU];
  if (!entry) return NaN;
  if (String(entry.base || '').toUpperCase() !== base) return NaN;
  const factor = Number(entry.factorToBase);
  return (isFinite(factor) && factor > 0) ? factor : NaN;
}

function _getItemUomOptions(itemCode) {
  const info = _getItemUomInfo(itemCode);
  const base = String(info.baseUom || 'KG').trim().toUpperCase() || 'KG';
  const list = [base];

  const alt = String(info.altUom || '').trim().toUpperCase();
  if (alt && list.indexOf(alt) === -1) list.push(alt);

  const meta = _getUomMasterMeta();
  const fromMaster = meta.byBase[base] || [];
  fromMaster.forEach(function (u) {
    if (u && list.indexOf(u) === -1) list.push(u);
  });

  return list;
}

function _getAllActiveUomCodes() {
  const meta = _getUomMasterMeta();
  const codes = Array.isArray(meta.allCodes) ? meta.allCodes.slice() : [];
  if (codes.length > 0) return codes;
  return ['KG'];
}

function _getItemUomInfo(itemCode) {
  const info = { baseUom: '', altUom: '', factorToKg: 1 };
  const code = String(itemCode || '').trim();
  if (!code) {
    info.baseUom = 'KG';
    return info;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);
  if (!sheet) {
    info.baseUom = 'KG';
    return info;
  }

  const data = _getSheetValuesCached(sheet.getName());
  if (data.length < 2) {
    info.baseUom = 'KG';
    return info;
  }

  const cols = _getItemSheetColumnMap(sheet);
  const itemCodeIdx = cols.ITEM_CODE;
  const uomIdx = cols.UOM_CODE;
  const altIdx = cols.ALT_UOM;
  const convIdx = cols.CONVERSION_FACTOR_TO_KG;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][itemCodeIdx] || '').trim().toUpperCase() === code.toUpperCase()) {
      info.baseUom = String(data[i][uomIdx] || '').trim().toUpperCase();
      if (typeof altIdx === 'number') info.altUom = String(data[i][altIdx] || '').trim();
      if (typeof convIdx === 'number') {
        const factor = Number(data[i][convIdx]);
        if (isFinite(factor) && factor > 0) info.factorToKg = factor;
      }
      break;
    }
  }

  if (!info.baseUom) {
    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    if (invSheet) {
      const invData = _getSheetValuesCached(invSheet.getName());
      const target = code.toUpperCase();
      for (let i = 1; i < invData.length; i++) {
        const rowCode = String(invData[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
        if (rowCode !== target) continue;
        const rowUom = String((CONFIG.INVENTORY_COLS.UOM !== undefined ? invData[i][CONFIG.INVENTORY_COLS.UOM] : '') || '').trim().toUpperCase();
        if (rowUom) {
          info.baseUom = rowUom;
          break;
        }
      }
    }
  }

  if (!info.baseUom) info.baseUom = 'KG';
  return info;
}

function _convertToKg(itemCode, qty, inputUom) {
  const num = Number(qty);
  if (!isFinite(num)) return 0;

  const info = _getItemUomInfo(itemCode);
  const baseUom = String(info.baseUom || 'KG').trim().toUpperCase() || 'KG';
  const inUom = String(inputUom || baseUom).trim().toUpperCase() || baseUom;
  if (inUom === baseUom) return num;

  const alt = String(info.altUom || '').trim().toUpperCase();
  if (inUom && alt && inUom === alt) {
    return num * (Number(info.factorToKg) || 1);
  }

  const masterFactor = _getUomFactorToBase(inUom, baseUom);
  if (isFinite(masterFactor) && masterFactor > 0) {
    return num * masterFactor;
  }

  return num;
}

function _normalizeCapacityUom(value) {
  const raw = String(value || 'KG').trim().toUpperCase();
  if (!raw) return 'KG';
  if (['KG', 'KGS', 'KILOGRAM', 'KILOGRAMS'].indexOf(raw) >= 0) return 'KG';
  if (['NOS', 'NO', 'PCS', 'PC', 'EA', 'EACH', 'UNIT', 'UNITS', 'PIECE', 'PIECES'].indexOf(raw) >= 0) return 'NOS';
  if (['L', 'LTR', 'LITRE', 'LITRES', 'LITER', 'LITERS'].indexOf(raw) >= 0) return 'L';
  return raw;
}

function _getSupportedCapacityUoms() {
  return ['KG', 'NOS', 'L'];
}

function _pickPrimaryCapacityUom(capacities, usageByUom) {
  const caps = capacities || {};
  const usageMap = usageByUom || {};
  const ordered = _getSupportedCapacityUoms().filter(function (uom) {
    return Number(caps[uom] || 0) > 0;
  });
  if (ordered.length === 0) return 'KG';

  let best = ordered[0];
  let bestPct = -1;
  ordered.forEach(function (uom) {
    const cap = Number(caps[uom] || 0);
    const used = Number((((usageMap || {})[uom] || {}).qty) || 0);
    const pct = cap > 0 ? (used / cap) : 0;
    if (pct > bestPct) {
      best = uom;
      bestPct = pct;
    }
  });
  return best;
}

function _extractBinCapacitiesFromRow(row, colMap) {
  const cols = colMap || {};
  const raw = row || [];
  const capacities = {};
  const legacyCapIdx = typeof cols['MAX_CAPACITY_KG'] === 'number'
    ? cols['MAX_CAPACITY_KG']
    : (typeof cols['MAX_CAPACITY'] === 'number' ? cols['MAX_CAPACITY'] : 5);
  const legacyCapUomIdx = typeof cols['CAPACITY_UOM'] === 'number'
    ? cols['CAPACITY_UOM']
    : (typeof cols['MAX_CAPACITY_UOM'] === 'number' ? cols['MAX_CAPACITY_UOM'] : 6);
  const kgIdx = typeof cols['MAX_CAPACITY_KG'] === 'number' ? cols['MAX_CAPACITY_KG'] : null;
  const nosIdx = typeof cols['MAX_CAPACITY_NOS'] === 'number' ? cols['MAX_CAPACITY_NOS'] : null;
  const lIdx = typeof cols['MAX_CAPACITY_L'] === 'number' ? cols['MAX_CAPACITY_L'] : null;
  const hasDedicatedCols = kgIdx !== null || nosIdx !== null || lIdx !== null;

  if (kgIdx !== null) {
    const kg = Number(raw[kgIdx]) || 0;
    if (kg > 0) capacities.KG = kg;
  }
  if (nosIdx !== null) {
    const nos = Number(raw[nosIdx]) || 0;
    if (nos > 0) capacities.NOS = nos;
  }
  if (lIdx !== null) {
    const liters = Number(raw[lIdx]) || 0;
    if (liters > 0) capacities.L = liters;
  }

  const legacyCapacity = Number(raw[legacyCapIdx]) || 0;
  const legacyUom = _normalizeCapacityUom(raw[legacyCapUomIdx] || 'KG');
  if (legacyCapacity > 0 && (!hasDedicatedCols || !capacities[legacyUom])) {
    capacities[legacyUom] = legacyCapacity;
  }

  return capacities;
}

function _getDeclaredBinCapacityUom(row, colMap, capacities) {
  const cols = colMap || {};
  const raw = row || [];
  const explicitUomIdx = typeof cols['CAPACITY_UOM'] === 'number'
    ? cols['CAPACITY_UOM']
    : (typeof cols['MAX_CAPACITY_UOM'] === 'number' ? cols['MAX_CAPACITY_UOM'] : null);
  const normalized = _normalizeCapacityUom(explicitUomIdx === null ? '' : raw[explicitUomIdx]);
  if (!normalized) return '';
  const caps = capacities || {};
  if (Number(caps[normalized] || 0) > 0) return normalized;
  return normalized;
}

function _getBinCapacityMetrics(meta, usageEntry) {
  const capacities = (meta && meta.capacities) ? meta.capacities : {};
  const usageByUom = (usageEntry && usageEntry.byUom) ? usageEntry.byUom : {};
  const capacityLines = _getSupportedCapacityUoms()
    .filter(function (uom) { return Number(capacities[uom] || 0) > 0; })
    .map(function (uom) {
      const maxCapacity = Number(capacities[uom] || 0);
      const currentUsage = Number((((usageByUom || {})[uom] || {}).qty) || 0);
      const availableCapacity = maxCapacity > 0 ? Math.max(0, maxCapacity - currentUsage) : 0;
      const utilizationPct = maxCapacity > 0 ? Math.max(0, (currentUsage / maxCapacity) * 100) : 0;
      return {
        uom: uom,
        maxCapacity: maxCapacity,
        currentUsage: currentUsage,
        availableCapacity: availableCapacity,
        utilizationPct: utilizationPct
      };
    });

  const primaryUom = _pickPrimaryCapacityUom(capacities, usageByUom);
  const primaryLine = capacityLines.filter(function (line) {
    return line.uom === primaryUom;
  })[0] || {
    uom: primaryUom,
    maxCapacity: 0,
    currentUsage: 0,
    availableCapacity: 0,
    utilizationPct: 0
  };

  return {
    capacityLines: capacityLines,
    capacityUom: primaryLine.uom,
    maxCapacity: primaryLine.maxCapacity,
    currentUsage: primaryLine.currentUsage,
    availableCapacity: primaryLine.availableCapacity,
    utilizationPct: primaryLine.utilizationPct,
    usageByUom: usageByUom,
    hasInventory: !!(usageEntry && usageEntry.hasInventory),
    capacityWarning: ''
  };
}

function _resolveBinCapacitySlot(meta, itemCode, inputUom) {
  const capacities = (meta && meta.capacities) ? meta.capacities : {};
  const configured = _getSupportedCapacityUoms().filter(function (uom) {
    return Number(capacities[uom] || 0) > 0;
  });
  if (configured.length === 0) {
    return { uom: '', capacity: 0, configuredUoms: [] };
  }

  const info = _getItemUomInfo(itemCode);
  const candidates = [];
  function pushCandidate(value) {
    const normalized = _normalizeCapacityUom(value || '');
    if (!normalized) return;
    if (candidates.indexOf(normalized) === -1) candidates.push(normalized);
  }

  pushCandidate(inputUom);
  pushCandidate(info.baseUom);
  pushCandidate(info.altUom);

  for (let i = 0; i < candidates.length; i++) {
    const uom = candidates[i];
    if (configured.indexOf(uom) >= 0) {
      return { uom: uom, capacity: Number(capacities[uom] || 0), configuredUoms: configured };
    }
  }

  const testFromUom = String(inputUom || info.baseUom || '').trim().toUpperCase() || String(info.baseUom || 'KG').trim().toUpperCase();
  for (let j = 0; j < configured.length; j++) {
    const targetUom = configured[j];
    const probe = _convertItemQtyToTargetUom(itemCode, 1, testFromUom, targetUom);
    if (isFinite(probe)) {
      return { uom: targetUom, capacity: Number(capacities[targetUom] || 0), configuredUoms: configured };
    }
  }

  return { uom: '', capacity: 0, configuredUoms: configured, missing: true };
}

function _convertItemQtyToTargetUom(itemCode, qty, inputUom, targetUom) {
  const num = Number(qty);
  if (!isFinite(num)) return NaN;

  const info = _getItemUomInfo(itemCode);
  const baseUom = String(info.baseUom || 'KG').trim().toUpperCase() || 'KG';
  const fromUom = String(inputUom || baseUom).trim().toUpperCase() || baseUom;
  const toUom = _normalizeCapacityUom(targetUom || baseUom);
  const altUom = String(info.altUom || '').trim().toUpperCase();
  const baseFamily = _normalizeCapacityUom(baseUom);
  const fromFamily = _normalizeCapacityUom(fromUom);
  const altFamily = _normalizeCapacityUom(altUom);

  if (fromFamily === toUom) return num;

  let qtyInBase = num;
  if (fromFamily !== baseFamily) {
    if (fromUom && altUom && fromFamily === altFamily) {
      const altFactor = Number(info.factorToKg) || 1;
      qtyInBase = num * altFactor;
    } else {
      const fromFactor = _getUomFactorToBase(fromUom, baseUom);
      if (!(isFinite(fromFactor) && fromFactor > 0)) return NaN;
      qtyInBase = num * fromFactor;
    }
  }

  if (toUom === baseFamily) return qtyInBase;

  if (toUom && altUom && toUom === altFamily) {
    const altFactor = Number(info.factorToKg) || 1;
    if (!(isFinite(altFactor) && altFactor > 0)) return NaN;
    return qtyInBase / altFactor;
  }

  const toFactor = _getUomFactorToBase(toUom, baseUom);
  if (!(isFinite(toFactor) && toFactor > 0)) return NaN;
  return qtyInBase / toFactor;
}

function _getInventoryQtyByBinMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const binId = String(data[i][CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
    if (!binId) continue;
    const qty = Number(data[i][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    map[binId] = (map[binId] || 0) + qty;
  }
  return map;
}

function _buildBinMasterMetaMap(binSheet, binData, colMap) {
  const idIdx = typeof colMap['BIN_ID'] === 'number' ? colMap['BIN_ID'] : 0;
  const rackIdx = typeof colMap['RACK_ID'] === 'number' ? colMap['RACK_ID'] : 1;
  const statusIdx = typeof colMap['BIN_STATUS'] === 'number' ? colMap['BIN_STATUS'] : null;
  const siteIdx = typeof colMap['SITE'] === 'number' ? colMap['SITE'] : 3;
  const locationIdx = typeof colMap['LOCATION'] === 'number' ? colMap['LOCATION'] : 4;
  const codeIdx = typeof colMap['BIN_CODE'] === 'number' ? colMap['BIN_CODE'] : 2;
  const out = {};

  for (let i = 1; i < binData.length; i++) {
    const binId = String(binData[i][idIdx] || '').trim();
    if (!binId) continue;
    const capacities = _extractBinCapacitiesFromRow(binData[i], colMap);
    const declaredUom = _getDeclaredBinCapacityUom(binData[i], colMap, capacities);
    const primaryUom = declaredUom || _pickPrimaryCapacityUom(capacities, {});
    out[binId] = {
      rowIndex: i + 1,
      rackId: String(binData[i][rackIdx] || '').trim(),
      binCode: String(binData[i][codeIdx] || binId).trim(),
      site: String(binData[i][siteIdx] || '').trim(),
      location: String(binData[i][locationIdx] || '').trim(),
      capacity: Number(capacities[primaryUom] || 0),
      capacityUom: primaryUom,
      declaredCapacityUom: declaredUom || primaryUom,
      capacities: capacities,
      statusIdx: statusIdx
    };
  }

  return out;
}

function _getInventoryUsageSummaryByBin(binMetaMap) {
  _ensureRequestCache();
  if (_REQ_CACHE.inventoryUsageSummaryByBin) return _REQ_CACHE.inventoryUsageSummaryByBin;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const byBin = {};
  const kgByBin = {};
  const kgUnknownByBin = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
    if (!binId) continue;

    const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (!isFinite(qty) || qty <= 0) continue;

    const itemCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
    const rowUom = String(
      (typeof CONFIG.INVENTORY_COLS.UOM === 'number' ? row[CONFIG.INVENTORY_COLS.UOM] : '') ||
      _getItemUomCode(itemCode)
    ).trim().toUpperCase() || _getItemUomCode(itemCode);

    const meta = (binMetaMap && binMetaMap[binId]) ? binMetaMap[binId] : null;
    if (meta) {
      if (!byBin[binId]) {
        byBin[binId] = {
          byUom: {},
          hasInventory: false,
          rowCount: 0
        };
      }
      byBin[binId].hasInventory = true;
      byBin[binId].rowCount += 1;
      Object.keys(meta.capacities || {}).forEach(function (targetUom) {
        if (!byBin[binId].byUom[targetUom]) {
          byBin[binId].byUom[targetUom] = { qty: 0, conversionMissCount: 0 };
        }
        const converted = _convertItemQtyToTargetUom(itemCode, qty, rowUom, targetUom);
        if (isFinite(converted)) byBin[binId].byUom[targetUom].qty += converted;
        else byBin[binId].byUom[targetUom].conversionMissCount += 1;
      });
    }

    const kgQty = _convertItemQtyToTargetUom(itemCode, qty, rowUom, 'KG');
    if (isFinite(kgQty)) {
      kgByBin[binId] = (kgByBin[binId] || 0) + kgQty;
    } else {
      kgUnknownByBin[binId] = true;
    }
  }

  _REQ_CACHE.inventoryUsageSummaryByBin = {
    byBin: byBin,
    kgByBin: kgByBin,
    kgUnknownByBin: kgUnknownByBin
  };
  return _REQ_CACHE.inventoryUsageSummaryByBin;
}

function _getInventoryUsageSummaryForBinSet(binMetaMap, binSet, cacheKey) {
  _ensureRequestCache();
  const key = cacheKey || ('BIN_USAGE_SET_' + Object.keys(binSet || {}).sort().join('|'));
  const cache = _REQ_CACHE.inventoryUsageSummaryByBinSet || (_REQ_CACHE.inventoryUsageSummaryByBinSet = {});
  if (cache[key]) return cache[key];

  const targets = binSet || {};
  const hasTargets = Object.keys(targets).length > 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const byBin = {};
  const kgByBin = {};
  const kgUnknownByBin = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
    if (!binId || (hasTargets && !targets[binId])) continue;

    const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (!isFinite(qty) || qty <= 0) continue;

    const itemCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
    const rowUom = String(
      (typeof CONFIG.INVENTORY_COLS.UOM === 'number' ? row[CONFIG.INVENTORY_COLS.UOM] : '') ||
      _getItemUomCode(itemCode)
    ).trim().toUpperCase() || _getItemUomCode(itemCode);

    const meta = (binMetaMap && binMetaMap[binId]) ? binMetaMap[binId] : null;
    if (meta) {
      if (!byBin[binId]) byBin[binId] = { byUom: {}, hasInventory: false, rowCount: 0 };
      byBin[binId].hasInventory = true;
      byBin[binId].rowCount += 1;
      Object.keys(meta.capacities || {}).forEach(function (targetUom) {
        if (!byBin[binId].byUom[targetUom]) {
          byBin[binId].byUom[targetUom] = { qty: 0, conversionMissCount: 0 };
        }
        const converted = _convertItemQtyToTargetUom(itemCode, qty, rowUom, targetUom);
        if (isFinite(converted)) byBin[binId].byUom[targetUom].qty += converted;
        else byBin[binId].byUom[targetUom].conversionMissCount += 1;
      });
    }

    const kgQty = _convertItemQtyToTargetUom(itemCode, qty, rowUom, 'KG');
    if (isFinite(kgQty)) kgByBin[binId] = (kgByBin[binId] || 0) + kgQty;
    else kgUnknownByBin[binId] = true;
  }

  cache[key] = { byBin: byBin, kgByBin: kgByBin, kgUnknownByBin: kgUnknownByBin };
  return cache[key];
}

function _getBinMasterMeta(binId, binSheet, binData, colMap) {
  const metaMap = _buildBinMasterMetaMap(binSheet, binData, colMap);
  return metaMap[String(binId || '').trim()] || null;
}

function _getBinMasterMetaContext() {
  _ensureRequestCache();
  if (_REQ_CACHE.binMasterMetaContext) return _REQ_CACHE.binMasterMetaContext;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
  const binData = _getSheetValuesCached(binSheet.getName());
  const colMap = _getSheetHeaderMap(binSheet);
  const metaMap = _buildBinMasterMetaMap(binSheet, binData, colMap);

  _REQ_CACHE.binMasterMetaContext = {
    sheet: binSheet,
    data: binData,
    colMap: colMap,
    metaMap: metaMap
  };
  return _REQ_CACHE.binMasterMetaContext;
}

function _isZeroBufferBin(binId, meta) {
  const raw = String(binId || '').trim().toUpperCase();
  const code = String((meta && meta.binCode) || '').trim().toUpperCase();
  const rack = String((meta && meta.rackId) || '').trim().toUpperCase();
  return raw.indexOf('ZERO') === 0 || code.indexOf('ZERO') === 0 || rack === 'ZERO';
}

function _assertBinCapacity(binId, addQty, binDeltaMap, itemCode, inputUom) {
  const ctx = _getBinMasterMetaContext();
  const metaMap = ctx.metaMap || {};
  const meta = metaMap[String(binId || '').trim()] || null;
  if (!meta) return;

  const binKey = String(binId || '').trim();
  if (binDeltaMap && binKey) binDeltaMap[binKey] = true;
  if (_isZeroBufferBin(binKey, meta)) {
    return { convertedQty: 0, capacityUom: meta.capacityUom || '', capacity: Number(meta.capacity || 0) };
  }
  const rawAddQty = Number(addQty) || 0;
  const capDeltaMap = binDeltaMap
    ? (binDeltaMap.__capacityByBin || (binDeltaMap.__capacityByBin = {}))
    : null;
  const slot = _resolveBinCapacitySlot(meta, itemCode, inputUom);

  if (!slot.uom || !slot.capacity) {
    if (slot.missing) {
      throw new Error(`Bin ${binId} has no matching capacity configured for ${String(itemCode || 'this item').trim()}. Add ${slot.configuredUoms.join('/') || 'a'} capacity for its UOM in Bin_Master.`);
    }
    _assertRackSwlForBin(binId, addQty, binDeltaMap, itemCode, inputUom);
    return { convertedQty: 0, capacityUom: '', capacity: 0 };
  }

  const deltaKey = binKey + '||' + slot.uom;
  const priorDelta = capDeltaMap ? Number(capDeltaMap[deltaKey] || 0) : 0;
  let addQtyInCapacityUom = 0;
  if (rawAddQty > 0) {
    const sourceItemCode = String(itemCode || '').trim();
    const sourceUom = String(inputUom || _getItemUomCode(sourceItemCode)).trim().toUpperCase();
    addQtyInCapacityUom = _convertItemQtyToTargetUom(sourceItemCode, rawAddQty, sourceUom, slot.uom);
    if (!isFinite(addQtyInCapacityUom)) {
      throw new Error(`Bin ${binId} capacity for ${slot.uom} cannot be compared for ${sourceItemCode || 'selected item'} from ${sourceUom || 'its current UOM'}.`);
    }
  }

  const usageSummary = _getInventoryUsageSummaryForBinSet(metaMap, (function () {
    const set = {};
    set[binKey] = true;
    return set;
  })(), 'CAPACITY_' + binKey);
  const usage = usageSummary.byBin[binKey] || { byUom: {}, hasInventory: false };
  const currentQty = Number(((((usage || {}).byUom || {})[slot.uom] || {}).qty) || 0);
  const projected = currentQty + priorDelta + addQtyInCapacityUom;
  if (projected > slot.capacity) {
    throw new Error(`Bin capacity exceeded for ${binId}. Max ${slot.capacity} ${slot.uom}, projected ${projected.toFixed(2)} ${slot.uom}.`);
  }

  if (capDeltaMap && addQtyInCapacityUom > 0) {
    capDeltaMap[deltaKey] = priorDelta + addQtyInCapacityUom;
  }

  _assertRackSwlForBin(binId, addQty, binDeltaMap, itemCode, inputUom);
  return { convertedQty: addQtyInCapacityUom, capacityUom: slot.uom, capacity: slot.capacity };
}

function _deriveBinStatus(metrics) {
  const info = metrics || {};
  const lines = Array.isArray(info.capacityLines) ? info.capacityLines : [];
  const hasInventory = info.hasInventory === true;
  for (let i = 0; i < lines.length; i++) {
    if (Number(lines[i].currentUsage || 0) > Number(lines[i].maxCapacity || 0)) return 'BLOCKED';
  }
  if (!hasInventory) return 'FREE';
  for (let j = 0; j < lines.length; j++) {
    if (Number(lines[j].maxCapacity || 0) > 0 && Number(lines[j].currentUsage || 0) >= Number(lines[j].maxCapacity || 0)) {
      return 'FULL';
    }
  }
  return 'PARTIAL';
}

function _updateBinAndRackStatuses(binIds) {
  if (!binIds || binIds.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const binCtx = _getBinMasterMetaContext();
  const binSheet = binCtx.sheet;
  const binData = binCtx.data || [];
  const binMetaMap = binCtx.metaMap || {};
  const targetSet = {};
  (binIds || []).forEach(function (b) {
    const id = String(b || '').trim();
    if (!id) return;
    if (_isZeroBufferBin(id, binMetaMap[id] || null)) return;
    targetSet[id] = true;
  });
  if (Object.keys(targetSet).length === 0) return;

  const racksTouched = {};
  const usageBinSet = {};
  Object.keys(targetSet).forEach(function (binId) { usageBinSet[binId] = true; });
  const hasAnyStatusCol = Object.keys(binMetaMap).some(function (binId) {
    return typeof (binMetaMap[binId] || {}).statusIdx === 'number';
  });
  if (!hasAnyStatusCol) {
    Logger.log('Bin status column not found; skipping bin status updates');
  }

  for (let i = 1; i < binData.length; i++) {
    const binId = String(binData[i][0] || '').trim();
    if (!binId || !targetSet[binId]) continue;
    const meta = binMetaMap[binId] || {};
    const rackId = String(meta.rackId || '').trim();
    if (rackId) racksTouched[rackId] = true;
  }

  const rackIds = Object.keys(racksTouched);
  if (rackIds.length > 0) {
    for (let i = 1; i < binData.length; i++) {
      const binId = String(binData[i][0] || '').trim();
      const meta = binMetaMap[binId] || {};
      if (rackIds.indexOf(String(meta.rackId || '').trim()) !== -1) usageBinSet[binId] = true;
    }
  }

  const usageSummary = _getInventoryUsageSummaryForBinSet(binMetaMap, usageBinSet, 'STATUS_' + Object.keys(usageBinSet).sort().join('|'));

  for (let i = 1; i < binData.length; i++) {
    const binId = String(binData[i][0] || '').trim();
    if (!binId || !targetSet[binId]) continue;
    const meta = binMetaMap[binId] || {};
    const usage = usageSummary.byBin[binId] || { byUom: {}, hasInventory: false };
    const status = _deriveBinStatus(_getBinCapacityMetrics(meta, usage));
    if (typeof meta.statusIdx === 'number') {
      binSheet.getRange(i + 1, meta.statusIdx + 1).setValue(status);
      _updateSheetCacheCell(CONFIG.SHEETS.BIN, i + 1, meta.statusIdx + 1, status);
    }
  }

  const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
  if (!rackSheet) return;
  const rackData = _getSheetValuesCached(rackSheet.getName());
  const rackColMap = _getSheetHeaderMap(rackSheet);
  const rackIdIdx = typeof rackColMap['RACK_ID'] === 'number' ? rackColMap['RACK_ID'] : 0;
  const rackStatusIdx = typeof rackColMap['RACK_STATUS'] === 'number' ? rackColMap['RACK_STATUS'] : null;
  const rackSwlIdx = typeof rackColMap['RACK_SWL_KG'] === 'number'
    ? rackColMap['RACK_SWL_KG']
    : (typeof rackColMap['RACK_SWL'] === 'number' ? rackColMap['RACK_SWL'] : null);
  if (rackStatusIdx === null) {
    Logger.log('Rack status column not found; skipping rack status updates');
    return;
  }

  if (rackIds.length === 0) return;

  rackIds.forEach(rackId => {
    let rackLoad = 0;
    for (let i = 1; i < binData.length; i++) {
      const binId = String(binData[i][0] || '').trim();
      const meta = binMetaMap[binId] || {};
      if (String(meta.rackId || '').trim() !== rackId) continue;
      rackLoad += Number(usageSummary.kgByBin[binId] || 0);
    }
    let rackSwl = 0;
    if (rackSwlIdx !== null) {
      for (let i = 1; i < rackData.length; i++) {
        if (String(rackData[i][rackIdIdx] || '').trim() === rackId) {
          rackSwl = Number(rackData[i][rackSwlIdx]) || 0;
          break;
        }
      }
    }
    let rackStatus = 'PARTIAL';
    if (rackLoad <= 0) rackStatus = 'FREE';
    else if (rackSwl > 0 && rackLoad >= rackSwl) rackStatus = 'FULL';

    for (let i = 1; i < rackData.length; i++) {
      if (String(rackData[i][rackIdIdx] || '').trim() === rackId) {
        rackSheet.getRange(i + 1, rackStatusIdx + 1).setValue(rackStatus);
        _updateSheetCacheCell(CONFIG.SHEETS.RACK, i + 1, rackStatusIdx + 1, rackStatus);
        break;
      }
    }
  });
}

function _assertRackSwlForBin(binId, addQty, binDeltaMap, itemCode, inputUom) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
  if (!rackSheet) return;
  const binCtx = _getBinMasterMetaContext();
  const binData = binCtx.data || [];
  const rackData = _getSheetValuesCached(rackSheet.getName());
  const binMetaMap = binCtx.metaMap || {};
  const rackColMap = _getSheetHeaderMap(rackSheet);
  const rackIdIdx = typeof rackColMap['RACK_ID'] === 'number' ? rackColMap['RACK_ID'] : 0;
  const rackSwlIdx = typeof rackColMap['RACK_SWL_KG'] === 'number'
    ? rackColMap['RACK_SWL_KG']
    : (typeof rackColMap['RACK_SWL'] === 'number' ? rackColMap['RACK_SWL'] : null);
  if (rackSwlIdx === null) return;

  const binKey = String(binId || '').trim();
  const rackId = String(((binMetaMap[binKey] || {}).rackId) || '').trim();
  if (!rackId) return;

  const rackBinSet = {};
  for (let i = 1; i < binData.length; i++) {
    const id = String(binData[i][0] || '').trim();
    const meta = binMetaMap[id] || {};
    if (String(meta.rackId || '').trim() === rackId) rackBinSet[id] = true;
  }
  const usageSummary = _getInventoryUsageSummaryForBinSet(binMetaMap, rackBinSet, 'RACK_' + rackId);

  let rackSwl = 0;
  for (let i = 1; i < rackData.length; i++) {
    if (String(rackData[i][rackIdIdx] || '').trim() === rackId) {
      rackSwl = Number(rackData[i][rackSwlIdx]) || 0;
      break;
    }
  }
  if (!rackSwl || rackSwl <= 0) return;

  const deltaMap = binDeltaMap
    ? (binDeltaMap.__rackKgByBin || (binDeltaMap.__rackKgByBin = {}))
    : null;
  let add = 0;
  const rawAddQty = Number(addQty) || 0;
  if (rawAddQty > 0 && itemCode) {
    const sourceUom = String(inputUom || _getItemUomCode(itemCode)).trim().toUpperCase();
    const addKg = _convertItemQtyToTargetUom(itemCode, rawAddQty, sourceUom, 'KG');
    if (isFinite(addKg)) {
      add = addKg;
    } else {
      Logger.log('Rack SWL skipped for ' + (itemCode || 'item') + ' in bin ' + binId + ' because quantity cannot be normalized to KG.');
      return;
    }
  }
  let rackLoad = 0;
  for (let i = 1; i < binData.length; i++) {
    const id = String(binData[i][0] || '').trim();
    if (!rackBinSet[id]) continue;
    const base = Number(usageSummary.kgByBin[id] || 0);
    const delta = Number((deltaMap && deltaMap[id]) || 0);
    rackLoad += base + delta;
  }
  const projected = rackLoad + add;
  if (projected > rackSwl) {
    throw new Error(`Rack SWL exceeded for rack ${rackId}. SWL ${rackSwl} KG, projected ${projected.toFixed(2)} KG.`);
  }
  if (deltaMap && add > 0) {
    deltaMap[binKey] = Number(deltaMap[binKey] || 0) + add;
  }
}

// =====================================================
// 6. DASHBOARD DATA APIs
// =====================================================

/**
 * VERSION RULES (read-only):
 * - Version is a production construct (not an inventory mutation).
 * - Missing version stays UNSPECIFIED for display.
 * - No auto-defaulting to V1 in read models.
 */
function _normalizeMovementVersion(value) {
  const v = String(value || '').trim();
  return v ? v.toUpperCase() : 'UNSPECIFIED';
}

function _buildBinMetaMap(ss) {
  const meta = {};
  let binSheet;
  try { binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN); } catch (e) { return meta; }
  const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);

  const rackMap = {};
  if (rackSheet) {
    const rackData = _getSheetValuesCached(rackSheet.getName());
    for (let i = 1; i < rackData.length; i++) {
      const rackId = String(rackData[i][0] || '').trim();
      if (!rackId) continue;
      rackMap[rackId] = {
        rackCode: String(rackData[i][1] || '').trim(),
        site: String(rackData[i][2] || '').trim(),
        location: String(rackData[i][3] || '').trim()
      };
    }
  }

  const binData = _getSheetValuesCached(binSheet.getName());
  for (let i = 1; i < binData.length; i++) {
    const binId = String(binData[i][0] || '').trim();
    if (!binId) continue;
    const rackId = String(binData[i][1] || '').trim();
    const rack = rackMap[rackId] || {};
    meta[binId] = {
      binCode: String(binData[i][2] || '').trim(),
      rackCode: rack.rackCode || '',
      site: rack.site || '',
      location: rack.location || ''
    };
  }
  return meta;
}

function getStockAggregationByVersion(options) {
  const opts = options || {};
  const build = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // FIX: inventory_id is source of truth
    // Movement-based aggregation was unsafe because it can skip valid inventory rows when batch/lot resolution is inconsistent.
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const invData = _getSheetValuesCached(invSheet.getName());
    const cols = _getInventorySheetColumnMap(invSheet);
    const binMeta = _buildBinMetaMap(ss);

    const itemMaps = _getItemMasterMaps();
    const itemUomQty = {};

    const agg = {};
    const addDelta = (itemCode, version, binId, delta, inwardDate, lastTransferDate) => {
      if (!itemCode || !binId || !delta) return;
      if (!agg[itemCode]) agg[itemCode] = {};
      if (!agg[itemCode][version]) agg[itemCode][version] = {};
      if (!agg[itemCode][version][binId]) {
        agg[itemCode][version][binId] = {
          quantity: 0,
          inwardTs: NaN,
          inwardDate: '',
          transferTs: NaN,
          lastTransferDate: ''
        };
      }
      const entry = agg[itemCode][version][binId];
      entry.quantity += delta;

      const inwardTs = _toTimestamp(inwardDate, NaN);
      if (isFinite(inwardTs) && (!isFinite(entry.inwardTs) || inwardTs < entry.inwardTs)) {
        entry.inwardTs = inwardTs;
        entry.inwardDate = inwardDate;
      } else if (!entry.inwardDate && inwardDate) {
        // Fallback for non-standard date formats
        entry.inwardDate = inwardDate;
      }

      const transferTs = _toTimestamp(lastTransferDate, NaN);
      if (isFinite(transferTs) && (!isFinite(entry.transferTs) || transferTs > entry.transferTs)) {
        entry.transferTs = transferTs;
        entry.lastTransferDate = lastTransferDate;
      } else if (!entry.lastTransferDate && lastTransferDate) {
        entry.lastTransferDate = lastTransferDate;
      }
    };

    for (let i = 1; i < invData.length; i++) {
      const row = invData[i];
      const qty = Number(row[cols.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;
      // FIX: inventory_id is source of truth
      // Aggregating by inventory rows avoids re-resolving batch/lot/gin, which was skipping valid stock.
      let itemCode = String(row[cols.ITEM_CODE_CACHE] || '').trim();
      if (!itemCode) {
        const rawItemId = String(row[cols.ITEM_ID] || '').trim();
        itemCode = itemMaps.idToCode[String(rawItemId || '').toUpperCase()] || rawItemId;
      }
      itemCode = _getCanonicalItemCode(itemCode, row[cols.ITEM_ID]);
      if (!itemCode) continue;

      const version = _normalizeMovementVersion(row[cols.VERSION]);
      const binId = String(row[cols.BIN_ID] || '').trim();
      if (!binId) continue;
      const rowUom = String((typeof cols.UOM === 'number' ? row[cols.UOM] : '') || '').trim().toUpperCase();
      const effectiveUom = rowUom || itemMaps.codeToUom[String(itemCode || '').trim().toUpperCase()] || 'KG';
      if (!itemUomQty[itemCode]) itemUomQty[itemCode] = {};
      itemUomQty[itemCode][effectiveUom] = (itemUomQty[itemCode][effectiveUom] || 0) + qty;
      const inwardDate = (typeof cols.INWARD_DATE === 'number')
        ? row[cols.INWARD_DATE]
        : '';
      const lastTransferDate = (typeof cols.LAST_TRANSFER_DATE === 'number')
        ? row[cols.LAST_TRANSFER_DATE]
        : '';
      addDelta(itemCode, version, binId, qty, inwardDate, lastTransferDate);
    }

    const items = [];
    Object.keys(agg).sort().forEach(itemCode => {
      const versionMap = agg[itemCode];
      const versions = [];
      let itemTotal = 0;
      Object.keys(versionMap).sort().forEach(version => {
        const binMap = versionMap[version];
        const bins = [];
        let vTotal = 0;
        Object.keys(binMap).sort().forEach(binId => {
          const entry = binMap[binId] || {};
          const quantity = Number(entry.quantity) || 0;
          if (quantity === 0) return;
          vTotal += quantity;
          const meta = binMeta[binId] || {};
          bins.push({
            binId: binId,
            binCode: meta.binCode || binId,
            site: meta.site || '',
            location: meta.location || '',
            rackCode: meta.rackCode || '',
            quantity: quantity,
            inwardDate: _formatDateForUi(entry.inwardDate),
            lastTransferDate: _formatDateForUi(entry.lastTransferDate)
          });
        });
        itemTotal += vTotal;
        versions.push({
          version: version,
          totalQty: vTotal,
          bins: bins
        });
      });
      items.push({
        itemCode: itemCode,
        uomCode: (function () {
          const byUom = itemUomQty[itemCode] || {};
          const keys = Object.keys(byUom);
          if (keys.length === 0) {
            return itemMaps.codeToUom[String(itemCode || '').trim().toUpperCase()] || _getItemUomCode(itemCode);
          }
          keys.sort(function (a, b) { return Number(byUom[b] || 0) - Number(byUom[a] || 0); });
          return keys[0] || itemMaps.codeToUom[String(itemCode || '').trim().toUpperCase()] || _getItemUomCode(itemCode);
        })(),
        totalQty: itemTotal,
        versions: versions
      });
    });
    return items;
  };

  if (opts.skipAuth === true) return _withRequestCache(build);
  return protect(function () {
    return build();
  });
}

function _getQaStatusSummaryByItem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const cols = _getInventorySheetColumnMap(invSheet);
  const summary = {};

  for (let i = 1; i < data.length; i++) {
    const qty = Number(data[i][cols.TOTAL_QUANTITY]) || 0;
    if (qty <= 0) continue;
    const itemCode = _getCanonicalItemCode(data[i][cols.ITEM_CODE_CACHE], data[i][cols.ITEM_ID]);
    if (!itemCode) continue;
    let status = _normalizeQaStatus(data[i][cols.QUALITY_STATUS] || 'PENDING');
    if (!summary[itemCode]) {
      summary[itemCode] = { approved: 0, pending: 0, rejected: 0, hold: 0, overridden: 0, closed: 0 };
    }
    if (status === 'APPROVED') summary[itemCode].approved += qty;
    else if (status === 'HOLD') summary[itemCode].hold += qty;
    else if (status === 'REJECTED') summary[itemCode].rejected += qty;
    else if (status === 'OVERRIDDEN') summary[itemCode].overridden += qty;
    else if (status === 'CLOSED') summary[itemCode].closed += qty;
    else summary[itemCode].pending += qty;
  }
  return summary;
}

function getWarehouseReadModel(options) {
  const opts = options || {};
  const build = () => {
    const items = getStockAggregationByVersion({ skipAuth: true }) || [];
    const qaSummaryByItem = _getQaStatusSummaryByItem();
    const itemMaps = _getItemMasterMaps();
    items.forEach(item => {
      const codeNorm = String(item.itemCode || '').trim().toUpperCase();
      item.itemName = String(itemMaps.codeToName[codeNorm] || item.itemCode || '');
      item.qaSummary = qaSummaryByItem[item.itemCode] || { approved: 0, pending: 0, rejected: 0, hold: 0, overridden: 0, closed: 0 };
    });
    return { items: items, qaSummaryByItem: qaSummaryByItem };
  };
  if (opts.skipAuth === true) return _withRequestCache(build);
  return protect(function () {
    return build();
  });
}


/**
 * Safely resolve dead stock sheet by name with fallbacks.
 * Returns the sheet object or null. NEVER creates a sheet; read-only operation.
 */
function _resolveDeadStockSheet(ss) {
  if (!ss) return null;
  var deadSheetName = String(CONFIG.SHEETS.DEAD_STOCK || 'dead_stock_log');
  var deadSheetNameNormLower = deadSheetName.trim().toLowerCase();

  // Strict: exact configured sheet name.
  var exactOps = ss.getSheetByName(deadSheetName);
  if (exactOps) {
    Logger.log('[DEAD_STOCK] Sheet selected by exact name: ' + deadSheetName);
    return exactOps;
  }

  // Strict: case-insensitive + trimmed exact match only.
  var sheets = ss.getSheets();
  for (var ei = 0; ei < sheets.length; ei++) {
    var nmRaw = String(sheets[ei].getName() || '');
    if (nmRaw.trim().toLowerCase() === deadSheetNameNormLower) {
      Logger.log('[DEAD_STOCK] Sheet selected by trimmed/case-insensitive name: ' + sheets[ei].getName());
      return sheets[ei];
    }
  }
  Logger.log('[DEAD_STOCK] Strict sheet not found: ' + deadSheetName +
    '. Available sheets: ' + sheets.map(function (s) { return s.getName(); }).join(', '));
  return null;
}

/**
 * Column index resolver with diagnostic logging.
 * Returns null (not a number) if column genuinely not found.
 */
function _safeColIdx(headerMap, candidates, fallbackIdx, fieldName) {
  for (var i = 0; i < candidates.length; i++) {
    var key = String(candidates[i] || '').trim().toUpperCase();
    if (key && typeof headerMap[key] === 'number') {
      return headerMap[key];
    }
  }
  if (typeof fallbackIdx === 'number') {
    Logger.log('[DEAD_STOCK] Column ' + fieldName + ' not found in headers; using CONFIG fallback index ' + fallbackIdx);
    return fallbackIdx;
  }
  Logger.log('[DEAD_STOCK] WARNING: Column ' + fieldName + ' not found in headers and no fallback defined. Returning null.');
  return null;
}

/**
 * Loose number parser for dead stock quantities.
 */
function _toNumberLooseDS(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  var txt = String(value || '').replace(/,/g, '').trim();
  if (!txt) return 0;
  var m = txt.match(/-?\d+(\.\d+)?/);
  return m ? (Number(m[0]) || 0) : 0;
}


function getWarehouseItemDashboardData() {
  return protect(function () {
    const model = getWarehouseReadModel({ skipAuth: true });
    return model.items || [];
  });
}

/**
 * Packaging Version Lifecycle System: Data API
 * Filters for items starting with 'PM' and merges with Packing_Version_Master metadata.
 */
function getPackingMaterialDashboardData() {
  return protect(function () {
    const cacheKey = 'PACKING_MATERIAL_DASHBOARD_V2';
    const cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    const model = getWarehouseReadModel({ skipAuth: true });
    const items = model.items || [];

    // Filter for Packing Materials only (Item Code starts with PM)
    const pmItems = items.filter(it => String(it.itemCode || '').trim().toUpperCase().startsWith('PM'));

    // Fetch Latest Mappings
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PACKING_VERSION_MASTER);
    const mappings = {};

    if (sheet) {
      _ensurePackingVersionMasterSchema(sheet);
      const data = sheet.getDataRange().getValues();
      const C = CONFIG.PACKING_VERSION_COLS;
      for (let i = 1; i < data.length; i++) {
        const itemCode = String(data[i][C.ITEM_CODE] || '').trim().toUpperCase();
        const version = String(data[i][C.SYSTEM_VERSION] || '').trim().toUpperCase();
        const binId = String(data[i][C.BIN_ID] || '').trim().toUpperCase();
        const isLatest = data[i][C.IS_LATEST] === true || String(data[i][C.IS_LATEST]).toUpperCase() === 'TRUE';

        if (isLatest && itemCode && version && binId) {
          const key = `${itemCode}||${version}||${binId}`;
          mappings[key] = {
            isLatest: true,
            reason: String(data[i][C.REASON] || ''),
            latestName: String(data[i][C.LATEST_NAME] || '')
          };
        }
      }
    }

    // Inject Latest status into Bins
    pmItems.forEach(it => {
      let itemHasLatest = false;
      (it.versions || []).forEach(v => {
        (v.bins || []).forEach(b => {
          const key = `${String(it.itemCode).toUpperCase()}||${String(v.version).toUpperCase()}||${String(b.binId).toUpperCase()}`;
          if (mappings[key]) {
            b.isLatest = true;
            b.latestReason = mappings[key].reason;
            b.latestName = mappings[key].latestName;
            itemHasLatest = true;
          } else {
            b.isLatest = false;
            b.latestName = '';
          }
        });
      });
      it.hasLatest = itemHasLatest;
    });

    _putScriptCacheJson(cacheKey, pmItems, 45);
    return pmItems;
  });
}

function _ensurePackingVersionMasterSchema(sheet) {
  if (!sheet) return;
  const C = CONFIG.PACKING_VERSION_COLS;
  const requiredHeaders = [];
  requiredHeaders[C.MAPPING_ID] = 'mapping_id';
  requiredHeaders[C.ITEM_CODE] = 'item_code';
  requiredHeaders[C.SYSTEM_VERSION] = 'system_version';
  requiredHeaders[C.BIN_ID] = 'bin_id';
  requiredHeaders[C.IS_LATEST] = 'is_latest';
  requiredHeaders[C.REASON] = 'reason';
  requiredHeaders[C.UPDATED_BY] = 'updated_by';
  requiredHeaders[C.UPDATED_AT] = 'updated_at';
  requiredHeaders[C.LATEST_NAME] = 'latest_name';

  const width = requiredHeaders.length;
  const current = sheet.getRange(1, 1, 1, Math.max(width, sheet.getLastColumn() || width)).getValues()[0] || [];
  let changed = false;
  for (let i = 0; i < requiredHeaders.length; i++) {
    if (!String(current[i] || '').trim()) {
      current[i] = requiredHeaders[i];
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(1, 1, 1, current.length).setValues([current]);
  }
}

/**
 * Assigns or updates the 'Latest' status for a specific PM bin.
 */
function setPackingVersionLatest(payload) {
  return withScriptLock(function () {
    return protect(function (user) {
      requireRole(SECURITY.ROLES.MANAGER);
      const itemCode = String(payload.itemCode || '').trim().toUpperCase();
      const version = String(payload.version || '').trim().toUpperCase();
      const binId = String(payload.binId || '').trim().toUpperCase();
      const isLatest = payload.isLatest === true;
      const reason = String(payload.reason || '').trim();
      const latestName = String(payload.latestName || payload.labelName || '').trim().toUpperCase();

      if (!itemCode || !version || !binId) throw new Error('Missing itemCode, version, or binId');
      if (isLatest && !latestName) throw new Error('Latest name required');

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(CONFIG.SHEETS.PACKING_VERSION_MASTER);
      if (!sheet) {
        sheet = ss.insertSheet(CONFIG.SHEETS.PACKING_VERSION_MASTER);
        sheet.appendRow(['mapping_id', 'item_code', 'system_version', 'bin_id', 'is_latest', 'reason', 'updated_by', 'updated_at', 'latest_name']);
      }
      _ensurePackingVersionMasterSchema(sheet);

      const data = sheet.getDataRange().getValues();
      const C = CONFIG.PACKING_VERSION_COLS;
      let foundRowIdx = -1;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][C.ITEM_CODE]).toUpperCase() === itemCode &&
          String(data[i][C.SYSTEM_VERSION]).toUpperCase() === version &&
          String(data[i][C.BIN_ID]).toUpperCase() === binId) {
          foundRowIdx = i + 1;
          break;
        }
      }

      const timestamp = new Date();
      const updatedBy = user.email;

      // If setting to latest, unmark all other entries for this item_code first
      if (isLatest) {
        for (let i = 1; i < data.length; i++) {
          const rowItemCode = String(data[i][C.ITEM_CODE] || '').trim().toUpperCase();
          const rowIsLatest = data[i][C.IS_LATEST] === true || String(data[i][C.IS_LATEST]).toUpperCase() === 'TRUE';
          if (rowItemCode === itemCode && rowIsLatest) {
            sheet.getRange(i + 1, C.IS_LATEST + 1).setValue(false);
            sheet.getRange(i + 1, C.UPDATED_BY + 1).setValue(updatedBy);
            sheet.getRange(i + 1, C.UPDATED_AT + 1).setValue(timestamp);
            sheet.getRange(i + 1, C.REASON + 1).setValue('Superseded by new latest');
            sheet.getRange(i + 1, C.LATEST_NAME + 1).setValue('');
          }
        }
      }

      if (foundRowIdx !== -1) {
        sheet.getRange(foundRowIdx, C.IS_LATEST + 1).setValue(isLatest);
        sheet.getRange(foundRowIdx, C.REASON + 1).setValue(reason);
        sheet.getRange(foundRowIdx, C.UPDATED_BY + 1).setValue(updatedBy);
        sheet.getRange(foundRowIdx, C.UPDATED_AT + 1).setValue(timestamp);
        sheet.getRange(foundRowIdx, C.LATEST_NAME + 1).setValue(isLatest ? latestName : '');
      } else {
        const mappingId = 'PM-L-' + timestamp.getTime();
        const newRow = [];
        newRow[C.MAPPING_ID] = mappingId;
        newRow[C.ITEM_CODE] = itemCode;
        newRow[C.SYSTEM_VERSION] = version;
        newRow[C.BIN_ID] = binId;
        newRow[C.IS_LATEST] = isLatest;
        newRow[C.REASON] = reason;
        newRow[C.UPDATED_BY] = updatedBy;
        newRow[C.UPDATED_AT] = timestamp;
        newRow[C.LATEST_NAME] = isLatest ? latestName : '';
        sheet.appendRow(newRow);
      }
      _clearSheetCache(CONFIG.SHEETS.PACKING_VERSION_MASTER);

      return { success: true };
    });
  });
}

function _applyPackingVersionMappingsFromInward(updates, userEmail) {
  if (!updates || updates.length === 0) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.PACKING_VERSION_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.PACKING_VERSION_MASTER);
    sheet.appendRow(['mapping_id', 'item_code', 'system_version', 'bin_id', 'is_latest', 'reason', 'updated_by', 'updated_at', 'latest_name']);
  }
  _ensurePackingVersionMasterSchema(sheet);

  const C = CONFIG.PACKING_VERSION_COLS;
  const width = C.LATEST_NAME + 1;
  const data = sheet.getDataRange().getValues();
  while (data.length < 1) data.push([]);
  while (data[0].length < width) data[0].push('');

  const now = new Date();
  const email = String(userEmail || Session.getActiveUser().getEmail() || '').trim();
  const rowByKey = {};
  const changedRows = {};
  for (let i = 1; i < data.length; i++) {
    while (data[i].length < width) data[i].push('');
    const key = [
      String(data[i][C.ITEM_CODE] || '').trim().toUpperCase(),
      String(data[i][C.SYSTEM_VERSION] || '').trim().toUpperCase(),
      String(data[i][C.BIN_ID] || '').trim().toUpperCase()
    ].join('||');
    rowByKey[key] = i;
  }

  const appendRows = [];
  updates.forEach(function (u) {
    const itemCode = String(u.itemCode || '').trim().toUpperCase();
    const version = String(u.version || '').trim().toUpperCase();
    const binId = String(u.binId || '').trim().toUpperCase();
    const labelType = String(u.labelType || '').trim().toUpperCase();
    const latestName = String(u.latestName || '').trim().toUpperCase();
    if (!itemCode || !version || !binId || !labelType) return;

    const isLatest = labelType === 'LATEST';
    if (isLatest && !latestName) return;

    if (isLatest) {
      for (let i = 1; i < data.length; i++) {
        const rowItem = String(data[i][C.ITEM_CODE] || '').trim().toUpperCase();
        const rowIsLatest = data[i][C.IS_LATEST] === true || String(data[i][C.IS_LATEST]).toUpperCase() === 'TRUE';
        if (rowItem === itemCode && rowIsLatest) {
          data[i][C.IS_LATEST] = false;
          data[i][C.REASON] = 'Superseded by PM inward ' + latestName;
          data[i][C.UPDATED_BY] = email;
          data[i][C.UPDATED_AT] = now;
          data[i][C.LATEST_NAME] = '';
          changedRows[i] = true;
        }
      }
    }

    const key = [itemCode, version, binId].join('||');
    let idx = rowByKey[key];
    if (typeof idx !== 'number') {
      const newRow = [];
      while (newRow.length < width) newRow.push('');
      newRow[C.MAPPING_ID] = 'PM-L-' + now.getTime() + '-' + Math.floor(Math.random() * 1000);
      newRow[C.ITEM_CODE] = itemCode;
      newRow[C.SYSTEM_VERSION] = version;
      newRow[C.BIN_ID] = binId;
      appendRows.push(newRow);
      idx = data.length + appendRows.length - 1;
      rowByKey[key] = idx;
    }

    const row = idx < data.length ? data[idx] : appendRows[idx - data.length];
    row[C.IS_LATEST] = isLatest;
    row[C.REASON] = isLatest ? ('PM inward latest label ' + latestName) : 'PM inward obsolete label';
    row[C.UPDATED_BY] = email;
    row[C.UPDATED_AT] = now;
    row[C.LATEST_NAME] = isLatest ? latestName : '';
    if (idx < data.length) changedRows[idx] = true;
  });

  Object.keys(changedRows).forEach(function (idxText) {
    const idx = Number(idxText);
    if (!isFinite(idx) || idx <= 0 || idx >= data.length) return;
    const row = data[idx].slice(0, width);
    while (row.length < width) row.push('');
    sheet.getRange(idx + 1, 1, 1, width).setValues([row]);
  });
  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, width).setValues(appendRows);
  }
  _clearSheetCache(CONFIG.SHEETS.PACKING_VERSION_MASTER);
  return updates.length;
}

function syncPackingLabelsFromInventoryLotLabels() {
  return withScriptLock(function () {
    const user = protect(() => requireRole(SECURITY.ROLES.MANAGER));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());
    const updates = [];

    for (let i = 1; i < data.length; i++) {
      const itemCode = String(data[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
      if (!/^PM/i.test(itemCode)) continue;

      const lotNo = String(data[i][CONFIG.INVENTORY_COLS.LOT_NO] || '').trim();
      const version = String(data[i][CONFIG.INVENTORY_COLS.VERSION] || '').trim();
      const binId = String(data[i][CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      if (!lotNo || !version || !binId) continue;

      const upper = lotNo.toUpperCase();
      if (upper.indexOf('LATEST:') === 0) {
        updates.push({
          itemCode: itemCode,
          version: version,
          binId: binId,
          labelType: 'LATEST',
          latestName: lotNo.slice(lotNo.indexOf(':') + 1).trim()
        });
      } else if (upper === 'OBSOLETE') {
        updates.push({
          itemCode: itemCode,
          version: version,
          binId: binId,
          labelType: 'OBSOLETE',
          latestName: ''
        });
      }
    }

    const synced = _applyPackingVersionMappingsFromInward(updates, user && user.email);
    return { success: true, scanned: Math.max(0, data.length - 1), synced: synced };
  });
}

/**
 * getTransferStockByItemCode
 * Used by TransferForm source picker.
 * Reads inventory the same way getWarehouseLayout does:
 *   - Direct column index access (no _resolveInventoryItemCode)
 *   - item_id → Item_Master map for code resolution (no CacheService)
 *   - Returns ONLY plain strings/numbers — no rowData, no Date objects
 *   - Safe for google.script.run serialization
 *
 * @param  {string} itemCode  e.g. "ST27"
 * @returns {Array<{batchId,batchNumber,binId,binCode,site,location,qty,uom,qaStatus,version}>}
 */
function getTransferStockByItemCode(itemCode) {
  return protect(function () {
    var code = String(itemCode || '').trim().toUpperCase();
    if (!code) return [];

    var cacheKey = 'TRANSFER_STOCK_V3_' + code;
    var cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    var invData = _getSheetValuesCached(invSheet.getName());

    // Build item_id → code map directly from Item_Master (same as getWarehouseLayout)
    var itemMap = {};
    var itemSheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);
    if (itemSheet) {
      var itemData = _getSheetValuesCached(itemSheet.getName());
      var itemCols = _getItemSheetColumnMap(itemSheet);
      for (var m = 1; m < itemData.length; m++) {
        var iId = String(itemData[m][itemCols.ITEM_ID] || '').trim();
        var iCode = String(itemData[m][itemCols.ITEM_CODE] || '').trim();
        if (iId) itemMap[iId] = iCode;
        if (iCode) itemMap['__CODE__' + iCode.toUpperCase()] = iCode;
      }
    }

    // Build bin_id → {binCode, site, location} map from Bin + Rack masters
    var binMeta = _buildBinMetaMap(ss);   // already returns plain strings

    var results = [];
    for (var i = 1; i < invData.length; i++) {
      var row = invData[i];
      var qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      // Resolve item code: item_id first, fall back to item_code_cache
      var rawId = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim();
      var rawCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
      var resolved = itemMap[rawId]
        || itemMap['__CODE__' + rawCode.toUpperCase()]
        || rawCode;

      if (String(resolved || '').trim().toUpperCase() !== code) continue;

      var batchId = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
      var batchInfo = _resolveBatchReference(code, batchId);
      var binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      var meta = binMeta[binId] || {};

      results.push({
        batchId: String((batchInfo && batchInfo.batchId) || batchId),
        batchNumber: String((batchInfo && batchInfo.batchNumber) || batchId),
        binId: binId,
        binCode: String(meta.binCode || binId),
        site: String(meta.site || ''),
        location: String(meta.location || ''),
        qty: qty,
        uom: String(row[CONFIG.INVENTORY_COLS.UOM] || 'NOS').trim() || 'NOS',
        qaStatus: String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').trim().toUpperCase(),
        version: String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim()
      });
    }

    // APPROVED/OVERRIDDEN first, then highest qty
    results.sort(function (a, b) {
      var okA = (a.qaStatus === 'APPROVED' || a.qaStatus === 'OVERRIDDEN') ? 1 : 0;
      var okB = (b.qaStatus === 'APPROVED' || b.qaStatus === 'OVERRIDDEN') ? 1 : 0;
      return okB !== okA ? okB - okA : b.qty - a.qty;
    });

    _putScriptCacheJson(cacheKey, results, 45);
    return results;
  });
}

function _buildWarehouseItemsFallbackFromInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const invData = _getSheetValuesCached(invSheet.getName());
  const cols = _getInventorySheetColumnMap(invSheet);
  const binMeta = _buildBinMetaMap(ss);
  const itemMaps = _getItemMasterMaps();
  const qaSummaryByItem = _getQaStatusSummaryByItem();
  const agg = {};

  for (let i = 1; i < invData.length; i++) {
    const row = invData[i];
    const qty = Number(row[cols.TOTAL_QUANTITY]) || 0;
    if (qty <= 0) continue;

    const status = _normalizeQaStatus(row[cols.QUALITY_STATUS] || 'PENDING');
    if (status === 'REJECTED' || status === 'HOLD') continue;

    let itemCode = String(row[cols.ITEM_CODE_CACHE] || '').trim();
    if (!itemCode) {
      const itemIdRaw = String(row[cols.ITEM_ID] || '').trim().toUpperCase();
      itemCode = itemMaps.idToCode[itemIdRaw] || itemIdRaw;
    }
    itemCode = _getCanonicalItemCode(itemCode, row[cols.ITEM_ID]);
    if (!itemCode) continue;

    const version = _normalizeMovementVersion(row[cols.VERSION]);
    const binId = String(row[cols.BIN_ID] || '').trim();
    if (!binId) continue;
    const rowUom = String((typeof cols.UOM === 'number' ? row[cols.UOM] : '') || '').trim().toUpperCase();
    const uom = rowUom || itemMaps.codeToUom[String(itemCode || '').trim().toUpperCase()] || 'KG';
    const inwardDate = (typeof cols.INWARD_DATE === 'number')
      ? _formatDateForUi(row[cols.INWARD_DATE])
      : '';
    const lastTransferDate = (typeof cols.LAST_TRANSFER_DATE === 'number')
      ? _formatDateForUi(row[cols.LAST_TRANSFER_DATE])
      : '';

    if (!agg[itemCode]) agg[itemCode] = { totalQty: 0, uom: uom, versions: {} };
    if (!agg[itemCode].versions[version]) agg[itemCode].versions[version] = {};
    if (!agg[itemCode].versions[version][binId]) {
      agg[itemCode].versions[version][binId] = {
        quantity: 0,
        inwardDate: '',
        lastTransferDate: ''
      };
    }

    agg[itemCode].totalQty += qty;
    const binEntry = agg[itemCode].versions[version][binId];
    binEntry.quantity += qty;
    if (!binEntry.inwardDate && inwardDate) binEntry.inwardDate = inwardDate;
    if (lastTransferDate) binEntry.lastTransferDate = lastTransferDate;
  }

  return Object.keys(agg).sort().map(function (itemCode) {
    const info = agg[itemCode];
    const codeNorm = String(itemCode || '').trim().toUpperCase();
    const versions = Object.keys(info.versions).sort().map(function (version) {
      const bins = Object.keys(info.versions[version]).sort().map(function (binId) {
        const entry = info.versions[version][binId];
        const meta = binMeta[binId] || {};
        return {
          binId: binId,
          binCode: meta.binCode || binId,
          site: meta.site || '',
          location: meta.location || '',
          rackCode: meta.rackCode || '',
          quantity: Number(entry.quantity) || 0,
          inwardDate: entry.inwardDate || '',
          lastTransferDate: entry.lastTransferDate || ''
        };
      });
      const versionQty = bins.reduce(function (a, b) { return a + (Number(b.quantity) || 0); }, 0);
      return {
        version: version,
        totalQty: versionQty,
        bins: bins
      };
    });
    return {
      itemCode: itemCode,
      itemName: String(itemMaps.codeToName[codeNorm] || itemCode),
      uomCode: info.uom || 'KG',
      totalQty: Number(info.totalQty) || 0,
      qaSummary: qaSummaryByItem[itemCode] || { approved: 0, pending: 0, rejected: 0, hold: 0, overridden: 0, closed: 0 },
      versions: versions
    };
  });
}

function _buildWarehouseItemsFallbackFromSnapshot() {
  const invAgg = _aggregateInventorySnapshotByItem();
  const qaSummaryByItem = _getQaStatusSummaryByItem();
  const itemMaps = _getItemMasterMaps();
  const emptyQa = { approved: 0, pending: 0, rejected: 0, hold: 0, overridden: 0, closed: 0 };

  return Object.keys(invAgg || {}).sort().map(function (itemCode) {
    const row = invAgg[itemCode] || {};
    const siteMap = row.bySiteLocation || {};
    const bins = Object.keys(siteMap).sort().map(function (key) {
      const parts = String(key || '').split('||');
      const site = String(parts[0] || '').trim();
      const location = String(parts[1] || '').trim();
      return {
        binId: 'ALL::' + site + '::' + location,
        binCode: 'ALL',
        site: site,
        location: location,
        rackCode: '',
        quantity: Number(siteMap[key]) || 0,
        inwardDate: '',
        lastTransferDate: ''
      };
    });

    const totalQty = Number(row.totalQty) || 0;
    const codeNorm = String(itemCode || '').trim().toUpperCase();
    return {
      itemCode: itemCode,
      itemName: String(itemMaps.codeToName[codeNorm] || itemCode),
      uomCode: String(itemMaps.codeToUom[codeNorm] || 'KG'),
      totalQty: totalQty,
      qaSummary: qaSummaryByItem[itemCode] || emptyQa,
      versions: [{
        version: 'UNSPECIFIED',
        totalQty: totalQty,
        bins: bins
      }]
    };
  });
}

function getWarehouseDashboardPayload() {
  return protect(function () {
    const model = getWarehouseReadModel({ skipAuth: true });
    let items = model.items || [];
    let fallbackUsed = '';
    if (!items.length) {
      const fallbackItems = _buildWarehouseItemsFallbackFromInventory();
      if (fallbackItems.length) {
        items = fallbackItems;
        fallbackUsed = 'inventory_rows';
      }
    }
    if (!items.length) {
      const snapshotItems = _buildWarehouseItemsFallbackFromSnapshot();
      if (snapshotItems.length) {
        items = snapshotItems;
        fallbackUsed = 'snapshot_aggregate';
      }
    }
    if (fallbackUsed) {
      Logger.log('[WAREHOUSE_DASHBOARD] Primary aggregation returned 0; fallback used=' + fallbackUsed + '; item_count=' + items.length);
    }
    const deadStock = _getDeadStockSummary();
    return {
      items: items,
      deadStock: deadStock
    };
  });
}

function getWorkerDashboardData() {
  return protect(function (user) {
    requireOperationalUser();
    return {
      role: 'WORKER',
      stats: _getInventorySummaryStats(false),
      movements: _getUserRecentMovements(user.email, 20)
    };
  });
}

function getManagerDashboardData() {
  return protect(function (user) {
    requireRole(SECURITY.ROLES.MANAGER);
    return {
      role: 'MANAGER',
      stats: _getInventorySummaryStats(true),
      activity: _getSystemRecentMovements(50),
      alerts: _getSystemAlerts()
    };
  });
}

function _getInventorySummaryStats(includeFinancials) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());

  let totalQty = 0;
  let pendingQuality = 0;
  const skuMap = {};

  for (let i = 1; i < data.length; i++) {
    const qty = Number(data[i][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (qty <= 0) continue;
    const status = _normalizeQaStatus(data[i][CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');
    const code = String(data[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
    if (status === 'APPROVED' || status === 'OVERRIDDEN') {
      totalQty += qty;
      if (code) skuMap[code] = true;
    } else if (status === 'PENDING' || status === 'HOLD') {
      pendingQuality += qty;
    }
  }

  return {
    totalQty: totalQty,
    uniqueSkus: Object.keys(skuMap).length,
    totalValue: includeFinancials ? 0 : undefined,
    pendingQA: pendingQuality
  };
}

function _getUserRecentMovements(email, limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try { sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT); } catch (e) { return []; }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const cap = Math.max(Number(limit) || 20, 1);
  const scanLimit = Math.max(cap * 10, 200);
  const startRow = Math.max(2, lastRow - scanLimit + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
  const results = [];

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][CONFIG.MOVEMENT_COLS.USER_EMAIL] === email) {
      results.push({
        id: data[i][CONFIG.MOVEMENT_COLS.MOVEMENT_ID],
        type: data[i][CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE],
        item: data[i][CONFIG.MOVEMENT_COLS.ITEM_ID],
        qty: data[i][CONFIG.MOVEMENT_COLS.QUANTITY],
        uom: String((CONFIG.MOVEMENT_COLS.UOM !== undefined ? data[i][CONFIG.MOVEMENT_COLS.UOM] : '') || '').trim(),
        date: data[i][CONFIG.MOVEMENT_COLS.TIMESTAMP]
      });
    }
  }
  return results.slice(0, cap);
}

function _getSystemRecentMovements(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  try { sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT); } catch (e) { return []; }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const startRow = Math.max(2, lastRow - limit);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();

  return data.reverse().map(row => ({
    id: row[CONFIG.MOVEMENT_COLS.MOVEMENT_ID],
    type: row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE],
    user: row[CONFIG.MOVEMENT_COLS.USER_EMAIL],
    qty: row[CONFIG.MOVEMENT_COLS.QUANTITY],
    uom: String((CONFIG.MOVEMENT_COLS.UOM !== undefined ? row[CONFIG.MOVEMENT_COLS.UOM] : '') || '').trim(),
    date: row[CONFIG.MOVEMENT_COLS.TIMESTAMP]
  }));
}

function _getSystemAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const alerts = [];

  // 1. Check Min Stock Levels
  try {
    const minStockMap = {}; // ItemCode -> { name, limit, current }
    const itemMaps = _getItemMasterMaps();
    Object.keys(itemMaps.codeToMinStock).forEach(function (codeNorm) {
      const minLevel = Number(itemMaps.codeToMinStock[codeNorm] || 0);
      if (!(minLevel > 0)) return;
      const code = itemMaps.codeDisplayByNorm[codeNorm] || codeNorm;
      minStockMap[code] = {
        name: String(itemMaps.codeToName[codeNorm] || code),
        uom: String(itemMaps.codeToUom[codeNorm] || 'KG').trim() || 'KG',
        min: minLevel,
        current: 0
      };
    });

    const invAgg = _aggregateInventorySnapshotByItem();
    Object.keys(invAgg).forEach(code => {
      if (minStockMap[code]) {
        minStockMap[code].current += Number(invAgg[code].totalQty || 0);
      }
    });

    // Generate Alerts
    for (const id in minStockMap) {
      const item = minStockMap[id];
      if (item.current < item.min) {
        alerts.push({
          type: 'CRITICAL',
          message: `Low Stock: ${item.name} (${item.current} / ${item.min} ${item.uom || 'KG'})`
        });
      }
    }

  } catch (e) {
    alerts.push({ type: 'WARNING', message: 'Failed to check stock levels: ' + e.message });
  }

  if (alerts.length === 0) {
    alerts.push({ type: 'INFO', message: 'System healthy. All stock levels normal.' });
  }

  return alerts;
}

// Snapshot aggregation for alerts (Inventory_Balance is the current state source).
function _aggregateInventorySnapshotByItem(options) {
  const opts = options || {};
  const allowedStatuses = Array.isArray(opts.allowedStatuses) ? opts.allowedStatuses.map(s => String(s).trim().toUpperCase()) : null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const agg = {};
  for (let i = 1; i < data.length; i++) {
    const qty = Number(data[i][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (qty <= 0) continue;
    const status = _normalizeQaStatus(data[i][CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');

    // Default: exclude REJECTED and HOLD unless explicitly asked for or allStatuses allowed
    if (allowedStatuses) {
      if (allowedStatuses.indexOf(status) === -1) continue;
    } else if (!opts.allStatuses) {
      if (status === 'REJECTED' || status === 'HOLD') continue;
    }
    const code = _getCanonicalItemCode(
      data[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE],
      data[i][CONFIG.INVENTORY_COLS.ITEM_ID]
    );
    if (!code) continue;
    const site = String(data[i][CONFIG.INVENTORY_COLS.SITE] || '').trim();
    const location = String(data[i][CONFIG.INVENTORY_COLS.LOCATION] || '').trim();
    if (!agg[code]) agg[code] = { totalQty: 0, bySiteLocation: {} };
    agg[code].totalQty += qty;
    const key = `${site}||${location}`;
    agg[code].bySiteLocation[key] = (agg[code].bySiteLocation[key] || 0) + qty;
  }
  return agg;
}


// One-time helper to ensure a daily trigger exists for min stock alerts.
function ensureMinStockAlertTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === 'sendMinStockAlerts');
  if (hasTrigger) return;
  ScriptApp.newTrigger('sendMinStockAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function _getActiveManagerAlertEmails_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aclSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ACCESS_CONTROL);
  const data = _getSheetValuesCached(aclSheet.getName());
  if (!data || data.length <= 1) return [];

  const headers = data[0] || [];
  const emailCol = Math.max(0, _findHeaderColumnLoose(headers, ['EMAIL', 'USER_EMAIL', 'USER EMAIL']));
  const roleCol = _findHeaderColumnLoose(headers, ['ROLE', 'USER_ROLE', 'USER ROLE']);
  const statusCol = _findHeaderColumnLoose(headers, ['STATUS', 'USER_STATUS', 'USER STATUS']);
  const resolvedRoleCol = roleCol >= 0 ? roleCol : 1;
  const resolvedStatusCol = statusCol >= 0 ? statusCol : 3;
  const seen = {};
  const emails = [];

  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][emailCol] || '').trim();
    const emailKey = email.toLowerCase();
    const role = String(data[i][resolvedRoleCol] || '').trim().toUpperCase();
    const status = String(data[i][resolvedStatusCol] || '').trim().toUpperCase();
    if (!email || seen[emailKey]) continue;
    if (role !== SECURITY.ROLES.MANAGER) continue;
    if (status && status !== 'ACTIVE') continue;
    seen[emailKey] = true;
    emails.push(email);
  }
  return emails;
}

function _getMinStockAlertRecipients_() {
  const props = PropertiesService.getScriptProperties();
  const managerEmails = _getActiveManagerAlertEmails_();
  if (managerEmails.length > 0) {
    props.setProperty('ALERT_EMAILS', managerEmails.join(','));
    return managerEmails;
  }

  const raw = String(props.getProperty('ALERT_EMAILS') || props.getProperty('ALERT_EMAIL') || '').trim();
  return raw.split(/[;,]/).map(function (e) { return e.trim(); }).filter(Boolean);
}

/**
 * Configures min-stock alert recipients from Access_Control_List managers
 * and ensures the daily trigger exists.
 *
 * @returns {{success:boolean, recipients:number, emails:string[]}}
 */
function setupMinStockAlertsForManagers() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);

    const emails = _getActiveManagerAlertEmails_();
    if (emails.length === 0) {
      throw new Error('No active MANAGER emails found in Access_Control_List.');
    }

    PropertiesService.getScriptProperties().setProperty('ALERT_EMAILS', emails.join(','));
    ensureMinStockAlertTrigger();

    return { success: true, recipients: emails.length, emails: emails };
  });
}

// One-time helper to ensure a monthly archival trigger exists for movement logs.
function ensureMovementArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === 'archiveOldMovementLogs');
  if (hasTrigger) return;
  ScriptApp.newTrigger('archiveOldMovementLogs')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
}

/**
 * Moves old Movement_Audit_Log rows to inventory_movement_archive.
 * Keeps recent rows in active sheet for fast dashboard queries.
 *
 * @param {number} daysToKeep optional, default 90
 * @returns {{success:boolean, archived:number, keptDays:number}}
 */
function archiveOldMovementLogs(daysToKeep) {
  protect(() => requireRole(SECURITY.ROLES.MANAGER));
  return withScriptLock(function () {
    return _withRequestCache(function () {
      const keepDays = Math.max(Number(daysToKeep) || 90, 30);
      const cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - keepDays);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const moveSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
      let archiveSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY_MOVEMENT_ARCHIVE);
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(CONFIG.SHEETS.INVENTORY_MOVEMENT_ARCHIVE);
      }

      if (archiveSheet.getLastRow() === 0) {
        const hdr = [];
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.MOVEMENT_ID] = 'movement_id';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.MOVEMENT_DATE] = 'movement_date';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.ITEM_ID] = 'item_id';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.BATCH_ID] = 'batch_id';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_SITE] = 'from_site';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_LOCATION] = 'from_location';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_RACK] = 'from_rack';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_BIN] = 'from_bin';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_SITE] = 'to_site';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_LOCATION] = 'to_location';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_RACK] = 'to_rack';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_BIN] = 'to_bin';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.QUANTITY] = 'quantity';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.UOM] = 'uom';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.MOVEMENT_TYPE] = 'movement_type';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.REMARKS] = 'remarks';
        hdr[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.CREATED_BY] = 'created_by';
        archiveSheet.getRange(1, 1, 1, CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.CREATED_BY + 1).setValues([hdr]);
      }

      const data = _getSheetValuesCached(moveSheet.getName());
      if (!data || data.length <= 1) return { success: true, archived: 0, keptDays: keepDays };

      const rowsToArchive = [];
      const deleteRows = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tsRaw = row[CONFIG.MOVEMENT_COLS.TIMESTAMP];
        const ts = (tsRaw instanceof Date) ? tsRaw : new Date(tsRaw);
        if (!(ts instanceof Date) || isNaN(ts.getTime())) continue;
        if (ts >= cutoff) continue;

        const out = [];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.MOVEMENT_ID] = row[CONFIG.MOVEMENT_COLS.MOVEMENT_ID];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.MOVEMENT_DATE] = ts;
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.ITEM_ID] = row[CONFIG.MOVEMENT_COLS.ITEM_ID];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.BATCH_ID] = row[CONFIG.MOVEMENT_COLS.BATCH_ID];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_SITE] = '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_LOCATION] = '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_RACK] = '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.FROM_BIN] = row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_SITE] = '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_LOCATION] = '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_RACK] = '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.TO_BIN] = row[CONFIG.MOVEMENT_COLS.TO_BIN_ID];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.QUANTITY] = row[CONFIG.MOVEMENT_COLS.QUANTITY];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.UOM] =
          (CONFIG.MOVEMENT_COLS.UOM !== undefined) ? row[CONFIG.MOVEMENT_COLS.UOM] : '';
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.MOVEMENT_TYPE] = row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.REMARKS] = row[CONFIG.MOVEMENT_COLS.REMARKS];
        out[CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.CREATED_BY] = row[CONFIG.MOVEMENT_COLS.USER_EMAIL];
        rowsToArchive.push(out);
        deleteRows.push(i + 1);
      }

      if (rowsToArchive.length === 0) return { success: true, archived: 0, keptDays: keepDays };

      const width = CONFIG.INVENTORY_MOVEMENT_ARCHIVE_COLS.CREATED_BY + 1;
      rowsToArchive.forEach(function (r) {
        while (r.length < width) r.push('');
      });
      const startRow = archiveSheet.getLastRow() + 1;
      archiveSheet.getRange(startRow, 1, rowsToArchive.length, width).setValues(rowsToArchive);

      deleteRows.sort(function (a, b) { return b - a; });
      let rangeHigh = deleteRows[0];
      let rangeLow = deleteRows[0];
      for (let i = 1; i < deleteRows.length; i++) {
        const r = deleteRows[i];
        if (r === rangeLow - 1) {
          rangeLow = r;
          continue;
        }
        moveSheet.deleteRows(rangeLow, rangeHigh - rangeLow + 1);
        rangeHigh = r;
        rangeLow = r;
      }
      moveSheet.deleteRows(rangeLow, rangeHigh - rangeLow + 1);

      return { success: true, archived: rowsToArchive.length, keptDays: keepDays };
    });
  });
}

function setupMovementLogArchival() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);
    ensureMovementArchiveTrigger();
    return { success: true, message: 'Monthly movement log archival trigger configured.' };
  });
}

/**
 * One-time setup for the Packing_Version_Master sheet.
 * Creates the sheet and adds headers if they don't exist.
 */
function setupPackingVersionMaster() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG.SHEETS.PACKING_VERSION_MASTER;
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = ['mapping_id', 'item_code', 'system_version', 'bin_id', 'is_latest', 'reason', 'updated_by', 'updated_at'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
      return { success: true, message: 'Sheet "' + sheetName + '" created with headers.' };
    } else {
      return { success: true, message: 'Sheet "' + sheetName + '" already exists.' };
    }
  });
}

/**
 * Get dashboard KPIs for Manager overview
 * @returns {Object} Dashboard statistics
 */
function getDashboardKPIs() {
  return protect(function (user) {
    const stats = _getInventorySummaryStats(false);
    const totalStock = Number(stats.totalQty || 0);
    const uniqueItems = Number(stats.uniqueSkus || 0);
    const pendingQA = Number(stats.pendingQA || 0);

    return {
      totalStock: totalStock,
      uniqueItems: uniqueItems,
      pendingQA: pendingQA,
      uniqueSKUs: uniqueItems,
      estValue: 0 // Placeholder - can calculate if price data available
    };
  });
}

function _movementTypeMatchesDashboardTab(movementType, tabType) {
  const mt = String(movementType || '').trim().toUpperCase();
  const tab = String(tabType || '').trim().toUpperCase();
  if (!tab) return true;
  if (tab === 'INWARD') return mt === 'INWARD' || mt === 'INWARD_RECEIPT' || mt === 'PROD_RETURN' || mt === 'RETURN_PROD';
  if (tab === 'INTER_UNIT') return mt === 'INTER_UNIT' || mt === 'TRANSFER' || mt === 'TRANSFER_QUARANTINE';
  if (tab === 'DISPATCH') return mt === 'DISPATCH' || mt === 'CONSUMPTION' || mt === 'PROD_CONSUMPTION';
  return mt === tab;
}

function getInventorySummary() {
  return protect(function () {
    const approved = _aggregateInventorySnapshotByItem({ allowedStatuses: ['APPROVED', 'OVERRIDDEN'] });
    const itemMaps = _getItemMasterMaps();
    const out = {};

    Object.keys(approved).forEach(function (code) {
      const norm = String(code || '').trim().toUpperCase();
      const group = String(itemMaps.codeToGroup[norm] || 'OTHERS').trim() || 'OTHERS';
      if (!out[group]) out[group] = { totalQty: 0, items: 0 };
      out[group].totalQty += Number((approved[code] || {}).totalQty || 0);
      out[group].items += 1;
    });

    return out;
  });
}

function getInventoryByCategory(groupCode) {
  return protect(function () {
    const target = String(groupCode || '').trim().toUpperCase();
    if (!target) return [];

    const itemMaps = _getItemMasterMaps();
    const items = getWarehouseItemDashboardData() || [];
    const rows = [];

    items.forEach(function (it) {
      const code = String(it.itemCode || '').trim();
      const norm = code.toUpperCase();
      const group = String(itemMaps.codeToGroup[norm] || '').trim().toUpperCase();
      if (group !== target) return;

      const seenBins = {};
      (it.versions || []).forEach(function (v) {
        (v.bins || []).forEach(function (b) {
          const id = String(b.binId || '').trim();
          if (id) seenBins[id] = true;
        });
      });

      rows.push({
        itemCode: code,
        itemName: String(itemMaps.codeToName[norm] || code),
        totalQty: Number(it.totalQty || 0),
        uom: String(it.uomCode || itemMaps.codeToUom[norm] || 'KG'),
        bins: Object.keys(seenBins).length
      });
    });

    rows.sort(function (a, b) { return b.totalQty - a.totalQty; });
    return rows;
  });
}

function getRecentMovements(type) {
  return protect(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
    const data = _getSheetValuesCached(sheet.getName());
    const itemMaps = _getItemMasterMaps();
    let binMeta = {};
    try { binMeta = _buildBinMetaMap(ss); } catch (e) { binMeta = {}; }
    const rows = [];
    const limit = 50;

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const movementType = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').trim().toUpperCase();
      if (!_movementTypeMatchesDashboardTab(movementType, type)) continue;

      const itemId = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
      const itemCode = itemMaps.idToCode[itemId.toUpperCase()] || itemId;
      const codeNorm = String(itemCode || '').trim().toUpperCase();
      const fromBinId = String(row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID] || '').trim();
      const toBinId = String(row[CONFIG.MOVEMENT_COLS.TO_BIN_ID] || '').trim();
      const qty = Number(row[CONFIG.MOVEMENT_COLS.QUANTITY] || 0);
      const uom = String((CONFIG.MOVEMENT_COLS.UOM !== undefined ? row[CONFIG.MOVEMENT_COLS.UOM] : '') || itemMaps.codeToUom[codeNorm] || 'KG').trim() || 'KG';

      rows.push({
        itemCode: itemCode,
        itemName: String(itemMaps.codeToName[codeNorm] || itemCode),
        fromBin: (binMeta[fromBinId] && binMeta[fromBinId].binCode) ? binMeta[fromBinId].binCode : fromBinId,
        toBin: (binMeta[toBinId] && binMeta[toBinId].binCode) ? binMeta[toBinId].binCode : toBinId,
        quantity: qty,
        uom: uom,
        dateString: row[CONFIG.MOVEMENT_COLS.TIMESTAMP] || new Date()
      });

      if (rows.length >= limit) break;
    }
    return rows;
  });
}

function getItemMovementHistory(itemCode, limit, offset) {
  return protect(function () {
    const code = String(itemCode || '').trim().toUpperCase();
    if (!code) return { movements: [], hasMore: false };

    const take = Math.max(1, Number(limit) || 20);
    const skip = Math.max(0, Number(offset) || 0);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
    const data = _getSheetValuesCached(sheet.getName());
    const itemMaps = _getItemMasterMaps();

    const movements = [];
    let seen = 0;
    let hasMore = false;

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const itemId = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
      const resolvedCode = String(itemMaps.idToCode[itemId.toUpperCase()] || itemId).trim().toUpperCase();
      if (resolvedCode !== code) continue;

      if (seen < skip) {
        seen++;
        continue;
      }

      if (movements.length >= take) {
        hasMore = true;
        break;
      }

      const qty = Number(row[CONFIG.MOVEMENT_COLS.QUANTITY] || 0);
      const uom = String((CONFIG.MOVEMENT_COLS.UOM !== undefined ? row[CONFIG.MOVEMENT_COLS.UOM] : '') || itemMaps.codeToUom[code] || 'KG').trim() || 'KG';
      movements.push({
        type: String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || ''),
        quantity: qty,
        uom: uom,
        date: row[CONFIG.MOVEMENT_COLS.TIMESTAMP] || new Date(),
        remarks: String(row[CONFIG.MOVEMENT_COLS.REMARKS] || '')
      });
    }

    return { movements: movements, hasMore: hasMore };
  });
}

/**
 * Returns the last 5 unique TRANSFER routes made by the current user.
 * Used by TransferForm "Recent" strip — read-only, no lock needed.
 * @returns {Array<{itemCode, batchId, fromBinId, fromBinCode, toBinId, toBinCode}>}
 */
function getMyRecentTransfers() {
  return protect(function (user) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MOVEMENT);
    if (!sheet) return [];

    const data = _getSheetValuesCached(sheet.getName());
    const maps = _getItemMasterMaps();
    const seen = {};
    const result = [];

    for (let i = data.length - 1; i >= 1 && result.length < 5; i--) {
      const row = data[i];
      const type = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').toUpperCase();
      if (type !== 'TRANSFER' && type !== 'TRANSFER_QUARANTINE') continue;

      const rowEmail = String(row[CONFIG.MOVEMENT_COLS.USER_EMAIL] || '').trim().toLowerCase();
      if (rowEmail !== String(user.email || '').trim().toLowerCase()) continue;

      const itemId = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
      const batchId = String(row[CONFIG.MOVEMENT_COLS.BATCH_ID] || '').trim();
      const fromBin = String(row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID] || '').trim();
      const toBin = String(row[CONFIG.MOVEMENT_COLS.TO_BIN_ID] || '').trim();
      const itemCode = maps.idToCode[itemId.toUpperCase()] || itemId;

      const key = itemCode + '|' + batchId + '|' + fromBin + '|' + toBin;
      if (seen[key]) continue;
      seen[key] = true;

      const fromMeta = _getBinMetaById(fromBin);
      const toMeta = _getBinMetaById(toBin);

      result.push({
        itemCode: itemCode,
        batchId: batchId,
        fromBinId: fromBin,
        fromBinCode: String((fromMeta && fromMeta.binCode) || fromBin),
        toBinId: toBin,
        toBinCode: String((toMeta && toMeta.binCode) || toBin)
      });
    }
    return result;
  });
}

function searchInventoryItems(query) {
  return protect(function (user) {
    const allItems = getItems();
    if (!query) return [];
    const q = query.toLowerCase();
    return allItems.filter(item =>
      item.code.toLowerCase().includes(q) || item.name.toLowerCase().includes(q)
    ).slice(0, 20);
  });
}

function getItemDetails(itemCode) {
  return protect(function (user) {
    const visRows = getInventoryReadView({ itemCode: itemCode });
    const stock = (visRows || []).map(r => ({
      batch: r.batchNumber || r.batchId,
      batchId: r.batchId,
      gin: r.ginNo,
      version: r.version,
      binId: r.binId,
      qty: Number(r.quantity) || 0,
      quality: r.qualityStatus,
      date: r.qualityDate
    }));
    return { itemCode: itemCode, stock: stock };
  });
}

/**
 * Clear dashboard cache to force refresh after transactions
 */
function clearDashboardCache() {
  try {
    const sc = CacheService.getScriptCache();
    sc.removeAll([
      AI_CONFIG.CACHE_KEY,
      'LOCATIONS_MASTER_V1',
      'BINS_FAST_V3',
      'BINS_V6',
      'ACTIVE_PRODUCTION_ORDERS_V1',
      'PACKING_MATERIAL_DASHBOARD_V1',
      'PACKING_MATERIAL_DASHBOARD_V2'
    ]);
    // NOTE: User auth and item caches expire naturally in 300s.
    // We cannot enumerate all keys to delete them (GAS limitation).
    // For immediate effect on a specific user: call invalidateUserAuthCache(email)

    return { success: true };
  } catch (e) {
    Logger.log('Cache clear error: ' + e.message);
    return { success: false };
  }
}

// =====================================================
// QA APPROVAL WORKFLOW
// =====================================================

/**
 * QA Approval/Rejection
 * Allows Manager to update quality status of inventory
 * @param {Object} payload - {inventoryId, newStatus, remarks}
 * @returns {Object} {success: true}
 */
function updateQualityStatus(payload) {
  return withScriptLock(function () {
    protect(() => assertQualityManagerAccess());

    if (!payload.inventoryId) throw new Error('Inventory ID required');
    if (!payload.newStatus) throw new Error('Status required');

    // Validate status value
    const validStatuses = ['APPROVED', 'HOLD', 'PENDING', 'OVERRIDDEN', 'CLOSED', 'REJECTED'];
    let newStatus = String(payload.newStatus || '').trim().toUpperCase();
    if (!validStatuses.includes(newStatus)) {
      throw new Error(`Invalid status. Must be: ${validStatuses.join(', ')}`);
    }
    if (newStatus === 'OVERRIDDEN') {
      throw new Error('OVERRIDDEN is set only via QA override');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());
    const headerMap = _getSheetHeaderMap(invSheet);
    const holdReasonIdx = headerMap['HOLD_REASON'] ?? headerMap['QA_HOLD_REASON'];
    const holdByIdx = headerMap['HOLD_BY'] ?? headerMap['QA_HOLD_BY'];
    const holdAtIdx = headerMap['HOLD_AT'] ?? headerMap['QA_HOLD_AT'];
    const approvedByIdx = headerMap['QA_APPROVED_BY'] ?? headerMap['APPROVED_BY'];
    const approvedAtIdx = headerMap['QA_APPROVED_AT'] ?? headerMap['APPROVED_AT'];

    let found = false;
    const targetInventoryId = String(payload.inventoryId);
    const hintedRowIndex = Number(payload.rowIndex);
    const candidateIndexes = [];
    if (isFinite(hintedRowIndex) && hintedRowIndex > 1 && hintedRowIndex <= data.length) {
      candidateIndexes.push(hintedRowIndex - 1);
    }
    for (let idx = 1; idx < data.length; idx++) {
      if (candidateIndexes.indexOf(idx) === -1) candidateIndexes.push(idx);
    }

    for (let ci = 0; ci < candidateIndexes.length; ci++) {
      const i = candidateIndexes[ci];
      const rowInvId = String(data[i][CONFIG.INVENTORY_COLS.INVENTORY_ID]);

      if (rowInvId === targetInventoryId) {
        const currentStatus = _normalizeQaStatus(data[i][CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');
        if (currentStatus === 'CLOSED') {
          throw new Error('CLOSED rows are immutable');
        }
        if (currentStatus === 'OVERRIDDEN' && newStatus !== 'CLOSED') {
          throw new Error('OVERRIDDEN stock can only transition to CLOSED');
        }
        if (newStatus === 'CLOSED' && currentStatus !== 'OVERRIDDEN') {
          throw new Error('Only OVERRIDDEN stock can be CLOSED');
        }

        const qd = payload.qualityDate ? new Date(payload.qualityDate) : new Date();
        const rowData = data[i].slice();
        const changedCells = [];
        const userEmail = Session.getActiveUser().getEmail();
        const currentQty = Number(data[i][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        const hasApprovedQtyInput = (
          payload &&
          typeof payload.approvedQty !== 'undefined' &&
          payload.approvedQty !== null &&
          String(payload.approvedQty).trim() !== ''
        );
        const rawApprovedQty = hasApprovedQtyInput ? Number(payload.approvedQty) : NaN;
        const hasApprovedQty = hasApprovedQtyInput && isFinite(rawApprovedQty) && rawApprovedQty > 0;
        const partialApprove = (
          newStatus === 'APPROVED' &&
          hasApprovedQty &&
          currentQty > 0 &&
          rawApprovedQty < (currentQty - 0.000001)
        );

        if (newStatus === 'APPROVED' && hasApprovedQtyInput && !hasApprovedQty) {
          throw new Error('Approved quantity must be greater than 0.');
        }

        if (newStatus === 'APPROVED' && hasApprovedQty && rawApprovedQty > currentQty + 0.000001) {
          throw new Error('Approved quantity cannot exceed available quantity.');
        }

        rowData[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = newStatus;
        rowData[CONFIG.INVENTORY_COLS.QUALITY_DATE] = qd;
        changedCells.push({ col: CONFIG.INVENTORY_COLS.QUALITY_STATUS, val: newStatus });
        changedCells.push({ col: CONFIG.INVENTORY_COLS.QUALITY_DATE, val: qd });

        if (partialApprove) {
          rowData[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = rawApprovedQty;
          changedCells.push({ col: CONFIG.INVENTORY_COLS.TOTAL_QUANTITY, val: rawApprovedQty });
        }

        if (newStatus === 'HOLD' || newStatus === 'REJECTED') {
          if (typeof holdReasonIdx === 'number') {
            const reason = String(payload.remarks || '');
            rowData[holdReasonIdx] = reason;
            changedCells.push({ col: holdReasonIdx, val: reason });
          }
          if (typeof holdByIdx === 'number') {
            rowData[holdByIdx] = userEmail;
            changedCells.push({ col: holdByIdx, val: userEmail });
          }
          if (typeof holdAtIdx === 'number') {
            const at = new Date();
            rowData[holdAtIdx] = at;
            changedCells.push({ col: holdAtIdx, val: at });
          }
        }
        if (newStatus === 'APPROVED') {
          if (typeof approvedByIdx === 'number') {
            rowData[approvedByIdx] = userEmail;
            changedCells.push({ col: approvedByIdx, val: userEmail });
          }
          if (typeof approvedAtIdx === 'number') {
            const at = new Date();
            rowData[approvedAtIdx] = at;
            changedCells.push({ col: approvedAtIdx, val: at });
          }
        }

        let writeWidth = rowData.length;
        changedCells.forEach(function (c) { writeWidth = Math.max(writeWidth, Number(c.col || 0) + 1); });
        while (rowData.length < writeWidth) rowData.push('');
        invSheet.getRange(i + 1, 1, 1, writeWidth).setValues([rowData.slice(0, writeWidth)]);
        changedCells.forEach(function (c) {
          _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, i + 1, Number(c.col) + 1, c.val);
        });

        const qaEvents = [{
          inventoryId: rowInvId,
          itemCode: String(data[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || ''),
          batchId: String(data[i][CONFIG.INVENTORY_COLS.BATCH_ID] || ''),
          binId: String(data[i][CONFIG.INVENTORY_COLS.BIN_ID] || ''),
          prevStatus: currentStatus,
          newStatus: newStatus,
          action: partialApprove
            ? 'APPROVE_PARTIAL'
            : (newStatus === 'REJECTED'
              ? 'REJECT'
              : (newStatus === 'HOLD' ? 'HOLD' : (newStatus === 'APPROVED' ? 'APPROVE' : 'STATUS_CHANGE'))),
          overrideReason: String(payload.remarks || ''),
          overriddenBy: userEmail,
          overriddenAt: new Date(),
          remarks: String(payload.remarks || ''),
          quantity: partialApprove ? rawApprovedQty : currentQty
        }];

        if (partialApprove) {
          const remainingQty = Math.max(0, currentQty - rawApprovedQty);
          if (remainingQty > 0.000001) {
            const holdRow = data[i].slice();
            const holdInventoryId = 'INV-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
            holdRow[CONFIG.INVENTORY_COLS.INVENTORY_ID] = holdInventoryId;
            holdRow[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = remainingQty;
            holdRow[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = 'HOLD';
            holdRow[CONFIG.INVENTORY_COLS.QUALITY_DATE] = qd;
            if (typeof holdReasonIdx === 'number') holdRow[holdReasonIdx] = String(payload.remarks || 'Partial QA approval - balance moved to HOLD');
            if (typeof holdByIdx === 'number') holdRow[holdByIdx] = userEmail;
            if (typeof holdAtIdx === 'number') holdRow[holdAtIdx] = new Date();

            let holdWidth = Math.max(writeWidth, holdRow.length);
            while (holdRow.length < holdWidth) holdRow.push('');
            invSheet.getRange(invSheet.getLastRow() + 1, 1, 1, holdWidth).setValues([holdRow.slice(0, holdWidth)]);
            _appendSheetCacheRow(CONFIG.SHEETS.INVENTORY, holdRow.slice(0, holdWidth));

            qaEvents.push({
              inventoryId: holdInventoryId,
              itemCode: String(holdRow[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || ''),
              batchId: String(holdRow[CONFIG.INVENTORY_COLS.BATCH_ID] || ''),
              binId: String(holdRow[CONFIG.INVENTORY_COLS.BIN_ID] || ''),
              prevStatus: currentStatus,
              newStatus: 'HOLD',
              action: 'HOLD_PARTIAL_BALANCE',
              overrideReason: String(payload.remarks || ''),
              overriddenBy: userEmail,
              overriddenAt: new Date(),
              remarks: 'Balance qty moved to HOLD after partial approval',
              quantity: remainingQty
            });
          }
        }

        _appendQaEventsBatch(qaEvents);

        found = true;
        break;
      }
    }

    if (!found) throw new Error('Inventory record not found');

    return { success: true, message: `Status updated to ${payload.newStatus}` };
  });
}

/**
 * Get list of inventory items pending QA approval
 * @returns {Array} List of pending inventory records
 */
function getPendingQAItems() {
  return getQaInventoryView();
}

function _getLatestOverrideEventsByInventoryId() {
  const map = {};
  let sheet;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.QA_EVENTS);
  } catch (e) {
    Logger.log('QA_Events missing: ' + e.message);
    return map;
  }
  const data = _getSheetValuesCached(sheet.getName());
  for (let i = 1; i < data.length; i++) {
    const invId = String(data[i][CONFIG.QA_EVENTS_COLS.INVENTORY_ID] || '').trim();
    if (!invId) continue;
    const newStatus = String(data[i][CONFIG.QA_EVENTS_COLS.NEW_STATUS] || '').trim().toUpperCase();
    const action = String(data[i][CONFIG.QA_EVENTS_COLS.ACTION] || '').trim().toUpperCase();
    if (newStatus !== 'OVERRIDDEN' && action !== 'OVERRIDE') continue;
    const tsVal = data[i][CONFIG.QA_EVENTS_COLS.OVERRIDDEN_AT] || data[i][CONFIG.QA_EVENTS_COLS.TIMESTAMP];
    const ts = tsVal instanceof Date ? tsVal.getTime() : Date.parse(tsVal);
    if (!map[invId] || (isFinite(ts) && ts > map[invId].ts)) {
      map[invId] = {
        ts: ts,
        overriddenBy: String(data[i][CONFIG.QA_EVENTS_COLS.OVERRIDDEN_BY] || ''),
        overriddenAt: tsVal || '',
        overrideReason: String(data[i][CONFIG.QA_EVENTS_COLS.OVERRIDE_REASON] || ''),
        overrideAction: String(data[i][CONFIG.QA_EVENTS_COLS.ACTION] || ''),
        overrideMovementId: String(data[i][CONFIG.QA_EVENTS_COLS.MOVEMENT_ID] || '')
      };
    }
  }
  return map;
}

function getQaInventoryView() {
  _beginRequestCache();
  try {
    // Auth guard FIRST
    assertQualityManagerAccess();

    // Load overrides
    const overrideMap = _getLatestOverrideEventsByInventoryId() || {};

    // Load inventory read view
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());

    // Build visibility rows inline (no nested protect calls)
    const norm = v => String(v || '').trim().toUpperCase();
    const itemIdToCode = {};
    const itemCodeToUom = {};
    const visRows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      let resolvedItemCode = _resolveInventoryItemCode(
        row,
        row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE],
        row[CONFIG.INVENTORY_COLS.ITEM_ID],
        itemIdToCode
      );
      resolvedItemCode = _getCanonicalItemCode(resolvedItemCode, row[CONFIG.INVENTORY_COLS.ITEM_ID]);
      const rowCode = norm(resolvedItemCode);
      if (!rowCode) continue;
      const rowUomRaw = String((CONFIG.INVENTORY_COLS.UOM !== undefined ? row[CONFIG.INVENTORY_COLS.UOM] : '') || '').trim();
      let rowUom = rowUomRaw;
      if (!rowUom) {
        if (!itemCodeToUom[rowCode]) itemCodeToUom[rowCode] = _getItemUomCode(rowCode);
        rowUom = itemCodeToUom[rowCode] || 'KG';
      }

      const batchInfo = _resolveBatchReference(resolvedItemCode, row[CONFIG.INVENTORY_COLS.BATCH_ID]);

      visRows.push({
        rowIndex: i + 1,
        rowData: row,
        inventoryId: String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || ''),
        itemId: String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || ''),
        itemCode: String(resolvedItemCode || '').trim(),
        batchId: String((batchInfo && batchInfo.batchId) || row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        binId: norm(row[CONFIG.INVENTORY_COLS.BIN_ID]),
        ginNo: String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim(),
        site: String(row[CONFIG.INVENTORY_COLS.SITE] || '').trim(),
        location: String(row[CONFIG.INVENTORY_COLS.LOCATION] || '').trim(),
        quantity: qty,
        uom: rowUom,
        qualityStatus: String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || '').trim(),
        qualityDate: row[CONFIG.INVENTORY_COLS.QUALITY_DATE] || '',
        version: String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim()
      });
    }

    const result = [];
    visRows.forEach(r => {
      let status = _normalizeQaStatus(r.qualityStatus || 'PENDING');
      // Exclude resolved rows from QA worklist.
      if (status === 'APPROVED' || status === 'REJECTED') {
        return;
      }

      const invId = String(r.inventoryId || '');
      const override = overrideMap[invId] || {};

      // Convert dates to ISO strings for JSON serialization
      const qualityDateVal = r.qualityDate;
      const qualityDateStr = qualityDateVal ? (qualityDateVal instanceof Date ? qualityDateVal.toISOString() : String(qualityDateVal)) : '';
      const overriddenAtVal = override.overriddenAt;
      const overriddenAtStr = overriddenAtVal ? (overriddenAtVal instanceof Date ? overriddenAtVal.toISOString() : String(overriddenAtVal)) : '';

      result.push({
        rowIndex: r.rowIndex,
        inventoryId: invId,
        itemId: String(r.itemId || ''),
        itemCode: String(r.itemCode || ''),
        batchId: String(r.batchId || ''),
        batchNumber: _getBatchDisplayNumber(r.itemCode, r.batchId),
        binId: String(r.binId || ''),
        ginNo: String(r.ginNo || ''),
        quantity: Number(r.quantity) || 0,
        uom: String(r.uom || ''),
        qualityStatus: status,
        qualityDate: qualityDateStr,
        inwardDate: qualityDateStr,
        overriddenBy: String(override.overriddenBy || ''),
        overriddenAt: overriddenAtStr,
        overrideReason: String(override.overrideReason || ''),
        overrideAction: String(override.overrideAction || ''),
        overrideMovementId: String(override.overrideMovementId || '')
      });
    });
    return result;

  } catch (error) {
    Logger.log('[QA_BACKEND] EXCEPTION: ' + error.toString());
    Logger.log('[QA_BACKEND] Message: ' + error.message);
    throw error;
  } finally {
    _endRequestCache();
  }
}


/**
 * Inventory read view for form loaders (visibility-only).
 * VISIBILITY ONLY - FIFO & QA intentionally bypassed for UI visibility.
 * NEVER call _readInventoryState here.
 */
function getInventoryReadView(filter) {
  return protect(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());
    const rows = [];

    const f = filter || {};
    const norm = v => String(v || '').trim().toUpperCase();
    const fItem = norm(f.itemCode || '');
    const batchFilterInfo = String(f.batchId || '').trim()
      ? _resolveBatchReference(fItem || f.itemCode || '', f.batchId)
      : null;
    const acceptedBatchKeys = batchFilterInfo ? (batchFilterInfo.acceptedKeys || {}) : null;
    const fBin = norm(f.binId || '');
    const itemIdToCode = {};
    const itemCodeToUom = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      const resolvedItemCode = _resolveInventoryItemCode(
        row,
        row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE],
        row[CONFIG.INVENTORY_COLS.ITEM_ID],
        itemIdToCode
      );
      const rowCode = norm(resolvedItemCode);
      if (!rowCode) continue;
      const rowBatch = norm(row[CONFIG.INVENTORY_COLS.BATCH_ID]);
      const rowBin = norm(row[CONFIG.INVENTORY_COLS.BIN_ID]);

      if (fItem && rowCode !== fItem) continue;
      if (acceptedBatchKeys && !acceptedBatchKeys[rowBatch]) continue;
      if (fBin && rowBin !== fBin) continue;

      const batchInfo = _resolveBatchReference(resolvedItemCode, row[CONFIG.INVENTORY_COLS.BATCH_ID]);

      rows.push({
        rowIndex: i + 1,
        rowData: row,
        inventoryId: String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || ''),
        itemId: String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || ''),
        itemCode: String(resolvedItemCode || '').trim(),
        batchId: String((batchInfo && batchInfo.batchId) || row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        batchNumber: String((batchInfo && batchInfo.batchNumber) || row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        binId: String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim(),
        ginNo: String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim(),
        site: String(row[CONFIG.INVENTORY_COLS.SITE] || '').trim(),
        location: String(row[CONFIG.INVENTORY_COLS.LOCATION] || '').trim(),
        quantity: qty,
        uom: (function () {
          const rowUom = String((CONFIG.INVENTORY_COLS.UOM !== undefined ? row[CONFIG.INVENTORY_COLS.UOM] : '') || '').trim();
          if (rowUom) return rowUom;
          const key = rowCode;
          if (!itemCodeToUom[key]) itemCodeToUom[key] = _getItemUomCode(resolvedItemCode);
          return itemCodeToUom[key] || 'KG';
        })(),
        qualityStatus: String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || '').trim(),
        qualityDate: row[CONFIG.INVENTORY_COLS.QUALITY_DATE] || '',
        version: String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim()
      });
    }
    return rows;
  });
}

/**
 * getTransferSourcesForItem
 * Lean server-side lookup for TransferForm source picker.
 * Returns ONLY the fields needed by the UI — no rowData, no rowIndex.
 * Called on item-code entry instead of scanning client-side inventoryData.
 *
 * @param  {string} itemCode
 * @returns {Array<{batchId, binId, binCode, site, location, quantity, uom, qualityStatus, version}>}
 */
function getTransferSourcesForItem(itemCode) {
  return protect(function () {
    const code = String(itemCode || '').trim().toUpperCase();
    if (!code) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());
    const binMeta = _buildBinMetaMap(ss);
    const itemIdToCode = {};
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      const resolvedCode = _resolveInventoryItemCode(
        row,
        row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE],
        row[CONFIG.INVENTORY_COLS.ITEM_ID],
        itemIdToCode
      );
      if (!resolvedCode || String(resolvedCode).trim().toUpperCase() !== code) continue;

      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      const batchInfo = _resolveBatchReference(resolvedCode, row[CONFIG.INVENTORY_COLS.BATCH_ID]);
      const meta = binMeta[binId] || {};

      result.push({
        batchId: String((batchInfo && batchInfo.batchId) || row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        batchNumber: String((batchInfo && batchInfo.batchNumber) || row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        binId: binId,
        binCode: String(meta.binCode || binId),
        site: String(meta.site || row[CONFIG.INVENTORY_COLS.SITE] || '').trim(),
        location: String(meta.location || row[CONFIG.INVENTORY_COLS.LOCATION] || '').trim(),
        quantity: qty,
        uom: String((CONFIG.INVENTORY_COLS.UOM !== undefined
          ? row[CONFIG.INVENTORY_COLS.UOM] : '') || '').trim() || _getItemUomCode(resolvedCode),
        qualityStatus: String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').trim(),
        version: String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim()
      });
    }

    // Sort: APPROVED first, then qty desc (mirrors client sort)
    result.sort(function (a, b) {
      var okA = (a.qualityStatus === 'APPROVED' || a.qualityStatus === 'OVERRIDDEN') ? 1 : 0;
      var okB = (b.qualityStatus === 'APPROVED' || b.qualityStatus === 'OVERRIDDEN') ? 1 : 0;
      if (okB !== okA) return okB - okA;
      return b.quantity - a.quantity;
    });

    return result;
  });
}

function getFifoAvailability(itemCode, batchId, binId, requiredQty, operationType) {
  return protect(function () {
    requireRole(SECURITY.ROLES.WORKER);
    _assertItemCodeBatch(itemCode, batchId, 'FIFO availability');
    if (!binId) {
      return {
        totalQty: 0,
        fifoEligibleQty: 0,
        approvedQty: 0,
        pendingQty: 0,
        rejectedQty: 0,
        holdQty: 0,
        blockedByPending: false,
        blockedByHold: false,
        blockedByRejected: false,
        fulfillable: false,
        failureReason: 'Bin required'
      };
    }

    // VISIBILITY ONLY - FIFO & QA intentionally bypassed for row loading.
    const visRows = getInventoryReadView({ itemCode: itemCode, batchId: batchId, binId: binId });
    const rows = (visRows || []).map(r => ({
      rowIndex: r.rowIndex,
      currentQty: Number(r.quantity) || 0,
      rowData: r.rowData
    }));
    if (rows.length === 0) {
      _logInventoryVisibility('FIFO availability', itemCode, batchId, binId, [], [], 'No matching inventory rows');
      return {
        totalQty: 0,
        fifoEligibleQty: 0,
        approvedQty: 0,
        pendingQty: 0,
        rejectedQty: 0,
        holdQty: 0,
        blockedByPending: false,
        blockedByHold: false,
        blockedByRejected: false,
        fulfillable: false,
        failureReason: 'No stock exists'
      };
    }
    const policy = getOperationPolicy(operationType);
    const evalRes = _evaluateFifoEligibility(itemCode, batchId, rows, requiredQty, policy);

    return {
      totalQty: evalRes.totalQty,
      fifoEligibleQty: evalRes.fifoEligibleQty,
      approvedQty: evalRes.approvedQty,
      pendingQty: evalRes.pendingQty,
      rejectedQty: evalRes.rejectedQty,
      holdQty: evalRes.holdQty,
      blockedByPending: evalRes.blockedByPending,
      blockedByHold: evalRes.blockedByHold,
      blockedByRejected: evalRes.blockedByRejected,
      fulfillable: evalRes.fulfillable,
      failureReason: evalRes.failureReason
    };
  });
}

function getQaSummary(itemCode, batchId, binId) {
  protect(() => requireRole(SECURITY.ROLES.WORKER));
  _assertItemCodeBatch(itemCode, batchId, 'QA summary');
  if (!binId) {
    return { total: 0, approved: 0, pending: 0, rejected: 0, hold: 0, version: '', versionConflict: false };
  }

  // VISIBILITY ONLY - FIFO & QA intentionally bypassed for row loading.
  const visRows = getInventoryReadView({ itemCode: itemCode, batchId: batchId, binId: binId });
  const rows = (visRows || []);
  if (!rows || rows.length === 0) {
    _logInventoryVisibility('QA summary', itemCode, batchId, binId, [], [], 'No matching inventory rows');
  }
  let total = 0;
  let approved = 0;
  let pending = 0;
  let rejected = 0;
  let overridden = 0;
  let closed = 0;
  let hold = 0;

  rows.forEach(r => {
    const qty = Number(r.quantity) || 0;
    if (qty <= 0) return;
    total += qty;
    let status = _normalizeQaStatus(r.qualityStatus || 'PENDING');
    if (status === 'APPROVED') approved += qty;
    else if (status === 'HOLD') hold += qty;
    else if (status === 'REJECTED') rejected += qty;
    else if (status === 'OVERRIDDEN') overridden += qty;
    else if (status === 'CLOSED') closed += qty;
    else pending += qty;
  });

  const versionKeys = [];
  rows.forEach(r => {
    const v = String(r.version || '').trim();
    if (v && versionKeys.indexOf(v.toUpperCase()) === -1) versionKeys.push(v.toUpperCase());
  });
  const versionConflict = versionKeys.length > 1;
  const version = versionConflict ? '' : (versionKeys[0] || '');

  return {
    total: total,
    approved: approved,
    pending: pending + hold,
    rejected: rejected,
    overridden: overridden,
    closed: closed,
    hold: hold,
    version: version,
    versionConflict: versionConflict
  };
}

/**
 * Audit / invariant checks (manual)
 * Validates:
 * - Inventory totals = sum of movements
 * - Ledger issued = sum of consumption
 * - Transfer OUT = Transfer IN
 */
/*function runAuditChecks() {
  return protect(function (user) {
    requireRole(SECURITY.ROLES.MANAGER);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    //const movSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
    const ledSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);

    const invData = _getSheetValuesCached(invSheet.getName());
    //const movData = _getSheetValuesCached(movSheet.getName());
    const ledData = _getSheetValuesCached(ledSheet.getName());

    let inventoryTotal = 0;
    for (let i = 1; i < invData.length; i++) {
      inventoryTotal += Number(invData[i][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    }

    let movementNet = 0;
    let transferOut = 0;
    let transferIn = 0;
    let consumptionTotal = 0;

    for (let i = 1; i < movData.length; i++) {
      const type = String(movData[i][CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '');
      const qty = Number(movData[i][CONFIG.MOVEMENT_COLS.QUANTITY]) || 0;
      if (!qty) continue;

      if (type === CONFIG.MOVEMENT_TYPES.INWARD) movementNet += qty;
      if (type === CONFIG.MOVEMENT_TYPES.RETURN_PROD) movementNet += qty;
      if (type === CONFIG.MOVEMENT_TYPES.DISPATCH) movementNet -= qty;
      if (type === CONFIG.MOVEMENT_TYPES.CONSUMPTION) {
        movementNet -= qty;
        consumptionTotal += qty;
      }
      if (type === CONFIG.MOVEMENT_TYPES.TRANSFER || type === CONFIG.MOVEMENT_TYPES.TRANSFER_QUARANTINE) {
        transferOut += qty;
        transferIn += qty;
      }
    }

    let ledgerIssuedTotal = 0;
    for (let i = 1; i < ledData.length; i++) {
      ledgerIssuedTotal += Number(ledData[i][CONFIG.PRODUCTION_LEDGER_COLS.QTY_ISSUED]) || 0;
    }

    const tol = 0.0001;
    const inventoryMatches = Math.abs(inventoryTotal - movementNet) <= tol;
    const ledgerMatches = Math.abs(ledgerIssuedTotal - consumptionTotal) <= tol;
    const transferMatches = Math.abs(transferOut - transferIn) <= tol;

    return {
      inventoryTotal: inventoryTotal,
      movementNet: movementNet,
      inventoryMatches: inventoryMatches,
      ledgerIssuedTotal: ledgerIssuedTotal,
      consumptionTotal: consumptionTotal,
      ledgerMatches: ledgerMatches,
      transferOut: transferOut,
      transferIn: transferIn,
      transferMatches: transferMatches,
      checkedAt: new Date()
    };
  });
}*/

function processTransferRequest(form) {
  // Wrapper for submitTransferV2 to match client side name
  const transferForm = {
    items: [{
      itemId: form.itemId,
      itemCode: form.itemCode,
      batchId: form.batchId,
      binId: form.fromBinId || form.fromBinCode, // client sends code or ID?
      quantity: form.quantity
    }],
    toSite: form.toSite,
    toBinId: form.toBinCode || form.toBinId, // CRITICAL: Destination bin for 2-way transfer
    toLocation: form.toLocation,
    remarks: form.remarks
  };
  return submitTransferV2(transferForm);
}

function _buildInwardInventoryKey(itemCode, batchId, binId, version, ginNo) {
  return [
    String(itemCode || '').trim().toUpperCase(),
    String(batchId || '').trim().toUpperCase(),
    String(binId || '').trim().toUpperCase(),
    String(version || '').trim().toUpperCase(),
    String(ginNo || '').trim().toUpperCase()
  ].join('||');
}


// =====================================================
// 7. PUBLIC TRANSACTION ENDPOINTS
// =====================================================

function submitInwardV2(form) {
  return withScriptLock(function () {
    const inwardUser = protect(() => requireOperationalUser());
    if (!form.items || !Array.isArray(form.items)) throw new Error("Invalid items");
    const ginNo = String(form.ginNo || form.billNo || '').trim();
    if (!ginNo) throw new Error("Bill No. required");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    // Pre-check movement sheet existence
    _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);

    const expandedItems = [];
    form.items.forEach(function (item, idx) {
      const splits = Array.isArray(item && item.binSplits) ? item.binSplits : [];
      if (!splits.length) {
        expandedItems.push(item);
        return;
      }
      const baseCode = String((item && item.itemCode) || '').trim();
      if (!baseCode) throw new Error('Item ' + (idx + 1) + ': Item Code required for split bins');
      const expectedKg = _convertToKg(baseCode, item.quantity, item.uom);
      let splitKg = 0;

      splits.forEach(function (sp, sIdx) {
        const spQty = Number(sp && sp.quantity);
        if (!isFinite(spQty) || spQty <= 0) {
          throw new Error('Item ' + (idx + 1) + ': Split ' + (sIdx + 1) + ' quantity must be > 0');
        }
        const binId = String((sp && sp.binId) || '').trim();
        if (!binId) {
          throw new Error('Item ' + (idx + 1) + ': Split ' + (sIdx + 1) + ' bin required');
        }
        splitKg += _convertToKg(baseCode, spQty, sp.uom || item.uom);
        expandedItems.push(Object.assign({}, item, {
          quantity: spQty,
          uom: sp.uom || item.uom,
          binId: binId,
          site: (sp && sp.site) || item.site,
          location: (sp && sp.location) || item.location
        }));
      });

      if (isFinite(expectedKg) && expectedKg > 0 && Math.abs(splitKg - expectedKg) > 0.001) {
        throw new Error('Item ' + (idx + 1) + ': split-bin quantity total must match entered quantity');
      }
    });

    const binDelta = {};
    const batchRefByItem = {};
    const batchByItem = {};
    const batchMetaCache = {};
    const itemIdCache = {};
    const movementLogs = [];
    const packingLabelUpdates = [];
    const inventoryUpdatesByRow = {};
    const inventoryRowsToCreateByKey = {};
    const inwardDateRows = {};
    const targetItemCodes = {};
    const targetItemIds = {};
    expandedItems.forEach(function (item) {
      const code = _getCanonicalItemCode(String((item && item.itemCode) || '').trim());
      if (!code) return;
      targetItemCodes[code.toUpperCase()] = true;
      try {
        const id = _getValidatedItemId(code);
        if (id) {
          targetItemIds[String(id).trim().toUpperCase()] = true;
          itemIdCache[code.toUpperCase()] = id;
        }
      } catch (e) {
        // Main validation below will throw a clear item error if needed.
      }
    });

    expandedItems.forEach(function (item) {
      const itemCodeRaw = String(item.itemCode || '').trim();
      const itemCode = _getCanonicalItemCode(itemCodeRaw);
      if (!itemCode) throw new Error('Inward requires Item Code');
      const qtyKg = _convertToKg(itemCode, item.quantity, item.uom);
      if (!isFinite(qtyKg) || qtyKg <= 0) {
        throw new Error('Inward quantity must be greater than 0 for ' + itemCode);
      }
      const batchNumber = String(item.batchNumber || item.batchId || '').trim();
      if (!batchNumber) {
        throw new Error('Batch No required for inward receipt: ' + itemCode);
      }
      const qualityStatus = _normalizeQaStatus(String(item.qualityStatus || 'PENDING'));
      const qualityDate = item.qualityDate ? new Date(item.qualityDate) : '';
      const version = String(item.version || 'V1').trim();
      const lotNo = String(item.lotNo || '').trim();
      const expiryDate = _normalizeInwardExpiryDate(item.expiryDate);

      if (!/^V\d+$/i.test(version)) {
        throw new Error('Version must be V1, V2, etc.');
      }

      if (batchRefByItem[itemCode] && batchRefByItem[itemCode].toUpperCase() !== batchNumber.toUpperCase()) {
        throw new Error(`Inward batch_id mismatch for ${itemCode}`);
      }
      batchRefByItem[itemCode] = batchNumber;
      const batchCacheKey = itemCode.toUpperCase() + '||' + batchNumber.toUpperCase();
      let batchMeta = batchMetaCache[batchCacheKey];
      if (!batchMeta) {
        batchMeta = _upsertBatchMasterFromInward(itemCode, batchNumber, {
          rawExpiryDate: item.expiryDate,
          expiryDate: expiryDate,
          lotNo: lotNo,
          ginNo: ginNo,
          version: version
        });
        batchMetaCache[batchCacheKey] = batchMeta;
      }
      const resolvedBatchId = String((batchMeta && batchMeta.batchId) || '').trim();
      if (!resolvedBatchId) {
        throw new Error('Failed to resolve Batch_Master ID for ' + itemCode + ' / ' + batchNumber);
      }
      batchByItem[itemCode] = resolvedBatchId;
      const ids = _assertItemCodeBatch(itemCode, resolvedBatchId, 'Inward');
      const itemCodeKey = ids.itemCode.toUpperCase();
      const lookup = itemIdCache[itemCodeKey] || (itemIdCache[itemCodeKey] = _getValidatedItemId(ids.itemCode));

      _assertBinCapacity(item.binId, item.quantity, binDelta, itemCode, item.uom);
      const inwardKey = _buildInwardInventoryKey(ids.itemCode, ids.batchId, item.binId, version, ginNo);
      let existing = inventoryRowsToCreateByKey[inwardKey] || null;

      if (String(form.inwardMode || '').trim().toLowerCase() === 'pm') {
        const labelType = String(item.pmLabelType || '').trim().toUpperCase();
        const labelName = String(item.pmLabelName || '').trim();
        if (labelType === 'LATEST' || labelType === 'OBSOLETE') {
          packingLabelUpdates.push({
            itemCode: ids.itemCode,
            version: version,
            binId: item.binId,
            labelType: labelType,
            latestName: labelName
          });
        }
      }

      movementLogs.push({
        type: CONFIG.MOVEMENT_TYPES.INWARD,
        itemId: lookup, batchId: ids.batchId,
        version: version, ginNo: ginNo,
        toBinId: item.binId, quantity: qtyKg,
        qualityStatus: qualityStatus, remarks: form.remarks
      });

      if (existing && existing.rowIndex) {
        existing.currentQty = Number(existing.currentQty || 0) + qtyKg;
        inventoryUpdatesByRow[existing.rowIndex] = {
          rowIndex: existing.rowIndex,
          newQty: existing.currentQty
        };
        if (typeof CONFIG.INVENTORY_COLS.INWARD_DATE === 'number' &&
          !existing.rowData[CONFIG.INVENTORY_COLS.INWARD_DATE]) {
          inwardDateRows[existing.rowIndex] = true;
        }
      } else if (existing && typeof existing.quantity !== 'undefined') {
        existing.quantity = Number(existing.quantity || 0) + qtyKg;
      } else {
        inventoryRowsToCreateByKey[inwardKey] = {
          itemId: lookup,
          itemCode: ids.itemCode,
          batchId: ids.batchId, ginNo: ginNo,
          version: version, qualityStatus: qualityStatus,
          qualityDate: qualityDate,
          binId: item.binId, site: item.site, location: item.location,
          quantity: qtyKg,
          uom: String(item.uom || '').trim().toUpperCase(),
          lotNo: lotNo,
          expiryDate: expiryDate
        };
      }
    });

    const updateChanges = Object.keys(inventoryUpdatesByRow).map(function (rowIndex) {
      return inventoryUpdatesByRow[rowIndex];
    });
    if (updateChanges.length > 0) {
      _persistInventoryStatesBatch(updateChanges, sheet);
    }

    const createRows = Object.keys(inventoryRowsToCreateByKey).map(function (key) {
      return inventoryRowsToCreateByKey[key];
    });
    if (createRows.length > 0) {
      _appendInventoryRowsBatch(createRows, sheet);
    }

    const inwardDateRowIndexes = Object.keys(inwardDateRows).map(function (r) { return Number(r); })
      .filter(function (r) { return isFinite(r) && r > 1; })
      .sort(function (a, b) { return a - b; });
    if (inwardDateRowIndexes.length > 0 && typeof CONFIG.INVENTORY_COLS.INWARD_DATE === 'number') {
      const inwardDateCol = CONFIG.INVENTORY_COLS.INWARD_DATE + 1;
      const now = new Date();
      inwardDateRowIndexes.forEach(function (rowIndex) {
        sheet.getRange(rowIndex, inwardDateCol).setValue(now);
        _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, rowIndex, inwardDateCol, now);
      });
    }

    if (movementLogs.length > 0) {
      _appendMovementLogsBatch(movementLogs);
    }

    if (packingLabelUpdates.length > 0) {
      _applyPackingVersionMappingsFromInward(packingLabelUpdates, inwardUser && inwardUser.email);
    }

    try { _updateBinAndRackStatuses(Object.keys(binDelta)); } catch (e) { Logger.log('Bin status update failed: ' + e.message); }
    _clearInwardLookupCaches(Object.keys(targetItemCodes));
    return { success: true, batchIds: batchByItem };
  });
}

function submitDispatchV2(form) {
  return withScriptLock(function () {
    const user = protect(() => requireOperationalUser());
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);

    const expandedItems = [];
    form.items.forEach(function (item, idx) {
      const splits = Array.isArray(item && item.sourceSplits) ? item.sourceSplits : [];
      if (!splits.length) {
        expandedItems.push(item);
        return;
      }
      const expectedKg = _convertToKg(item.itemCode, item.quantity, item.uom);
      let splitKg = 0;
      splits.forEach(function (sp, sIdx) {
        if (!sp.binId) throw new Error('Item ' + (idx + 1) + ': Split ' + (sIdx + 1) + ' bin required');
        const spQty = Number(sp.quantity);
        if (spQty <= 0) throw new Error('Item ' + (idx + 1) + ': Split ' + (sIdx + 1) + ' quantity must be > 0');
        splitKg += _convertToKg(item.itemCode, spQty, item.uom);
        expandedItems.push(Object.assign({}, item, {
          binId: sp.binId,
          quantity: spQty
        }));
      });
      if (Math.abs(splitKg - expectedKg) > 0.001) {
        throw new Error('Item ' + (idx + 1) + ': source split quantities must match row total');
      }
    });

    expandedItems.forEach(function (item) {
      const qaOverride = form && form.qaOverride === true;
      const overrideReason = String(form && form.qaOverrideReason || '').trim();
      if (qaOverride && user.role !== SECURITY.ROLES.MANAGER) {
        throw new Error('QA override requires manager approval');
      }
      if (qaOverride && !overrideReason) {
        throw new Error('QA override reason required');
      }
      const qtyKg = _convertToKg(item.itemCode, item.quantity, item.uom);
      if (qtyKg > 50000) {
        throw new Error('Quantity exceeds single-transaction limit of 50,000 KG - contact manager');
      }

      // CRITICAL FIX: Server-authoritative item validation
      const lookup = _getValidatedItemId(item.itemCode);
      _assertItemCodeBatch(item.itemCode, item.batchId, 'Dispatch');

      // EXECUTION ONLY - FIFO + QA enforced here.
      const allRows = _readInventoryState(item.itemCode, item.batchId, item.binId);
      const rows = allRows.filter(r => r.currentQty > 0);
      if (rows.length === 0) {
        _logInventoryVisibility('Dispatch', item.itemCode, item.batchId, item.binId, allRows, rows, 'No matching inventory rows');
        throw new Error('No stock exists');
      }
      const evalRes = _validateQualityEligibility(item.itemCode, item.batchId, rows, qtyKg, 'DISPATCH', { allowUnapproved: qaOverride });

      const orderedRows = evalRes.orderedRows || [];
      const itemVersion = _getSingleInventoryVersion(orderedRows, `Dispatch ${item.itemCode}`);
      const changes = [];
      const logs = [];
      const qaStatusChanges = [];
      const qaEvents = [];
      let qaEventChange = null;
      const qaMeta = _buildQaOverrideMeta(user, 'DISPATCH', overrideReason, qaOverride);
      const overrideContext = {
        itemCode: item.itemCode,
        batchId: item.batchId,
        binId: item.binId,
        overrideReason: overrideReason,
        userEmail: user.email,
        remarks: _formatRemarksWithOverride(form.remarks, qaMeta)
      };
      let remaining = qtyKg;

      for (const row of orderedRows) {
        if (remaining <= 0) break;
        const pick = Math.min(row.currentQty, remaining);
        if (qaOverride) {
          _applyQaOverrideForRow(row, overrideContext, qaStatusChanges, qaEvents);
        }

        logs.push({
          type: CONFIG.MOVEMENT_TYPES.DISPATCH,
          itemId: lookup,
          batchId: item.batchId,
          version: itemVersion,
          ginNo: row.rowData[CONFIG.INVENTORY_COLS.GIN_NO],
          fromBinId: item.binId, quantity: pick,
          remarks: _formatRemarksWithOverride(form.remarks, qaMeta)
        });

        changes.push({
          rowIndex: row.rowIndex,
          prevQty: row.currentQty,
          prevUpdated: row.rowData[CONFIG.INVENTORY_COLS.LAST_UPDATED],
          newQty: row.currentQty - pick
        });
        remaining -= pick;
      }

      if (remaining > 0) {
        throw new Error('FIFO quantity insufficient');
      }

      try {
        _persistInventoryStatesBatch(changes, sheet);
        _applyQaStatusChanges(sheet, qaStatusChanges);
        qaEventChange = _appendQaEventsBatch(qaEvents);
        _appendMovementLogsBatch(logs);
        try { _updateBinAndRackStatuses([item.binId]); } catch (e) { Logger.log('Bin status update failed: ' + e.message); }
      } catch (e) {
        _rollbackInventoryChanges(changes, sheet);
        _rollbackQaStatusChanges(sheet, qaStatusChanges);
        _rollbackQaEvents(qaEventChange);
        throw e;
      }
      try { _cleanupZeroQtyInventoryRows(changes, sheet); } catch (e) { Logger.log('Zero-qty cleanup failed: ' + e.message); }
    });
    return { success: true };
  });
}

function _findOutwardMovementForReversal(criteria) {
  const opts = criteria || {};
  const targetCode = String(opts.itemCode || '').trim().toUpperCase();
  const targetQty = Number(opts.quantity);
  const targetMovementId = String(opts.movementId || '').trim();
  const targetType = String(opts.movementType || '').trim().toUpperCase();
  const occurrence = Math.max(1, Number(opts.occurrence || 1) || 1);
  if (!targetCode && !targetMovementId) throw new Error('Provide itemCode or movementId for reversal');
  if (!targetMovementId && (!isFinite(targetQty) || targetQty <= 0)) throw new Error('Provide positive quantity for reversal');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
  const data = _getSheetValuesCached(movSheet.getName());
  const itemMaps = _getItemMasterMaps();
  const used = {};
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const movementId = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_ID] || '').trim();
    const remarks = String(row[CONFIG.MOVEMENT_COLS.REMARKS] || '');
    const reverseMatch = remarks.match(/REVERSAL_OF=([^;\s]+)/i);
    if (reverseMatch && reverseMatch[1]) used[String(reverseMatch[1]).trim()] = true;
    if (targetMovementId && movementId !== targetMovementId) continue;

    const movementType = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').trim().toUpperCase();
    if (movementType !== CONFIG.MOVEMENT_TYPES.DISPATCH && movementType !== CONFIG.MOVEMENT_TYPES.CONSUMPTION) continue;
    if (targetType && movementType !== targetType) continue;

    const itemId = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
    const code = String(itemMaps.idToCode[itemId.toUpperCase()] || itemId).trim().toUpperCase();
    const qty = Number(row[CONFIG.MOVEMENT_COLS.QUANTITY]) || 0;
    if (targetCode && code !== targetCode) continue;
    if (isFinite(targetQty) && Math.abs(qty - targetQty) > 0.000001) continue;
    matches.push({ rowIndex: i + 1, row: row, itemCode: code, movementId: movementId, movementType: movementType, quantity: qty });
  }

  const available = matches.filter(function (m) { return !used[m.movementId]; });
  if (available.length < occurrence) {
    throw new Error('No unreversed outward movement found for ' + (targetMovementId || (targetCode + ' qty ' + targetQty)));
  }
  return available[occurrence - 1];
}

function _restoreInventoryForOutwardMovement(match, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const invData = _getSheetValuesCached(invSheet.getName());
  const row = match.row;
  const itemCode = match.itemCode;
  const itemId = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
  const batchId = String(row[CONFIG.MOVEMENT_COLS.BATCH_ID] || '').trim();
  const version = String(row[CONFIG.MOVEMENT_COLS.VERSION] || '').trim();
  const ginNo = String(row[CONFIG.MOVEMENT_COLS.GIN_NO] || '').trim();
  const binId = String(row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID] || '').trim();
  const qty = Number(row[CONFIG.MOVEMENT_COLS.QUANTITY]) || 0;
  const uom = String(row[CONFIG.MOVEMENT_COLS.UOM] || _getItemUomCode(itemCode) || 'KG').trim().toUpperCase();
  if (!itemCode || !batchId || !binId || qty <= 0) throw new Error('Selected movement is missing item/batch/bin/quantity details');

  let targetRowIndex = -1;
  for (let i = 1; i < invData.length; i++) {
    const invRow = invData[i];
    const rowCode = String(invRow[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
    const rowBatch = String(invRow[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
    const rowBin = String(invRow[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
    const rowVersion = String(invRow[CONFIG.INVENTORY_COLS.VERSION] || '').trim();
    const rowGin = String(invRow[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim();
    if (rowCode === itemCode && rowBatch === batchId && rowBin === binId && rowVersion === version && rowGin === ginNo) {
      targetRowIndex = i + 1;
      break;
    }
  }

  if (targetRowIndex > 1) {
    const current = Number(invData[targetRowIndex - 1][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    const restored = current + qty;
    invSheet.getRange(targetRowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1).setValue(restored);
    invSheet.getRange(targetRowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1).setValue(new Date());
    _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, targetRowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1, restored);
    _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, targetRowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1, new Date());
  } else {
    _appendInventoryRowsBatch([{
      itemId: itemId || itemCode,
      itemCode: itemCode,
      batchId: batchId,
      ginNo: ginNo || 'REVERSAL-' + match.movementId,
      version: version || 'V1',
      qualityStatus: 'APPROVED',
      qualityDate: new Date(),
      binId: binId,
      site: '',
      location: '',
      quantity: qty,
      uom: uom,
      lotNo: '',
      expiryDate: _getDefaultInwardExpiryDate()
    }], invSheet);
  }

  _appendMovementLogsBatch([{
    type: CONFIG.MOVEMENT_TYPES.INWARD,
    itemId: itemId || itemCode,
    batchId: batchId,
    version: version,
    ginNo: ginNo || 'REVERSAL-' + match.movementId,
    toBinId: binId,
    quantity: qty,
    uom: uom,
    qualityStatus: 'APPROVED',
    remarks: 'REVERSAL_OF=' + match.movementId + '; ' + String(reason || 'Outward reversed by manager')
  }]);

  try { _updateBinAndRackStatuses([binId]); } catch (e) { Logger.log('Bin status update failed: ' + e.message); }
  return { movementId: match.movementId, itemCode: itemCode, batchId: batchId, binId: binId, quantityRestored: qty, uom: uom };
}

function reverseOutwardMovement(criteria) {
  return withScriptLock(function () {
    protect(function () { return requireRole(SECURITY.ROLES.MANAGER); });
    const match = _findOutwardMovementForReversal(criteria || {});
    const result = _restoreInventoryForOutwardMovement(match, (criteria && criteria.reason) || '');
    return { success: true, reversed: result };
  });
}

function reversePM311FirstOutward147() {
  return reverseOutwardMovement({
    itemCode: 'PM311',
    quantity: 147,
    occurrence: 1,
    reason: 'Reverse mistaken first PM311 outward 147 NOS'
  });
}

function _findInwardMovementForInventoryRow_(invRow) {
  const movSheet = _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.MOVEMENT);
  const data = _getSheetValuesCached(movSheet.getName());
  const itemId = String(invRow[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim();
  const batchId = String(invRow[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
  const version = String(invRow[CONFIG.INVENTORY_COLS.VERSION] || '').trim();
  const ginNo = String(invRow[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim();
  const binId = String(invRow[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
  const qty = Number(invRow[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const movementType = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').trim().toUpperCase();
    if (movementType !== CONFIG.MOVEMENT_TYPES.INWARD) continue;
    if (String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim() !== itemId) continue;
    if (String(row[CONFIG.MOVEMENT_COLS.BATCH_ID] || '').trim() !== batchId) continue;
    if (String(row[CONFIG.MOVEMENT_COLS.VERSION] || '').trim() !== version) continue;
    if (String(row[CONFIG.MOVEMENT_COLS.GIN_NO] || '').trim() !== ginNo) continue;
    if (String(row[CONFIG.MOVEMENT_COLS.TO_BIN_ID] || '').trim() !== binId) continue;
    if (Math.abs((Number(row[CONFIG.MOVEMENT_COLS.QUANTITY]) || 0) - qty) > 0.000001) continue;
    return String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_ID] || '').trim();
  }
  return '';
}

function _hasInwardInventoryReversal_(inventoryId) {
  const data = _getSheetValuesCached(CONFIG.SHEETS.MOVEMENT);
  const token = 'INWARD_REVERSAL_OF_INV=' + String(inventoryId || '').trim();
  for (let i = 1; i < data.length; i++) {
    const remarks = String(data[i][CONFIG.MOVEMENT_COLS.REMARKS] || '');
    if (remarks.indexOf(token) !== -1) return true;
  }
  return false;
}

function _findInventoryRowForInwardReversal_(criteria) {
  const invSheet = _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const inventoryId = String(criteria.inventoryId || '').trim();
  const itemCode = String(criteria.itemCode || '').trim().toUpperCase();
  const batchId = String(criteria.batchId || '').trim();
  const ginNo = String(criteria.ginNo || '').trim();
  const version = String(criteria.version || '').trim();
  const binId = String(criteria.binId || '').trim();
  const quantity = Number(criteria.quantity);
  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowInventoryId = String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || '').trim();
    const rowCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
    const rowBatch = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
    const rowGin = String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim();
    const rowVersion = String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim();
    const rowBin = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
    const rowQty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (inventoryId && rowInventoryId !== inventoryId) continue;
    if (itemCode && rowCode !== itemCode) continue;
    if (batchId && rowBatch !== batchId) continue;
    if (ginNo && rowGin !== ginNo) continue;
    if (version && rowVersion !== version) continue;
    if (binId && rowBin !== binId) continue;
    if (isFinite(quantity) && Math.abs(rowQty - quantity) > 0.000001) continue;
    if (rowQty <= 0) continue;
    matches.push({ rowIndex: i + 1, row: row });
  }

  if (matches.length !== 1) {
    throw new Error('Expected exactly one duplicate inward inventory row for ' + itemCode + ', found ' + matches.length + '. Use inventoryId for exact reversal.');
  }
  return matches[0];
}

function reverseDuplicateInwardRows(criteria) {
  return withScriptLock(function () {
    protect(function () { return requireRole(SECURITY.ROLES.MANAGER); });
    const items = Array.isArray(criteria && criteria.items) ? criteria.items : [];
    if (items.length === 0) throw new Error('Provide at least one inward inventory row to reverse.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const reversed = [];
    const logs = [];
    const touchedBins = {};
    const reason = String(criteria.reason || 'Duplicate inward reversed by manager').trim();

    items.forEach(function (item) {
      const match = _findInventoryRowForInwardReversal_(item || {});
      const row = match.row;
      const inventoryId = String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || '').trim();
      if (_hasInwardInventoryReversal_(inventoryId)) {
        throw new Error('Inventory row already reversed earlier: ' + inventoryId);
      }

      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) throw new Error('Inventory row has no stock to reverse: ' + inventoryId);
      const now = new Date();
      invSheet.getRange(match.rowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1).setValue(0);
      invSheet.getRange(match.rowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1).setValue(now);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, match.rowIndex, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1, 0);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, match.rowIndex, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1, now);

      const movementId = _findInwardMovementForInventoryRow_(row);
      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      touchedBins[binId] = true;
      logs.push({
        type: CONFIG.MOVEMENT_TYPES.INWARD_REVERSAL,
        itemId: String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim(),
        batchId: String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        version: String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim(),
        ginNo: String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim(),
        fromBinId: binId,
        quantity: qty,
        uom: String(row[CONFIG.INVENTORY_COLS.UOM] || '').trim(),
        qualityStatus: String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || '').trim(),
        remarks: 'INWARD_REVERSAL_OF_INV=' + inventoryId + (movementId ? '; REVERSAL_OF=' + movementId : '') + '; ' + reason
      });
      reversed.push({
        inventoryId: inventoryId,
        itemCode: String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim(),
        batchId: String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
        ginNo: String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim(),
        version: String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim(),
        binId: binId,
        quantityReversed: qty,
        uom: String(row[CONFIG.INVENTORY_COLS.UOM] || '').trim()
      });
    });

    _appendMovementLogsBatch(logs);
    try { _updateBinAndRackStatuses(Object.keys(touchedBins).filter(Boolean)); } catch (e) { Logger.log('Bin status update failed: ' + e.message); }
    return { success: true, reversed: reversed };
  });
}

function reversePM1050PM1051DuplicateInwardMay3() {
  return reverseDuplicateInwardRows({
    reason: 'Reverse duplicate inward for PM1050/PM1051 invoice 000020260413A requested by warehouse',
    items: [
      { inventoryId: 'INV-1777806298436-339-0', itemCode: 'PM1050', batchId: '1847', ginNo: '000020260413A', version: 'V3', binId: 'ZERO - NOS', quantity: 10000 },
      { inventoryId: 'INV-1777806298436-92-1', itemCode: 'PM1051', batchId: '1846', ginNo: '000020260413A', version: 'V3', binId: 'ZERO - NOS', quantity: 9000 }
    ]
  });
}

function submitConsumptionV2(form) {
  return withScriptLock(function () {
    const user = protect(() => requireOperationalUser());
    if (!form.prodOrderNo) throw new Error("Production Order No required");

    // VALIDATION: Ensure Production Order is APPROVED before consumption
    _validateProductionOrderApproved(form.prodOrderNo);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const qaOverride = form && form.qaOverride === true;
    const overrideReason = String(form && form.qaOverrideReason || '').trim();
    if (qaOverride && user.role !== SECURITY.ROLES.MANAGER) {
      throw new Error('QA override requires manager approval');
    }
    if (qaOverride && !overrideReason) {
      throw new Error('QA override reason required');
    }

    // Expand bin-only rows into concrete batch rows (server-authoritative).
    const expandedItems = [];
    (form.items || []).forEach(function (item) {
      const itemCode = String(item.itemCode || '').trim();
      if (!itemCode) throw new Error('Item Code required');

      const qtyKgInput = _convertToKg(itemCode, item.quantity, item.uom);
      if (!isFinite(qtyKgInput) || qtyKgInput <= 0) {
        throw new Error('Invalid quantity for ' + itemCode);
      }
      if (qtyKgInput > 50000) {
        throw new Error('Quantity exceeds single-transaction limit of 50,000 KG - contact manager');
      }

      const explicitBatchId = String(item.batchId || '').trim();
      if (explicitBatchId) {
        expandedItems.push(Object.assign({}, item, {
          itemCode: itemCode,
          batchId: explicitBatchId,
          quantity: qtyKgInput,
          uom: 'KG'
        }));
        return;
      }

      // Bin-only path: resolve eligible rows across all batches in the selected bin.
      _validateBinExists(item.binId);
      const anyRows = _readInventoryState(itemCode, '', item.binId).filter(function (r) {
        return (Number(r.currentQty) || 0) > 0;
      });
      if (anyRows.length === 0) {
        _logInventoryVisibility('Consumption-binOnly', itemCode, '', item.binId, anyRows, anyRows, 'No matching inventory rows');
        throw new Error('No stock exists for ' + itemCode + ' in selected bin');
      }

      const evalAny = _validateQualityEligibility(
        itemCode,
        '__ANY__',
        anyRows,
        qtyKgInput,
        'CONSUMPTION',
        { allowUnapproved: qaOverride }
      );

      let remainingAny = qtyKgInput;
      const byBatch = {};
      (evalAny.orderedRows || []).forEach(function (row) {
        if (remainingAny <= 0) return;
        const rowBatch = String(row.rowData[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
        if (!rowBatch) return;
        const pick = Math.min(Number(row.currentQty) || 0, remainingAny);
        if (pick <= 0) return;
        byBatch[rowBatch] = (Number(byBatch[rowBatch]) || 0) + pick;
        remainingAny -= pick;
      });

      if (remainingAny > 0.000001) {
        throw new Error('FIFO quantity insufficient for ' + itemCode + ' in selected bin');
      }

      const splitBatches = Object.keys(byBatch);
      if (splitBatches.length === 0) {
        throw new Error('No valid batch found for ' + itemCode + ' in selected bin');
      }

      splitBatches.forEach(function (batchId) {
        expandedItems.push(Object.assign({}, item, {
          itemCode: itemCode,
          batchId: batchId,
          quantity: Number(byBatch[batchId]) || 0,
          uom: 'KG'
        }));
      });
    });

    expandedItems.forEach(function (item) {
      const ledgerBatchId = String(item.ledgerBatchId || '').trim();
      const ledgerRowIndex = Number(item.ledgerRowIndex || 0);
      const qtyKg = _convertToKg(item.itemCode, item.quantity, item.uom);
      if (qtyKg > 50000) {
        throw new Error('Quantity exceeds single-transaction limit of 50,000 KG - contact manager');
      }

      // Validate item code for inventory, but update the exact PO ledger item ID
      // returned by the server-side PO loader when present.
      const lookup = _getValidatedItemId(item.itemCode);
      const ledgerItemKey = String(item.ledgerItemId || '').trim() || lookup;
      const ledgerState = ledgerRowIndex
        ? _getProductionLedgerStateByRowIndex(form.prodOrderNo, ledgerRowIndex, ledgerItemKey, (ledgerBatchId || item.batchId), '')
        : _getProductionLedgerState(form.prodOrderNo, ledgerItemKey, (ledgerBatchId || item.batchId), '');
      if (ledgerState.rowIndex === -1) {
        throw new Error('Production ledger line not found for PO ' + form.prodOrderNo + ' / ' + item.itemCode);
      }
      const ledgerStatus = String(ledgerState.status || '').trim().toUpperCase();
      if (ledgerStatus !== 'APPROVED' && ledgerStatus !== 'OPEN') {
        throw new Error('Cannot consume PO line for ' + item.itemCode + '. Current status: ' + ledgerState.status);
      }
      _assertItemCodeBatch(item.itemCode, item.batchId, 'Consumption');

      // Validate bin exists before creating inventory
      _validateBinExists(item.binId);
      // EXECUTION ONLY - FIFO + QA enforced here.
      const allRows = _readInventoryState(item.itemCode, item.batchId, item.binId);
      const rows = allRows.filter(r => r.currentQty > 0);
      if (rows.length === 0) {
        _logInventoryVisibility('Consumption', item.itemCode, item.batchId, item.binId, allRows, rows, 'No matching inventory rows');
        throw new Error('No stock exists');
      }
      const evalRes = _validateQualityEligibility(item.itemCode, item.batchId, rows, qtyKg, 'CONSUMPTION', { allowUnapproved: qaOverride });

      const orderedRows = evalRes.orderedRows || [];
      const itemVersion = _getSingleInventoryVersion(orderedRows, `Consumption ${item.itemCode}`);
      const changes = [];
      const logs = [];
      let ledgerChange = null;
      let prodTransferChange = null;
      const qaStatusChanges = [];
      const qaEvents = [];
      let qaEventChange = null;
      const qaMeta = _buildQaOverrideMeta(user, 'CONSUMPTION', overrideReason, qaOverride);
      const overrideContext = {
        itemCode: item.itemCode,
        batchId: item.batchId,
        binId: item.binId,
        overrideReason: overrideReason,
        userEmail: user.email,
        remarks: _formatRemarksWithOverride(form.remarks, qaMeta)
      };
      let remaining = qtyKg;

      for (const row of orderedRows) {
        if (remaining <= 0) break;
        const pick = Math.min(row.currentQty, remaining);
        if (qaOverride) {
          _applyQaOverrideForRow(row, overrideContext, qaStatusChanges, qaEvents);
        }

        logs.push({
          type: CONFIG.MOVEMENT_TYPES.CONSUMPTION,
          itemId: lookup,
          batchId: item.batchId,
          version: itemVersion,
          ginNo: row.rowData[CONFIG.INVENTORY_COLS.GIN_NO],
          prodOrderRef: form.prodOrderNo, fromBinId: item.binId,
          quantity: pick, remarks: _formatRemarksWithOverride(form.remarks, qaMeta)
        });

        changes.push({
          rowIndex: row.rowIndex,
          prevQty: row.currentQty,
          prevUpdated: row.rowData[CONFIG.INVENTORY_COLS.LAST_UPDATED],
          newQty: row.currentQty - pick
        });
        remaining -= pick;
      }

      if (remaining > 0) {
        throw new Error('FIFO quantity insufficient');
      }

      try {
        // Production transfer audit row (bridge between PO, inventory, movement)
        const meta = orderedRows[0] && orderedRows[0].rowData ? orderedRows[0].rowData : [];
        const fromBinId = item.binId;
        const binMeta = _getBinMetaById(fromBinId);
        const batchMeta = _getBatchById(item.batchId, item.itemCode) || {};
        prodTransferChange = _appendProductionTransfersBatch([{
          transferType: 'CONSUMPTION',
          productionOrderNo: form.prodOrderNo,
          productionArea: String(form.productionArea || 'PRODUCTION').trim(),
          itemCode: item.itemCode,
          itemId: lookup,
          batchNumber: batchMeta.batchNumber || item.batchId,
          batchId: item.batchId,
          lotNumber: batchMeta.lotNumber || item.lotNo || '',
          quantity: qtyKg,
          uom: _getItemUomCode(item.itemCode),
          fromSite: String(meta[CONFIG.INVENTORY_COLS.SITE] || '').trim(),
          fromLocation: String(meta[CONFIG.INVENTORY_COLS.LOCATION] || '').trim(),
          fromBinId: fromBinId,
          fromBinCode: binMeta.binCode || fromBinId,
          toSite: String(form.productionArea || 'PRODUCTION').trim(),
          toLocation: String(form.productionArea || 'PRODUCTION').trim(),
          toBinId: String(form.productionArea || 'PRODUCTION').trim(),
          toBinCode: String(form.productionArea || 'PRODUCTION').trim(),
          returnedByName: '',
          createdBy: user.email,
          createdAt: new Date(),
          remarks: String(form.remarks || ''),
          status: 'COMPLETED'
        }]);

        _persistInventoryStatesBatch(changes, sheet);
        _applyQaStatusChanges(sheet, qaStatusChanges);
        qaEventChange = _appendQaEventsBatch(qaEvents);
        ledgerChange = _updateProductionLedger(
          form.prodOrderNo,
          ledgerItemKey,
          (ledgerBatchId || item.batchId),
          itemVersion,
          qtyKg,
          0,
          0,
          ledgerRowIndex
        );
        _appendMovementLogsBatch(logs);
        try { _updateBinAndRackStatuses([item.binId]); } catch (e) { Logger.log('Bin status update failed: ' + e.message); }
      } catch (e) {
        _rollbackInventoryChanges(changes, sheet);
        _rollbackQaStatusChanges(sheet, qaStatusChanges);
        _rollbackQaEvents(qaEventChange);
        _rollbackProductionLedger(ledgerChange);
        _rollbackProductionTransfers(prodTransferChange);
        throw e;
      }
      try { _cleanupZeroQtyInventoryRows(changes, sheet); } catch (e) { Logger.log('Zero-qty cleanup failed: ' + e.message); }
    });
    return { success: true };
  });
}


/**
 * CRITICAL FIX: Updated transfer function to use itemCode directly if validation fails
 */
/**
 * FIXED: submitTransferV2 - Handles single and split-bin transfers
 * 
 * CRITICAL FIXES:
 * 1. Update existing inventory rows instead of creating new ones
 * 2. Log actual destination bin IDs (not "SPLIT")
 * 3. Proper FIFO deduction from source bins
 * 4. Proper credit to destination bins (update existing or create new)
 * 
 * REPLACE existing submitTransferV2 function (around line 5089) with this:
 */

function submitTransferV2(form) {
  return withScriptLock(function () {
    protect(() => requireRole(SECURITY.ROLES.MANAGER));

    if (!form.items || !Array.isArray(form.items)) {
      throw new Error("Invalid items");
    }
    if (!form.toSite) {
      throw new Error("Destination site required");
    }

    const destSplits = Array.isArray(form.destSplits) ? form.destSplits : [];
    if (!form.toBinId && destSplits.length === 0) {
      throw new Error("Destination Bin ID required");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const invData = _getSheetValuesCached(invSheet.getName());

    // 1. EXPAND SOURCE ITEMS (Handle source split bins)
    const expandedItems = [];
    form.items.forEach(function (item, idx) {
      const splits = Array.isArray(item && item.sourceSplits) ? item.sourceSplits : [];
      if (!splits.length) {
        expandedItems.push(item);
        return;
      }

      const expectedKg = _convertToKg(item.itemCode, item.quantity, item.uom);
      let splitKg = 0;

      splits.forEach(function (sp, sIdx) {
        if (!sp.binId) {
          throw new Error('Item ' + (idx + 1) + ': Split ' + (sIdx + 1) + ' bin required');
        }
        const spQty = Number(sp.quantity);
        if (spQty <= 0) {
          throw new Error('Item ' + (idx + 1) + ': Split ' + (sIdx + 1) + ' quantity must be > 0');
        }
        splitKg += _convertToKg(item.itemCode, spQty, item.uom);
        expandedItems.push(Object.assign({}, item, {
          binId: sp.binId,
          quantity: spQty
        }));
      });

      if (Math.abs(splitKg - expectedKg) > 0.001) {
        throw new Error('Item ' + (idx + 1) + ': source split quantities must match row total');
      }
    });

    const transferTs = new Date();
    const userEmail = Session.getActiveUser().getEmail();

    // 2. BUILD TRANSFER PLAN
    // We'll collect what needs to move from where to where
    const transferPlan = []; // { fromBin, toBin, itemCode, batchId, version, qty, ginNo, qa, qd, itemId }

    // If single destination, distribute proportionally
    if (destSplits.length === 0) {
      // Single destination bin
      expandedItems.forEach(function (item) {
        const itemCode = String(item.itemCode || '').trim();
        const batchId = String(item.batchId || '').trim();
        const fromBinId = String(item.binId || '').trim();
        const qtyKg = _convertToKg(itemCode, item.quantity, item.uom);

        if (!itemCode || !batchId || !fromBinId) {
          throw new Error('Item code, batch ID, and source bin required');
        }

        let itemId;
        try { itemId = _getValidatedItemId(itemCode); } catch (e) { itemId = itemCode; }

        // Find source rows via FIFO
        const allRows = _readInventoryState(itemCode, batchId, fromBinId);
        const rows = allRows.filter(r => r.currentQty > 0);

        if (rows.length === 0) {
          throw new Error('No stock exists for ' + itemCode + ' in bin ' + fromBinId);
        }

        const evalRes = _validateTransferEligibility(itemCode, batchId, rows, qtyKg);
        const orderedRows = evalRes.orderedRows || [];
        const itemVersion = _getSingleInventoryVersion(orderedRows, 'Transfer ' + itemCode);

        let remaining = qtyKg;
        for (const row of orderedRows) {
          if (remaining <= 0) break;
          const pick = Math.min(row.currentQty, remaining);

          transferPlan.push({
            fromBin: fromBinId,
            toBin: form.toBinId,
            itemCode: itemCode,
            batchId: batchId,
            version: itemVersion,
            qty: pick,
            ginNo: String(row.rowData[CONFIG.INVENTORY_COLS.GIN_NO] || ''),
            qa: String(row.rowData[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').toUpperCase(),
            qd: row.rowData[CONFIG.INVENTORY_COLS.QUALITY_DATE] || '',
            itemId: itemId,
            uom: item.uom,
            sourceRowIndex: row.rowIndex,
            inwardDate: row.rowData[CONFIG.INVENTORY_COLS.INWARD_DATE] || row.rowData[CONFIG.INVENTORY_COLS.LAST_UPDATED] || transferTs
          });

          remaining -= pick;
        }

        if (remaining > 0.001) {
          throw new Error('Insufficient stock for ' + itemCode + ' in bin ' + fromBinId);
        }
      });
    } else {
      // Split destination bins
      // Calculate total quantity needed per destination
      const destTotals = {};
      destSplits.forEach(function (split) {
        const binId = String(split.binId || '').trim();
        if (!binId) throw new Error('Destination bin ID required in split');
        const qty = Number(split.quantity) || 0;
        if (qty <= 0) throw new Error('Destination split quantity must be > 0');

        if (!destTotals[binId]) destTotals[binId] = 0;
        destTotals[binId] += qty; // Already in base units (KG/NOS)
      });

      // Collect all available stock from source bins
      const stockPool = [];
      let totalAvailable = 0;

      expandedItems.forEach(function (item) {
        const itemCode = String(item.itemCode || '').trim();
        const batchId = String(item.batchId || '').trim();
        const fromBinId = String(item.binId || '').trim();
        const qtyKg = _convertToKg(itemCode, item.quantity, item.uom);

        let itemId;
        try { itemId = _getValidatedItemId(itemCode); } catch (e) { itemId = itemCode; }

        const allRows = _readInventoryState(itemCode, batchId, fromBinId);
        const rows = allRows.filter(r => r.currentQty > 0);

        if (rows.length === 0) {
          throw new Error('No stock in ' + fromBinId + ' for ' + itemCode);
        }

        const evalRes = _validateTransferEligibility(itemCode, batchId, rows, qtyKg);
        const orderedRows = evalRes.orderedRows || [];
        const itemVersion = _getSingleInventoryVersion(orderedRows, 'Transfer ' + itemCode);

        let remaining = qtyKg;
        for (const row of orderedRows) {
          if (remaining <= 0) break;
          const pick = Math.min(row.currentQty, remaining);

          stockPool.push({
            fromBin: fromBinId,
            itemCode: itemCode,
            batchId: batchId,
            version: itemVersion,
            qty: pick,
            ginNo: String(row.rowData[CONFIG.INVENTORY_COLS.GIN_NO] || ''),
            qa: String(row.rowData[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').toUpperCase(),
            qd: row.rowData[CONFIG.INVENTORY_COLS.QUALITY_DATE] || '',
            itemId: itemId,
            uom: item.uom,
            sourceRowIndex: row.rowIndex,
            inwardDate: row.rowData[CONFIG.INVENTORY_COLS.INWARD_DATE] || row.rowData[CONFIG.INVENTORY_COLS.LAST_UPDATED] || transferTs
          });

          totalAvailable += pick;
          remaining -= pick;
        }

        if (remaining > 0.001) {
          throw new Error('Insufficient stock for ' + itemCode + ' in bin ' + fromBinId);
        }
      });

      // Distribute stock to destination bins
      Object.keys(destTotals).forEach(function (destBin) {
        let needed = destTotals[destBin];

        for (let i = 0; i < stockPool.length && needed > 0.001; i++) {
          const stock = stockPool[i];
          if (stock.qty <= 0) continue;

          const take = Math.min(stock.qty, needed);

          transferPlan.push({
            fromBin: stock.fromBin,
            toBin: destBin,
            itemCode: stock.itemCode,
            batchId: stock.batchId,
            version: stock.version,
            qty: take,
            ginNo: stock.ginNo,
            qa: stock.qa,
            qd: stock.qd,
            itemId: stock.itemId,
            uom: stock.uom,
            sourceRowIndex: stock.sourceRowIndex,
            inwardDate: stock.inwardDate
          });

          stock.qty -= take;
          needed -= take;
        }

        if (needed > 0.001) {
          throw new Error('Not enough stock to fill destination bin ' + destBin);
        }
      });
    }

    // 3. VALIDATE DESTINATION BINS EXIST
    const destBinIds = {};
    transferPlan.forEach(function (t) {
      destBinIds[t.toBin] = true;
    });
    Object.keys(destBinIds).forEach(function (binId) {
      _validateBinExists(binId);
    });

    // 4. CHECK VERSION COMPATIBILITY AT DESTINATION
    const destVersionMap = {};
    transferPlan.forEach(function (t) {
      const key = t.itemCode + '|' + t.batchId + '|' + t.toBin;
      if (!destVersionMap[key]) {
        const batchInfo = _resolveBatchReference(t.itemCode, t.batchId);
        const acceptedBatchKeys = batchInfo.acceptedKeys || {};
        // Check what version exists in destination bin
        const destRows = invData.filter(function (row) {
          return String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim() === t.itemCode &&
            !!acceptedBatchKeys[String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim().toUpperCase()] &&
            String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim() === t.toBin &&
            Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] || 0) > 0;
        });

        if (destRows.length > 0) {
          const existingVersion = String(destRows[0][CONFIG.INVENTORY_COLS.VERSION] || '').trim().toUpperCase();
          if (existingVersion && existingVersion !== t.version) {
            throw new Error('Destination bin ' + t.toBin + ' has version ' + existingVersion + ' but transferring version ' + t.version);
          }
          destVersionMap[key] = existingVersion || t.version;
        } else {
          destVersionMap[key] = t.version;
        }
      }
    });

    // 5. EXECUTE TRANSFER (Debit + Credit)
    const inventoryUpdates = {}; // rowIndex -> new quantity
    const movementLogs = [];

    // Group transfer plan by source row for efficient deduction
    const debitMap = {}; // rowIndex -> total qty to deduct
    transferPlan.forEach(function (t) {
      if (!debitMap[t.sourceRowIndex]) debitMap[t.sourceRowIndex] = 0;
      debitMap[t.sourceRowIndex] += t.qty;
    });

    // Apply debits
    Object.keys(debitMap).forEach(function (rowIdx) {
      const idx = Number(rowIdx);
      const row = invData[idx - 1]; // Sheet is 1-indexed, array is 0-indexed
      const currentQty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      const newQty = currentQty - debitMap[rowIdx];

      if (newQty < -0.001) {
        throw new Error('Cannot deduct more than available from inventory row ' + idx);
      }

      inventoryUpdates[idx] = Math.max(0, newQty);
    });

    // Apply credits and log movements
    transferPlan.forEach(function (t) {
      const destBinMeta = _getBinMetaById(t.toBin);
      const destSite = destBinMeta.site || form.toSite || '';
      const destLocation = destBinMeta.location || form.toLocation || '';

      // Find or create destination inventory row
      let destRowIdx = null;
      const batchInfo = _resolveBatchReference(t.itemCode, t.batchId);
      const acceptedBatchKeys = batchInfo.acceptedKeys || {};
      for (let i = 1; i < invData.length; i++) {
        const row = invData[i];
        if (String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim() === t.itemCode &&
          !!acceptedBatchKeys[String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim().toUpperCase()] &&
          String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim() === t.toBin &&
          String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim().toUpperCase() === t.version &&
          String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim() === t.ginNo) {
          destRowIdx = i + 1;
          break;
        }
      }

      if (destRowIdx) {
        // Update existing row
        const currentQty = Number(invData[destRowIdx - 1][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        const newQty = currentQty + t.qty;
        inventoryUpdates[destRowIdx] = newQty;
      } else {
        // Create new row
        const newRow = [];
        newRow[CONFIG.INVENTORY_COLS.INVENTORY_ID] = 'INV-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
        newRow[CONFIG.INVENTORY_COLS.ITEM_ID] = t.itemId;
        newRow[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] = t.itemCode;
        newRow[CONFIG.INVENTORY_COLS.BATCH_ID] = t.batchId;
        newRow[CONFIG.INVENTORY_COLS.GIN_NO] = t.ginNo;
        newRow[CONFIG.INVENTORY_COLS.VERSION] = t.version;
        newRow[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = t.qa;
        newRow[CONFIG.INVENTORY_COLS.QUALITY_DATE] = t.qd;
        newRow[CONFIG.INVENTORY_COLS.BIN_ID] = t.toBin;
        newRow[CONFIG.INVENTORY_COLS.SITE] = destSite;
        newRow[CONFIG.INVENTORY_COLS.LOCATION] = destLocation;
        newRow[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = t.qty;
        newRow[CONFIG.INVENTORY_COLS.UOM] = t.uom || 'KG';
        newRow[CONFIG.INVENTORY_COLS.LOT_NO] = '';
        newRow[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = '';
        newRow[CONFIG.INVENTORY_COLS.LAST_UPDATED] = transferTs;
        newRow[CONFIG.INVENTORY_COLS.INWARD_DATE] = t.inwardDate;
        newRow[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = transferTs;

        invSheet.appendRow(newRow);
        _appendSheetCacheRow(CONFIG.SHEETS.INVENTORY, newRow);
      }

      // Log movement
      const movType = (t.qa === 'APPROVED' || t.qa === 'OVERRIDDEN') ?
        CONFIG.MOVEMENT_TYPES.TRANSFER :
        CONFIG.MOVEMENT_TYPES.TRANSFER_QUARANTINE;

      movementLogs.push({
        type: movType,
        itemId: t.itemId,
        batchId: t.batchId,
        version: t.version,
        ginNo: t.ginNo,
        fromBinId: t.fromBin,
        toBinId: t.toBin, // <-- ACTUAL DESTINATION BIN, NOT "SPLIT"
        quantity: t.qty,
        remarks: form.remarks || 'Transfer to ' + form.toSite
      });
    });

    // 6. WRITE UPDATES TO SHEET
    Object.keys(inventoryUpdates).forEach(function (rowIdx) {
      const idx = Number(rowIdx);
      const newQty = inventoryUpdates[idx];

      invSheet.getRange(idx, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1).setValue(newQty);
      invSheet.getRange(idx, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1).setValue(transferTs);
      invSheet.getRange(idx, CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE + 1).setValue(transferTs);

      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, idx, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1, newQty);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, idx, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1, transferTs);
      _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, idx, CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE + 1, transferTs);
    });

    // 7. LOG MOVEMENTS
    _appendMovementLogsBatch(movementLogs);

    return {
      success: true,
      message: 'Transfer completed: ' + transferPlan.length + ' movements logged',
      movements: transferPlan.length
    };
  });
}
// _creditToBin: removed (orphaned dead code from previous edit)



/**
 * Internal helper to find ledger row
 */
function _isProductionLedgerStatus(v) {
  const t = String(v || '').trim().toUpperCase();
  return t === 'OPEN' || t === 'APPROVED' || t === 'REJECTED' || t === 'CLOSED';
}

function _detectProductionLedgerStatusIdx(row, fallbackIdx) {
  const candidates = [fallbackIdx, 8, 9, 10, 11];
  for (let i = 0; i < candidates.length; i++) {
    const idx = Number(candidates[i]);
    if (!isFinite(idx) || idx < 0) continue;
    if (_isProductionLedgerStatus(row[idx])) return idx;
  }
  return (typeof fallbackIdx === 'number') ? fallbackIdx : CONFIG.PRODUCTION_LEDGER_COLS.STATUS;
}

function _detectProductionLedgerUpdatedIdx(row, statusIdx, fallbackIdx) {
  const candidates = [fallbackIdx, statusIdx + 1, 9, 10, 11, 12];
  for (let i = 0; i < candidates.length; i++) {
    const idx = Number(candidates[i]);
    if (!isFinite(idx) || idx < 0 || idx === statusIdx) continue;
    const v = row[idx];
    if (v instanceof Date) return idx;
    if (String(v || '').trim() && !isNaN(Date.parse(v))) return idx;
  }
  return (typeof fallbackIdx === 'number') ? fallbackIdx : CONFIG.PRODUCTION_LEDGER_COLS.LAST_UPDATED;
}

function _emptyProductionLedgerState() {
  return {
    rowIndex: -1, requested: 0, issued: 0, returned: 0, rejected: 0,
    status: 'NOT_FOUND', statusIdx: CONFIG.PRODUCTION_LEDGER_COLS.STATUS,
    updatedIdx: CONFIG.PRODUCTION_LEDGER_COLS.LAST_UPDATED,
    version: '', updateBatch: false, updateVersion: false
  };
}

function _resolveProductionLedgerItemKeys(itemIdOrCode) {
  let targetId = String(itemIdOrCode || '').trim().toUpperCase();
  let targetCode = String(itemIdOrCode || '').trim().toUpperCase();
  try {
    const item = getItemByCodeCached(itemIdOrCode) || _getItemByIdCached(itemIdOrCode);
    if (item) {
      targetId = String(item.id || '').trim().toUpperCase();
      targetCode = String(item.code || '').trim().toUpperCase();
    }
  } catch (e) { }
  return { id: targetId, code: targetCode };
}

function _buildProductionLedgerState(rowIndex, row, searchBatch, searchVersion) {
  const cols = CONFIG.PRODUCTION_LEDGER_COLS;
  const rowBatch = String(row[cols.BATCH_ID] || '').trim().toUpperCase();
  const rowVersion = String(row[cols.VERSION] || '').trim().toUpperCase();
  const statusIdx = _detectProductionLedgerStatusIdx(row, cols.STATUS);
  const updatedIdx = _detectProductionLedgerUpdatedIdx(row, statusIdx, cols.LAST_UPDATED);
  const batchIsPlaceholder = (rowBatch === '' || rowBatch === 'TBD');
  const versionIsPlaceholder = (rowVersion === '' || rowVersion === 'TBD');

  return {
    rowIndex: rowIndex,
    requested: Number(row[cols.QTY_REQUESTED]) || 0,
    issued: Number(row[cols.QTY_ISSUED]) || 0,
    returned: Number(row[cols.QTY_RETURNED]) || 0,
    rejected: Number(row[cols.QTY_REJECTED]) || 0,
    status: row[statusIdx],
    statusIdx: statusIdx,
    updatedIdx: updatedIdx,
    version: rowVersion,
    updateBatch: !!searchBatch && batchIsPlaceholder,
    updateVersion: !!searchVersion && versionIsPlaceholder
  };
}

function _chooseProductionLedgerCandidate(candidates) {
  let best = null;
  let bestOut = -Infinity;
  (candidates || []).forEach(function (candidate) {
    const status = String(candidate.status || '').trim().toUpperCase();
    const active = (status === 'OPEN' || status === 'APPROVED');
    const outstanding = Number(candidate.requested || 0) - Number(candidate.issued || 0) + Number(candidate.returned || 0);
    if (!best) {
      best = candidate;
      bestOut = outstanding;
      return;
    }
    const bestStatus = String(best.status || '').trim().toUpperCase();
    const bestActive = (bestStatus === 'OPEN' || bestStatus === 'APPROVED');
    if (active && !bestActive) {
      best = candidate;
      bestOut = outstanding;
    } else if (active === bestActive && outstanding > bestOut) {
      best = candidate;
      bestOut = outstanding;
    }
  });
  return best || null;
}

function _getProductionLedgerStateByRowIndex(orderRef, rowIndex, itemId, batchId, version) {
  const targetRowIndex = Number(rowIndex);
  if (!isFinite(targetRowIndex) || targetRowIndex <= 1) return _emptyProductionLedgerState();

  const searchBatch = String(batchId || '').trim().toUpperCase();
  const searchVersion = String(version || '').trim().toUpperCase();
  const orderRows = _getProductionLedgerOrderRows(orderRef);

  for (let i = 0; i < orderRows.length; i++) {
    if (Number(orderRows[i].rowIndex) !== targetRowIndex) continue;
    const row = orderRows[i].row;
    // Exact row index is emitted by the server-side PO loader for this PO.
    // Do not re-resolve item code here; legacy/duplicate Item_Master IDs were
    // the original reason valid loaded POs failed during final consumption.
    return _buildProductionLedgerState(orderRows[i].rowIndex, row, searchBatch, searchVersion);
  }

  return _emptyProductionLedgerState();
}

function _getProductionLedgerState(orderRef, itemId, batchId, version) {
  const keys = _resolveProductionLedgerItemKeys(itemId);
  const searchBatch = String(batchId || '').trim().toUpperCase();
  const searchVersion = String(version || '').trim().toUpperCase();
  const candidates = [];
  const candidatesAnyBatch = [];
  const orderRows = _getProductionLedgerOrderRows(orderRef);
  const cols = CONFIG.PRODUCTION_LEDGER_COLS;

  for (let i = 0; i < orderRows.length; i++) {
    const row = orderRows[i].row;
    const rowItemId = String(row[cols.ITEM_ID] || '').trim().toUpperCase();
    if (rowItemId !== keys.id && rowItemId !== keys.code) continue;

    const rowBatch = String(row[cols.BATCH_ID] || '').trim().toUpperCase();
    const rowVersion = String(row[cols.VERSION] || '').trim().toUpperCase();
    const batchIsPlaceholder = (rowBatch === '' || rowBatch === 'TBD');
    const versionIsPlaceholder = (rowVersion === '' || rowVersion === 'TBD');
    const batchMatches = searchBatch ? (rowBatch === searchBatch || batchIsPlaceholder) : true;
    const versionMatches = searchVersion ? (rowVersion === searchVersion || versionIsPlaceholder) : true;
    const versionMatchesAnyBatch = searchVersion ? (rowVersion === searchVersion || versionIsPlaceholder) : true;

    if (versionMatchesAnyBatch) {
      candidatesAnyBatch.push(_buildProductionLedgerState(orderRows[i].rowIndex, row, '', searchVersion));
    }
    if (batchMatches && versionMatches) {
      candidates.push(_buildProductionLedgerState(orderRows[i].rowIndex, row, searchBatch, searchVersion));
    }
  }

  if (candidates.length > 0) return _chooseProductionLedgerCandidate(candidates);
  if (searchBatch && candidatesAnyBatch.length > 0) return _chooseProductionLedgerCandidate(candidatesAnyBatch);
  return _emptyProductionLedgerState();
}

/**
 * Updates production actuals
 */
function _updateProductionLedger(orderRef, itemId, batchId, version, qtyIssuedInc, qtyReturnedInc, qtyRejectedInc, ledgerRowIndex) {
  const state = ledgerRowIndex
    ? _getProductionLedgerStateByRowIndex(orderRef, ledgerRowIndex, itemId, batchId, version)
    : _getProductionLedgerState(orderRef, itemId, batchId, version);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
  const cols = CONFIG.PRODUCTION_LEDGER_COLS;
  const issuedInc = Number(qtyIssuedInc) || 0;
  const returnedInc = Number(qtyReturnedInc) || 0;
  const rejectedInc = Number(qtyRejectedInc) || 0;

  if (state.rowIndex === -1) {
    if (!version) throw new Error('Ledger version required');
    const requested = Math.max(issuedInc, 0);
    const issued = issuedInc;
    const returned = returnedInc;
    const rejected = rejectedInc;
    const net = requested - issued + returned;
    const status = 'OPEN';
    const rowIndex = sheet.getLastRow() + 1;
    const newRow = [
      orderRef,
      itemId,
      batchId || 'TBD',
      version,
      requested,
      issued,
      returned,
      net,
      rejected,
      status,
      new Date()
    ];
    sheet.appendRow(newRow);
    _appendSheetCacheRow(CONFIG.SHEETS.PRODUCTION_LEDGER, newRow);
    delete _REQ_CACHE.productionLedgerRowsByOrder;
    _clearScriptCachesForSheet(CONFIG.SHEETS.PRODUCTION_LEDGER);
    return { action: 'create', rowIndex: rowIndex };
  }

  if (state.version && version && state.version !== version && state.version !== 'TBD') {
    throw new Error(`Version mismatch for ledger ${orderRef} / ${itemId}`);
  }

  const change = {
    action: 'update',
    rowIndex: state.rowIndex,
    prevIssued: state.issued,
    prevReturned: state.returned,
    prevRejected: state.rejected || 0,
    prevNet: (state.requested || 0) - state.issued + state.returned,
    prevStatus: state.status,
    prevUpdated: undefined,
    prevBatch: undefined,
    prevVersion: undefined
  };

  const newIssued = state.issued + issuedInc;
  const newReturned = state.returned + returnedInc;
  const newRejected = (state.rejected || 0) + rejectedInc;
  const net = (state.requested || 0) - newIssued + newReturned;
  const nextStatus = (state.requested > 0 && net <= 0) ? 'CLOSED' : state.status;
  const updatedAt = new Date();
  const statusIdx = (typeof state.statusIdx === 'number') ? state.statusIdx : cols.STATUS;
  const updatedIdx = (typeof state.updatedIdx === 'number') ? state.updatedIdx : cols.LAST_UPDATED;

  const ledgerData = _getSheetValuesCached(sheet.getName());
  const rowData = (ledgerData[state.rowIndex - 1] || []).slice();
  const requiredCols = Math.max(updatedIdx + 1, statusIdx + 1, cols.QTY_REJECTED + 1, rowData.length);
  while (rowData.length < requiredCols) rowData.push('');

  change.prevUpdated = rowData[updatedIdx];
  if (state.updateBatch && batchId) {
    change.prevBatch = rowData[cols.BATCH_ID];
    rowData[cols.BATCH_ID] = batchId;
  }
  if (state.updateVersion && version) {
    change.prevVersion = rowData[cols.VERSION];
    rowData[cols.VERSION] = version;
  }

  rowData[cols.QTY_ISSUED] = newIssued;
  rowData[cols.QTY_RETURNED] = newReturned;
  rowData[cols.NET_OUTSTANDING] = net;
  rowData[cols.QTY_REJECTED] = newRejected;
  rowData[statusIdx] = nextStatus;
  rowData[updatedIdx] = updatedAt;
  sheet.getRange(state.rowIndex, 1, 1, requiredCols).setValues([rowData]);

  if (state.updateBatch && batchId) {
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, cols.BATCH_ID + 1, batchId);
  }
  if (state.updateVersion && version) {
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, cols.VERSION + 1, version);
  }
  _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, cols.QTY_ISSUED + 1, newIssued);
  _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, cols.QTY_RETURNED + 1, newReturned);
  _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, cols.NET_OUTSTANDING + 1, net);
  _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, cols.QTY_REJECTED + 1, newRejected);
  _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, statusIdx + 1, nextStatus);
  _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, state.rowIndex, updatedIdx + 1, updatedAt);
  _clearScriptCachesForSheet(CONFIG.SHEETS.PRODUCTION_LEDGER);
  return change;
}



// =====================================================
// PRODUCTION REQUEST FLOW
// =====================================================


/**
 * Get active production orders for manager visibility.
 * Returns both APPROVED and legacy OPEN lines so the manager screen can
 * act as a read-only order summary instead of an approval queue.
 */

function getPendingProductionOrders() {
  _beginRequestCache();
  try {
    requireRole(SECURITY.ROLES.MANAGER);

    const cacheKey = 'ACTIVE_PRODUCTION_ORDERS_V1';
    const cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
    const data = _getSheetValuesCached(sheet.getName());
    const itemMaps = _getItemMasterMaps();

    const result = [];
    const validStatusSet = { OPEN: true, APPROVED: true, REJECTED: true, CLOSED: true };

    function detectStatusIdx(row) {
      const candidates = [CONFIG.PRODUCTION_LEDGER_COLS.STATUS, 8, 9, 10, 11];
      for (let c = 0; c < candidates.length; c++) {
        const idx = Number(candidates[c]);
        if (!isFinite(idx) || idx < 0) continue;
        const v = String(row[idx] || '').trim().toUpperCase();
        if (validStatusSet[v]) return idx;
      }
      return CONFIG.PRODUCTION_LEDGER_COLS.STATUS;
    }

    function detectUpdatedIdx(row, statusIdx) {
      const candidates = [CONFIG.PRODUCTION_LEDGER_COLS.LAST_UPDATED, statusIdx + 1, 9, 10, 11, 12];
      for (let c = 0; c < candidates.length; c++) {
        const idx = Number(candidates[c]);
        if (!isFinite(idx) || idx < 0 || idx === statusIdx) continue;
        const v = row[idx];
        if (v instanceof Date) return idx;
        if (String(v || '').trim() && !isNaN(Date.parse(v))) return idx;
      }
      return CONFIG.PRODUCTION_LEDGER_COLS.LAST_UPDATED;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const statusIdx = detectStatusIdx(row);
      const status = String(row[statusIdx] || '').trim().toUpperCase();

      // Production orders are now auto-approved. Keep legacy OPEN rows visible
      // so older orders still appear and remain consumable.
      if (status === 'CLOSED' || status === 'REJECTED') {
        continue;
      }

      const orderRef = String(row[CONFIG.PRODUCTION_LEDGER_COLS.ORDER_ID] || '');
      const itemId = String(row[CONFIG.PRODUCTION_LEDGER_COLS.ITEM_ID] || '');
      // Then inside loop:
      const itemCode = itemMaps.idToCode[String(itemId || '').toUpperCase()] || itemId;
      const itemName = itemMaps.codeToName[String(itemCode || '').trim().toUpperCase()] || itemCode;
      const uomCode = itemMaps.codeToUom[String(itemCode || '').trim().toUpperCase()] || _getItemUomCode(itemCode);
      const batchId = String(row[CONFIG.PRODUCTION_LEDGER_COLS.BATCH_ID] || '');
      const batchNumber = batchId ? _getBatchDisplayNumber(itemCode, batchId) : '';
      const quantity = Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REQUESTED]) || Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_ISSUED]) || 0;
      const rejected = Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REJECTED]) || 0;
      const outstanding = Number(row[CONFIG.PRODUCTION_LEDGER_COLS.NET_OUTSTANDING]) || 0;
      const lastUpdated = row[detectUpdatedIdx(row, statusIdx)] || new Date();

      result.push({
        orderRef: orderRef,
        itemId: itemId,
        itemCode: itemCode,
        itemName: itemName,
        uomCode: uomCode,
        batchId: batchId,
        batchNumber: batchNumber,
        quantity: quantity,
        rejected: rejected,
        outstanding: outstanding,
        status: status,
        lastUpdated: lastUpdated instanceof Date ? lastUpdated.toISOString() : String(lastUpdated)
      });
    }

    _putScriptCacheJson(cacheKey, result, 30);
    return result;

  } catch (error) {
    Logger.log('[PO_BACKEND] EXCEPTION: ' + error.toString());
    Logger.log('[PO_BACKEND] Message: ' + error.message);
    throw error;
  } finally {
    _endRequestCache();
  }
}

/**
 * Update production order approval status
 * @param {Object} payload - {orderRef, itemId?, batchId?, newStatus, remarks}
 * @returns {Object} Success response
 */
function updateProductionOrderStatus(payload) {
  return withScriptLock(function () {
    protect(() => requireRole(SECURITY.ROLES.MANAGER));

    if (!payload.orderRef) throw new Error('Order Ref required');
    if (!payload.newStatus) throw new Error('Status required');
    const targetOrderRef = String(payload.orderRef || '').trim();
    const targetItemId = String(payload.itemId || '').trim().toUpperCase();
    const targetBatchId = String(payload.batchId || '').trim().toUpperCase();

    const validStatuses = ['APPROVED', 'REJECTED', 'OPEN'];
    const newStatus = String(payload.newStatus || '').trim().toUpperCase();
    if (!validStatuses.includes(newStatus)) {
      throw new Error(`Invalid status. Must be: ${validStatuses.join(', ')}`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
    const data = _getSheetValuesCached(sheet.getName());
    const cols = CONFIG.PRODUCTION_LEDGER_COLS;
    const validStatusSet = { OPEN: true, APPROVED: true, REJECTED: true, CLOSED: true };

    function detectStatusIdx(row) {
      const candidates = [cols.STATUS, 8, 9, 10, 11];
      for (let c = 0; c < candidates.length; c++) {
        const idx = Number(candidates[c]);
        if (!isFinite(idx) || idx < 0) continue;
        const v = String(row[idx] || '').trim().toUpperCase();
        if (validStatusSet[v]) return idx;
      }
      return cols.STATUS;
    }

    function detectUpdatedIdx(row, statusIdx) {
      const candidates = [cols.LAST_UPDATED, statusIdx + 1, 9, 10, 11, 12];
      for (let c = 0; c < candidates.length; c++) {
        const idx = Number(candidates[c]);
        if (!isFinite(idx) || idx < 0 || idx === statusIdx) continue;
        const v = row[idx];
        if (v instanceof Date) return idx;
        if (String(v || '').trim() && !isNaN(Date.parse(v))) return idx;
      }
      return cols.LAST_UPDATED;
    }

    const candidates = [];
    for (let i = 1; i < data.length; i++) {
      const rowOrderRef = String(data[i][cols.ORDER_ID] || '').trim();
      if (rowOrderRef !== targetOrderRef) continue;

      if (targetItemId) {
        const rowItemId = String(data[i][cols.ITEM_ID] || '').trim().toUpperCase();
        if (rowItemId !== targetItemId) continue;
      }

      if (targetBatchId) {
        const rowBatchId = String(data[i][cols.BATCH_ID] || '').trim().toUpperCase();
        if (rowBatchId !== targetBatchId) continue;
      }

      const statusIdx = detectStatusIdx(data[i]);
      const updatedIdx = detectUpdatedIdx(data[i], statusIdx);
      candidates.push({ rowIndex: i, statusIdx: statusIdx, updatedIdx: updatedIdx });
    }

    if (candidates.length === 0) throw new Error('Production Order line not found');

    let target = null;
    for (let c = 0; c < candidates.length; c++) {
      const cand = candidates[c];
      const status = String(data[cand.rowIndex][cand.statusIdx] || '').trim().toUpperCase();
      if (status === 'OPEN') {
        target = cand;
        break;
      }
    }

    if (!target) {
      const statuses = candidates.map(function (idx) {
        return String(data[idx.rowIndex][idx.statusIdx] || '').trim().toUpperCase() || 'UNKNOWN';
      });
      throw new Error('Cannot change order line. Only OPEN lines can be approved/rejected. Current: ' + statuses.join(', '));
    }

    sheet.getRange(target.rowIndex + 1, target.statusIdx + 1).setValue(newStatus);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, target.rowIndex + 1, target.statusIdx + 1, newStatus);

    const now = new Date();
    sheet.getRange(target.rowIndex + 1, target.updatedIdx + 1).setValue(now);
    _updateSheetCacheCell(CONFIG.SHEETS.PRODUCTION_LEDGER, target.rowIndex + 1, target.updatedIdx + 1, now);

    _clearScriptCachesForSheet(CONFIG.SHEETS.PRODUCTION_LEDGER);
    return { success: true, message: `Production Order ${newStatus}`, orderRef: targetOrderRef };
  });
}

/**
 * Validates that a Production Order is ready before allowing consumption/return.
 * APPROVED is the standard status. Legacy OPEN rows are treated as ready so
 * older production orders continue to work after approval removal.
 * @param {String} orderRef - Production Order reference ID
 * @throws {Error} If PO doesn't exist or has no ready lines
 */
function _validateProductionOrderApproved(orderRef) {
  const cols = CONFIG.PRODUCTION_LEDGER_COLS;
  const rows = _getProductionLedgerOrderRows(orderRef);

  let found = false;
  let hasReadyLine = false;
  const statuses = {};

  for (let i = 0; i < rows.length; i++) {
    found = true;
    const row = rows[i].row;
    const statusIdx = _detectProductionLedgerStatusIdx(row, cols.STATUS);
    const status = String(row[statusIdx] || '').trim().toUpperCase();
    statuses[status || 'UNKNOWN'] = true;
    if (status === 'APPROVED' || status === 'OPEN') hasReadyLine = true;
  }

  if (!found) throw new Error(`Production Order ${orderRef} not found`);
  if (!hasReadyLine) {
    const list = Object.keys(statuses).join(', ');
    throw new Error(`Production Order ${orderRef} is not ready for issue/return. Current status: ${list}.`);
  }
}

function _resolveInventoryItemCode(row, itemCodeCache, itemId, itemIdToCode) {
  const rawCode = String(itemCodeCache || '').trim();
  const rawId = String(itemId || '').trim();
  if (rawId) {
    // Modern rows: ITEM_ID is present. If cache differs from ID, trust cache as code.
    if (rawCode && rawCode.toUpperCase() !== rawId.toUpperCase()) return rawCode;

    if (!itemIdToCode[rawId]) {
      itemIdToCode[rawId] = String(_getItemCodeById(rawId) || '').trim();
    }
    return itemIdToCode[rawId] || rawCode;
  }

  // Legacy rows fallback: ITEM_ID is blank; ITEM_CODE_CACHE may contain either code or ID.
  if (!rawCode) return '';

  const codeProbeKey = '__CODE__' + rawCode.toUpperCase();
  if (!(codeProbeKey in itemIdToCode)) {
    const byCode = getItemByCodeCached(rawCode);
    itemIdToCode[codeProbeKey] = byCode ? String(byCode.code || '').trim() : '';
  }
  if (itemIdToCode[codeProbeKey]) return itemIdToCode[codeProbeKey];

  const idProbeKey = '__ID__' + rawCode.toUpperCase();
  if (!(idProbeKey in itemIdToCode)) {
    itemIdToCode[idProbeKey] = String(_getItemCodeById(rawCode) || '').trim();
  }
  return itemIdToCode[idProbeKey] || rawCode;
}

function getWarehouseLayout() {
  return _withRequestCache(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
    const rackSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.RACK);
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const itemSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ITEM); // ADD THIS

    // --- Existing Rack Map (unchanged) ---
    const rackData = _getSheetValuesCached(rackSheet.getName());
    const rackMap = {};
    for (let i = 1; i < rackData.length; i++) {
      const rId = String(rackData[i][0]);
      rackMap[rId] = { id: rId, code: rackData[i][1], site: rackData[i][2], location: rackData[i][3] };
    }

    // --- NEW: Item Map (ID -> { code, name, uom }) ---
    const itemData = _getSheetValuesCached(itemSheet.getName());
    const itemCols = _getItemSheetColumnMap(itemSheet);
    const itemMap = {};
    for (let i = 1; i < itemData.length; i++) {
      const iId = String(itemData[i][itemCols.ITEM_ID] || '').trim();
      const code = String(itemData[i][itemCols.ITEM_CODE] || '').trim();
      const name = String(itemData[i][itemCols.ITEM_NAME] || '').trim();
      const uom = String(itemData[i][itemCols.UOM_CODE] || 'KG').trim().toUpperCase() || 'KG';
      if (iId) {
        itemMap[iId] = { code: code, name: name, uomCode: uom };
      }
      if (code) {
        const codeKey = '__CODE__' + code.toUpperCase();
        itemMap[codeKey] = { code: code, name: name, uomCode: uom };
      }
    }

    // --- Existing Bin Map (unchanged) ---
    const binData = _getSheetValuesCached(binSheet.getName());
    const binHeaders = _getSheetHeaderMap(binSheet);
    const binMetaMap = _buildBinMasterMetaMap(binSheet, binData, binHeaders);
    const usageSummary = _getInventoryUsageSummaryByBin(binMetaMap);
    const bins = {};
    for (let i = 1; i < binData.length; i++) {
      const bId = String(binData[i][0]);
      const rId = String(binData[i][1]);
      const rack = rackMap[rId];
      if (rack) {
        const meta = binMetaMap[bId] || {};
        const usage = usageSummary.byBin[bId] || { byUom: {}, hasInventory: false };
        const metrics = _getBinCapacityMetrics(meta, usage);
        bins[bId] = {
          id: bId, code: binData[i][2], rackCode: rack.code,
          site: rack.site, location: rack.location,
          maxCapacity: Number(metrics.maxCapacity || 0),
          capacityUom: _normalizeCapacityUom(meta.declaredCapacityUom || meta.capacityUom || metrics.capacityUom || 'KG'),
          capacities: meta.capacities || {},
          currentUsage: Number(metrics.currentUsage || 0),
          capacityLines: metrics.capacityLines || [],
          utilizationPct: Number(metrics.utilizationPct || 0),
          capacityWarning: String(metrics.capacityWarning || ''),
          items: []
        };
      } else {
        Logger.log('Orphan Bin: ' + bId + ' -> Rack ' + rId);
      }
    }

    // --- EXTENDED Inventory Aggregation ---
    const invData = _getSheetValuesCached(invSheet.getName());
    for (let i = 1; i < invData.length; i++) {
      const row = invData[i];
      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID]);
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]);
      if (!bins[binId] || qty <= 0) continue;

      // NEW: Collect item info per bin row
      const itemId = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim();
      const rawCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
      const itemMeta = itemMap[itemId] || itemMap['__CODE__' + rawCode.toUpperCase()] || { code: rawCode || itemId, name: 'Unknown', uomCode: 'KG' };
      const batchRef = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
      const batchInfo = _resolveBatchReference(itemMeta.code, batchRef);
      const batchId = String((batchInfo && batchInfo.batchId) || batchRef).trim();
      const qaStatus = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').trim();
      const rowUom = String((CONFIG.INVENTORY_COLS.UOM !== undefined ? row[CONFIG.INVENTORY_COLS.UOM] : '') || '').trim();

      // Merge rows for same item+batch in same bin (can be split across multiple inv rows)
      const version = String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim().toUpperCase();
      const existingIdx = bins[binId].items.findIndex(
        it => it.itemId === itemId && it.batchId === batchId && it.version === version
      );
      if (existingIdx >= 0) {
        bins[binId].items[existingIdx].qty += qty;
      } else {
        bins[binId].items.push({
          itemId: itemId,
          itemCode: itemMeta.code,
          itemName: itemMeta.name,
          uomCode: rowUom || itemMeta.uomCode || 'KG',
          batchId: batchId,
          batchNumber: String((batchInfo && batchInfo.batchNumber) || batchRef || '').trim(),
          version: version,
          qty: qty,
          qaStatus: qaStatus
        });
      }
    }

    // --- NEW: Inject Packing Version 'Latest' Status ---
    const packingMappings = {};
    const pmSheet = ss.getSheetByName(CONFIG.SHEETS.PACKING_VERSION_MASTER);
    if (pmSheet) {
      const pmData = pmSheet.getDataRange().getValues();
      const C = CONFIG.PACKING_VERSION_COLS;
      for (let i = 1; i < pmData.length; i++) {
        const isLatest = pmData[i][C.IS_LATEST] === true || String(pmData[i][C.IS_LATEST]).toUpperCase() === 'TRUE';
        if (isLatest) {
          const key = `${String(pmData[i][C.ITEM_CODE]).toUpperCase()}||${String(pmData[i][C.SYSTEM_VERSION]).toUpperCase()}||${String(pmData[i][C.BIN_ID]).toUpperCase()}`;
          packingMappings[key] = String(pmData[i][C.REASON] || 'Latest');
        }
      }
    }

    Object.keys(bins).forEach(bid => {
      bins[bid].items.forEach(it => {
        const key = `${String(it.itemCode).toUpperCase()}||${String(it.version).toUpperCase()}||${String(bid).toUpperCase()}`;
        if (packingMappings[key]) {
          it.isLatest = true;
          it.latestReason = packingMappings[key];
        }
      });
    });

    return Object.values(bins).map(function (b) {
      let usageUom = String(b.capacityUom || 'KG').trim().toUpperCase() || 'KG';
      return {
        ...b,
        usageUom: usageUom,
        available: Number((((b.capacityLines || [])[0] || {}).availableCapacity) || Math.max(0, Number(b.maxCapacity || 0) - Number(b.currentUsage || 0)))
      };
    });
  });
}

// =====================================================
// 8. WEB APP ROUTING (doGet)
// =====================================================

function doGet(e) {
  return _withRequestCache(function () {
    let user;
    if (GLOBAL_DEV_BYPASS) {
      user = { email: 'kadambari.purohit@cultivator.in', role: SECURITY.ROLES.MANAGER, fullName: 'System Owner (Bypass)' };
    } else {
      try { user = getEffectiveUser(); } catch (err) {
        const t = HtmlService.createTemplateFromFile('Unauthorized');
        t.email = Session.getActiveUser().getEmail();
        return t.evaluate();
      }
    }

    const formType = e.parameter.form || 'dashboard';
    let htmlFile;

    if (formType === 'qa_approval' && !isQualityManager(user)) {
      const t = HtmlService.createTemplateFromFile('Unauthorized');
      t.email = user.email;
      return t.evaluate();
    }

    if (user.role === SECURITY.ROLES.WORKER) {
      switch (formType) {
        case 'inward': htmlFile = e.parameter.mode ? 'InwardForm' : 'InwardLanding'; break;
        case 'transfer': htmlFile = 'TransferForm'; break;
        case 'dispatch': htmlFile = 'DispatchForm'; break;
        case 'production_request': htmlFile = 'ProductionRequestForm'; break;
        case 'warehouse': htmlFile = 'WarehouseLayout'; break;
        case 'inventory_dashboard': htmlFile = 'WarehouseItemDashboard'; break;
        case 'packing_dashboard': htmlFile = 'PackingMaterialDashboard'; break;
        case 'dashboard': htmlFile = 'WorkerDashboard'; break;
        default: htmlFile = 'WorkerDashboard';
      }
    }

    else if (user.role === SECURITY.ROLES.MANAGER) {
      switch (formType) {
        case 'inward': htmlFile = e.parameter.mode ? 'InwardForm' : 'InwardLanding'; break;
        case 'transfer': htmlFile = 'TransferForm'; break;
        case 'dispatch': htmlFile = 'DispatchForm'; break;
        case 'warehouse': htmlFile = 'WarehouseLayout'; break;
        case 'inventory_dashboard': htmlFile = 'WarehouseItemDashboard'; break;
        case 'packing_dashboard': htmlFile = 'PackingMaterialDashboard'; break;
        case 'dashboard': htmlFile = 'ManagerDashboard'; break;
        case 'production_request': htmlFile = 'ProductionRequestForm'; break;
        case 'production_approval': htmlFile = 'ProductionApprovalForm'; break;
        case 'qa_approval': htmlFile = 'QAApprovalForm'; break;
        case 'bulk_upload': htmlFile = 'BulkUploadMaster'; break;
        case 'reports':
          htmlFile = 'ReportsScreen';
          break;

        case 'ai_assistant': htmlFile = 'AIAssistant'; break;
        default: htmlFile = 'ManagerDashboard';
      }
    }

    else if (user.role === SECURITY.ROLES.QUALITY_MANAGER) {
      switch (formType) {
        case 'qa_approval': htmlFile = 'QAApprovalForm'; break;
        case 'warehouse': htmlFile = 'WarehouseLayout'; break;
        case 'inventory_dashboard': htmlFile = 'WarehouseItemDashboard'; break;
        case 'packing_dashboard': htmlFile = 'PackingMaterialDashboard'; break;
       case 'production_approval': htmlFile = 'ProductionApprovalForm'; break;
        case 'ai_assistant': htmlFile = 'AIAssistant'; break;
        case 'dashboard': htmlFile = 'QADashboard'; break;
        case 'bulk_upload': htmlFile = 'BulkUploadMaster'; break;
        case 'reports':
          htmlFile = 'ReportsScreen';
          break;

        default: htmlFile = 'QADashboard';
      }
    }
    else {
      htmlFile = 'Unauthorized';
    }

    const template = HtmlService.createTemplateFromFile(htmlFile);
    template.userEmail = user.email;
    template.userRole = user.role;
    template.inwardMode = String((e.parameter && e.parameter.mode) || '').trim().toLowerCase();
    if (htmlFile === 'Unauthorized') {
      template.email = user.email;
    }

    return template.evaluate()
      .setTitle('WMS - ' + formType.toUpperCase())
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  });
}

// Manual edit audit (sheet-level tamper logging). Uses QA_Events as audit sink.
function onEdit(e) {
  return _withRequestCache(function () {
    try {
      if (!e || !e.range) return;
      const sheet = e.range.getSheet();
      const sheetName = sheet.getName();
      _clearScriptCachesForSheet(sheetName);
      if (sheetName === CONFIG.SHEETS.QA_EVENTS || sheetName === CONFIG.SHEETS.MOVEMENT) return;
      if (sheetName !== CONFIG.SHEETS.INVENTORY && sheetName !== CONFIG.SHEETS.BATCH) return;
      const row = e.range.getRow();
      if (row <= 1) return; // skip header

      let inventoryId = '';
      let itemCode = '';
      let batchId = '';
      let binId = '';
      if (sheetName === CONFIG.SHEETS.INVENTORY) {
        const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        inventoryId = String(rowData[CONFIG.INVENTORY_COLS.INVENTORY_ID] || '');
        itemCode = String(rowData[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '');
        batchId = String(rowData[CONFIG.INVENTORY_COLS.BATCH_ID] || '');
        binId = String(rowData[CONFIG.INVENTORY_COLS.BIN_ID] || '');
      } else if (sheetName === CONFIG.SHEETS.BATCH) {
        const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        itemCode = String(rowData[CONFIG.BATCH_COLS.ITEM_CODE] || '');
        batchId = String(rowData[CONFIG.BATCH_COLS.BATCH_ID] || '');
      }

      const cell = e.range.getA1Notation();
      const oldVal = typeof e.oldValue === 'undefined' ? '' : String(e.oldValue);
      const newVal = typeof e.value === 'undefined' ? '' : String(e.value);
      const remarks = `Manual edit in ${sheetName} ${cell}: '${oldVal}' â†’ '${newVal}'`;
      _appendQaEventsBatch([{
        inventoryId: inventoryId,
        itemCode: itemCode,
        batchId: batchId,
        binId: binId,
        prevStatus: '',
        newStatus: '',
        action: 'TAMPER',
        overrideReason: '',
        overriddenBy: Session.getActiveUser().getEmail(),
        overriddenAt: new Date(),
        remarks: remarks
      }]);
    } catch (err) {
      Logger.log('onEdit audit failed: ' + err.message);
    }
  });
}

// =====================================================
// AUTO-VERSION LOOKUP (Read-only helper â€” Section 8 addendum)
// Paste this at the bottom of Code.js, before Section 9 if it exists.
// Does NOT modify any transactional function.
// =====================================================

/**
 * getNextVersionForItemCode
 *
 * Scans Batch_Master for every row matching the given item code,
 * finds the highest existing version number (V1, V2, V3â€¦),
 * and returns the NEXT version string.
 *
 * Called by InwardForm.html when the user finishes typing an item code.
 * READ-ONLY â€” touches no transactional state.
 *
 * @param  {string} itemCode  â€” Item code as typed by user (case-insensitive).
 * @returns {string}           â€” "V1" if never seen before, "V2", "V3", etc.
 */
function getNextVersionForItemCode(itemCode) {
  return _withRequestCache(function () {
    const code = _getCanonicalItemCode(itemCode).toUpperCase();
    if (!code) return 'V1';

    const cacheKey = 'INWARD_NEXT_VERSION_V4_' + code;
    const cached = _getScriptCacheJson(cacheKey);
    if (cached && cached.version) return cached.version;

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    const batchSheet = ss.getSheetByName(CONFIG.SHEETS.BATCH);

    if (!invSheet && !batchSheet) return 'V1';

    let maxVersion = 0;
    const itemMaps = _getItemMasterMaps();
    const codeNorm = code.toUpperCase();

    function scanInventory(data) {
      for (let i = 1; i < data.length; i++) {
        let rowCode = String(data[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
        if (!rowCode) {
          const itemId = String(data[i][CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();
          rowCode = String(itemMaps.idToCode[itemId] || '').trim().toUpperCase();
        }
        rowCode = String(itemMaps.codeDisplayByNorm[rowCode] || rowCode).trim().toUpperCase();
        if (rowCode !== codeNorm) continue;

        const raw = String(data[i][CONFIG.INVENTORY_COLS.VERSION] || '').trim().toUpperCase();
        const m = raw.match(/^V(\d+)$/);
        if (m) maxVersion = Math.max(maxVersion, Number(m[1]));
      }
    }

    function scanBatchMaster(data) {
      for (let i = 1; i < data.length; i++) {
        let rowCode = String(data[i][CONFIG.BATCH_COLS.ITEM_CODE] || '').trim().toUpperCase();
        rowCode = String(itemMaps.codeDisplayByNorm[rowCode] || rowCode).trim().toUpperCase();
        if (rowCode !== codeNorm) continue;

        const raw = String(data[i][CONFIG.BATCH_COLS.VERSION] || '').trim().toUpperCase();
        const m = raw.match(/^V(\d+)$/);
        if (m) maxVersion = Math.max(maxVersion, Number(m[1]));
      }
    }

    if (invSheet) {
      const invData = _getSheetValuesCached(invSheet.getName());
      scanInventory(invData);
    }

    if (batchSheet) {
      const batchData = _getSheetValuesCached(batchSheet.getName());
      scanBatchMaster(batchData);
    }

    const nextVersion = 'V' + (maxVersion + 1);
    _putScriptCacheJson(cacheKey, { version: nextVersion }, 60);
    return nextVersion;
  });
}

/**
 * Backward-compatible helper for InwardForm.
 * Returns both the next version and the item name for a given item code.
 *
 * @param {string} itemCode
 * @returns {{version:string,name:string,code:string}}
 */
function getInwardItemDetails(itemCode) {
  return _withRequestCache(function () {
    const code = _getCanonicalItemCode(itemCode);
    if (!code) {
      return { version: 'V1', name: '', code: '' };
    }

    const item = getItemByCodeCached(code);
    return {
      version: getNextVersionForItemCode(code) || 'V1',
      name: String((item && item.name) || code).trim(),
      code: code
    };
  });
}


// 10.1  VERSION WARNINGS  


/**
 * 
 * Returns a warning if an OLDER version of the same item still has stock,
 * so the worker knows they are consuming e.g. V2 while V1 is still on shelves.
 *
 * @param {string} itemCode      - Item code being consumed/dispatched
 * @param {string} currentVersion - Version currently selected (e.g. "V2")
 * @returns {{ message: string }} - Empty string if no warning needed
 */
function getVersionDispatchWarning(itemCode, currentVersion) {
  return protect(function () {
    if (!itemCode || !currentVersion) return { message: '' };

    const code = String(itemCode).trim().toUpperCase();
    const uomCode = _getItemUomCode(code);
    const current = String(currentVersion).trim().toUpperCase();
    const currentNum = _parseVersionNum(current);
    if (currentNum === null) return { message: '' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const invData = _getSheetValuesCached(invSheet.getName());

    const olderBins = {};

    for (var i = 1; i < invData.length; i++) {
      const row = invData[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      const rowCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
      if (rowCode !== code) continue;

      const rowVer = String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim().toUpperCase();
      const rowNum = _parseVersionNum(rowVer);
      if (rowNum === null || rowNum >= currentNum) continue;

      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      const binMeta = _getBinMetaById(binId);
      const binLabel = binMeta ? (binMeta.binCode || binId) : binId;

      if (!olderBins[rowVer]) olderBins[rowVer] = [];
      const existing = olderBins[rowVer].find(function (b) { return b.binId === binId; });
      if (existing) {
        existing.qty += qty;
      } else {
        olderBins[rowVer].push({ binId: binId, binLabel: binLabel, qty: qty });
      }
    }

    const versions = Object.keys(olderBins);
    if (versions.length === 0) return { message: '' };

    const summaries = versions.sort(function (a, b) {
      const an = _parseVersionNum(a) || 0;
      const bn = _parseVersionNum(b) || 0;
      return an - bn;
    }).map(function (ver) {
      const bins = olderBins[ver] || [];
      const total = bins.reduce(function (acc, b) { return acc + (Number(b.qty) || 0); }, 0);
      return { version: ver, binCount: bins.length, totalQty: total };
    });

    const preview = summaries.slice(0, 3).map(function (s) {
      return s.version + ' ' + s.totalQty.toFixed(2) + ' ' + uomCode + ' in ' + s.binCount + ' bin(s)';
    }).join(' | ');
    const extra = summaries.length > 3 ? ' +' + (summaries.length - 3) + ' more version(s)' : '';

    return {
      message: 'Older stock exists: ' + preview + extra + '. You are consuming ' + current + '.'
    };
  });
}

function getVersionInwardWarning(itemCode, nextVersion) {
  return protect(function () {
    if (!itemCode || !nextVersion) return { message: '', bins: [] };

    const code = String(itemCode).trim().toUpperCase();
    const cacheKey = 'INWARD_VERSION_WARN_V3_' + code + '_' + String(nextVersion || '').trim().toUpperCase();
    const cached = _getScriptCacheJson(cacheKey);
    if (cached) return cached;

    const uomCode = _getItemUomCode(code);
    const nextNum = _parseVersionNum(String(nextVersion).trim().toUpperCase());
    if (nextNum === null || nextNum <= 1) return { message: '', bins: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const invData = _getSheetValuesCached(invSheet.getName());
    const binMetaById = _buildBinMetaMap(ss);

    const existing = {};

    for (var i = 1; i < invData.length; i++) {
      const row = invData[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      const rowCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
      if (rowCode !== code) continue;

      const rowVer = String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim().toUpperCase();
      if (!rowVer || rowVer === 'UNSPECIFIED') continue;

      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      const binMeta = binMetaById[binId] || null;
      const binLabel = binMeta ? (binMeta.binCode || binId) : binId;

      if (!existing[rowVer]) existing[rowVer] = [];
      const found = existing[rowVer].find(function (b) { return b.binId === binId; });
      if (found) {
        found.qty += qty;
      } else {
        existing[rowVer].push({ binId: binId, binLabel: binLabel, qty: qty });
      }
    }

    const versions = Object.keys(existing);
    if (versions.length === 0) {
      const empty = { message: '', bins: [] };
      _putScriptCacheJson(cacheKey, empty, 45);
      return empty;
    }

    const summaries = versions.sort(function (a, b) {
      const an = _parseVersionNum(a) || 0;
      const bn = _parseVersionNum(b) || 0;
      return an - bn;
    }).map(function (ver) {
      const bins = existing[ver] || [];
      const total = bins.reduce(function (acc, b) { return acc + (Number(b.qty) || 0); }, 0);
      return {
        version: ver,
        binCount: bins.length,
        totalQty: total,
        bins: bins
      };
    });

    const preview = summaries.slice(0, 3).map(function (s) {
      return s.version + ' ' + s.totalQty.toFixed(2) + ' ' + uomCode + ' in ' + s.binCount + ' bin(s)';
    }).join(' | ');
    const extra = summaries.length > 3 ? ' +' + (summaries.length - 3) + ' more version(s)' : '';
    const details = summaries.map(function (s) {
      const topBins = (s.bins || []).slice(0, 4).map(function (b) {
        return b.binLabel + ' (' + Number(b.qty || 0).toFixed(2) + ' ' + uomCode + ')';
      }).join(', ');
      const moreBins = (s.bins || []).length > 4 ? ' + ' + ((s.bins || []).length - 4) + ' more bin(s)' : '';
      return s.version + ': ' + s.totalQty.toFixed(2) + ' ' + uomCode + ' in ' + s.binCount + ' bin(s)' +
        (topBins ? ' [' + topBins + moreBins + ']' : '');
    });

    const result = {
      message: 'Older stock on shelves: ' + preview + extra + '. Receiving ' + nextVersion + '.',
      bins: versions,
      details: details,
      totalOlderQty: summaries.reduce(function (acc, s) { return acc + (Number(s.totalQty) || 0); }, 0)
    };
    _putScriptCacheJson(cacheKey, result, 45);
    return result;
  });
}

function _parseVersionNum(v) {
  if (!v) return null;
  const m = String(v).trim().toUpperCase().match(/^V(\d+)$/);
  return m ? parseInt(m[1], 10) : null;
}


// 10.2  PO ITEMS LOOKUP  (for consumption auto-populate)

/**
 * Called by DispatchForm when user types a PO number in Consumption mode.
 * Returns the full item list for that ready PO so the form can auto-fill.
 *
 * Returns error objects instead of throwing â€” client checks .error field.
 *
 * @param {string} poNo - Production Order reference (e.g. "PO-20250215-042")
 * @returns {{ items: Array, poNo: string, status: string } | { error: string }}
 */
function getApprovedProductionOrderItems(poNo) {
  return protect(function () {
    return _withRequestCache(function () {
      if (!poNo) return { error: 'Production Order number required.' };

      const ref = String(poNo).trim();
      const cols = CONFIG.PRODUCTION_LEDGER_COLS;
      const rows = _getProductionLedgerOrderRows(ref);

      var poStatus = null;
      var matchedRows = [];

      for (var i = 0; i < rows.length; i++) {
        const row = rows[i].row;
        const statusIdx = _detectProductionLedgerStatusIdx(row, cols.STATUS);
        const status = String(row[statusIdx] || '').trim().toUpperCase();
        poStatus = status; // last seen status for this PO (all rows should match)

        if (status !== 'APPROVED' && status !== 'OPEN') continue;

        const itemId = String(row[cols.ITEM_ID] || '').trim();
        const itemCode = _getItemCodeById(itemId) || itemId;
        const ledgerBatch = String(row[cols.BATCH_ID] || '').trim();
        const jobOrderNo = (ledgerBatch && ledgerBatch.toUpperCase() !== 'TBD') ? ledgerBatch : '';
        const version = String(row[cols.VERSION] || '').trim();
        const requested = Number(row[cols.QTY_REQUESTED]) || 0;
        const issued = Number(row[cols.QTY_ISSUED]) || 0;
        const returned = Number(row[cols.QTY_RETURNED]) || 0;
        const rejected = Number(row[cols.QTY_REJECTED]) || 0;
        const netOutstandingRaw = Number(row[cols.NET_OUTSTANDING]);
        const remaining = isFinite(netOutstandingRaw) ? Math.max(0, netOutstandingRaw) : Math.max(0, requested - issued + returned);

        matchedRows.push({
          ledgerRowIndex: rows[i].rowIndex,
          itemId: itemId,
          itemCode: itemCode,
          uomCode: _getItemUomCode(itemCode),
          batchId: (ledgerBatch.toUpperCase() === 'TBD' ? '' : ledgerBatch),
          jobOrderNo: jobOrderNo,
          version: version,
          requested: requested,
          issued: issued,
          returned: returned,
          rejected: rejected,
          remaining: remaining
        });
      }

      if (matchedRows.length === 0) {
        if (poStatus && poStatus !== 'APPROVED' && poStatus !== 'OPEN') {
          return { error: 'PO ' + ref + ' is not ready yet. Current status: ' + poStatus + '.' };
        }
        return { error: 'Production Order "' + ref + '" not found.' };
      }

      // Filter out items already fully issued
      const pendingItems = matchedRows.filter(function (r) { return r.remaining > 0; });
      if (pendingItems.length === 0) {
        return { error: 'All items for PO ' + ref + ' have already been fully issued.' };
      }

      return {
        poNo: ref,
        status: 'READY',
        items: pendingItems
      };
    });
  });
}

function _buildConsumptionSourceKey(batchId, binId) {
  return encodeURIComponent(String(batchId || '')) + '||' + encodeURIComponent(String(binId || ''));
}

function _getConsumptionSourceOptionsForItems(items) {
  const wantedCodes = {};
  const wantedIds = {};
  (items || []).forEach(function (item) {
    const code = String(item && item.itemCode || '').trim().toUpperCase();
    const id = String(item && item.itemId || '').trim().toUpperCase();
    if (code) wantedCodes[code] = true;
    if (id) wantedIds[id] = true;
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const data = _getSheetValuesCached(invSheet.getName());
  const binMeta = _buildBinMetaMap(ss);
  const itemIdToCode = {};
  const grouped = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
    if (qty <= 0) continue;

    const resolvedCode = String(_resolveInventoryItemCode(
      row,
      row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE],
      row[CONFIG.INVENTORY_COLS.ITEM_ID],
      itemIdToCode
    ) || '').trim();
    const codeU = resolvedCode.toUpperCase();
    const rowIdU = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();
    if (!wantedCodes[codeU] && !wantedIds[rowIdU]) continue;

    const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
    if (!binId) continue;

    const batchInfo = _resolveBatchReference(resolvedCode, row[CONFIG.INVENTORY_COLS.BATCH_ID]);
    const batchId = String((batchInfo && batchInfo.batchId) || row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
    const qa = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || '').trim().toUpperCase();
    const key = codeU;
    const optionKey = _buildConsumptionSourceKey('', binId);
    const meta = binMeta[binId] || {};

    if (!grouped[key]) grouped[key] = {};
    if (!grouped[key][optionKey]) {
      grouped[key][optionKey] = {
        optionKey: optionKey,
        binId: binId,
        batchId: '',
        binCode: String(meta.binCode || binId).trim(),
        displayLabel: String(meta.displayLabel || meta.binCode || binId).trim(),
        capacityUom: String(meta.capacityUom || '').trim(),
        site: String(row[CONFIG.INVENTORY_COLS.SITE] || meta.site || '').trim(),
        location: String(row[CONFIG.INVENTORY_COLS.LOCATION] || meta.location || '').trim(),
        totalQty: 0,
        approvedQty: 0,
        pendingQty: 0,
        rejectedQty: 0,
        batchBuckets: {}
      };
    }

    const opt = grouped[key][optionKey];
    opt.totalQty += qty;
    if (qa === 'APPROVED' || qa === 'OVERRIDDEN') opt.approvedQty += qty;
    else if (qa === 'REJECTED' || qa === 'HOLD') opt.rejectedQty += qty;
    else opt.pendingQty += qty;

    const bucketKey = batchId || '__NO_BATCH__';
    if (!opt.batchBuckets[bucketKey]) {
      opt.batchBuckets[bucketKey] = { batchId: batchId, totalQty: 0, approvedQty: 0 };
    }
    opt.batchBuckets[bucketKey].totalQty += qty;
    if (qa === 'APPROVED' || qa === 'OVERRIDDEN') {
      opt.batchBuckets[bucketKey].approvedQty += qty;
    }
  }

  const out = {};
  Object.keys(grouped).forEach(function (codeU) {
    out[codeU] = Object.keys(grouped[codeU]).map(function (k) {
      const opt = grouped[codeU][k];
      opt.batchBuckets = Object.keys(opt.batchBuckets).map(function (bk) {
        return opt.batchBuckets[bk];
      }).sort(function (a, b) {
        if (b.approvedQty !== a.approvedQty) return b.approvedQty - a.approvedQty;
        return b.totalQty - a.totalQty;
      });
      return opt;
    }).sort(function (a, b) {
      if (b.approvedQty !== a.approvedQty) return b.approvedQty - a.approvedQty;
      return b.totalQty - a.totalQty;
    });
  });
  return out;
}

function getConsumptionLoadData(poNo) {
  return protect(function () {
    return _withRequestCache(function () {
      const po = getApprovedProductionOrderItems(poNo);
      if (!po || po.error) return po;

      const sourceOptionsByItem = _getConsumptionSourceOptionsForItems(po.items || []);
      po.items = (po.items || []).map(function (item) {
        const codeU = String(item && item.itemCode || '').trim().toUpperCase();
        const copy = Object.assign({}, item);
        copy.sourceOptions = sourceOptionsByItem[codeU] || [];
        return copy;
      });
      return po;
    });
  });
}

/**
 * Debug helper for PO consumption source resolution.
 * Use from Apps Script console:
 *   Logger.log(JSON.stringify(debugPoConsumptionSources('PO-YYYYMMDD-XXXX'), null, 2));
 */
function debugPoConsumptionSources(poNo) {
  return protect(function () {
    const po = getApprovedProductionOrderItems(poNo);
    if (!po || po.error) return po || { error: 'PO lookup failed' };

    const inv = getInventoryReadView() || [];
    const out = (po.items || []).map(function (it) {
      const codeU = String(it.itemCode || '').trim().toUpperCase();
      const idRaw = String(it.itemId || '').trim();

      let matchedRows = 0;
      let qtyRows = 0;
      let usableRows = 0;
      let totalQty = 0;
      let missingBatchRows = 0;
      let missingBinRows = 0;

      inv.forEach(function (r) {
        const matchByCode = String(r.itemCode || '').trim().toUpperCase() === codeU;
        const matchById = String(r.itemId || '').trim() === idRaw;
        if (!matchByCode && !matchById) return;

        matchedRows += 1;
        const qty = Number(r.quantity) || 0;
        if (qty <= 0) return;
        qtyRows += 1;

        const hasBatch = !!String(r.batchId || '').trim();
        const hasBin = !!String(r.binId || '').trim();
        if (!hasBatch) { missingBatchRows += 1; return; }
        if (!hasBin) { missingBinRows += 1; return; }

        usableRows += 1;
        totalQty += qty;
      });

      return {
        itemCode: it.itemCode,
        itemId: it.itemId,
        requested: Number(it.requested || 0),
        remaining: Number(it.remaining || 0),
        matchedRows: matchedRows,
        qtyRows: qtyRows,
        usableRows: usableRows,
        totalQty: totalQty,
        missingBatchRows: missingBatchRows,
        missingBinRows: missingBinRows
      };
    });

    return { poNo: String(poNo || '').trim(), items: out };
  });
}


// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// 10.3  QTY CAP ENFORCEMENT  (BUG-2 FIX)
//
// Add ONE call to this function inside submitConsumptionV2,
// just BEFORE the call to _updateProductionLedger.
//
// Find this line in submitConsumptionV2 (around line 3380):
//   ledgerChange = _updateProductionLedger(form.prodOrderNo, lookup, item.batchId, itemVersion, qtyKg, 0);
//
// Add ONE line above it:
//   _assertConsumptionWithinBudget(form.prodOrderNo, lookup, item.batchId, qtyKg);
//
// That is the ONLY change to submitConsumptionV2.
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

/**
 * Throws if issuing qtyKg for this PO + item would exceed the requested qty.
 * Called inside submitConsumptionV2, inside the lock, before ledger update.
 *
 * @param {string} orderRef
 * @param {string} itemId
 * @param {string} batchId
 * @param {number} qtyKgToIssue
 */
function _assertConsumptionWithinBudget(orderRef, itemId, batchId, qtyKgToIssue) {
  const state = _getProductionLedgerState(orderRef, itemId, batchId, '');

  // If no ledger row exists yet, it will be auto-created â€” no cap to enforce
  if (state.rowIndex === -1) return;

  const requested = Number(state.requested) || 0;
  // If QTY_REQUESTED was never set (old POs), skip enforcement
  if (requested === 0) return;

  const alreadyIssued = Number(state.issued) || 0;
  const returned = Number(state.returned) || 0;
  const remaining = requested - alreadyIssued + returned;

  if (qtyKgToIssue > remaining + 0.001) { // 0.001 tolerance for float rounding
    throw new Error(
      'Cannot issue ' + qtyKgToIssue.toFixed(3) + ' KG for ' + orderRef + ' / ' + (itemId || batchId) + '.' +
      ' Requested: ' + requested.toFixed(3) + ' KG, already issued: ' + alreadyIssued.toFixed(3) + ' KG, remaining: ' + remaining.toFixed(3) + ' KG.'
    );
  }
}




// 10.5  STOCK SUMMARY HELPER  (called by ProductionRequestForm advisory check)
//
// READ-ONLY. Returns approved + pending qty for an item code.
// Does not throw â€” returns null if item not found.

/**
 * @param {string} itemCode
 * @returns {{ approved: number, pending: number } | null}
 */
function getItemStockSummary(itemCode) {
  return protect(function () {
    if (!itemCode) return null;
    const code = String(itemCode).trim().toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    if (!invSheet) return null;

    const data = _getSheetValuesCached(invSheet.getName());
    var approved = 0;
    var pending = 0;

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;
      const rowCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
      if (rowCode !== code) continue;
      const status = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || '').trim().toUpperCase();
      if (status === 'APPROVED' || status === 'OVERRIDDEN') {
        approved += qty;
      } else {
        pending += qty;
      }
    }

    if (approved === 0 && pending === 0) return null;
    return { approved: approved, pending: pending };
  });
}

// ============================================================================
// COPY THIS ENTIRE BLOCK TO Code.js
// Paste AFTER getItemStockSummary function (around line 5288)
// ============================================================================

/**
 * NEW FUNCTION - MISSING FROM YOUR CODE
 * Returns all batches with stock for a given item code
 * 
 * @param {string} itemCode
 * @returns {Array<{batchId:string, itemId:string, approved:number, pending:number, total:number, binCount:number}>}
 */
function getItemBatchStockSummary(itemCode) {
  return protect(function () {
    if (!itemCode) return [];
    const code = String(itemCode).trim().toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    if (!invSheet) return [];

    const data = _getSheetValuesCached(invSheet.getName());
    const batchMap = {};

    // Scan inventory for all batches with this item code
    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;

      const rowCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
      if (rowCode !== code) continue;

      const batchRef = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
      if (!batchRef) continue;
      const batchId = String((_resolveBatchReference(code, batchRef).batchId) || batchRef).trim();

      const status = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || '').trim().toUpperCase();
      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();

      // CRITICAL: Get item_id from inventory row (not Item_Master)
      const rowItemId = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim();

      if (!batchMap[batchId]) {
        batchMap[batchId] = {
          batchId: batchId,
          itemId: rowItemId,
          approved: 0,
          pending: 0,
          total: 0,
          binSet: {},
          fifoTs: Number.POSITIVE_INFINITY,
          fifoGin: '',
          fifoInvTs: Number.POSITIVE_INFINITY
        };
      }

      if (status === 'APPROVED' || status === 'OVERRIDDEN') {
        batchMap[batchId].approved += qty;
      } else {
        batchMap[batchId].pending += qty;
      }
      batchMap[batchId].total += qty;

      if (binId) {
        batchMap[batchId].binSet[binId] = true;
      }

      const qdTs = _toTimestamp(row[CONFIG.INVENTORY_COLS.QUALITY_DATE], NaN);
      const invTs = _getInventoryCreatedTs(row);
      const gin = String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim();
      const candidateTs = isFinite(qdTs) ? qdTs : (isFinite(invTs) ? invTs : Number.POSITIVE_INFINITY);
      const currentTs = Number(batchMap[batchId].fifoTs);
      if (candidateTs < currentTs ||
        (candidateTs === currentTs && gin && (!batchMap[batchId].fifoGin || gin.localeCompare(batchMap[batchId].fifoGin) < 0))) {
        batchMap[batchId].fifoTs = candidateTs;
        batchMap[batchId].fifoGin = gin;
        batchMap[batchId].fifoInvTs = isFinite(invTs) ? invTs : Number.POSITIVE_INFINITY;
      }
    }

    // Convert map to array
    const result = [];
    Object.keys(batchMap).forEach(function (batchId) {
      const batch = batchMap[batchId];
      result.push({
        batchId: batch.batchId,
        batchNumber: _getBatchDisplayNumber(code, batch.batchId),
        itemId: batch.itemId,
        approved: batch.approved,
        pending: batch.pending,
        total: batch.total,
        binCount: Object.keys(batch.binSet).length,
        fifoTs: batch.fifoTs,
        fifoGin: batch.fifoGin,
        fifoInvTs: batch.fifoInvTs
      });
    });

    // FIFO first: oldest quality date / inventory anchor first, not largest stock.
    result.sort(function (a, b) {
      const at = Number(a.fifoTs);
      const bt = Number(b.fifoTs);
      const aHas = isFinite(at);
      const bHas = isFinite(bt);
      if (aHas && bHas && at !== bt) return at - bt;
      if (aHas !== bHas) return aHas ? -1 : 1;
      const ag = String(a.fifoGin || '');
      const bg = String(b.fifoGin || '');
      if (ag && bg && ag !== bg) return ag.localeCompare(bg);
      if (!!ag !== !!bg) return ag ? -1 : 1;
      return String(a.batchNumber || a.batchId || '').localeCompare(String(b.batchNumber || b.batchId || ''));
    });

    return result;
  });
}

// =====================================================
// SECTION 11: DATA MAPPING FIXES
// Paste at bottom of Code.js, after Section 10.
//
// Fixes:
//   FIX-1  Return Ref / Note now passes through and is saved
//   FIX-2  Production request remarks saved to Movement log
//   FIX-3  Lot No saved to Inventory (new column)
//   FIX-4  Expiry Date saved to Inventory (new column)
//   FIX-5  initRackInventory - populate bins from physical stock
//
// Schema changes required in your Google Sheet:
//   Inventory sheet:     add col N (index 13) = LOT_NO
//                        add col O (index 14) = EXPIRY_DATE
//   (All existing cols 0-12 stay in place â€” just add two at the right)
//
// After adding those columns, add these two lines to CONFIG.INVENTORY_COLS:
//   LOT_NO:      13,
//   EXPIRY_DATE: 14,
// =====================================================



// 11.1  FIXED submitReturnV2
//
// Changes vs original:
//   a) Accepts form.refNote and uses it as the GIN_NO (user-entered ref)
//      Falls back to 'RET-' + prodOrderNo if blank (backward compatible)
//   b) Sequential multi-item returns â€” no parallel lock contention
//      (InwardForm now calls submitReturnBatch instead)


/**
 * Single-item production return.
 * NOW ACCEPTS: form.refNote  (the "Return Ref / Note" field from InwardForm)
 *
 * @param {Object} form
 *   prodOrderNo {string}
 *   refNote     {string}   â† NEW â€” Return Ref / Note entered by operator
 *   itemCode    {string}
 *   batchId     {string}   or batchNumber
 *   quantity    {number}
 *   binId       {string}
 *   site        {string}
 *   location    {string}
 *   remarks     {string}
 *   lotNo       {string}
 *   qualityStatus {string} optional, default PENDING
 *   qualityDate {string}   optional
 */
function submitReturnV2Fixed(form) {
  return withScriptLock(function () {
    protect(function () { return requireOperationalUser(); });

    if (!form.prodOrderNo) throw new Error('Production Order No required');

    const batchId = String(form.batchId || form.batchNumber || '').trim();
    if (!batchId) throw new Error('Batch No required for return');

    const ids = _assertItemCodeBatch(form.itemCode, batchId, 'Return');
    const qtyKg = _convertToKg(ids.itemCode, form.quantity, form.uom);
    const lookup = _getValidatedItemId(ids.itemCode);
    const isDeadStock = form.isDeadStock === true;
    const deadStockReason = String(form.deadStockReason || '').trim();
    if (isDeadStock && !deadStockReason) {
      throw new Error('Dead stock reason required');
    }

    const ledger = _getProductionLedgerState(form.prodOrderNo, lookup, ids.batchId, '');
    const outstanding = ledger.issued - ledger.returned - (ledger.rejected || 0);
    if (qtyKg > outstanding + 0.001) {
      throw new Error('Cannot return ' + qtyKg.toFixed(3) + ' KG. Max outstanding: ' + outstanding.toFixed(3) + ' KG');
    }
    if (!ledger.version) throw new Error('Version required for return â€” no ledger row found for this PO + item + batch');

    const ginForInventory = String(form.refNote || '').trim() || ('RET-' + form.prodOrderNo);
    if (!isDeadStock) {
      _assertBinCapacity(form.binId, form.quantity, null, ids.itemCode, form.uom);
    }

    let created = null;
    const logs = [{
      type: isDeadStock ? CONFIG.MOVEMENT_TYPES.RETURN_DEAD : CONFIG.MOVEMENT_TYPES.RETURN_PROD,
      itemId: lookup,
      batchId: ids.batchId,
      version: ledger.version,
      ginNo: ginForInventory,
      prodOrderRef: form.prodOrderNo,
      toBinId: isDeadStock ? '' : form.binId,
      quantity: qtyKg,
      remarks: isDeadStock
        ? ('DEAD_STOCK: ' + deadStockReason + (form.remarks ? ' | ' + form.remarks : ''))
        : form.remarks
    }];

    if (!isDeadStock) {
      created = _createInventoryRowV2({
        itemId: lookup,
        itemCode: ids.itemCode,
        batchId: ids.batchId,
        ginNo: ginForInventory,
        version: ledger.version,
        qualityStatus: _normalizeQaStatus(form.qualityStatus || 'PENDING'),
        qualityDate: form.qualityDate ? new Date(form.qualityDate) : '',
        binId: form.binId,
        site: form.site,
        location: form.location,
        quantity: qtyKg,
        lotNo: String(form.lotNo || ''),
        expiryDate: form.expiryDate ? new Date(form.expiryDate) : '',
        returnRowIndex: true
      });
    }

    let ledgerChange = null;
    let prodTransferChange = null;
    let deadStockChange = null;
    const changes = created ? [{ action: 'create', rowIndex: created.rowIndex }] : [];

    try {
      const binMeta = _getBinMetaById(form.binId);
      prodTransferChange = _appendProductionTransfersBatch([{
        transferType: isDeadStock ? 'RETURN_DEAD_STOCK' : 'RETURN',
        productionOrderNo: form.prodOrderNo,
        productionArea: String(form.productionArea || 'PRODUCTION').trim(),
        itemCode: ids.itemCode,
        itemId: lookup,
        batchNumber: batchId,
        batchId: ids.batchId,
        lotNumber: form.lotNo || '',
        quantity: qtyKg,
        uom: _getItemUomCode(ids.itemCode),
        fromSite: String(form.productionArea || 'PRODUCTION').trim(),
        fromLocation: String(form.productionArea || 'PRODUCTION').trim(),
        fromBinId: String(form.productionArea || 'PRODUCTION').trim(),
        fromBinCode: String(form.productionArea || 'PRODUCTION').trim(),
        toSite: isDeadStock ? 'DEAD_STOCK' : String(form.site || ''),
        toLocation: isDeadStock ? 'DEAD_STOCK' : String(form.location || ''),
        toBinId: isDeadStock ? '' : String(form.binId || ''),
        toBinCode: isDeadStock ? 'DEAD_STOCK' : (binMeta.binCode || form.binId),
        returnedByName: Session.getActiveUser().getEmail(),
        createdBy: Session.getActiveUser().getEmail(),
        createdAt: new Date(),
        remarks: String(form.refNote ? '[Ref: ' + form.refNote + '] ' : '') +
          (isDeadStock ? ('[Dead Stock: ' + deadStockReason + '] ') : '') + String(form.remarks || ''),
        status: 'COMPLETED'
      }]);

      ledgerChange = _updateProductionLedger(
        form.prodOrderNo,
        lookup,
        ids.batchId,
        ledger.version,
        0,
        isDeadStock ? 0 : qtyKg,
        isDeadStock ? qtyKg : 0
      );

      if (isDeadStock) {
        deadStockChange = _appendDeadStockRows([{
          prodOrderRef: form.prodOrderNo,
          itemId: lookup,
          itemCode: ids.itemCode,
          batchId: ids.batchId,
          quantity: qtyKg,
          uom: _getItemUomCode(ids.itemCode),
          reason: deadStockReason,
          remarks: form.remarks || '',
          userEmail: Session.getActiveUser().getEmail()
        }]);
      }

      _appendMovementLogsBatch(logs);

      try { if (!isDeadStock) _updateBinAndRackStatuses([form.binId]); } catch (e) {
        Logger.log('[RETURN] Bin status update failed: ' + e.message);
      }
    } catch (e) {
      if (changes.length > 0) {
        _rollbackInventoryChanges(changes, _getSheetOrThrow(SpreadsheetApp.getActiveSpreadsheet(), CONFIG.SHEETS.INVENTORY));
      }
      _rollbackProductionLedger(ledgerChange);
      _rollbackProductionTransfers(prodTransferChange);
      _rollbackDeadStock(deadStockChange);
      throw e;
    }

    return { success: true, ginNo: ginForInventory, deadStock: isDeadStock };
  });
}


/**
 * FIX-5: Multi-item return â€” sequential, safe with script locks.
 * InwardForm calls this instead of Promise.all(submitReturnV2).
 *
 * @param {Object} form
 *   prodOrderNo {string}
 *   refNote     {string}   â† same ref note applied to all items in this return
 *   items       {Array}    each: { itemCode, batchId|batchNumber, quantity, binId, site, location, lotNo }
 *   remarks     {string}
 */
function submitReturnBatch(form) {
  return withScriptLock(function () {
    protect(function () { return requireOperationalUser(); });

    if (!form.prodOrderNo) throw new Error('Production Order No required');
    if (!Array.isArray(form.items) || form.items.length === 0) throw new Error('At least one item required');

    const ginBase = String(form.refNote || '').trim() || ('RET-' + form.prodOrderNo);
    const isDeadStock = form.isDeadStock === true;
    const deadStockReason = String(form.deadStockReason || '').trim();
    if (isDeadStock && !deadStockReason) {
      throw new Error('Dead stock reason required');
    }
    const results = [];

    form.items.forEach(function (item, idx) {
      const batchId = String(item.batchId || item.batchNumber || '').trim();
      if (!item.itemCode || !batchId) {
        throw new Error('Item ' + (idx + 1) + ': itemCode and batch required');
      }

      const ids = _assertItemCodeBatch(item.itemCode, batchId, 'ReturnBatch');
      const qtyKg = _convertToKg(ids.itemCode, item.quantity, item.uom);
      const lookup = _getValidatedItemId(ids.itemCode);

      const ledger = _getProductionLedgerState(form.prodOrderNo, lookup, ids.batchId, '');
      const outstanding = ledger.issued - ledger.returned - (ledger.rejected || 0);
      if (qtyKg > outstanding + 0.001) {
        throw new Error('Item ' + item.itemCode + ': cannot return ' + qtyKg.toFixed(3) + ' KG. Max: ' + outstanding.toFixed(3) + ' KG');
      }
      if (!ledger.version) throw new Error('Version not found for ' + item.itemCode + ' in ledger');

      if (!isDeadStock) {
        _assertBinCapacity(item.binId, item.quantity, null, ids.itemCode, item.uom);
      }

      // Append suffix for each line within same return batch for traceability
      const ginNo = ginBase + (form.items.length > 1 ? '-L' + (idx + 1) : '');

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
      let created = null;
      if (!isDeadStock) {
        created = _createInventoryRowV2({
          itemId: lookup,
          itemCode: ids.itemCode,
          batchId: ids.batchId,
          ginNo: ginNo,
          version: ledger.version,
          qualityStatus: _normalizeQaStatus(form.qualityStatus || 'PENDING'),
          qualityDate: form.qualityDate ? new Date(form.qualityDate) : '',
          binId: item.binId,
          site: item.site,
          location: item.location,
          quantity: qtyKg,
          lotNo: String(item.lotNo || ''),
          expiryDate: item.expiryDate ? new Date(item.expiryDate) : '',
          returnRowIndex: true
        });
      }

      const changes = created ? [{ action: 'create', rowIndex: created.rowIndex }] : [];
      let ledgerChange = null;
      let prodTransferChange = null;
      let deadStockChange = null;

      try {
        const binMeta = _getBinMetaById(item.binId);
        prodTransferChange = _appendProductionTransfersBatch([{
          transferType: isDeadStock ? 'RETURN_DEAD_STOCK' : 'RETURN',
          productionOrderNo: form.prodOrderNo,
          productionArea: 'PRODUCTION',
          itemCode: ids.itemCode,
          itemId: lookup,
          batchNumber: batchId,
          batchId: ids.batchId,
          lotNumber: item.lotNo || '',
          quantity: qtyKg,
          uom: _getItemUomCode(ids.itemCode),
          fromSite: 'PRODUCTION', fromLocation: 'PRODUCTION',
          fromBinId: 'PRODUCTION', fromBinCode: 'PRODUCTION',
          toSite: isDeadStock ? 'DEAD_STOCK' : String(item.site || ''),
          toLocation: isDeadStock ? 'DEAD_STOCK' : String(item.location || ''),
          toBinId: isDeadStock ? '' : String(item.binId || ''),
          toBinCode: isDeadStock ? 'DEAD_STOCK' : (binMeta.binCode || item.binId),
          returnedByName: Session.getActiveUser().getEmail(),
          createdBy: Session.getActiveUser().getEmail(),
          createdAt: new Date(),
          remarks: String(form.refNote ? '[Ref: ' + form.refNote + '] ' : '') +
            (isDeadStock ? ('[Dead Stock: ' + deadStockReason + '] ') : '') + String(form.remarks || ''),
          status: 'COMPLETED'
        }]);

        ledgerChange = _updateProductionLedger(
          form.prodOrderNo,
          lookup,
          ids.batchId,
          ledger.version,
          0,
          isDeadStock ? 0 : qtyKg,
          isDeadStock ? qtyKg : 0
        );

        if (isDeadStock) {
          deadStockChange = _appendDeadStockRows([{
            prodOrderRef: form.prodOrderNo,
            itemId: lookup,
            itemCode: ids.itemCode,
            batchId: ids.batchId,
            quantity: qtyKg,
            uom: _getItemUomCode(ids.itemCode),
            reason: deadStockReason,
            remarks: form.remarks || '',
            userEmail: Session.getActiveUser().getEmail()
          }]);
        }

        _appendMovementLogsBatch([{
          type: isDeadStock ? CONFIG.MOVEMENT_TYPES.RETURN_DEAD : CONFIG.MOVEMENT_TYPES.RETURN_PROD,
          itemId: lookup,
          batchId: ids.batchId,
          version: ledger.version,
          ginNo: ginNo,
          prodOrderRef: form.prodOrderNo,
          toBinId: isDeadStock ? '' : item.binId,
          quantity: qtyKg,
          remarks: isDeadStock
            ? ('DEAD_STOCK: ' + deadStockReason + (form.remarks ? ' | ' + form.remarks : ''))
            : form.remarks
        }]);

        try { if (!isDeadStock) _updateBinAndRackStatuses([item.binId]); } catch (e) { }

        results.push({ itemCode: ids.itemCode, ginNo: ginNo, qty: qtyKg, deadStock: isDeadStock });

      } catch (e) {
        if (changes.length > 0) _rollbackInventoryChanges(changes, sheet);
        _rollbackProductionLedger(ledgerChange);
        _rollbackProductionTransfers(prodTransferChange);
        _rollbackDeadStock(deadStockChange);
        throw e;
      }
    });

    return { success: true, returned: results };
  });
}


// 11.2  FIXED _createInventoryRow (now includes lotNo + expiryDate)
//
// Rename the existing _createInventoryRow to _createInventoryRowLegacy
// and use this as the new default. OR simply add the new params to the
// existing function â€” both approaches work. We use a V2 name here to
// avoid modifying the original.

function _createInventoryRowV2(params) {
  _validateBinExists(params.binId);
  if (!params.version) throw new Error('Inventory version required');
  if (!params.ginNo) throw new Error('Bill No. required for inventory');
  if (!params.itemCode || !params.batchId) throw new Error('Item Code and Batch No required for inventory');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  const id = 'INV-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
  const rowIndex = sheet.getLastRow() + 1;

  const row = [];
  row[CONFIG.INVENTORY_COLS.INVENTORY_ID] = id;
  row[CONFIG.INVENTORY_COLS.ITEM_ID] = params.itemId || params.itemCode;
  row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] = params.itemCode || params.itemId;
  row[CONFIG.INVENTORY_COLS.BATCH_ID] = params.batchId;
  row[CONFIG.INVENTORY_COLS.GIN_NO] = params.ginNo;
  row[CONFIG.INVENTORY_COLS.VERSION] = params.version || '';
  row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = params.qualityStatus || 'PENDING';
  row[CONFIG.INVENTORY_COLS.QUALITY_DATE] = params.qualityDate || '';
  row[CONFIG.INVENTORY_COLS.BIN_ID] = params.binId;
  row[CONFIG.INVENTORY_COLS.SITE] = params.site || 'Main';
  row[CONFIG.INVENTORY_COLS.LOCATION] = params.location || 'Main';
  row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = params.quantity;
  if (CONFIG.INVENTORY_COLS.UOM !== undefined) {
    row[CONFIG.INVENTORY_COLS.UOM] = _getItemUomCode(params.itemCode || params.itemId || '');
  }
  // FIX-3: Lot No
  if (CONFIG.INVENTORY_COLS.LOT_NO !== undefined) {
    row[CONFIG.INVENTORY_COLS.LOT_NO] = params.lotNo || '';
  }
  // FIX-4: Expiry Date
  if (CONFIG.INVENTORY_COLS.EXPIRY_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = _normalizeInwardExpiryDate(params.expiryDate);
  }
  row[CONFIG.INVENTORY_COLS.LAST_UPDATED] = new Date();
  if (CONFIG.INVENTORY_COLS.INWARD_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.INWARD_DATE] = params.inwardDate || new Date();
  }
  if (CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = params.lastTransferDate || '';
  }

  sheet.appendRow(row);
  _appendSheetCacheRow(CONFIG.SHEETS.INVENTORY, row);

  if (params.returnRowIndex === true) return { id: id, rowIndex: rowIndex };
  return id;
}


// ============================================================================
// REPLACE submitProductionRequestV3 IN Code.js (starts at line 5629)
// Delete the entire existing function and paste this instead
// ============================================================================

function submitProductionRequestV3(form) {
  return withScriptLock(function () {
    protect(function () { return requireOperationalUser(); });

    if (!form || !Array.isArray(form.items) || form.items.length === 0) {
      throw new Error('At least one item is required.');
    }
    if (form.items.length > 20) throw new Error('Maximum 20 items per production order.');

    // Fail-fast validation
    form.items.forEach(function (item, idx) {
      const label = 'Item ' + (idx + 1);
      if (!item.itemCode || !String(item.itemCode).trim()) throw new Error(label + ': Item Code is required.');
      if (!Number(item.quantity) || Number(item.quantity) <= 0) throw new Error(label + ' (' + item.itemCode + '): Valid quantity required.');
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);

    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    const rand = Math.floor(Math.random() * 9000 + 1000).toString();
    const orderRef = 'PO-' + dateStr + '-' + rand;

    const rowsToWrite = [];
    const movementLogs = [];

    form.items.forEach(function (item) {
      const itemCode = String(item.itemCode).trim();

      // CRITICAL FIX: Extract batchId in correct priority order
      const batchId = String(item.batchId || item.jobOrderNo || item.jobOrder || '').trim();

      // CRITICAL FIX: Get item_id from inventory where this batch exists
      let itemId;
      if (batchId && batchId !== '' && batchId !== 'TBD') {
        // Look up item_id from inventory for this specific batch
        const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
        if (invSheet) {
          const invData = _getSheetValuesCached(invSheet.getName());
          const codeU = itemCode.toUpperCase();
          const batchU = batchId.toUpperCase();

          // Find first inventory row with this item code + batch combination
          for (let i = 1; i < invData.length; i++) {
            const rowCode = String(invData[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
            const rowBatch = String(invData[i][CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim().toUpperCase();
            const rowQty = Number(invData[i][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;

            if (rowCode === codeU && rowBatch === batchU && rowQty > 0) {
              itemId = String(invData[i][CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim();
              break;
            }
          }
        }
      }

      // Fallback to Item_Master only if batch not specified or not found in inventory
      if (!itemId) {
        itemId = getItemIdByCode(itemCode) || itemCode;
      }

      const qty = Number(item.quantity);
      const qtyKg = _convertToKg(itemCode, qty, item.uom);
      const finalBatchId = batchId || 'TBD';

      rowsToWrite.push([
        orderRef, itemId, finalBatchId, 'TBD',
        qtyKg, 0, 0, qtyKg, 0, 'APPROVED', new Date()
      ]);

      // Store remarks as a PROD_REQUEST movement log entry
      if (form.remarks && String(form.remarks).trim()) {
        movementLogs.push({
          type: 'PROD_REQUEST',
          itemId: itemId,
          batchId: finalBatchId,
          version: '',
          ginNo: '',
          prodOrderRef: orderRef,
          fromBinId: '',
          toBinId: '',
          quantity: qtyKg,
          remarks: String(form.remarks).trim()
        });
      }
    });

    // Write all ledger rows atomically
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
    rowsToWrite.forEach(function (r) { _appendSheetCacheRow(CONFIG.SHEETS.PRODUCTION_LEDGER, r); });
    _clearScriptCachesForSheet(CONFIG.SHEETS.PRODUCTION_LEDGER);

    // Write remarks to movement log (non-blocking)
    if (movementLogs.length > 0) {
      try { _appendMovementLogsBatch(movementLogs); } catch (e) {
        Logger.log('[PO_REQUEST] Remarks movement log failed: ' + e.message);
      }
    }

    return { success: true, orderRef: orderRef, itemCount: rowsToWrite.length };
  });
}

// 11.4  RACK INITIALIZATION
//
// Populates existing physical stock into the WMS without creating
// a purchase inward transaction.
//
// Use case: "We already have material in these bins, we want the
// system to reflect that without re-running all historical inwards."
//
// Requires: MANAGER role
// Rejects: any bin that already has inventory rows (prevents double-init)
// Creates: STOCK_INIT movement type (auditable, distinguishable from INWARD)
// â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

/**
 * @param {Object} payload
 *   rows {Array} each row:
 *     itemCode    {string}  required
 *     batchId     {string}  required
 *     version     {string}  e.g. "V1"
 *     binId       {string}  required
 *     site        {string}
 *     location    {string}
 *     quantity    {number}  in selected UOM (stored internally in base UOM)
 *     uom         {string}  optional (e.g. KG, NOS)
 *     lotNo       {string}  optional
 *     expiryDate  {string}  optional YYYY-MM-DD
 *     qualityStatus {string} default 'APPROVED' (physical stock assumed QA-cleared)
 *   overwrite {boolean}  default false â€” if true, allows re-init of already-stocked bins
 *   remarks   {string}   reason for initialization
 *
 * @returns {{ success: true, created: number, skipped: number, errors: string[] }}
 */
function initRackInventory(payload) {
  return withScriptLock(function () {
    protect(function () { requireRole(SECURITY.ROLES.MANAGER); });

    if (!payload || !Array.isArray(payload.rows) || payload.rows.length === 0) {
      throw new Error('No rows provided for rack initialization.');
    }
    if (payload.rows.length > 200) {
      throw new Error('Maximum 200 rows per initialization batch.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const overwrite = payload.overwrite === true;
    const remarks = String(payload.remarks || 'Physical stock initialization').trim();
    const initGin = 'INIT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');

    const errors = [];
    const logsToWrite = [];
    var created = 0;
    var skipped = 0;

    payload.rows.forEach(function (row, idx) {
      const label = 'Row ' + (idx + 1);

      // Validate required fields
      if (!row.itemCode || !String(row.itemCode).trim()) { errors.push(label + ': itemCode required'); return; }
      if (!row.batchId || !String(row.batchId).trim()) { errors.push(label + ': batchId required'); return; }
      if (!row.binId || !String(row.binId).trim()) { errors.push(label + ': binId required'); return; }
      if (!row.quantity || Number(row.quantity) <= 0) { errors.push(label + ': quantity must be > 0'); return; }

      const itemCode = String(row.itemCode).trim();
      const batchId = String(row.batchId).trim();
      const binId = String(row.binId).trim();
      const qtyInput = Number(row.quantity);
      const version = String(row.version || 'V1').trim();
      const qaStatus = _normalizeQaStatus(String(row.qualityStatus || 'APPROVED'));
      const lotNo = String(row.lotNo || '');
      const site = String(row.site || '');
      const location = String(row.location || '');
      const qty = _convertToKg(itemCode, qtyInput, row.uom);

      // Check for existing inventory in this bin for this item+batch
      if (!overwrite) {
        try {
          const existing = _readInventoryState(itemCode, batchId, binId, version)
            .filter(function (r) { return Number(r.currentQty) > 0; });
          if (existing.length > 0) {
            skipped++;
            Logger.log('[INIT] Skipped ' + itemCode + ' in ' + binId + ' â€” already has inventory');
            return;
          }
        } catch (e) {
          // _readInventoryState may throw if item not found â€” treat as "no existing"
        }
      }

      // Validate bin exists
      try { _validateBinExists(binId); } catch (e) {
        errors.push(label + ': bin "' + binId + '" does not exist â€” ' + e.message);
        return;
      }

      // Resolve item ID
      let itemId;
      try { itemId = _getValidatedItemId(itemCode); } catch (e) {
        errors.push(label + ': item "' + itemCode + '" not found in item master');
        return;
      }

      // Create inventory row
      try {
        _createInventoryRowV2({
          itemId: itemId,
          itemCode: itemCode,
          batchId: batchId,
          ginNo: initGin,
          version: version,
          qualityStatus: qaStatus,
          qualityDate: row.qualityDate ? new Date(row.qualityDate) : new Date(),
          binId: binId,
          site: site,
          location: location,
          quantity: qty,
          lotNo: lotNo,
          expiryDate: row.expiryDate ? new Date(row.expiryDate) : ''
        });

        logsToWrite.push({
          type: 'STOCK_INIT',
          itemId: itemId,
          batchId: batchId,
          version: version,
          ginNo: initGin,
          prodOrderRef: '',
          fromBinId: '',
          toBinId: binId,
          quantity: qty,
          qualityStatus: qaStatus,
          remarks: remarks
        });

        created++;
      } catch (e) {
        errors.push(label + ' (' + itemCode + '): ' + e.message);
      }
    });

    // Write movement logs for all successful rows
    if (logsToWrite.length > 0) {
      try { _appendMovementLogsBatch(logsToWrite); } catch (e) {
        Logger.log('[INIT] Movement log write failed: ' + e.message);
      }
    }

    // Update bin/rack statuses for all touched bins
    const touchedBins = payload.rows
      .filter(function (r) { return r.binId; })
      .map(function (r) { return String(r.binId).trim(); });
    const uniqueBins = Array.from(new Set(touchedBins));
    try { _updateBinAndRackStatuses(uniqueBins); } catch (e) {
      Logger.log('[INIT] Rack status update failed: ' + e.message);
    }

    if (errors.length > 0 && created === 0) {
      throw new Error('Initialization failed:\n' + errors.join('\n'));
    }

    return {
      success: true,
      created: created,
      skipped: skipped,
      errors: errors,
      ginRef: initGin
    };
  });
}


// 
// 11.5  getInitializationTemplate
// Returns current inventory state per bin for cross-check before init
// 

/**
 * Returns a count of items per bin to help staff verify
 * which bins already have stock vs which are truly empty.
 */
function getBinStockSummaryForInit() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());
    const binMap = {};

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;
      const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
      if (!binId) continue;
      binMap[binId] = (binMap[binId] || 0) + qty;
    }

    return Object.keys(binMap).map(function (binId) {
      return { binId: binId, totalQty: binMap[binId] };
    }).sort(function (a, b) { return b.totalQty - a.totalQty; });
  });
}
// 
// END SECTION 11
// 

// =====================================================
// SECTION 12: UOM DISPLAY + FORM HELPERS
// Paste at bottom of Code.js.
//
// Purpose:
//   getItemUomForForm  — called by every form on item-code blur
//   _getItemUomCode    — internal helper used by dashboards
//
// No schema changes required.
// No existing logic is changed.
// UOM display is purely additive — the backend already stores
// quantities in the item's base UOM via _convertToKg.
//
// The UOM displayed in forms is for worker clarity ONLY.
// The stored value is always in the base UOM (KG for weight items,
// NOS for count items, etc.) as defined in Item_Master.
// =====================================================

/**
 * Returns the UOM code for an item so the form can display
 * "Quantity (NOS)" instead of "Quantity (KG)" for count items.
 *
 * Called by: InwardForm, DispatchForm, TransferForm, ProductionRequestForm
 * on item-code blur/change.
 *
 * @param {string} itemCode
 * @returns {{ uomCode: string, itemCode: string }} or null
 *
 * NOTE: Returns uomCode from Item_Master.UOM_CODE (header-driven lookup).
 * The "factorToKg" conversion only applies when the item also
 * has an ALT_UOM column and a CONVERSION_FACTOR_TO_KG column.
 * For NOS items, factor = 1 (no conversion needed).
 */
function getItemUomForForm(itemCode) {
  return protect(function () {
    const code = String(itemCode || '').trim().toUpperCase();
    const allOptions = _getAllActiveUomCodes();
    if (!code) {
      const defaultUom = String(allOptions[0] || 'KG').trim().toUpperCase() || 'KG';
      return { uomCode: defaultUom, altUom: '', uomOptions: allOptions, allUomOptions: allOptions, itemCode: code };
    }

    const info = _getItemUomInfo(itemCode); // existing function in Code.js
    const options = _getItemUomOptions(itemCode);
    return {
      uomCode: info.baseUom || 'KG',
      altUom: info.altUom || '',
      uomOptions: options,
      allUomOptions: allOptions,
      itemCode: code
    };
  });
}

/**
 * Returns the UOM code string for an item from Item_Master.
 * Used internally by dashboard/report functions to format qty fields.
 *
 * @param {string} itemCode
 * @returns {string}  e.g. "KG", "NOS", "FT", "MTR"
 */
function _getItemUomCode(itemCode) {
  const info = _getItemUomInfo(itemCode);
  return info.baseUom || 'KG';
}


// ─────────────────────────────────────────────────────────────────────
// SECTION 12.2 — UOM-AWARE DASHBOARD QUERY HELPER
//
// Use this in any server-side function that returns quantities
// to the dashboards so they can display the right unit label.
//
// Dashboard usage example (read-only, no schema change):
//
//   const inv = getInventoryWithUom(); // returns rows with uomCode field
//   rows.forEach(r => label = r.qty + ' ' + r.uomCode);
// ─────────────────────────────────────────────────────────────────────

/**
 * Enriches inventory rows with a uomCode field from Item_Master.
 * Called by dashboard functions to display correct unit labels.
 *
 * This function reads the existing inventory snapshot and looks up
 * the UOM for each unique item code in a single pass — no extra
 * sheet reads per row.
 *
 * @returns {Array} inventory rows each with { ...existingFields, uomCode: string }
 */
function getInventoryWithUom() {
  return protect(function () {
    const rows = getInventoryReadView(); // existing function

    // Build a UOM cache in one pass over all unique item codes
    const uomCache = {};
    rows.forEach(function (r) {
      const code = String(r.itemCode || '').trim().toUpperCase();
      if (!uomCache[code]) {
        uomCache[code] = _getItemUomCode(r.itemCode);
      }
    });

    return rows.map(function (r) {
      const code = String(r.itemCode || '').trim().toUpperCase();
      return Object.assign({}, r, { uomCode: uomCache[code] || 'KG' });
    });
  });
}

/**
 * Creates (or refreshes if empty) UOM_Master with recommended structure.
 * Sheet columns:
 *   UOM_CODE | UOM_NAME | BASE_UOM | FACTOR_TO_BASE | IS_ACTIVE | ALLOW_DECIMAL | DISPLAY_ORDER | NOTES
 */
function setupUomMasterTemplate() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.UOM_MASTER || 'UOM_Master');
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEETS.UOM_MASTER || 'UOM_Master');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      const header = ['UOM_CODE', 'UOM_NAME', 'BASE_UOM', 'FACTOR_TO_BASE', 'IS_ACTIVE', 'ALLOW_DECIMAL', 'DISPLAY_ORDER', 'NOTES'];
      const seed = [
        ['KG', 'Kilogram', 'KG', 1, 'TRUE', 'TRUE', 1, 'Base weight unit'],
        ['G', 'Gram', 'KG', 0.001, 'TRUE', 'TRUE', 2, '1000 g = 1 kg'],
        ['TON', 'Metric Ton', 'KG', 1000, 'TRUE', 'TRUE', 3, '1 ton = 1000 kg'],
        ['NOS', 'Numbers', 'NOS', 1, 'TRUE', 'FALSE', 4, 'Base count unit'],
        ['PCS', 'Pieces', 'NOS', 1, 'TRUE', 'FALSE', 5, 'Alias of NOS'],
        ['L', 'Litre', 'L', 1, 'TRUE', 'TRUE', 6, 'Base liquid unit'],
        ['ML', 'Millilitre', 'L', 0.001, 'TRUE', 'TRUE', 7, '1000 ml = 1 L']
      ];
      sheet.clearContents();
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
      sheet.getRange(2, 1, seed.length, header.length).setValues(seed);
      sheet.setFrozenRows(1);
      return { success: true, message: 'UOM_Master template created', rows: seed.length };
    }

    return { success: true, message: 'UOM_Master already present', rows: lastRow - 1 };
  });
}
// ─────────────────────────────────────────────────────────────────────
// END SECTION 12
// ─────────────────────────────────────────────────────────────────────

/**
 * TESTING FUNCTION: Use this to verify your lookup works
 * 
 * How to test:
 * 1. Go to Apps Script Editor
 * 2. Select this function from dropdown
 * 3. Click Run
 * 4. Check Execution Log for results
 */
function testBatchLookup() {
  // Replace with actual Item Code + Batch No + Lot No from Batch_Master
  const testItemCode = 'ITEM001';
  const testBatchNo = 'BATCH001';
  const testLotNo = 'LOT12345';

  const result = resolveBatchComposite(testItemCode, testBatchNo, testLotNo);

  if (result) {
    Logger.log('âœ“ Batch found!');
    Logger.log('Batch ID: ' + result.batchId);
    Logger.log('Batch Number: ' + result.batchNumber);
    Logger.log('Item Code: ' + result.itemCode);
    Logger.log('Expiry Date: ' + result.expiryDate);
  } else {
    Logger.log('âœ— Lot number not found: ' + testLotNo);
    Logger.log('Check your Batch_Master sheet for available lot numbers');
  }
}
function debugQaCount() {
  const rows = getQaInventoryView();
  Logger.log('QA rows: ' + (rows || []).length);
  Logger.log(rows);
}


// ======================================================================
// SECTION 9: AI ASSISTANT MODULE
// Production-Grade Hybrid RAG â€” Pure Retrieval Mode (No External APIs)
//
// ARCHITECTURE CONTRACT:
//   - READ-ONLY. Zero writes to any transactional sheet.
//   - NO LockService. Reads are idempotent and non-critical.
//   - FULLY ISOLATED. Zero imports from Sections 1-8 except:
//       getEffectiveUser()  â†’ auth
//       SECURITY.ROLES      â†’ role constants
//       ROLE_RANK           â†’ rank comparisons
//   - Gemini integration is STUBBED. Architecture is in place.
//     Set AI_CONFIG.GEMINI_ENABLED = true and add API key to enable.
//   - All functions prefixed _ai* (private) or askWmsAI (public endpoint).
// ======================================================================

// -----------------------------------------------------------------------
// 9.1  AI MODULE CONFIGURATION
// -----------------------------------------------------------------------

/** @const {Object} AI_CONFIG â€” All AI tuning lives here. Zero hardcoding below. */
const AI_CONFIG = {

  // --- Sheet names ---
  KNOWLEDGE_SHEET: 'WMS_AI_Knowledge',
  QUERY_LOG_SHEET: 'AI_Query_Log',

  // --- Knowledge sheet column indices (0-based, mirrors exact sheet schema) ---
  COLS: {
    DOC_ID: 0,
    DOC_VERSION: 1,
    MODULE: 2,
    PROCESS_STAGE: 3,
    ENTITY_SCOPE: 4,
    ROLE_SCOPE: 5,
    KEYWORDS: 6,
    ARCH_DEPENDENCIES: 7,
    CONTENT: 8,
    RELATED_SHEETS: 9,
    RELATED_FUNCTIONS: 10,
    SECURITY_LEVEL: 11,
    DEPRECATED: 12,
    LAST_UPDATED: 13
  },

  // --- Retrieval tuning ---
  TOP_N: 5,     // Return top N scored results
  MIN_SCORE: 1,     // Minimum score to include a result
  CACHE_TTL_SECONDS: 300,   // CacheService TTL for knowledge sheet (5 min)
  CACHE_KEY: 'AI_KB_V1_WMS_KNOWLEDGE',

  // --- Scoring weights (per spec) ---
  SCORE: {
    KEYWORD_MATCH: 3,
    MODULE_MATCH: 2,
    ENTITY_MATCH: 1,
    PHRASE_BONUS: 2   // bonus when query phrase appears verbatim in CONTENT
  },

  // --- Security levels ---
  SEC_LEVEL: {
    MANAGER_ONLY: 'MANAGER_ONLY',
    ALL: 'ALL'
  },

  // --- Role scope sentinel ---
  ROLE_SCOPE_ALL: 'ALL',

  // --- Query Log columns (auto-created sheet) ---
  LOG_COLS: [
    'QUERY_ID', 'TIMESTAMP', 'USER_EMAIL', 'USER_ROLE',
    'QUERY_RAW', 'QUERY_NORMALIZED', 'MODE', 'RESULT_COUNT',
    'SOURCES_RETURNED', 'FLAGGED', 'FLAG_REASON', 'RESPONSE_TIME_MS'
  ],

  // --- Abuse detection: block queries matching ANY of these tokens ---
  // Each entry is evaluated as a whole-word or whole-phrase match.
  ABUSE_TOKENS: [
    'delete', 'bypass', 'disable lock', 'edit sheet',
    'drop table', 'truncate', 'script inject',
    'ignore previous', 'forget instructions'
  ],

  // --- Caution tokens: logged but not blocked (for analytics) ---
  CAUTION_TOKENS: ['override', 'disable', 'lock'],

  // -------------------------------------------------------------------------
  // GEMINI STUB â€” disabled by default.
  // To enable: set GEMINI_ENABLED = true and populate GEMINI_API_KEY.
  // When enabled, Layer 2 fires only when Layer 1 returns < GEMINI_THRESHOLD_CHARS
  // of combined content, or when query contains synthesis signal words.
  // -------------------------------------------------------------------------
  GEMINI_ENABLED: false,
  GEMINI_API_KEY: '',   // PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
  GEMINI_MODEL: 'gemini-1.5-flash',
  GEMINI_MAX_TOKENS: 512,
  GEMINI_THRESHOLD_CHARS: 200,  // Trigger Gemini if Layer 1 context < this
  GEMINI_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/',

  // --- System prompt for Gemini (used only when GEMINI_ENABLED = true) ---
  GEMINI_SYSTEM_PROMPT: [
    'You are the WMS Operations Assistant for a pharmaceutical warehouse.',
    'Your ONLY knowledge source is the context provided below.',
    'Rules:',
    '1. Answer ONLY from the provided SOP content. Never invent facts.',
    '2. If the answer is not in the context, say exactly: "I don\'t have information about that in the knowledge base."',
    '3. Never suggest bypassing system controls, locks, QA approvals, or ACL.',
    '4. Always cite the SOP source (e.g., "Per SOP-001...").',
    '5. Maximum 150 words.',
    '6. If the question involves a blocked transaction or QA hold, always end with:',
    '   "For any blocked transaction, contact your supervisor before proceeding."',
    '7. Never reveal user emails, quantities, batch IDs from context.',
    'Context:'
  ].join('\n')
};

// -----------------------------------------------------------------------
// 9.2  PUBLIC ENTRY POINT
// -----------------------------------------------------------------------

/**
 * askWmsAI â€” Public entry point. Called by AIAssistant.html via google.script.run.
 *
 * @param  {string} question â€” Raw user query from textarea.
 * @returns {Object} { mode, answer, sources, flagged, flagReason, resultCount }
 */
function askWmsAI(question) {
  const startMs = new Date().getTime();
  let user;

  // --- Auth gate: must be an active WMS user ---
  try {
    user = getEffectiveUser();
  } catch (authErr) {
    return _aiSafeError('AUTH_FAILED', 'Authentication failed. Please refresh and try again.', [], startMs);
  }

  // Supervisory-only access (Quality Manager and Manager).
  const userRank = ROLE_RANK[String(user.role || '').toUpperCase()] || 0;
  const minAiRank = ROLE_RANK[SECURITY.ROLES.QUALITY_MANAGER] || 2;
  if (userRank < minAiRank) {
    return _aiSafeError(
      'ACCESS_DENIED',
      'AI assistant is available for QA Manager / Manager roles only.',
      [],
      startMs
    );
  }

  // --- Input sanitisation ---
  const rawQuery = String(question || '').trim();
  if (!rawQuery || rawQuery.length < 3) {
    return _aiSafeError('QUERY_TOO_SHORT', 'Please enter at least 3 characters.', [], startMs);
  }
  if (rawQuery.length > 500) {
    return _aiSafeError('QUERY_TOO_LONG', 'Query exceeds 500 characters. Please be more specific.', [], startMs);
  }

  // --- Layer 0: Abuse detection (immediate gate, no logging on hard block) ---
  const abuseCheck = _aiDetectAbuse(rawQuery);
  if (abuseCheck.blocked) {
    _aiLogQuery(user, rawQuery, '', 'BLOCKED', [], 0, true, abuseCheck.reason, new Date().getTime() - startMs);
    return {
      mode: 'BLOCKED',
      answer: 'This query cannot be processed. Please contact your supervisor if you need assistance with this topic.',
      sources: [],
      flagged: true,
      flagReason: abuseCheck.reason,
      resultCount: 0
    };
  }

  // --- Layer 1: Keyword retrieval from WMS_AI_Knowledge ---
  const normalizedQuery = _aiNormalizeQuery(rawQuery);

  // --- Live-data mode: answer from Inventory / PO data first ---
  try {
    const live = _aiTryLiveDataAnswer(rawQuery, normalizedQuery, user);
    if (live && live.matched) {
      const elapsedLive = new Date().getTime() - startMs;
      _aiLogQuery(
        user,
        rawQuery,
        normalizedQuery,
        live.mode || 'LIVE_DATA',
        live.sources || [],
        Number(live.resultCount) || 0,
        abuseCheck.caution,
        abuseCheck.reason,
        elapsedLive
      );
      return {
        mode: live.mode || 'LIVE_DATA',
        answer: live.answer || 'No live data found for your query.',
        sources: live.sources || [],
        flagged: abuseCheck.caution,
        flagReason: abuseCheck.caution ? abuseCheck.reason : '',
        resultCount: Number(live.resultCount) || 0
      };
    }
  } catch (liveErr) {
    Logger.log('[AI_LIVE_ERR] Falling back to knowledge retrieval: ' + liveErr.message);
  }

  let retrievedDocs;
  try {
    retrievedDocs = _aiSearchKnowledge(normalizedQuery, user);
  } catch (searchErr) {
    Logger.log('[AI_SEARCH_ERR] ' + searchErr.message);
    return _aiSafeError('SEARCH_FAILED', 'Knowledge base is temporarily unavailable. Please try again shortly.', [], startMs);
  }

  if (!retrievedDocs || retrievedDocs.length === 0) {
    _aiLogQuery(user, rawQuery, normalizedQuery, 'RETRIEVAL', [], 0, false, '', new Date().getTime() - startMs);
    return {
      mode: 'RETRIEVAL',
      answer: 'No matching information found in the knowledge base for your query.\n\nTip: Try different keywords, e.g., "inward receipt", "dispatch steps", "QA approval process".',
      sources: [],
      flagged: false,
      flagReason: '',
      resultCount: 0
    };
  }

  // --- Build ranked context from retrieved docs ---
  const ranked = _aiRankResults(retrievedDocs);
  const context = _aiBuildContext(ranked);
  const sources = ranked.map(function (r) { return r.docId; });

  // --- Layer 2: Gemini synthesis (stubbed â€” fires only if enabled) ---
  let mode = 'RETRIEVAL';
  let answer = context.formattedAnswer;

  if (AI_CONFIG.GEMINI_ENABLED && AI_CONFIG.GEMINI_API_KEY) {
    const contextTooShort = context.rawContentLength < AI_CONFIG.GEMINI_THRESHOLD_CHARS;
    const needsSynthesis = _aiRequiresSynthesis(normalizedQuery);
    if (contextTooShort || needsSynthesis) {
      try {
        const geminiAnswer = _aiCallGemini(rawQuery, context.geminiContext);
        if (geminiAnswer && geminiAnswer.length > 10) {
          answer = geminiAnswer;
          mode = 'GEMINI';
        }
      } catch (geminiErr) {
        Logger.log('[AI_GEMINI_ERR] Falling back to retrieval: ' + geminiErr.message);
        // Fallback gracefully â€” retrieval answer is already set
      }
    }
  }

  // --- Audit log ---
  const elapsedMs = new Date().getTime() - startMs;
  _aiLogQuery(user, rawQuery, normalizedQuery, mode, sources, ranked.length, abuseCheck.caution, abuseCheck.reason, elapsedMs);

  return {
    mode: mode,
    answer: answer,
    sources: sources,
    flagged: abuseCheck.caution,
    flagReason: abuseCheck.caution ? abuseCheck.reason : '',
    resultCount: ranked.length
  };
}

// -----------------------------------------------------------------------
// 9.3  LAYER 1 â€” KEYWORD RETRIEVAL
// -----------------------------------------------------------------------

/**
 * _aiSearchKnowledge â€” Loads knowledge sheet (cached), applies role filter,
 * scores each row against the normalized query, returns top N.
 *
 * @param  {string} normalizedQuery â€” Lowercased, cleaned query string.
 * @param  {Object} user            â€” getEffectiveUser() result.
 * @returns {Array}  Raw scored doc objects (unsorted).
 */
function _aiSearchKnowledge(normalizedQuery, user) {
  const rows = _aiLoadKnowledgeSheet();
  const cols = AI_CONFIG.COLS;
  const userRole = String(user.role || '').toUpperCase();
  const userRank = ROLE_RANK[userRole] || 0;
  const queryTokens = normalizedQuery.split(/\s+/).filter(function (t) { return t.length > 1; });

  const results = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];

    // --- Guard: skip empty rows ---
    if (!row[cols.DOC_ID]) continue;

    // --- Filter 1: Deprecated check ---
    var deprecated = String(row[cols.DEPRECATED] || '').trim().toUpperCase();
    if (deprecated === 'TRUE' || deprecated === '1' || deprecated === 'YES') continue;

    // --- Filter 2: Role scope check ---
    var roleScope = String(row[cols.ROLE_SCOPE] || 'ALL').trim().toUpperCase();
    if (roleScope !== AI_CONFIG.ROLE_SCOPE_ALL) {
      // May be comma-separated list of allowed roles
      var allowedRoles = roleScope.split(',').map(function (r) { return r.trim(); });
      if (allowedRoles.indexOf(userRole) === -1) continue;
    }

    // --- Filter 3: Security level check ---
    var secLevel = String(row[cols.SECURITY_LEVEL] || 'ALL').trim().toUpperCase();
    if (secLevel === AI_CONFIG.SEC_LEVEL.MANAGER_ONLY) {
      // Only MANAGER (rank 3) can see MANAGER_ONLY docs
      if (userRank < ROLE_RANK[SECURITY.ROLES.MANAGER]) continue;
    }

    // --- Scoring ---
    var score = _aiScoreRow(row, queryTokens, normalizedQuery, cols);
    if (score < AI_CONFIG.MIN_SCORE) continue;

    results.push({
      score: score,
      docId: String(row[cols.DOC_ID]).trim(),
      docVersion: String(row[cols.DOC_VERSION] || 'v1.0').trim(),
      module: String(row[cols.MODULE] || '').trim(),
      processStage: String(row[cols.PROCESS_STAGE] || '').trim(),
      entityScope: String(row[cols.ENTITY_SCOPE] || '').trim(),
      keywords: String(row[cols.KEYWORDS] || '').trim(),
      archDependencies: String(row[cols.ARCH_DEPENDENCIES] || '').trim(),
      content: String(row[cols.CONTENT] || '').trim(),
      relatedSheets: String(row[cols.RELATED_SHEETS] || '').trim(),
      relatedFunctions: String(row[cols.RELATED_FUNCTIONS] || '').trim(),
      securityLevel: secLevel,
      lastUpdated: row[cols.LAST_UPDATED]
    });
  }

  return results;
}

/**
 * _aiScoreRow â€” Calculates relevance score for a single knowledge row.
 * Scoring: +3 per keyword match, +2 module match, +1 entity match, +2 phrase bonus.
 *
 * @param  {Array}  row          â€” Raw sheet row.
 * @param  {Array}  queryTokens  â€” Individual normalized query words.
 * @param  {string} fullQuery    â€” Full normalized query for phrase matching.
 * @param  {Object} cols         â€” Column index map.
 * @returns {number} Score (0 = no match).
 */
function _aiScoreRow(row, queryTokens, fullQuery, cols) {
  var score = 0;

  // --- Keyword matching (+3 per matched keyword) ---
  var keywordsRaw = String(row[cols.KEYWORDS] || '').toLowerCase();
  var keywords = keywordsRaw.split(',').map(function (k) { return k.trim(); }).filter(Boolean);
  for (var ki = 0; ki < keywords.length; ki++) {
    for (var ti = 0; ti < queryTokens.length; ti++) {
      if (keywords[ki] === queryTokens[ti] || keywords[ki].indexOf(queryTokens[ti]) !== -1) {
        score += AI_CONFIG.SCORE.KEYWORD_MATCH;
        break; // count each keyword once per query token pass
      }
    }
  }

  // --- Module matching (+2 if query contains module name) ---
  var module = String(row[cols.MODULE] || '').toLowerCase();
  if (module && fullQuery.indexOf(module) !== -1) {
    score += AI_CONFIG.SCORE.MODULE_MATCH;
  }
  // Also check each token against module
  for (var mt = 0; mt < queryTokens.length; mt++) {
    if (module && module.indexOf(queryTokens[mt]) !== -1 && queryTokens[mt].length > 3) {
      score += AI_CONFIG.SCORE.MODULE_MATCH;
      break;
    }
  }

  // --- Entity scope matching (+1 if query contains entity) ---
  var entity = String(row[cols.ENTITY_SCOPE] || '').toLowerCase();
  if (entity) {
    var entityTokens = entity.split(',').map(function (e) { return e.trim(); });
    for (var et = 0; et < entityTokens.length; et++) {
      if (fullQuery.indexOf(entityTokens[et]) !== -1 && entityTokens[et].length > 2) {
        score += AI_CONFIG.SCORE.ENTITY_MATCH;
        break;
      }
    }
  }

  // --- Phrase bonus (+2 if full query phrase appears in CONTENT) ---
  var content = String(row[cols.CONTENT] || '').toLowerCase();
  if (fullQuery.length > 5 && content.indexOf(fullQuery) !== -1) {
    score += AI_CONFIG.SCORE.PHRASE_BONUS;
  }

  return score;
}

// -----------------------------------------------------------------------
// 9.4  RANKING
// -----------------------------------------------------------------------

/**
 * _aiRankResults â€” Sorts by score descending, returns top N.
 *
 * @param  {Array} results â€” Unranked scored doc objects.
 * @returns {Array} Top N results sorted by score.
 */
function _aiRankResults(results) {
  var sorted = results.slice().sort(function (a, b) {
    // Primary: score descending
    if (b.score !== a.score) return b.score - a.score;
    // Secondary: prefer more recently updated docs
    var aDate = a.lastUpdated instanceof Date ? a.lastUpdated.getTime() : 0;
    var bDate = b.lastUpdated instanceof Date ? b.lastUpdated.getTime() : 0;
    return bDate - aDate;
  });
  return sorted.slice(0, AI_CONFIG.TOP_N);
}

// -----------------------------------------------------------------------
// 9.5  CONTEXT BUILDER
// -----------------------------------------------------------------------

/**
 * _aiBuildContext â€” Transforms ranked results into:
 *   (a) formattedAnswer â€” Human-readable markdown for direct display.
 *   (b) geminiContext   â€” Compressed context string for Gemini (when enabled).
 *   (c) rawContentLength â€” Total character count for Gemini threshold check.
 *
 * @param  {Array} ranked â€” Top N ranked result objects.
 * @returns {Object} { formattedAnswer, geminiContext, rawContentLength }
 */
function _aiBuildContext(ranked) {
  if (!ranked || ranked.length === 0) {
    return { formattedAnswer: '', geminiContext: '', rawContentLength: 0 };
  }

  var sections = [];
  var geminiParts = [];
  var totalChars = 0;

  for (var i = 0; i < ranked.length; i++) {
    var doc = ranked[i];
    totalChars += doc.content.length;

    // --- Primary display section (first result is the main answer) ---
    var section = [];
    section.push('â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”');
    section.push('ðŸ“‹ SOURCE: ' + doc.docId + ' (' + doc.docVersion + ')');
    if (doc.module) section.push('ðŸ“¦ MODULE: ' + doc.module);
    if (doc.processStage) section.push('ðŸ”„ PROCESS STAGE: ' + doc.processStage);
    if (doc.archDependencies && doc.archDependencies !== '') {
      section.push('ðŸ”— DEPENDENCIES: ' + doc.archDependencies);
    }
    section.push('');
    section.push(doc.content);
    if (doc.relatedSheets && doc.relatedSheets !== '') {
      section.push('');
      section.push('ðŸ“Š Related Sheets: ' + doc.relatedSheets);
    }
    if (doc.relatedFunctions && doc.relatedFunctions !== '') {
      section.push('âš™ï¸ Related Functions: ' + doc.relatedFunctions);
    }

    if (i === 0) {
      // First result is the primary answer â€” show in full
      sections.push(section.join('\n'));
    } else {
      // Additional results shown as supplementary references
      sections.push('\n' + section.join('\n'));
    }

    // --- Compressed Gemini context (key fields only, saves tokens) ---
    geminiParts.push(
      'SOURCE: ' + doc.docId + ' (' + doc.docVersion + ')\n' +
      'MODULE: ' + doc.module + '\n' +
      'CONTENT: ' + doc.content.substring(0, 800) // cap per doc for Gemini context window
    );
  }

  var suffix = ranked.length > 1
    ? '\n\nâ”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”\n' +
    'ðŸ“Œ ' + ranked.length + ' reference(s) found. Primary result shown above.\n' +
    'Sources: ' + ranked.map(function (r) { return r.docId; }).join(', ')
    : '';

  return {
    formattedAnswer: sections.join('\n') + suffix,
    geminiContext: geminiParts.join('\n\n---\n\n'),
    rawContentLength: totalChars
  };
}

// -----------------------------------------------------------------------
// 9.6  LAYER 2 â€” GEMINI STUB (architecture in place, disabled)
// -----------------------------------------------------------------------

/**
 * _aiCallGemini â€” Calls Gemini REST API.
 * CURRENTLY DISABLED. Set AI_CONFIG.GEMINI_ENABLED = true to activate.
 *
 * @param  {string} userQuery    â€” Original user question.
 * @param  {string} contextBlock â€” Compressed SOP context from _aiBuildContext.
 * @returns {string} Gemini answer text.
 * @throws  {Error} On API failure (caller must catch and fallback).
 */
function _aiCallGemini(userQuery, contextBlock) {
  if (!AI_CONFIG.GEMINI_ENABLED || !AI_CONFIG.GEMINI_API_KEY) {
    throw new Error('Gemini not configured');
  }

  var prompt = AI_CONFIG.GEMINI_SYSTEM_PROMPT + '\n\n' + contextBlock + '\n\nQuestion: ' + userQuery;

  var payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      maxOutputTokens: AI_CONFIG.GEMINI_MAX_TOKENS,
      temperature: 0.1,  // Low temperature = factual, consistent
      topP: 0.8
    },
    safetySettings: [
      { category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_LOW_AND_ABOVE' },
      { category: 'HARM_CATEGORY_HATE_SPEECH', threshold: 'BLOCK_LOW_AND_ABOVE' },
      { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_LOW_AND_ABOVE' },
      { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_LOW_AND_ABOVE' }
    ]
  };

  var url = AI_CONFIG.GEMINI_ENDPOINT +
    AI_CONFIG.GEMINI_MODEL + ':generateContent?key=' +
    AI_CONFIG.GEMINI_API_KEY;

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Gemini API returned HTTP ' + code + ': ' + response.getContentText().substring(0, 200));
  }

  var json = JSON.parse(response.getContentText());

  // Safety: validate response structure before accessing
  if (!json.candidates || !json.candidates[0] ||
    !json.candidates[0].content || !json.candidates[0].content.parts) {
    throw new Error('Gemini returned unexpected response structure');
  }

  var text = json.candidates[0].content.parts[0].text || '';
  return text.trim();
}

/**
 * _aiRequiresSynthesis â€” Heuristic: does this query need cross-SOP synthesis?
 * Synthesis keywords signal that the user wants comparison or explanation,
 * not just a procedure lookup.
 *
 * @param  {string} normalizedQuery
 * @returns {boolean}
 */
function _aiRequiresSynthesis(normalizedQuery) {
  var synthesisSignals = [
    'difference between', 'compare', 'explain why', 'what is the reason',
    'how does', 'why does', 'what happens when', 'relationship between'
  ];
  for (var i = 0; i < synthesisSignals.length; i++) {
    if (normalizedQuery.indexOf(synthesisSignals[i]) !== -1) return true;
  }
  return false;
}

// -----------------------------------------------------------------------
// 9.7  ABUSE DETECTION
// -----------------------------------------------------------------------

/**
 * _aiDetectAbuse â€” Checks query against blocked and caution token lists.
 * Blocked queries return { blocked: true }. Caution queries are logged but allowed.
 *
 * @param  {string} rawQuery â€” Original user query (not normalized).
 * @returns {Object} { blocked: bool, caution: bool, reason: string }
 */
function _aiDetectAbuse(rawQuery) {
  var q = rawQuery.toLowerCase();

  // --- Hard block list (per specification) ---
  var blocked = AI_CONFIG.ABUSE_TOKENS;
  for (var i = 0; i < blocked.length; i++) {
    // Use indexOf for phrase matching (handles multi-word patterns)
    if (q.indexOf(blocked[i]) !== -1) {
      return {
        blocked: true,
        caution: false,
        reason: 'Query matched blocked pattern: "' + blocked[i] + '"'
      };
    }
  }

  // --- Caution list (logged, allowed) ---
  var caution = AI_CONFIG.CAUTION_TOKENS;
  for (var j = 0; j < caution.length; j++) {
    // Word-boundary check: surround with spaces or start/end of string
    var pattern = '(?:^|\\s)' + caution[j] + '(?:\\s|$)';
    if (new RegExp(pattern).test(q)) {
      return {
        blocked: false,
        caution: true,
        reason: 'Caution token detected: "' + caution[j] + '"'
      };
    }
  }

  return { blocked: false, caution: false, reason: '' };
}

// -----------------------------------------------------------------------
// 9.8  QUERY NORMALISATION
// -----------------------------------------------------------------------

/**
 * _aiNormalizeQuery â€” Lowercases, strips punctuation except commas/hyphens,
 * collapses whitespace. Produces a clean token-ready string.
 *
 * @param  {string} raw
 * @returns {string} Normalized query.
 */
function _aiNormalizeQuery(raw) {
  return String(raw || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s,\-]/g, ' ')  // keep alphanumeric, spaces, commas, hyphens
    .replace(/\s+/g, ' ')             // collapse whitespace
    .trim();
}

// -----------------------------------------------------------------------
// 9.9  KNOWLEDGE SHEET LOADER (with CacheService)
// -----------------------------------------------------------------------

/**
 * _aiLoadKnowledgeSheet â€” Returns all rows from WMS_AI_Knowledge.
 * Cache-first: checks CacheService before hitting Sheets.
 * Falls back to direct read on cache miss or parse error.
 *
 * IMPORTANT: Does NOT use _getSheetValuesCached (request-scope cache).
 * The AI module manages its own persistent cache independently to avoid
 * polluting transactional request caches with large knowledge data.
 *
 * @returns {Array} 2D array of sheet values (row 0 = header).
 * @throws  {Error} If sheet is missing.
 */
function _aiLoadKnowledgeSheet() {
  var scriptCache = CacheService.getScriptCache();

  // --- Attempt cache read ---
  try {
    var cached = scriptCache.get(AI_CONFIG.CACHE_KEY);
    if (cached) {
      var parsed = JSON.parse(cached);
      if (Array.isArray(parsed) && parsed.length > 1) {
        return parsed;
      }
    }
  } catch (cacheReadErr) {
    Logger.log('[AI_CACHE_READ] Parse failed, falling through to sheet read: ' + cacheReadErr.message);
  }

  // --- Cache miss: read from sheet ---
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AI_CONFIG.KNOWLEDGE_SHEET);
  if (!sheet) {
    throw new Error('Knowledge sheet "' + AI_CONFIG.KNOWLEDGE_SHEET + '" not found. Please create it and populate SOPs.');
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    Logger.log('[AI_KB] Knowledge sheet is empty or header-only.');
    return values || [[]];
  }

  // --- Attempt cache write (may exceed 100KB for large knowledge bases â€” handled gracefully) ---
  try {
    var serialized = JSON.stringify(values);
    if (serialized.length <= 90000) {
      scriptCache.put(AI_CONFIG.CACHE_KEY, serialized, AI_CONFIG.CACHE_TTL_SECONDS);
    } else {
      // Knowledge base > 90KB: skip persistent cache, rely on request processing time
      // This happens at ~150+ SOPs with verbose content. At that point, consider
      // splitting into category-partitioned cache keys.
      Logger.log('[AI_CACHE_WRITE] Knowledge base exceeds 90KB (' + serialized.length + ' bytes). Persistent cache skipped.');
    }
  } catch (cacheWriteErr) {
    Logger.log('[AI_CACHE_WRITE] Failed: ' + cacheWriteErr.message);
  }

  return values;
}

/**
 * aiInvalidateKnowledgeCache â€” Forces cache refresh for WMS_AI_Knowledge.
 * Call this after adding/editing SOPs in the sheet.
 * Exposed as a callable function for manual trigger from script editor.
 */
function aiInvalidateKnowledgeCache() {
  try {
    CacheService.getScriptCache().remove(AI_CONFIG.CACHE_KEY);
    Logger.log('[AI_CACHE] Knowledge cache invalidated successfully.');
    return { success: true, message: 'Knowledge cache cleared. Next query will reload from sheet.' };
  } catch (e) {
    Logger.log('[AI_CACHE] Invalidation failed: ' + e.message);
    return { success: false, message: 'Cache invalidation failed: ' + e.message };
  }
}

// -----------------------------------------------------------------------
// 9.10  QUERY AUDIT LOG
// -----------------------------------------------------------------------

/**
 * _aiLogQuery â€” Appends one row to AI_Query_Log.
 * Auto-creates the sheet with correct headers if absent.
 * Never throws â€” logging failures must not break the query response.
 *
 * @param {Object} user           â€” getEffectiveUser() result.
 * @param {string} rawQuery       â€” Original query.
 * @param {string} normalizedQuery
 * @param {string} mode           â€” "RETRIEVAL" | "GEMINI" | "BLOCKED"
 * @param {Array}  sources        â€” Doc IDs returned.
 * @param {number} resultCount
 * @param {boolean} flagged
 * @param {string} flagReason
 * @param {number} responseTimeMs
 */
function _aiLogQuery(user, rawQuery, normalizedQuery, mode, sources, resultCount, flagged, flagReason, responseTimeMs) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(AI_CONFIG.QUERY_LOG_SHEET);

    // --- Auto-create log sheet if missing ---
    if (!logSheet) {
      logSheet = ss.insertSheet(AI_CONFIG.QUERY_LOG_SHEET);
      logSheet.appendRow(AI_CONFIG.LOG_COLS);
      logSheet.setFrozenRows(1);
      logSheet.getRange(1, 1, 1, AI_CONFIG.LOG_COLS.length)
        .setBackground('#4a4a8a')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      Logger.log('[AI_LOG] AI_Query_Log sheet created.');
    }

    var queryId = 'AIQ-' + new Date().getTime() + '-' +
      Math.random().toString(36).slice(2, 6).toUpperCase();

    var row = [
      queryId,
      new Date(),
      user.email || 'unknown',
      user.role || 'unknown',
      rawQuery.substring(0, 500),         // cap raw query
      normalizedQuery.substring(0, 500),  // cap normalized query
      mode,
      resultCount,
      (sources || []).join(', '),
      flagged ? 'TRUE' : 'FALSE',
      flagReason || '',
      responseTimeMs
    ];

    logSheet.appendRow(row);
  } catch (logErr) {
    // Logging failure is non-fatal. Log to Apps Script logger only.
    Logger.log('[AI_LOG_ERR] Failed to write query log: ' + logErr.message);
  }
}

// -----------------------------------------------------------------------
// 9.11  HELPER: SAFE ERROR RESPONSE
// -----------------------------------------------------------------------

/**
 * _aiSafeError â€” Builds a safe error response object.
 * Never exposes internal error messages to the client.
 *
 * @param  {string} code      â€” Internal error code (logged, not returned to client).
 * @param  {string} message   â€” User-safe message.
 * @param  {Array}  sources
 * @param  {number} startMs   â€” Request start timestamp for timing.
 * @returns {Object}
 */
function _aiSafeError(code, message, sources, startMs) {
  Logger.log('[AI_ERR] Code: ' + code + ' | Elapsed: ' + (new Date().getTime() - startMs) + 'ms');
  return {
    mode: 'ERROR',
    answer: message,
    sources: sources || [],
    flagged: false,
    flagReason: '',
    resultCount: 0
  };
}

// -----------------------------------------------------------------------
// 9.12  ANALYTICS (Manager-Only)
// -----------------------------------------------------------------------

/**
 * getAiQueryStats â€” Returns summary analytics for the AI query log.
 * Manager-only. Called by ManagerDashboard or AIAssistant (manager view).
 *
 * @returns {Object} { totalQueries, uniqueUsers, topQueries, flaggedCount, avgResponseMs, modeBreakdown }
 */
function getAiQueryStats() {
  // --- Auth: manager only ---
  var user = getEffectiveUser();
  if (ROLE_RANK[user.role] < ROLE_RANK[SECURITY.ROLES.MANAGER]) {
    throw new Error('Access denied: Manager role required for AI analytics.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AI_CONFIG.QUERY_LOG_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { totalQueries: 0, uniqueUsers: 0, topQueries: [], flaggedCount: 0, avgResponseMs: 0, modeBreakdown: {} };
  }

  var data = sheet.getDataRange().getValues();
  // Col indices match LOG_COLS order
  var COL = { TS: 1, EMAIL: 2, ROLE: 3, QUERY: 4, MODE: 6, COUNT: 7, FLAGGED: 9, RESPONSE_MS: 11 };

  var totalQueries = 0;
  var flaggedCount = 0;
  var totalMs = 0;
  var users = {};
  var queryCounts = {};
  var modeCounts = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[COL.TS]) continue;

    totalQueries++;
    users[String(row[COL.EMAIL])] = true;
    if (String(row[COL.FLAGGED]).toUpperCase() === 'TRUE') flaggedCount++;
    totalMs += Number(row[COL.RESPONSE_MS]) || 0;

    var q = String(row[COL.QUERY] || '').substring(0, 60);
    queryCounts[q] = (queryCounts[q] || 0) + 1;

    var mode = String(row[COL.MODE] || 'UNKNOWN');
    modeCounts[mode] = (modeCounts[mode] || 0) + 1;
  }

  // Top 10 queries by frequency
  var topQueries = Object.keys(queryCounts)
    .map(function (q) { return { query: q, count: queryCounts[q] }; })
    .sort(function (a, b) { return b.count - a.count; })
    .slice(0, 10);

  return {
    totalQueries: totalQueries,
    uniqueUsers: Object.keys(users).length,
    topQueries: topQueries,
    flaggedCount: flaggedCount,
    avgResponseMs: totalQueries > 0 ? Math.round(totalMs / totalQueries) : 0,
    modeBreakdown: modeCounts
  };
}

// =====================================================
// SECTION 13: HTML EMAIL ALERT ENGINE
//
// DEPLOYMENT:
//   1. Paste this entire file at the bottom of Code.js
//   2. In sendMinStockAlerts(), replace the MailApp.sendEmail line with
//      the call shown in the comment block below.
//
// CHANGES TO sendMinStockAlerts() (3 lines to change):
//
//   OLD:
//     const subject = 'WMS Min Stock Alert';
//     const body = 'The following items are at or below minimum stock:\n\n' + lines.join('\n');
//     MailApp.sendEmail(emails.join(','), subject, body);
//
//   NEW:
//     const subject = _buildAlertEmailSubject(notifiedItems.length);
//     const htmlBody = _buildAlertEmailHtml(notifiedItems, emails);
//     MailApp.sendEmail({
//       to: emails.join(','),
//       subject: subject,
//       body: _buildAlertEmailPlainText(notifiedItems), // plain-text fallback
//       htmlBody: htmlBody
//     });
//
// The notifiedItems array must be built instead of just pushing
// to lines[]. See the modified sendMinStockAlerts below.
// =====================================================

/**
 * REPLACEMENT FOR sendMinStockAlerts() in Code.js.
 *
 * This is a drop-in replacement that:
 * - Sends HTML email (with plain-text fallback)
 * - Builds structured items data instead of flat text lines
 * - One-per-day dedup is unchanged
 */
function sendMinStockAlerts() {
  return _withRequestCache(function () {
    try {
      const props = PropertiesService.getScriptProperties();
      const emails = _getMinStockAlertRecipients_();
      if (emails.length === 0) {
        Logger.log('Min stock alert skipped: ALERT_EMAILS not configured.');
        return;
      }

      // Business-facing alerts must match the warehouse item dashboard:
      // aggregate stock by item_code, even if legacy item_ids differ.
      const itemMaps = _getItemMasterMaps();
      const codeToMinStock = itemMaps.codeToMinStock || {};

      // Build inventory totals keyed by item_code from raw Inventory_Balance rows.
      const invByCode = {};
      {
        const ss2 = SpreadsheetApp.getActiveSpreadsheet();
        const invSheet2 = _getSheetOrThrow(ss2, CONFIG.SHEETS.INVENTORY);
        const invRows = _getSheetValuesCached(invSheet2.getName());
        for (let ii = 1; ii < invRows.length; ii++) {
          const qty = Number(invRows[ii][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
          if (qty <= 0) continue;
          const status = _normalizeQaStatus(invRows[ii][CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');
          if (status === 'REJECTED' || status === 'HOLD') continue;
          const rawItemId = String(invRows[ii][CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();
          const rawCode = String(invRows[ii][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || itemMaps.idToCode[rawItemId] || '').trim().toUpperCase();
          if (!rawCode) continue;
          const site = String(invRows[ii][CONFIG.INVENTORY_COLS.SITE] || '').trim();
          const location = String(invRows[ii][CONFIG.INVENTORY_COLS.LOCATION] || '').trim();
          if (!invByCode[rawCode]) invByCode[rawCode] = { totalQty: 0, bySiteLocation: {} };
          invByCode[rawCode].totalQty += qty;
          const slKey = site + '||' + location;
          invByCode[rawCode].bySiteLocation[slKey] = (invByCode[rawCode].bySiteLocation[slKey] || 0) + qty;
        }
      }

      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const notifiedItems = [];
      let belowThresholdCount = 0;
      let dedupedCount = 0;

      // Check each item_code that has min_stock configured and stock in inventory.
      Object.keys(codeToMinStock).forEach(function (codeNorm) {
        const minStock = Number(codeToMinStock[codeNorm] || 0);
        if (!(minStock > 0)) return;
        const invEntry = invByCode[codeNorm];
        if (!invEntry) return; // no stock in inventory at all — skip (nothing to alert on)
        const current = Number(invEntry.totalQty || 0);
        if (current > minStock) return;
        belowThresholdCount += 1;

        const lastKey = 'MIN_STOCK_ALERT_LAST_' + codeNorm;
        const lastSent = String(props.getProperty(lastKey) || '').trim();
        const legacySentToday = Object.keys(itemMaps.idToItemInfo || {}).some(function (idKey) {
          const legacyInfo = (itemMaps.idToItemInfo || {})[idKey];
          if (!legacyInfo || String(legacyInfo.codeNorm || '').trim().toUpperCase() !== codeNorm) return false;
          return String(props.getProperty('MIN_STOCK_ALERT_LAST_' + idKey) || '').trim() === today;
        });
        if (lastSent === today || legacySentToday) {
          dedupedCount += 1;
          return;
        }

        const bySite = invEntry.bySiteLocation || {};
        const locations = Object.keys(bySite).map(function (k) {
          const parts = k.split('||');
          return { site: parts[0] || '-', location: parts[1] || '-', qty: Number(bySite[k] || 0) };
        });
        notifiedItems.push({
          code: itemMaps.codeDisplayByNorm[codeNorm] || codeNorm,
          codeNorm: codeNorm,
          name: itemMaps.codeToName[codeNorm] || (itemMaps.codeDisplayByNorm[codeNorm] || codeNorm),
          current: current,
          min: minStock,
          uom: itemMaps.codeToUom[codeNorm] || 'KG',
          locations: locations
        });
      });

      if (notifiedItems.length === 0) {
        Logger.log(
          'Min stock alert: no items below threshold today. ' +
          'items_with_minstock=' + Object.keys(codeToMinStock).length +
          ', below_threshold=' + belowThresholdCount +
          ', deduped_today=' + dedupedCount
        );
        return;
      }

      const subject = _buildAlertEmailSubject(notifiedItems.length);
      const htmlBody = _buildAlertEmailHtml(notifiedItems, emails);
      const plainBody = _buildAlertEmailPlainText(notifiedItems);

      MailApp.sendEmail({
        to: emails.join(','),
        subject: subject,
        body: plainBody,   // plain-text fallback for old mail clients
        htmlBody: htmlBody
      });

      // Mark sent for today — key is item_id (codeNorm field stores item_id after fix)
      notifiedItems.forEach(function (item) {
        props.setProperty('MIN_STOCK_ALERT_LAST_' + String(item.codeNorm || '').trim().toUpperCase(), today);
      });

      Logger.log('Min stock alert sent for ' + notifiedItems.length + ' item(s).');
    } catch (e) {
      Logger.log('Min stock alert failed: ' + e.message);
    }
  });
}

function getMinStockAlertPreview() {
  // Uses item_code aggregation - same logic as sendMinStockAlerts.
  return _withRequestCache(function () {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const itemMaps = _getItemMasterMaps();
    const codeToMinStock = itemMaps.codeToMinStock || {};

    // Build inventory totals keyed by item_code, matching the email and dashboard.
    const invByCode = {};
    const ss2 = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet2 = _getSheetOrThrow(ss2, CONFIG.SHEETS.INVENTORY);
    const invRows = _getSheetValuesCached(invSheet2.getName());
    for (let ii = 1; ii < invRows.length; ii++) {
      const qty = Number(invRows[ii][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      if (qty <= 0) continue;
      const status = _normalizeQaStatus(invRows[ii][CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');
      if (status === 'REJECTED' || status === 'HOLD') continue;
      const rawItemId = String(invRows[ii][CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();
      const rawCode = String(invRows[ii][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || itemMaps.idToCode[rawItemId] || '').trim().toUpperCase();
      if (!rawCode) continue;
      if (!invByCode[rawCode]) invByCode[rawCode] = { totalQty: 0 };
      invByCode[rawCode].totalQty += qty;
    }

    const rows = [];
    Object.keys(codeToMinStock).forEach(function (codeNorm) {
      const minStock = Number(codeToMinStock[codeNorm] || 0);
      if (!(minStock > 0)) return;
      const invEntry = invByCode[codeNorm];
      if (!invEntry) return;
      const current = Number(invEntry.totalQty || 0);
      const lastSent = String(props.getProperty('MIN_STOCK_ALERT_LAST_' + codeNorm) || '').trim();
      const legacySentToday = Object.keys(itemMaps.idToItemInfo || {}).some(function (idKey) {
        const legacyInfo = (itemMaps.idToItemInfo || {})[idKey];
        if (!legacyInfo || String(legacyInfo.codeNorm || '').trim().toUpperCase() !== codeNorm) return false;
        return String(props.getProperty('MIN_STOCK_ALERT_LAST_' + idKey) || '').trim() === today;
      });
      rows.push({
        itemCode: itemMaps.codeDisplayByNorm[codeNorm] || codeNorm,
        current: current,
        min: minStock,
        uom: itemMaps.codeToUom[codeNorm] || 'KG',
        belowOrEqualMin: current <= minStock,
        alreadySentToday: lastSent === today || legacySentToday,
        lastSentDate: lastSent || ''
      });
    });

    const below = rows.filter(function (r) { return r.belowOrEqualMin; });
    const willSend = below.filter(function (r) { return !r.alreadySentToday; });

    return {
      today: today,
      inventoryTrackedItems: Object.keys(invByCode).length,
      withMinConfigInInventory: rows.length,
      belowThresholdCount: below.length,
      willSendCount: willSend.length,
      items: rows.sort(function (a, b) { return String(a.itemCode).localeCompare(String(b.itemCode)); })
    };
  });
}

function _clearMinStockAlertDedupKeys(itemCodes) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let targetKeys = [];

  if (Array.isArray(itemCodes) && itemCodes.length > 0) {
    const norms = itemCodes.map(function (c) {
      return String(c || '').trim().toUpperCase();
    }).filter(Boolean);
    targetKeys = norms.map(function (code) { return 'MIN_STOCK_ALERT_LAST_' + code; });
  } else {
    targetKeys = Object.keys(all).filter(function (k) {
      return String(k || '').indexOf('MIN_STOCK_ALERT_LAST_') === 0;
    });
  }

  const deleted = [];
  targetKeys.forEach(function (k) {
    if (Object.prototype.hasOwnProperty.call(all, k)) {
      props.deleteProperty(k);
      deleted.push(k);
    }
  });
  return deleted;
}

function clearMinStockAlertDedup(itemCodes) {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);
    const deleted = _clearMinStockAlertDedupKeys(itemCodes);
    return {
      success: true,
      deletedCount: deleted.length,
      deletedKeys: deleted
    };
  });
}

function getMinStockAlertStatus() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);
    const props = PropertiesService.getScriptProperties();
    const emails = _getMinStockAlertRecipients_();
    const dedupKeys = Object.keys(props.getProperties()).filter(function (k) {
      return String(k || '').indexOf('MIN_STOCK_ALERT_LAST_') === 0;
    });
    return {
      recipientsConfigured: emails.length,
      recipients: emails,
      dedupKeysCount: dedupKeys.length,
      preview: getMinStockAlertPreview()
    };
  });
}


// ─────────────────────────────────────────────────────────────────────
// 13.1  SUBJECT LINE
// ─────────────────────────────────────────────────────────────────────
function _buildAlertEmailSubject(count) {
  const tz = Session.getScriptTimeZone();
  const date = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy');
  return '⚠ WMS Stock Alert — ' + count + ' item' + (count === 1 ? '' : 's') + ' below minimum · ' + date;
}


// ─────────────────────────────────────────────────────────────────────
// 13.2  PLAIN-TEXT FALLBACK
// ─────────────────────────────────────────────────────────────────────
function _buildAlertEmailPlainText(items) {
  const tz = Session.getScriptTimeZone();
  const date = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm');
  const lines = ['WMS MIN STOCK ALERT — ' + date, '='.repeat(50), ''];
  items.forEach(function (item) {
    const pct = item.min > 0 ? Math.round((item.current / item.min) * 100) : 0;
    lines.push(item.code + ' — ' + item.name);
    lines.push('  Current: ' + item.current + ' ' + item.uom + '  |  Minimum: ' + item.min + ' ' + item.uom + '  (' + pct + '% of min)');
    if (item.locations && item.locations.length > 0) {
      item.locations.forEach(function (loc) {
        lines.push('  · ' + loc.site + ' / ' + loc.location + ': ' + loc.qty + ' ' + item.uom);
      });
    }
    lines.push('');
  });
  lines.push('This alert is sent once per day per item.');
  return lines.join('\n');
}


// ─────────────────────────────────────────────────────────────────────
// 13.3  HTML EMAIL BUILDER
// ─────────────────────────────────────────────────────────────────────
function _buildAlertEmailHtml(items, emails) {
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy');
  let dashUrl = '';
  try { dashUrl = ScriptApp.getService().getUrl(); } catch (e) { }

  // ── Build each item card ──────────────────────────────────────────
  const itemsHtml = items.map(function (item) {
    const pct = item.min > 0 ? Math.min(100, Math.round((item.current / item.min) * 100)) : 0;
    const isCritical = item.current === 0 || pct <= 20;
    const fillClass = isCritical ? 'red' : 'amber';

    // Location chips
    let locHtml = '';
    if (item.locations && item.locations.length > 0) {
      const chips = item.locations.map(function (loc) {
        return '<span class="loc-chip">' + _esc(loc.site) + ' / ' + _esc(loc.location) +
          ' <span class="qty-inline">- ' + _fmtAlertQty(loc.qty, item.uom) + '</span></span>';
      }).join('');
      locHtml = '<div class="section-label" style="margin-top:10px;margin-bottom:6px">Location - </div>' +
        '<div class="locations">' + chips + '</div>';
    } else {
      locHtml = '<div class="loc-chip" style="margin-top:8px;background:#FEF2F2;border-color:#FECACA;color:#B91C1C">No stock in any bin</div>';
    }

    return [
      '<div class="item-card">',
      '<div class="item-head">',
      '<div>',
      '<div class="item-code">' + _esc(item.code) + '</div>',
      '<div class="item-name">' + _esc(item.name) + '</div>',
      '</div>',
      '</div>',
      '<div class="item-body">',
      '<div class="qty-row">',
      '<div class="qty-box">',
      '<div class="label">Current Stock</div>',
      '<div class="value danger">' + _fmtAlertQty(item.current, item.uom) + '</div>',
      '</div>',
      '<div class="qty-box">',
      '<div class="label">Minimum Level</div>',
      '<div class="value min">' + _fmtAlertQty(item.min, item.uom) + '</div>',
      '</div>',
      '</div>',
      '<div class="fill-bar-wrap">',
      '<div class="fill-label">',
      '<span>Stock level vs minimum</span>',
      '<span>' + pct + '%</span>',
      '</div>',
      '<div class="fill-bar">',
      '<div class="fill ' + fillClass + '" style="width:' + pct + '%"></div>',
      '</div>',
      '</div>',
      locHtml,
      '</div>',
      '</div>'
    ].join('\n');
  }).join('\n');

  // ── Assemble full email ───────────────────────────────────────────
  const recipientDisplay = emails.length > 2
    ? emails.slice(0, 2).join(', ') + ' +' + (emails.length - 2) + ' more'
    : emails.join(', ');

  return [
    '<!DOCTYPE html><html lang="en">',
    '<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">',
    '<style>',
    '*{margin:0;padding:0;box-sizing:border-box;}',
    'body{background:#F0F2F5;font-family:Segoe UI,Arial,sans-serif;-webkit-text-size-adjust:100%;}',
    '.wrapper{width:100%;background:#F0F2F5;padding:32px 16px;}',
    '.card{max-width:620px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08);}',
    '.header{background:#1E1E2E;padding:24px 28px;display:flex;align-items:center;gap:14px;}',
    '.logo{width:44px;height:44px;background:#B45309;border-radius:8px; display:flex;align-items:center;justify-content:center;font-size:33px;flex-shrink:0;}',
    '.hd-title{color:#fff;font-size:1.05rem;font-weight:700; margin-left:10px;}',
    '.hd-sub{color:#A0A0B8;font-size:0.76rem;margin-left:10px; margin-top:2px;}',
    '.banner{background:#FEF9EC;border-left:5px solid #F59E0B;padding:14px 28px;display:flex;align-items:center;gap:12px;}',
    '.ban-icon{font-size:1.5rem;}',
    '.ban-hl{color:#92400E;font-size:.95rem;font-weight:700;}',
    '.ban-sub{color:#B45309;font-size:.75rem;margin-top:2px;}',
    '.body{padding:24px 28px;}',
    '.sec-label{font-size:.63rem;font-weight:800;letter-spacing:1.2px;text-transform:uppercase;color:#9CA3AF;margin-bottom:12px;}',
    '.item-card{border:1px solid #E5E7EB;border-radius:12px;margin-bottom:12px;overflow:hidden;}',
    '.item-head{background:#F9FAFB;padding:11px 15px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #E5E7EB;}',
    '.item-code{font-size:.86rem;font-weight:800;color:#111827;}',
    '.item-name{font-size:.7rem;color:#6B7280;margin-top:2px;}',
    '.badge{border-radius:6px;padding:3px 9px;font-size:.68rem;font-weight:700;}',
    '.badge-crit{background:#FEE2E2;color:#B91C1C;}',
    '.badge-low{background:#FEF3C7;color:#92400E;}',
    '.item-body{padding:12px 15px;}',
    '.qty-row{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px;}',
    '.qty-box{background:#F9FAFB;border-radius:8px;padding:8px 11px;}',
    '.ql{font-size:.6rem;text-transform:uppercase;letter-spacing:.8px;color:#9CA3AF;font-weight:700;margin-bottom:2px;}',
    '.qv{font-size:.95rem;font-weight:800;color:#111827;}',
    '.qv.danger{color:#DC2626;}',
    '.qv.min{color:#6B7280;}',
    '.uom{font-size:.64rem;color:#9CA3AF;margin-left:2px;}',
    '.fill-wrap{margin-bottom:8px;}',
    '.fill-lbl{display:flex;justify-content:space-between;font-size:.66rem;color:#6B7280;margin-bottom:3px;}',
    '.fill-bar{height:7px;background:#E5E7EB;border-radius:8px;overflow:hidden;}',
    '.fill{height:100%;border-radius:8px;}',
    '.fill-red{background:linear-gradient(90deg,#EF4444,#DC2626);}',
    '.fill-amber{background:linear-gradient(90deg,#F59E0B,#D97706);}',
    '.locs{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px;}',
    '.chip{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:5px;padding:2px 8px;font-size:.66rem;color:#1E40AF;font-weight:600;}',
    '.chip-none{background:#FEF2F2;border-color:#FECACA;color:#B91C1C;}',
    '.chip .qi{color:#6B7280;font-weight:400;}',
    'hr{border:none;border-top:1px solid #F3F4F6;margin:20px 0;}',
    '.footer{padding:0 28px 24px;}',
    '.cta{display:block;background:#1E1E2E;color:#fff!important;text-align:center;text-decoration:none;',
    '     border-radius:9px;padding:11px;font-size:.82rem;font-weight:700;margin-bottom:18px;}',
    '.ts{font-size:.67rem;color:#9CA3AF;text-align:center;line-height:1.5;}',
    '@media(max-width:480px){.qty-row{grid-template-columns:1fr;}}',
    '</style></head>',
    '<body><div class="wrapper"><div class="card">',

    // Header
    '<div class="header">',
    '<span class="logo">⚠</span>',
    '<div><div class="hd-title">Cultivator Natural Products WMS</div><div class="hd-sub">Warehouse Management System — Automated Email Alert</div></div>',
    '</div>',

    // Banner
    '<div class="banner">',
    '<div><div class="ban-hl">' + items.length + ' item' + (items.length === 1 ? '' : 's') + ' at or below minimum stock level</div>',
    '<div class="ban-sub">Generated on ' + dateStr + ' · Immediate replenishment recommended</div></div>',
    '</div>',

    // Items
    '<div class="body">',
    '<div class="sec-label">Items Requiring Attention</div>',
    itemsHtml,
    '</div>',

    '<hr>',
    '<div class="footer">',
    '<p class="ts">Sent once per day per item when stock ≤ minimum.<br>',
    'Recipients: ' + _esc(recipientDisplay) + '<br>',
    'WMS v5.0 · Cultivator Natural Products · Do not reply to this automated message.</p>',
    '</div>',

    '</div></div></body></html>'
  ].join('');
}


// ─────────────────────────────────────────────────────────────────────
// 13.4  HTML ESCAPE HELPER (safe for email insertion)
// ─────────────────────────────────────────────────────────────────────
function _esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
// ─────────────────────────────────────────────────────────────────────
// END SECTION 13
// ─────────────────────────────────────────────────────────────────────

function _fmtAlertQty(qty, uom) {
  const num = Number(qty) || 0;
  const shown = Math.abs(num % 1) < 0.000001 ? String(Math.round(num)) : num.toFixed(2);
  return shown + (uom ? ' ' + _esc(uom) : '');
}

function _appendInventoryRow(rowArray, sheet) {
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  }
  sheet.appendRow(rowArray);
}

function _normalizeExpiryDateForReports(value) {
  if (!value) return _getDefaultInwardExpiryDate();
  const parsed = value instanceof Date ? value : new Date(value);
  if (!(parsed instanceof Date) || isNaN(parsed.getTime()) || parsed.getTime() === 0) {
    return _getDefaultInwardExpiryDate();
  }
  return parsed;
}

/**
 * REPORTS SCREEN BACKEND FUNCTIONS
 * Add these 5 functions to the bottom of Code.js
 * These replace the broken AlertsScreen and DeadStockScreen functions
 */

// ============================================================================
// FUNCTION 1: DEAD STOCK REPORT
// Reads from Dead_Stock_Log sheet (primary) and Movement_Audit_Log (fallback)
// ============================================================================
function getDeadStockReportData() {
  return protect(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rows = [];
    const byItem = {};
    let totalQty = 0;

    // --- PRIMARY: Read from Dead_Stock_Log sheet ---
    var dsSheet = _resolveDeadStockSheet(ss);
    if (dsSheet) {
      var dsData = dsSheet.getDataRange().getValues();
      var dsHeader = (dsData[0] || []).map(function (h) { return String(h || '').trim().toUpperCase(); });
      var colMap = {};
      dsHeader.forEach(function (h, i) { colMap[h] = i; });
      // Detect key columns robustly
      var cItemCode = colMap['ITEM_CODE'] !== undefined ? colMap['ITEM_CODE'] : (colMap['ITEM_ID'] !== undefined ? colMap['ITEM_ID'] : 1);
      var cBatch = colMap['BATCH_ID'] !== undefined ? colMap['BATCH_ID'] : 2;
      var cQty = colMap['QUANTITY'] !== undefined ? colMap['QUANTITY'] : (colMap['TOTAL_QUANTITY'] !== undefined ? colMap['TOTAL_QUANTITY'] : 3);
      var cUom = colMap['UOM'] !== undefined ? colMap['UOM'] : 4;
      var cReason = colMap['REASON'] !== undefined ? colMap['REASON'] : (colMap['REMARKS'] !== undefined ? colMap['REMARKS'] : 5);
      var cTs = colMap['TIMESTAMP'] !== undefined ? colMap['TIMESTAMP'] : (colMap['LAST_UPDATED'] !== undefined ? colMap['LAST_UPDATED'] : 0);
      var cPo = colMap['PROD_ORDER_REF'] !== undefined ? colMap['PROD_ORDER_REF'] : (colMap['PO_NO'] !== undefined ? colMap['PO_NO'] : -1);

      for (var i = 1; i < dsData.length; i++) {
        var row = dsData[i];
        var qty = Number(row[cQty]) || 0;
        if (qty <= 0) continue;
        var itemCode = String(row[cItemCode] || '').trim();
        if (!itemCode) continue;
        var batchId = String(row[cBatch] || '').trim();
        var uom = String(row[cUom] || 'NOS').trim();
        var reason = String(row[cReason] || 'Not specified').trim();
        var poNo = cPo >= 0 ? String(row[cPo] || '').trim() : '';
        var ts = row[cTs];
        var tsStr = ts instanceof Date ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd-MMM-yyyy') : String(ts || '');
        var record = { timestamp: tsStr, prodOrderRef: poNo, itemCode: itemCode, batchId: batchId, quantity: qty, uom: uom, reason: reason };
        rows.push(record);
        totalQty += qty;
        if (!byItem[itemCode]) byItem[itemCode] = { itemCode: itemCode, totalQty: 0, count: 0 };
        byItem[itemCode].totalQty += qty;
        byItem[itemCode].count++;
      }
    }

    // --- FALLBACK: Movement_Audit_Log with PROD_RETURN_DEAD ---
    if (rows.length === 0) {
      var movSheet = ss.getSheetByName(CONFIG.SHEETS.MOVEMENT);
      if (movSheet) {
        var movData = movSheet.getDataRange().getValues();
        for (var j = 1; j < movData.length; j++) {
          var mrow = movData[j];
          var movType = String(mrow[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').trim().toUpperCase();
          if (movType !== 'PROD_RETURN_DEAD') continue;
          var mQty = Number(mrow[CONFIG.MOVEMENT_COLS.QUANTITY]) || 0;
          if (mQty <= 0) continue;
          var mItemId = String(mrow[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
          var mCode = mItemId;
          try { if (!isNaN(mItemId) && parseInt(mItemId) > 0) mCode = _getItemCodeById(parseInt(mItemId)) || mItemId; } catch (e) { }
          var mTs = mrow[CONFIG.MOVEMENT_COLS.TIMESTAMP];
          var mTsStr = mTs instanceof Date ? Utilities.formatDate(mTs, Session.getScriptTimeZone(), 'dd-MMM-yyyy') : String(mTs || '');
          var mRemarks = String(mrow[CONFIG.MOVEMENT_COLS.REMARKS] || '').trim();
          var mReason = 'Not specified';
          if (mRemarks.indexOf('DEAD_STOCK:') !== -1) { var parts = mRemarks.split('DEAD_STOCK:'); if (parts[1]) mReason = parts[1].split('|')[0].trim(); }
          var mRecord = { timestamp: mTsStr, prodOrderRef: String(mrow[CONFIG.MOVEMENT_COLS.PROD_ORDER_REF] || ''), itemCode: mCode, batchId: String(mrow[CONFIG.MOVEMENT_COLS.BATCH_ID] || ''), quantity: mQty, uom: String(mrow[CONFIG.MOVEMENT_COLS.UOM] || 'KG'), reason: mReason };
          rows.push(mRecord);
          totalQty += mQty;
          if (!byItem[mCode]) byItem[mCode] = { itemCode: mCode, totalQty: 0, count: 0 };
          byItem[mCode].totalQty += mQty;
          byItem[mCode].count++;
        }
      }
    }

    return { rows: rows, byItem: Object.keys(byItem).map(function (k) { return byItem[k]; }), totalQty: totalQty };
  });
}


// ============================================================================
// FUNCTION 2: REJECTED ITEMS REPORT
// ============================================================================
function getRejectedItemsReportData() {
  return protect(function () {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const qaSheet = ss.getSheetByName(CONFIG.SHEETS.QA_EVENTS);
      const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);

      if (!qaSheet) {
        return { rejections: [], summary: [], _error: 'QA_Events sheet not found' };
      }

      // Load QA Events
      const qaData = qaSheet.getDataRange().getValues();
      const qaCols = CONFIG.QA_EVENTS_COLS;

      // Load Inventory for historical quantity lookup
      const invData = invSheet ? invSheet.getDataRange().getValues() : [];
      const invMap = {};
      if (invData.length > 1) {
        for (let j = 1; j < invData.length; j++) {
          const invId = String(invData[j][CONFIG.INVENTORY_COLS.INVENTORY_ID] || '');
          const invQty = Number(invData[j][CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
          if (invId) invMap[invId] = invQty;
        }
      }

      const grouped = {};

      for (let i = 1; i < qaData.length; i++) {
        const row = qaData[i];
        const action = String(row[qaCols.ACTION] || '').trim().toUpperCase();
        const status = String(row[qaCols.NEW_STATUS] || '').trim().toUpperCase();
        const invId = String(row[qaCols.INVENTORY_ID] || '').trim();
        const itemCode = String(row[qaCols.ITEM_CODE] || '').trim();
        const batchId = String(row[qaCols.BATCH_ID] || '').trim();
        const binId = String(row[qaCols.BIN_ID] || '').trim();

        // Get quantity: try new column first, then historical lookup
        let qty = 0;
        if (typeof qaCols.QUANTITY === 'number' && row[qaCols.QUANTITY] > 0) {
          qty = Number(row[qaCols.QUANTITY]);
        } else if (invId && invMap[invId]) {
          qty = invMap[invId];
        }

        // Use Overridden At or Timestamp
        let ts = row[qaCols.OVERRIDDEN_AT] || row[qaCols.TIMESTAMP];
        if (!ts || !(ts instanceof Date)) ts = new Date();

        // Grouping key: Item + Batch + Timestamp (truncated to minutes)
        const tsMinute = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyyMMddHHmm');
        const groupKey = `${itemCode}|${batchId}|${tsMinute}`;

        if (!grouped[groupKey]) {
          grouped[groupKey] = {
            timestamp: Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm'),
            itemCode: itemCode,
            batchId: batchId,
            binId: binId,
            approvedQty: 0,
            rejectedQty: 0,
            reason: String(row[qaCols.OVERRIDE_REASON] || '').trim(),
            remarks: String(row[qaCols.REMARKS] || '').trim(),
            user: String(row[qaCols.OVERRIDDEN_BY] || '').trim(),
            hasRejection: false
          };
        }

        const g = grouped[groupKey];

        // Categorize qty based on action/status
        if (action === 'REJECT' || action === 'REJECTED' || status === 'REJECTED') {
          g.rejectedQty += qty;
          g.hasRejection = true;
          // Capture reason for rejection specifically
          const r = String(row[qaCols.OVERRIDE_REASON] || '').trim();
          if (r && r !== '-') g.reason = r;
        } else if (action === 'APPROVE' || action === 'APPROVE_PARTIAL' || status === 'APPROVED') {
          g.approvedQty += qty;
        }
      }

      // Filter: Only show "Rejected Items" flows (where at least one rejection occurred)
      const finalReport = Object.keys(grouped)
        .map(k => grouped[k])
        .filter(g => g.hasRejection)
        .sort((a, b) => b.timestamp.localeCompare(a.timestamp));

      return {
        rejections: finalReport,
        summary: []
      };
    } catch (e) {
      Logger.log('[REJECTED_REPORT] Error: ' + e.toString());
      return { rejections: [], _error: e.toString() };
    }
  });
}


// ============================================================================
// FUNCTION 3: EXPIRED STOCK REPORT
// ============================================================================
function getExpiredStockReportData() {
  return protect(function () {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
      if (!invSheet) return { expired: [] };

      var data = invSheet.getDataRange().getValues();
      var expired = [];
      var now = new Date();
      var tz = Session.getScriptTimeZone();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        var expiryRaw = row[CONFIG.INVENTORY_COLS.EXPIRY_DATE];
        if (qty <= 0 || !expiryRaw) continue;
        var expiry = _normalizeExpiryDateForReports(expiryRaw);
        if (expiry < now) {
          var daysOverdue = Math.floor((now - expiry) / 86400000);
          expired.push({
            itemCode: String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim(),
            batchId: String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
            binId: String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim(),
            quantity: qty,
            uom: String(row[CONFIG.INVENTORY_COLS.UOM] || 'KG').trim(),
            expiryDate: Utilities.formatDate(expiry, tz, 'dd-MMM-yyyy'),
            daysOverdue: daysOverdue
          });
        }
      }
      return { expired: expired };
    } catch (e) {
      Logger.log('[EXPIRED_REPORT] Error: ' + e.message);
      return { expired: [], _error: e.message };
    }
  });
}


// ============================================================================
// FUNCTION 4: EXPIRING SOON REPORT
// ============================================================================
function getExpiringStockReportData() {
  return protect(function () {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
      if (!invSheet) return { expiring: [] };

      var data = invSheet.getDataRange().getValues();
      var expiring = [];
      var now = new Date();
      var thirtyDaysFromNow = new Date(now.getTime() + 30 * 86400000);
      var tz = Session.getScriptTimeZone();

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        var expiryRaw = row[CONFIG.INVENTORY_COLS.EXPIRY_DATE];
        if (qty <= 0 || !expiryRaw) continue;
        var expiry = _normalizeExpiryDateForReports(expiryRaw);
        if (expiry > now && expiry <= thirtyDaysFromNow) {
          var daysUntilExpiry = Math.ceil((expiry - now) / 86400000);
          expiring.push({
            itemCode: String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim(),
            batchId: String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
            binId: String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim(),
            quantity: qty,
            uom: String(row[CONFIG.INVENTORY_COLS.UOM] || 'KG').trim(),
            expiryDate: Utilities.formatDate(expiry, tz, 'dd-MMM-yyyy'),
            daysUntilExpiry: daysUntilExpiry
          });
        }
      }
      return { expiring: expiring };
    } catch (e) {
      Logger.log('[EXPIRING_REPORT] Error: ' + e.message);
      return { expiring: [], _error: e.message };
    }
  });
}


// ============================================================================
// FUNCTION 5: LOW STOCK REPORT
// ============================================================================
function getLowStockReportData() {
  return protect(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    const itemSheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);

    if (!invSheet || !itemSheet) {
      return { lowStock: [] };
    }

    const invData = invSheet.getDataRange().getValues();
    const itemData = itemSheet.getDataRange().getValues();

    const minStockMap = {};
    const itemNameMap = {};
    for (let i = 1; i < itemData.length; i++) {
      const code = String(itemData[i][CONFIG.ITEM_COLS.ITEM_CODE] || '').trim().toUpperCase();
      const minStock = Number(itemData[i][CONFIG.ITEM_COLS.MIN_STOCK_LEVEL]) || 0;
      const itemName = String(itemData[i][CONFIG.ITEM_COLS.ITEM_NAME] || '').trim();
      minStockMap[code] = minStock;
      itemNameMap[code] = itemName;
    }

    const stockByItem = {};
    for (let i = 1; i < invData.length; i++) {
      const row = invData[i];
      const itemCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim().toUpperCase();
      const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
      const status = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').toUpperCase();

      if (status === 'APPROVED' || status === 'OVERRIDDEN') {
        if (!stockByItem[itemCode]) stockByItem[itemCode] = 0;
        stockByItem[itemCode] += qty;
      }
    }

    const lowStock = [];
    Object.keys(minStockMap).forEach(itemCode => {
      const minStock = minStockMap[itemCode];
      const currentStock = stockByItem[itemCode] || 0;

      if (minStock > 0 && currentStock < minStock) {
        const deficit = minStock - currentStock;
        const percentBelow = Math.round((deficit / minStock) * 100);

        let severity = 'info';
        if (percentBelow > 80) severity = 'critical';
        else if (percentBelow > 50) severity = 'warning';

        lowStock.push({
          itemCode: itemCode,
          itemName: itemNameMap[itemCode] || '',
          currentStock: currentStock,
          minStock: minStock,
          deficit: deficit,
          percentBelow: percentBelow,
          severity: severity
        });
      }
    });

    return { lowStock: lowStock };
  });
}

function _getMovementReportWindow(referenceDate, period, customStart, customEnd) {
  if (period === 'CUSTOM' && customStart && customEnd) {
    const start = new Date(customStart);
    const end = new Date(customEnd);
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    // Cap at 6 months
    const diffMs = end.getTime() - start.getTime();
    const sixMonthsMs = 183 * 24 * 60 * 60 * 1000;
    if (diffMs > sixMonthsMs) {
      throw new Error('Date range cannot exceed 6 months.');
    }
    return { period: 'CUSTOM', start: start, end: end };
  }

  const ref = referenceDate ? new Date(referenceDate) : new Date();
  const safeRef = (ref instanceof Date && !isNaN(ref.getTime())) ? ref : new Date();
  const start = new Date(safeRef);
  const end = new Date(safeRef);
  const mode = String(period || 'DAILY').trim().toUpperCase();

  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);

  if (mode === 'MONTHLY') {
    start.setDate(1);
    end.setMonth(start.getMonth() + 1, 0);
    end.setHours(23, 59, 59, 999);
    return { period: 'MONTHLY', start: start, end: end };
  }

  if (mode === 'WEEKLY') {
    const day = start.getDay();
    const diffToMonday = (day + 6) % 7;
    start.setDate(start.getDate() - diffToMonday);
    end.setTime(start.getTime());
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    return { period: 'WEEKLY', start: start, end: end };
  }

  return { period: 'DAILY', start: start, end: end };
}

function _getMaterialMovementFlow(movementType) {
  const type = String(movementType || '').trim().toUpperCase();
  if (type === CONFIG.MOVEMENT_TYPES.INWARD || type === CONFIG.MOVEMENT_TYPES.RETURN_PROD) return 'INWARD';
  if (type === CONFIG.MOVEMENT_TYPES.DISPATCH || type === CONFIG.MOVEMENT_TYPES.CONSUMPTION || type === CONFIG.MOVEMENT_TYPES.INWARD_REVERSAL) return 'OUTWARD';
  if (type === CONFIG.MOVEMENT_TYPES.TRANSFER || type === CONFIG.MOVEMENT_TYPES.TRANSFER_QUARANTINE) return 'TRANSFER';
  return '';
}

function getMaterialMovementReportData(filters) {
  return protect(function () {
    const opts = filters || {};
    const requestedFlow = String(opts.flow || 'ALL').trim().toUpperCase();
    const windowInfo = _getMovementReportWindow(opts.referenceDate, opts.period, opts.startDate, opts.endDate);
    const start = windowInfo.start;
    const end = windowInfo.end;
    const tz = Session.getScriptTimeZone();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.MOVEMENT);
    const data = _getSheetValuesCached(sheet.getName());
    const itemMaps = _getItemMasterMaps();
    const binMeta = _buildBinMetaMap(ss);
    const currentStockNow = (function () {
      const exact = {};
      const byBin = {};
      const byItem = {};
      const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
      const invRows = _getSheetValuesCached(invSheet.getName());
      for (let r = 1; r < invRows.length; r++) {
        const inv = invRows[r];
        const qty = Number(inv[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        if (qty <= 0) continue;
        const qa = _normalizeQaStatus(inv[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING');
        if (qa !== 'APPROVED' && qa !== 'OVERRIDDEN') continue;
        const invItemId = String(inv[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim().toUpperCase();
        const code = String(inv[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || itemMaps.idToCode[invItemId] || '').trim().toUpperCase();
        const batch = String(inv[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim().toUpperCase();
        const bin = String(inv[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim().toUpperCase();
        if (!code) continue;
        const exactKey = [code, batch, bin].join('|');
        const binKey = [code, bin].join('|');
        exact[exactKey] = (exact[exactKey] || 0) + qty;
        byBin[binKey] = (byBin[binKey] || 0) + qty;
        byItem[code] = (byItem[code] || 0) + qty;
      }
      return { exact: exact, byBin: byBin, byItem: byItem };
    })();
    const rows = [];
    let totalQty = 0;
    let totalToProduction = 0;
    const itemSet = {};
    const flowCounts = { INWARD: 0, OUTWARD: 0, TRANSFER: 0 };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tsRaw = row[CONFIG.MOVEMENT_COLS.TIMESTAMP];
      const ts = tsRaw instanceof Date ? tsRaw : new Date(tsRaw);
      if (!(ts instanceof Date) || isNaN(ts.getTime())) continue;
      if (ts < start || ts > end) continue;

      const movementType = String(row[CONFIG.MOVEMENT_COLS.MOVEMENT_TYPE] || '').trim().toUpperCase();
      const flow = _getMaterialMovementFlow(movementType);
      if (!flow) continue;
      if (requestedFlow !== 'ALL' && requestedFlow !== flow) continue;

      const itemId = String(row[CONFIG.MOVEMENT_COLS.ITEM_ID] || '').trim();
      const itemCodeResolved = _getCanonicalItemCode(itemMaps.idToCode[itemId.toUpperCase()] || itemId, itemId);
      const itemCode = String(itemCodeResolved || '').trim();
      const codeNorm = itemCode.toUpperCase();
      const itemName = String(itemMaps.codeToName[codeNorm] || itemCode || itemId);
      const batchId = String(row[CONFIG.MOVEMENT_COLS.BATCH_ID] || '').trim();
      const batchNumber = _getBatchDisplayNumber(itemCode, batchId);
      const qty = Number(row[CONFIG.MOVEMENT_COLS.QUANTITY]) || 0;
      const uom = String(row[CONFIG.MOVEMENT_COLS.UOM] || itemMaps.codeToUom[codeNorm] || 'KG').trim() || 'KG';
      const prodOrderRef = String(row[CONFIG.MOVEMENT_COLS.PROD_ORDER_REF] || '').trim();
      const fromBinId = String(row[CONFIG.MOVEMENT_COLS.FROM_BIN_ID] || '').trim();
      const toBinId = String(row[CONFIG.MOVEMENT_COLS.TO_BIN_ID] || '').trim();
      const fromMeta = binMeta[fromBinId] || {};
      const toMeta = binMeta[toBinId] || {};
      const primaryMeta = flow === 'OUTWARD' ? fromMeta : toMeta;
      const primaryBinId = flow === 'OUTWARD' ? fromBinId : toBinId;
      const toProductionQty = movementType === CONFIG.MOVEMENT_TYPES.CONSUMPTION ? qty : 0;
      const stockCode = codeNorm;
      const stockBatch = batchId.toUpperCase();
      const stockBin = String(primaryBinId || '').trim().toUpperCase();
      const exactStockKey = [stockCode, stockBatch, stockBin].join('|');
      const binStockKey = [stockCode, stockBin].join('|');
      const currentStockInBinNow = stockBin
        ? (currentStockNow.exact[exactStockKey] != null ? currentStockNow.exact[exactStockKey] : (currentStockNow.byBin[binStockKey] || 0))
        : (currentStockNow.byItem[stockCode] || 0);
      const currentItemStockNow = currentStockNow.byItem[stockCode] || 0;

      rows.push({
        timestamp: ts.toISOString(),
        dateLabel: Utilities.formatDate(ts, tz, 'dd-MMM-yyyy HH:mm'),
        flow: flow,
        movementType: movementType,
        itemCode: itemCode,
        itemName: itemName,
        batchId: batchId,
        batchNumber: batchNumber,
        quantity: qty,
        currentStockNow: currentStockInBinNow,
        currentItemStockNow: currentItemStockNow,
        uom: uom,
        site: String(primaryMeta.site || ''),
        location: String(primaryMeta.location || ''),
        binId: String(primaryBinId || ''),
        binCode: String(primaryMeta.binCode || primaryBinId || ''),
        fromSite: String(fromMeta.site || ''),
        fromLocation: String(fromMeta.location || ''),
        fromBinId: fromBinId,
        fromBinCode: String(fromMeta.binCode || fromBinId || ''),
        toSite: String(toMeta.site || ''),
        toLocation: String(toMeta.location || ''),
        toBinId: toBinId,
        toBinCode: String(toMeta.binCode || toBinId || ''),
        prodOrderRef: prodOrderRef,
        ginNo: String(row[CONFIG.MOVEMENT_COLS.GIN_NO] || '').trim(),
        qualityStatus: String(row[CONFIG.MOVEMENT_COLS.QUALITY_STATUS] || '').trim(),
        toProductionQty: toProductionQty,
        remarks: String(row[CONFIG.MOVEMENT_COLS.REMARKS] || '').trim()
      });

      totalQty += qty;
      totalToProduction += toProductionQty;
      if (itemCode) itemSet[codeNorm] = true;
      flowCounts[flow] = (flowCounts[flow] || 0) + 1;
    }

    rows.sort(function (a, b) {
      return String(b.timestamp || '').localeCompare(String(a.timestamp || ''));
    });

    return {
      filters: {
        flow: requestedFlow,
        period: windowInfo.period,
        startDate: Utilities.formatDate(start, tz, 'yyyy-MM-dd'),
        endDate: Utilities.formatDate(end, tz, 'yyyy-MM-dd')
      },
      summary: {
        totalRows: rows.length,
        totalQty: totalQty,
        totalToProduction: totalToProduction,
        uniqueItems: Object.keys(itemSet).length,
        inwardCount: flowCounts.INWARD || 0,
        outwardCount: flowCounts.OUTWARD || 0,
        transferCount: flowCounts.TRANSFER || 0
      },
      rows: rows
    };
  });
}

/**
 * BULK UPLOAD BACKEND FUNCTION
 * Add this to Code.js
 * Handles bulk import of items or batches with column mapping
 */

function bulkUploadMasterData(data) {
  return withScriptLock(function () {
    protect(() => requireRole(SECURITY.ROLES.MANAGER));

    try {
      const type = String(data.type || 'items').toLowerCase();
      const rows = data.rows || [];
      const columnMap = data.columnMap || {};

      if (rows.length === 0) {
        return { success: false, message: 'No data to upload' };
      }

      if (Object.keys(columnMap).length === 0) {
        return { success: false, message: 'Column mapping is required' };
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();

      if (type === 'items') {
        return bulkUploadItems(ss, rows, columnMap);
      } else if (type === 'batches') {
        return bulkUploadBatches(ss, rows, columnMap);
      } else {
        return { success: false, message: 'Invalid upload type' };
      }

    } catch (error) {
      Logger.log('[BULK_UPLOAD] Error: ' + error.message);
      return { success: false, message: error.message };
    }
  });
}


/**
 * Bulk upload items to Item_Master sheet
 */
function bulkUploadItems(ss, rows, columnMap) {
  const itemSheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);

  if (!itemSheet) {
    return { success: false, message: 'Item_Master sheet not found' };
  }

  // Get existing items to check duplicates
  const existingData = itemSheet.getDataRange().getValues();
  const existingCodes = {};
  let maxId = 0;

  for (let i = 1; i < existingData.length; i++) {
    const code = String(existingData[i][CONFIG.ITEM_COLS.ITEM_CODE] || '').trim().toUpperCase();
    const id = Number(existingData[i][CONFIG.ITEM_COLS.ITEM_ID]) || 0;

    if (code) existingCodes[code] = true;
    if (id > maxId) maxId = id;
  }

  const newRows = [];
  const errors = [];
  let currentId = maxId;

  // Process each row
  rows.forEach(function (row, index) {
    try {
      const itemData = {};

      // Extract data based on column mapping
      Object.keys(columnMap).forEach(function (colIndex) {
        const fieldKey = columnMap[colIndex];
        itemData[fieldKey] = String(row[colIndex] || '').trim();
      });

      // Validate required fields
      const itemCode = itemData.item_code ? itemData.item_code.trim().toUpperCase() : '';
      const itemName = itemData.item_name || '';
      const category = itemData.category || '';
      const uom = itemData.uom ? itemData.uom.toUpperCase() : '';
      const minStock = Number(itemData.min_stock) || 0;

      if (!itemCode || itemCode.length < 3) {
        errors.push('Row ' + (index + 1) + ': Invalid item code (min 3 chars)');
        return;
      }

      if (!itemName) {
        errors.push('Row ' + (index + 1) + ': Item name is required');
        return;
      }

      if (existingCodes[itemCode]) {
        errors.push('Row ' + (index + 1) + ': Item ' + itemCode + ' already exists in master');
        return;
      }

      // Check duplicate in current batch
      if (newRows.some(r => r[CONFIG.ITEM_COLS.ITEM_CODE] === itemCode)) {
        errors.push('Row ' + (index + 1) + ': Duplicate item code in current upload: ' + itemCode);
        return;
      }

      // Build new row
      currentId++;
      const newRow = [];
      newRow[CONFIG.ITEM_COLS.ITEM_ID] = currentId;
      newRow[CONFIG.ITEM_COLS.ITEM_CODE] = itemCode;
      newRow[CONFIG.ITEM_COLS.ITEM_NAME] = itemName;
      newRow[CONFIG.ITEM_COLS.ITEM_GROUP_CODE] = category;
      newRow[CONFIG.ITEM_COLS.UOM_CODE] = uom;
      newRow[CONFIG.ITEM_COLS.MIN_STOCK_LEVEL] = minStock;
      newRow[6] = 0; // weight
      newRow[7] = 'BULK_UPLOAD'; // remarks
      newRow[CONFIG.ITEM_COLS.STATUS] = 'Active';

      newRows.push(newRow);

    } catch (rowError) {
      errors.push('Row ' + (index + 1) + ': ' + rowError.message);
    }
  });

  // If there are errors, don't upload anything
  if (errors.length > 0) {
    return {
      success: false,
      message: 'Validation failed',
      errors: errors
    };
  }

  // Write all rows at once
  if (newRows.length > 0) {
    const startRow = itemSheet.getLastRow() + 1;
    const range = itemSheet.getRange(startRow, 1, newRows.length, newRows[0].length);
    range.setValues(newRows);

    // Clear cache
    _clearSheetCache(CONFIG.SHEETS.ITEM);

    Logger.log('[BULK_UPLOAD] Created ' + newRows.length + ' items');
  }

  return {
    success: true,
    created: newRows.length,
    message: 'Successfully uploaded ' + newRows.length + ' items'
  };
}


/**
 * Bulk upload batches to Batch_Master sheet
 */
function bulkUploadBatches(ss, rows, columnMap) {
  const batchSheet = ss.getSheetByName(CONFIG.SHEETS.BATCH);

  if (!batchSheet) {
    return { success: false, message: 'Batch_Master sheet not found' };
  }

  // Get existing batches to find max ID for numbering
  const existingData = batchSheet.getDataRange().getValues();
  const existingBatches = {};
  let maxBatchId = 0;

  for (let i = 1; i < existingData.length; i++) {
    const itemCode = String(existingData[i][CONFIG.BATCH_COLS.ITEM_CODE] || '').trim().toUpperCase();
    const batchNum = String(existingData[i][CONFIG.BATCH_COLS.BATCH_NUMBER] || '').trim().toUpperCase();
    const key = itemCode + '|' + batchNum;
    existingBatches[key] = true;

    // Track max numeric ID
    const bId = Number(existingData[i][CONFIG.BATCH_COLS.BATCH_ID]);
    if (isFinite(bId) && bId > maxBatchId) maxBatchId = bId;
  }

  // Get existing items to check if item_code is valid
  const itemSheet = ss.getSheetByName(CONFIG.SHEETS.ITEM);
  const itemMasterData = itemSheet ? itemSheet.getDataRange().getValues() : [];
  const validItemCodes = {};
  for (let j = 1; j < itemMasterData.length; j++) {
    validItemCodes[String(itemMasterData[j][CONFIG.ITEM_COLS.ITEM_CODE] || '').trim().toUpperCase()] = true;
  }

  const newRows = [];
  const errors = [];
  let currentBatchId = maxBatchId;

  // Process each row
  rows.forEach(function (row, index) {
    try {
      const batchData = {};

      // Extract data
      Object.keys(columnMap).forEach(function (colIndex) {
        const fieldKey = columnMap[colIndex];
        batchData[fieldKey] = String(row[colIndex] || '').trim();
      });

      const itemCode = batchData.item_code ? batchData.item_code.toUpperCase() : '';
      const batchNumber = batchData.batch_number || '';
      const expiryDate = batchData.expiry_date || '';

      if (!itemCode) {
        errors.push('Row ' + (index + 1) + ': Item code is required');
        return;
      }

      if (!batchNumber) {
        errors.push('Row ' + (index + 1) + ': Batch number is required');
        return;
      }

      if (!validItemCodes[itemCode]) {
        errors.push('Row ' + (index + 1) + ': Item ' + itemCode + ' not found in Item_Master');
        return;
      }

      const key = itemCode + '|' + batchNumber;
      if (existingBatches[key]) {
        errors.push('Row ' + (index + 1) + ': Batch ' + batchNumber + ' already exists for item ' + itemCode);
        return;
      }

      // Check duplicate in current batch
      if (newRows.some(r => r[CONFIG.BATCH_COLS.ITEM_CODE] === itemCode &&
        r[CONFIG.BATCH_COLS.BATCH_NUMBER] === batchNumber)) {
        errors.push('Row ' + (index + 1) + ': Duplicate batch in current upload');
        return;
      }

      // Parse Date robustly
      let parsedDate = '';
      if (expiryDate) {
        const dt = new Date(expiryDate);
        if (!isNaN(dt.getTime())) {
          parsedDate = dt;
        } else {
          // Try DD-MM-YYYY or DD/MM/YYYY
          const parts = expiryDate.split(/[-/]/);
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            const dt2 = new Date(year, month, day);
            if (!isNaN(dt2.getTime())) parsedDate = dt2;
          }
        }
      }

      // Build new row
      currentBatchId++;
      const newRow = [];
      newRow[CONFIG.BATCH_COLS.BATCH_ID] = currentBatchId;
      newRow[CONFIG.BATCH_COLS.ITEM_CODE] = itemCode;
      newRow[CONFIG.BATCH_COLS.BATCH_NUMBER] = batchNumber;
      newRow[CONFIG.BATCH_COLS.BATCH_EXP_DATE] = parsedDate;
      newRow[CONFIG.BATCH_COLS.LOT_NUMBER] = '';
      newRow[CONFIG.BATCH_COLS.GIN_NO] = 'BULK-' + new Date().getTime();
      newRow[CONFIG.BATCH_COLS.QUALITY_DATE] = new Date();
      newRow[CONFIG.BATCH_COLS.QUALITY_STATUS] = 'PENDING';
      newRow[CONFIG.BATCH_COLS.VERSION] = 'V1';
      newRow[CONFIG.BATCH_COLS.VERSION_PARENT_ID] = '';
      newRow[CONFIG.BATCH_COLS.IS_VERSION] = false;

      newRows.push(newRow);

    } catch (rowError) {
      errors.push('Row ' + (index + 1) + ': ' + rowError.message);
    }
  });

  if (errors.length > 0) {
    return {
      success: false,
      message: 'Validation failed',
      errors: errors
    };
  }

  // Write all rows
  if (newRows.length > 0) {
    const startRow = batchSheet.getLastRow() + 1;
    const range = batchSheet.getRange(startRow, 1, newRows.length, newRows[0].length);
    range.setValues(newRows);

    _clearSheetCache(CONFIG.SHEETS.BATCH);

    Logger.log('[BULK_UPLOAD] Created ' + newRows.length + ' batches');
  }

  return {
    success: true,
    created: newRows.length,
    message: 'Successfully uploaded ' + newRows.length + ' batches'
  };
}

/**  Run checkBulkDupes() in editor to see how many batches are already in Batch_Master */

function checkBulkDupes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bm = ss.getSheetByName('Batch_Master').getDataRange().getValues();
  const existing = new Set(bm.slice(1).map(r => r[1] + '|' + r[2]));  // itemCode|batchNumber
  const tpl = ss.getSheetByName('Bulk_Upload_Template').getDataRange().getValues();
  const dupes = tpl.slice(1).filter(r => existing.has(r[0] + '|' + r[1]));
  Logger.log('Dupes: ' + dupes.length + ' / ' + (tpl.length - 1));
  return dupes.map(r => r[0] + '|' + r[1]);
}

/**
 * BULK UPLOAD INVENTORY FROM SHEET
 * 
 * Paste this function at the bottom of Code.js (before the last closing brace)
 * 
 * Usage:
 * 1. Create sheet "Bulk_Upload_Template" with headers
 * 2. Fill data rows
 * 3. Run this function (or click button)
 * 4. Validates all rows, creates inventory entries
 */

function processBulkUploadFast() {
  return withScriptLock(function () {
    protect(() => requireRole(SECURITY.ROLES.MANAGER));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Bulk_Upload_Template');
    if (!templateSheet) throw new Error('Bulk_Upload_Template sheet not found');

    const data = templateSheet.getDataRange().getValues();
    if (data.length < 2) throw new Error('No data in template');

    // ── Build lookup maps ONCE (O(n) not O(n²)) ──────────────────────────
    const itemSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ITEM);
    const itemData = _getSheetValuesCached(itemSheet.getName());
    const itemCols = _getItemSheetColumnMap(itemSheet);
    const itemIdByCode = {};   // 'PM1779' → '15073'
    for (let i = 1; i < itemData.length; i++) {
      const code = String(itemData[i][itemCols.ITEM_CODE] || '').trim().toUpperCase();
      const id = String(itemData[i][itemCols.ITEM_ID] || '').trim();
      if (code && id) itemIdByCode[code] = id;
    }

    const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
    const binData = _getSheetValuesCached(binSheet.getName());
    const binHeaders = _getSheetHeaderMap(binSheet);
    const binMetaRows = _buildBinMasterMetaMap(binSheet, binData, binHeaders);
    const binMetaMap = {};
    Object.keys(binMetaRows).forEach(function (bid) {
      const meta = binMetaRows[bid] || {};
      binMetaMap[bid] = {
        site: String(meta.site || '').trim(),
        location: String(meta.location || '').trim(),
        capacity: Number(meta.capacity || 0),
        capacityUom: _normalizeCapacityUom(meta.declaredCapacityUom || meta.capacityUom || 'KG')
      };
    });

    // ── Parse header ─────────────────────────────────────────────────────
    const hdr = {};
    data[0].forEach((h, i) => { hdr[String(h || '').trim().toLowerCase().replace(/ /g, '_')] = i; });
    const required = ['item_code', 'batch_id', 'bin_id', 'quantity', 'uom'];
    const missing = required.filter(c => hdr[c] === undefined);
    if (missing.length) throw new Error('Missing columns: ' + missing.join(', '));

    const errors = [], validRows = [];
    const now = new Date();
    const binDelta = {};   // track capacity delta within this batch

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const itemCode = String(row[hdr['item_code']] || '').trim().toUpperCase();
      if (!itemCode) continue;   // skip blank rows

      const rn = i + 1;
      const batchId = String(row[hdr['batch_id']] || '').trim();
      const binId = String(row[hdr['bin_id']] || '').trim();
      const qty = Number(row[hdr['quantity']]) || 0;
      const uom = String(row[hdr['uom']] || 'NOS').trim().toUpperCase();
      const qa = String(row[hdr['quality_status']] || 'PENDING').trim().toUpperCase();
      const ginNo = String(row[hdr['gin_no']] || 'BULK-' + now.getTime()).trim();
      const version = String(row[hdr['version']] || 'V1').trim().toUpperCase();

      if (!batchId) { errors.push('Row ' + rn + ': Batch ID required'); continue; }
      if (!binId) { errors.push('Row ' + rn + ': Bin ID required'); continue; }
      if (qty <= 0) { continue; }   // skip zero-qty rows silently (nothing to store)

      // O(1) lookup
      const itemId = itemIdByCode[itemCode];
      if (!itemId) { errors.push('Row ' + rn + ': Item ' + itemCode + ' not in Item_Master'); continue; }

      const binInfo = binMetaMap[binId];
      if (!binInfo) { errors.push('Row ' + rn + ': Bin ' + binId + ' not in Bin_Master'); continue; }

      const validQA = { PENDING: 1, APPROVED: 1, HOLD: 1, REJECTED: 1 };
      if (!validQA[qa]) { errors.push('Row ' + rn + ': Invalid QA: ' + qa); continue; }

      // Capacity check (only if bin has a real capacity limit)
      const qtyKg = _convertToKg(itemCode, qty, uom);
      try {
        _assertBinCapacity(binId, qty, binDelta, itemCode, uom);
      } catch (e) {
        errors.push('Row ' + rn + ': ' + e.message);
        continue;
      }

      validRows.push({
        itemId, itemCode, batchId, binId,
        qty: qtyKg, site: binInfo.site, location: binInfo.location,
        uom, qa, ginNo, version,
        lotNo: String(row[hdr['lot_no']] || '').trim(),
        expiryDate: row[hdr['expiry_date']] instanceof Date ? row[hdr['expiry_date']] : '',
        remarks: String(row[hdr['remarks']] || '').trim()
      });
    }

    if (errors.length > 0 && validRows.length === 0)
      throw new Error(errors.length + ' validation errors:\n' + errors.slice(0, 20).join('\n'));

    // ── Write inventory rows ──────────────────────────────────────────────
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const moveLogs = [];

    validRows.forEach(function (item) {
      const invId = 'INV-' + now.getTime() + '-' + Math.floor(Math.random() * 9999);
      const newRow = [];
      newRow[CONFIG.INVENTORY_COLS.INVENTORY_ID] = invId;
      newRow[CONFIG.INVENTORY_COLS.ITEM_ID] = item.itemId;
      newRow[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] = item.itemCode;
      newRow[CONFIG.INVENTORY_COLS.BATCH_ID] = item.batchId;
      newRow[CONFIG.INVENTORY_COLS.GIN_NO] = item.ginNo;
      newRow[CONFIG.INVENTORY_COLS.VERSION] = item.version;
      newRow[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = item.qa;
      newRow[CONFIG.INVENTORY_COLS.QUALITY_DATE] = item.qa === 'APPROVED' ? now : '';
      newRow[CONFIG.INVENTORY_COLS.BIN_ID] = item.binId;
      newRow[CONFIG.INVENTORY_COLS.SITE] = item.site;
      newRow[CONFIG.INVENTORY_COLS.LOCATION] = item.location;
      newRow[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = item.qty;
      newRow[CONFIG.INVENTORY_COLS.UOM] = item.uom;
      newRow[CONFIG.INVENTORY_COLS.LOT_NO] = item.lotNo;
      newRow[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = item.expiryDate;
      newRow[CONFIG.INVENTORY_COLS.LAST_UPDATED] = now;
      newRow[CONFIG.INVENTORY_COLS.INWARD_DATE] = now;
      newRow[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = '';
      invSheet.appendRow(newRow);

      moveLogs.push({
        type: CONFIG.MOVEMENT_TYPES.INWARD,
        itemId: item.itemId, batchId: item.batchId,
        version: item.version, ginNo: item.ginNo,
        toBinId: item.binId, quantity: item.qty,
        qualityStatus: item.qa, remarks: 'BULK_UPLOAD: ' + (item.remarks || '')
      });
    });

    if (moveLogs.length > 0) {
      try { _appendMovementLogsBatch(moveLogs); } catch (e) {
        Logger.log('[BULK] Movement log failed: ' + e.message);
      }
    }

    // Clear template after success
    if (validRows.length > 0 && data.length > 1)
      templateSheet.getRange(2, 1, data.length - 1, data[0].length).clearContent();

    Logger.log('[BULK_UPLOAD] Processed: ' + validRows.length
      + (errors.length ? ' | Skipped: ' + errors.length : ''));

    return {
      success: true, processed: validRows.length,
      skipped: errors.length, errors: errors.slice(0, 50)
    };
  });
}

function processBulkUpload() {
  return withScriptLock(function () {
    protect(() => requireRole(SECURITY.ROLES.MANAGER));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Bulk_Upload_Template');

    if (!templateSheet) {
      throw new Error('Bulk_Upload_Template sheet not found');
    }

    const data = templateSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('No data in Bulk_Upload_Template');
    }

    const header = data[0];
    const headerMap = {};
    header.forEach((h, i) => {
      const key = String(h || '').trim().toLowerCase().replace(/ /g, '_');
      headerMap[key] = i;
    });

    const requiredCols = ['item_code', 'batch_id', 'bin_id', 'quantity', 'uom'];
    const missingCols = requiredCols.filter(col => headerMap[col] === undefined);

    if (missingCols.length > 0) {
      throw new Error('Missing columns: ' + missingCols.join(', '));
    }

    const errors = [];
    const validRows = [];
    const now = new Date();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1;

      const itemCode = String(row[headerMap['item_code']] || '').trim();
      if (!itemCode) continue;

      const batchId = String(row[headerMap['batch_id']] || '').trim();
      const binId = String(row[headerMap['bin_id']] || '').trim();
      const qty = Number(row[headerMap['quantity']]) || 0;
      const uom = String(row[headerMap['uom']] || 'KG').trim().toUpperCase();
      const qa = String(row[headerMap['quality_status']] || 'PENDING').trim().toUpperCase();
      const lotNo = String(row[headerMap['lot_no']] || '').trim();
      const expiryDate = row[headerMap['expiry_date']];
      const ginNo = String(row[headerMap['gin_no']] || 'BULK-' + now.getTime()).trim();
      const version = String(row[headerMap['version']] || 'V1').trim().toUpperCase();
      const remarks = String(row[headerMap['remarks']] || '').trim();

      if (!batchId) {
        errors.push('Row ' + rowNum + ': Batch ID required');
        continue;
      }
      if (!binId) {
        errors.push('Row ' + rowNum + ': Bin ID required');
        continue;
      }
      if (qty <= 0) {
        errors.push('Row ' + rowNum + ': Quantity must be > 0');
        continue;
      }

      let itemId;
      try {
        itemId = _getValidatedItemId(itemCode);
      } catch (e) {
        errors.push('Row ' + rowNum + ': Item ' + itemCode + ' not in Item_Master');
        continue;
      }

      try {
        _validateBinExists(binId);
      } catch (e) {
        errors.push('Row ' + rowNum + ': Bin ' + binId + ' not in Bin_Master');
        continue;
      }

      const validQA = ['PENDING', 'APPROVED', 'HOLD', 'REJECTED'];
      if (!validQA.includes(qa)) {
        errors.push('Row ' + rowNum + ': Invalid QA status: ' + qa);
        continue;
      }

      const binSheet = ss.getSheetByName(CONFIG.SHEETS.BIN);
      const binData = _getSheetValuesCached(binSheet.getName());

      let site = '';
      let location = '';

      for (let j = 1; j < binData.length; j++) {
        if (String(binData[j][0] || '').trim() === binId) {
          site = String(binData[j][3] || '').trim();
          location = String(binData[j][4] || '').trim();
          break;
        }
      }

      const qtyKg = _convertToKg(itemCode, qty, uom);

      try {
        _assertBinCapacity(binId, qty, null, itemCode, uom);
      } catch (e) {
        errors.push('Row ' + rowNum + ': ' + e.message);
        continue;
      }

      validRows.push({
        itemId: itemId,
        itemCode: itemCode,
        batchId: batchId,
        binId: binId,

        qty: qtyKg,
        site: site,
        location: location,
        uom: uom,
        qa: qa,
        lotNo: lotNo,
        expiryDate: expiryDate && expiryDate instanceof Date ? expiryDate : '',
        ginNo: ginNo,
        version: version,
        remarks: remarks
      });
    }

    if (errors.length > 0) {
      Logger.log('[BULK_UPLOAD] Validation errors: ' + errors.length);
      return {
        success: false,
        processed: 0,
        errors: errors,
        message: 'Fix errors and retry'
      };
    }

    if (validRows.length === 0) {
      return {
        success: false,
        processed: 0,
        errors: ['No valid rows'],
        message: 'No data to process'
      };
    }

    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const movementLogs = [];
    let created = 0;

    validRows.forEach(function (item) {
      const invData = _getSheetValuesCached(invSheet.getName());
      let existingRow = null;

      for (let i = 1; i < invData.length; i++) {
        const row = invData[i];
        if (String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim() === item.itemCode &&
          String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim() === item.batchId &&
          String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim() === item.binId &&
          String(row[CONFIG.INVENTORY_COLS.VERSION] || '').trim() === item.version &&
          String(row[CONFIG.INVENTORY_COLS.GIN_NO] || '').trim() === item.ginNo) {
          existingRow = {
            index: i + 1,
            currentQty: Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0
          };
          break;
        }
      }

      if (existingRow) {
        const newQty = existingRow.currentQty + item.qty;
        invSheet.getRange(existingRow.index, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1).setValue(newQty);
        invSheet.getRange(existingRow.index, CONFIG.INVENTORY_COLS.LAST_UPDATED + 1).setValue(now);
        _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, existingRow.index, CONFIG.INVENTORY_COLS.TOTAL_QUANTITY + 1, newQty);
      } else {
        const invId = 'INV-' + now.getTime() + '-' + Math.floor(Math.random() * 1000);
        const newRow = [];
        newRow[CONFIG.INVENTORY_COLS.INVENTORY_ID] = invId;
        newRow[CONFIG.INVENTORY_COLS.ITEM_ID] = item.itemId;
        newRow[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] = item.itemCode;
        newRow[CONFIG.INVENTORY_COLS.BATCH_ID] = item.batchId;
        newRow[CONFIG.INVENTORY_COLS.GIN_NO] = item.ginNo;
        newRow[CONFIG.INVENTORY_COLS.VERSION] = item.version;
        newRow[CONFIG.INVENTORY_COLS.QUALITY_STATUS] = item.qa;
        newRow[CONFIG.INVENTORY_COLS.QUALITY_DATE] = now;
        newRow[CONFIG.INVENTORY_COLS.BIN_ID] = item.binId;
        newRow[CONFIG.INVENTORY_COLS.SITE] = item.site;
        newRow[CONFIG.INVENTORY_COLS.LOCATION] = item.location;
        newRow[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY] = item.qty;
        newRow[CONFIG.INVENTORY_COLS.UOM] = item.uom;
        newRow[CONFIG.INVENTORY_COLS.LOT_NO] = item.lotNo;
        newRow[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = item.expiryDate;
        newRow[CONFIG.INVENTORY_COLS.LAST_UPDATED] = now;
        newRow[CONFIG.INVENTORY_COLS.INWARD_DATE] = now;
        newRow[CONFIG.INVENTORY_COLS.LAST_TRANSFER_DATE] = '';

        invSheet.appendRow(newRow);
        _appendSheetCacheRow(CONFIG.SHEETS.INVENTORY, newRow);
      }

      movementLogs.push({
        type: 'INWARD_RECEIPT',
        itemId: item.itemId,
        batchId: item.batchId,
        version: item.version,
        ginNo: item.ginNo,
        fromBinId: '',
        toBinId: item.binId,
        quantity: item.qty,
        remarks: 'BULK_UPLOAD: ' + (item.remarks || 'Initial stock')
      });

      created++;
    });

    _appendMovementLogsBatch(movementLogs);

    if (data.length > 1) {
      templateSheet.getRange(2, 1, data.length - 1, header.length).clearContent();
    }

    Logger.log('[BULK_UPLOAD] Processed: ' + created + ' items');

    return {
      success: true,
      processed: created,
      errors: [],
      message: 'Loaded ' + created + ' inventory entries'
    };
  });
}


// Paste this into Apps Script editor and run it once.
// It re-runs ONLY the validation (no writes) and logs every failed row.
function diagnoseFailedBulkRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName('Bulk_Upload_Template');
  if (!templateSheet) { Logger.log('No template sheet'); return; }

  var data = templateSheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('No data'); return; }

  // Build item lookup
  var itemSheet = ss.getSheetByName('Item_Master');
  var itemData = itemSheet.getDataRange().getValues();
  var itemIdByCode = {};
  for (var m = 1; m < itemData.length; m++) {
    var c = String(itemData[m][1] || '').trim().toUpperCase(); // col B = item_code
    var d = String(itemData[m][0] || '').trim();               // col A = item_id
    if (c && d) itemIdByCode[c] = d;
  }

  // Build bin lookup
  var binSheet = ss.getSheetByName('Bin_Master');
  var binData = binSheet.getDataRange().getValues();
  var binSet = {};
  for (var b = 1; b < binData.length; b++) {
    var bid = String(binData[b][0] || '').trim();
    if (bid) binSet[bid] = true;
  }

  // Parse header
  var hdr = {};
  data[0].forEach(function (h, i) {
    hdr[String(h || '').trim().toLowerCase().replace(/ /g, '_')] = i;
  });

  var VALID_QA = { PENDING: 1, APPROVED: 1, HOLD: 1, REJECTED: 1 };
  var errors = [];
  var passed = 0;
  var skipped = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rn = i + 1;
    var itemCode = String(row[hdr['item_code']] || '').trim().toUpperCase();
    if (!itemCode) { skipped++; continue; }

    var batchId = String(row[hdr['batch_id']] || '').trim();
    var binId = String(row[hdr['bin_id']] || '').trim();
    var qty = Number(row[hdr['quantity']]) || 0;
    var uom = String(row[hdr['uom']] || '').trim().toUpperCase();
    var qa = String(row[hdr['quality_status']] || 'PENDING').trim().toUpperCase();

    var reason = '';
    if (!batchId) reason = 'Batch ID missing';
    else if (!binId) reason = 'Bin ID missing';
    else if (qty <= 0) reason = 'Qty = ' + qty + ' (zero/negative)';
    else if (!itemIdByCode[itemCode]) reason = 'Item ' + itemCode + ' not in Item_Master';
    else if (!VALID_QA[qa]) reason = 'Invalid QA: ' + qa;
    else if (binId !== 'NULL' && !binSet[binId]) reason = 'Bin "' + binId + '" not in Bin_Master';

    if (reason) {
      errors.push('Row ' + rn + ' | ' + itemCode + ' | Batch: ' + batchId + ' | ' + reason);
    } else {
      passed++;
    }
  }

  Logger.log('PASSED: ' + passed + ' | FAILED: ' + errors.length + ' | BLANK SKIPPED: ' + skipped);
  Logger.log('--- FAILED ROWS ---');
  errors.forEach(function (e) { Logger.log(e); });

  // Also write failures to a new sheet so you can see them in the spreadsheet
  var failSheet = ss.getSheetByName('Bulk_Upload_Failures');
  if (!failSheet) failSheet = ss.insertSheet('Bulk_Upload_Failures');
  failSheet.clearContents();
  failSheet.getRange(1, 1).setValue('Row | ItemCode | Batch | Reason');
  if (errors.length > 0) {
    var rows = errors.map(function (e) { return [e]; });
    failSheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }
  Logger.log('Results written to "Bulk_Upload_Failures" sheet');
}
