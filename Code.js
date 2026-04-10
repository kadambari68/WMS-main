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
    AI_QUERY_LOG: 'AI_Query_Log'
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

  MOVEMENT_TYPES: {
    INWARD: 'INWARD_RECEIPT',
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
    codeToMinStock: {},
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
}

function _clearSheetCache(sheetName) {
  _ensureRequestCache();
  const name = String(sheetName || '').trim();
  if (!name) return;
  delete _REQ_CACHE.values[name];
  delete _REQ_CACHE.headers[name];
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

function _getItemCodeById(itemId) {
  if (!itemId) return null;
  const targetId = String(itemId || '').trim().toUpperCase();
  const maps = _getItemMasterMaps();
  return maps.idToCode[targetId] || null;
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
    return locs;
  });
}

/**
 * Get all bins with hierarchy (Site/Location from Rack)
 */
function getBins() {
  return _withRequestCache(function () {
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
          site: String(rackData[i][2]), // Site (Col C)
          location: String(rackData[i][3]) // Location (Col D)
        };
      }
    }

    const data = _getSheetValuesCached(binSheet.getName());
    const headerMap = _getSheetHeaderMap(binSheet);
    const capIdx = (typeof headerMap['MAX_CAPACITY_KG'] === 'number')
      ? headerMap['MAX_CAPACITY_KG']
      : ((typeof headerMap['MAX_CAPACITY'] === 'number') ? headerMap['MAX_CAPACITY'] : 3);
    const capUomIdx = (typeof headerMap['CAPACITY_UOM'] === 'number')
      ? headerMap['CAPACITY_UOM']
      : ((typeof headerMap['MAX_CAPACITY_UOM'] === 'number') ? headerMap['MAX_CAPACITY_UOM'] : null);
    const qtyByBin = _getInventoryQtyByBinMap();
    const bins = [];

    for (let i = 1; i < data.length; i++) {
      const binId = String(data[i][0] || '').trim();
      if (!binId) continue;
      const rackRef = String(data[i][1]).trim(); // Rack Ref (Col B)
      const binCode = data[i][2]; // Bin Code (Col C)
      const maxCapacity = Number(data[i][capIdx]) || 0;
      const currentUsage = Number(qtyByBin[binId] || 0);
      const availableCapacity = (maxCapacity > 0) ? Math.max(0, maxCapacity - currentUsage) : 0;
      const capacityUom = (capUomIdx === null)
        ? 'KG'
        : (String(data[i][capUomIdx] || 'KG').trim().toUpperCase() || 'KG');

      // Resolve Site/Location from Rack Map or Defaults
      const meta = rackMap[rackRef] || { site: 'Main', location: 'Main' };

      bins.push({
        binId: binId,
        binCode: binCode,
        site: meta.site,
        location: meta.location,
        maxCapacity: maxCapacity,
        currentUsage: currentUsage,
        availableCapacity: availableCapacity,
        capacityUom: capacityUom
      });
    }
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
  const id = String(batchId || '').trim();
  if (!id) return null;
  const code = String(itemCode || '').trim().toUpperCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.BATCH);
  if (!sheet) return null;
  const data = _getSheetValuesCached(sheet.getName());
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][CONFIG.BATCH_COLS.BATCH_ID] || '').trim();
    if (rowId !== id) continue;
    if (code && String(data[i][CONFIG.BATCH_COLS.ITEM_CODE] || '').trim().toUpperCase() !== code) continue;
    return _buildBatchMetaFromRow(data[i]);
  }
  return null;
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

function getRacks() {
  return _withRequestCache(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
    const binSheet = ss.getSheetByName(CONFIG.SHEETS.BIN);
    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);

    if (!rackSheet || !binSheet) return [];

    // Load all data
    const rackData = _getSheetValuesCached(rackSheet.getName());
    const binData = _getSheetValuesCached(binSheet.getName());
    const invData = invSheet ? _getSheetValuesCached(invSheet.getName()) : [];

    // Build bin map: bin_id -> { max_capacity, items[] }
    const binMap = {};
    for (let i = 1; i < binData.length; i++) {
      const binId = String(binData[i][0] || '').trim();
      if (!binId) continue;
      const binCode = binData[i][2] || '';
      const maxCapacity = Number(binData[i][3]) || 0;
      binMap[binId] = {
        binId: binId,
        code: binCode,  // Frontend expects 'code' not 'binCode'
        binCode: binCode,
        maxCapacity: maxCapacity,
        currentUsage: 0,
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

    // Populate inventory items into bins
    if (invData.length > 0) {
      for (let i = 1; i < invData.length; i++) {
        const row = invData[i];
        const binId = String(row[CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
        const itemCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
        const qty = Number(row[CONFIG.INVENTORY_COLS.TOTAL_QUANTITY]) || 0;
        const qaStatus = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').trim();

        if (binId && binMap[binId] && itemCode && qty > 0) {
          binMap[binId].currentUsage += qty;
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
          binObj.available = binObj.maxCapacity - binObj.currentUsage;
          rackBins.push(binObj);
        }
      }

      racks.push({
        rackCode: rackCode,
        rackId: rackId,
        site: site,
        location: location,
        maxCapacity: maxCapacity,
        currentUsage: rackBins.reduce((sum, b) => sum + b.currentUsage, 0),
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
  if (lock.tryLock(10000)) {
    try {
      return callback();
    } catch (e) {
      Logger.log('Transaction Error: ' + e.toString());
      throw e;
    } finally {
      lock.releaseLock();
      _endRequestCache();
    }
  } else {
    _endRequestCache();
    throw new Error('System is busy. Please try again.');
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
    const batchMatch = searchBatch ? (rowBatch === searchBatch) : true;

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
  const requestedU = String(requested).toUpperCase();
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

    if (requestedU && rowBatchU === requestedU) matchedRequested = true;
    if (!batchMap[rowBatchU]) batchMap[rowBatchU] = rowBatchRaw;
  }

  if (requestedU && matchedRequested) return requested;
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
    row[CONFIG.INVENTORY_COLS.UOM] = _getItemUomCode(params.itemCode || params.itemId || '');
  }
  if (CONFIG.INVENTORY_COLS.LOT_NO !== undefined) {
    row[CONFIG.INVENTORY_COLS.LOT_NO] = params.lotNo || '';
  }
  if (CONFIG.INVENTORY_COLS.EXPIRY_DATE !== undefined) {
    row[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = params.expiryDate || '';
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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ITEM);
  const data = _getSheetValuesCached(sheet.getName());

  // Search Item_Master for matching Item_Code
  for (let i = 1; i < data.length; i++) { // Skip header
    const rowItemCode = String(data[i][1]); // Assuming Item_Code is column B (index 1)
    if (rowItemCode.toUpperCase() === String(itemCode).toUpperCase()) {
      const itemId = String(data[i][0]); // Item_ID is column A (index 0)
      if (!itemId) throw new Error(`Item ${itemCode} has no ID in master`);
      return itemId;
    }
  }

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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
  const data = _getSheetValuesCached(sheet.getName());

  // Search Bin_Master for matching Bin_ID
  for (let i = 1; i < data.length; i++) { // Skip header
    const rowBinId = String(data[i][0]); // Assuming Bin_ID is column A
    if (rowBinId === String(binId)) {
      return true;
    }
  }

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

function _getBinMasterMeta(binId, binSheet, binData, colMap) {
  const idIdx = typeof colMap['BIN_ID'] === 'number' ? colMap['BIN_ID'] : 0;
  const rackIdx = typeof colMap['RACK_ID'] === 'number' ? colMap['RACK_ID'] : 1;
  const capIdx = typeof colMap['MAX_CAPACITY_KG'] === 'number'
    ? colMap['MAX_CAPACITY_KG']
    : (typeof colMap['MAX_CAPACITY'] === 'number' ? colMap['MAX_CAPACITY'] : 3);
  const statusIdx = typeof colMap['BIN_STATUS'] === 'number' ? colMap['BIN_STATUS'] : null;

  for (let i = 1; i < binData.length; i++) {
    if (String(binData[i][idIdx] || '').trim() === String(binId || '').trim()) {
      return {
        rowIndex: i + 1,
        rackId: String(binData[i][rackIdx] || '').trim(),
        capacityKg: Number(binData[i][capIdx]) || 0,
        statusIdx: statusIdx
      };
    }
  }
  return null;
}

function _assertBinCapacity(binId, addQtyKg, binDeltaMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
  const binData = _getSheetValuesCached(binSheet.getName());
  const colMap = _getSheetHeaderMap(binSheet);
  const meta = _getBinMasterMeta(binId, binSheet, binData, colMap);
  if (!meta) return;

  if (!meta.capacityKg || meta.capacityKg <= 0) {
    Logger.log(`Bin capacity missing for ${binId}; skipping capacity check`);
    return;
  }

  const currentMap = _getInventoryQtyByBinMap();
  const currentQty = Number(currentMap[String(binId || '').trim()] || 0);
  const delta = (binDeltaMap && binDeltaMap[String(binId || '').trim()]) || 0;
  const projected = currentQty + delta + (Number(addQtyKg) || 0);
  if (projected > meta.capacityKg) {
    throw new Error(`Bin capacity exceeded for ${binId}. Max ${meta.capacityKg} KG, projected ${projected} KG.`);
  }
  _assertRackSwlForBin(binId, addQtyKg, binDeltaMap);
}

function _deriveBinStatus(qty, capacityKg) {
  const q = Number(qty) || 0;
  const cap = Number(capacityKg) || 0;
  if (cap > 0 && q > cap) return 'BLOCKED';
  if (q <= 0) return 'FREE';
  if (cap > 0 && q >= cap) return 'FULL';
  return 'PARTIAL';
}

function _updateBinAndRackStatuses(binIds) {
  if (!binIds || binIds.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const binSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.BIN);
  const binData = _getSheetValuesCached(binSheet.getName());
  const binColMap = _getSheetHeaderMap(binSheet);
  const idIdx = typeof binColMap['BIN_ID'] === 'number' ? binColMap['BIN_ID'] : 0;
  const rackIdx = typeof binColMap['RACK_ID'] === 'number' ? binColMap['RACK_ID'] : 1;
  const capIdx = typeof binColMap['MAX_CAPACITY_KG'] === 'number'
    ? binColMap['MAX_CAPACITY_KG']
    : (typeof binColMap['MAX_CAPACITY'] === 'number' ? binColMap['MAX_CAPACITY'] : 3);
  const statusIdx = typeof binColMap['BIN_STATUS'] === 'number' ? binColMap['BIN_STATUS'] : null;

  const qtyMap = _getInventoryQtyByBinMap();
  const targetSet = {};
  (binIds || []).forEach(b => { if (b) targetSet[String(b).trim()] = true; });

  const racksTouched = {};
  if (statusIdx === null) {
    Logger.log('Bin status column not found; skipping bin status updates');
  }

  for (let i = 1; i < binData.length; i++) {
    const binId = String(binData[i][idIdx] || '').trim();
    if (!binId || !targetSet[binId]) continue;
    const rackId = String(binData[i][rackIdx] || '').trim();
    if (rackId) racksTouched[rackId] = true;
    const capacity = Number(binData[i][capIdx]) || 0;
    const qty = Number(qtyMap[binId] || 0);
    const status = _deriveBinStatus(qty, capacity);
    if (statusIdx !== null) {
      binSheet.getRange(i + 1, statusIdx + 1).setValue(status);
      _updateSheetCacheCell(CONFIG.SHEETS.BIN, i + 1, statusIdx + 1, status);
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

  const rackIds = Object.keys(racksTouched);
  if (rackIds.length === 0) return;

  rackIds.forEach(rackId => {
    let rackLoad = 0;
    for (let i = 1; i < binData.length; i++) {
      if (String(binData[i][rackIdx] || '').trim() !== rackId) continue;
      const binId = String(binData[i][idIdx] || '').trim();
      rackLoad += Number(qtyMap[binId] || 0);
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

function _assertRackSwlForBin(binId, addQtyKg, binDeltaMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const binSheet = ss.getSheetByName(CONFIG.SHEETS.BIN);
  const rackSheet = ss.getSheetByName(CONFIG.SHEETS.RACK);
  if (!binSheet || !rackSheet) return;
  const binData = _getSheetValuesCached(binSheet.getName());
  const rackData = _getSheetValuesCached(rackSheet.getName());
  const binColMap = _getSheetHeaderMap(binSheet);
  const rackColMap = _getSheetHeaderMap(rackSheet);
  const idIdx = typeof binColMap['BIN_ID'] === 'number' ? binColMap['BIN_ID'] : 0;
  const rackIdx = typeof binColMap['RACK_ID'] === 'number' ? binColMap['RACK_ID'] : 1;
  const rackIdIdx = typeof rackColMap['RACK_ID'] === 'number' ? rackColMap['RACK_ID'] : 0;
  const rackSwlIdx = typeof rackColMap['RACK_SWL_KG'] === 'number'
    ? rackColMap['RACK_SWL_KG']
    : (typeof rackColMap['RACK_SWL'] === 'number' ? rackColMap['RACK_SWL'] : null);
  if (rackSwlIdx === null) return;

  let rackId = '';
  for (let i = 1; i < binData.length; i++) {
    if (String(binData[i][idIdx] || '').trim() === String(binId || '').trim()) {
      rackId = String(binData[i][rackIdx] || '').trim();
      break;
    }
  }
  if (!rackId) return;

  let rackSwl = 0;
  for (let i = 1; i < rackData.length; i++) {
    if (String(rackData[i][rackIdIdx] || '').trim() === rackId) {
      rackSwl = Number(rackData[i][rackSwlIdx]) || 0;
      break;
    }
  }
  if (!rackSwl || rackSwl <= 0) return;

  const qtyMap = _getInventoryQtyByBinMap();
  const deltaMap = binDeltaMap || {};
  const add = Number(addQtyKg) || 0;
  let rackLoad = 0;
  for (let i = 1; i < binData.length; i++) {
    if (String(binData[i][rackIdx] || '').trim() !== rackId) continue;
    const id = String(binData[i][idIdx] || '').trim();
    const base = Number(qtyMap[id] || 0);
    const delta = Number(deltaMap[id] || 0);
    rackLoad += base + delta;
  }
  if (add) {
    rackLoad += add;
  }
  if (rackLoad > rackSwl) {
    throw new Error(`Rack SWL exceeded for rack ${rackId}. SWL ${rackSwl} KG, projected ${rackLoad} KG.`);
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
    const itemCode = String(data[i][cols.ITEM_CODE_CACHE] || '').trim();
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
    items.forEach(item => {
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
    const code = String(data[i][CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
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

/**
 * Configures min-stock alert recipients from Access_Control_List managers
 * and ensures the daily trigger exists.
 *
 * @returns {{success:boolean, recipients:number, emails:string[]}}
 */
function setupMinStockAlertsForManagers() {
  return protect(function () {
    requireRole(SECURITY.ROLES.MANAGER);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aclSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.ACCESS_CONTROL);
    const data = _getSheetValuesCached(aclSheet.getName());
    const emails = [];

    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][0] || '').trim();
      const role = String(data[i][1] || '').trim().toUpperCase();
      const status = String(data[i][3] || '').trim().toUpperCase();
      if (!email) continue;
      if (status && status !== 'ACTIVE') continue;
      if (role !== SECURITY.ROLES.MANAGER) continue;
      emails.push(email);
    }

    const unique = Array.from(new Set(emails));
    if (unique.length === 0) {
      throw new Error('No active MANAGER emails found in Access_Control_List.');
    }

    PropertiesService.getScriptProperties().setProperty('ALERT_EMAILS', unique.join(','));
    ensureMinStockAlertTrigger();

    return { success: true, recipients: unique.length, emails: unique };
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
      batch: r.batchId,
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
    // Clear the AI knowledge cache if it exists

    sc.remove(AI_CONFIG.CACHE_KEY);
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
    for (let i = 1; i < data.length; i++) {
      const rowInvId = String(data[i][CONFIG.INVENTORY_COLS.INVENTORY_ID]);

      if (rowInvId === String(payload.inventoryId)) {
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
        if (newStatus === 'APPROVED') {
          // Validate capacity/SWL before approval (no quantity change, but prevents approving invalid storage).
          const binId = String(data[i][CONFIG.INVENTORY_COLS.BIN_ID] || '').trim();
          if (binId) {
            _assertBinCapacity(binId, 0);
            _assertRackSwlForBin(binId);
          }
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
    Logger.log('[QA_BACKEND] Checking QA access...');
    assertQualityManagerAccess();

    // Load overrides
    Logger.log('[QA_BACKEND] Loading override events...');
    const overrideMap = _getLatestOverrideEventsByInventoryId() || {};
    Logger.log('[QA_BACKEND] Override map size: ' + Object.keys(overrideMap).length);

    // Load inventory read view
    Logger.log('[QA_BACKEND] Loading inventory read view...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const data = _getSheetValuesCached(invSheet.getName());
    Logger.log('[QA_BACKEND] Inventory sheet rows: ' + data.length);

    // Build visibility rows inline (no nested protect calls)
    const norm = v => String(v || '').trim().toUpperCase();
    const itemIdToCode = {};
    const itemCodeToUom = {};
    const visRows = [];

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
      const rowUomRaw = String((CONFIG.INVENTORY_COLS.UOM !== undefined ? row[CONFIG.INVENTORY_COLS.UOM] : '') || '').trim();
      let rowUom = rowUomRaw;
      if (!rowUom) {
        if (!itemCodeToUom[rowCode]) itemCodeToUom[rowCode] = _getItemUomCode(rowCode);
        rowUom = itemCodeToUom[rowCode] || 'KG';
      }

      visRows.push({
        rowIndex: i + 1,
        rowData: row,
        inventoryId: String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || ''),
        itemId: String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || ''),
        itemCode: String(resolvedItemCode || '').trim(),
        batchId: norm(row[CONFIG.INVENTORY_COLS.BATCH_ID]),
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
    Logger.log('[QA_BACKEND] Visible rows loaded: ' + visRows.length);

    const result = [];
    visRows.forEach(r => {
      let status = _normalizeQaStatus(r.qualityStatus || 'PENDING');
      // Exclude resolved rows from QA worklist.
      if (status === 'APPROVED' || status === 'REJECTED') {
        Logger.log('[QA_BACKEND] Excluding resolved row: ' + r.inventoryId + ' status=' + status);
        return;
      }

      const invId = String(r.inventoryId || '');
      const override = overrideMap[invId] || {};
      Logger.log('[QA_BACKEND] Adding ' + invId + ' with status ' + status);

      // Convert dates to ISO strings for JSON serialization
      const qualityDateVal = r.qualityDate;
      const qualityDateStr = qualityDateVal ? (qualityDateVal instanceof Date ? qualityDateVal.toISOString() : String(qualityDateVal)) : '';
      const overriddenAtVal = override.overriddenAt;
      const overriddenAtStr = overriddenAtVal ? (overriddenAtVal instanceof Date ? overriddenAtVal.toISOString() : String(overriddenAtVal)) : '';

      result.push({
        inventoryId: invId,
        itemId: String(r.itemId || ''),
        itemCode: String(r.itemCode || ''),
        batchId: String(r.batchId || ''),
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
    Logger.log('[QA_BACKEND] Final result length: ' + result.length);
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
    const fBatch = norm(f.batchId || '');
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
      if (fBatch && rowBatch !== fBatch) continue;
      if (fBin && rowBin !== fBin) continue;

      rows.push({
        rowIndex: i + 1,
        rowData: row,
        inventoryId: String(row[CONFIG.INVENTORY_COLS.INVENTORY_ID] || ''),
        itemId: String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || ''),
        itemCode: String(resolvedItemCode || '').trim(),
        batchId: String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim(),
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


// =====================================================
// 7. PUBLIC TRANSACTION ENDPOINTS
// =====================================================

function submitInwardV2(form) {
  return withScriptLock(function () {
    protect(() => requireOperationalUser());
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
    const batchByItem = {};
    expandedItems.forEach(function (item) {
      const itemCode = String(item.itemCode || '').trim();
      if (!itemCode) throw new Error('Inward requires Item Code');
      const qtyKg = _convertToKg(itemCode, item.quantity, item.uom);
      if (!isFinite(qtyKg) || qtyKg <= 0) {
        throw new Error('Inward quantity must be greater than 0 for ' + itemCode);
      }
      // Manual inward entry: batch is entered by user; no Batch_Master lookup here.
      let batchId = String(item.batchId || item.batchNumber || '').trim();
      if (!batchId) {
        throw new Error('Batch No required for inward receipt: ' + itemCode);
      }
      const qualityStatus = _normalizeQaStatus(String(item.qualityStatus || 'PENDING'));
      const qualityDate = item.qualityDate ? new Date(item.qualityDate) : '';
      const version = String(item.version || 'V1').trim();
      const lotNo = String(item.lotNo || '').trim();
      const expiryDate = item.expiryDate ? new Date(item.expiryDate) : '';

      if (!/^V\d+$/i.test(version)) {
        throw new Error('Version must be V1, V2, etc.');
      }

      if (batchByItem[itemCode] && batchByItem[itemCode] !== batchId) {
        throw new Error(`Inward batch_id mismatch for ${itemCode}`);
      }
      batchByItem[itemCode] = batchId;
      const ids = _assertItemCodeBatch(itemCode, batchId, 'Inward');
      const lookup = _getValidatedItemId(ids.itemCode);

      _assertBinCapacity(item.binId, qtyKg, binDelta);
      binDelta[item.binId] = (binDelta[item.binId] || 0) + qtyKg;

      const existing = _readInventoryState(
        ids.itemCode, ids.batchId, item.binId, version
      ).filter(r => String(r.rowData[CONFIG.INVENTORY_COLS.GIN_NO]) === ginNo);



      _appendMovementLog({
        type: CONFIG.MOVEMENT_TYPES.INWARD,
        itemId: lookup, batchId: ids.batchId,
        version: version, ginNo: ginNo,
        toBinId: item.binId, quantity: qtyKg,
        qualityStatus: qualityStatus, remarks: form.remarks
      });

      if (existing.length > 0) {
        const row = existing[0];
        _persistInventoryState(row.rowIndex, row.currentQty + qtyKg, sheet);
        if (typeof CONFIG.INVENTORY_COLS.INWARD_DATE === 'number') {
          const inwardExisting = row.rowData[CONFIG.INVENTORY_COLS.INWARD_DATE];
          if (!inwardExisting) {
            const now = new Date();
            sheet.getRange(row.rowIndex, CONFIG.INVENTORY_COLS.INWARD_DATE + 1).setValue(now);
            _updateSheetCacheCell(CONFIG.SHEETS.INVENTORY, row.rowIndex, CONFIG.INVENTORY_COLS.INWARD_DATE + 1, now);
          }
        }
      } else {
        _createInventoryRow({
          itemId: lookup,
          itemCode: ids.itemCode,
          batchId: ids.batchId, ginNo: ginNo,
          version: version, qualityStatus: qualityStatus,
          qualityDate: qualityDate,
          binId: item.binId, site: item.site, location: item.location,
          quantity: qtyKg,
          lotNo: lotNo,
          expiryDate: expiryDate
        });
      }
    });
    try { _updateBinAndRackStatuses(Object.keys(binDelta)); } catch (e) { Logger.log('Bin status update failed: ' + e.message); }
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
      if (qtyKg > 10000) {
        throw new Error('Quantity exceeds single-transaction limit of 10,000 KG - contact manager');
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
      if (qtyKgInput > 10000) {
        throw new Error('Quantity exceeds single-transaction limit of 10,000 KG - contact manager');
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
      const qtyKg = _convertToKg(item.itemCode, item.quantity, item.uom);
      if (qtyKg > 10000) {
        throw new Error('Quantity exceeds single-transaction limit of 10,000 KG - contact manager');
      }

      // CRITICAL FIX: Server-authoritative item validation
      // Ignore client-provided itemId entirely - only trust itemCode
      const lookup = _getValidatedItemId(item.itemCode);
      const ledgerState = _getProductionLedgerState(form.prodOrderNo, lookup, (ledgerBatchId || item.batchId), '');
      if (ledgerState.rowIndex === -1) {
        throw new Error('Production ledger line not found for PO ' + form.prodOrderNo + ' / ' + item.itemCode);
      }
      if (String(ledgerState.status || '').trim().toUpperCase() !== 'APPROVED') {
        throw new Error('Cannot consume unapproved PO line for ' + item.itemCode + '. Current status: ' + ledgerState.status);
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
          lookup,
          (ledgerBatchId || item.batchId),
          itemVersion,
          qtyKg,
          0
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
        // Check what version exists in destination bin
        const destRows = invData.filter(function (row) {
          return row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] === t.itemCode &&
            row[CONFIG.INVENTORY_COLS.BATCH_ID] === t.batchId &&
            row[CONFIG.INVENTORY_COLS.BIN_ID] === t.toBin &&
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
      for (let i = 1; i < invData.length; i++) {
        const row = invData[i];
        if (String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim() === t.itemCode &&
          String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim() === t.batchId &&
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

function _getProductionLedgerState(orderRef, itemId, batchId, version) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
  const data = _getSheetValuesCached(sheet.getName());

  // Resolve target identifier to both ID and Code
  let targetId = String(itemId || '').trim().toUpperCase();
  let targetCode = String(itemId || '').trim().toUpperCase();
  try {
    const item = getItemByCodeCached(itemId) || _getItemByIdCached(itemId);
    if (item) {
      targetId = String(item.id || '').trim().toUpperCase();
      targetCode = String(item.code || '').trim().toUpperCase();
    }
  } catch (e) { }

  const searchBatch = String(batchId || '').trim().toUpperCase();
  const searchVersion = String(version || '').trim().toUpperCase();
  const candidates = [];
  const candidatesAnyBatch = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowOrderId = String(row[CONFIG.PRODUCTION_LEDGER_COLS.ORDER_ID]);
    const rowItemId = String(row[CONFIG.PRODUCTION_LEDGER_COLS.ITEM_ID]).trim().toUpperCase();

    if (rowOrderId === String(orderRef) && (rowItemId === targetId || rowItemId === targetCode)) {
      const rowBatch = String(row[CONFIG.PRODUCTION_LEDGER_COLS.BATCH_ID] || '').trim().toUpperCase();
      const rowVersion = String(row[CONFIG.PRODUCTION_LEDGER_COLS.VERSION] || '').trim().toUpperCase();
      const statusIdx = _detectProductionLedgerStatusIdx(row, CONFIG.PRODUCTION_LEDGER_COLS.STATUS);
      const updatedIdx = _detectProductionLedgerUpdatedIdx(row, statusIdx, CONFIG.PRODUCTION_LEDGER_COLS.LAST_UPDATED);
      const status = row[statusIdx];

      const batchIsPlaceholder = (rowBatch === '' || rowBatch === 'TBD');
      const versionIsPlaceholder = (rowVersion === '' || rowVersion === 'TBD');

      const batchMatches = searchBatch ? (rowBatch === searchBatch || batchIsPlaceholder) : true;
      const versionMatches = searchVersion ? (rowVersion === searchVersion || versionIsPlaceholder) : true;
      const versionMatchesAnyBatch = searchVersion ? (rowVersion === searchVersion || versionIsPlaceholder) : true;

      if (versionMatchesAnyBatch) {
        candidatesAnyBatch.push({
          rowIndex: i + 1,
          requested: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REQUESTED]) || 0,
          issued: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_ISSUED]) || 0,
          returned: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_RETURNED]) || 0,
          rejected: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REJECTED]) || 0,
          status: status,
          statusIdx: statusIdx,
          updatedIdx: updatedIdx,
          version: rowVersion,
          updateBatch: false,
          updateVersion: !!searchVersion && versionIsPlaceholder
        });
      }

      if (batchMatches && versionMatches) {
        candidates.push({
          rowIndex: i + 1,
          requested: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REQUESTED]) || 0,
          issued: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_ISSUED]) || 0,
          returned: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_RETURNED]) || 0,
          rejected: Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REJECTED]) || 0,
          status: status,
          statusIdx: statusIdx,
          updatedIdx: updatedIdx,
          version: rowVersion,
          updateBatch: !!searchBatch && batchIsPlaceholder,
          updateVersion: !!searchVersion && versionIsPlaceholder
        });
      }
    }
  }

  if (candidates.length > 1) {
    throw new Error(`Duplicate production ledger rows found for ${orderRef} / ${itemId} / ${batchId || 'TBD'} / ${version || 'TBD'}`);
  }
  if (candidates.length === 1) return candidates[0];
  if (searchBatch) {
    if (candidatesAnyBatch.length > 1) {
      throw new Error(`Multiple production ledger rows found for ${orderRef} / ${itemId}. Cannot infer unique Job Order row.`);
    }
    if (candidatesAnyBatch.length === 1) return candidatesAnyBatch[0];
  }
  return {
    rowIndex: -1, requested: 0, issued: 0, returned: 0, rejected: 0,
    status: 'NOT_FOUND', statusIdx: CONFIG.PRODUCTION_LEDGER_COLS.STATUS,
    updatedIdx: CONFIG.PRODUCTION_LEDGER_COLS.LAST_UPDATED,
    version: '', updateBatch: false, updateVersion: false
  };
}

/**
 * Updates production actuals
 */
function _updateProductionLedger(orderRef, itemId, batchId, version, qtyIssuedInc, qtyReturnedInc, qtyRejectedInc) {
  const state = _getProductionLedgerState(orderRef, itemId, batchId, version);

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
  return change;
}



// =====================================================
// PRODUCTION REQUEST FLOW
// =====================================================


/**
 * Get pending production orders for manager approval
 * @returns {Array} List of pending production orders
 */

function getPendingProductionOrders() {
  _beginRequestCache();
  try {
    // Auth guard - Manager only
    Logger.log('[PO_BACKEND] Checking manager access...');
    requireRole(SECURITY.ROLES.MANAGER);

    Logger.log('[PO_BACKEND] Loading production ledger...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
    const data = _getSheetValuesCached(sheet.getName());
    const itemMaps = _getItemMasterMaps();
    Logger.log('[PO_BACKEND] Production ledger rows: ' + data.length);

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

      // Include OPEN (pending approval) and exclude APPROVED/CLOSED
      if (status === 'APPROVED' || status === 'CLOSED' || status === 'REJECTED') {
        continue;
      }

      const orderRef = String(row[CONFIG.PRODUCTION_LEDGER_COLS.ORDER_ID] || '');
      const itemId = String(row[CONFIG.PRODUCTION_LEDGER_COLS.ITEM_ID] || '');
      // Then inside loop:
      const itemCode = itemMaps.idToCode[String(itemId || '').toUpperCase()] || itemId;
      const uomCode = itemMaps.codeToUom[String(itemCode || '').trim().toUpperCase()] || _getItemUomCode(itemCode);
      const batchId = String(row[CONFIG.PRODUCTION_LEDGER_COLS.BATCH_ID] || '');
      const quantity = Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REQUESTED]) || Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_ISSUED]) || 0;
      const rejected = Number(row[CONFIG.PRODUCTION_LEDGER_COLS.QTY_REJECTED]) || 0;
      const outstanding = Number(row[CONFIG.PRODUCTION_LEDGER_COLS.NET_OUTSTANDING]) || 0;
      const lastUpdated = row[detectUpdatedIdx(row, statusIdx)] || new Date();

      Logger.log('[PO_BACKEND] Adding pending PO: ' + orderRef + ' status=' + status);

      result.push({
        orderRef: orderRef,
        itemId: itemId,
        itemCode: itemCode,
        uomCode: uomCode,
        batchId: batchId,
        quantity: quantity,
        rejected: rejected,
        outstanding: outstanding,
        status: status,
        lastUpdated: lastUpdated instanceof Date ? lastUpdated.toISOString() : String(lastUpdated)
      });
    }

    Logger.log('[PO_BACKEND] Final pending POs: ' + result.length);
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
    Logger.log('[PO_BACKEND] Update request: ' + payload.orderRef + ' -> ' + payload.newStatus);

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

    Logger.log('[PO_BACKEND] Updated line for ' + targetOrderRef + ' to ' + newStatus);
    return { success: true, message: `Production Order ${newStatus}`, orderRef: targetOrderRef };
  });
}

/**
 * Validates that a Production Order is APPROVED before allowing consumption/return
 * @param {String} orderRef - Production Order reference ID
 * @throws {Error} If PO doesn't exist or is not APPROVED
 */
function _validateProductionOrderApproved(orderRef) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
  const data = _getSheetValuesCached(sheet.getName());
  const cols = CONFIG.PRODUCTION_LEDGER_COLS;

  let found = false;
  let hasApproved = false;
  const statuses = {};

  for (let i = 1; i < data.length; i++) {
    const rowOrderRef = String(data[i][cols.ORDER_ID] || '').trim();
    if (rowOrderRef === String(orderRef)) {
      found = true;
      const statusIdx = _detectProductionLedgerStatusIdx(data[i], cols.STATUS);
      const status = String(data[i][statusIdx] || '').trim().toUpperCase();
      statuses[status || 'UNKNOWN'] = true;
      if (status === 'APPROVED') hasApproved = true;
    }
  }

  if (!found) throw new Error(`Production Order ${orderRef} not found`);
  if (!hasApproved) {
    const list = Object.keys(statuses).join(', ');
    throw new Error(`Production Order ${orderRef} has no approved lines yet. Current status: ${list}. Please get manager approval first.`);
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
    const capIdx = (typeof binHeaders['MAX_CAPACITY_KG'] === 'number')
      ? binHeaders['MAX_CAPACITY_KG']
      : ((typeof binHeaders['MAX_CAPACITY'] === 'number') ? binHeaders['MAX_CAPACITY'] : null);
    const capUomIdx = (typeof binHeaders['CAPACITY_UOM'] === 'number')
      ? binHeaders['CAPACITY_UOM']
      : ((typeof binHeaders['MAX_CAPACITY_UOM'] === 'number') ? binHeaders['MAX_CAPACITY_UOM'] : null);
    const bins = {};
    for (let i = 1; i < binData.length; i++) {
      const bId = String(binData[i][0]);
      const rId = String(binData[i][1]);
      const rack = rackMap[rId];
      if (rack) {
        const capUom = capUomIdx === null ? '' : String(binData[i][capUomIdx] || '').trim();
        const resolvedCap = capIdx !== null ? (Number(binData[i][capIdx]) || 0) : 0;
        bins[bId] = {
          id: bId, code: binData[i][2], rackCode: rack.code,
          site: rack.site, location: rack.location,
          maxCapacity: resolvedCap,
          capacityUom: capUom || 'KG',
          currentUsage: 0,
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

      bins[binId].currentUsage += qty;

      // NEW: Collect item info per bin row
      const itemId = String(row[CONFIG.INVENTORY_COLS.ITEM_ID] || '').trim();
      const rawCode = String(row[CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE] || '').trim();
      const itemMeta = itemMap[itemId] || itemMap['__CODE__' + rawCode.toUpperCase()] || { code: rawCode || itemId, name: 'Unknown', uomCode: 'KG' };
      const batchId = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
      const qaStatus = String(row[CONFIG.INVENTORY_COLS.QUALITY_STATUS] || 'PENDING').trim();
      const rowUom = String((CONFIG.INVENTORY_COLS.UOM !== undefined ? row[CONFIG.INVENTORY_COLS.UOM] : '') || '').trim();

      // Merge rows for same item+batch in same bin (can be split across multiple inv rows)
      const existingIdx = bins[binId].items.findIndex(
        it => it.itemId === itemId && it.batchId === batchId
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
          qty: qty,
          qaStatus: qaStatus
        });
      }
    }

    return Object.values(bins).map(function (b) {
      const uomTotals = {};
      (b.items || []).forEach(function (it) {
        const u = String(it.uomCode || '').trim().toUpperCase();
        if (!u) return;
        uomTotals[u] = (uomTotals[u] || 0) + Number(it.qty || 0);
      });
      const uomKeys = Object.keys(uomTotals);
      let usageUom = String(b.capacityUom || 'KG').trim().toUpperCase() || 'KG';
      if (uomKeys.length === 1) {
        usageUom = uomKeys[0];
      } else if (uomKeys.length > 1) {
        uomKeys.sort(function (a, c) { return Number(uomTotals[c] || 0) - Number(uomTotals[a] || 0); });
        usageUom = uomKeys[0] || usageUom;
      }
      return {
        ...b,
        usageUom: usageUom,
        available: Math.max(0, b.maxCapacity - b.currentUsage)
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
        case 'inward': htmlFile = 'InwardForm'; break;
        case 'transfer': htmlFile = 'TransferForm'; break;
        case 'dispatch': htmlFile = 'DispatchForm'; break;
        case 'production_request': htmlFile = 'ProductionRequestForm'; break;
        case 'warehouse': htmlFile = 'WarehouseLayout'; break;
        case 'inventory_dashboard': htmlFile = 'WarehouseItemDashboard'; break;
        case 'dashboard': htmlFile = 'WorkerDashboard'; break;
        default: htmlFile = 'WorkerDashboard';
      }
    }

    else if (user.role === SECURITY.ROLES.MANAGER) {
      switch (formType) {
        case 'inward': htmlFile = 'InwardForm'; break;
        case 'transfer': htmlFile = 'TransferForm'; break;
        case 'dispatch': htmlFile = 'DispatchForm'; break;
        case 'warehouse': htmlFile = 'WarehouseLayout'; break;
        case 'inventory_dashboard': htmlFile = 'WarehouseItemDashboard'; break;
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
        case 'ai_assistant': htmlFile = 'AIAssistant'; break;
        case 'dashboard': htmlFile = 'QAApprovalForm'; break;
        case 'bulk_upload': htmlFile = 'BulkUploadMaster'; break;
        case 'reports':
          htmlFile = 'ReportsScreen';
          break;

        default: htmlFile = 'QAApprovalForm';
      }
    }
    else {
      htmlFile = 'Unauthorized';
    }

    const template = HtmlService.createTemplateFromFile(htmlFile);
    template.userEmail = user.email;
    template.userRole = user.role;
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
    const code = String(itemCode || '').trim().toUpperCase();
    if (!code) return 'V1';

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const invSheet = ss.getSheetByName(CONFIG.SHEETS.INVENTORY);
    const movSheet = ss.getSheetByName(CONFIG.SHEETS.MOVEMENT);

    if (!invSheet && !movSheet) return 'V1';

    let maxVersion = 0;

    function scan(data, codeCol, versionCol) {
      for (let i = 1; i < data.length; i++) {
        const rowCode = String(data[i][codeCol] || '').trim().toUpperCase();
        if (rowCode !== code) continue;

        const raw = String(data[i][versionCol] || '').trim().toUpperCase();
        const m = raw.match(/^V(\d+)$/);
        if (m) maxVersion = Math.max(maxVersion, Number(m[1]));
      }
    }

    if (invSheet) {
      const invData = invSheet.getDataRange().getValues();
      scan(invData, CONFIG.INVENTORY_COLS.ITEM_CODE_CACHE, CONFIG.INVENTORY_COLS.VERSION);
    }

    if (movSheet) {
      const movData = movSheet.getDataRange().getValues();
      scan(movData, CONFIG.MOVEMENT_COLS.ITEM_CODE_CACHE || CONFIG.MOVEMENT_COLS.ITEM_ID, CONFIG.MOVEMENT_COLS.VERSION);
    }

    return 'V' + (maxVersion + 1);
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
    const uomCode = _getItemUomCode(code);
    const nextNum = _parseVersionNum(String(nextVersion).trim().toUpperCase());
    if (nextNum === null || nextNum <= 1) return { message: '', bins: [] };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
    const invData = _getSheetValuesCached(invSheet.getName());

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
      const binMeta = _getBinMetaById(binId);
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
    if (versions.length === 0) return { message: '', bins: [] };

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

    return {
      message: 'Older stock on shelves: ' + preview + extra + '. Receiving ' + nextVersion + '.',
      bins: versions,
      details: details,
      totalOlderQty: summaries.reduce(function (acc, s) { return acc + (Number(s.totalQty) || 0); }, 0)
    };
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
 * Returns the full item list for that approved PO so the form can auto-fill.
 *
 * Returns error objects instead of throwing â€” client checks .error field.
 *
 * @param {string} poNo - Production Order reference (e.g. "PO-20250215-042")
 * @returns {{ items: Array, poNo: string, status: string } | { error: string }}
 */
function getApprovedProductionOrderItems(poNo) {
  return protect(function () {
    if (!poNo) return { error: 'Production Order number required.' };

    const ref = String(poNo).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.PRODUCTION_LEDGER);
    const data = _getSheetValuesCached(sheet.getName());
    const cols = CONFIG.PRODUCTION_LEDGER_COLS;

    var poStatus = null;
    var matchedRows = [];

    for (var i = 1; i < data.length; i++) {
      const row = data[i];
      const rowRef = String(row[cols.ORDER_ID] || '').trim();
      if (rowRef !== ref) continue;

      const statusIdx = _detectProductionLedgerStatusIdx(row, cols.STATUS);
      const status = String(row[statusIdx] || '').trim().toUpperCase();
      poStatus = status; // last seen status for this PO (all rows should match)

      if (status !== 'APPROVED') continue;

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
      if (poStatus && poStatus !== 'APPROVED') {
        return { error: 'PO ' + ref + ' is not approved yet. Current status: ' + poStatus + '. Ask your manager to approve it first.' };
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
      status: 'APPROVED',
      items: pendingItems
    };
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

      const batchId = String(row[CONFIG.INVENTORY_COLS.BATCH_ID] || '').trim();
      if (!batchId) continue;

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
          binSet: {}
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
    }

    // Convert map to array
    const result = [];
    Object.keys(batchMap).forEach(function (batchId) {
      const batch = batchMap[batchId];
      result.push({
        batchId: batch.batchId,
        itemId: batch.itemId,
        approved: batch.approved,
        pending: batch.pending,
        total: batch.total,
        binCount: Object.keys(batch.binSet).length
      });
    });

    // Sort by total stock descending
    result.sort(function (a, b) { return b.total - a.total; });

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
      _assertBinCapacity(form.binId, qtyKg);
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
        _assertBinCapacity(item.binId, qtyKg);
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
    row[CONFIG.INVENTORY_COLS.EXPIRY_DATE] = params.expiryDate || '';
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
              Logger.log('[PROD_REQ] Found item_id=' + itemId + ' for ' + itemCode + ' batch ' + batchId + ' in inventory');
              break;
            }
          }
        }
      }

      // Fallback to Item_Master only if batch not specified or not found in inventory
      if (!itemId) {
        itemId = getItemIdByCode(itemCode) || itemCode;
        Logger.log('[PROD_REQ] Using Item_Master item_id=' + itemId + ' for ' + itemCode + ' (no batch or batch not in inventory)');
      }

      const qty = Number(item.quantity);
      const qtyKg = _convertToKg(itemCode, qty, item.uom);
      const finalBatchId = batchId || 'TBD';

      rowsToWrite.push([
        orderRef, itemId, finalBatchId, 'TBD',
        qtyKg, 0, 0, qtyKg, 0, 'OPEN', new Date()
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
      const raw = String(props.getProperty('ALERT_EMAILS') || props.getProperty('ALERT_EMAIL') || '').trim();
      if (!raw) {
        Logger.log('Min stock alert skipped: ALERT_EMAILS not configured.');
        return;
      }
      const emails = raw.split(/[;,]/).map(function (e) { return e.trim(); }).filter(Boolean);
      if (emails.length === 0) {
        Logger.log('Min stock alert skipped: no valid emails.');
        return;
      }

      const minMap = {};
      const itemMaps = _getItemMasterMaps();
      Object.keys(itemMaps.codeToMinStock).forEach(function (codeNorm) {
        const minLevel = Number(itemMaps.codeToMinStock[codeNorm] || 0);
        if (!(minLevel > 0)) return;
        const code = itemMaps.codeDisplayByNorm[codeNorm] || codeNorm;
        minMap[String(codeNorm || '').trim().toUpperCase()] = {
          code: String(code || codeNorm).trim(),
          codeNorm: String(codeNorm || '').trim().toUpperCase(),
          name: String(itemMaps.codeToName[codeNorm] || code),
          min: minLevel,
          uom: String(itemMaps.codeToUom[codeNorm] || 'KG').trim() || 'KG'
        };
      });

      const invAgg = _aggregateInventorySnapshotByItem();
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const notifiedItems = []; // NEW: structured data instead of text lines
      const invCodes = Object.keys(invAgg || {});
      let belowThresholdCount = 0;
      let dedupedCount = 0;

      // Alert only for items that currently exist in Inventory_Balance (>0 qty rows).
      invCodes.forEach(function (inventoryCode) {
        const codeNorm = String(inventoryCode || '').trim().toUpperCase();
        const minInfo = minMap[codeNorm];
        if (!minInfo) return;
        const current = Number((invAgg[inventoryCode] || {}).totalQty || 0);
        if (current > minInfo.min) return;
        belowThresholdCount += 1;

        const lastKey = 'MIN_STOCK_ALERT_LAST_' + minInfo.codeNorm;
        const lastSent = String(props.getProperty(lastKey) || '').trim();
        if (lastSent === today) {
          dedupedCount += 1;
          return;
        }

        const bySite = (invAgg[inventoryCode] || {}).bySiteLocation || {};
        const locations = Object.keys(bySite).map(function (k) {
          const parts = k.split('||');
          return {
            site: parts[0] || '-',
            location: parts[1] || '-',
            qty: Number(bySite[k] || 0)
          };
        });

        notifiedItems.push({
          code: minInfo.code,
          codeNorm: minInfo.codeNorm,
          name: minInfo.name,
          current: current,
          min: minInfo.min,
          uom: minInfo.uom,
          locations: locations
        });
      });

      if (notifiedItems.length === 0) {
        Logger.log(
          'Min stock alert: no items below threshold today. ' +
          `inventory_items=${invCodes.length}, below_threshold=${belowThresholdCount}, deduped_today=${dedupedCount}`
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

      // Mark sent for today
      notifiedItems.forEach(function (item) {
        props.setProperty('MIN_STOCK_ALERT_LAST_' + String(item.codeNorm || item.code || '').trim().toUpperCase(), today);
      });

      Logger.log('Min stock alert sent for ' + notifiedItems.length + ' item(s).');
    } catch (e) {
      Logger.log('Min stock alert failed: ' + e.message);
    }
  });
}

function getMinStockAlertPreview() {
  return _withRequestCache(function () {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const itemMaps = _getItemMasterMaps();
    const minMap = {};

    Object.keys(itemMaps.codeToMinStock).forEach(function (codeNorm) {
      const minLevel = Number(itemMaps.codeToMinStock[codeNorm] || 0);
      if (!(minLevel > 0)) return;
      const code = itemMaps.codeDisplayByNorm[codeNorm] || codeNorm;
      minMap[String(codeNorm || '').trim().toUpperCase()] = {
        code: String(code || codeNorm).trim(),
        codeNorm: String(codeNorm || '').trim().toUpperCase(),
        min: minLevel,
        uom: String(itemMaps.codeToUom[codeNorm] || 'KG').trim() || 'KG'
      };
    });

    const invAgg = _aggregateInventorySnapshotByItem();
    const rows = [];
    Object.keys(invAgg || {}).forEach(function (inventoryCode) {
      const codeNorm = String(inventoryCode || '').trim().toUpperCase();
      const minInfo = minMap[codeNorm];
      if (!minInfo) return;
      const current = Number((invAgg[inventoryCode] || {}).totalQty || 0);
      const lastSent = String(props.getProperty('MIN_STOCK_ALERT_LAST_' + minInfo.codeNorm) || '').trim();
      rows.push({
        itemCode: minInfo.code,
        current: current,
        min: minInfo.min,
        uom: minInfo.uom,
        belowOrEqualMin: current <= minInfo.min,
        alreadySentToday: lastSent === today,
        lastSentDate: lastSent || ''
      });
    });

    const below = rows.filter(function (r) { return r.belowOrEqualMin; });
    const willSend = below.filter(function (r) { return !r.alreadySentToday; });

    return {
      today: today,
      inventoryTrackedItems: Object.keys(invAgg || {}).length,
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
    const raw = String(props.getProperty('ALERT_EMAILS') || props.getProperty('ALERT_EMAIL') || '').trim();
    const emails = raw ? raw.split(/[;,]/).map(function (e) { return e.trim(); }).filter(Boolean) : [];
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
          ' <span class="qty-inline">— ' + loc.qty + ' ' + _esc(item.uom) + '</span></span>';
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
      '<div class="value danger">' + item.current + '<span class="uom">' + _esc(item.uom) + '</span></div>',
      '</div>',
      '<div class="qty-box">',
      '<div class="label">Minimum Level</div>',
      '<div class="value min">' + item.min + '<span class="uom">' + _esc(item.uom) + '</span></div>',
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

function _appendInventoryRow(rowArray, sheet) {
  if (!sheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = _getSheetOrThrow(ss, CONFIG.SHEETS.INVENTORY);
  }
  sheet.appendRow(rowArray);
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
        var expiry = expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
        if (isNaN(expiry.getTime())) continue;
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
        var expiry = expiryRaw instanceof Date ? expiryRaw : new Date(expiryRaw);
        if (isNaN(expiry.getTime())) continue;
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
    for (let i = 1; i < itemData.length; i++) {
      const code = String(itemData[i][CONFIG.ITEM_COLS.ITEM_CODE] || '').trim().toUpperCase();
      const minStock = Number(itemData[i][CONFIG.ITEM_COLS.MIN_STOCK_LEVEL]) || 0;
      minStockMap[code] = minStock;
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
        _assertBinCapacity(binId, qtyKg);
      } catch (e) {
        errors.push('Row ' + rowNum + ': Bin capacity exceeded');
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
