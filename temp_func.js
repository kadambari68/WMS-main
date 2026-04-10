function getRejectedItemsReportData() {
  return protect(function () {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const qaSheet = ss.getSheetByName(CONFIG.SHEETS.QA_EVENTS);

      if (!qaSheet) {
        return { rejections: [], summary: [], _error: 'QA_Events sheet not found' };
      }

      const data = qaSheet.getDataRange().getValues();
      const rejections = [];
      const summary = {};

      // Supported reporting actions
      const TARGET_ACTIONS = ['REJECT', 'HOLD', 'REJECTED', 'APPROVE_PARTIAL', 'HOLD_PARTIAL_BALANCE'];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const action = String(row[CONFIG.QA_EVENTS_COLS.ACTION] || '').trim().toUpperCase();

        if (TARGET_ACTIONS.indexOf(action) === -1) continue;

        const itemCode = String(row[CONFIG.QA_EVENTS_COLS.ITEM_CODE] || '').trim();
        const batchId = String(row[CONFIG.QA_EVENTS_COLS.BATCH_ID] || '').trim();
        const binId = String(row[CONFIG.QA_EVENTS_COLS.BIN_ID] || '').trim();

        // Use Overridden At or Timestamp
        let ts = row[CONFIG.QA_EVENTS_COLS.OVERRIDDEN_AT] || row[CONFIG.QA_EVENTS_COLS.TIMESTAMP];
        let tsStr = ts instanceof Date ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm') : String(ts || '-');

        const reason = String(row[CONFIG.QA_EVENTS_COLS.OVERRIDE_REASON] || '').trim();
        const prevStatus = String(row[CONFIG.QA_EVENTS_COLS.PREV_STATUS] || '').trim();
        const newStatus = String(row[CONFIG.QA_EVENTS_COLS.NEW_STATUS] || '').trim();
        const remarks = String(row[CONFIG.QA_EVENTS_COLS.REMARKS] || '').trim();
        const UserEmail = String(row[CONFIG.QA_EVENTS_COLS.OVERRIDDEN_BY] || '').trim();

        const record = {
          timestamp: tsStr,
          itemCode: itemCode,
          batchId: batchId,
          binId: binId,
          action: action,
          prevStatus: prevStatus,
          newStatus: newStatus,
          reason: reason,
          remarks: remarks,
          user: UserEmail
        };

        rejections.push(record);

        if (itemCode) {
          if (!summary[itemCode]) {
            summary[itemCode] = { itemCode: itemCode, rejected: 0, held: 0, partial: 0 };
          }
          if (action === 'REJECT' || action === 'REJECTED') summary[itemCode].rejected++;
          else if (action === 'HOLD') summary[itemCode].held++;
          else if (action.indexOf('PARTIAL') !== -1) summary[itemCode].partial++;
        }
      }

      const summaryList = Object.keys(summary).map(k => summary[k]);

      return {
        rejections: rejections,
        summary: summaryList
      };
    } catch (e) {
      Logger.log('[REJECTED_REPORT] Error: ' + e.toString());
      return { rejections: [], summary: [], _error: e.toString() };
    }
  });
}
