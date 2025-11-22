// Log Approved Paid Leaves
function logApprovedPaidLeave(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName(ALL_RECORDS);

  if (!formResponsesSheet) {
    Logger.log("Error: 'All Records' sheet not found.");
    return;
  }

  let approvedPaidLeavesSheet = ss.getSheetByName(APPROVED_PAID_LEAVES);

  // Create "Approved Paid Leaves" sheet if not exists
  if (!approvedPaidLeavesSheet) {
    Logger.log("Creating 'Approved Paid Leaves' sheet...");
    approvedPaidLeavesSheet = ss.insertSheet(APPROVED_PAID_LEAVES);

    // Copy headers from "All Records"
    const headers = formResponsesSheet.getRange(1, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];
    approvedPaidLeavesSheet.appendRow(headers);
  }

  // Validate row number
  if (!row || isNaN(row) || row < 2) {
    Logger.log(`Error: Invalid row number: ${row}`);
    return;
  }

  const lastColumn = formResponsesSheet.getLastColumn();
  const rowData = formResponsesSheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  // Ensure rowData contains values
  if (!rowData || rowData.length === 0) {
    Logger.log(`Error: No data found for row ${row}.`);
    return;
  }

  // Log for debugging
  Logger.log(`Logging paid leave for row ${row}: ${JSON.stringify(rowData)}`);

  approvedPaidLeavesSheet.appendRow(rowData);

  Logger.log("Successfully logged paid leave.");
}

// Log Approved Unpaid Leaves
function logApprovedUnpaidLeave(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName(ALL_RECORDS);

  if (!formResponsesSheet) {
    Logger.log("Error: 'All Records' sheet not found.");
    return;
  }

  let approvedUnpaidLeavesSheet = ss.getSheetByName(APPROVED_UNPAID_LEAVES);

  // Create "Approved Unpaid Leaves" sheet if not exists
  if (!approvedUnpaidLeavesSheet) {
    Logger.log("Creating 'Approved Unpaid Leaves' sheet...");
    approvedUnpaidLeavesSheet = ss.insertSheet(APPROVED_UNPAID_LEAVES);

    // Copy headers from "All Records"
    const headers = formResponsesSheet.getRange(1, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];
    approvedUnpaidLeavesSheet.appendRow(headers);
  }

  // Validate row number
  if (!row || isNaN(row) || row < 2) {
    Logger.log(`Error: Invalid row number: ${row}`);
    return;
  }

  const lastColumn = formResponsesSheet.getLastColumn();
  const rowData = formResponsesSheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  // Ensure rowData contains values
  if (!rowData || rowData.length === 0) {
    Logger.log(`Error: No data found for row ${row}.`);
    return;
  }

  // Log for debugging
  Logger.log(`Logging paid leave for row ${row}: ${JSON.stringify(rowData)}`);

  approvedUnpaidLeavesSheet.appendRow(rowData);

  Logger.log("Successfully logged paid leave.");
}

// Log Rejected Leave Applications
function logRejectedLeaveApplications(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = ss.getSheetByName(ALL_RECORDS);

  if (!formResponsesSheet) {
    Logger.log("Error: 'All Records' sheet not found.");
    return;
  }

  let rejectedLeaveApplicationsSheet = ss.getSheetByName(REJECTED_LEAVE_APPLICATIONS);

  // Check and create "Rejected Leave Applications" sheet if not exists
  if (!rejectedLeaveApplicationsSheet) {
    Logger.log("Rejected Leaves sheet does not exist. Creating it...");
    rejectedLeaveApplicationsSheet = ss.insertSheet(REJECTED_LEAVE_APPLICATIONS);
    
    // Copy headers from "All Records"
    const headers = formResponsesSheet.getRange(1, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];
    rejectedLeaveApplicationsSheet.appendRow(headers);
  }

  // Validate row number
  if (!row || isNaN(row) || row < 2) {
    Logger.log(`Error: Invalid row number: ${row}`);
    return;
  }

  const lastColumn = formResponsesSheet.getLastColumn();
  const rowData = formResponsesSheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  // Ensure rowData contains values
  if (!rowData || rowData.length === 0) {
    Logger.log(`Error: No data found for row ${row}.`);
    return;
  }

  // Log for debugging
  Logger.log(`Logging rejected leave applications leave for row ${row}: ${JSON.stringify(rowData)}`);

  rejectedLeaveApplicationsSheet.appendRow(rowData);

  Logger.log("Successfully logged rejected leave applications leave.");
}