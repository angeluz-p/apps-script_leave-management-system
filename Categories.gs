function copyToApprovedPaidLeaves(requestData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const approvedPaidLeaves = ss.getSheetByName(APPROVED_PAID_LEAVES);

    if (!approvedPaidLeaves) {
      debugLog("Error: 'Approved Paid Leaves' sheet not found");
      return false;
    }

    // Check if Request ID already exists in Approved Paid Leaves sheet
    /* const dataRange = approvedPaidLeaves.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][9] === requestData.requestId) { 
        const existingRow = i + 1;
        debugLog(`Request ID ${requestData.requestId} already exists in Approved Paid Leaves sheet at row ${existingRow}. Updating status columns.`);
        
        approvedPaidLeaves.getRange(existingRow, 11).setValue(requestData.notifiedStatus);
        approvedPaidLeaves.getRange(existingRow, 12).setValue(requestData.forPurchaseStatus);
        approvedPaidLeaves.getRange(existingRow, 13).setValue(requestData.completedStatus);
        
        debugLog(`Successfully updated status columns for Request ID ${requestData.requestId} at row ${existingRow}`);
        return true; // Return true since the record was successfully updated
      }
    } */

    // Get the next empty row in Approved Paid Leaves sheet
    const nextRow = approvedPaidLeaves.getLastRow() + 1;

    // Prepare the data array for the new row
    const newRowData = [
      requestData.employeeEmail,
      requestData.employeeName,
      requestData.jobTitle,
      requestData.department,
      requestData.leaveType,
      requestData.subLeaveType,
      requestData.startDate,
      requestData.endDate,
      requestData.leaveHoursDay,
      requestData.employeeReason,
      requestData.employeeAttachments,
      requestData.supervisorEmail,
      requestData.acNo,
      requestData.requestId,
    ];

    // Add the data to Approved Paid Leaves sheet
    approvedPaidLeaves
      .getRange(nextRow, 1, 1, newRowData.length)
      .setValues([newRowData]);

    debugLog(
      `Successfully copied Request ID ${requestData.requestId} to Approved Paid Leaves sheet at row ${nextRow}`
    );
    return true;
  } catch (error) {
    Logger.log(`Error copying to Approved Paid Leaves sheet: ${error.message}`);
    return false;
  }
}

function copyToApprovedUnpaidLeaves(requestData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const approvedUnpaidLeaves = ss.getSheetByName(APPROVED_UNPAID_LEAVES);

    if (!approvedUnpaidLeaves) {
      debugLog("Error: 'Approved Unpaid Leavess' sheet not found");
      return false;
    }

    // Check if Request ID already exists in Approved Unpaid Leaves sheet
    /* const dataRange = approvedUnpaidLeaves.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][9] === requestData.requestId) { 
        const existingRow = i + 1;
        debugLog(`Request ID ${requestData.requestId} already exists in Approved Unpaid Leaves sheet at row ${existingRow}. Updating status columns.`);
        
        approvedUnpaidLeaves.getRange(existingRow, 11).setValue(requestData.notifiedStatus);
        approvedUnpaidLeaves.getRange(existingRow, 12).setValue(requestData.forPurchaseStatus);
        approvedUnpaidLeaves.getRange(existingRow, 13).setValue(requestData.completedStatus);
        
        debugLog(`Successfully updated status columns for Request ID ${requestData.requestId} at row ${existingRow}`);
        return true; // Return true since the record was successfully updated
      }
    } */

    // Get the next empty row in Approved Unpaid Leaves sheet
    const nextRow = approvedUnpaidLeaves.getLastRow() + 1;

    // Prepare the data array for the new row
    const newRowData = [
      requestData.employeeEmail,
      requestData.employeeName,
      requestData.jobTitle,
      requestData.department,
      requestData.leaveType,
      requestData.subLeaveType,
      requestData.startDate,
      requestData.endDate,
      requestData.leaveHoursDay,
      requestData.employeeReason,
      requestData.employeeAttachments,
      requestData.supervisorEmail,
      requestData.acNo,
      requestData.requestId,
    ];

    // Add the data to Approved Unpaid Leaves sheet
    approvedUnpaidLeaves
      .getRange(nextRow, 1, 1, newRowData.length)
      .setValues([newRowData]);

    debugLog(
      `Successfully copied Request ID ${requestData.requestId} to Approved Unpaid Leaves sheet at row ${nextRow}`
    );
    return true;
  } catch (error) {
    Logger.log(
      `Error copying to Approved Unpaid Leaves sheet: ${error.message}`
    );
    return false;
  }
}

function copyToRejectedLeaves(requestData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rejectedLeaves = ss.getSheetByName(REJECTED_LEAVE_APPLICATIONS);

    if (!rejectedLeaves) {
      debugLog("Error: 'Rejected Leaves' sheet not found");
      return false;
    }

    // Check if Request ID already exists in Rejected Leavess sheet
    /* const dataRange = rejectedLeaves.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][9] === requestData.requestId) { 
        const existingRow = i + 1;
        debugLog(`Request ID ${requestData.requestId} already exists in Rejected Leaves sheet at row ${existingRow}. Updating status columns.`);
        
        rejectedLeaves.getRange(existingRow, 11).setValue(requestData.notifiedStatus);
        rejectedLeaves.getRange(existingRow, 12).setValue(requestData.forPurchaseStatus);
        rejectedLeaves.getRange(existingRow, 13).setValue(requestData.completedStatus);
        
        debugLog(`Successfully updated status columns for Request ID ${requestData.requestId} at row ${existingRow}`);
        return true; // Return true since the record was successfully updated
      }
    } */

    // Get the next empty row in Rejected Leaves sheet
    const nextRow = rejectedLeaves.getLastRow() + 1;

    // Prepare the data array for the new row
    const newRowData = [
      requestData.employeeEmail,
      requestData.employeeName,
      requestData.jobTitle,
      requestData.department,
      requestData.leaveType,
      requestData.subLeaveType,
      requestData.startDate,
      requestData.endDate,
      requestData.leaveHoursDay,
      requestData.employeeReason,
      requestData.employeeAttachments,
      requestData.supervisorEmail,
      requestData.acNo,
      requestData.requestId,
    ];

    // Add the data to Rejected Leaves sheet
    rejectedLeaves
      .getRange(nextRow, 1, 1, newRowData.length)
      .setValues([newRowData]);

    debugLog(
      `Successfully copied Request ID ${requestData.requestId} to Rejected Leaves sheet at row ${nextRow}`
    );
    return true;
  } catch (error) {
    Logger.log(`Error copying to Rejected Leaves sheet: ${error.message}`);
    return false;
  }
}
