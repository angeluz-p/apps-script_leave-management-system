function copyToApprovedPaidLeaves(requestData) {
  try {
    const approvedPaidLeaves =
      ACTIVE_SHEET.getSheetByName(APPROVED_PAID_LEAVES);

    if (!approvedPaidLeaves) {
      debugLog("Error: 'Approved Paid Leaves' sheet not found");
      return false;
    }

    // Check if Request ID already exists in Approved Paid Leaves sheet
    const dataRange = approvedPaidLeaves.getDataRange();
    const values = dataRange.getValues();

    for (let i = 1; i < values.length; i++) {
      // Skip header row
      if (values[i][CAT_COL_A_REQUEST_ID] === requestData.requestId) {
        const existingRow = i + 1;
        debugLog(
          `Request ID ${requestData.requestId} already exists in Approved Paid Leaves sheet at row ${existingRow}. Skipping duplicate entry.`
        );
        return true; // Return true to indicate the record already exists (not an error)
      }
    }

    // Get the next empty row in Approved Paid Leaves sheet
    const nextRow = approvedPaidLeaves.getLastRow() + 1;
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MM/dd/yyyy hh:mm:ss a"
    );

    // Prepare the data array for the new row
    const newRowData = [
      timestamp,
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
    const approvedUnpaidLeaves = ACTIVE_SHEET.getSheetByName(
      APPROVED_UNPAID_LEAVES
    );

    if (!approvedUnpaidLeaves) {
      debugLog("Error: 'Approved Unpaid Leaves' sheet not found");
      return false;
    }

    // Check if Request ID already exists in Approved Unpaid Leaves sheet
    const dataRange = approvedUnpaidLeaves.getDataRange();
    const values = dataRange.getValues();

    for (let i = 1; i < values.length; i++) {
      // Skip header row
      if (values[i][CAT_COL_A_REQUEST_ID] === requestData.requestId) {
        const existingRow = i + 1;
        debugLog(
          `Request ID ${requestData.requestId} already exists in Approved Unpaid Leaves sheet at row ${existingRow}. Skipping duplicate entry.`
        );
        return true; // Return true to indicate the record already exists (not an error)
      }
    }

    // Get the next empty row in Approved Unpaid Leaves sheet
    const nextRow = approvedUnpaidLeaves.getLastRow() + 1;
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MM/dd/yyyy hh:mm:ss a"
    );

    // Prepare the data array for the new row
    const newRowData = [
      timestamp,
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
    const rejectedLeaves = ACTIVE_SHEET.getSheetByName(
      REJECTED_LEAVE_APPLICATIONS
    );

    if (!rejectedLeaves) {
      debugLog("Error: 'Rejected Leaves' sheet not found");
      return false;
    }

    // Check if Request ID already exists in Rejected Leaves sheet
    const dataRange = rejectedLeaves.getDataRange();
    const values = dataRange.getValues();

    for (let i = 1; i < values.length; i++) {
      // Skip header row
      if (values[i][CAT_COL_A_REQUEST_ID] === requestData.requestId) {
        const existingRow = i + 1;
        debugLog(
          `Request ID ${requestData.requestId} already exists in Rejected Leaves sheet at row ${existingRow}. Skipping duplicate entry.`
        );
        return true; // Return true to indicate the record already exists (not an error)
      }
    }

    // Get the next empty row in Rejected Leaves sheet
    const nextRow = rejectedLeaves.getLastRow() + 1;
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "MM/dd/yyyy hh:mm:ss a"
    );

    // Prepare the data array for the new row
    const newRowData = [
      timestamp,
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

function copyToCancelledLeaves(requestLeaveData) {
  const cancelledSheet = ACTIVE_SHEET.getSheetByName(
    CANCELLED_LEAVE_APPLICATIONS
  );

  if (!cancelledSheet) {
    debugLog("Error: 'Cancelled Leaves' sheet not found.");
    return;
  }

  // Check if Request ID already exists
  const existingData = targetSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][CAT_COL_A_REQUEST_ID] === requestLeaveData.requestId) {
      debugLog(
        `Request ID ${requestLeaveData.requestId} already exists in Automatically Rejected Leaves. Skipping.`
      );
      return;
    }
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MM/dd/yyyy hh:mm:ss a"
  );

  // Prepare the data to be added
  const newRow = [
    timestamp,
    requestLeaveData.employeeEmail,
    requestLeaveData.employeeName,
    requestLeaveData.jobTitle,
    requestLeaveData.department,
    requestLeaveData.leaveType,
    requestLeaveData.subLeaveType,
    requestLeaveData.startDate,
    requestLeaveData.endDate,
    requestLeaveData.leaveHoursDay,
    requestLeaveData.employeeReason,
    requestLeaveData.employeeAttachments,
    requestLeaveData.supervisorEmail,
    requestLeaveData.acNo,
    requestLeaveData.requestId,
  ];

  // Append the row to the Cancelled Leave Applications sheet
  cancelledSheet.appendRow(newRow);

  debugLog(
    `Added Request ID ${requestLeaveData.requestId} to Cancelled Leave Applications`
  );
}

function copyToAutomaticallyRejectedLeaves(requestLeaveData, rejectionReason) {
  const targetSheet = ACTIVE_SHEET.getSheetByName(
    AUTOMATICALLY_REJECTED_LEAVES
  );

  if (!targetSheet) {
    debugLog("Error: 'Automatically Rejected Leaves' sheet not found.");
    return;
  }

  // Check if Request ID already exists
  const existingData = targetSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][CAT_COL_A_REQUEST_ID] === requestLeaveData.requestId) {
      debugLog(
        `Request ID ${requestLeaveData.requestId} already exists in Automatically Rejected Leaves. Skipping.`
      );
      return;
    }
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MM/dd/yyyy hh:mm:ss a"
  );

  // Append the data with rejection reason
  const newRow = [
    timestamp,
    requestLeaveData.employeeEmail,
    requestLeaveData.employeeName,
    requestLeaveData.jobTitle,
    requestLeaveData.department,
    requestLeaveData.leaveType,
    requestLeaveData.subLeaveType,
    requestLeaveData.startDate,
    requestLeaveData.endDate,
    requestLeaveData.leaveHoursDay,
    requestLeaveData.employeeReason,
    requestLeaveData.employeeAttachments,
    requestLeaveData.supervisorEmail,
    requestLeaveData.acNo,
    requestLeaveData.requestId,
    rejectionReason,
  ];

  targetSheet.appendRow(newRow);

  debugLog(
    `Request ID ${requestLeaveData.requestId} copied to Automatically Rejected Leaves`
  );
}
