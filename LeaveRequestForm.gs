// ============================================
// Edit action - automatically update for certain condition
// ============================================
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const editedColumn = e.range.getColumn();
  const editedRow = e.range.getRow();
  const ss = ACTIVE_SHEET;

  const leaveBalanceUpdate =
    sheetName === ALL_RECORDS &&
    [
      COL_N_EMAIL_ADDRESS,
      COL_N_MAIN_LEAVE,
      COL_N_SUB_LEAVE,
      COL_N_START_DATE,
      COL_N_END_DATE,
      COL_N_LEAVE_DURATION,
      COL_N_MAIN_STATUS,
    ].includes(editedColumn);
  const manualUpdate =
    sheetName === ALL_RECORDS &&
    editedColumn === COL_N_MAIN_STATUS &&
    editedRow > 1;

  if (manualUpdate) {
    const newStatus = e.range.getValue();
    const requestId = sheet.getRange(editedRow, COL_N_REQUEST_ID).getValue();

    // Only process if there's a valid request ID and status changed
    if (requestId && newStatus) {
      // Show processing toast based on status
      let processingMessage = "";
      if (newStatus === "Cancelled") {
        processingMessage = "Processing cancellation...";
      } else if (newStatus === "Approved") {
        processingMessage = "Processing approval...";
      } else if (newStatus === "Rejected") {
        processingMessage = "Processing rejection...";
      }

      if (processingMessage) {
        ss.toast(processingMessage, "Status Update", -1);

        try {
          handleManualStatusChange(editedRow, newStatus, requestId);

          // Show success toast based on status
          let successMessage = "";
          if (newStatus === "Cancelled") {
            successMessage =
              "Leave request successfully cancelled and moved to Cancelled Leaves";
          } else if (newStatus === "Approved") {
            successMessage =
              "Leave request successfully approved and categorized";
          } else if (newStatus === "Rejected") {
            successMessage =
              "Leave request successfully rejected and moved to Rejected Leaves";
          }

          ss.toast(successMessage, "Success", 5);
        } catch (error) {
          // Show error toast
          ss.toast(
            "Error processing status change: " + error.message,
            "Error",
            10
          );
          debugLog("Error in onEdit status change: " + error.message);
        }
      }
    }
  }

  if (leaveBalanceUpdate) {
    // Get the leave type and sub-type to check if balance update is needed
    const leaveType = sheet.getRange(editedRow, COL_N_MAIN_LEAVE).getValue();
    const subLeaveType = sheet.getRange(editedRow, COL_N_SUB_LEAVE).getValue();
    const status = sheet.getRange(editedRow, COL_N_MAIN_STATUS).getValue();

    // Only show toast for leaves that affect balances (Approved Paid leaves)
    const affectsBalance =
      status === "Approved" &&
      ((leaveType === "Vacation Leave" && subLeaveType === "VL/SL") ||
        (leaveType === "Sick Leave" && subLeaveType === "VL/SL") ||
        (leaveType === "Vacation Leave" && subLeaveType === "Unexcused") ||
        (leaveType === "Sick Leave" && subLeaveType === "Unexcused") ||
        (leaveType === "Sick Leave" && subLeaveType === "Emergency Leave") ||
        (leaveType === "Vacation Leave" &&
          subLeaveType === "Emergency Leave") ||
        (leaveType === "Other" && subLeaveType === "Bereavement Leave") ||
        (leaveType === "Other" && subLeaveType === "Paternity Leave") ||
        (leaveType === "Other" && subLeaveType === "Maternity Leave"));

    if (affectsBalance) {
      // Show processing toast only for approved paid leaves
      ss.toast("Updating leave balances...", "Processing", -1);
    }

    try {
      updateLeaveBalances();

      if (affectsBalance) {
        // Show success toast only for approved paid leaves
        ss.toast("Leave balances updated successfully", "Success", 3);
      }
    } catch (error) {
      // Always show error toast if update fails
      ss.toast("Error updating leave balances: " + error.message, "Error", 10);
      debugLog("Error updating leave balances: " + error.message);
    }
  }
}

// ============================================
// Form Submission
// ============================================
function onFormSubmit(e) {
  const ss = ACTIVE_SHEET;
  const sheet = ss.getSheetByName(ALL_RECORDS);
  const balancesSheet = ss.getSheetByName(LEAVE_BALANCES);
  const balancesData = balancesSheet.getDataRange().getValues();

  if (!sheet || !balancesSheet) {
    Logger.log("Error: One or both sheets not found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  const row = sheet
    .getRange(lastRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  debugLog("Row Data: " + row);

  const existingRequestId = sheet
    .getRange(lastRow, COL_N_REQUEST_ID)
    .getValue();
  let requestId;

  if (!existingRequestId) {
    requestId = generateRequestIdEnhanced(sheet);
    sheet.getRange(lastRow, COL_N_REQUEST_ID).setValue(requestId);
    debugLog(
      `Generated new Request ID ${requestId} in row ${lastRow}, column Q`
    );
  } else {
    requestId = existingRequestId;
    debugLog(
      `Using existing Request ID ${requestId} in row ${lastRow}, column Q`
    );
  }

  // Add dropdown to the new row's Status column
  const statusOptions = [
    "Approved",
    "Rejected",
    "Cancelled",
    "Automatically Rejected",
  ];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(lastRow, COL_N_MAIN_STATUS).setDataValidation(rule);

  const employeeEmail = row[COL_A_EMAIL_ADDRESS];
  const employeeName = row[COL_A_FULL_NAME];
  let employeeFormatName = formatName(employeeName);

  const leaveType = row[COL_A_MAIN_LEAVE];
  const subLeaveType = row[COL_A_SUB_LEAVE];
  const leaveStartDate = row[COL_A_START_DATE];
  const leaveEndDate = row[COL_A_END_DATE];
  const leaveHoursDay = row[COL_A_LEAVE_DURATION];
  const employeeReason = row[COL_A_REASON];
  const employeeAttachments = row[COL_A_ATTACHMENT];
  const supervisorEmail = row[COL_A_APPROVER]?.trim(); // Ensure it exists and remove extra spaces
  const statusColumn = COL_N_MAIN_STATUS; // Change this based on your Google Sheet column index for "Status"
  const forwardedStatusColumn = COL_N_FORWARDING_STATUS;
  const acNo = row[COL_A_AC_NO]; // ADD THIS if it's missing
  const jobTitle = row[COL_A_JOB_TITLE]; // ADD THIS if it's missing
  const department = row[COL_A_DEPARTMENT]; // ADD THIS if it's missing

  const requestLeaveData = {
    employeeEmail: employeeEmail,
    employeeName: employeeName,
    jobTitle: jobTitle,
    department: department,
    leaveType: leaveType,
    subLeaveType: subLeaveType,
    startDate: leaveStartDate,
    endDate: leaveEndDate,
    leaveHoursDay: leaveHoursDay,
    employeeReason: employeeReason,
    employeeAttachments: employeeAttachments,
    supervisorEmail: supervisorEmail,
    acNo: acNo,
    requestId: requestId,
  };

  // action link - will call on the createApprovalUrl function
  const approvalUrl = createApprovalUrl(requestId, "approve");
  const rejectionUrl = createApprovalUrl(requestId, "reject");

  const formattedStartDate = formatDate(leaveStartDate);
  const formattedEndDate = formatDate(leaveEndDate);

  // Log values for debugging
  debugLog(`Employee Name: ${employeeName}`);
  debugLog(`Employee Email: ${employeeEmail}`);
  debugLog(`Supervisor Email: ${supervisorEmail}`);

  // Validate supervisor's email
  if (!supervisorEmail || !supervisorEmail.includes("@")) {
    Logger.log("Error: Invalid Supervisor Email - " + supervisorEmail);
    return;
  }

  if (!employeeEmail || !leaveHoursDay) {
    Logger.log("Error: Missing employee email or leave input.");
    return;
  }

  // Convert leave input into numeric value
  let parsedLeaveHoursDay = parseLeaveHours(leaveHoursDay);

  // ============================================
  // Check if employee email exists in "Leave Balances" sheet
  // ============================================
  const emailExists = balancesData.some(
    (row) =>
      row[A_BALANCE_EMAIL] === employeeEmail ||
      ALLOWED_EMAILS.includes(employeeEmail)
  ); // Column B - Employee Email

  if (!emailExists) {
    debugLog(
      `Leave request automatically rejected for ${employeeName} (${employeeEmail}) due to email mismatch.`
    );

    // Update "Status" column in "All Records"
    sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
    sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

    copyToAutomaticallyRejectedLeaves(
      requestLeaveData,
      "Invalid Email Address"
    );

    // Send rejection email
    const templateData = {
      receiverName: employeeFormatName,
      bodyMessage: `
      <p>The email you provided (${employeeEmail}) does not match any records in the system.</p>
      <p>Please <strong>submit a new request</strong> and use your company email (all lower case).</p>
      <p>Contact HR for assistance.</p>
      <p>Thank you.</p>`,
    };

    try {
      sendTemplatedEmail(
        TEST_DEV_ACC,
        "Leave Request Automatically Rejected - Invalid Email Address",
        "TemplateGeneralEmailBody",
        templateData,
        [],
        true
      );
    } catch (error) {
      Logger.log("Failed to send rejection email: " + error.message);
    }
    return;
  }

  const leaveStartDateObj = new Date(parseSheetDate(leaveStartDate));
  const leaveEndDateObj = new Date(parseSheetDate(leaveEndDate));

  const allowedStart = PERIOD_ALLOWED_START_DATE;
  const allowedEnd = PERIOD_ALLOWED_END_DATE;

  debugLog(`Final Parsed Leave Start Date: ${leaveStartDateObj}`);
  debugLog(`Final Parsed Leave End Date: ${leaveEndDateObj}`);
  debugLog(`Allowed Start: ${allowedStart}`);
  debugLog(`Allowed End: ${allowedEnd}`);

  // ============================================
  // Check if the leave request is outside the allowed range
  // ============================================
  if (
    !isExemptedEmail(employeeEmail) &&
    !isParaplannerEmail(employeeEmail) &&
    !isOtherExemptedEmail(employeeEmail) && // Apply rejection to non-exempted employees
    (leaveStartDateObj < allowedStart || // Reject if before Dec 26 {previous year}
      leaveStartDateObj > allowedEnd || // Reject if after Dec 25 {current year}
      (leaveStartDateObj <= allowedEnd && leaveEndDateObj > allowedEnd))
  ) {
    debugLog(
      `Leave request automatically rejected for ${employeeFormatName} (Start: ${leaveStartDateObj.toDateString()}, End: ${leaveEndDateObj.toDateString()})`
    );

    sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
    sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

    //Copy to Automatically Rejected Leaves with Reason
    copyToAutomaticallyRejectedLeaves(
      requestLeaveData,
      "Outside Allowed Leave Period"
    );

    // Send rejection email to employee
    const templateData = {
      receiverName: employeeFormatName,
      bodyMessage: `
      <p>Your leave request has been <strong>automatically rejected</strong> as the selected leave does not fall within the leave period of the current year.</p>
      <div class="info-card">
          <div class="info-card-title">Leave Details</div>
          <div class="info-card-body">
              <table class="info-table" cellpadding="0" cellspacing="0">
                  <tr>
                      <td class="info-label">Leave Period:</td>
                      <td class="info-value">${formatDate(
                        allowedStart
                      )} to ${formatDate(allowedEnd)}</td>
                  </tr>
                  <tr class="last-row">
                      <td class="info-label">Requested Leave:</td>
                      <td class="info-value">${formatDate(
                        leaveStartDateObj
                      )} to ${formatDate(leaveEndDateObj)}</td>
                  </tr>
              </table>
          </div>
      </div>
      <p style="margin-top: 20px;">If you wish to file leave outside of this period, kindly contact your supervisor to check if there are any forms available.</p>
      <p>Thank you.</p>
      `,
    };

    if (employeeEmail) {
      try {
        sendTemplatedEmail(
          TEST_DEV_ACC,
          "Leave Request Automatically Rejected - Outside Allowed Leave Period",
          "TemplateGeneralEmailBody",
          templateData,
          [],
          true
        );
      } catch (error) {
        Logger.log("Failed to send rejection email: " + error.message);
      }
    } else {
      Logger.log("No email found for employee, skipping email send.");
    }
    return; // Stop further execution
  }

  // ============================================
  // Only check balance for certain leave types
  // ============================================
  const balanceCheckedLeaves = MAIN_LEAVE_BALANCE_CHECK;

  if (balanceCheckedLeaves.includes(leaveType)) {
    let remainingSL = null;
    let remainingVL = null;

    // Define half-year periods for RBA accounts
    const currentYear = new Date().getFullYear();
    const janToJuneStart = new Date(currentYear, 0, 1);
    const janToJuneEnd = new Date(currentYear, 5, 30);
    const julyToDecStart = new Date(currentYear, 6, 1);
    const julyToDecEnd = new Date(currentYear, 11, 31);

    const isJanToJune =
      leaveStartDate >= janToJuneStart && leaveEndDate <= janToJuneEnd;
    const isJulyToDec =
      leaveStartDate >= julyToDecStart && leaveStartDate <= julyToDecEnd;

    if (isExemptedEmail(employeeEmail)) {
      // Use "RBA Leave Balances" sheet for exempted emails
      const rbaBalancesSheet = ACTIVE_SHEET.getSheetByName(RBA_LEAVE_BALANCES);
      const rbaBalancesData = rbaBalancesSheet.getDataRange().getValues();

      // Find the employee in the RBA Leave Balances sheet
      for (let i = 1; i < rbaBalancesData.length; i++) {
        // Skip header row
        if (rbaBalancesData[i][A_BALANCE_EMAIL] === employeeEmail) {
          if (isJanToJune) {
            if (leaveType === "Vacation Leave" && subLeaveType === "VL/SL") {
              remainingVL = rbaBalancesData[i][A_JAN_JUN_REMAINING_VL];
            }
            if (leaveType === "Sick Leave" && subLeaveType === "VL/SL") {
              remainingSL = rbaBalancesData[i][A_JAN_JUN_REMAINING_SL];
            }
            // For Emergency Leave and Unexcused
            if (
              (leaveType === "Sick Leave" &&
                subLeaveType === "Emergency Leave") ||
              (leaveType === "Vacation Leave" &&
                subLeaveType === "Emergency Leave") ||
              (leaveType === "Vacation Leave" &&
                subLeaveType === "Unexcused") ||
              (leaveType === "Sick Leave" && subLeaveType === "Unexcused")
            ) {
              remainingVL = rbaBalancesData[i][A_JAN_JUN_REMAINING_VL];
              remainingSL = rbaBalancesData[i][A_JAN_JUN_REMAINING_SL];
            }
            // For "Other" leave type with Emergency/Unexcused
            if (
              leaveType === "Other" &&
              (subLeaveType === "Emergency Leave" ||
                subLeaveType === "Unexcused")
            ) {
              remainingVL = rbaBalancesData[i][A_JAN_JUN_REMAINING_VL];
              remainingSL = rbaBalancesData[i][A_JAN_JUN_REMAINING_SL];
            }
          } else {
            if (leaveType === "Vacation Leave" && subLeaveType === "VL/SL") {
              remainingVL = rbaBalancesData[i][A_JUL_DEC_REMAINING_VL];
            }
            if (leaveType === "Sick Leave" && subLeaveType === "VL/SL") {
              remainingSL = rbaBalancesData[i][A_JUL_DEC_REMAINING_SL];
            }
            // For Emergency Leave and Unexcused
            if (
              (leaveType === "Sick Leave" &&
                subLeaveType === "Emergency Leave") ||
              (leaveType === "Vacation Leave" &&
                subLeaveType === "Emergency Leave") ||
              (leaveType === "Vacation Leave" &&
                subLeaveType === "Unexcused") ||
              (leaveType === "Sick Leave" && subLeaveType === "Unexcused")
            ) {
              remainingVL = rbaBalancesData[i][A_JUL_DEC_REMAINING_VL];
              remainingSL = rbaBalancesData[i][A_JUL_DEC_REMAINING_SL];
            }
            // For "Other" leave type with Emergency/Unexcused
            if (
              leaveType === "Other" &&
              (subLeaveType === "Emergency Leave" ||
                subLeaveType === "Unexcused")
            ) {
              remainingVL = rbaBalancesData[i][A_JUL_DEC_REMAINING_VL];
              remainingSL = rbaBalancesData[i][A_JUL_DEC_REMAINING_SL];
            }
          }
        }
      }
    } else {
      // Use regular "Leave Balances" sheet for non-exempted emails
      for (let i = 1; i < balancesData.length; i++) {
        // Skip header row
        if (balancesData[i][A_BALANCE_EMAIL] === employeeEmail) {
          if (leaveType === "Vacation Leave" && subLeaveType === "VL/SL") {
            remainingVL = balancesData[i][A_REMAINING_VL];
          }
          if (leaveType === "Sick Leave" && subLeaveType === "VL/SL") {
            remainingSL = balancesData[i][A_REMAINING_SL];
          }
          // For Emergency Leave and Unexcused
          if (
            (leaveType === "Sick Leave" &&
              subLeaveType === "Emergency Leave") ||
            (leaveType === "Vacation Leave" &&
              subLeaveType === "Emergency Leave") ||
            (leaveType === "Vacation Leave" && subLeaveType === "Unexcused") ||
            (leaveType === "Sick Leave" && subLeaveType === "Unexcused")
          ) {
            remainingVL = balancesData[i][A_REMAINING_VL];
            remainingSL = balancesData[i][A_REMAINING_SL];
          }
          // For "Other" leave type with Emergency/Unexcused
          if (
            leaveType === "Other" &&
            (subLeaveType === "Emergency Leave" || subLeaveType === "Unexcused")
          ) {
            remainingVL = balancesData[i][A_REMAINING_VL];
            remainingSL = balancesData[i][A_REMAINING_SL];
          }
          break;
        }
      }
    }

    const isRBA = isExemptedEmail(employeeEmail);

    // Generate balance messages
    let balanceMessageEmergency = null;
    let balanceMessageUnexcused = null;

    if (leaveType === "Sick Leave" || leaveType === "Vacation Leave") {
      balanceMessageEmergency = getBalanceMessage(
        leaveType,
        "Emergency Leave",
        remainingSL === null && remainingVL === null
          ? null
          : `Sick Leave: ${remainingSL}<br>Vacation Leave: ${remainingVL}`,
        isRBA,
        isJanToJune,
        isJulyToDec
      );
      balanceMessageUnexcused = getBalanceMessage(
        leaveType,
        "Unexcused",
        remainingSL === null && remainingVL === null
          ? null
          : `Vacation Leave: ${remainingVL}<br>Sick Leave: ${remainingSL}`,
        isRBA,
        isJanToJune,
        isJulyToDec
      );
    }
    let balanceMessageSick = getBalanceMessage(
      "Sick Leave",
      "VL/SL",
      remainingSL,
      isRBA,
      isJanToJune,
      isJulyToDec
    );
    let balanceMessageVacation = getBalanceMessage(
      "Vacation Leave",
      "VL/SL",
      remainingVL,
      isRBA,
      isJanToJune,
      isJulyToDec
    );

    let balanceMessage = "";

    if (leaveType === "Vacation Leave" && subLeaveType === "VL/SL") {
      balanceMessage = balanceMessageVacation;
    } else if (leaveType === "Sick Leave" && subLeaveType === "VL/SL") {
      balanceMessage = balanceMessageSick;
    } else if (
      (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
      subLeaveType === "Emergency Leave"
    ) {
      balanceMessage = balanceMessageEmergency;
    } else if (
      (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
      subLeaveType === "Unexcused"
    ) {
      balanceMessage = balanceMessageUnexcused;
    }

    let insufficientBalance = false;

    // ============================================
    // Additional validation for Unexcused Leave
    // ============================================
    if (subLeaveType === "Unexcused") {
      // If "Other" is selected but employee has available VL or SL balance
      if (
        leaveType === "Other" &&
        ((remainingVL !== null && remainingVL > 0) ||
          (remainingSL !== null && remainingSL > 0))
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining Balance:</td>
            <td class="info-value">VL: ${remainingVL || 0}<br>SL: ${
          remainingSL || 0
        }</td>
        </tr>`;
      }

      // If "Sick Leave" is selected but VL balance is still available
      if (
        leaveType === "Sick Leave" &&
        remainingVL !== null &&
        remainingVL > 0
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining VL Balance:</td>
            <td class="info-value">${remainingVL}</td>
        </tr>`;
      }

      //If "Vacation Leave" is selected but VL balance is 0, should use SL
      if (
        leaveType === "Vacation Leave" &&
        (remainingVL === null || remainingVL === 0) &&
        remainingSL !== null &&
        remainingSL > 0
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining Balance:</td>
            <td class="info-value">VL: 0<br>SL: ${remainingSL}</td>
        </tr>`;
      }

      // If both VL and SL are 0
      if (
        (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
        (remainingVL === null || remainingVL === 0) &&
        (remainingSL === null || remainingSL === 0)
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining Balance:</td>
            <td class="info-value">VL: 0<br>SL: 0</td>
        </tr>`;
      }
    }

    // ============================================
    // Additional validation for Emergency Leave
    // ============================================
    if (subLeaveType === "Emergency Leave") {
      // If "Other" is selected but employee has available VL or SL balance
      if (
        leaveType === "Other" &&
        ((remainingSL !== null && remainingSL > 0) ||
          (remainingVL !== null && remainingVL > 0))
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining Balance:</td>
            <td class="info-value">SL: ${remainingSL || 0}<br>VL: ${
          remainingVL || 0
        }</td>
        </tr>`;
      }

      // If "Vacation Leave" is selected but SL balance is still available
      if (
        leaveType === "Vacation Leave" &&
        remainingSL !== null &&
        remainingSL > 0
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining SL Balance:</td>
            <td class="info-value">${remainingSL}</td>
        </tr>`;
      }

      // If "Sick Leave" is selected but SL balance is 0, should use VL
      if (
        leaveType === "Sick Leave" &&
        (remainingSL === null || remainingSL === 0) &&
        remainingVL !== null &&
        remainingVL > 0
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining Balance:</td>
            <td class="info-value">SL: 0<br>VL: ${remainingVL}</td>
        </tr>`;
      }

      // If both SL and VL are 0
      if (
        (leaveType === "Sick Leave" || leaveType === "Vacation Leave") &&
        (remainingSL === null || remainingSL === 0) &&
        (remainingVL === null || remainingVL === 0)
      ) {
        insufficientBalance = true;
        balanceMessage = `
        <tr class="last-row">
            <td class="info-label">Remaining Balance:</td>
            <td class="info-value">SL: 0<br>VL: 0</td>
        </tr>`;
      }
    }

    if (
      leaveType === "Sick Leave" &&
      subLeaveType === "VL/SL" &&
      (remainingSL === null || parsedLeaveHoursDay > remainingSL)
    ) {
      insufficientBalance = true;
    } else if (
      leaveType === "Vacation Leave" &&
      subLeaveType === "VL/SL" &&
      (remainingVL === null || parsedLeaveHoursDay > remainingVL)
    ) {
      insufficientBalance = true;
    } else if (
      (leaveType === "Vacation Leave" && subLeaveType === "Emergency Leave") ||
      (leaveType === "Sick Leave" && subLeaveType === "Emergency Leave")
    ) {
      const totalEL = (remainingSL || 0) + (remainingVL || 0);
      if (parsedLeaveHoursDay > totalEL) {
        insufficientBalance = true;
      }
    } else if (
      (leaveType === "Vacation Leave" && subLeaveType === "Unexcused") ||
      (leaveType === "Sick Leave" && subLeaveType === "Unexcused")
    ) {
      const totalUnexcused = (remainingVL || 0) + (remainingSL || 0);
      if (parsedLeaveHoursDay > totalUnexcused) {
        insufficientBalance = true;
      }
    }

    // ============================================
    // Check if leave request exceeds available balance
    // ============================================
    if (insufficientBalance) {
      debugLog(
        `Leave request automatically rejected for ${employeeFormatName} due to insufficient balance.`
      );

      // Update Status column
      sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
      sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

      let balanceRejectionReason = "Insufficient Leave Balance";
      if (
        (subLeaveType === "Unexcused" || subLeaveType === "Emergency Leave") &&
        (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
        (remainingVL === null || remainingVL === 0) &&
        (remainingSL === null || remainingSL === 0)
      ) {
        balanceRejectionReason = "No Remaining Leave Balance";
      }
      // Used "Other" but has available balance
      else if (
        (subLeaveType === "Unexcused" && leaveType === "Other") ||
        (subLeaveType === "Emergency Leave" && leaveType === "Other")
      ) {
        balanceRejectionReason = "Incorrect Leave Type Selection";
      }
      // Wrong main leave type (has primary balance but used secondary)
      else if (
        (subLeaveType === "Unexcused" && leaveType === "Sick Leave") ||
        (subLeaveType === "Emergency Leave" && leaveType === "Vacation Leave")
      ) {
        balanceRejectionReason = "Wrong Main Leave Type";
      }
      // Primary balance is 0 but secondary has balance
      else if (
        (subLeaveType === "Unexcused" &&
          leaveType === "Vacation Leave" &&
          (remainingVL === null || remainingVL === 0)) ||
        (subLeaveType === "Emergency Leave" &&
          leaveType === "Sick Leave" &&
          (remainingSL === null || remainingSL === 0))
      ) {
        balanceRejectionReason = "Wrong Main Leave Type";
      }

      copyToAutomaticallyRejectedLeaves(
        requestLeaveData,
        balanceRejectionReason
      );

      // Send Rejection Email
      const templateData = {
        receiverName: employeeFormatName,
        bodyMessage: `
        <p>Your leave request has been <strong>automatically rejected</strong> ${
          // CHECK BOTH ZERO FIRST
          subLeaveType === "Unexcused" &&
          (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
          (remainingVL === null || remainingVL === 0) &&
          (remainingSL === null || remainingSL === 0)
            ? "because you have <strong>no remaining VL and SL balance</strong>."
            : subLeaveType === "Emergency Leave" &&
              (leaveType === "Sick Leave" || leaveType === "Vacation Leave") &&
              (remainingSL === null || remainingSL === 0) &&
              (remainingVL === null || remainingVL === 0)
            ? "because you have <strong>no remaining SL and VL balance</strong>."
            : (subLeaveType === "Unexcused" && leaveType === "Other") ||
              (subLeaveType === "Emergency Leave" && leaveType === "Other")
            ? "because you still have available leave balance. Please use your VL/SL first before filing under 'Other'."
            : subLeaveType === "Unexcused" && leaveType === "Sick Leave"
            ? "because you selected <strong>SL</strong> but still have available <strong>VL</strong> balance."
            : subLeaveType === "Emergency Leave" &&
              leaveType === "Vacation Leave"
            ? "because you selected <strong>VL</strong> but still have available <strong>SL</strong> balance."
            : subLeaveType === "Unexcused" &&
              leaveType === "Vacation Leave" &&
              (remainingVL === null || remainingVL === 0)
            ? "because your <strong>VL balance is 0</strong>."
            : subLeaveType === "Emergency Leave" &&
              leaveType === "Sick Leave" &&
              (remainingSL === null || remainingSL === 0)
            ? "because your <strong>SL balance is 0</strong>."
            : "due to insufficient remaining leave balance."
        }</p>
        <div class="info-card">
            <div class="info-card-title">Leave Details</div>
            <div class="info-card-body">
                <table class="info-table" cellpadding="0" cellspacing="0">
                    <tr>
                        <td class="info-label">Main Leave Type:</td>
                        <td class="info-value">${leaveType}</td>
                    </tr>
                    <tr>
                        <td class="info-label">Leave Sub-Category:</td>
                        <td class="info-value">${subLeaveType}</td>
                    </tr>
                    <tr>
                        <td class="info-label">Leave Duration Request:</td>
                        <td class="info-value">${leaveHoursDay}</td>
                    </tr>
                    ${balanceMessage}
                    ${
                      // BOTH ZERO - NO CARD NOTE NEEDED, just show the balance
                      subLeaveType === "Unexcused" &&
                      (leaveType === "Vacation Leave" ||
                        leaveType === "Sick Leave") &&
                      (remainingVL === null || remainingVL === 0) &&
                      (remainingSL === null || remainingSL === 0)
                        ? ``
                        : subLeaveType === "Emergency Leave" &&
                          (leaveType === "Sick Leave" ||
                            leaveType === "Vacation Leave") &&
                          (remainingSL === null || remainingSL === 0) &&
                          (remainingVL === null || remainingVL === 0)
                        ? ``
                        : subLeaveType === "Unexcused" &&
                          (leaveType === "Sick Leave" ||
                            leaveType === "Vacation Leave")
                        ? `
                        <tr class="card-note">
                          <td colspan="2">
                            <p>Unexcused leave deducts from your VL first. If there's no remaining balance for VL, then it will be deducted from your SL.</p>
                          </td>
                        </tr>`
                        : subLeaveType === "Emergency Leave" &&
                          (leaveType === "Vacation Leave" ||
                            leaveType === "Sick Leave")
                        ? `
                        <tr class="card-note">
                          <td colspan="2">
                            <p>Emergency leave deducts from your SL first. If there's no remaining balance for SL, then it will be deducted from your VL.</p>
                          </td>
                        </tr>`
                        : ""
                    }
                </table>
            </div>
        </div>
        ${
          // BOTH ZERO - DIRECT TO USE "OTHER"
          subLeaveType === "Unexcused" &&
          (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
          (remainingVL === null || remainingVL === 0) &&
          (remainingSL === null || remainingSL === 0)
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Select <strong>Other</strong> as Main Leave Type</li>
                <li>Keep <strong>Unexcused</strong> as Sub-Category</li>
              </ul>
            </div>`
            : subLeaveType === "Emergency Leave" &&
              (leaveType === "Sick Leave" || leaveType === "Vacation Leave") &&
              (remainingSL === null || remainingSL === 0) &&
              (remainingVL === null || remainingVL === 0)
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Select <strong>Other</strong> as Main Leave Type</li>
                <li>Keep <strong>Emergency Leave</strong> as Sub-Category</li>
              </ul>
            </div>`
            : // CHECK ZERO BALANCE CASES FIRST
            subLeaveType === "Emergency Leave" &&
              leaveType === "Other" &&
              (remainingSL === null || remainingSL === 0) &&
              remainingVL > 0
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Select <strong>Vacation Leave</strong> as Main Leave Type</li>
                <li>Keep <strong>Emergency Leave</strong> as Sub-Category</li>
              </ul>
              <br>
              <p style="margin-top: 10px;"><strong>Note:</strong> Since your SL balance is 0, Emergency Leave will use your VL balance.</p>
            </div>`
            : subLeaveType === "Unexcused" &&
              leaveType === "Other" &&
              (remainingVL === null || remainingVL === 0) &&
              remainingSL > 0
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Select <strong>Sick Leave</strong> as Main Leave Type</li>
                <li>Keep <strong>Unexcused</strong> as Sub-Category</li>
              </ul>
              <br>
              <p style="margin-top: 10px;"><strong>Note:</strong> Since your VL balance is 0, Unexcused Leave will use your SL balance.</p>
            </div>`
            : // THEN DEFAULT CASES (when primary balance exists)
            subLeaveType === "Emergency Leave" && leaveType === "Other"
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Select <strong>Sick Leave</strong> as Main Leave Type</li>
                <li>Keep <strong>Emergency Leave</strong> as Sub-Category</li>
              </ul>
              <br>
              <p style="margin-top: 10px;"><strong>Note:</strong> Only use "Other" when you have <strong>no</strong> VL and SL balance remaining.</p>
            </div>`
            : subLeaveType === "Unexcused" && leaveType === "Other"
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Select <strong>Vacation Leave</strong> as Main Leave Type</li>
                <li>Keep <strong>Unexcused</strong> as Sub-Category</li>
              </ul>
              <br>
              <p style="margin-top: 10px;"><strong>Note:</strong> Only use "Other" when you have <strong>no</strong> VL and SL balance remaining.</p>
            </div>`
            : subLeaveType === "Unexcused" && leaveType === "Sick Leave"
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Change Main Leave Type to <strong>Vacation Leave</strong></li>
                <li>Keep <strong>Unexcused</strong> as Sub-Category</li>
              </ul>
            </div>`
            : subLeaveType === "Emergency Leave" &&
              leaveType === "Vacation Leave"
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Change Main Leave Type to <strong>Sick Leave</strong></li>
                <li>Keep <strong>Emergency Leave</strong> as Sub-Category</li>
              </ul>
            </div>`
            : subLeaveType === "Unexcused" &&
              leaveType === "Vacation Leave" &&
              (remainingVL === null || remainingVL === 0)
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Change Main Leave Type to <strong>Sick Leave</strong></li>
                <li>Keep <strong>Unexcused</strong> as Sub-Category</li>
              </ul>
            </div>`
            : subLeaveType === "Emergency Leave" &&
              leaveType === "Sick Leave" &&
              (remainingSL === null || remainingSL === 0)
            ? `
            <div class="external-note">
              <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
              <ul style="margin-top: 0;">
                <li>Change Main Leave Type to <strong>Vacation Leave</strong></li>
                <li>Keep <strong>Emergency Leave</strong> as Sub-Category</li>
              </ul>
            </div>`
            : ""
        }
        <p style="margin-top: 20px;">Please resubmit your request with the correct leave type. Thank you.</p>
        `,
      };

      let emailSubject = "Leave Request Automatically Rejected - ";

      // Both VL and SL are 0
      if (
        (subLeaveType === "Unexcused" || subLeaveType === "Emergency Leave") &&
        (leaveType === "Vacation Leave" || leaveType === "Sick Leave") &&
        (remainingVL === null || remainingVL === 0) &&
        (remainingSL === null || remainingSL === 0)
      ) {
        emailSubject += "No Remaining Leave Balance";
      }
      // Used "Other" but has available balance
      else if (
        (subLeaveType === "Unexcused" && leaveType === "Other") ||
        (subLeaveType === "Emergency Leave" && leaveType === "Other")
      ) {
        emailSubject += "Incorrect Leave Type Selection";
      }
      // Wrong main leave type (has primary balance but used secondary)
      else if (
        (subLeaveType === "Unexcused" && leaveType === "Sick Leave") ||
        (subLeaveType === "Emergency Leave" && leaveType === "Vacation Leave")
      ) {
        emailSubject += "Wrong Main Leave Type";
      }
      // Primary balance is 0 but secondary has balance
      else if (
        (subLeaveType === "Unexcused" &&
          leaveType === "Vacation Leave" &&
          (remainingVL === null || remainingVL === 0)) ||
        (subLeaveType === "Emergency Leave" &&
          leaveType === "Sick Leave" &&
          (remainingSL === null || remainingSL === 0))
      ) {
        emailSubject += "Wrong Main Leave Type";
      }
      // Default - insufficient balance for VL/SL requests
      else {
        emailSubject += "Insufficient Leave Balance";
      }

      if (employeeEmail) {
        try {
          sendTemplatedEmail(
            TEST_DEV_ACC,
            emailSubject,
            "TemplateGeneralEmailBody",
            templateData,
            [],
            true
          );
        } catch (error) {
          Logger.log("Failed to send rejection email: " + error.message);
        }
      }
      return; // Stop further execution
    }
  }

  // ============================================
  // AUTOMATIC REJECTION: Other Leaves with wrong main leave type
  // ============================================
  const otherLeaveTypes = OTHER_LEAVE_TYPES_WITH_BALANCE;

  if (otherLeaveTypes.includes(subLeaveType)) {
    // Check if main leave type is NOT "Other"
    if (leaveType !== "Other") {
      debugLog(
        `Leave request automatically rejected for ${employeeFormatName} - ${subLeaveType} must use "Other" as main leave type.`
      );

      sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
      sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

      copyToAutomaticallyRejectedLeaves(
        requestLeaveData,
        "Wrong Main Leave Type"
      );

      // Send rejection email
      const templateData = {
        receiverName: employeeFormatName,
        bodyMessage: `
        <p>Your leave request has been <strong>automatically rejected</strong> because you selected <strong>${subLeaveType}</strong> under <strong>${leaveType}</strong>.</p>
        <div class="info-card">
            <div class="info-card-title">Leave Details</div>
            <div class="info-card-body">
                <table class="info-table" cellpadding="0" cellspacing="0">
                    <tr>
                        <td class="info-label">Main Leave Type:</td>
                        <td class="info-value">${leaveType}</td>
                    </tr>
                    <tr class="last-row">
                        <td class="info-label">Leave Sub-Category:</td>
                        <td class="info-value">${subLeaveType}</td>
                    </tr>
                    <tr class="card-note">
                      <td colspan="2">
                        <p>${subLeaveType} must always be filed under the "Other" main leave type.</p>
                      </td>
                    </tr>
                </table>
            </div>
        </div>
        <div class="external-note">
          <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
          <ul style="margin-top: 0;">
            <li>Select <strong>Other</strong> as Main Leave Type</li>
            <li>Keep <strong>${subLeaveType}</strong> as Sub-Category</li>
          </ul>
        </div>
        <p>Please resubmit your request with the correct leave type. Thank you.</p>
        `,
      };

      try {
        sendTemplatedEmail(
          TEST_DEV_ACC,
          "Leave Request Automatically Rejected - Wrong Main Leave Type",
          "TemplateGeneralEmailBody",
          templateData,
          [],
          true
        );
      } catch (error) {
        Logger.log("Failed to send rejection email: " + error.message);
      }
      return; // Stop further execution
    }

    // Check balance for Other Leaves (Bereavement, Parental, Maternal)
    const otherLeaveBalancesSheet =
      ACTIVE_SHEET.getSheetByName(OTHER_LEAVE_BALANCES);
    if (!otherLeaveBalancesSheet) {
      Logger.log("Error: 'Other Leave Balances' sheet not found.");
      return;
    }

    const otherBalancesData = otherLeaveBalancesSheet
      .getDataRange()
      .getValues();
    let remainingBalance = null;
    let leaveTypeName = "";

    // Find the employee's balance
    for (let i = 1; i < otherBalancesData.length; i++) {
      if (otherBalancesData[i][A_BALANCE_EMAIL] === employeeEmail) {
        if (subLeaveType === "Bereavement Leave") {
          remainingBalance = otherBalancesData[i][A_REMAINING_BL];
          leaveTypeName = "Bereavement Leave";
        } else if (subLeaveType === "Paternity Leave") {
          remainingBalance = otherBalancesData[i][A_REMAINING_PL];
          leaveTypeName = "Paternity Leave";
        } else if (subLeaveType === "Maternity Leave") {
          remainingBalance = otherBalancesData[i][A_REMAINING_ML];
          leaveTypeName = "Maternity Leave";
        }
        break;
      }
    }

    // Check if balance is insufficient
    if (remainingBalance === null || parsedLeaveHoursDay > remainingBalance) {
      debugLog(
        `Leave request automatically rejected for ${employeeFormatName} due to insufficient ${leaveTypeName} balance.`
      );

      sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
      sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

      copyToAutomaticallyRejectedLeaves(
        requestLeaveData,
        `Insufficient ${leaveTypeName} Balance`
      );

      // Send rejection email
      const templateData = {
        receiverName: employeeFormatName,
        bodyMessage: `
        <p>Your leave request has been <strong>automatically rejected</strong> due to insufficient ${leaveTypeName} balance.</p>
        <div class="info-card">
            <div class="info-card-title">Leave Details</div>
            <div class="info-card-body">
                <table class="info-table" cellpadding="0" cellspacing="0">
                    <tr>
                        <td class="info-label">Main Leave Type:</td>
                        <td class="info-value">${leaveType}</td>
                    </tr>
                    <tr>
                        <td class="info-label">Leave Sub-Category:</td>
                        <td class="info-value">${subLeaveType}</td>
                    </tr>
                    <tr>
                        <td class="info-label">Leave Duration Request:</td>
                        <td class="info-value">${leaveHoursDay}</td>
                    </tr>
                    <tr class="last-row">
                        <td class="info-label">Remaining Leave Balance:</td>
                        <td class="info-value">${
                          remainingBalance !== null ? remainingBalance : 0
                        }</td>
                    </tr>
                </table>
            </div>
        </div>
        <p style="margin-top: 20px;">Please check your leave balance and resubmit with the correct duration. Thank you.</p>
        `,
      };

      try {
        sendTemplatedEmail(
          TEST_DEV_ACC,
          `Leave Request Automatically Rejected - Insufficient ${leaveTypeName} Balance`,
          "TemplateGeneralEmailBody",
          templateData,
          [],
          true
        );
      } catch (error) {
        Logger.log("Failed to send rejection email: " + error.message);
      }
      return; // Stop further execution
    }
  }

  // ============================================
  // AUTOMATIC REJECTION: Wrong Main Leave Type for Other Leaves (WITHOUT balance checking)
  // ============================================
  const otherLeaveTypesWithoutBalance = OTHER_LEAVE_TYPES_WITHOUT_BALANCE;

  if (otherLeaveTypesWithoutBalance.includes(subLeaveType)) {
    // Check if main leave type is NOT "Other"
    if (leaveType !== "Other") {
      debugLog(
        `Leave request automatically rejected for ${employeeFormatName} - ${subLeaveType} must use "Other" as main leave type.`
      );

      sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
      sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

      copyToAutomaticallyRejectedLeaves(
        requestLeaveData,
        "Wrong Main Leave Type"
      );

      // Send rejection email
      const templateData = {
        receiverName: employeeFormatName,
        bodyMessage: `
        <p>Your leave request has been <strong>automatically rejected</strong> because you selected <strong>${subLeaveType}</strong> under <strong>${leaveType}</strong>.</p>
        <div class="info-card">
            <div class="info-card-title">Leave Details</div>
            <div class="info-card-body">
                <table class="info-table" cellpadding="0" cellspacing="0">
                    <tr>
                        <td class="info-label">Main Leave Type:</td>
                        <td class="info-value">${leaveType}</td>
                    </tr>
                    <tr class="last-row">
                        <td class="info-label">Leave Sub-Category:</td>
                        <td class="info-value">${subLeaveType}</td>
                    </tr>
                    <tr class="card-note">
                      <td colspan="2">
                        <p>${subLeaveType} must always be filed under the "Other" main leave type.</p>
                      </td>
                    </tr>
                </table>
            </div>
        </div>
        <div class="external-note">
          <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
          <ul style="margin-top: 0;">
            <li>Select <strong>Other</strong> as Main Leave Type</li>
            <li>Keep <strong>${subLeaveType}</strong> as Sub-Category</li>
          </ul>
        </div>
        <p>Please resubmit your request with the correct leave type. Thank you.</p>
        `,
      };

      try {
        sendTemplatedEmail(
          TEST_DEV_ACC,
          "Leave Request Automatically Rejected - Wrong Main Leave Type",
          "TemplateGeneralEmailBody",
          templateData,
          [],
          true
        );
      } catch (error) {
        Logger.log("Failed to send rejection email: " + error.message);
      }
      return; // Stop further execution
    }

    // No balance checking needed - just continue to approval process
    debugLog(`${subLeaveType} validated - no balance check required.`);
  }

  // ============================================
  // AUTOMATIC REJECTION: VL/SL must use Vacation Leave or Sick Leave as main type
  // ============================================
  if (subLeaveType === "VL/SL") {
    // Check if main leave type is NOT "Vacation Leave" or "Sick Leave"
    if (leaveType !== "Vacation Leave" && leaveType !== "Sick Leave") {
      debugLog(
        `Leave request automatically rejected for ${employeeFormatName} - VL/SL must use "Vacation Leave" or "Sick Leave" as main leave type.`
      );

      sheet.getRange(lastRow, statusColumn).setValue("Automatically Rejected");
      sheet.getRange(lastRow, forwardedStatusColumn).setValue("Not Forwarded");

      copyToAutomaticallyRejectedLeaves(
        requestLeaveData,
        "Wrong Main Leave Type"
      );

      // Send rejection email
      const templateData = {
        receiverName: employeeFormatName,
        bodyMessage: `
        <p>Your leave request has been <strong>automatically rejected</strong> because you selected <strong>VL/SL</strong> under <strong>${leaveType}</strong>.</p>
        <div class="info-card">
            <div class="info-card-title">Leave Details</div>
            <div class="info-card-body">
                <table class="info-table" cellpadding="0" cellspacing="0">
                    <tr>
                        <td class="info-label">Main Leave Type:</td>
                        <td class="info-value">${leaveType}</td>
                    </tr>
                    <tr class="last-row">
                        <td class="info-label">Leave Sub-Category:</td>
                        <td class="info-value">${subLeaveType}</td>
                    </tr>
                    <tr class="card-note">
                      <td colspan="2">
                        <p>When using the VL/SL sub-category, the main leave type must be either 'Vacation Leave' or 'Sick Leave'.</p>
                      </td>
                    </tr>
                </table>
            </div>
        </div>
        <div class="external-note">
          <p style="margin-bottom: 5px;"><strong>Steps to resubmit:</strong></p>
          <ul style="margin-top: 0;">
            <li>Select <strong>Vacation Leave</strong> or <strong>Sick Leave</strong> as Main Leave Type</li>
            <li>Keep <strong>VL/SL</strong> as Sub-Category</li>
          </ul>
        </div>
        <p>Please resubmit your request with the correct leave type. Thank you.</p>
        `,
      };

      try {
        sendTemplatedEmail(
          TEST_DEV_ACC,
          "Leave Request Automatically Rejected - Wrong Main Leave Type",
          "TemplateGeneralEmailBody",
          templateData,
          [],
          true
        );
      } catch (error) {
        Logger.log("Failed to send rejection email: " + error.message);
      }
      return; // Stop further execution
    }
  }

  // ============================================
  // Proceed with approval email if leave request is valid
  // ============================================
  const templateData = {
    bodyMessage:
      "Kindly review the leave request and take the necessary action at your earliest convenience. Please let the employee know if you need any further details. Thank you!",
    receiverName: employeeFormatName,
    employeeEmail: employeeEmail,
    leaveType: leaveType,
    subLeaveType: subLeaveType,
    formattedStartDate: formattedStartDate,
    formattedEndDate: formattedEndDate,
    leaveHoursDay: leaveHoursDay,
    employeeReason: employeeReason,
    approvalUrl: approvalUrl,
    rejectionUrl: rejectionUrl,
  };

  // Handle multiple attachments (Google Drive File ID or URL)
  const attachments = getAttachmentBlobs(employeeAttachments);

  sendTemplatedEmail(
    TEST_DEV_ACC,
    "Leave Request Approval Needed",
    "TemplateApprovalEmailBody",
    templateData,
    attachments,
    true
  );
}

// ============================================
// Get Data
// ============================================
function doGet(e) {
  // Log the entire e object to see its structure
  Logger.log("Received e: " + JSON.stringify(e));

  if (!e || !e.parameter || !e.parameter.id || !e.parameter.action) {
    debugLog("Error: Missing parameters.");
    return HtmlService.createHtmlOutput("Error: Missing parameters.");
  }

  const requestId = e.parameter.id;
  const action = e.parameter.action;

  // Log the parameters for debugging
  debugLog(`Received Request ID: ${requestId}`);
  debugLog(`Received action: ${action}`);

  if (!requestId) {
    return HtmlService.createHtmlOutput("Error: Invalid request ID parameter.");
  }

  const sheet = ACTIVE_SHEET.getSheetByName(ALL_RECORDS);
  if (!sheet) {
    debugLog("Error: Sheet not found.");
    return HtmlService.createHtmlOutput("Error: Sheet not found.");
  }

  // Find the row with matching requestId
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  let row = -1;

  let status = "";
  let existingStatus = null;
  let forwardingStatus = null;
  let message = "";

  for (let i = 1; i < allData.length; i++) {
    // Start from 1 to skip header
    if (allData[i][COL_A_REQUEST_ID] === requestId) {
      row = i + 1; // +1 because array is 0-indexed but sheets are 1-indexed
      break;
    }
  }

  if (row === -1) {
    Logger.log(`Error: Request ID ${requestId} not found.`);
    return HtmlService.createHtmlOutput("Error: Request not found.");
  }

  debugLog(`Found Request ID ${requestId} at row ${row}`);

  // ============================================
  // CHECK FOR DUPLICATE ACTIONS FIRST
  // ============================================
  existingStatus = sheet.getRange(row, COL_N_MAIN_STATUS).getValue();
  forwardingStatus = sheet.getRange(row, COL_N_FORWARDING_STATUS).getValue();

  const isProcessed = existingStatus !== "";
  const alreadyProcessed =
    forwardingStatus === "Forwarded" || forwardingStatus === "Not Forwarded";

  debugLog(
    `Checking Row: ${row}, Status: ${existingStatus}, Forwarding: ${forwardingStatus}`
  );

  // ✅ Prevent duplicate actions
  if (action === "forward" && isProcessed && alreadyProcessed) {
    debugLog(`Leave already forwarded: ${forwardingStatus}`);
    message = "This leave request has already been forwarded.";

    const errorAlertTemplate =
      HtmlService.createTemplateFromFile("TemplateAlertUrl");
    errorAlertTemplate.alertType = "error";
    errorAlertTemplate.header = "<p>Action Not Allowed</p>";
    errorAlertTemplate.message = `<p>${message}</p>`;
    errorAlertTemplate.requestId = requestId;

    return HtmlService.createHtmlOutput(
      errorAlertTemplate.evaluate().getContent()
    );
  }

  if ((action === "approve" || action === "reject") && isProcessed) {
    debugLog(`Leave already processed: ${existingStatus}`);
    message = "This leave request has already been processed.";

    const errorAlertTemplate =
      HtmlService.createTemplateFromFile("TemplateAlertUrl");
    errorAlertTemplate.alertType = "error";
    errorAlertTemplate.header = "<p>Action Not Allowed</p>";
    errorAlertTemplate.message = `<p>${message}</p>`;
    errorAlertTemplate.requestId = requestId;

    return HtmlService.createHtmlOutput(
      errorAlertTemplate.evaluate().getContent()
    );
  }

  // ============================================
  // GET ALL DATA AFTER VALIDATION PASSES
  // ============================================
  const timeStamp = sheet.getRange(row, COL_N_TIMESTAMP).getValue();
  const employeeEmail = sheet.getRange(row, COL_N_EMAIL_ADDRESS).getValue();
  const employeeName = sheet.getRange(row, COL_N_FULL_NAME).getValue();
  let employeeFormatName = formatName(employeeName);

  const jobTitle = sheet.getRange(row, COL_N_JOB_TITLE).getValue();
  const department = sheet.getRange(row, COL_N_DEPARTMENT).getValue();

  const leaveType = sheet.getRange(row, COL_N_MAIN_LEAVE).getValue();
  const subLeaveType = sheet.getRange(row, COL_N_SUB_LEAVE).getValue();
  const leaveStartDate = sheet.getRange(row, COL_N_START_DATE).getValue();
  const leaveEndDate = sheet.getRange(row, COL_N_END_DATE).getValue();
  const noOfDaysOrHours = sheet.getRange(row, COL_N_LEAVE_DURATION).getValue();
  const employeeReason = sheet.getRange(row, COL_N_REASON).getValue();
  const employeeAttachments = sheet.getRange(row, COL_N_ATTACHMENT).getValue();
  const supervisorEmail = sheet.getRange(row, COL_N_APPROVER).getValue();
  const supervisorName = getSupervisorName(supervisorEmail);
  const acNo = sheet.getRange(row, COL_N_AC_NO).getValue();

  const formattedStartDate = formatDate(leaveStartDate);
  const formattedEndDate = formatDate(leaveEndDate);

  // Log the retrieved values for debugging
  debugLog(`Employee email: ${employeeEmail}`);
  debugLog(`Leave start date: ${leaveStartDate}`);
  debugLog(`Leave end date: ${leaveEndDate}`);

  // Handle multiple attachments (Google Drive File ID or URL)
  const attachments = getAttachmentBlobs(employeeAttachments);

  // Prepare data for copying to Categories sheet
  const requestLeaveData = {
    employeeEmail: employeeEmail,
    employeeName: employeeName,
    jobTitle: jobTitle,
    department: department,
    leaveType: leaveType,
    subLeaveType: subLeaveType,
    startDate: leaveStartDate,
    endDate: leaveEndDate,
    leaveHoursDay: noOfDaysOrHours,
    employeeReason: employeeReason,
    employeeAttachments: employeeAttachments,
    supervisorEmail: supervisorEmail,
    acNo: acNo,
    requestId: requestId,
  };

  // ============================================
  // PROCESS ACTIONS
  // ============================================
  if (action === "approve") {
    sheet.getRange(row, COL_N_MAIN_STATUS).setValue("Approved");

    updateLeaveBalances();

    // Determine which leave balance sheet to use based on email
    let leaveBalancesSheet;
    let employeeLeaveBalance = null;
    let vacationUsed = 0;
    let vacationRemaining = 0;
    let sickLeaveUsed = 0;
    let sickLeaveRemaining = 0;

    if (isExemptedEmail(employeeEmail)) {
      // Use RBA Leave Balances for exempted emails
      leaveBalancesSheet = ACTIVE_SHEET.getSheetByName(RBA_LEAVE_BALANCES);
      if (!leaveBalancesSheet) {
        debugLog("ERROR: RBA Leave Balances sheet not found");
        return;
      }

      const lastRow = leaveBalancesSheet.getLastRow();
      if (lastRow < 3) {
        debugLog("ERROR: No data in RBA Leave Balances sheet");
        return;
      }

      // Get data starting from row 3 (skip merged headers)
      const leaveRecords = leaveBalancesSheet
        .getRange(3, 1, lastRow - 2, leaveBalancesSheet.getLastColumn())
        .getValues();

      // Find the employee in the RBA Leave Balances sheet
      for (let i = 0; i < leaveRecords.length; i++) {
        if (leaveRecords[i][A_BALANCE_EMAIL] === employeeEmail) {
          employeeLeaveBalance = leaveRecords[i];
          break;
        }
      }

      // Define half-year periods for RBA accounts
      const currentYear = new Date().getFullYear();
      const janToJuneStart = new Date(currentYear, 0, 1);
      const janToJuneEnd = new Date(currentYear, 5, 30);
      const julyToDecStart = new Date(currentYear, 6, 1);
      const julyToDecEnd = new Date(currentYear, 11, 31);

      const isJanToJune =
        leaveStartDate >= janToJuneStart && leaveEndDate <= janToJuneEnd;
      const isJulyToDec =
        leaveStartDate >= julyToDecStart && leaveStartDate <= julyToDecEnd;

      if (isJanToJune) {
        vacationUsed = employeeLeaveBalance
          ? employeeLeaveBalance[A_JAN_JUN_USED_VL]
          : 0;
        vacationRemaining = employeeLeaveBalance
          ? employeeLeaveBalance[A_JAN_JUN_REMAINING_VL]
          : 0;
        sickLeaveUsed = employeeLeaveBalance
          ? employeeLeaveBalance[A_JAN_JUN_USED_SL]
          : 0;
        sickLeaveRemaining = employeeLeaveBalance
          ? employeeLeaveBalance[A_JAN_JUN_REMAINING_SL]
          : 0;
      } else {
        vacationUsed = employeeLeaveBalance
          ? employeeLeaveBalance[A_JUL_DEC_USED_VL]
          : 0;
        vacationRemaining = employeeLeaveBalance
          ? employeeLeaveBalance[A_JUL_DEC_REMAINING_VL]
          : 0;
        sickLeaveUsed = employeeLeaveBalance
          ? employeeLeaveBalance[A_JUL_DEC_USED_SL]
          : 0;
        sickLeaveRemaining = employeeLeaveBalance
          ? employeeLeaveBalance[A_JUL_DEC_REMAINING_SL]
          : 0;
      }

      const period = isJanToJune
        ? "January to June"
        : isJulyToDec
        ? "July to December"
        : "";
      note = `
      <tr class="card-note">
        <td colspan="4">
          <p>The displayed remaining balance covers the <strong>${period}</strong> period.</p>
        </td>
      </tr>
      `;
    } else {
      // Use regular Leave Balances for non-exempted emails
      leaveBalancesSheet = ACTIVE_SHEET.getSheetByName(LEAVE_BALANCES);
      if (!leaveBalancesSheet) {
        debugLog("ERROR: Leave Balances sheet not found");
        return;
      }

      const lastRow = leaveBalancesSheet.getLastRow();
      if (lastRow < 3) {
        debugLog("ERROR: No data in Leave Balances sheet");
        return;
      }

      // Get data starting from row 3 (skip merged headers)
      const leaveRecords = leaveBalancesSheet
        .getRange(3, 1, lastRow - 2, leaveBalancesSheet.getLastColumn())
        .getValues();

      // Now loop through the data
      for (let i = 0; i < leaveRecords.length; i++) {
        if (leaveRecords[i][A_BALANCE_EMAIL] === employeeEmail) {
          employeeLeaveBalance = leaveRecords[i];
          break;
        }
      }

      // Extract vacation and sick leave balances and used balances
      vacationUsed = employeeLeaveBalance ? employeeLeaveBalance[A_USED_VL] : 0;
      vacationRemaining = employeeLeaveBalance
        ? employeeLeaveBalance[A_REMAINING_VL]
        : 0;
      sickLeaveUsed = employeeLeaveBalance
        ? employeeLeaveBalance[A_USED_SL]
        : 0;
      sickLeaveRemaining = employeeLeaveBalance
        ? employeeLeaveBalance[A_REMAINING_SL]
        : 0;

      note = "";
    }

    // Conditionally generate the leave balance table if leaveType is Vacation Leave or Sick Leave
    let leaveBalanceTable = "";
    if (leaveType === "Vacation Leave" || leaveType === "Sick Leave") {
      leaveBalanceTable = `
        <div class="balance-table-wrapper">
          <!-- Vacation Leave Table -->
          <table class="balance-table balance-table-vl" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;text-align:center;margin-bottom:15px;">
            <thead>
              <tr>
                <th colspan="2" class="vl-header" style="background-color:#4C1130;color:#fff;padding:12px;font-size:16px;font-weight:bold;">Vacation Leave</th>
              </tr>
              <tr>
                <th class="vl-subheader" style="background-color:#D5A6BD;font-size:14px;font-weight:600;padding:12px;width:50%;">Used</th>
                <th class="vl-subheader" style="background-color:#D5A6BD;font-size:14px;font-weight:600;padding:12px;width:50%;">Remaining</th>
              </tr>
            </thead>
            <tbody>
              <tr style="background:#d1d5db30;">
                <td style="padding:15px 10px;font-size:16px;border-top:1px solid #e5e7eb;color:#555;">${vacationUsed}</td>
                <td style="padding:15px 10px;font-size:16px;border-top:1px solid #e5e7eb;color:#555;">${vacationRemaining}</td>
              </tr>
            </tbody>
          </table>
          
          <!-- Sick Leave Table -->
          <table class="balance-table balance-table-sl" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;text-align:center;">
            <thead>
              <tr>
                <th colspan="2" class="sl-header" style="background-color:#073763;color:#fff;padding:12px;font-size:16px;font-weight:bold;">Sick Leave</th>
              </tr>
              <tr>
                <th class="sl-subheader" style="background-color:#9FC5E8;font-size:14px;font-weight:600;padding:12px;width:50%;">Used</th>
                <th class="sl-subheader" style="background-color:#9FC5E8;font-size:14px;font-weight:600;padding:12px;width:50%;">Remaining</th>
              </tr>
            </thead>
            <tbody>
              <tr style="background:#d1d5db30;">
                <td style="padding:15px 10px;font-size:16px;border-top:1px solid #e5e7eb;color:#555;">${sickLeaveUsed}</td>
                <td style="padding:15px 10px;font-size:16px;border-top:1px solid #e5e7eb;color:#555;">${sickLeaveRemaining}</td>
              </tr>
              ${note}
            </tbody>
          </table>
        </div>
      `;

      // Add a note if Emergency Leave is selected
      if (subLeaveType === "Emergency Leave") {
        leaveBalanceTable += `
        <div class="external-note">
          <p style="margin-bottom: 5px;">Note that <strong>emergency leave</strong>: </p>
          <ul style="margin-top: 0;">
          <li>Is deducted from your SL balance first</li>
          <li>If there's no remaining balance for SL, then it will be deducted from your VL</li>
          </ul>
        </div>
        `;
      }

      // Add a note if Unexcused Leave is selected
      if (subLeaveType === "Unexcused") {
        leaveBalanceTable += `
        <div class="external-note">
          <p style="margin-bottom: 5px;">Note that <strong>unexcused leave</strong>: </p>
          <ul style="margin-top: 0;">
          <li>Is deducted from your VL balance first</li>
          <li>If there's no remaining balance for VL, then it will be deducted from your SL</li>
          </ul>
        </div>
        `;
      }
    }

    // For Other leave types - show only the specific leave balance
    if (
      leaveType === "Other" &&
      (subLeaveType === "Bereavement Leave" ||
        subLeaveType === "Paternity Leave" ||
        subLeaveType === "Maternity Leave")
    ) {
      // Get Other Leave Balances sheet
      const otherLeaveBalancesSheet =
        ACTIVE_SHEET.getSheetByName(OTHER_LEAVE_BALANCES);

      if (!otherLeaveBalancesSheet) {
        debugLog("ERROR: Other Leave Balances sheet not found");
        return;
      }

      const lastRow = otherLeaveBalancesSheet.getLastRow();
      if (lastRow < 3) {
        debugLog("ERROR: No data in Other Leave Balances sheet");
        return;
      }

      const otherLeaveRecords = otherLeaveBalancesSheet
        .getRange(3, 1, lastRow - 2, otherLeaveBalancesSheet.getLastColumn())
        .getValues();

      let leaveUsed = 0;
      let leaveRemaining = 0;
      let leaveTitle = "";
      let headerColor = "";
      let subHeaderColor = "";

      // Find the employee in the Other Leave Balances sheet
      for (let i = 0; i < otherLeaveRecords.length; i++) {
        if (otherLeaveRecords[i][A_BALANCE_EMAIL] === employeeEmail) {
          if (subLeaveType === "Bereavement Leave") {
            leaveUsed = otherLeaveRecords[i][A_USED_BL] || 0;
            leaveRemaining = otherLeaveRecords[i][A_REMAINING_BL] || 0;
            leaveTitle = "Bereavement Leave";
            headerColor = "#6B3410";
            subHeaderColor = "#D4A574";
          } else if (subLeaveType === "Paternity Leave") {
            leaveUsed = otherLeaveRecords[i][A_USED_PL] || 0;
            leaveRemaining = otherLeaveRecords[i][A_REMAINING_PL] || 0;
            leaveTitle = "Paternity Leave";
            headerColor = "#38761D";
            subHeaderColor = "#93C47D";
          } else if (subLeaveType === "Maternity Leave") {
            leaveUsed = otherLeaveRecords[i][A_USED_ML] || 0;
            leaveRemaining = otherLeaveRecords[i][A_REMAINING_ML] || 0;
            leaveTitle = "Maternity Leave";
            headerColor = "#741B47";
            subHeaderColor = "#C27BA0";
          }
          break;
        }
      }

      leaveBalanceTable = `
        <div class="balance-table-wrapper">
          <!-- ${leaveTitle} Table -->
          <table class="balance-table" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;text-align:center;">
            <thead>
              <tr>
                <th colspan="2" style="background-color:${headerColor};color:#fff;padding:12px;font-size:16px;font-weight:bold;">${leaveTitle}</th>
              </tr>
              <tr>
                <th style="background-color:${subHeaderColor};font-size:14px;font-weight:600;padding:12px;width:50%;">Used</th>
                <th style="background-color:${subHeaderColor};font-size:14px;font-weight:600;padding:12px;width:50%;">Remaining</th>
              </tr>
            </thead>
            <tbody>
              <tr style="background:#d1d5db30;">
                <td style="padding:15px 10px;font-size:16px;border-top:1px solid #e5e7eb;color:#555;">${leaveUsed}</td>
                <td style="padding:15px 10px;font-size:16px;border-top:1px solid #e5e7eb;color:#555;">${leaveRemaining}</td>
              </tr>
            </tbody>
          </table>
        </div>
      `;
    }

    //Email to Supervisor
    const templateDataSupervisor = {
      receiverName: supervisorName,
      bodyMessage: `
      <p>You've <strong>approved</strong> ${employeeFormatName}'s ${leaveType} - ${subLeaveType} request for ${formattedStartDate} to ${formattedEndDate}.</p>
      <p>Please advise if there's anything specific you need during his/her absence.</p>
      ${leaveBalanceTable}`,
    };
    sendTemplatedEmail(
      TEST_DEV_ACC,
      `Approved Leave for ${employeeFormatName}`,
      "TemplateGeneralEmailBody",
      templateDataSupervisor,
      attachments,
      false
    );

    // Email to employee
    const templateDataEmployee = {
      receiverName: employeeFormatName,
      bodyMessage: `
      <p>Great news! Your ${leaveType} - ${subLeaveType} from ${formattedStartDate} to ${formattedEndDate} has been <strong>approved</strong> by ${supervisorName}.</p>
      <p>Please advise if there's anything specific you need during his/her absence.</p>
      ${leaveBalanceTable}`,
    };
    sendTemplatedEmail(
      TEST_EMPLOYEE_ACC,
      `${leaveType} Application Status`,
      "TemplateGeneralEmailBody",
      templateDataEmployee,
      attachments,
      false
    );

    if (
      (leaveType === "Vacation Leave" && subLeaveType === "Unexcused") ||
      (leaveType === "Sick Leave" && subLeaveType === "Unexcused") ||
      leaveType === "Vacation Leave" ||
      leaveType === "Sick Leave" ||
      (leaveType === "Other" && subLeaveType === "Maternity Leave") ||
      (leaveType === "Other" && subLeaveType === "Paternity Leave") ||
      (leaveType === "Other" && subLeaveType === "Bereavement Leave") ||
      (leaveType === "Vacation Leave" && subLeaveType === "Emergency Leave") ||
      (leaveType === "Sick Leave" && subLeaveType === "Emergency Leave")
    ) {
      copyToApprovedPaidLeaves(requestLeaveData);
    } else if (
      (leaveType === "Other" && subLeaveType === "Unexcused") ||
      (leaveType === "Other" && subLeaveType === "Emergency Leave") ||
      subLeaveType === "Unpaid Leave" ||
      subLeaveType === "Half-day" ||
      subLeaveType === "Undertime"
    ) {
      copyToApprovedUnpaidLeaves(requestLeaveData);
    }

    const approvedAlertTemplate =
      HtmlService.createTemplateFromFile("TemplateAlertUrl");
    approvedAlertTemplate.alertType = "approve";
    approvedAlertTemplate.header = "<p>Leave Request Approved</p>";
    approvedAlertTemplate.message = `<p><strong>${employeeFormatName}</strong>'s leave request has been successfully approved.</p>`;
    approvedAlertTemplate.requestId = requestId;

    return HtmlService.createHtmlOutput(
      approvedAlertTemplate.evaluate().getContent()
    );
  }

  if (action === "reject") {
    sheet.getRange(row, COL_N_MAIN_STATUS).setValue("Rejected");
    sheet.getRange(row, COL_N_FORWARDING_STATUS).setValue("Not Forwarded");

    // Email to Employee
    const templateDataEmployee = {
      receiverName: employeeFormatName,
      bodyMessage: `
      <p>Your ${leaveType} - ${subLeaveType} from ${formattedStartDate} to ${formattedEndDate} was <strong>rejected</strong>. Please contact your approver, ${supervisorName} for more information.</p>`,
    };
    sendTemplatedEmail(
      TEST_EMPLOYEE_ACC,
      `${leaveType} Application Status`,
      "TemplateGeneralEmailBody",
      templateDataEmployee,
      attachments,
      false
    );

    // Email to Supervisor
    const templateDataSupervisor = {
      receiverName: supervisorName,
      bodyMessage: `
      <p>You've <strong>rejected</strong> ${employeeFormatName}'s ${leaveType} - ${subLeaveType} request for ${formattedStartDate} to ${formattedEndDate}.</p>`,
    };
    sendTemplatedEmail(
      TEST_DEV_ACC,
      `Rejected Leave for ${employeeFormatName}`,
      "TemplateGeneralEmailBody",
      templateDataSupervisor,
      attachments,
      false
    );

    copyToRejectedLeaves(requestLeaveData); // Log the rejected leave to "Rejected Leaves" sheet

    const rejectAlertTemplate =
      HtmlService.createTemplateFromFile("TemplateAlertUrl");
    rejectAlertTemplate.alertType = "reject";
    rejectAlertTemplate.header = "<p>Leave Request Rejected</p>";
    rejectAlertTemplate.message = `<p><strong>${employeeFormatName}</strong>'s leave application has been successfully rejected.</p>`;
    rejectAlertTemplate.requestId = requestId;

    return HtmlService.createHtmlOutput(
      rejectAlertTemplate.evaluate().getContent()
    );
  }

  if (action === "forward") {
    sheet.getRange(row, COL_N_FORWARDING_STATUS).setValue("Forwarded");

    // Email to accounting
    const templateDataAccounting = {
      receiverName: "Receiver Name",
      bodyMessage: `
      <p>Please note that ${employeeFormatName}'s ${leaveType} - ${subLeaveType} from <strong>${formattedStartDate} to ${formattedEndDate}</strong> has been <strong>approved</strong> by ${supervisorName}.</p>
      <p>Please update your records. Thanks!</p>
      <p style="display: none;">View <a href="${SHEET_URL}">Leave Request Records</a></p>
      `,
    };
    sendTemplatedEmail(
      TEST_ACCOUNTING_EMAIL,
      `Approved Leave for ${employeeFormatName}`,
      "TemplateGeneralEmailBody",
      templateDataAccounting,
      [],
      false,
      TEST_HR_EMAIL
    );

    // Email to SB
    const templateDataSb = {
      receiverName: "Receiver Name",
      bodyMessage: `
      <p>${employeeFormatName}'s ${leaveType} - ${subLeaveType} request for <strong>${formattedStartDate} to ${formattedEndDate}</strong> has been <strong>approved</strong> by ${supervisorName}.</p>
      <p>Please advise if there's anything specific you need during his/her absence.</p>
      <p style="display: none;">View <a href="${SHEET_URL}">Leave Request Records</a></p>
      `,
    };
    sendTemplatedEmail(
      TEST_SB_EMAIL,
      `Approved Leave for ${employeeFormatName}`,
      "TemplateGeneralEmailBody",
      templateDataSb,
      attachments,
      false,
      TEST_SD_EMAIL
    );

    addLeaveToCalendar(
      employeeEmail,
      employeeFormatName,
      leaveStartDate,
      leaveEndDate,
      leaveType,
      supervisorName,
      employeeReason,
      noOfDaysOrHours
    );

    const forwardAlertTemplate =
      HtmlService.createTemplateFromFile("TemplateAlertUrl");
    forwardAlertTemplate.alertType = "forward";
    forwardAlertTemplate.header = "<p>Leave Request Forwarded</p>";
    forwardAlertTemplate.message = `<p><strong>${employeeFormatName}</strong>'s leave application has been successfully forwarded.</p>`;
    forwardAlertTemplate.requestId = requestId;

    return HtmlService.createHtmlOutput(
      forwardAlertTemplate.evaluate().getContent()
    );
  }
}
