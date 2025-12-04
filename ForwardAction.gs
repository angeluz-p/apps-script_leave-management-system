function checkApprovalStatus() {
  const now = new Date();
  const currentHour = now.getHours(); // Get current hour in 24-hour format

  if (currentHour < 6 || currentHour >= 17) {
    debugLog("checkApprovalStatus: Outside allowed time range (6 AM - 5 PM)");
    return; // Stop execution if outside the allowed time range
  }

  const sheet = ACTIVE_SHEET.getSheetByName(ALL_RECORDS);

  if (!sheet) {
    debugLog("Error: ALL_RECORDS sheet not found");
    return;
  }

  const data = sheet.getDataRange().getValues(); // Get all data in the sheet

  for (let row = 1; row < data.length; row++) {
    // Start from row 1 (index 0 is header)
    const supervisorEmail = data[row][COL_A_APPROVER];
    const status = data[row][COL_A_MAIN_STATUS];
    const forwardedStatus = data[row][COL_A_FORWARDING_STATUS];
    const requestId = data[row][COL_A_REQUEST_ID];

    // Check if approved but not yet forwarded
    if (
      status &&
      status.trim() === "Approved" &&
      (!forwardedStatus || forwardedStatus.trim() === "")
    ) {
      // Additional validation: ensure Request ID exists
      if (!requestId || requestId.trim() === "") {
        debugLog(
          `Row ${
            row + 1
          }: Approved request missing Request ID, generating now...`
        );
        const newRequestId = generateRequestIdEnhanced(sheet);
        sheet.getRange(row + 1, COL_N_REQUEST_ID).setValue(newRequestId);
      }

      try {
        sendApprovalEmail(row + 1, sheet); // Send email (row +1 because index starts from 0)
        sheet.getRange(row + 1, COL_N_FORWARDING_STATUS).setValue("Pending");
        debugLog(
          `Row ${row + 1}: Forwarding email sent, status set to Pending`
        );
      } catch (error) {
        Logger.log(
          `Error sending approval email for row ${row + 1}: ${error.message}`
        );
      }
    }
  }
}

function sendApprovalEmail(row, sheet) {
  // Get existing Request ID from the specified row (not lastRow)
  const existingRequestId = sheet.getRange(row, COL_N_REQUEST_ID).getValue();
  let requestId;

  if (!existingRequestId) {
    requestId = generateRequestIdEnhanced(sheet);
    sheet.getRange(row, COL_N_REQUEST_ID).setValue(requestId);
    debugLog(
      `Generated new Request ID ${requestId} in row ${row}, column ${COL_N_REQUEST_ID}`
    );
  } else {
    requestId = existingRequestId;
    debugLog(
      `Using existing Request ID ${requestId} in row ${row}, column ${COL_N_REQUEST_ID}`
    );
  }

  const employeeName = sheet.getRange(row, COL_N_FULL_NAME).getValue();
  let employeeFormatName = formatName(employeeName);
  const leaveType = sheet.getRange(row, COL_N_MAIN_LEAVE).getValue();
  const subLeaveType = sheet.getRange(row, COL_N_SUB_LEAVE).getValue();
  const leaveStartDate = sheet.getRange(row, COL_N_START_DATE).getValue();
  const leaveEndDate = sheet.getRange(row, COL_N_END_DATE).getValue();
  const employeeAttachments = sheet.getRange(row, COL_N_ATTACHMENT).getValue();
  const supervisorEmail = sheet.getRange(row, COL_N_APPROVER).getValue();
  const supervisorName = getSupervisorName(supervisorEmail);

  const formattedStartDate = formatDate(leaveStartDate);
  const formattedEndDate = formatDate(leaveEndDate);

  // Generate the forwarded URL using requestId
  const forwardedUrl = createApprovalUrl(requestId, "forward");
  const attachments = getAttachmentBlobs(employeeAttachments);

  const templateData = {
    receiverName: "Receiver Name",
    bodyMessage: `
    <p>${employeeFormatName}'s ${leaveType} - ${subLeaveType} request for <strong>${formattedStartDate} to ${formattedEndDate}</strong> has been <strong>approved</strong> by ${supervisorName}.</p>
    <p>Please advise if there's anything specific you need during his/her absence.</p>
    <a href="${forwardedUrl}" class="btn btn-forward">Forward Leave</a>`,
  };

  sendTemplatedEmail(
    TEST_SD_EMAIL,
    `Approved Leave for ${employeeFormatName}`,
    "TemplateGeneralEmailBody",
    templateData,
    attachments,
    true
  );
}
