function checkApprovalStatus() {

  const now = new Date();
  const currentHour = now.getHours(); // Get current hour in 24-hour format

  if (currentHour < 6 || currentHour >= 17) {
    return; // Stop execution if outside the allowed time range
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ALL_RECORDS);
  const data = sheet.getDataRange().getValues(); // Get all data in the sheet

  for (let row = 1; row < data.length; row++) { // Start from row 2 (assuming row 1 has headers)
    const supervisorEmail = data[row][COL_A_APPROVER];
    const status = data[row][COL_A_MAIN_STATUS]; 
    const forwardedStatus = data[row][COL_A_FORWARDING_STATUS];

    if (status && status.trim() === "Approved" && (!forwardedStatus || forwardedStatus.trim() === "")) {
      sendApprovalEmail(row + 1, sheet); // Send email (row +1 because index starts from 0)
      sheet.getRange(row + 1, COL_N_FORWARDING_STATUS).setValue("Pending");
    }
  }
}

function sendApprovalEmail(row, sheet) {
  const employeeName = sheet.getRange(row, COL_N_FULL_NAME).getValue(); // Column C
  let employeeFormatName = formatName(employeeName);

  const leaveType = sheet.getRange(row, COL_N_MAIN_LEAVE).getValue(); // Column F
  const subLeaveType = sheet.getRange(row, COL_N_SUB_LEAVE).getValue(); // Column G
  const leaveStartDate = formatDate(sheet.getRange(row, COL_N_START_DATE).getValue()); // Column H
  const leaveEndDate = formatDate(sheet.getRange(row, COL_N_END_DATE).getValue()); // Column I
  const employeeAttachments = sheet.getRange(row, COL_N_ATTACHMENT).getValue(); // Column L
  const supervisorEmail = sheet.getRange(row, COL_N_APPROVER).getValue(); // Column M
  const supervisorName = getSupervisorName(supervisorEmail);

  const formattedStartDate = formatDate(leaveStartDate);
  const formattedEndDate = formatDate(leaveEndDate);

  // Generate the forwarded URL
  const forwardedUrl = createApprovalUrl(row, "forward");

  const attachments = getAttachmentBlobs(employeeAttachments);

  const templateData = {
    receiverName: "Name",
    bodyMessage: `
    <p>${employeeFormatName}'s ${leaveType} - ${subLeaveType} request for <strong>${formattedStartDate} to ${formattedEndDate}</strong> has been <strong>approved</strong> by ${supervisorName}.</p>
    <p>Please advise if there's anything specific you need during his/her absence.</p>
    <a href="${forwardedUrl}" class="btn btn-forward">Forward Leave</a>`
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