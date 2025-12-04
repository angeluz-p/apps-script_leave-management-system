function manageApproverAccess() {
  const now = new Date();
  const currentHour = now.getHours(); // 0 to 23

  if (currentHour < 5 || currentHour === 0) {
    return; // Stop execution before 5:00 AM or exactly at midnight
  }

  const ss = ACTIVE_SHEET;
  const sheet = ss.getSheetByName(ALL_RECORDS);

  if (!sheet) {
    Logger.log("Error: Sheet not found.");
    return;
  }

  const data = sheet.getDataRange().getValues(); // Get all data from the sheet
  const editorEmails = ["approver1@gmail.com", "approver2@gmail.com"];

  let grantedEmails = new Set(); // Use a Set to avoid duplicate entries

  for (let i = 1; i < data.length; i++) {
    // Start from row 2 (skip headers)
    let approverEmail = data[i][COL_A_APPROVER]; // Column M (Approver Email)
    let status = data[i][COL_A_MAIN_STATUS]; // Column N (Status)

    if (
      editorEmails.includes(approverEmail) &&
      (!status || status.trim() === "")
    ) {
      grantedEmails.add(approverEmail);
    }
  }

  editorEmails.forEach((email) => {
    if (grantedEmails.has(email)) {
      ss.addEditor(email);
      Logger.log(`Editor access granted to: ${email}`);
    } else {
      ss.removeEditor(email);
      Logger.log(`Editor access removed for: ${email}`);
    }
  });
}
