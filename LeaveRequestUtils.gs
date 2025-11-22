// Modify the scriptUrl when creating a new copy of the sheet
function createApprovalUrl(row, action) {
  const scriptUrl = "https://script.google.com/macros/s/AKfycbyaQfBPr41qCRu4q-EOSHEKBofBVpJ9zs0gkkruDkrz/dev";
  return `${scriptUrl}?row=${row}&action=${action}`;
}

/**
 * Change these configurations once to update everywhere
 */ 

// Get Sheet URL dynamically
const sheet = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_URL = sheet.getUrl();

// Sheet name configuration
const ALL_RECORDS = "All Records";
const LEAVE_BALANCES = "Leave Balances";
const OTHER_LEAVE_BALANCES = "Other Leave Balances";
const RBA_LEAVE_BALANCES = "RBA Leave Balances";
const APPROVED_PAID_LEAVES = "Approved Paid Leaves";
const APPROVED_UNPAID_LEAVES = "Approved Unpaid Leaves";
const REJECTED_LEAVE_APPLICATIONS = "Rejected Leave Applications";

// Email configuration
const TEST_ACCOUNTING_EMAIL = "example@gmail.com";
const TEST_SB_EMAIL = "example@gmail.com";
const TEST_SD_EMAIL = "example@gmail.com";
const TEST_HR_EMAIL = "example@gmail.com";

const TEST_DEV_ACC = "example@gmail.com";
const TEST_EMPLOYEE_ACC = "example@gmail.com";
const TEST_ALIAS_ACC = "example@gmail.com";

// List of additional allowed emails for filing a leave not listed on the "Leave Balances" sheet
const ALLOWED_EMAILS = ["allowed1@gmail.com", "allowed2@gmail.com"];

// Sheet columns configuration
// ARRAY
const COL_A_TIMESTAMP = 0;
const COL_A_EMAIL_ADDRESS = 1;
const COL_A_FULL_NAME = 2;
const COL_A_JOB_TITLE = 3;
const COL_A_DEPARTMENT = 4;
const COL_A_MAIN_LEAVE = 5;
const COL_A_SUB_LEAVE = 6;
const COL_A_START_DATE = 7;
const COL_A_END_DATE = 8;
const COL_A_LEAVE_DURATION = 9;
const COL_A_REASON = 10;
const COL_A_ATTACHMENT = 11;
const COL_A_APPROVER = 12;
const COL_A_MAIN_STATUS = 13;
const COL_A_FORWARDING_STATUS = 14;
const COL_A_AC_NO = 15;
const COL_A_REQUEST_ID = 16;

// NOT ARRAY
const COL_N_TIMESTAMP = 1;
const COL_N_EMAIL_ADDRESS = 2;
const COL_N_FULL_NAME = 3;
const COL_N_JOB_TITLE = 4;
const COL_N_DEPARTMENT = 5;
const COL_N_MAIN_LEAVE = 6;
const COL_N_SUB_LEAVE = 7;
const COL_N_START_DATE = 8;
const COL_N_END_DATE = 9;
const COL_N_LEAVE_DURATION = 10;
const COL_N_REASON = 11;
const COL_N_ATTACHMENT = 12;
const COL_N_APPROVER = 13;
const COL_N_MAIN_STATUS = 14;
const COL_N_FORWARDING_STATUS = 15;
const COL_N_AC_NO = 16;
const COL_N_REQUEST_ID = 17;

// Leave Balances Configurations
const A_BALANCE_EMAIL = 1;
const BALANCE_EMAIL = 2;

const USED_VL = 3;
const REMAINING_VL = 4;
const USED_SL = 5;
const REMAINING_SL = 6;

const A_USED_VL = 2;
const A_REMAINING_VL = 3;
const A_USED_SL = 4;
const A_REMAINING_SL = 5;

// RBA
const JAN_JUN_USED_VL = 3;
const JAN_JUN_REMAINING_VL = 4;
const JAN_JUN_USED_SL = 5;
const JAN_JUN_REMAINING_SL = 6;

const JUL_DEC_USED_VL = 8;
const JUL_DEC_REMAINING_VL = 9;
const JUL_DEC_USED_SL = 10;
const JUL_DEC_REMAINING_SL = 11;

const A_JAN_JUN_USED_VL = 2;
const A_JAN_JUN_REMAINING_VL = 3;
const A_JAN_JUN_USED_SL = 4;
const A_JAN_JUN_REMAINING_SL = 5;

const A_JUL_DEC_USED_VL = 7;
const A_JUL_DEC_REMAINING_VL = 8;
const A_JUL_DEC_USED_SL = 9;
const A_JUL_DEC_REMAINING_SL = 10;

// OTHER LEAVE
const USED_BL = 3;
const REMAINING_BL = 4;
const USED_PL = 5;
const REMAINING_PL = 6;
const USED_ML = 7;
const REMAINING_ML = 8;

const A_USED_BL = 2;
const A_REMAINING_BL = 3;
const A_USED_PL = 4;
const A_REMAINING_PL = 5;
const A_USED_ML = 6;
const A_REMAINING_ML = 7;

// Covered period configuration
// Date range configuration - automatically updates based on current date
const today = new Date();
const currentYear = today.getFullYear();

// Check if we're currently in the period from Dec 26 onwards
const dec26ThisYear = new Date(currentYear, 11, 26); // Dec 26 of current year

let cycleStartYear, cycleEndYear;

if (today >= dec26ThisYear) {
  // We're after Dec 26 this year, so new cycle starts this year
  cycleStartYear = currentYear;
  cycleEndYear = currentYear + 1;
} else {
  // We're before Dec 26 this year, so cycle started last year
  cycleStartYear = currentYear - 1;
  cycleEndYear = currentYear;
}

const PERIOD_ALLOWED_START_DATE = new Date(cycleStartYear, 11, 26); // Dec 26 - Previous Year
const PERIOD_ALLOWED_END_DATE = new Date(cycleEndYear, 11, 25);   // Dec 25 - Current Year

const DEBUG_MODE = false;

function debugLog(message) {
  if (DEBUG_MODE) {
    Logger.log(message);
  }
}

// Include external HTML/CSS files in templates
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Smart email sender - automatically chooses system or user email
function sendTemplatedEmail(recipient, subject, templateName, templateData, attachments = [], useSystemSender = false, cc = null, bcc = null) {
  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    
    Object.keys(templateData).forEach(key => {
      template[key] = templateData[key];
    });
    
    const htmlBody = template.evaluate().getContent();
    
    const emailOptions = {
      htmlBody: htmlBody,
      attachments: attachments.length > 0 ? attachments : undefined
    };
    
    // Add 'from' parameter only if using system sender
    if (useSystemSender) {
      emailOptions.from = TEST_DEV_ACC;
      emailOptions.name = "Leave Management";
    }

    if (cc) {
      emailOptions.cc = cc;
    }
    
    if (bcc) {
      emailOptions.bcc = bcc;
    }
    
    GmailApp.sendEmail(recipient, subject, "", emailOptions);
    
    const senderInfo = useSystemSender ? TEST_DEV_ACC : Session.getActiveUser().getEmail();
    const ccInfo = cc ? `, CC: ${cc}` : '';
    Logger.log(`✅ Email sent to ${recipient}${ccInfo} from ${senderInfo}: ${subject}`);
    return true;
  } catch (error) {
    Logger.log(`❌ Failed to send email to ${recipient}: ${error.message}`);
    return false;
  }
}

// Retrieves Google Drive file attachments from a comma-separated list of URLs/IDs
function getAttachmentBlobs(employeeAttachments) {
  const attachments = [];
  
  if (!employeeAttachments) {
    return attachments;
  }
  
  const attachmentLinks = employeeAttachments.split(',');
  
  attachmentLinks.forEach(link => {
    try {
      let fileId = extractDriveFileId(link.trim());
      if (fileId) {
        let file = DriveApp.getFileById(fileId);
        attachments.push(file.getBlob());
      }
    } catch (error) {
      Logger.log("Error fetching attachment: " + error.message);
    }
  });
  
  return attachments;
}

// Function to extract Google Drive File ID from a URL
function extractDriveFileId(url) {
  const match = url.match(/[-\w]{25,}/); // Extracts the file ID from the Google Drive URL
  return match ? match[0] : null;
}

function parseSheetDate(sheetDate) {
  if (!sheetDate) return null; // Handle empty or null values

  // ✅ If it's already a Date object, return it directly
  if (sheetDate instanceof Date && !isNaN(sheetDate.getTime())) {
    return sheetDate;
  }

  // ✅ If it's a string, try parsing it
  if (typeof sheetDate === "string") {
    sheetDate = sheetDate.trim(); // Remove spaces/tabs
    let parsedDate = new Date(sheetDate);

    if (!isNaN(parsedDate.getTime())) {
      return parsedDate;
    } else {
      Logger.log(`❌ Invalid string date format: ${sheetDate}`);
      return null;
    }
  } 
  
  // ✅ If it's a number (Google Sheets serial date), convert it
  if (typeof sheetDate === "number") {
    return new Date(1899, 11, 30 + sheetDate);
  }

  Logger.log(`❌ Unexpected date format: ${sheetDate}`);
  return null;
}

// Format name from "Last, First Middle" to "First Middle Last"
function formatName(name) {
  if (!name) return ""; // Handle empty or undefined names

  function capitalize(str) {
    return str
      .trim()
      .split(/\s+/) // Split by spaces
      .map(word =>
        word
          .split("-") // Split hyphenated names
          .map(part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()) // Capitalize first letter, lowercase rest
          .join("-") // Rejoin with hyphen
      )
      .join(" "); // Rejoin words with spaces
  }

  if (name.includes(",")) {
    const parts = name.split(",");
    return capitalize(parts[1].trim()) + " " + capitalize(parts[0].trim()); // Swap and capitalize
  } else {
    return capitalize(name); // Capitalize properly if no comma
  }
}

function formatDate(date) {
    const options = { year: 'numeric', month: 'long', day: 'numeric' };
    return new Date(date).toLocaleDateString('en-US', options);
}

function parseLeaveHours(leaveInput) {
  if (typeof leaveInput === "number") return leaveInput; // If already a number, return it

  leaveInput = leaveInput.toString().trim().toLowerCase();

  if (leaveInput.includes("full day")) return 1;
  if (leaveInput.includes("half day")) return 0.5;

  // Extract only the first valid number from the input
  const match = leaveInput.match(/(\d+(\.\d+)?)/);

  return match ? parseFloat(match[1]) : NaN; // Convert to number or return NaN if not found
}

function getBalanceMessage(leaveType, subLeaveType, balance, isRba = false, isJanToJune = false, isJulyToDec = false) {
  let note = "";
  if (isRba) {
      const period = isJanToJune ? "January to June" : isJulyToDec ? "July to December" : "";
      note = `
      <tr class="card-note">
        <td colspan="2">
          <p>ℹ️  The displayed remaining balance covers the <strong>${period}</strong> period.</p>
        </td>
      </tr>`;
  }

  return `
    <tr class="last-row">
        <td class="info-label">Remaining Balance:</td>
        <td class="info-value">${balance === null ? '<span style="color: red;"><strong>No balance</strong> corresponds to the email you entered on the form.</span>' : balance}</td>
    </tr>
    ${note}`;
}