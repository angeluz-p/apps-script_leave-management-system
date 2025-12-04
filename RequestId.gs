function generateRequestId(sheet) {
  const currentYear = SHEET_YEAR;

  // Get all existing request IDs from Column Q
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  let maxNumber = 0;

  // Find the highest number for the current year
  for (let i = 1; i < values.length; i++) {
    // Skip header row
    const existingId = values[i][COL_A_REQUEST_ID]; // Column Q (0-based index 12)
    if (existingId && typeof existingId === "string") {
      const match = existingId.match(/^(\d{4})-(\d{4})$/);
      if (match) {
        const year = parseInt(match[1]);
        const number = parseInt(match[2]);

        if (year === currentYear && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
  }

  // Increment the number
  const nextNumber = maxNumber + 1;
  const requestId = `${currentYear}-${nextNumber.toString().padStart(4, "0")}`;

  debugLog(`Generated Request ID: ${requestId}`);
  return requestId;
}

/**
 * Reset Request ID counter - Run this function manually to reset the counter
 * This will preserve existing Request IDs and ensure new submissions continue from the highest number
 */
function resetRequestIdCounter() {
  const sheet = ACTIVE_SHEET.getSheetByName(ALL_RECORDS);
  const currentYear = SHEET_YEAR;

  // Get the data range
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const numRows = dataRange.getNumRows();

  if (numRows > 1) {
    // If there are data rows (beyond header)
    // First, find the highest existing request ID number
    let maxNumber = 0;

    for (let i = 1; i < values.length; i++) {
      // Skip header row
      const existingId = values[i][COL_A_REQUEST_ID]; // Column Q (0-based index 12)
      if (existingId && typeof existingId === "string") {
        const match = existingId.match(/^(\d{4})-(\d{4})$/);
        if (match) {
          const year = parseInt(match[1]);
          const number = parseInt(match[2]);

          if (year === currentYear && number > maxNumber) {
            maxNumber = number;
          }
        }
      }
    }

    // Store the maximum number in a script property for future reference
    PropertiesService.getScriptProperties().setProperty(
      "lastRequestNumber_" + currentYear,
      maxNumber.toString()
    );

    // Keep existing Request IDs - DO NOT clear Column Q

    debugLog(
      `Request ID counter has been reset. Highest existing number: ${maxNumber}. Next submission will start from ${currentYear}-${(
        maxNumber + 1
      )
        .toString()
        .padStart(4, "0")}.`
    );
  } else {
    debugLog("No data rows found to reset.");
  }
}

/**
 * Enhanced generateRequestId that considers stored maximum when no existing IDs are found
 */
function generateRequestIdEnhanced(sheet) {
  const currentYear = SHEET_YEAR;

  // Get all existing request IDs from Column Q
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  let maxNumber = 0;

  // Find the highest number for the current year
  for (let i = 1; i < values.length; i++) {
    // Skip header row
    const existingId = values[i][COL_A_REQUEST_ID]; // Column Q (0-based index 12)
    if (existingId && typeof existingId === "string") {
      const match = existingId.match(/^(\d{4})-(\d{4})$/);
      if (match) {
        const year = parseInt(match[1]);
        const number = parseInt(match[2]);

        if (year === currentYear && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
  }

  // If no existing IDs found, check if we have a stored maximum from previous reset
  if (maxNumber === 0) {
    const storedMax = PropertiesService.getScriptProperties().getProperty(
      "lastRequestNumber_" + currentYear
    );
    if (storedMax) {
      maxNumber = parseInt(storedMax);
      debugLog(
        `No existing request IDs found. Using stored maximum: ${maxNumber}`
      );
    }
  }

  // Increment the number
  const nextNumber = maxNumber + 1;
  const requestId = `${currentYear}-${nextNumber.toString().padStart(4, "0")}`;

  debugLog(`Generated Request ID: ${requestId}`);
  return requestId;
}

/**
 * Generate Request IDs for empty cells only - Run this manually to fill missing Request IDs
 * This will only add Request IDs to rows that don't already have them
 */
function fillMissingRequestIds() {
  const sheet = ACTIVE_SHEET.getSheetByName(ALL_RECORDS);
  const currentYear = SHEET_YEAR;

  // Get the data range
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const numRows = dataRange.getNumRows();

  if (numRows > 1) {
    // If there are data rows (beyond header)
    // First, find the highest existing request ID number
    let maxNumber = 0;

    for (let i = 1; i < values.length; i++) {
      // Skip header row
      const existingId = values[i][COL_A_REQUEST_ID]; // Column Q (0-based index 12)
      if (existingId && typeof existingId === "string") {
        const match = existingId.match(/^(\d{4})-(\d{4})$/);
        if (match) {
          const year = parseInt(match[1]);
          const number = parseInt(match[2]);

          if (year === currentYear && number > maxNumber) {
            maxNumber = number;
          }
        }
      }
    }

    // Check for stored maximum if current max is 0
    if (maxNumber === 0) {
      const storedMax = PropertiesService.getScriptProperties().getProperty(
        "lastRequestNumber_" + currentYear
      );
      if (storedMax) {
        maxNumber = parseInt(storedMax);
      }
    }

    let generatedCount = 0;

    // Fill missing Request IDs
    for (let i = 2; i <= numRows; i++) {
      const existingId = sheet.getRange(i, COL_N_REQUEST_ID).getValue();

      if (!existingId || existingId === "") {
        // Generate new Request ID for empty cell
        maxNumber++;
        const requestId = `${currentYear}-${maxNumber
          .toString()
          .padStart(4, "0")}`;
        sheet.getRange(i, COL_N_REQUEST_ID).setValue(requestId);
        generatedCount++;
        debugLog(`Generated Request ID ${requestId} for row ${i}`);
      }
    }

    // Update the stored maximum
    PropertiesService.getScriptProperties().setProperty(
      "lastRequestNumber_" + currentYear,
      maxNumber.toString()
    );

    debugLog(
      `Filled ${generatedCount} missing Request IDs. Next submission will use ${currentYear}-${(
        maxNumber + 1
      )
        .toString()
        .padStart(4, "0")}.`
    );
  } else {
    debugLog("No data rows found.");
  }
}

/**
 * Clear the stored request number counter for the current year
 * Use this if you want to completely start over from 0001
 */
function clearStoredCounter() {
  const currentYear = SHEET_YEAR;
  PropertiesService.getScriptProperties().deleteProperty(
    "lastRequestNumber_" + currentYear
  );
  debugLog(
    `Cleared stored counter for year ${currentYear}. Next request will start from 0001.`
  );
}
