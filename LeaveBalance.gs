function updateLeaveBalances() {
  const leaveRecordsSheet = ACTIVE_SHEET.getSheetByName(ALL_RECORDS);
  const leaveBalancesSheet = ACTIVE_SHEET.getSheetByName(LEAVE_BALANCES); // VL, SL, Emergency, Unexcused
  const otherLeaveBalancesSheet =
    ACTIVE_SHEET.getSheetByName(OTHER_LEAVE_BALANCES); // Bereavement, Parental, Maternal
  const rbaLeaveBalancesSheet = ACTIVE_SHEET.getSheetByName(RBA_LEAVE_BALANCES); // VL and SL for RBA

  if (!leaveRecordsSheet) {
    debugLog("The 'All Records' sheet was not found!");
    return;
  }

  if (!leaveBalancesSheet && !rbaLeaveBalancesSheet) {
    debugLog(
      "Neither 'Leave Balances' nor 'RBA Leave Balances' sheets were found!"
    );
    return;
  }

  const leaveData = leaveRecordsSheet.getDataRange().getValues();

  // Define RBA accounts and their separate tab
  const rbaAccounts = {
    "example@test.com": RBA_LEAVE_BALANCES,
    "example@test.com": RBA_LEAVE_BALANCES,
  };

  // Store leave balances per employee
  const leaveBalances = {};
  const otherLeaveBalances = {};

  // Define half-year periods for RBA accounts
  const currentYear = new Date().getFullYear();
  const janToJuneStart = new Date(currentYear, 0, 1);
  const janToJuneEnd = new Date(currentYear, 5, 30);
  const julyToDecStart = new Date(currentYear, 6, 1);
  const julyToDecEnd = new Date(currentYear, 11, 31);

  // Initialize leave balances for all employees in "Leave Balances"
  if (leaveBalancesSheet) {
    const lastRowLB = leaveBalancesSheet.getLastRow();
    if (lastRowLB >= 3) {
      const leaveBalancesData = leaveBalancesSheet
        .getRange(3, 1, lastRowLB - 2, leaveBalancesSheet.getLastColumn())
        .getValues();
      leaveBalancesData.forEach((row) => {
        const employeeEmail = row[A_BALANCE_EMAIL];
        if (employeeEmail) {
          if (rbaAccounts[employeeEmail]) {
            // Initialize RBA accounts with half-year tracking
            leaveBalances[employeeEmail] = {
              JanToJune_VL: 0,
              JanToJune_SL: 0,
              JanToJune_Emergency: 0,
              JanToJune_Unexcused: 0,
              JulyToDec_VL: 0,
              JulyToDec_SL: 0,
              JulyToDec_Emergency: 0,
              JulyToDec_Unexcused: 0,
              Total_VL: 0,
              Total_SL: 0,
            };
          } else {
            // Initialize regular employees with simple tracking
            leaveBalances[employeeEmail] = {
              "Vacation Leave": 0,
              "Sick Leave": 0,
              "Emergency Leave": 0,
              Unexcused: 0,
            };
          }
        }
      });
    }
  }

  // Initialize RBA employees from "RBA Leave Balances" sheet
  if (rbaLeaveBalancesSheet) {
    const lastRowRBA = rbaLeaveBalancesSheet.getLastRow();
    if (lastRowRBA >= 3) {
      const rbaBalancesData = rbaLeaveBalancesSheet
        .getRange(3, 1, lastRowRBA - 2, rbaLeaveBalancesSheet.getLastColumn())
        .getValues();
      rbaBalancesData.forEach((row) => {
        const employeeEmail = row[A_BALANCE_EMAIL];
        if (
          employeeEmail &&
          rbaAccounts[employeeEmail] &&
          !leaveBalances[employeeEmail]
        ) {
          leaveBalances[employeeEmail] = {
            JanToJune_VL: 0,
            JanToJune_SL: 0,
            JanToJune_Emergency: 0,
            JanToJune_Unexcused: 0,
            JulyToDec_VL: 0,
            JulyToDec_SL: 0,
            JulyToDec_Emergency: 0,
            JulyToDec_Unexcused: 0,
            Total_VL: 0,
            Total_SL: 0,
          };
        }
      });
    }
  }

  // Initialize other leave balances for all employees in "Other Leave Balances"
  if (otherLeaveBalancesSheet) {
    const lastRowOLB = otherLeaveBalancesSheet.getLastRow();
    if (lastRowOLB >= 3) {
      const otherBalancesData = otherLeaveBalancesSheet
        .getRange(3, 1, lastRowOLB - 2, otherLeaveBalancesSheet.getLastColumn())
        .getValues();
      otherBalancesData.forEach((row) => {
        const employeeEmail = row[A_BALANCE_EMAIL];
        if (employeeEmail) {
          otherLeaveBalances[employeeEmail] = {
            "Bereavement Leave": 0,
            "Paternity Leave": 0,
            "Maternity Leave": 0,
          };
        }
      });
    }
  }

  // SINGLE LOOP through "All Records" to collect approved leave balances for ALL leave types
  leaveData.slice(1).forEach((row) => {
    const email = row[COL_A_EMAIL_ADDRESS]; // Column B - Employee Email
    const leaveType = row[COL_A_MAIN_LEAVE]; // Column F - Type of Leave
    const subLeaveType = row[COL_A_SUB_LEAVE]; // Column G - Sub Leave
    const startDate = new Date(row[COL_A_START_DATE]); // Column H - Start Date
    const endDate = new Date(row[COL_A_END_DATE]); // Column I - End Date
    let days = row[COL_A_LEAVE_DURATION]; // Column J - No. of Days
    const status = row[COL_A_MAIN_STATUS]; // Column N - Status

    if (!email || !leaveType || !days || status !== "Approved") return;

    // Convert "Full Day" and numeric values
    if (typeof days === "string") {
      days =
        days.toLowerCase() === "full day"
          ? 1
          : parseFloat(days.match(/[\d\.]+/)?.[0] || 0);
    }
    if (isNaN(days)) return;

    const isJanToJune = startDate >= janToJuneStart && endDate <= janToJuneEnd;
    const isJulyToDec =
      startDate >= julyToDecStart && startDate <= julyToDecEnd;

    // Process VL, SL, Emergency, and Unexcused Leave
    if (leaveType === "Vacation Leave" || leaveType === "Sick Leave") {
      // Create entry if not exists
      if (!leaveBalances[email]) {
        if (rbaAccounts[email]) {
          leaveBalances[email] = {
            JanToJune_VL: 0,
            JanToJune_SL: 0,
            JanToJune_Emergency: 0,
            JanToJune_Unexcused: 0,
            JulyToDec_VL: 0,
            JulyToDec_SL: 0,
            JulyToDec_Emergency: 0,
            JulyToDec_Unexcused: 0,
            Total_VL: 0,
            Total_SL: 0,
          };
        } else {
          leaveBalances[email] = {
            "Vacation Leave": 0,
            "Sick Leave": 0,
            "Emergency Leave": 0,
            Unexcused: 0,
          };
        }
      }

      // Process RBA accounts differently
      if (rbaAccounts[email]) {
        // Categorize leave types for RBA accounts with half-year tracking
        if (leaveType === "Vacation Leave" && subLeaveType === "VL/SL") {
          if (isJanToJune) leaveBalances[email].JanToJune_VL += days;
          if (isJulyToDec) leaveBalances[email].JulyToDec_VL += days;
          leaveBalances[email].Total_VL += days;
        } else if (leaveType === "Sick Leave" && subLeaveType === "VL/SL") {
          if (isJanToJune) leaveBalances[email].JanToJune_SL += days;
          if (isJulyToDec) leaveBalances[email].JulyToDec_SL += days;
          leaveBalances[email].Total_SL += days;
        } else if (
          (leaveType === "Vacation Leave" &&
            subLeaveType === "Emergency Leave") ||
          (leaveType === "Sick Leave" && subLeaveType === "Emergency Leave")
        ) {
          // Track emergency leave by half-year period
          if (isJanToJune) leaveBalances[email].JanToJune_Emergency += days;
          if (isJulyToDec) leaveBalances[email].JulyToDec_Emergency += days;
        } else if (
          (leaveType === "Vacation Leave" && subLeaveType === "Unexcused") ||
          (leaveType === "Sick Leave" && subLeaveType === "Unexcused")
        ) {
          // Track Unexcused by half-year period
          if (isJanToJune) leaveBalances[email].JanToJune_Unexcused += days;
          if (isJulyToDec) leaveBalances[email].JulyToDec_Unexcused += days;
        }
      } else {
        // Regular employee leave tracking
        if (leaveType === "Vacation Leave" && subLeaveType === "VL/SL")
          leaveBalances[email]["Vacation Leave"] += days;
        if (leaveType === "Sick Leave" && subLeaveType === "VL/SL")
          leaveBalances[email]["Sick Leave"] += days;
        if (
          (leaveType === "Vacation Leave" &&
            subLeaveType === "Emergency Leave") ||
          (leaveType === "Sick Leave" && subLeaveType === "Emergency Leave")
        )
          leaveBalances[email]["Emergency Leave"] += days;
        if (
          (leaveType === "Vacation Leave" && subLeaveType === "Unexcused") ||
          (leaveType === "Sick Leave" && subLeaveType === "Unexcused")
        )
          leaveBalances[email]["Unexcused"] += days;
      }
    }

    // Process Other Leave Types (Bereavement, Parental, Maternal)
    if (
      (leaveType === "Other" && subLeaveType === "Bereavement Leave") ||
      (leaveType === "Other" && subLeaveType === "Paternity Leave") ||
      (leaveType === "Other" && subLeaveType === "Maternity Leave")
    ) {
      // Create entry if not exists
      if (!otherLeaveBalances[email]) {
        otherLeaveBalances[email] = {
          "Bereavement Leave": 0,
          "Paternity Leave": 0,
          "Maternity Leave": 0,
        };
      }

      // Track other leave types
      if (leaveType === "Other" && subLeaveType === "Bereavement Leave") {
        otherLeaveBalances[email]["Bereavement Leave"] += days;
      } else if (leaveType === "Other" && subLeaveType === "Paternity Leave") {
        otherLeaveBalances[email]["Paternity Leave"] += days;
      } else if (leaveType === "Other" && subLeaveType === "Maternity Leave") {
        otherLeaveBalances[email]["Maternity Leave"] += days;
      }
    }
  });

  // ===== UPDATE RBA LEAVE BALANCES SHEET ===== //
  if (rbaLeaveBalancesSheet) {
    const lastRowRBA2 = rbaLeaveBalancesSheet.getLastRow();
    if (lastRowRBA2 >= 3) {
      const rbaData = rbaLeaveBalancesSheet
        .getRange(3, 1, lastRowRBA2 - 2, rbaLeaveBalancesSheet.getLastColumn())
        .getValues();

      rbaData.forEach((row, index) => {
        const email = row[A_BALANCE_EMAIL]; // Column B - Employee Email

        if (!email || !rbaAccounts[email] || !leaveBalances[email]) return;

        const rowIndex = index + 3; // Start from row 3
        const balance = leaveBalances[email];

        // RBA accounts have max 10 VL and 5 SL per 6-month period
        const maxVL = 10; // Max VL per 6 months for RBA
        const maxSL = 5; // Max SL per 6 months for RBA

        // Get emergency leave by period
        const janToJuneEmergency = balance.JanToJune_Emergency || 0;
        const julyToDecEmergency = balance.JulyToDec_Emergency || 0;
        const janToJuneUnexcused = balance.JanToJune_Unexcused || 0;
        const julyToDecUnexcused = balance.JulyToDec_Unexcused || 0;

        // Jan-Jun calculation with emergency distribution (SL first, then VL)
        let emergencyToJanSL = Math.min(
          janToJuneEmergency,
          maxSL - balance.JanToJune_SL
        );
        let janToJuneSL = balance.JanToJune_SL + emergencyToJanSL;

        let remainingJanEmergency = janToJuneEmergency - emergencyToJanSL;
        let janToJuneVL = balance.JanToJune_VL + remainingJanEmergency;

        // Jan-Jun unexcused distribution (VL first, then SL)
        let unexcusedToJanVL = Math.min(
          janToJuneUnexcused,
          maxVL - janToJuneVL
        );
        janToJuneVL += unexcusedToJanVL;

        let remainingJanUnexcused = janToJuneUnexcused - unexcusedToJanVL;
        janToJuneSL += remainingJanUnexcused;

        // Ensure we don't exceed maximums
        janToJuneSL = Math.min(janToJuneSL, maxSL);
        janToJuneVL = Math.min(janToJuneVL, maxVL);

        // Jul-Dec calculation with emergency distribution (SL first, then VL)
        let emergencyToJulSL = Math.min(
          julyToDecEmergency,
          maxSL - balance.JulyToDec_SL
        );
        let julyToDecSL = balance.JulyToDec_SL + emergencyToJulSL;

        let remainingJulEmergency = julyToDecEmergency - emergencyToJulSL;
        let julyToDecVL = balance.JulyToDec_VL + remainingJulEmergency;

        // Jul-Dec unexcused distribution (VL first, then SL)
        let unexcusedToJulVL = Math.min(
          julyToDecUnexcused,
          maxVL - julyToDecVL
        );
        julyToDecVL += unexcusedToJulVL;

        let remainingJulUnexcused = julyToDecUnexcused - unexcusedToJulVL;
        julyToDecSL += remainingJulUnexcused;

        // Ensure we don't exceed maximums
        julyToDecSL = Math.min(julyToDecSL, maxSL);
        julyToDecVL = Math.min(julyToDecVL, maxVL);

        // Calculate remaining leave
        const remainingJanToJuneVL = maxVL - janToJuneVL;
        const remainingJanToJuneSL = maxSL - janToJuneSL;
        const remainingJulyToDecVL = maxVL - julyToDecVL;
        const remainingJulyToDecSL = maxSL - julyToDecSL;

        // Update RBA Leave Balances sheet
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JAN_JUN_USED_VL)
          .setValue(janToJuneVL); // Col C - Jan-Jun VL Used
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JAN_JUN_REMAINING_VL)
          .setValue(remainingJanToJuneVL); // Col D - Jan-Jun VL Remaining
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JAN_JUN_USED_SL)
          .setValue(janToJuneSL); // Col E - Jan-Jun SL Used
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JAN_JUN_REMAINING_SL)
          .setValue(remainingJanToJuneSL); // Col F - Jan-Jun SL Remaining
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JUL_DEC_USED_VL)
          .setValue(julyToDecVL); // Col H - Jul-Dec VL Used
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JUL_DEC_REMAINING_VL)
          .setValue(remainingJulyToDecVL); // Col I - Jul-Dec VL Remaining
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JUL_DEC_USED_SL)
          .setValue(julyToDecSL); // Col J - Jul-Dec SL Used
        rbaLeaveBalancesSheet
          .getRange(rowIndex, JUL_DEC_REMAINING_SL)
          .setValue(remainingJulyToDecSL); // Col K - Jul-Dec SL Remaining
      });
    }
  }

  // ===== UPDATE REGULAR LEAVE BALANCES SHEET ===== //
  if (leaveBalancesSheet) {
    const lastRowLB2 = leaveBalancesSheet.getLastRow();
    if (lastRowLB2 >= 3) {
      const balancesData = leaveBalancesSheet
        .getRange(3, 1, lastRowLB2 - 2, leaveBalancesSheet.getLastColumn())
        .getValues();

      balancesData.forEach((row, index) => {
        const employeeEmail = row[A_BALANCE_EMAIL];

        // Skip excluded and RBA emails
        if (!employeeEmail || rbaAccounts[employeeEmail]) return;

        // Specific emails to exclude from leave balances
        const excludedEmails = EXCLUDED_EMAILS;
        if (excludedEmails.includes(employeeEmail)) return;

        const rowIndex = index + 3; // Start from row 3

        if (leaveBalances[employeeEmail]) {
          const usedVacationLeave =
            leaveBalances[employeeEmail]["Vacation Leave"] || 0;
          const usedSickLeave = leaveBalances[employeeEmail]["Sick Leave"] || 0;
          const usedEmergencyLeave =
            leaveBalances[employeeEmail]["Emergency Leave"] || 0;
          const usedUnexcusedLeave =
            leaveBalances[employeeEmail]["Unexcused"] || 0;

          // Set maximum leave allowances
          let maxSickLeave = 10; // Max SL for most employees
          let maxVacationLeave = 10; // Max VL for most employees

          // Employees with different leave balances (e.g., less than 10 or under accrual method)
          if (employeeEmail === "example@gmail.com") {
            maxVacationLeave = 17.5;
            maxSickLeave = 8.75;
          }

          if (employeeEmail === "example@gmail.com") {
            maxVacationLeave = 2.5;
            maxSickLeave = 2.5;
          }

          if (employeeEmail === "example@gmail.com") {
            maxVacationLeave = 14;
            maxSickLeave = 6;
          }

          if (employeeEmail === "example@gmail.com") {
            maxVacationLeave = 15;
            maxSickLeave = 7.5;
          }

          // Deduct emergency leave from SL first, then VL
          const emergencyToSL = Math.min(
            usedEmergencyLeave,
            maxSickLeave - usedSickLeave
          );
          let totalUsedSickLeave = usedSickLeave + emergencyToSL;

          const emergencyToVL = usedEmergencyLeave - emergencyToSL;
          let totalUsedVacationLeave = usedVacationLeave + emergencyToVL;

          // Deduct unexcused from VL first, then SL
          const unexcusedToVL = Math.min(
            usedUnexcusedLeave,
            maxVacationLeave - totalUsedVacationLeave
          );
          totalUsedVacationLeave += unexcusedToVL;

          const unexcusedToSL = usedUnexcusedLeave - unexcusedToVL;
          totalUsedSickLeave += unexcusedToSL;

          // Final used leave (ensure we don't exceed maximums)
          const finalUsedSickLeave = Math.min(totalUsedSickLeave, maxSickLeave);
          const finalUsedVacationLeave = Math.min(
            totalUsedVacationLeave,
            maxVacationLeave
          );

          // Remaining leaves
          const remainingSickLeave = maxSickLeave - finalUsedSickLeave;
          const remainingVacationLeave =
            maxVacationLeave - finalUsedVacationLeave;

          // Write to Leave Balances sheet
          leaveBalancesSheet
            .getRange(rowIndex, USED_VL)
            .setValue(finalUsedVacationLeave); // Col C - Used VL
          leaveBalancesSheet
            .getRange(rowIndex, REMAINING_VL)
            .setValue(remainingVacationLeave); // Col D - Remaining VL
          leaveBalancesSheet
            .getRange(rowIndex, USED_SL)
            .setValue(finalUsedSickLeave); // Col E - Used SL
          leaveBalancesSheet
            .getRange(rowIndex, REMAINING_SL)
            .setValue(remainingSickLeave); // Col F - Remaining SL
        }
      });
    }
  }

  // ===== UPDATE OTHER LEAVE BALANCES SHEET ===== //
  if (otherLeaveBalancesSheet) {
    const lastRowOLB2 = otherLeaveBalancesSheet.getLastRow();
    if (lastRowOLB2 >= 3) {
      const otherBalancesData = otherLeaveBalancesSheet
        .getRange(
          3,
          1,
          lastRowOLB2 - 2,
          otherLeaveBalancesSheet.getLastColumn()
        )
        .getValues();

      otherBalancesData.forEach((row, index) => {
        const employeeEmail = row[A_BALANCE_EMAIL];

        if (!employeeEmail) return;

        const rowIndex = index + 3; // Start from row 3

        if (otherLeaveBalances[employeeEmail]) {
          const usedBereavementLeave =
            otherLeaveBalances[employeeEmail]["Bereavement Leave"] || 0;
          const usedParentalLeave =
            otherLeaveBalances[employeeEmail]["Paternity Leave"] || 0;
          const usedMaternalLeave =
            otherLeaveBalances[employeeEmail]["Maternity Leave"] || 0;

          // Check if employee is from @test.com (no bereavement leave)
          const isRBAEmployee = employeeEmail.endsWith("@test.com");

          // Set maximum leave allowances
          const maxBereavementLeave = isRBAEmployee ? 0 : 5;
          const maxParentalLeave = 7;

          // Calculate used leave (ensure we don't exceed maximums)
          const finalUsedBereavementLeave = isRBAEmployee
            ? 0
            : Math.min(usedBereavementLeave, maxBereavementLeave);
          const finalUsedParentalLeave = Math.min(
            usedParentalLeave,
            maxParentalLeave
          );

          // Calculate remaining leaves
          const remainingBereavementLeave =
            maxBereavementLeave - finalUsedBereavementLeave;
          const remainingParentalLeave =
            maxParentalLeave - finalUsedParentalLeave;

          // Write to Other Leave Balances sheet
          if (!isRBAEmployee) {
            otherLeaveBalancesSheet
              .getRange(rowIndex, USED_BL)
              .setValue(finalUsedBereavementLeave); // Col C - Used Bereavement
            otherLeaveBalancesSheet
              .getRange(rowIndex, REMAINING_BL)
              .setValue(remainingBereavementLeave); // Col D - Remaining Bereavement
          }

          otherLeaveBalancesSheet
            .getRange(rowIndex, USED_PL)
            .setValue(finalUsedParentalLeave); // Col E - Used Parental
          otherLeaveBalancesSheet
            .getRange(rowIndex, REMAINING_PL)
            .setValue(remainingParentalLeave); // Col F - Remaining Parental

          // Only update Maternity Leave if employee has a valid Maternity Leave balance
          const currentUsedML = otherLeaveBalancesSheet
            .getRange(rowIndex, USED_ML)
            .getValue();
          if (currentUsedML !== "--") {
            otherLeaveBalancesSheet
              .getRange(rowIndex, USED_ML)
              .setValue(usedMaternalLeave);
          }
        }
      });
    }
  }

  Logger.log("Leave balances updated successfully!");
}
