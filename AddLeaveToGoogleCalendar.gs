function addLeaveToCalendar(employeeEmail, employeeFormatName, leaveStartDate, leaveEndDate, leaveType, supervisorName, employeeReason, noOfDaysOrHours){
  try {
    // Get the employee's primary calendar
    const calendar_id = "d927aa8d6b478e0a33beb506e01ba3442c0809dbec885678fac560ffb2b14c75@group.calendar.google.com" // Replace with your actual calendar ID
    const calendar = CalendarApp.getCalendarById(calendar_id);
    
    /* const calendar = CalendarApp.getDefaultCalendar(); */
    const startDate = new Date(leaveStartDate);
    const endDate = new Date(leaveEndDate);

    if (!calendar) {
      Logger.log("Error: Specified calendar not found.");
      return;
    }

    if (isNaN(startDate) || isNaN(endDate)) {
      Logger.log("Error: Invalid date format.");
      return;
    }

    // Check if an event for the same employee and leave period already exists
    const events = calendar.getEvents(
      new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate(), 0, 0, 0),
      new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 23, 59, 59)
    );
    const eventTitle = `${employeeFormatName} - ${leaveType}`;
    
    const existingEvent = events.find(event => {
      const titleMatch = event.getTitle() === eventTitle;
      const description = event.getDescription();
      const startMatch = event.getStartTime().getTime() === startDate.getTime();
      const endMatch = event.getEndTime().getTime() === endDate.getTime();
      
      // Ensure description is not null before checking for the email
      const emailMatch = description && description.includes(employeeEmail);

      return titleMatch && startMatch && endMatch && emailMatch;
    });

    if (existingEvent) {
      Logger.log("Leave request already exists in the calendar.");
      return;
    }
    
    // Ensure end date is set correctly (full-day event)
    endDate.setDate(endDate.getDate() + 1);

    // Create an event in Google Calendar
    calendar.createAllDayEvent(
      eventTitle, 
      startDate, 
      endDate,
      {
        description: `Employee Email: ${employeeEmail}\nAM/PM/All Day: ${noOfDaysOrHours}\nApprover: ${supervisorName}`}
    );

    Logger.log("Leave request successfully added to Google Calendar.");
  } catch (error) {
    Logger.log("Error adding event to Google Calendar: " + error.message);
  }
}