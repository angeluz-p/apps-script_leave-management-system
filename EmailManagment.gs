// Function to get supervisor name from the email
function getSupervisorName(email) {
  const supervisorMapping = {
    "supervisor1@gmail.com": "Supervisor Name",
    "supervisor2@gmail.com": "Supervisor Name"
  };

  // Normalize email (remove spaces, convert to lowercase)
  const normalizedEmail = email.trim().toLowerCase();

  return supervisorMapping[normalizedEmail] || "Supervisor Not Found";
}

function isExemptedEmail(email) {
  const exemptedEmails = [
    "exempted1@gmail.com",
    "exempted2@gmail.com"
  ];
  return exemptedEmails.includes(email);
}

function isOtherExemptedEmail(email) {
  const otherExemptedEmails = [
    "exempted1@gmail.com",
    "exempted1@gmail.com"
  ];
  return otherExemptedEmails.includes(email);
}

function isParaplannerEmail(email) {
  return email === "example@gmail.com";
}