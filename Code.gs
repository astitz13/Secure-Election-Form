function sendEmails() {
  // Get sheet for email list
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Emails");
  // Gather form name and link from corresponding cells
  const formName = sheet.getRange("E1").getValue();
  const formLink = sheet.getRange("E2").getValue();
  // Get range for email list
  const listRange = sheet.getRange("A2:A");
  // Gather unique values for email list
  // First, the first column of each row is taken with index 0.
  // Then, values are filtered to be strings, of first occurrence/unique, and with a non-zero length (not empty)
  const data = listRange.getValues().map(row => row[0]).filter((value, idx, arr) => typeof(value) == "string" && arr.indexOf(value) == idx && value.length > 0);
  // Generate verification codes to be 20-character capital alphanumeric codes
  const verificationCodes = data.map(n => Array.from({length: 20}).fill(1).map(n => (chars => chars[Math.floor(Math.random() * chars.length)])("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890")).join(""));

  // Loop through indices of our email list
  for(let i in data) {
    // Get email address from index of email list
    const address = data[i];
    // Get corresponding verification code
    const code = verificationCodes[i];
    // Send email
    GmailApp.sendEmail(address, `Verification Code for Election Form "${formName}"`, `You have been invited to vote in the election, ${formName}!\n\nYour verification code is: ${code}\n\nLink to form: ${formLink}\n\nAs you place your vote, please remember to only vote once and do not share this verification code with anyone. Also, please make sure you entered your verification code correctly. It is recommended to copy and paste the code from this email. If you fail to provide a valid code, your vote will not be counted. If multiple submissions use the same code, the election must be re-run to ensure security.\n\nThank you for voting!`);
  }

  // Variable to store shuffled verification codes
  let shuffledCodes = [];
  // Loop until all verification codes are shuffled
  for(let i = verificationCodes.length; i > 0;) {
    // Pick random index of verification codes
    const idx = Math.floor(Math.random() * verificationCodes.length);

    // If the index is already used, try another
    if(verificationCodes[idx] == "used") continue;

    // Add randomly-picked verification code to shuffled codes
    shuffledCodes.push(verificationCodes[idx]);
    // Say the newly used code is used to prevent duplicates
    verificationCodes[idx] = "used";

    // One down, however many left to go!
    // This will only run if the continue before is not called
    i--;
  }

  // Loop through shuffled verification codes
  for(let i = 0; i < shuffledCodes.length; i++) {
    // Store the verification codes in B2:B range
    sheet.getRange("B" + (i + 2)).setValue(shuffledCodes[i]);
  }
}
