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
    GmailApp.sendEmail(address, `Verification Code for Election Form ${formName}`, `You have been invited to vote in the election, ${formName}!\n\nYour verification code is: ${code}\n\nLink to form: ${formLink}\n\nAs you place your vote, please remember to only vote once and do not share this verification code with anyone. Also, please make sure you entered your verification code correctly. It is recommended to copy and paste the code from this email. If you fail to provide a valid code, your vote will not be counted. If multiple submissions use the same code, the election must be re-run to ensure security.\n\nThank you for voting!`);
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

  // Delete sent emails
  const threads = GmailApp.search(`subject:"Verification Code for Election Form ${formName}"`, 0, data.length);
  GmailApp.moveThreadsToTrash(threads);
  for(let thread of threads) {
    Logger.log(thread.getId())
    Gmail.Users.Threads.remove("me", thread.getId());
  }
}

function generateColumn(idx) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let res = "";
  let n = idx - 1;
  while(n > 0) {
    res = chars[n % chars.length] + res;
    n = Math.floor(n / chars.length);
  }
  return res;
}

function createForm() {


  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const formSheets = sheets.filter(n=>n.getName().startsWith("Form Responses"));
  const mainFormName = formSheets[0].getName();
  if(!(sheets.map(n=>n.getName()).includes("Form Verified"))) {
    spreadsheet.insertSheet("Form Verified");
  }
  if(!(sheets.map(n=>n.getName()).includes("Emails"))) {
    spreadsheet.insertSheet("Emails");
  }
  const verifiedSheet = spreadsheet.getSheetByName("Form Verified");
  const emailsSheet = spreadsheet.getSheetByName("Emails");
  const formSheet = formSheets[0];
  emailsSheet.getRange("A1").setValue("Email List");
  emailsSheet.getRange("B1").setValue("Valid Codes");
  emailsSheet.getRange("D1").setValue("Form Name:");
  emailsSheet.getRange("D2").setValue("Form Link:");
  emailsSheet.getRange("B2:B").clearContent();
  const formHeaderRow = formSheet.getRange("A1:1").getValues()[0];
  const formWidth = formHeaderRow.indexOf("");
  const lastColumn = generateColumn(formWidth);

  verifiedSheet.getRange("A1").setValue(`=ARRAYFORMULA('${mainFormName}'!A1:${lastColumn})`);
  verifiedSheet.getRange("R1C" + (formWidth + 1)).setValue(`=ARRAYFORMULA(A1:${lastColumn}1)`);
  verifiedSheet.getRange("R2C" + (formWidth + 1)).setValue(`=FILTER(A$2:${lastColumn},MATCH(B$2:B,Emails!B1:B,0)>1)`);

  verifiedSheet.getRange("R1C" + (2*formWidth + 1)).setValue("Duplicates");
  verifiedSheet.getRange("R1C" + (2*formWidth + 2)).setValue("Recount Needed");

  const emailColumn = generateColumn(formWidth + 2);
  let dupes = [];
  for(let row = 2; row <= 1000; row++) {
    dupes.push([`=COUNTIF(${emailColumn}$2:$${emailColumn},${emailColumn}${row})`]);
  }
  verifiedSheet.getRange("R2C" + (2*formWidth + 1) + ":R1000C" + (2*formWidth + 1)).setValues(dupes);
  const duplicatesColumn = generateColumn(2*formWidth + 1);
  verifiedSheet.getRange("R2C"+ (2*formWidth + 2)).setValue(`=IF(COUNTIF(${duplicatesColumn}2:${duplicatesColumn},">1")>0,"Yes","No")`);
  verifiedSheet.hideColumns(1,formWidth);
}

function implementPlurality() {
  const active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(active.getName() != "Form Verified") {
    Logger.log("Error: not on Form Verified sheet");
    return;
  }
  const activeRange = SpreadsheetApp.getActiveRange();
  const width = activeRange.getWidth();
  const startCol = activeRange.getColumn();
  for(let col = startCol, i = 0; i < width; i++, col++) {
    const question = active.getRange("R1C" + col).getValue();
    const headerRow = active.getRange("A1:1").getValues()[0];
    const headerWidth = headerRow.indexOf("");
    const colName = generateColumn(col);

    active.getRange("R1C"+(headerWidth + 1)).setValue(question);
    active.getRange("R2C"+(headerWidth + 1)).setValue(`=UNIQUE($${colName}2:$${colName})`);
    
    active.getRange("R1C"+(headerWidth + 2)).setValue("Total Votes:");
    let votes = [];
    for(let row = 2; row <= 1000; row++) {
      votes.push([`=IFERROR(COUNTIF($${colName}2:$${colName}, ${generateColumn(headerWidth + 1) + row}),"")`])
    }
    active.getRange("R2C"+(headerWidth + 2)+":R1000C"+(headerWidth + 2)).setValues(votes);

    const candidateCol = generateColumn(headerWidth + 1);
    const voteCol = generateColumn(headerWidth + 2);
    active.getRange("R1C"+(headerWidth + 3)).setValue("Winners");
    active.getRange("R2C"+(headerWidth + 3)).setValue(`=SORT($${candidateCol}2:$${voteCol}, 2, FALSE)`);
    active.getRange("R1C"+(headerWidth + 4)).setValue("Votes");
  }
}

function implementFPTP() {
  const active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(active.getName() != "Form Verified") {
    Logger.log("Error: not on Form Verified sheet");
    return;
  }
  const activeRange = SpreadsheetApp.getActiveRange();
  const width = activeRange.getWidth();
  const startCol = activeRange.getColumn();
  for(let col = startCol, i = 0; i < width; i++, col++) {
    const question = active.getRange("R1C" + col).getValue();
    const headerRow = active.getRange("A1:1").getValues()[0];
    const headerWidth = headerRow.indexOf("");
    const colName = generateColumn(col);

    active.getRange("R1C"+(headerWidth + 1)).setValue(question);
    active.getRange("R2C"+(headerWidth + 1)).setValue(`=UNIQUE($${colName}2:$${colName})`);
    
    active.getRange("R1C"+(headerWidth + 2)).setValue("Total Votes:");
    let votes = [];
    for(let row = 2; row <= 1000; row++) {
      votes.push([`=IFERROR(COUNTIF($${colName}2:$${colName}, ${generateColumn(headerWidth + 1) + row}),"")`])
    }
    active.getRange("R2C"+(headerWidth + 2)+":R1000C"+(headerWidth + 2)).setValues(votes);

    const candidateCol = generateColumn(headerWidth + 1);
    const voteCol = generateColumn(headerWidth + 2);
    active.getRange("R1C"+(headerWidth + 3)).setValue("Percents");
    active.getRange("R2C"+(headerWidth + 3)).setValue(`=ARRAYFORMULA(${voteCol}2:${voteCol}/SUM(${voteCol}2:${voteCol}))`);
    active.getRange("R1C"+(headerWidth + 4)).setValue("Winner:");
    const pctCol = generateColumn(headerWidth + 3);
    active.getRange("R2C"+(headerWidth + 4)).setValue(`=FILTER(${candidateCol}2:${candidateCol}, ${pctCol}2:${pctCol} > 0.5)`);
  }
}

// Function to add a button to send emails
function onOpen() {
  let ui = SpreadsheetApp.getUi();

  ui.createMenu('Secure Elections')
      .addItem('Create Form Spreadsheet', 'createForm')
      .addItem('Send Emails', 'sendEmails')
      .addSubMenu(ui.createMenu("Implement Vote Counting")
          .addItem("Plurality", "implementPlurality")
          .addItem("First Past The Post", "implementFPTP"))
      .addToUi();
}
