// Column letters to numbers
function colLetterToNumber(letter) {
  let num = 0;
  letter = letter.toUpperCase();
  for (let i = 0; i < letter.length; i++) {
    num = num * 26 + (letter.charCodeAt(i) - 64);
  }
  return num;
}

// === adding a checkkbox thing
function insertCheckbox(range, value) {   
  range.setDataValidation( SpreadsheetApp.newDataValidation().requireCheckbox().build() ); range.setValue(value); 
}

// ==== FORM SUBMIT: email + prepare checkbox on col S ====
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet   = ss.getSheetByName('EngSoc Finance Form');
  const ledgerSheet = ss.getSheetByName('General Ledger');

  const row = e.range.getRow();      // the row that was inserted
  const v   = e.values;              // values from the form row

  const timestamp = v[colLetterToNumber("A") - 1];
  const email = v[colLetterToNumber("B") - 1];
  const fullName = v[colLetterToNumber("C") - 1];
  const instructions = v[colLetterToNumber("D") - 1];
  const portfolio = v[colLetterToNumber("E") - 1];
  const request = v[colLetterToNumber("F") - 1];
  const eventName = v[colLetterToNumber("G") - 1];
  const category = v[colLetterToNumber("H") - 1];
  const moneyStr = v[colLetterToNumber("I") - 1];
  const receipt = v[colLetterToNumber("J") - 1];
  const address = v[colLetterToNumber("K") - 1];
  let cheque = v[colLetterToNumber("L") - 1];
  const vendorName = v[colLetterToNumber("M") - 1];
  const contact = v[colLetterToNumber("N") - 1];
  const phoneNumber = v[colLetterToNumber("O") - 1];
  const vendorEmail = v[colLetterToNumber("P") - 1];
  const voidCheque = v[colLetterToNumber("Q") - 1];
  const invoice = v[colLetterToNumber("R") - 1];
  const studentID = v[colLetterToNumber("S") - 1];
  const pr_itemReq = v[colLetterToNumber("T") - 1];
  const pr_invoiceQuote = v[colLetterToNumber("U") - 1];
  const vpfa_checked = v[colLetterToNumber("V") - 1]; // checkbox is added
  const vpfa_date = v[colLetterToNumber("W") - 1]; // onEdit, date is added when checked for vpfa_checked

  const money = Number(moneyStr) || 0;

  // ----- Email (text + HTML) -----
  const subject = "Budget Update: " + fullName;

  let textMessage =
    "A new Finance Form has been submitted.\n\n" +
    "Name: " + fullName +
    "\nEmail: " + email +
    "\nInstructed by: " + instructions +
    "\nFinance Request: " + request +
    "\nPortfolio: " + portfolio +
    "\nEvent Name: " + (eventName || "") +
    "\nTransaction Pended: $" + money.toFixed(2) +
    "\nCategory: " + (category || "") +
    "\nStudent ID: " + studentID;

  if (address && cheque) {
    textMessage += "\nAddress: " + address;
    if (cheque === "Yes") {
      textMessage += "\nCheque status: Sent to you";
      cheque = "Sent to you";
    } else {
      textMessage += "\nCheque status: Will send to you";
      cheque = "Will send to you";
    }
  }

  if (voidCheque && invoice) {
    textMessage +=
      "\nVendor Name: " + (vendorName || "") +
      "\nPrimary Contact: " + (contact || "") +
      "\nPrimary Phone Number: " + (phoneNumber || "") +
      "\nVendor Email: " + (vendorEmail || "") +
      "\nVoice cheque and invoice Status: Will send to you";
  } else { 
    textMessage += "\nNot a third-party payment form.";
  }

  if (pr_itemReq && pr_invoiceQuote) {
    textMessage +=
    "\nItem Requested: " + (pr_itemReq || "") + 
    "\nInvoice or Quote: " + (pr_invoiceQuote || "");
  }

  textMessage += "\n\nPlease review it in the spreadsheet and check it off once done!\n" + ss.getUrl();

  const htmlMessage = `
    <p>A new <b>Finance Form</b> has been submitted.</p>
    <table border="0" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
      <tr><td><b>Person who submitted:</b></td><td>${fullName}</td></tr>
      <tr><td><b>Email:</b></td><td>${email}</td></tr>
      <tr><td><b>Instructed by:</b></td><td>${instructions}</td></tr>
      <tr><td><b>Finance Request:</b></td><td>${request}</td></tr>
      <tr><td><b>Portfolio:</b></td><td>${portfolio}</td></tr>
      <tr><td><b>Event Name:</b></td><td>${eventName || ""}</td></tr>
      ${address && cheque ? `<tr><td><b>Student ID:</b></td><td>${studentID}</td></tr>` : ""}
      <tr><td><b>Transaction Pended:</b></td><td><b>$${money.toFixed(2)}</b></td></tr>
      <tr><td><b>Category:</b></td><td>${category || ""}</td></tr>
      ${address && cheque ? `<tr><td><b>Address:</b></td><td>${address}</td></tr>` : ""}
      ${address && cheque ? `<tr><td><b>Cheque Status:</b></td><td>${cheque}</td></tr>` : ""}
      ${vendorEmail && contact ? `<tr><td><b>Vendor Name:</b></td><td>${vendorName}</td></tr>` : ""}
      ${vendorEmail && contact ? `<tr><td><b>Primary Contact Name:</b></td><td>${contact}</td></tr>` : ""}
      ${vendorEmail && contact ? `<tr><td><b>Vendor Email:</b></td><td>${vendorEmail}</td></tr>` : ""}
      ${vendorEmail && contact ? `<tr><td><b>Void Cheque and Invoice Status:</b></td><td>Will send to me :)</td></tr>` : ""}
    ${pr_invoiceQuote && pr_itemReq ? `<tr><td><b>Item Requested:</b></td><td>${pr_itemReq}</td></tr>` : ""}
    ${pr_invoiceQuote && pr_itemReq ? `<tr><td><b>Invoice/Quote Status:</b></td></td>Will send to me :)</td></tr>` : ""}
    </table>
    <p>Please review it in the spreadsheet and check it off once done!<br>
    <a href="${ss.getUrl()}">Open the spreadsheet</a></p>
  `;

  MailApp.sendEmail({
    to: "",
    subject,
    body: textMessage,
    htmlBody: htmlMessage
  });

  // Prep VPFA checkbox on V or THIS form row 
  ensureColumnExists(formSheet, colLetterToNumber("V"));         // V
  const vCell = formSheet.getRange(row, colLetterToNumber("V")); // VPFA checked
  vCell.insertCheckboxes();
  vCell.setValue(false);
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Only respond to edits on the Finance Form sheet
  if (sheet.getName() !== "EngSoc Finance Form") return;

  // Only respond when VPFA checkbox (column V = 22) is checked TRUE
  if (col !== 22 || String(e.value).toUpperCase() !== "TRUE") return;

  // Read the entire edited form row
  const lastCol = sheet.getLastColumn();
  const r = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const timestamp = r[colLetterToNumber("A") - 1];
  const fullName  = r[colLetterToNumber("C") - 1];
  const portfolio = r[colLetterToNumber("E") - 1];
  const request   = r[colLetterToNumber("F") - 1];
  let   eventName = r[colLetterToNumber("G") - 1] || "";
  const category  = r[colLetterToNumber("H") - 1];
  const money     = Number(r[colLetterToNumber("I") - 1]) || 0;
  const description = r[colLetterToNumber("G") - 1] || "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // CASE 1: UPDATE GENERAL LEDGER

  const ledgerSheet = ss.getSheetByName("General Ledger");
  const nextLedgerRow = getNextAvailableRow(ledgerSheet);

  // Stamp W with timestamp
  ensureColumnExists(sheet, colLetterToNumber("W"));
  sheet.getRange(row, colLetterToNumber("W")).setValue(new Date());

  // Format date
  const date = new Date(timestamp);
  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d/yyyy');

  // Negative spending
  const negativeAmount = -Math.abs(money);

  // Running balance formula
  // const balanceCell = ledgerSheet.getRange(nextLedgerRow, colLetterToNumber("J"));
  // balanceCell.setFormula(`=SUM(J${nextLedgerRow - 1}, I${nextLedgerRow})`);

  // Normalize type of form
  let typeForm = eventName.toLowerCase();
  if (typeForm.includes("event")) typeForm = "Reimbursement";
  else if (typeForm.includes("misc")) typeForm = "Reimbursement";
  else if (typeForm.includes("third")) typeForm = "3rd Party Payment";
  else if (typeForm.includes("purchase")) typeForm = "Purchase Req.";
  else typeForm = "Conference";

  // Write row to General Ledger
  ledgerSheet.getRange(nextLedgerRow, 1).setValue(formattedDate);
  ledgerSheet.getRange(nextLedgerRow, 2).setValue(request);
  ledgerSheet.getRange(nextLedgerRow, 3).setValue(portfolio);
  ledgerSheet.getRange(nextLedgerRow, 4).setValue(category);
  ledgerSheet.getRange(nextLedgerRow, 5).setValue(`Submitted by: ${fullName} - ${eventName}`);
  ledgerSheet.getRange(nextLedgerRow, 9).setValue(10000); // added trans. no.
  ledgerSheet.getRange(nextLedgerRow, 10).setValue(negativeAmount);
  ledgerSheet.getRange(nextLedgerRow, 11).setValue(`=SUM(J${nextLedgerRow - 1}, I${nextLedgerRow})`);
  ledgerSheet.getRange(nextLedgerRow, 12).setValue("Form");
  ledgerSheet.getRange(nextLedgerRow, 13).setValue(typeForm);
  insertCheckbox(ledgerSheet.getRange(nextLedgerRow, 14), false);
  ledgerSheet.getRange(nextLedgerRow, 15).setValue("Yes");

  ledgerSheet.getRange(nextLedgerRow, 1, 1, 14).setFontFamily("Ubuntu").setFontSize(10);

  // CASE 2: UPDATE PORTFOLIO TRACKER

  const tracker = ss.getSheetByName(portfolio);
  if (!tracker) {
    SpreadsheetApp.getUi().alert("No sheet found for portfolio: " + portfolio);
    return;
  }

  const nextTrackerRow = getNextAvailableRow(tracker);

  // Normalize event type for tracker
  const low = eventName.toLowerCase();
  if (low.includes("event")) eventName = "Event";
  else if (low.includes("misc")) eventName = "Misc.";
  else if (low.includes("third")) eventName = "3rd Party Payment";
  else if (low.includes("purchase")) eventName = "Purchase Req.";
  else eventName = "Conference";

  const firstName = (fullName || "").split(" ")[0] || "";

  tracker.appendRow([
    timestamp,                     // A Date
    "edit event",                  // B Purpose
    eventName,                     // C Category
    "",                            // D Estimate
    -Math.abs(money),              // E Actual
    `=SUM(F${nextTrackerRow - 1}, E${nextTrackerRow})`, // F Balance
    `${firstName}: ${description}`, // G Description
    new Date()                     // H Reviewed At
  ]);

  tracker.getRange(nextTrackerRow, 1, 1, 8).setFontFamily("Ubuntu").setFontSize(10);
}


// ==== helpers ====
function ensureColumnExists(sheet, col) {
  const max = sheet.getMaxColumns();
  if (col > max) sheet.insertColumnsAfter(max, col - max);
}

function getPreviousBalance(sheet, currentRow) {
  for (let row = currentRow - 1; row >= 3; row--) {
    const v = sheet.getRange(row, 3).getValue(); // Column C (Balance)
    if (v !== null && v !== '') return Number(v) || 0;
  }
  return 0;
}

function getNextAvailableRow(sheet) {
  const lastRow = sheet.getLastRow();
  // Search from row 5 for the first empty slot; fall back to lastRow+1
  for (let row = 5; row <= lastRow + 10; row++) {
    const a = sheet.getRange(row, 1).getValue(); // Date
    const d = sheet.getRange(row, 4).getValue(); // Amount
    const e = sheet.getRange(row, 5).getValue(); // Portfolio
    if ((!a && a !== 0) && (!d && d !== 0) && (!e && e !== 0)) return row;
  }
  return lastRow + 1;
}
