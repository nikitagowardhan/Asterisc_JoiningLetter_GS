const DOCID = '19-qWf6rPnfdGyh_7QUCjs9fehvj6vz6fV2nSBpleuKY';
const FOLDERID = '1-HP0ZaL6i1DY4uQcHKqlU5W3yYlTyeG3';
const SHEETID = '1yJ6dAv3jkoQn4Ejj2WCoSm6GGXzXlvuZLeALQfIuPks'; // Replace with your actual Google Sheet ID

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Mail');
  menu.addItem('Generate JoiningLetter', 'generateReceipt');
  menu.addItem('Send Mail', 'sendMail');
  menu.addToUi();
}

function generateReceipt() {
  setDateInColumnE();
  const sheet = SpreadsheetApp.openById(SHEETID).getSheets()[0];
  if (!sheet) {
    console.error("Sheet not found. Aborting function.");
    return;
  }

  const temp = DriveApp.getFileById(DOCID);
  const folder = DriveApp.getFolderById(FOLDERID);

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Exclude header row

  rows.forEach((row, index) => {
    const srNo = row[0]; // Gets the serial number from the first column
    const name = row[1]; // Gets the name from the second column
    const email = row[2]; // Gets the email from the third column
    const refNo = row[3]; // Gets the reference number from the fourth column
    const startDateOfJoining = row[5]; // Gets the start date of joining from the sixth column (column F)
    const pdfLink = row[6]; // Gets the PDF link from the seventh column
    const waLink = row[7]; // Gets the WhatsApp link from the eighth column
    const verify = row[8]; // Gets the verify status from the ninth column
    const mailSend = row[9]; // Gets the mail send status from the tenth column
    const waGroupLink = row[10]; // Gets the WhatsApp Group link from the eleventh column

    if (!pdfLink && email && startDateOfJoining) { // Check if start date is provided
      try {
        // Make a copy of the template document
        const file = temp.makeCopy(folder);
        const doc = DocumentApp.openById(file.getId());
        const body = doc.getBody();

        // Replace placeholders with actual data
        const formattedStartDate = Utilities.formatDate(new Date(startDateOfJoining), Session.getScriptTimeZone(), 'dd-MM-yyyy');
        const formattedCurrentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

        body.replaceText('{{Ref No}}', refNo);
        body.replaceText('{{Current Date}}', formattedCurrentDate); // Replace with formatted current date
        body.replaceText('{{Full Name}}', name);
        body.replaceText('{{Start Date}}', formattedStartDate);

        const pdfName = `${name}_${refNo}.pdf`;
        doc.setName(pdfName);

        // Convert doc to pdf
        const blob = doc.getAs(MimeType.PDF);
        doc.saveAndClose();
        const pdf = folder.createFile(blob).setName('joiningletter_' + pdfName);

        // Set the PDF file to be shareable with anyone
        pdf.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

        // Set PDF link in the sheet
        const url = pdf.getUrl();
        sheet.getRange(index + 2, 7).setValue(url); // PDF link in the seventh column (G)

        // Generate WhatsApp link
        const customMessage = "Dear " + name + ",\n\n" +
          "We are delighted to welcome you to our company. " +
          "Your start date will be " + formattedStartDate + ". " +
          "Please review the attached joining letter for more details.\n" +
          "You can view the joining letter by clicking the following link: " + url + ".\n\n" +
          "We look forward to having you join our team and contribute to our success.\n\n" +
          "Sincerely,\n" +
          "Chandrakant Bobade,\n" +
          "Director,\n" +
          "Asterisc Team";
        const encodedMessage = encodeURIComponent(customMessage);

        const walink = `https://wa.me/?text=${encodedMessage}`;
        sheet.getRange(index + 2, 8).setValue(walink); // WhatsApp link in the eighth column (H)

        // Log and trash the temporary file
        Logger.log(row);
        file.setTrashed(true);
      } catch (error) {
        console.error("Error generating receipt for row " + (index + 2) + ": " + error.message);
      }
    }
  });
}

function setDateInColumnE() {
  const sheet = SpreadsheetApp.openById(SHEETID).getSheets()[0];
  if (!sheet) {
    console.error("Sheet not found. Aborting function.");
    return;
  }

  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  const data = sheet.getRange('B2:E').getValues(); // Adjusted to range B2:E to match the columns accurately

  data.forEach((row, index) => {
    const rowIndex = index + 2; // Adjust for header row
    const bValue = row[0]; // B column value
    const eValue = row[3]; // E column value

    if (bValue && !eValue) {
      sheet.getRange(`E${rowIndex}`).setValue(currentDate); // Set date in column E
    }
  });
}

function sendMail() {
  const sheet = SpreadsheetApp.openById(SHEETID).getSheets()[0];
  if (!sheet) {
    console.error("Sheet not found. Aborting function.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Exclude header row

  rows.forEach((row, index) => {
    const email = row[2]; // Gets the email from the third column
    const pdfLinkCell = sheet.getRange(index + 2, 7); // PDF link in the seventh column (G)

    if (row[8] && !row[9]) { // Check the verify status (column 9) and mail send status (column 10)
      try {
        const subject = 'Congratulations 🌟 on your success! We are pleased to present you with the Joining Letter';

        // Extract the PDF link from the cell formula
        const pdfLink = pdfLinkCell.getValue();

        // WhatsApp link (assuming it's in the eighth column)
        const waLink = row[7];

        // WhatsApp Group link (assuming it's in the eleventh column)
        const waGroupLink = row[10];

        // HTML mail template
        var messageHtmlBody = HtmlService.createTemplateFromFile('mail_template.html');
        messageHtmlBody.row = row;
        messageHtmlBody.WAlink = waLink;
        messageHtmlBody.PDFlink = pdfLink; // Pass the PDF link to the template

        // Evaluate the HTML template and get the content as a string
        var messageBody = messageHtmlBody.evaluate().getContent();

        // Check if the email is valid before sending
        if (isValidEmail(email)) {
          // Send email with attachments
          MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: messageBody,
          });

          // Log the email sent and update sheet with status
          Logger.log("Email sent to " + email);
          var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
          var formattedTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
          sheet.getRange(index + 2, 10).setValue("Mail Sent " + formattedDate + " - " + formattedTime);
        } else {
          Logger.log("Invalid email address: " + email);
          sheet.getRange(index + 2, 10).setValue("Invalid Email");
        }
      } catch (error) {
        console.error("Error sending email for row " + (index + 2) + ": " + error.message);
      }
    }
  });
}

// Validate email format
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
