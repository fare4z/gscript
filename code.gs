function onSubmit(e) {
  // Get the form response data
  const responses = e.namedValues;
 
  // Get the active Google Sheet and the responses sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responsesSheet = ss.getSheetByName('Form Responses 1'); // Assuming your form responses are in 'Form Responses 1'

  // Find the last row with data and the current row
  const lastRow = responsesSheet.getLastRow();
  const currentRow = responsesSheet.getActiveRange().getRowIndex();

  // Set the default status to "Pending"
  responsesSheet.getRange('E' + currentRow).setValue('Pending'); // Assuming 'Status' is in column E
}

function generateCertificatesFromSlides() {
  // 1. Get the Google Sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // 2. Define the column headers
  const headers = data[0];
  const nameColumnIndex = headers.indexOf("Nama Penuh");
  const emailColumnIndex = headers.indexOf("Email");
  const organizationColumnIndex = headers.indexOf("Jabatan");
  const statusColumnIndex = headers.indexOf("Status");
  const icColumnIndex = headers.indexOf("No KP");
  const runningNumberColumnIndex = headers.indexOf("NO_SIRI"); // Add this line

  // 3. Load certificate template (replace with your Slides template ID)
  const certificateTemplateId = "replace with ID";  // <--- IMPORTANT: Replace with your Slides template ID
  const certificateTemplate = DriveApp.getFileById(certificateTemplateId);

  // 4. Create a folder for generated certificates
  const outputFolderId = "replace with folder ID";  // <--- IMPORTANT: Replace with your output folder ID
  let outputFolder;
  if (outputFolderId) {
    outputFolder = DriveApp.getFolderById(outputFolderId);
  } else {
    outputFolder = DriveApp.createFolder("Generated Certificates");
  }

  // 5. Get the last running number from the sheet and increment it
  let lastRunningNumber = 0;
  if (runningNumberColumnIndex > -1) { // Check if the column exists
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const runningNumber = row[runningNumberColumnIndex];
      if (runningNumber) {
        const numberPart = parseInt(runningNumber.split("-")[1], 10);
        if (!isNaN(numberPart) && numberPart > lastRunningNumber) {
          lastRunningNumber = numberPart;
        }
      }
    }
  }
  let currentRunningNumber = lastRunningNumber + 1;

  // 6. Loop through the rows and generate certificates for approved attendees
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusColumnIndex];

    console.log(status);

    if (status === "Approved") {
      const name = row[nameColumnIndex];
      const email = row[emailColumnIndex];
      const organization = row[organizationColumnIndex];
      const nokp = row[icColumnIndex];
   

      // 7. Generate the running number string
      const currentDate = new Date();
      const year = currentDate.getFullYear();
      const month = String(currentDate.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
      const runningNumberString = `${year}${month}-${String(currentRunningNumber).padStart(3, '0')}`;

      // 8. Create a copy of the template slide
      const newCertificateFile = certificateTemplate.makeCopy(`${name} Certificate`, outputFolder);
      const newCertificateId = newCertificateFile.getId();
      const presentation = SlidesApp.openById(newCertificateId);
      const slide = presentation.getSlides()[0]; // Get the first slide

      // 9. Replace placeholders in the slide with attendee data
      replaceAllTextInSlide(slide, "{{Nama Penuh}}", name);
      replaceAllTextInSlide(slide, "{{No KP}}", nokp);
      replaceAllTextInSlide(slide, "{{no_siri}}", runningNumberString); // Add this line

      // 10. Save the presentation
      presentation.saveAndClose();

      // 11. Convert the presentation to PDF
      const pdfFile = newCertificateFile.getAs(MimeType.PDF);

      // 12. Send the certificate via email
      MailApp.sendEmail({
        to: email,
        subject: "Your Certificate of Attendance",
        body: "Please find your certificate attached. Your Certificate Number is " + runningNumberString, // Include running number in email
        attachments: [pdfFile],
      });

      // 13. (Optional) Delete the temporary Google Slides file (keep the PDF)
      DriveApp.getFileById(newCertificateId).setTrashed(true);

      // 14. Update the sheet
      if (runningNumberColumnIndex > -1) { // Check if the column exists
        sheet.getRange(i + 1, runningNumberColumnIndex + 1).setValue(runningNumberString); // Write running number to sheet
      }
      sheet.getRange(i + 1, statusColumnIndex + 1).setValue("Certificate Generated");
      SpreadsheetApp.flush();

      currentRunningNumber++; // Increment for the next certificate
    }
  }
  Logger.log("Certificates generated and sent.");
}

/**
 * Replaces all occurrences of a placeholder text in a slide.
 * @param {Slide} slide The slide to modify.
 * @param {string} placeholder The text to replace.
 * @param {string} value The replacement text.
 */
function replaceAllTextInSlide(slide, placeholder, value) {
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    if (shape.getText()) {
      shape.getText().replaceAllText(placeholder, value);
    }
  }
}