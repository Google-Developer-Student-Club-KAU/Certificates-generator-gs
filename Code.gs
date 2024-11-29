const SPREADSHEET_ID   = ''; // id of the sheet that contains the attendees data
const SHEETNAME = ''; // name of sheet

const CERTS_FOLDER_ID  = ''; // id of the folder in which certificates will be stored
const CERT_SLIDE_ID    = ''; // id of the google slide certificate

const TEXTBOX_PLACEHOLDER    = 'name here'; // the placeholder text of the textbox we want to change. e.g. -> "name here".
const TEXTBOX_MAX_CHARACTERS = 22; // manually test and specify the max number of characters the textbox can fit without changing its shape

const NAME_COL  = 2; // the name column index in the sheet. starts from zero.
const EMAIL_COL = 4; // the email column index in the sheet. starts from zero.

/*
things to consider:
  - this code, by default, reads the active sheet (first sheet) in the Spreadsheet file
*/

function main() {
  var data = getSpreadSheetData(SPREADSHEET_ID);


  var folder = DriveApp.getFolderById(CERTS_FOLDER_ID);
  var cert_slide = DriveApp.getFileById(CERT_SLIDE_ID);


  var name_textbox_idx = findTextBoxByName(
    SlidesApp
      .openById(cert_slide.getId())
      .getSlides()[0], CHECKBOX_PLACEHOLDER
    );

  
 for (let i = 1; i < data.length; i++) {
    var name = prettifyName(data[i][NAME_COL]);
    var email = prettifyName(data[i][EMAIL_COL]);
    Logger.log('[%d] Processing | %s - %s', i, name, email)

    //makes a temporary copy of the google slide and store it the drive folder.
    var cert = cert_slide.makeCopy(cert_slide.getName() + "_" + i, )

    //opens the new copy we made
    var slide_doc = SlidesApp.openById(cert.getId());
    var slide     = slide_doc.getSlides()[0];

    // change the textbox value
    var textbox = slide.getShapes()[name_textbox_idx];
    adjustNameForCertificate(textbox, name);
     
    slide_doc.saveAndClose();

    // create a pdf out of the slide and change its name
    let pdf = cert.getBlob().getAs("application/pdf");
    pdf.setName(name.trim() +".pdf");

    // store the pdf in our drive folder and store the url;
    var cert_file = folder.createFile(pdf);
    cert_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // remove temporary the cert.slide
    cert.setTrashed(true);

    var line1 = `Hey ${name}, Thank you for your attendance...etc`;

    sendEmail(email, 'Tech Talk Certificates', line1, 'All the best to you in your future endeavors.', cert_file.getUrl());
    Logger.log('[%d] Done | %s - %s', i, name, email);
  }
}


// util funcs
function prettifyName(text) {
  let words = text.toLowerCase().trim().split(/\s+/); // match one -or more- spaces
  
  if (words.length === 4) {
    words.splice(2, 1); // remove the third name
  }
  
  // capitalize the first letter of each name
  words = words.map(word => word.charAt(0).toUpperCase() + word.slice(1));
  
  return words.join(' ');
}



function findTextBoxByName(slide, targetText) {
  const textBoxes = slide.getShapes();

  for (let i = 0; i < textBoxes.length; i++) {
    const textBox = textBoxes[i];

    // look for textboxes only
    if (textBox.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
      if (textBox.getText().asString().trim() === targetText) {
        return i
      }
    }
  }

  // If no matching textbox was found
  return null;
}

function adjustNameForCertificate(textbox, name, maxFontSize = 65) {
  let fontSize = maxFontSize;
  const text = textbox.getText();
  const minFontSize = 35; 

  fontSize = 45.5;

  while (textbox.getText().getRuns().length > 1 && fontSize > minFontSize) {
    fontSize -= 1;
    text.getTextStyle().setFontSize(fontSize);
  }

  text.setText(name);
}

function getSpreadSheetData(spreadsheetId) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(SHEETNAME);
  var dataRange = sheet.getDataRange();
  return dataRange.getValues()
}



function sendEmail(to, name, title, line1, line2, link) {
  let emailImages = {"email_header": EMAIL_HEADER};
  let emailBody = HtmlService.createTemplateFromFile("Template");

  emailBody.title = title; 
  emailBody.body_line1 = line1;
  emailBody.body_line2 = line2;
  emailBody.name = name;
  emailBody.link = link;

  const htmlBody = emailBody.evaluate().getContent();

  MailApp.sendEmail({
    to: to,
    subject: title,
    htmlBody: htmlBody,
    inlineImages: emailImages
  });
}



