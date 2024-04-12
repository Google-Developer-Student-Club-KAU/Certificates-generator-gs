const SPREADSHEET_ID   = ''; // id of the sheet that contains the attendees data
const CERTS_FOLDER_ID  = ''; // id of the folder in which certificates will be stored
const CERT_SLIDE_ID    = ''; // id of the google slide certificate

const CHECKBOX_PLACEHOLDER    = 'name here'; // the placeholder text of the checkbox we want to change. e.g. -> "name here".
const CHECKBOX_MAX_CHARACTERS = 22; // manually test and specify the max number of characters the textbox can fit without changing its shape
const CHECKBOX_FONT_SIZE      = 65; // font size of the name checkbox

const NAME_COL  = 2; // the name column index in the sheet. starts from zero.
const EMAIL_COL = 4; // the email column index in the sheet. starts from zero.

/*
things to consider:
  - this code, by default, reads the active sheet (first sheet) in the Spreadsheet file
*/

function main() {
  const spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const page = spreadSheet.getActiveSheet(); // or by using the sheet name or sheet index
  const data = page.getDataRange().getValues();

  const folder = DriveApp.getFolderById(CERTS_FOLDER_ID);

  const cert_template_file   = DriveApp.getFileById(CERT_SLIDE_ID);
  const cert_template_slide  = SlidesApp.openById(CERT_SLIDE_ID).getSlides()[0];

  const name_textbox_idx = findTextBoxByName(
    cert_template_slide, CHECKBOX_PLACEHOLDER
  );
  const textbox_font_size = cert_template_slide
    .getShapes()[name_textbox_idx]
    .getText()
    .getTextStyle()
    .getFontSize();

  
  for (let i = 1; i < data.length; i++) {
    //grabs the names of the participant: i-th row, 0-th col
    var name  = prettifyName(data[i][NAME_COL]);
    var email = data[i][EMAIL_COL].trim();
    Logger.log('[%d] Processing | %s - %s', i, name, email)

    //makes a temporary copy of the google slide and store it the drive folder.
    var cert = cert_template_file.makeCopy(i + "_" + cert_template_file.getName(), folder)

    // opens the new copy we made
    var slide_doc = SlidesApp.openById(cert.getId());
    var slide     = slide_doc.getSlides()[0];

    // declare the textbox and write the attendee name
    var textbox = slide.getShapes()[name_textbox_idx];
    writeName(textbox, name, textbox_font_size)

    slide_doc.saveAndClose();
    
    // create a pdf out of the slide and change its name
    let pdf = cert.getBlob().getAs("application/pdf");
    pdf.setName(name.trim() +".pdf");

    // store the pdf in our drive folder and store the url;
    var cert_pdf = folder.createFile(pdf);
    cert_pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // remove temporary the cert.slide
    cert.setTrashed(true);

    // send email
    // var line1 = "Please find attached your certificate for the ....";
    // sendEmail(email, '<course name> Certificate', line1, 'All the best to you in your future endeavors!', cert_pdf.getUrl());

    Logger.log('[%d] Done | %s - %s', i, name, email);
  }
}


// util funcs
function prettifyName(text) {
  const words = text.toLowerCase().trim().split(/\s+/); // match one -or more- spaces

  for (let i = 0; i < words.length; i++) {
    words[i] = words[i].charAt(0).toUpperCase() + words[i].slice(1);
  }

  return words.join(' ');
}

function writeName(textbox, name) {
  textbox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
  var maxWidth  = CHECKBOX_FONT_SIZE * (CHECKBOX_MAX_CHARACTERS-1); // minus one just in case

  var nameLength = name.length;
  var newFontSize = Math.min(Math.floor(maxWidth/nameLength) - 3, CHECKBOX_FONT_SIZE);
  
  Logger.log('Name-Length: [%d], Adjusted-Font-Size: [%d]', nameLength, newFontSize);

  textbox.getText()
    .setText(name)
    .getTextStyle().setFontSize(newFontSize);
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

function sendEmail(to, title, subtitle, line1, line2, link) {
  var email_footer = DriveApp.getFileById("1IPRPDn7fItnvERuzQF1lCjm_qG1-85m5").getAs("image/png"); // keep as is
  var email_header = DriveApp.getFileById("1wGUV62JfmTGvmN7wKXC8Jw1pi2_h7xsX").getAs("image/png"); // keep as is
  let emailImages = {"email_header": email_header, "email_footer": email_footer};
  let emailBody = HtmlService.createTemplateFromFile("template");
  emailBody.title = title;
  emailBody.subtitle = subtitle;
  emailBody.body_line1 = line1;
  emailBody.link = link;
  emailBody.body_line2 = line2;

  const htmlBody = emailBody.evaluate().getContent();

  MailApp.sendEmail(
    {
      to: to,
      subject: title,
      htmlBody: htmlBody,
      inlineImages: emailImages
    }
  )
}

