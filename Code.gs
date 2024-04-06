const SPREADSHEET_ID  = ''; // id of the sheet that contains the attendees data
const CERTS_FOLDER_ID = ''; // id of the folder in which certificates will be stored
const CERT_SLIDE_ID   = ''; // id of the google slide certificate

const CHECKBOX_PLACEHOLDER = ''; // the placeholder text of the checkbox we want to change. e.g. -> "name here".

/*
things to consider:
  - this code generates certificates per the 0-th column index (namely the "name" column), you may need to change that.
  - this code, by default, reads the active sheet (first sheet) in the Spreadsheet file, you may need to change that.
*/

function main() {
  var spreadSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var page = spreadSheet.getActiveSheet(); // or by using the sheet name (getSheetByName) or sheet index (getSheets()[indx])
  var data = page.getDataRange().getValues();

  var folder = DriveApp.getFolderById(CERTS_FOLDER_ID);

  var cert_slide = DriveApp.getFileById(CERT_SLIDE_ID);
  var name_textbox_idx = findTextBoxByName(
    SlidesApp
      .openById(cert_slide.getId())
      .getSlides()[0], CHECKBOX_PLACEHOLDER
    );

  
 for (let i = 1; i < data.length; i++) {
    //grabs the names of the participant: i-th row, 0-th col
    var name = prettifyName(data[i][0]);

    //makes a temporary copy of the google slide and store it the drive folder.
    var cert = cert_slide.makeCopy(cert_slide.getName() + "_" + i, )

    //opens the new copy we made
    var slide_doc = SlidesApp.openById(cert.getId());
    var slide     = slide_doc.getSlides()[0];

    // change the textbox value
    slide.getShapes()[name_textbox_idx]
      .getText()
      .setText(name);
    
    slide_doc.saveAndClose();

    // create a pdf out of the slide and change its name
    let pdf = cert.getBlob().getAs("application/pdf");
    pdf.setName(name.trim() +".pdf");

    // store the pdf in our drive folder
    folder.createFile(pdf);

    // remove temporary the cert.slide
    cert.setTrashed(true);

    console.log(i, name);
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

