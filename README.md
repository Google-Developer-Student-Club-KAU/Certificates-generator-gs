# Certificates-generator-gs
A google apps script that generates certificates from a spreadsheet of attendee data and a Google Slides certificate template (exported from canva).

## Setup:
Note: you must export the certificate template as `.pttx` and then open it and save it with Google Slides.
  - Create a copy of this script: Open Google Apps Script (https://script.google.com) and paste the provided code into a new project.
  - Export the certificate from Canva to Google Drive as `.pttx`:
    - Open the file using Google Slides.
    - **THEN** go to *File* > *Save as Google Slides*.
    - That's it, grap the id of the file to use it in the code.

## Replace placeholders:
`SPREADSHEET_ID`: Id of the sheet that contains the attendees data.<br>
`CERTS_FOLDER_ID`: Id of the folder in which certificates will be stored.<br>
`CERT_SLIDE_ID`: Id of the google slide certificate.<br>
`CHECKBOX_PLACEHOLDER`: The exact text of the placeholder on the certificate template where the name should go.<br>
`NAME_COL` & `EMAIL_COL`: The name & email column index in the sheet. starts from zero.

## Usage:
Run the script. The script will generate a PDF certificate for each attendee listed in your spreadsheet and place them in the Drive folder.



### Additional Notes:
The script assumes the `name` & `email` data are read from the 2nd and 4th columns of your spreadsheet, respectively. Make sure to change their indexes in the code if needed.

The script processes data from the active sheet of your spreadsheet. Modify it if you need it to use a specific sheet (see the code for more hints).
