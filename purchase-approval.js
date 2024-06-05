const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
function onFormSubmit(e) {
  let currentRow = sheet.getLastRow();
  let formData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  // let uniqueID = formData[0].toString();
  let th = formData[4];
  let subject = "New Office Essentials Purchase Request";
  let name = "Purchase Request";
  let body = `
  <head>
    <style>
      a {
        text-decoration: none;
        font-size: 1.1em;
        padding: 0.5em 1em;
        border: 1px solid #888;
        border-radius: 1em;
      }
    </style>
  </head>
  <body>
    <p>A new purchase request has been submitted.</p>
    <p>Employee Name: ${formData[2]}</p>
    <p>Item Name: ${formData[6]}</p>
    <p>Amount: ${formData[7]}</p>
    <p>Reason: ${formData[8]}</p>
    <span>Do you approve this request?</span>
    &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbxfsmI1wQcn79WR74FuKazJ6pHJZzq1_G0VrejpSbDlVDFJdiV8HWUwiwXZxGv5Vg3v/exec?row=${currentRow}">Respond</a>
    <br>
    <p>Thanks & Regards</p>
  </body>
  `;
  GmailApp.sendEmail(th, subject, body, {
    htmlBody: body,
    name: name
  });
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("th-response");
}

function thResponse(rowUI, say, comment) {
  sheet.getRange(rowUI, 10, 1, 1).setValue(say);
  sheet.getRange(rowUI, 11, 1, 1).setValue(comment);
}
