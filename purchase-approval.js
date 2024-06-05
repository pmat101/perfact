const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
let currentRow;
  
function onFormSubmit(e) {
  currentRow = sheet.getLastRow();
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
    <p>Do you approve this request?</p>
    <a href="https://script.google.com/macros/s/AKfycbwDWgpWTt4_KqszbwfprKU5j7XYWXBOjKzpuXUSIfsBBldwLSCgqZIbTfeWK-SBLhgz/exec?action=accept">Accept</a>
    &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbwDWgpWTt4_KqszbwfprKU5j7XYWXBOjKzpuXUSIfsBBldwLSCgqZIbTfeWK-SBLhgz/exec?action=reject">Reject</a>
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
  let action = e.parameter.action;
  if(action == "accept") {
    sheet.getRange(currentRow, 10, 1, 1).setValue("accepted");
  }
  else if(action == "reject") {
    sheet.getRange(currentRow, 10, 1, 1).setValue("rejected");
  }
  return HtmlService.createHtmlOutputFromFile("th-response");
}

function thResponse(comment) {
    sheet.getRange(currentRow, 11, 1, 1).setValue(comment);
}
