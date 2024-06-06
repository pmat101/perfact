const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
function onFormSubmit(e) {
  let currentRow = sheet.getLastRow();
  let formData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
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
    <p>Request No. ${currentRow}</p>
    <p>Employee Name: ${formData[2]}</p>
    <p>Item Name: ${formData[6]}</p>
    <p>Amount: ${formData[7]}</p>
    <p>Reason: ${formData[8]}</p>
    <span>Do you approve this request?</span>
    &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbykg6__N7yK8X7S-mhq4osizcxC-dd7zk8d0kq50z2kso9y2o-Hb4VO2XjpNymYr2A4Sw/exec?send=th">Respond</a>
    <br>
    <p>Thanks & Regards</p>
  </body>
  `;
  GmailApp.sendEmail(th, subject, body, {
    htmlBody: body,
    name: name
  });
}

function thResponse(uid, say, comment) {
  sheet.getRange(uid, 10, 1, 1).setValue(say);
  sheet.getRange(uid, 11, 1, 1).setValue(comment);
  if (say == "Accepted") {
    let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
    let bh = formData[5];
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
      <p>Request No. ${uid}</p>
      <p>Employee Name: ${formData[2]}</p>
      <p>Item Name: ${formData[6]}</p>
      <p>Amount: ${formData[7]}</p>
      <p>Reason: ${formData[8]}</p>
      <p>Team Head Response: ${comment}</p>
      <span>Do you approve this request?</span>
      &nbsp; &nbsp;
    <a href="https://script.google.com/macros/s/AKfycbykg6__N7yK8X7S-mhq4osizcxC-dd7zk8d0kq50z2kso9y2o-Hb4VO2XjpNymYr2A4Sw/exec?send=bh">Respond</a>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(bh, subject, body, {
      htmlBody: body,
      name: name
    });
  }
  else if (say == "Rejected") {
    let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
    let employee = formData[3];
    let subject = "Purchase Request Rejected";
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
      <p>Your purchase request for ${formData[6]} has not been accepted by your Team Head, for below given reason.</p>
      <p>Team Head Response: ${comment}</p>
      <p>For further clarifications please consult your manager</p>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(employee, subject, body, {
      htmlBody: body,
      name: name
    });
  }
}

function bhResponse(uid, say, comment) {
  sheet.getRange(uid, 12, 1, 1).setValue(say);
  sheet.getRange(uid, 13, 1, 1).setValue(comment);
  if (say == "Accepted") {
    let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
    let recipient = "accounts@perfactgroup.in";
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
      <p>Request No. ${uid}</p>
      <p>Employee Name: ${formData[2]}</p>
      <p>Item Name: ${formData[6]}</p>
      <p>Amount: ${formData[7]}</p>
      <p>Reason: ${formData[8]}</p>
      <p>Team Head Response: ${formData[10]}</p>
      <p>Business Head Response: ${comment}</p>
      <p>Please process this request.</p>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(recipient, subject, body, {
      htmlBody: body,
      name: name
    });
  }
  else if (say == "Rejected") {
    let formData = sheet.getRange(uid, 1, 1, sheet.getLastColumn()).getValues()[0];
    let employee = formData[3];
    let subject = "Purchase Request Rejected";
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
      <p>Your purchase request for ${formData[6]} has not been accepted by your Business Head, for below given reason.</p>
      <p>Business Head Response: ${comment}</p>
      <p>For further clarifications please consult your manager</p>
      <br>
      <p>Thanks & Regards</p>
    </body>
    `;
    GmailApp.sendEmail(employee, subject, body, {
      htmlBody: body,
      name: name
    });
  }
}

function doGet(e) {
  let send = e.parameter.send;
  if(send == "th"){
    return HtmlService.createHtmlOutputFromFile("th-response");
  }
  else if(send == "bh"){
    return HtmlService.createHtmlOutputFromFile("bh-response");
  }
  else{
    console.log("oops")
  }
}
