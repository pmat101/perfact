
function sendApprovalEmail(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let lastRow = sheet.getLastRow();
  let formData = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var email = formData[4];
  var subject = "New Office Essentials Purchase Request";
  var body = `
    <body>
      <p>A new purchase request has been submitted.</p>
      <br>
      <p>Employee Name: ${formData[2]}</p>
      <p>Item Name: ${formData[6]}</p>
      <p>Purchase Amount: ${formData[7]}</p>
      <p>Reason: ${formData[8]}</p>
      <br>
      <p>Do you approve this request?</p>
      <a href="https://script.google.com/a/macros/perfactgroup.in/s/AKfycbyPe1nuXXW72jfh_dcb2E4Fn_7M8LEujO-VwoDiTbuz-RdzSyO3k09hFhrrZg94JHsK/exec?action=approve&row=${lastRow}">Approve</a>
      </body>`;
  GmailApp.sendEmail(email, subject, body, {htmlBody: body});
}
