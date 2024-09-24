function dailyCheck() {
  ScriptApp.newTrigger("summary")
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
  .atHour(14)
  .nearMinute(0)
  .create(); 
}

function summary () {
  let docFile = DocumentApp.getActiveDocument()
  let docBody = docFile.getBody();       // now take each line and put in doc file and then convert to PDF
  docBody.clear();
  multi("Pool", "#990000", docBody);
  multi("Cove", "#990000", docBody);
  // multi("Ocean", "#b45f06", docBody);
  multi("Pond", "#0b5394", docBody);
  multi("Estuary", "#38761d", docBody);
  multi("Canal", "#38761d", docBody);
  // multi("Tributary", "#351c75", docBody);
  // multi("Delta", "#351c75", docBody);
  reservoir (docBody);
  glacier (docBody);
  let today = new Date();
  let str = today.getDate();
  str = str + "/" + (today.getMonth()+1) + "/" + today.getFullYear();
  let pdfFile = docFile.getAs('application/pdf').setName(`WPF-report-${str}`);
  const recipient = "topmanagement@perfactgroup.in";
  const subject = `Weekly Performance Update - [${str}]`;
  const cc = "it.council@perfactgroup.in";
  const name = "IT/ COUNCIL/ PERFACT";
  const alias = GmailApp.getAliases();
  let body = `
    <head></head>
    <body>
      <p>Dear Top Management,</p>
      <p>Please find attached the latest weekly performance report, consolidating data from all team WPFs submitted for the week ending [${str}].</p>
      <p>This report provides a comprehensive overview of team progress, key achievements, and any identified challenges.</p>
      <p>We believe this data will be valuable in tracking performance and making informed decisions.</p>
      <p>Please let us know if you have any questions or require further details.</p>
      <p>Thank you for your valuable input and support.</p>
      <br>
      <p>--------------------------</p>
      <p>Thanks & Regards</p>
      <br>
    </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    attachments: [pdfFile],
    from: alias[0],
    cc: cc
  });
}

function entities (entity) {
  const arr = ["&#x2f;", "/", "&rdquo;", '"', "&ldquo;", '"', "&#x28;", "(", "&#x29;", ")", "&amp;", "&", "&#x2b;", "+", "&#x3d;", "=", "&#x25;", "%", "&#x27;", '"', "&#x3a;", ":", "&#x3b;", ";"];
  for(let i=0; i<arr.length; i=i+2) {
    while(entity.match(arr[i]) != null) {
      let e1 = entity.match(arr[i]);
      entity = entity.replace(e1[0], arr[i+1]);
    }
  }
  return entity;
}

function glacier (docBody) {
  let msg = GmailApp.search(`subject:(Team Performance of Glacier for the week)`, 0, 1)[0].getMessages()[0].getBody();
  msg = msg.replace(`<div><span class="colour" style="color: rgb(0, 0, 0)"><span class="font" style="font-family: Verdana, arial, Helvetica, sans-serif"><span class="size" style="font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; white-space: normal; text-decoration-color: initial; float: none">Dear sir,<br /><br /><span class="colour" style="color: rgb(0, 0, 0)"><span class="font" style="font-family: Verdana, arial, Helvetica, sans-serif"><span class="size" style="font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; white-space: normal; text-decoration-color: initial; float: none">`, "");
  let regEx = new RegExp(`${'has been submitted successfully for your review'}(.*?)${'<tr><td valign="top" style="padding-top: 3px;">Number of BDs filled<'}`);
  msg = msg.replace(regEx, '<table><tr><td valign="top" style="padding-top: 3px;">Number of BDs filled<');
  regEx = new RegExp(`${'<tr><td valign="top">Upload Weekly PPT<'}(.*?)${'Thanks and Regards<br /><br /></div><br />'}`);
  msg = msg.replace(regEx, "</table>");
  msg = msg.replace(`&nbsp;`, " ");
  regEx = new RegExp(`${'style="'}(.*?)${'"'}`);
  let flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  regEx = new RegExp(`${' valign="top'}(.*?)${'"'}`);
  flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  regEx = new RegExp(`${'Team Performance of'}(.*?)${'\\)'}`);
  let str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str[0] = str[0].replace(`&nbsp;`, " ");
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#434343";
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  str[0] = entities(str[0]);
  docBody.appendParagraph(str[0].toString()).setAttributes(style);
  docBody.appendParagraph("\r\r");
  msg = msg.replace("<table>", "");
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  let headings = msg.match(regEx);
  let countCols = headings[0].split("<th >");
  let cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  let values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  let countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  let table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  headings = msg.match(regEx);
  countCols = headings[0].split("<th >");
  cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  headings = msg.match(regEx);
  countCols = headings[0].split("<th >");
  cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  let subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i])
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendHorizontalRule()
  docBody.appendParagraph("\r\r");;
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
}

function reservoir (docBody) {
  let msg = GmailApp.search(`subject:(Team Performance of Reservoir for the week)`, 0, 1)[0].getMessages()[0].getBody();
  msg = msg.replace(`<div><span class="colour" style="color: rgb(0, 0, 0)"><span class="font" style="font-family: Verdana, arial, Helvetica, sans-serif"><span class="size" style="font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; white-space: normal; text-decoration-color: initial; float: none">Dear sir,<br /><br /><span class="colour" style="color: rgb(0, 0, 0)"><span class="font" style="font-family: Verdana, arial, Helvetica, sans-serif"><span class="size" style="font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; white-space: normal; text-decoration-color: initial; float: none">`, "");
  let regEx = new RegExp(`${'has been submitted successfully for your review'}(.*?)${'<tr><td valign="top" style="padding-top: 3px;">Number of TFs filled<'}`);
  msg = msg.replace(regEx, '<table><tr><td valign="top" style="padding-top: 3px;">Number of TFs filled<');
  regEx = new RegExp(`${'<tr><td valign="top">Upload Weekly PPT<'}(.*?)${'Thanks and Regards<br /><br /></div><br />'}`);
  msg = msg.replace(regEx, "</table>");
  msg = msg.replace(`&nbsp;`, " ");
  regEx = new RegExp(`${'style="'}(.*?)${'"'}`);
  let flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  regEx = new RegExp(`${' valign="top'}(.*?)${'"'}`);
  flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  regEx = new RegExp(`${'Team Performance of'}(.*?)${'\\)\\&nbsp;'}`);
  let str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str[0] = str[0].replace(`&nbsp;`, " ");
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#b45f06";
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  docBody.appendParagraph(str[0].toString()).setAttributes(style);
  docBody.appendParagraph("\r\r");
  msg = msg.replace("<table>", "");
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  let headings = msg.match(regEx);
  let countCols = headings[0].split("<th >");
  let cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  let values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  let countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  let table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  headings = msg.match(regEx);
  countCols = headings[0].split("<th >");
  cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  headings = msg.match(regEx);
  countCols = headings[0].split("<th >");
  cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td >", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td><table ><thead><tr >"}(.*?)${"</tr></thead>"}`);
  headings = msg.match(regEx);
  countCols = headings[0].split("<th >");
  cells = [];
  for(let i=0; i<countCols.length - 1; i++) {
    cells[i] = [];
    regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
    let th = headings[0].match(regEx);
    headings[0] = headings[0].replace(th[0], "");
    let temp = th[0].replace("<th >", "");
    temp = temp.replace("</th>", "");
    cells[0][i] = temp;
  }
  regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  values = msg.match(regEx);
  values[0] = values[0].replace("<tbody>", "");
  values[0] = values[0].replace("</tbody>", "");
  countRows = values[0].split("<tr >");
  for(let i=1; i<=(countRows.length-1); i++) {
    cells[i] = [];
    let temp;
    for(let j=0; j < countCols.length-1; j++){
      let regEx1 = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let regEx2 = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
      if((values[0].match(regEx1) != null) && j==0){
        let td1 = values[0].match(regEx1);
        values[0] = values[0].replace(td1[0], "");
        temp = td1[0].replace("<th >", "");
        temp = temp.replace("</th>", "");
      }
      else if(values[0].match(regEx2) != null) {
        let td2 = values[0].match(regEx2);
        values[0] = values[0].replace(td2[0], "");
        temp = td2[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
      }
      temp = entities(temp);
      cells[i][j] = temp;
    }
  }
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</table></td></tr>", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  //   recurring piece
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  let subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  regEx = new RegExp(`${"<tr><td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  subStr = subStr[0].replace("<tr><td>", "");
  subStr = subStr.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  subStr = entities(subStr);
  docBody.appendParagraph(subStr).setAttributes(style);
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  subStr = str[0].match(regEx);
  str[0] = str[0].replace(subStr[0], "");
  str[0] = str[0].replace("<td>", "");
  str = str[0].replace("</td></tr>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendHorizontalRule();
  docBody.appendParagraph("\r\r");
  regEx = new RegExp(`${"<tr>"}(.*?)${"</tr>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  //   recurring piece
}

function multi (name, colour, docBody) {
  let msg = GmailApp.search(`subject:(Team Performance of ${name} for the week)`, 0, 1)[0].getMessages()[0].getBody();
  msg = msg.replace(`<div><span class="colour" style="color: rgb(0, 0, 0)"><span class="font" style="font-family: Verdana, arial, Helvetica, sans-serif"><span class="size" style="font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; white-space: normal; text-decoration-color: initial; float: none">Dear sir,<br /><br /><span class="colour" style="color: rgb(0, 0, 0)"><span class="font" style="font-family: Verdana, arial, Helvetica, sans-serif"><span class="size" style="font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; white-space: normal; text-decoration-color: initial; float: none">`, "");
  let regEx = new RegExp(`${'has been submitted successfully for your review'}(.*?)${'<tr><td valign="top" style="padding-top: 3px;">Number of TFs filled<'}`);
  msg = msg.replace(regEx, '<table><tr><td valign="top" style="padding-top: 3px;">Number of TFs filled<');
  regEx = new RegExp(`${'<tr><td valign="top">Upload Weekly PPT<'}(.*?)${'Thanks and Regards<br /><br /></div><br />'}`);
  msg = msg.replace(regEx, "</table>");
  msg = msg.replace("&nbsp;", " ");
  regEx = new RegExp(`${'style="'}(.*?)${'"'}`);
  let flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  regEx = new RegExp(`${' valign="top'}(.*?)${'"'}`);
  flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  regEx = new RegExp(`${'Team Performance of'}(.*?)${'\\) '}`);
  let str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = colour;
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  str[0] = entities(str[0]);
  docBody.appendParagraph(str[0]).setAttributes(style);
  docBody.appendParagraph("\r\r");
  msg = msg.replace("<table>", "");
  regEx = new RegExp(`${"Number of TFs filled</td><td >:</td><td><table >"}(.*?)${"</table>"}`);
  str = msg.match(regEx);
  if(str.length > 1) {   // TF table
    regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    str = str[0].replace("<tr><td >", "");
    str = str.replace("</td>", ":");
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
    style[DocumentApp.Attribute.FONT_SIZE] = 14;
    style[DocumentApp.Attribute.BOLD] = false;
    str = entities(str);
    docBody.appendParagraph(str).setAttributes(style);
    regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    regEx = new RegExp(`${"<thead>"}(.*?)${"</thead>"}`);
    let headings = msg.match(regEx);
    let countCols = headings[0].split("<th >");
    let cells = [];
    for(let i=0; i<countCols.length - 1; i++) {
      cells[i] = [];
      regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let th = headings[0].match(regEx);
      headings[0] = headings[0].replace(th[0], "");
      let temp = th[0].replace("<th >", "");
      temp = temp.replace("</th>", "");
      cells[0][i] = temp;
    }
    regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
    let values = msg.match(regEx);
    let countRows = values[0].split("<td >");
    for(let i=1; i<=(countRows.length-1)/(countCols.length-1); i++) {
      cells[i] = [];
      for(let j=0; j < countCols.length-1; j++){
        regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
        let td = values[0].match(regEx);
        values[0] = values[0].replace(td[0], "");
        let temp = td[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
        temp = entities(temp);
        cells[i][j] = temp;
      }
    }
    regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    msg = msg.replace("</table></td></tr>", "");
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    let table = docBody.appendTable(cells).setAttributes(style);
    docBody.appendParagraph("\r");
  } else {
      regEx = new RegExp(`${"<tr><td>Number of TFs filled"}(.*?)${":</td><td></td></tr>"}`);
      str = msg.match(regEx);
      msg = msg.replace(str[0], "");
  }    // TF table
  regEx = new RegExp(`${"Milestone achieved this week</td><td >:</td><td><table >"}(.*?)${"</table>"}`);
  str = msg.match(regEx);
  if(str != null) {   // Milestone table
    regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    str = str[0].replace("<tr><td >", "");
    str = str.replace("</td>", ":");
    style[DocumentApp.Attribute.FONT_SIZE] = 14;
    str = entities(str);
    docBody.appendParagraph(str).setAttributes(style);
    regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    regEx = new RegExp(`${"<thead>"}(.*?)${"</thead>"}`);
    headings = msg.match(regEx);
    countCols = headings[0].split("<th >");
    cells = [];
    for(let i=0; i<countCols.length - 1; i++) {
      cells[i] = [];
      regEx = new RegExp(`${"<th >"}(.*?)${"</th>"}`);
      let th = headings[0].match(regEx);
      headings[0] = headings[0].replace(th[0], "");
      let temp = th[0].replace("<th >", "");
      temp = temp.replace("</th>", "");
      cells[0][i] = temp;
    }
    regEx = new RegExp(`${"<td>"}(.*?)${"</thead>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
    values = msg.match(regEx);
    countRows = values[0].split("<td >");
    for(let i=1; i<=(countRows.length-1)/(countCols.length-1); i++) {
      for(let j=0; j < countCols.length-1; j++){
        regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
        let td = values[0].match(regEx);
        values[0] = values[0].replace(td[0], "");
        let temp = td[0].replace("<td >", "");
        temp = temp.replace("</td>", "");
        temp = entities(temp);
        cells[i][j] = temp;
      }
    }
    regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    msg = msg.replace("</table></td></tr>", "");
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    table = docBody.appendTable(cells).setAttributes(style);
    docBody.appendParagraph("\r");
  } else {
      regEx = new RegExp(`${"<tr><td>Milestone achieved this week"}(.*?)${":</td><td></td></tr>"}`);
      str = msg.match(regEx);
      msg = msg.replace(str[0], "");
  }    // milestone table
  regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = str[0].replace("<tr><td>", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str[0] = str[0].replace("<td>", "");
  str[0] = str[0].replace("</td>", "");
  let arr = str[0].split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</tr><tr>", "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</tr><tr>", "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</tr><tr>", "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</tr><tr>", "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</tr><tr>", "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", ":");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  str = entities(str);
  docBody.appendParagraph(str).setAttributes(style);
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<td>", "");
  str = str.replace("</td>", "");
  arr = str.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendHorizontalRule();
  docBody.appendParagraph("\r\r");
  regEx = new RegExp(`${"<td>"}(.*?)${"</td>"}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  msg = msg.replace("</tr><tr>", "");
}
