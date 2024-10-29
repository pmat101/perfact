
function dailyCheck() {
  ScriptApp.newTrigger("summary")
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.TUESDAY)
  .atHour(18)
  .nearMinute(0)
  .create(); 
}

function doGet (e) {
  let msg = GmailApp.search(`subject:(Weekly Performance Report -)`, 0, 1)[0];
  if(msg == undefined){    // If unavailable
    return;
  }
  msg = msg.getMessages()[0].getSubject();    // Get msg subject
  let regEx = new RegExp(`${'<div><span class="colour"'}(.*?)${'float: none">'}`);
  let str = msg.match(regEx);
  msg = msg.replace("Weekly Performance Report - ", "");
  msg = msg.replace("[", "");
  msg = msg.replace("]", "");
  let getPDF = DriveApp.getFilesByName(`WPF-report-${msg}`).next();   // WPF-report-${str}
  let pdfID = getPDF.getId();
  return HtmlService.createHtmlOutput(`
    <head>
      <style>
        body {
          background-color: #313131;
        }
        iframe {
          position: absolute;
          width: 100vw;
          height: 100vh;
          border: none;
        }
      </style>
    </head>
    <iframe src="https://drive.google.com/file/d/${pdfID}/preview" allow="autoplay"></iframe>
`);
}

function summary () {
  let docFile = DocumentApp.getActiveDocument();
  let docBody = docFile.getBody();       // now take each line and put in doc file and then convert to PDF
  docBody.clear();
  let pool = multi("Pool", "#741b47", docBody);
  let cove = multi("Cove", "#741b47", docBody);
  let ocean = multi("Ocean", "#bf9000", docBody);
  let pond = multi("Pond", "#38761d", docBody);
  let estuary = multi("Estuary", "#1155cc", docBody);
  let canal = multi("Canal", "#1155cc", docBody);
  let tributary = multi("Tributary", "#990000", docBody);
  let delta = multi("Delta", "#351c75", docBody);
  let reserv = reservoir ("Reservoir", "#b45f06", docBody);
  let lagoon = reservoir ("Lagoon", "#b45f06", docBody);
  glacier (docBody);
  fountain (docBody);    // remaining
  const summarize = [
    ["TEAM", "Bills Generated", "Total Working Hours"],
    ["Pool", pool[0], pool[1]],
    ["Cove", cove[0], cove[1]],
    ["Ocean", ocean[0],ocean[1]],
    ["Pond", pond[0], pond[1]],
    ["Estuary", estuary[0], estuary[1]],
    ["Canal", canal[0], canal[1]],
    ["Tributary", tributary[0], tributary[1]],
    ["Delta", delta[0], delta[1]],
    ["Reservoir", reserv[0], reserv[1]],
    ["Lagoon", lagoon[0], lagoon[1]],
    ["TOTAL", parseInt(pool[0])+parseInt(cove[0])+parseInt(ocean[0])+parseInt(pond[0])+parseInt(estuary[0])+parseInt(canal[0])+parseInt(tributary[0])+parseInt(delta[0])+parseInt(reserv[0])+parseInt(lagoon[0]), parseInt(pool[1])+parseInt(cove[1])+parseInt(ocean[1])+parseInt(pond[1])+parseInt(estuary[1])+parseInt(canal[1])+parseInt(tributary[1])+parseInt(delta[1])+parseInt(reserv[1])+parseInt(lagoon[1])]
  ];
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000";
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  let title = "Team Performance Summary";
  docBody.appendParagraph(title).setAttributes(style);
  docBody.appendParagraph("\r\r");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  let table = docBody.appendTable(summarize).setAttributes(style);
  docBody.appendHorizontalRule();
  docBody.appendParagraph("\r\r");
  docFile.saveAndClose();
  
  Utilities.sleep(29000);
  let today = new Date();
  let str = today.getDate();
  str = str + "/" + (today.getMonth()+1) + "/" + today.getFullYear();
  let pdfFile = docFile.getAs('application/pdf');
  const folder = DriveApp.getFolderById('18I57n_cPfehpfJKFGFPKXeLhewk3W7qQ'); // Folder ID: 18I57n_cPfehpfJKFGFPKXeLhewk3W7qQ
  let newPDF = folder.createFile(pdfFile);
  newPDF = newPDF.setName(`WPF-report-${str}`);
  Utilities.sleep(29000);
  const recipient = "gov.council@perfactgroup.in";
  const subject = `Weekly Performance Report - [${str}]`;
  const cc = "it.council@perfactgroup.in";
  const name = "IT/ COUNCIL/ PERFACT";
  const alias = GmailApp.getAliases();
  let body = `
    <head></head>
    <body>
      <p>Dear Governing Council members,</p>
      <p>Please find attached the latest weekly performance report, consolidating data from all team WPFs submitted for the week ending [${str}].</p>
      <p>This report provides a comprehensive overview of team progress, key achievements, and any identified challenges.</p>
      <p>We believe this data will be valuable in tracking performance and making informed decisions.</p>
      <p>This updated report is also available on the <a href="intranet.perfactgroup.in">INTRANET</a> under the MIS tab, accessible only to the Governing and IT Councils.</p>
      <p>Please let us know if you have any questions or require further details.</p>
      <p style="font-size: 1.1em"><strong>Please Note</strong>: <em>Summary table of the bills generated and time invested by each team in the past week is available at the bottom of the report.</em></p>
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
    attachments: [newPDF],
    from: alias[0],
    cc: cc
  });
  
}

function bill(msg) {
  let regEx = new RegExp(`${"TF07"}(.*?)${"</tr>"}`);
  let str = msg.match(regEx);
  if(str != null) {
    str = str[0].replace("TF07</td>", "");
    regEx = new RegExp(`${"<td  >"}(.*?)${"</td>"}`);
    let subStr = str.match(regEx);
    str = str.replace(subStr[0], "");
    regEx = new RegExp(`${"<td  >"}(.*?)${"</td>"}`);
    subStr = str.match(regEx);
    subStr = subStr[0].replace("<td  >", "");
    subStr = subStr.replace("</td>", "");
    return subStr;
  } else return 0;
}

function reservoirBill(msg) {
  let totality = 0;
  let regEx = new RegExp(`${"<thead>"}(.*?)${"</tbody>"}`);
  let str = msg.match(regEx);
  for(let i=0; i<5; i++) {
    regEx = new RegExp(`${"<td  >"}(.*?)${"</td>"}`);
    let subStr = str[0].match(regEx);
    subStr = subStr[0].replace("<td  >", "");
    subStr = subStr.replace("</td>", "");
    totality = totality + parseInt(subStr);
    str[0] = str[0].replace(subStr, "");
    str[0] = str[0].replace("<td  >", "");
    str[0] = str[0].replace("</td>", "");
  }
  return totality;
}

function tWO(msg) {
    let regEx = new RegExp(`${"<tr>"}(.*?)${"</td>"}`);
    let str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    regEx = new RegExp(`${"<td >"}(.*?)${"</td>"}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
    msg = msg.replace("<td >", "");
    msg = msg.replace("</td></tr>", "");
    return msg;
}

function entities (entity) {
  const arr = ["&#x2f;", "/", "&rdquo;", '"', "&ldquo;", '"', "&#x28;", "(", "&#x29;", ")", "&amp;", "&", "&#x2b;", "+", "&#x3d;", "=", "&#x25;", "%", "&#x27;", '"', "&#x3a;", ":", "&#x3b;", ";", "&nbsp;", " "];
  for(let i=0; i<arr.length; i=i+2) {
    while(entity.match(arr[i]) != null) {
      let e1 = entity.match(arr[i]);
      entity = entity.replace(e1[0], arr[i+1]);
    }
  }
  return entity;
}

function tables (msg) {
  let regEx = new RegExp(`${"<thead>"}(.*?)${"</thead>"}`);
  let str = msg.match(regEx);
  str = str[0].replace("<thead><tr >", "");
  str = str.replace("</tr></thead>", "");
  str = entities(str);
  let tHead = str.split("</th><th >");
  for(let i=0; i<tHead.length; i++) {
      tHead[i] = tHead[i].replace("<th >", "");
      tHead[i] = tHead[i].replace("</th>", "");
      tHead[i] = entities(tHead[i]);
  }
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<tbody><tr >", "");
  str = str.replace("</tr></tbody>", "");
  let tBody = str.split("</tr><tr >");
  let cells = [];
  for(let i=0; i<tBody.length; i++) {
    let tRow = tBody[i].split("</td><td  >");
    cells[i] = [];
    for(let j=0; j<tRow.length; j++) {
      cells[i][j] = tRow[j];
      cells[i][j] = cells[i][j].replace("<td  >", "");
      cells[i][j] = cells[i][j].replace("</td>", "");
      cells[i][j] = cells[i][j].replaceAll("<br />", " ");
      cells[i][j] = entities(cells[i][j]);
    }
  }
  cells.unshift(tHead);
  return cells;
}

function matrix (msg) {
  let regEx = new RegExp(`${"<thead>"}(.*?)${"</thead>"}`);
  let str = msg.match(regEx);
  str = str[0].replace("<thead><tr >", "");
  str = str.replace("</tr></thead>", "");
  str = entities(str);
  let tHead = str.split("</th><th >");
  for(let i=0; i<tHead.length; i++) {
      tHead[i] = tHead[i].replace("<th >", "");
      tHead[i] = tHead[i].replace("</th>", "");
      tHead[i] = entities(tHead[i]);
  }
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${"<tbody>"}(.*?)${"</tbody>"}`);
  str = msg.match(regEx);
  str = str[0].replace("<tbody><tr >", "");
  str = str.replace("</tr></tbody>", "");
  let tBody = str.split("</tr><tr >");
  let cells = [];
  for(let i=0; i<tBody.length; i++) {
    let tRow = tBody[i].split("<td  >");
    cells[i] = [];
    for(let j=0; j<tRow.length; j++) {
      cells[i][j] = tRow[j];
      cells[i][j] = cells[i][j].replace("<td  >", "");
      cells[i][j] = cells[i][j].replace("</td>", "");
      cells[i][j] = cells[i][j].replace("<th >", "");
      cells[i][j] = cells[i][j].replace("</th>", "");
      cells[i][j] = entities(cells[i][j]);
    }
  }
  cells.unshift(tHead);
  return cells;
}

function single (str) {
  str = str.replace("<tr>", "");
  str = str.replaceAll("<td >", "");
  str = str.replaceAll("<td  >", "");
  str = str.replaceAll("</td>", " ");
  str = str.replace("</tr>", "");
  str = str.replace("<table", "");
  return str;
}

function multiLine (msg, docBody, style) {
  let regEx = new RegExp(`${'<tr>'}(.*?)${'</td>'}`);
  let str = msg.match(regEx);
  let line = single(str[0]);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'<td >'}(.*?)${'</td>'}`);
  str = msg.match(regEx);
  line = line + single(str[0]);
  line = entities(line);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'<td >'}(.*?)${'</td></tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  let arr = line.split("<br />");
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  for(let i=0; i<arr.length; i++) {
    arr[i] = entities(arr[i]);
    docBody.appendListItem(arr[i]).setAttributes(style);
  }
  docBody.appendParagraph("\r");
}

function multi (name, colour, docBody) {
  let tf07;
  let past = new Date();
  past.setDate(past.getDate() - 7);
  let pastDate = `${past.getFullYear()}/${past.getMonth()+1}/${past.getDate()}`;
  let future = new Date();
  future.setDate(future.getDate() + 1);
  let futureDate = `${future.getFullYear()}/${future.getMonth()+1}/${future.getDate()}`;
  let msg = GmailApp.search(`subject:(Team Performance of ${name} for the week) after:${pastDate} before:${futureDate} `, 0, 1)[0];
  if(msg == undefined){    // If unavailable
    return;
  }
  msg = msg.getMessages()[0].getBody();    // Get msg body
  let regEx = new RegExp(`${'<div><span class="colour"'}(.*?)${'float: none">'}`);
  let str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'Dear sir'}(.*?)${'float: none">'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'style="'}(.*?)${'"'}`);   // Remove all style attributes
  let flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  msg = msg.replaceAll('valign="top"', "");
  regEx = new RegExp(`${'<tr><td >Upload'}(.*?)${'Regards<br /><br /></div><br />'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'Team Performance'}(.*?)${'\\) '}`);   // Heading
  str = msg.match(regEx);
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = colour;
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  str[0] = entities(str[0]);
  docBody.appendParagraph(str[0]).setAttributes(style);
  docBody.appendParagraph("\r\r");
  regEx = new RegExp(`${'Team Performance'}(.*?)${'TFs filled this week'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "TFs filled this week");
  regEx = new RegExp(`${'TFs filled this week'}(.*?)${'</tr>'}`);    // TFs
  str = msg.match(regEx);
  str = str[0].replaceAll("</td>", "");
  str = str.replaceAll("<td >", "");
  str = str.replaceAll("</tr>", "");
  let subReg = new RegExp(`${'TF'}(.*?)${':'}`);
  let subStr = str.match(subReg);
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  subStr = entities(subStr[0]);
  docBody.appendParagraph(subStr).setAttributes(style);

  // TFs filled this week
  str = str.replace(subStr, "");
  if(str == "Yes") {
    docBody.editAsText().appendText(" Yes");
    regEx = new RegExp(`${'<tr><td  >Number of TFs filled'}(.*?)${'<table >'}`);
    str = msg.match(regEx);
    subStr = str[0].replaceAll("</td>", "");
    subStr = subStr.replaceAll("<td  >", "");
    subStr = subStr.replaceAll("<td >", "");
    subStr = subStr.replaceAll("<tr>", "");
    subStr = subStr.replaceAll("<table >", "");
    docBody.appendParagraph(subStr).setAttributes(style);
    msg = msg.replace(str[0], "");
    tf07 = bill(msg);
    let cells = tables(msg);
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    let table = docBody.appendTable(cells).setAttributes(style);
    docBody.appendParagraph("\r");
    regEx = new RegExp(`${'<thead>'}(.*?)${'</table></td></tr>'}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
  } else {
    docBody.editAsText().appendText(" No");
    docBody.appendParagraph("\r");
    tf07 = 0;
  }
  regEx = new RegExp(`${'TFs filled this week'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Milestone
  regEx = new RegExp(`${'<tr><td >Milestone'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  str = str[0].replaceAll("</td>", "");
  str = str.replaceAll("<td >", "");
  str = str.replaceAll("<tr>", "");
  str = str.replaceAll("</tr>", "");
  subReg = new RegExp(`${'Milestone'}(.*?)${':'}`);
  subStr = str.match(subReg);
  subStr = entities(subStr[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  docBody.appendParagraph(subStr).setAttributes(style);
  str = str.replace(subStr, "");
  if(str == "Yes") {
    docBody.editAsText().appendText(" Yes");
    regEx = new RegExp(`${'<tr><td  >Details of Milestone'}(.*?)${'<table >'}`);
    str = msg.match(regEx);
    subStr = str[0].replaceAll("</td>", "");
    subStr = subStr.replaceAll("<td  >", "");
    subStr = subStr.replaceAll("<td >", "");
    subStr = subStr.replaceAll("<tr>", "");
    subStr = subStr.replaceAll("<table >", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
    docBody.appendParagraph(subStr).setAttributes(style);
    msg = msg.replace(str[0], "");
    let cells = tables(msg);
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    let table = docBody.appendTable(cells).setAttributes(style);
    docBody.appendParagraph("\r");
    regEx = new RegExp(`${'<thead>'}(.*?)${'</table></td></tr>'}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
  } else {
    docBody.editAsText().appendText(" No");
    docBody.appendParagraph("\r");
  }
  regEx = new RegExp(`${'<tr><td >Milestone'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'<tr><td  >'}(.*?)${'<table >'}`);    // Team work info
  str = msg.match(regEx);
  subStr = str[0].replaceAll("</td>", "");
  subStr = subStr.replaceAll("<td  >", "");
  subStr = subStr.replaceAll("<td >", "");
  subStr = subStr.replaceAll("<tr>", "");
  subStr = subStr.replaceAll("<table >", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  docBody.appendParagraph(subStr).setAttributes(style);
  msg = msg.replace(str[0], "");
  let cells = tables(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  let table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</table></td></tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Total working hours
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  let totality = tWO(str[0]);
  let line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");

  // Total active projects
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");

  // Targets planned this week
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Targets achieved this week
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Highlights
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Low points - complaints
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Challenges
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Bottlenecks
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");

  // other reasons
  regEx = new RegExp(`${'<tr><td >other reasons'}(.*?)${':</td>'}`);
  if(msg.match(regEx) != null) {
    multiLine(msg, docBody, style);
    regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
  }

  // Projected targets
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  
  docBody.appendHorizontalRule();
  docBody.appendParagraph("\r\r");
  return [tf07, totality];
}

function glacier (docBody) {
  let past = new Date();
  past.setDate(past.getDate() - 7);
  let pastDate = `${past.getFullYear()}/${past.getMonth()+1}/${past.getDate()}`;
  let future = new Date();
  future.setDate(future.getDate() + 1);
  let futureDate = `${future.getFullYear()}/${future.getMonth()+1}/${future.getDate()}`;
  let msg = GmailApp.search(`subject:(Team Performance of Glacier for the week) after:${pastDate} before:${futureDate}`, 0, 1)[0];
  if(msg == undefined){    // If unavailable
    return;
  }
  msg = msg.getMessages()[0].getBody();    // Get msg body
  let regEx = new RegExp(`${'<div><span class="colour"'}(.*?)${'float: none">'}`);
  let str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'Dear sir'}(.*?)${'float: none">'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'style="'}(.*?)${'"'}`);   // Remove all style attributes
  let flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  msg = msg.replaceAll('valign="top"', "");
  regEx = new RegExp(`${'<tr><td >Upload'}(.*?)${'Regards<br /><br /></div><br />'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'Team Performance'}(.*?)${'\\) '}`);   // Heading
  str = msg.match(regEx);
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = "#434343";
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  str[0] = entities(str[0]);
  docBody.appendParagraph(str[0]).setAttributes(style);
  docBody.appendParagraph("\r\r");
  regEx = new RegExp(`${'Team Performance'}(.*?)${'BDs filled this week'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "BDs filled this week");
  regEx = new RegExp(`${'BDs filled this week'}(.*?)${'</tr>'}`);    // BDs
  str = msg.match(regEx);
  str = str[0].replaceAll("</td>", "");
  str = str.replaceAll("<td >", "");
  str = str.replaceAll("</tr>", "");
  let subReg = new RegExp(`${'BD'}(.*?)${':'}`);
  let subStr = str.match(subReg);
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  subStr = entities(subStr[0]);
  docBody.appendParagraph(subStr).setAttributes(style);

  // No. of BDs
  str = str.replace(subStr, "");
  if(str == "Yes") {
    docBody.editAsText().appendText(" Yes");
    regEx = new RegExp(`${'<tr><td  >Number of BDs filled'}(.*?)${'<table >'}`);
    str = msg.match(regEx);
    subStr = str[0].replaceAll("</td>", "");
    subStr = subStr.replaceAll("<td  >", "");
    subStr = subStr.replaceAll("<td >", "");
    subStr = subStr.replaceAll("<tr>", "");
    subStr = subStr.replaceAll("<table >", "");
    subStr = entities(subStr);
    docBody.appendParagraph(subStr).setAttributes(style);
    msg = msg.replace(str[0], "");
    let cells = tables(msg);
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    let table = docBody.appendTable(cells).setAttributes(style);
    docBody.appendParagraph("\r");
    regEx = new RegExp(`${'<thead>'}(.*?)${'</table></td></tr>'}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
  } else {
    docBody.editAsText().appendText(" No");
    docBody.appendParagraph("\r");
  }
  regEx = new RegExp(`${'BDs filled this week'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Proposal sent
  regEx = new RegExp(`${'<tr>'}(.*?)${'<table'}`);
  str = msg.match(regEx);
  let line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  msg = msg.replace('>', "");
  msg = msg.replace(' ', "");
  msg = msg.replace('</table></td></tr>', "");
  cells = matrix(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</tbody>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Bills raised
  regEx = new RegExp(`${'<tr>'}(.*?)${'<table'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  msg = msg.replace('>', "");
  msg = msg.replace(' ', "");
  msg = msg.replace('</table></td></tr>', "");
  cells = matrix(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</tbody>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Client complaints
  regEx = new RegExp(`${'<tr><td >Client Complaint'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  str = str[0].replaceAll("</td>", "");
  str = str.replaceAll("<td >", "");
  str = str.replaceAll("<tr>", "");
  str = str.replaceAll("</tr>", "");
  subReg = new RegExp(`${'Client Complaint'}(.*?)${':'}`);
  subStr = str.match(subReg);
  subStr = entities(subStr[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  docBody.appendParagraph(subStr).setAttributes(style);
  str = str.replace(subStr, "");
  if(str == "Yes") {
    docBody.editAsText().appendText(" Yes");
    regEx = new RegExp(`${'<tr><td  >Details of Milestone'}(.*?)${'<table >'}`);
    str = msg.match(regEx);
    subStr = str[0].replaceAll("</td>", "");
    subStr = subStr.replaceAll("<td  >", "");
    subStr = subStr.replaceAll("<td >", "");
    subStr = subStr.replaceAll("<tr>", "");
    subStr = subStr.replaceAll("<table >", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
    docBody.appendParagraph(subStr).setAttributes(style);
    msg = msg.replace(str[0], "");
    cells = tables(msg);
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    let table = docBody.appendTable(cells).setAttributes(style);
    docBody.appendParagraph("\r");
    regEx = new RegExp(`${'<thead>'}(.*?)${'</table></td></tr>'}`);
    str = msg.match(regEx);
    msg = msg.replace(str[0], "");
  } else {
    docBody.editAsText().appendText(" No");
    docBody.appendParagraph("\r");
  }
  regEx = new RegExp(`${'<tr><td >Client Complaint'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Targets planned this week
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Targets achieved this week
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Highlights
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Low points
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // challenges
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Projected Targets for next week
  multiLine(msg, docBody, style);
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  
  docBody.appendHorizontalRule();
  docBody.appendParagraph("\r\r");
}

function reservoir (name, colour, docBody) {
  let past = new Date();
  past.setDate(past.getDate() - 7);
  let pastDate = `${past.getFullYear()}/${past.getMonth()+1}/${past.getDate()}`;
  let future = new Date();
  future.setDate(future.getDate() + 1);
  let futureDate = `${future.getFullYear()}/${future.getMonth()+1}/${future.getDate()}`;
  let msg = GmailApp.search(`subject:(Team Performance of ${name} for the week) after:${pastDate} before:${futureDate}`, 0, 1)[0];
  if(msg == undefined){    // If unavailable
    return;
  }
  msg = msg.getMessages()[0].getBody();    // Get msg body
  let regEx = new RegExp(`${'<div><span class="colour"'}(.*?)${'float: none">'}`);
  let str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'Dear sir'}(.*?)${'float: none">'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'style="'}(.*?)${'"'}`);   // Remove all style attributes
  let flag = 0;
  while(flag == 0){
    let matchText = msg.match(regEx);
    if(matchText != null){
      msg = msg.replace(matchText[0], "");
    } else {
      flag = 1;
    }
  }
  msg = msg.replaceAll('valign="top"', "");
  regEx = new RegExp(`${'<tr><td >Upload'}(.*?)${'Regards<br><br></div><br />'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");
  regEx = new RegExp(`${'Team Performance'}(.*?)${'\\)'}`);   // Heading
  str = msg.match(regEx);
  let style = {};
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = colour;
  style[DocumentApp.Attribute.FONT_SIZE] = 18;
  style[DocumentApp.Attribute.BOLD] = true;
  str = entities(str[0]);
  docBody.appendParagraph(str).setAttributes(style);
  docBody.appendParagraph("\r\r");
  regEx = new RegExp(`${'Team Performance'}(.*?)${'Number of submissions'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "Number of submissions");

  // No of submissions
  regEx = new RegExp(`${'Number of submissions'}(.*?)${'<table'}`);
  str = msg.match(regEx);
  let line = single(str[0]);
  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  style[DocumentApp.Attribute.BOLD] = false;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  msg = msg.replace('>', "");
  msg = msg.replace(' ', "");
  msg = msg.replace('</table></td></tr>', "");
  let submit = reservoirBill(msg);
  cells = matrix(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</tbody>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Targets planned
  regEx = new RegExp(`${'<tr>'}(.*?)${'<table'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  msg = msg.replace('>', "");
  msg = msg.replace(' ', "");
  msg = msg.replace('</table></td></tr>', "");
  cells = matrix(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</tbody>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Targets achieved
  regEx = new RegExp(`${'<tr>'}(.*?)${'<table'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  msg = msg.replace('>', "");
  msg = msg.replace(' ', "");
  msg = msg.replace('</table></td></tr>', "");
  cells = matrix(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</tbody>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // projected targets
  regEx = new RegExp(`${'<tr>'}(.*?)${'<table'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  msg = msg.replace(str[0], "");
  msg = msg.replace('>', "");
  msg = msg.replace(' ', "");
  msg = msg.replace('</table></td></tr>', "");
  cells = matrix(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</tbody>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Team work info
  regEx = new RegExp(`${'<tr><td  >'}(.*?)${'<table >'}`);
  str = msg.match(regEx);
  subStr = str[0].replaceAll("</td>", "");
  subStr = subStr.replaceAll("<td  >", "");
  subStr = subStr.replaceAll("<td >", "");
  subStr = subStr.replaceAll("<tr>", "");
  subStr = subStr.replaceAll("<table >", "");
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  docBody.appendParagraph(subStr).setAttributes(style);
  msg = msg.replace(str[0], "");
  cells = tables(msg);
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  table = docBody.appendTable(cells).setAttributes(style);
  docBody.appendParagraph("\r");
  regEx = new RegExp(`${'<thead>'}(.*?)${'</table></td></tr>'}`);
  str = msg.match(regEx);
  msg = msg.replace(str[0], "");

  // Total working hours
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  let totality = tWO(str[0]);
  line = single(str[0]);
  style[DocumentApp.Attribute.FONT_SIZE] = 14;
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");

  // Total active projects
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");


  // Highlights
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");


  // Lowlights
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");


  // Challanges faced 
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");


  // Bottlenecks
  regEx = new RegExp(`${'<tr>'}(.*?)${'</tr>'}`);
  str = msg.match(regEx);
  line = single(str[0]);
  line = entities(line);
  docBody.appendParagraph(line).setAttributes(style);
  docBody.appendParagraph("\r");
  msg = msg.replace(str[0], "");

  docBody.appendHorizontalRule();
  docBody.appendParagraph("\r\r");
  return [submit, totality];
}

function fountain (docBody) {}
