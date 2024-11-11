function dailyCheck() {
  ScriptApp.newTrigger("sendEmail")
  .timeBased()
  .everyDays(1)
  .atHour(11)
  .nearMinute(0)
  .create();
}

function sendEmail() {
  const today = new Date();
  const currentDate = today.getDate();
  const currentMonth = today.getMonth();
  
  // 02. SOP for Recording EAC Meetings via MS Teams (discuss with Shweta ma'am)
  // 04. PCODE list (remind Sakhsi Singhal for remaining PCODEs)
  
  // dates and months are numbers but months are zero indexed
  if(currentDate == 1 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    markOutOfOffice();
  }
  else if(currentDate == 2){
    hybridWorkModel();
  }
  else if(currentDate == 3 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    tg01();
  }
  else if(currentDate == 5 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    pgApprovedVendorList();
  }
  else if(currentDate == 7 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    sopCommonPurchase();
  }
  else if(currentDate == 7 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    tg02();
  }
  else if(currentDate == 8){
    clientVisitChecklistZohoForm();
  }
  else if(currentDate == 9 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    shareWindowMSTeams();
  }
  else if(currentDate == 10){
    bookMeetingSlotsWithNBPB();
  }
  else if(currentDate == 11 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    externalEmail();
  }
  else if(currentDate == 11 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    shareTabGoogleMeet();
  }
  else if(currentDate == 12){
    setupGmailForOfflineUse();
  }
  else if(currentDate == 13 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    pdfGear();
  }
  else if(currentDate == 13 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    governingCouncilSystem();
  }
  else if(currentDate == 14){
    saveTeamsRecording();
  }
  else if(currentDate == 15 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    chatAudioVideoCalls();
  }
  else if(currentDate == 15 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    monthlyParty();
  }
  else if(currentDate == 16){
    empReqForm();
  }
  else if(currentDate == 17 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    fileShareRule();
  }
  else if(currentDate == 17 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    travelBooking();
  }
  else if(currentDate == 19 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    msOfficeEIA();
  }
  else if(currentDate == 20){
    gatePassPolicy();
  }
  else if(currentDate == 21 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    presentPPTinMSTeams();
  }
  else if(currentDate == 21 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    laptopBag();
  }
  else if(currentDate == 22){
    masterPPTnGuidelines();
  }
  else if(currentDate == 23 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    tg03();
  }
  else if(currentDate == 23 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    officeSupplies();
  }
  else if(currentDate == 24){
    activeZohoForms();
  }
  else if(currentDate == 25 && (currentMonth == 0 || currentMonth == 3 || currentMonth == 6 || currentMonth == 9)){
    conferenceRoom();
  }
  else if(currentDate == 26){
    guidelinesSmoothCheckIn();
  }
  else if(currentDate == 27 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    purchaseForm();
  }
  else if(currentDate == 27 && (currentMonth == 2 || currentMonth == 5 || currentMonth == 8 || currentMonth == 11)){
    emailEtiquette();
  }
  else if(currentDate == 28){
    maintainingHygieneInWorkplace();
  }
}

function markOutOfOffice () {
  const imgBlob2 = DriveApp.getFilesByName("sop-to-mark-out-of-office-2.png").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("sop-to-mark-out-of-office-3.png").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("sop-to-mark-out-of-office-4.png").next().getBlob();
  const imgBlob6 = DriveApp.getFilesByName("sop-to-mark-out-of-office-6.png").next().getBlob();
  let inlineImages = {
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
    'img6': imgBlob6
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = `[Reminder] SOP: Marking "Out of Office" in Google Calendar`;
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please find the SOP below to mark yourself "Out of Office" in Google Calendar. This will let you on set up a notification to inform colleagues about your unavailability during leaves, site visits or other commitments.</p>
    <br>
    <h1 style="font-weight:bold; text-align:center">SOP to mark “Out of Office” in Google Calendar</h1>
    <br>
    <h3><strong>Step1 -</strong> Open your web browser and go to Google Calendar [https://calendar.google.com/]</h3>
    <br>
    <h3><strong>Step2 -</strong> Click on the 'Create' event button or any time-block in the calendar.</h3>
    <img src="cid:img2">
    <br>
    <h3><strong>Step3 -</strong> Select "Out of office" from the menu.</h3>
    <img src="cid:img3">
    <br>
    <h3><strong>Step4 -</strong> Specify a start and end time or choose the 'All day' option to pick the start and end dates for your absence.</h3>
    <img src="cid:img4">
    <br>
    <h3><strong>Step5 -</strong> Choose to automatically decline new and/or existing meetings.</h3>
    <br>
    <h3><strong>Step6 -</strong> You can even customize your "Out of Office" decline message.</h3>
    <img src="cid:img6">
    <br>
    <h3><strong>Step7 -</strong> Save and enjoy your day out of the office.</h3>
    <br>
    <p>SOP link- https://docs.google.com/document/d/1HIFUKhFt2nERE7EPGisJ_FIJeN4ttn2bdilUFW0ieIU/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function hybridWorkModel() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Hybrid Work Model Policy and Roster Sheet Update";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the Hybrid Work Model Policy and updated Roster sheet. To facilitate smoother planning and coordination, we have transitioned to a monthly roster format. Please fill out the sheet for the entire month, providing clear details of your WFO and WFH days.</p>
    <br>
    <h1 style="font-weight:bold; text-align:center">Hybrid Work Model Policy</h1>
    <br>
    <p>This hybrid work model is designed to offer flexibility and productivity. It aims to strike a balance between in-office and remote work. We believe this model will enhance employee satisfaction, productivity and work-life balance.</p>
    <br>
    <h2><strong>Key Points of the Hybrid Work Model:</strong></h2>
    <ol>
      <li><strong style="font-size: 1.1em">Minimum In-Office Presence:</strong>
        <ul>
          <li>A minimum of 3 days of Work From Office (WFO) per week is mandatory for all employees.</li>
          <li>TAs will have to do WFO on all 5 days.</li>
        </ul>
      </li>
      <li><strong style="font-size: 1.1em">Work From Home (WFH) Guidelines:</strong>
        <ul>
          <li>A maximum of 2 days of Work From Home (WFH) is permitted between Monday and Friday within a week.</li>
          <li>WFH days cannot be consecutive. They must be on alternate days.</li>
          <li>Team Head and Deputy Team Head cannot take WFH on the same day.</li>
          <li>WFH isn't allowed for TAs.</li>
        </ul>
      </li>
      <li><strong style="font-size: 1.1em">Saturday Work and Rotational Offs:</strong>
        <ul>
          <li>Two Saturdays will be working days for all employees per month.</li>
          <li>One Saturday will be WFH and the other will be WFO.</li>
          <li>Two Rotational Saturdays Off (RSO) will be provided.</li>
          <li>Fifth Saturday (if any) will be a mandatory WFO day.</li>
          <li>TAs will be given 2 RSO and 2 WFO per month.</li>
        </ul>
      </li>
      <li><strong style="font-size: 1.1em">Holidays and Leaves:</strong>
        <ul>
          <li>Holidays and leaves cannot be counted towards WFO or WFH days.</li>
        </ul>
      </li>
      <li><strong style="font-size: 1.1em">Post-Maternity Leave:</strong>
        <ul>
          <li>Post-maternity leave, female employees must work from home for the first six months.</li>
          <li>During the subsequent six months, they must visit the office once a week.</li>
          <li>After one year of delivery, they must adhere to the standard hybrid work model guidelines.</li>
        </ul>
      </li>
    </ol>
    <br>
    <h2><strong>Scheduling and Approvals:</strong></h2>
    <ul>
      <li>The Team Head (TH) or Deputy Team Head (DTH) will be responsible for sharing the monthly WFO/WFH schedule with the team in the last week of the preceding month.</li>
      <li>Any exceptions to the policy must be approved by the Business Head (BH) or Top Management.</li>
      <li>BH and TH will declare the names of permanent WFH members.</li>
    </ul>
    <br>
    <p>Please adhere to the guidelines and cooperate with the team to ensure smooth implementation.</p>
    <br>
    <p>Link to Policy Doc- [https://docs.google.com/document/d/1hRPTwPwLjBcU21c8PiMwaU8NirqCI0FizdR7wBq2lIY/edit?usp=sharing]</p>
    <p>Link to Monthly Roster Sheet- [https://docs.google.com/spreadsheets/d/1KhXGu_LoY5g854RLzHhweJY9amdCYQRdAmMkyaHo62I/edit?usp=sharing]</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function tg01 () {
  const imgBlob1 = DriveApp.getFilesByName("TG01-chatSpaceGuidelinesSOP-IMG_0087.PNG").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1
  }
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] TG01 Chat Space Guidelines";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please see the useful SOP for TG-01 Chat spaces.</p>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">TG01 Chat space guidelines</h1>
    <br>
    <p style="font-weight:bold">Question:  What is the use of Chat Space?</p>
    <p>Answer:</p>
    <ul>
      <li>It allows you to group your most important members together and communicate with them. Day to day communication with clients, FAE, Team head , EIA coordinator readily we can update the team.</li>
      <li>The communication being done remains in the loops of Everybody who has stake in the project at the same time instead sending texting or emailing in chain manner.</li>
      <li>Additionally, you can assign tasks, organize project meetings, mention relevant links of the project, and as per the preference the type notification delivery can be selected.</li>
    </ul>
    <br>
    <p style="font-weight:bold">Question:  Who will create the Chat Space group?</p>
    <p>Answer:  BD head will create the chat space once the job is received and will add BH/Team Head as chat space manager.</p>
    <p>BD head then will add scope of work, timeline, contact information, deliverables and relevant information about the project and leave the space. BH/TH can then add other people.</p>
    <br>
    <p style="font-weight:bold">Question:  What will be the naming format for the Chat Space group?</p>
    <p>Answer:  Project Proponent Name- State Name- Business Head- EC Name- (Greenfield/Brownfield) New/Expansion/Amendent/ NIPL/Extension/Corrigendum project Category (5ga)- Team Name.</p>
    <img src="cid:img1">
    <br>
    <p style="font-weight:bold">Question:  What is to be added in the project description?</p>
    <p>Answer:  Parivesh portal login ID, Project Pcode and PP contact no. & email (as there is word limit in brief description). Rest relevant links, preference to VC mode, critical information (including Project Sheet) can be given in initial chat messages.</p>
    <br>
    <p style="font-weight:bold">Question:  Who can edit the project description?</p>
    <p>Answer:  Only the Space Manager can edit the Chat space.</p>
    <br>
    <p style="font-weight:bold">Question:  Who will be a member of chat space?</p>
    <p>Answer:  Space managers- Team Head, EIA Coordinator, BH, CEO/COO (reporting head) as per involvement</p>
    <p>Members- working team members (probationers allowed but trainees not allowed).</p>
    <br>
    <p style="font-weight:bold">Question:  How does it help in improving team interaction?</p>
    <p>Answer:  It reduces the no. of emails and gives a common platform for sharing information / key issues within the team space.</p>
    <br>
    <p style="font-weight:bold">Question:  Is it easy to share project documents?</p>
    <p>Answer:  Yes, all links for project drive can be pasted in chat space.</p>
    <br>
    <p style="font-weight:bold">Question:  Can tasks be given easily within a team?</p>
    <p>Answer:  Yes, tasks can be alloted to the project incharge along with timeline.</p>
    <br>
    <p style="font-weight:bold">Question:  Does it allow users to set specific notification settings?</p>
    <p>Answer:  Yes</p>
    <br>
    <p style="font-weight:bold">Question:  Can you research the chat spaces messages in google chat?</p>
    <p>Answer:  Yes you can find the keyword in chat space</p>
    <br>
    <p style="font-weight:bold">Question:  Once the google chat has been left can we rejoin the same chat space?</p>
    <p>Answer:  Yes</p>
    <br>
    <br>
    <p>Chat space training- https://support.google.com/a/users/answer/9247502?hl=en</p>
    <p>Cheat Sheet- https://support.google.com/a/users/answer/9299928?hl=en&ref_topic=9348682</p>
    <p>SOP link- https://docs.google.com/document/d/1XDm47_FbZxC8s0Q-aq-1_9wAF1oYCux_z6uN_UdvCW4/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    inlineImages: inlineImages,
    cc: cc
  });
}

function pgApprovedVendorList () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Vendors for any Repair / Maintenance at Perfact Group";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>At Perfact Group we prefer Quality of the Service / Work our utmost priority & always try to provide the best quality of Service to our Customers.</p>
    <p>We also expect the same from our vendors.</p>
    <p>Keeping in mind the quality of work, timeliness and dependability, we have developed some core vendors for every work in our Company.</p>
    <p>Please note that any repair/ maintenance work should be done only from the approved vendors as decided.</p>
    <p>If the approved is not able to do the work by whatsoever reason, only Urban Company experts should be scheduled for the same.</p>
    <p>No unapproved vendor should be called for any repair maintenance work without prior approval from Top Management.</p>
    <p>Please find below the link mentioned of PG Approved Vendor list Google sheet and all are requested to please enter the Vendor details in the same: https://docs.google.com/spreadsheets/d/1wiPocloQtPVNuJWBAEbLX59jXUjCyvLb8Of_agPukeM/edit?usp=sharing</p>
    <p>For any office repair/ maintenance work Admin Team will coordinate with the Vendor, all departments should raise a ticket to Freshdesk with CC to Team Admin for any issue.</p>
    <p>For any Lab Specific work i.e. in which Lab Section's intervention is required, only then Lab section will coordinate with the vendor</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function sopCommonPurchase() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] SOP (Common Purchase)";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
    <head></head>
    <body>
    <p>Dear all,</p>
    <p>Please find below the SOP for Common Purchase.</p>
    <br>
    <h1 style="font-weight:bold; text-align:center">Standard Operating Procedure: Common Purchase Process</h1>
    <br>
    <h2><strong>Purpose: </strong></h2>
    <p>The purpose of this Standard Operating Procedure (SOP) is to establish a standardized process for requesting and procuring items from the administration team. This SOP ensures efficient and timely fulfillment of employee requests while maintaining an accurate inventory and proper financial documentation.</p>
    <br>
    <h2><strong>Objectives: </strong></h2>
      <ol>
        <li>To streamline the process of requesting and procuring items from the administration team.</li>
        <li>To maintain an updated inventory of available items.</li>
        <li>To ensure proper financial documentation and payment processes.</li>
      </ol>
    <br>
    <h2><strong>Procedure: </strong></h2>
    <h3><em>Step 1: </em> Request Submission</h3>
    <ul>
      <li>When an employee requires any items from the admin, they must fill out the Common Request Form.</li>
      <li>The form should include the details of the requested item, quantity, and any specific requirements.</li>
    </ul>
    <br>
    <h3><em>Step 2: </em> Availability Check</h3>
    <ul>
      <li>The admin team will check the availability of the requested item in their stock.</li>
      <li>If the item is available, the admin team will issue it immediately to the employee.</li>
      <li>The admin team will update their stock accordingly and share the Combid (Common Purchase Bid) with the top management for purchasing approval, ensuring desired stock levels are maintained.</li>
    </ul>
    <br>
    <h3><em>Step 3: </em> Unavailability of Item</h3>
    <ul>
      <li>If the requested item is not available in stock, the admin team will gather quotes from different vendors.</li>
      <li>The admin team will create a Combid request that includes the vendor quotes.</li>
      <li>The Combid request will be sent to the top management for approval.</li>
    </ul>
    <br>
    <h3><em>Step 4: </em> Procurement Request</h3>
    <ul>
      <li>Upon receiving approval from the management, the admin team will prepare a procurement request.</li>
      <li>The procurement request will be sent via email to the accounts department.</li>
      <li>The email will include all necessary details, such as payment terms, vendor account information, and item delivery time.</li>
    </ul>
    <br>
    <h3><em>Step 5: </em> Payment and Delivery</h3>
    <ul>
      <li>The accounts department will make the payment to the vendor as per the provided details.</li>
      <li>Once the payment is made, the vendor will deliver the item to the office.</li>
    </ul>
    <br>
    <h3><em>Step 6: </em> Item Issuance and Stock Update</h3>
    <ul>
      <li>The admin team will receive the item from the vendor.</li>
      <li>They will issue the received item to the requester.</li>
      <li>The admin team will update their stock accordingly to maintain accurate inventory records.</li>
    </ul>
    <br>
    <h3><em>Step 7: </em> Financial Documentation</h3>
    <ul>
      <li>The admin team will raise the bill in the expense.zoho.com system for proper financial documentation.</li>
    </ul>
    <br>
    <h2><strong>Supporting Information: </strong></h2>
    <ul>
      <li>For detailed instructions and guidelines, refer to the complete SOP document available at the following link: [SOP Link: https://docs.google.com/document/d/1WIBMgxoNRILXN8_fUHfnIXGMY4iSA6i6aFp3RYtGWLg/edit?usp=sharing]</li>
      <li>To access the Common Purchase Form for submitting requests, please click on the following link: [Form Link: https://zfrmz.com/44MxSiwpKIxnFaNQBHWx]</li>
    </ul>
    <br>
    <p>By following this SOP, we aim to ensure a standardized and efficient process for common purchases within our organization. If you have any questions or require further clarification, please contact the administration team.</p>
    <p>Thank you for your cooperation in implementing this SOP effectively.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function tg02 () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] TG02 Guidelines- Document Version Naming and Sending Documents to External Parties";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>To ensure consistent and efficient document management, please adhere to the following naming conventions:</p>
    <ul>
      <li><strong>Version Control</strong>: For all documents saved as Word, PDF, Excel, or PowerPoint (docx/xlsx/pptx/pdf), please use the following suffix convention:
        <ul>
          <li><strong>R0</strong>: Initial draft by the first person</li>
          <li><strong>R1</strong>: Updated by the first reviewer</li>
          <li><strong>R2</strong>: Updated by the second reviewer, and so on</li>
        </ul>
      </li>
      <li><strong>Date Stamp</strong>: Append the date in the format DD MM YY (e.g., 09 03 24) to the filename.</li>
    </ul>
    <br>
    <p>Example:</p>
    <ul>
      <li><em>EIA - Arbuda Industries &lt;PCODE&gt; - R4 - 09 03 24</em></li>
    </ul>
    <br>
    <p>Additional Notes:</p>
    <ul>
      <li>While the filename remains unchanged for saved documents like Google Docs/Slides/Sheets, they may be named with a version identifier (e.g., "EIA Sent to Client on 23 08 22").</li>
    </ul>
    <br>
    <p>This naming convention ensures clear version tracking and facilitates efficient collaboration.</p>
    <p>Thank you for your cooperation!</p>
    <br><br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function clientVisitChecklistZohoForm() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Client Visit Checklist Zoho Form";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>As part of the ongoing efforts to enhance our client engagement and hospitality, we are introducing a new initiative to streamline the arrangements for client visits at our Head Office.</p>
    <p>Client Visit Checklist Zoho Form is implemented with immediate effect.</p>
    <p>Link of the form: https://zfrmz.com/Obv7IM1JlNpv4BRMoeMw</p>
    <p>Each team is required to fill & submit the form ahead of any scheduled client visit. This form is designed to gather essential information about the visit's requirements, ensuring that the Admin team can make the necessary arrangements to provide the best possible hospitality to our valued clients.</p>
    <p>This form will enable us to anticipate and address all the necessary details, allowing us to create a positive and seamless experience for our clients during their visit to our Head Office.</p>
    <p>Everyone's cooperation in adhering to this new process is greatly appreciated. For any further clarification, feel free to reach out to the Admin Team.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function shareWindowMSTeams () {
  const imgBlob1 = DriveApp.getFilesByName("sop-share-window-msteams-1.jpg").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("sop-share-window-msteams-2.jpg").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("sop-share-window-msteams-3.jpg").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("sop-share-window-msteams-4.jpg").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] SOP to share a window in MS Teams";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the SOP to share a program window in MS Teams meeting. This method promotes a more focused and private video call experience by sharing only relevant information on your screen, eliminating distractions from other applications and cluttered desktops.</p>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">SOP to share a program window in MS Teams meeting</h1>
    <br>
    <h3><strong>Step1 - </strong> join a MS Teams meeting and locate the controls at the top of the screen, click on the “Share” button to share your content.</h3>
    <img src="cid:img1">
    <br>
    <h3><strong>Step2 - </strong> you get the options to share your entire screen, a specific program window or a Powerpoint presentation. Select the “Window” option from the dropdown menu.</h3>
    <img src="cid:img2">
    <br>
    <h3><strong>Step3 - </strong> click on the specific window you want to share with the participants.</h3>
    <img src="cid:img3">
    <br>
    <h3><strong>Step4 - </strong> screen sharing in Teams does not include computer audio by default. If you want to share sound from your computer (e.g. to play a video), click the "Include sound" button before sharing.</h3>
    <img src="cid:img4">
    <br><br>
    <p>SOP link- https://docs.google.com/document/d/12bj5kNQyDcIevqFwZS9PBPIzQ39ioFDO20I0NQNZ_mw/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function bookMeetingSlotsWithNBPB() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Improved Scheduling for Internal Discussions, Client Meetings & Availability";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>This email combines important updates regarding scheduling meetings and communication with both Nipun sir and Praveen sir.</p>
    <p style="font-weight:bold">Internal Discussions with NB sir:</p>
    <ul>
      <li>To streamline internal discussions (excluding recurring meetings), a new calendar booking system has been implemented.</li>
      <li>Use this link to book a time slot for project discussions, ad hoc meetings, etc: https://calendar.app.google/eiVC4FMdy1evVRi99</li>
      <li>Generate a Teams link after booking time slot and add NB sir and all members required for the meeting</li>
      <li>This eliminates the need for direct messages, calendar conflicts and forgotten meeting links.</li>
    </ul>
    <br>
    <p style="font-weight:bold">External Client Meetings with NB sir:</p>
    <ul>
      <li>A separate booking page has been created for scheduling external client meetings with NB sir</li>
      <li>Use this link and add required details: https://calendar.app.google/wWiARwVbsvRNd6wCA</li>
      <li>Generate a Teams link after booking time slot and add NB sir and all internal/external members required for the meeting</li>
      <li>These links have also been added to the Master List of Forms and the Intranet for everyone's ease.</li>
    </ul>
    <br>
    <p style="font-weight:bold">PB sir's Availability:</p>
    <ul>
      <li>To address concerns about project delays, PB sir is introducing a dedicated time slot for quick discussions.</li>
      <li>He will be available daily from 5 PM to 6 PM (subject to change) for 5-minute focused discussions on urgent topics.</li>
      <li>To secure a time slot:
        <ul>
          <li>Briefly explain the discussion topic via WhatsApp.</li>
          <li>Send relevant details/links via email.</li>
          <li>Ensure clear and concise points for a productive 5-minute meeting.</li>
        </ul>
      </li>
    </ul>
    <br>
    <p>These changes aim to improve communication, scheduling efficiency, and project timelines.</p>
    <p>If you have any questions, please don't hesitate to reach out to the IT Department</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function externalEmail () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Enhancing Email Security and External Email Posting";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>In our continuous efforts to safeguard our organization's sensitive information and maintain a secure communication environment, we allow external email posting only for employees with the designation of Deputy Team Head (DTH) and above.</p>
    <p>This does not impact our inbound emails, Google Drive, Google Chat or other Google service settings for any user.</p>
    <p>For example an employee who is not given access to send email to anyone outside Perfact Group, can receive emails from an external source and use Google Drive, Google Chat & other Google services in general.</p>
    <p>We kindly request your cooperation in adhering to this email security protocol. If you have any questions, concerns, or require further clarification, please do not hesitate to reach out to the IT department or your immediate supervisor.</p>
    <p>Thank you for your understanding and ongoing commitment to the security of our organization's communications.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function shareTabGoogleMeet () {
  const imgBlob1 = DriveApp.getFilesByName("share-tab-google-meet-1.jpg").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("share-tab-google-meet-2.jpg").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("share-tab-google-meet-3.jpg").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] SOP to share a specific tab in Google Meet";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the SOP to share a specific tab in Google Meet. It offers a focused and efficient way to present information during video calls.</p>
    <p><strong>Benefits:</strong></p>
    <ul>
      <li>Direct attention to the specific website or document within the tab, eliminating distractions from other browser tabs and desktop clutter</li>
      <li>Maintain confidentiality by keeping other browsing activity hidden</li>
      <li>Focus on the content being discussed, promoting shorter and more productive meetings</li>
    </ul>
    <p><strong>Use Cases:</strong></p>
    <ul>
      <li>Client Presentations</li>
      <li>Daily Huddles</li>
      <li>Internal Meetings</li>
    </ul>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">SOP to share a specific tab in Google Meet</h1>
    <br>
    <h3><strong>Step1 - </strong> during an ongoing meeting in Google Meet, locate the controls at the bottom of the screen and click on the “Present Now” button</h3>
    <img src="cid:img1">
    <br>
    <h3><strong>Step2 - </strong> you get the options to share your entire screen, a specific program window or a single tab. Select the specific tab from your Chrome browser (tab audio is shared by default)</h3>
    <img src="cid:img2">
    <br>
    <h3><strong>Step3 - </strong> you can choose to share content on a different tab by opening that tab and clicking on the “Share this tab instead” button</h3>
    <img src="cid:img3">
    <br><br>
    <p>SOP link- https://docs.google.com/document/d/1_2MDrlApXOlwna6cgjxfeBsGF8iGBT2xet1PqOQES3k/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function setupGmailForOfflineUse() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Set up Gmail for Offline Use";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>Hope this email finds you well. Shared below is a useful tool that allows you to use your email offline. This can be particularly helpful if you are traveling or do not have access to a reliable internet connection.</p>
    <p>To use Gmail offline, you will need to first enable offline access in your Gmail settings. Here's how:</p>
    <ul>
      <li>Go to the gear icon in the top right corner of your Gmail account and click "Settings".</li>
      <li>In the "General" tab, scroll down to the "Offline" section.</li>
      <li>Click the "Enable offline mail" button.</li>
      <li>Follow the prompts to set up offline access.</li>
      <li>Once you have enabled offline access, you can use Gmail without an internet connection by going to https://mail.google.com/mail/u/0/?ui=2&zy=h in your web browser. You can read, search, and compose emails, as well as archive and delete messages. Any changes you make while offline will be synced with your account the next time you go online.</li>
    </ul>
    <p>I hope this is helpful. If you have any questions or need further assistance, don't hesitate to reach out.</p>
    <p>For further information, you can access the following link- https://binaryfork.com/use-gmail-offline-3846/</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function pdfGear () {
  const imgBlob1 = DriveApp.getFilesByName("install-pdf-gear-1.png").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("install-pdf-gear-2.png").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("install-pdf-gear-3.png").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("install-pdf-gear-4.png").next().getBlob();
  const imgBlob5 = DriveApp.getFilesByName("install-pdf-gear-5.png").next().getBlob();
  const imgBlob6 = DriveApp.getFilesByName("install-pdf-gear-6.png").next().getBlob();
  const imgBlob7 = DriveApp.getFilesByName("install-pdf-gear-7.png").next().getBlob();
  const imgBlob8 = DriveApp.getFilesByName("install-pdf-gear-8.png").next().getBlob();
  const imgBlob9 = DriveApp.getFilesByName("install-pdf-gear-9.png").next().getBlob();
  const imgBlob10 = DriveApp.getFilesByName("install-pdf-gear-10.png").next().getBlob();
  const imgBlob11 = DriveApp.getFilesByName("install-pdf-gear-11.png").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
    'img5': imgBlob5,
    'img6': imgBlob6,
    'img7': imgBlob7,
    'img8': imgBlob8,
    'img9': imgBlob9,
    'img10': imgBlob10,
    'img11': imgBlob11,
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] Installation Guide for PDF Gear Software";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the SOP to install PDF Gear, a handy tool for making PDF tasks easier like editing, converting to Excel/Word, compressing, and managing files.</p>
    <p><strong>Purpose: </strong> To streamline the process of working with PDF documents, such as editing, converting to Excel/Word formats, compressing files, and importing/exporting documents, improvements have been made for enhanced user convenience.</p>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">SOP to Install PDF Gear</h1>
    <br>
    <h3><strong>Step1 - </strong> Go to pdfgear.com on your web browser</h3>
    <img src="cid:img1">
    <br>
    <h3><strong>Step2 - </strong> Click on Free Download</h3>
    <img src="cid:img2">
    <br>
    <h3><strong>Step3 - </strong> Once the download is done, open the downloaded file</h3>
    <img src="cid:img3">
    <br>
    <h3><strong>Step4 - </strong> Click on "Yes" when asked to start the installation and follow the instructions</h3>
    <img src="cid:img4">
    <br>
    <h3><strong>Step5 - </strong> Click "I ACCEPT THE AGREEMENT" and then "Next"</h3>
    <img src="cid:img5">
    <br>
    <h3><strong>Step6 - </strong> Keep clicking "Next" until it's finished</h3>
    <img src="cid:img6">
    <br>
    <img src="cid:img7">
    <br>
    <img src="cid:img8">
    <br>
    <img src="cid:img9">
    <br>
    <h3><strong>Step10 - </strong> Click "Finish" to end the installation</h3>
    <img src="cid:img10">
    <br>
    <h3><strong>Step11 - </strong> After installing, right-click on any PDF file, select "Open with," and pick PDFgear. Make sure to set it as your default PDF viewer</h3>
    <img src="cid:img11">
    <br><br>
    <p>SOP link- https://docs.google.com/document/d/155_XYBNLqy-TKcWQW_UbKw9v0QeadhdFyOmaC37AGCw/edit?usp=sharing</p>
    <p>Feel free to reach out if you have any queries to Himanshu Kohli at 8882038491</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function governingCouncilSystem() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Governing Council and the updated Council System";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head>
    <style>
      table {
        border-collapse: collapse;
        width: 100%;
      }

      th, td {
        padding: 10px;
        text-align: left;
        border-bottom: 1px solid #ddd;
      }

      th {
        background-color: #f2f2f2;
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>This email provides a comprehensive update on our recently implemented governance structure, along with resources to enhance communication and participation.</p>
    <br>
    <h2>Strengthened Governance for Strategic Growth</h2>
    <p>As previously communicated, we are pleased to announce the establishment of the Governing Council, responsible for overseeing strategic decision-making. This high-level body leads eight specialized councils focused on critical areas such as accreditation and employee development.</p>
    <br>
    <h2>Enhanced Collaboration and Accessibility</h2>
    <ul>
      <li><h3>Streamlined Governance Structure:</h3>
        <ul>
          <li><strong>Governing Council: </strong> Provides strategic direction and oversees major decisions.</li>
          <li><strong>9 Specialized Councils: </strong> (Staff, Accreditation, EIA, Business, Lab, Recruitment, IT, Admin and Quality Control): Offer expertise in their respective domains.</li>
          <li><strong>Working Groups: </strong> Replace sub-councils to promote efficiency and focused action.</li>
        </ul>
      </li>
      <li><h3>Direct Council Communication: </h3> Each council now has a dedicated email address to facilitate your inquiries and suggestions:
        <ul>
        <br>
          <li><strong>Governing Council: </strong> gov.council@perfactgroup.in</li>
          <br>
          <li>
            <table>
              <thead>
                <tr>
                  <th>Council Name</th>
                  <th>Email Address</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>Staff Council</td>
                  <td>staff.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>Accreditation Council</td>
                  <td>acrd.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>EIA Council</td>
                  <td>eia.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>Business Council</td>
                  <td>bsns.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>Lab Council</td>
                  <td>lab.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>Recruitment Council</td>
                  <td>rct.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>IT Council</td>
                  <td>it.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>Admin Council</td>
                  <td>admin.council@perfactgroup.in</td>
                </tr>
                <tr>
                  <td>Quality Control Council</td>
                  <td>qc.council@perfactgroup.in</td>
                </tr>
              </tbody>
            </table>
          </li>
        </ul>
      </li>
    </ul>
    <br>
    <h2>Embrace Continuous Improvement</h2>
    <p>We actively seek your valuable feedback on this new structure. Your participation is critical to shaping Perfact's future. We encourage you to actively engage with your respective council or the Governing Council directly.</p>
    <br>
    <p>Detailed PDFs outlining the new system, goals, changes and list of council members are attached.</p>
    <p>Working together through collaborative governance, we can achieve exceptional results.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc,
    attachments:[
      DriveApp.getFilesByName("governing-Council-System-v2.6-Latest-20-5-24.pdf").next().getBlob(),
      DriveApp.getFilesByName("governing-Council-System-Structure-Latest-13-9-24.pdf").next().getBlob()
      ]
    });
}

function monthlyParty() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Team-wise monthly event schedule";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head>
    <style>
      h3 {
        background-color: #FF9900;
        padding: 0.5em 1em;
        display: inline;
      }
      table {
        border-collapse: collapse;
        width: 100%;
      }
      th, td {
        padding: 10px;
        text-align: center;
        border: 1px solid #ddd;
      }
      th {
        background-color: #00FFFF;
      }
      td {
        background-color: #B7E1CD;
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>We are pleased to inform you that we have curated a team-wise event schedule to help organize office events effectively. This schedule is intended to ensure that everyone in the company gets an opportunity to participate in team events and build stronger connections with their colleagues.</p>
    <p>To ensure smooth coordination, we kindly request that the team organizing the party sends an email to both the Account and Admin teams in advance, along with the proposed budget. This will help us keep track of the planned events and allocate resources accordingly. The email should include the following details:</p>
    <ul>
      <li>Proposed date and time of the party</li>
      <li>Venue and theme of the party</li>
      <li>Estimated number of attendees</li>
      <li>Budget and any sponsors (if applicable)</li>
    </ul>
    <br>
    <p>We would like to remind all the teams that the proposed budget should be within the allocated budget for the party calendar. If you require any assistance, please do not hesitate to contact the Admin team, who will be happy to help you with any concerns or questions you may have.</p>
    <p>Let us all work together to make these events enjoyable for everyone.</p>
    <br>
    <h3>Team Wise Calendar :-</h3>
    <hr>
    <table>
      <thead>
        <tr>
          <th>S. No</th>
          <th>Team Name</th>
          <th>Month</th>
          <th>Festival or Birthday</th>
          <th>Date</th>
          <th>Day</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>1</td>
          <td>Pond Team</td>
          <td>April</td>
          <td>Baisakhi</td>
          <td>13/04/2024</td>
          <td>Saturday</td>
        </tr>
        <tr>
          <td>2</td>
          <td>Estuary Team</td>
          <td>May</td>
          <td>International Labour Day</td>
          <td>01/05/2024</td>
          <td>Wednesday</td>
        </tr>
        <tr>
          <td>3</td>
          <td>Reservoir Team</td>
          <td>June</td>
          <td>World Environment Day</td>
          <td>05/06/2024</td>
          <td>Wednesday</td>
        </tr>
        <tr>
          <td>4</td>
          <td>Fountain Team</td>
          <td>July</td>
          <td>Guru Poornima</td>
          <td>19/07/2024</td>
          <td>Friday</td>
        </tr>
        <tr>
          <td>5</td>
          <td>Arctic Team</td>
          <td>August</td>
          <td>Independence Day</td>
          <td>14/08/2024</td>
          <td>Wednesday</td>
        </tr>
        <tr>
          <td>6</td>
          <td>Canal Team</td>
          <td>September</td>
          <td>Teacher's Day</td>
          <td>05/09/2024</td>
          <td>Thursday</td>
        </tr>
        <tr>
          <td>7</td>
          <td>Tributary Team</td>
          <td>October</td>
          <td>Dussehra / Diwali</td>
          <td>25/10/2024</td>
          <td>Wednesday</td>
        </tr>
        <tr>
          <td>8</td>
          <td>Glacier Team</td>
          <td>November</td>
          <td>Children's Day</td>
          <td>14/11/2024</td>
          <td>Thursday</td>
        </tr>
        <tr>
          <td>9</td>
          <td>Ocean Team</td>
          <td>December</td>
          <td>Christmas</td>
          <td>24/12/2024</td>
          <td>Tuesday</td>
        </tr>
        <tr>
          <td>10</td>
          <td>Pool Team</td>
          <td>January</td>
          <td>Lohri</td>
          <td>13/01/2025</td>
          <td>Monday</td>
        </tr>
        <tr>
          <td>11</td>
          <td>Creek Team</td>
          <td>February</td>
          <td>Vasant Panchami</td>
          <td>01/02/2025</td>
          <td>Saturday</td>
        </tr>
        <tr>
          <td>12</td>
          <td>Lakes Team</td>
          <td>March</td>
          <td>Holi</td>
          <td>12/03/2025</td>
          <td>Wednesday</td>
        </tr>
      </tbody>
    </table>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function empReqForm() {
  const subject = "[Reminder] Employee Requisition Form";
  const name = "IT Administrator / PERFACT";
  const recipient = "akta.chugh@perfactgroup.in, gmk@perfactgroup.in, neha.aggarwal@perfactgroup.in, santosh.pant@perfactgroup.in, nipunbhargava@perfactgroup.in, dr.anilkumar@perfactgroup.in, rachna.dogra@perfactgroup.in, ajay.pasricha@perfactgroup.in";
  const cc = "sreeja.sreekanth@perfactgroup.in, richa.aggarwal@perfactgroup.in, shweta.rajput@perfactgroup.in, cipia.mehta@perfactgroup.in, disha.patel@perfactgroup.in, sadaf.akhtar@perfactgroup.in, urvi.pritam@perfactgroup.in, saloni.sharma@perfactgroup.in, aarti.gupta@perfactgroup.in, chandrashekhar.jha@perfactgroup.in, deepika.arora@perfactgroup.in, shailly.luthra@perfactgroup.in, meerambika.behera@perfactgroup.in, tushar.bansal@perfactgroup.in, hr@perfactgroup.in, pranav.mathur@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>Hope this email finds you well. We are pleased to introduce the Employee Requisition Form to streamline the hiring process. This form will facilitate efficient communication of your team's manpower requirements to the Recruitment Council.</p>
    <p>The form will capture essential details about the required position, including job description, qualifications, and other relevant information. This standardized approach will expedite the recruitment process and ensure that all necessary information is readily available.</p>
    <p>Access the form through the Intranet or using this link: [https://zfrmz.com/e4fjPCQI0YVvL1wCtO0z]</p>
    <p>We encourage all teams to utilize this form for all future hiring requests.</p>
    <p>Thank you for your cooperation.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function fileShareRule () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Preferred Method for Sharing Data";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>I hope this message finds you well.</p>
    <p>To streamline our workflow and make it easier for everyone to reference the data shared, We'd like to request a small adjustment in how we share information when responding to requests from seniors or colleagues.</p>
    <p>Going forward, please follow these guidelines when sharing any data:</p>
    <ul>
      <li><strong>For internal sharing:</strong>
        <ul>
          <li>If the data is brief, paste it directly into the body of your email.</li>
          <li>If the data is lengthy or includes multiple details, please attach it as a PDF.</li>
          <li>Avoid sharing multiple links, sheets, or documents unless specifically requested.</li>
        </ul>
      </li>
      <li><strong>For external sharing:</strong>
        <ul>
          <li>Utilize our SharePoint platform to create a common sharing drive. Upload files to One Drive and share the link via email to people outside our firm.</li>
        </ul>
      </li>
    </ul>
    <p>This practice will not only speed up our workflow but also ensure that the data provided is final and most current version.</p>
    <p>Thank you for your cooperation in improving our efficiency.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function saveTeamsRecording () {
  const imgBlob1 = DriveApp.getFilesByName("sop-to-save-msteams-recording-1.png").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("sop-to-save-msteams-recording-2.png").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("sop-to-save-msteams-recording-3.png").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("sop-to-save-msteams-recording-4.png").next().getBlob();
  const imgBlob5 = DriveApp.getFilesByName("sop-to-save-msteams-recording-5.png").next().getBlob();
  const imgBlob6 = DriveApp.getFilesByName("sop-to-save-msteams-recording-6.png").next().getBlob();
  const imgBlob7 = DriveApp.getFilesByName("sop-to-save-msteams-recording-7.png").next().getBlob();
  const imgBlob8 = DriveApp.getFilesByName("sop-to-save-msteams-recording-8.png").next().getBlob();
  const imgBlob9 = DriveApp.getFilesByName("sop-to-save-msteams-recording-9.png").next().getBlob();
  const imgBlob10 = DriveApp.getFilesByName("sop-to-save-msteams-recording-10.png").next().getBlob();
  const imgBlob11 = DriveApp.getFilesByName("sop-to-save-msteams-recording-11.png").next().getBlob();
  const imgBlob12 = DriveApp.getFilesByName("sop-to-save-msteams-recording-12.png").next().getBlob();
  const imgBlob13 = DriveApp.getFilesByName("sop-to-save-msteams-recording-13.png").next().getBlob();
  const imgBlob14 = DriveApp.getFilesByName("sop-to-save-msteams-recording-14.png").next().getBlob();
  const imgBlob15 = DriveApp.getFilesByName("sop-to-save-msteams-recording-15.png").next().getBlob();
  const imgBlob16 = DriveApp.getFilesByName("sop-to-save-msteams-recording-16.png").next().getBlob();
  const imgBlob17 = DriveApp.getFilesByName("sop-to-save-msteams-recording-17.png").next().getBlob();
  const imgBlob18 = DriveApp.getFilesByName("sop-to-save-msteams-recording-18.png").next().getBlob();
  const imgBlob19 = DriveApp.getFilesByName("sop-to-save-msteams-recording-19.png").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
    'img5': imgBlob5,
    'img6': imgBlob6,
    'img7': imgBlob7,
    'img8': imgBlob8,
    'img9': imgBlob9,
    'img10': imgBlob10,
    'img11': imgBlob11,
    'img12': imgBlob12,
    'img13': imgBlob13,
    'img14': imgBlob14,
    'img15': imgBlob15,
    'img16': imgBlob16,
    'img17': imgBlob17,
    'img18': imgBlob18,
    'img19': imgBlob19,
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] SOP to Save MS Teams Recording";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
  <p>Dear all,</p>
  <p>Please find the below (SOP) to save MS team recording in Google Drive.</p>
  <br>
  <h1 style="font-weight:bold; text-align:center">SOP to Save MS Teams Recording</h1>
  <br>
  <h3><strong>Step1 -</strong> Join the meeting and click on "More options". Select "Record and transcribe" and start the recording.</h3>
  <img src="cid:img1">
  <br>
  <h3><strong>Step2 -</strong> Once the meeting is finished, click on "More options" again, select "Record and transcribe", and stop the recording.</h3>
  <img src="cid:img2">
  <br>
  <h3><strong>Step3 -</strong> Click on "Stop" and leave the meeting.</h3>
  <img src="cid:img3">
  <br>
  <h3><strong>Step4 -</strong> Open your Microsoft Teams software, go to the "Files" option, and click on "Add cloud storage".</h3>
  <img src="cid:img4">
  <br>
  <h3><strong>Step5 -</strong> Choose "Google Drive".</h3>
  <img src="cid:img5">
  <br>
  <h3><strong>Step6 -</strong> Enter your email address and click on "Sign in".</h3>
  <img src="cid:img6">
  <br>
  <h3><strong>Step7 -</strong> Enter your password and click on "Sign in". You may need to approve the sign-in from your phone.</h3>
  <img src="cid:img7">
  <br>
  <h3><strong>Step8 -</strong> Once you've signed in, click on "Files" and then select "OneDrive".</h3>
  <img src="cid:img8">
  <br>
  <h3><strong>Step9 -</strong> Locate the "Recording" folder and click on it.</h3>
  <img src="cid:img9">
  <br>
  <h3><strong>Step10 -</strong> Click on the meeting recording to select it and then click on "Move" from the top menu.</h3>
  <img src="cid:img10">
  <br>
  <h3><strong>Step11 -</strong> Choose "Google Drive".</h3>
  <img src="cid:img11">
  <br>
  <h3><strong>Step12 -</strong> Select the "Meet Recordings" folder.</h3>
  <img src="cid:img12">
  <br>
  <h3><strong>Step13 -</strong> Click on "Move".</h3>
  <img src="cid:img13">
  <br>
  <h3><strong>Step14 -</strong> Once the file has been moved to your Google Drive, open "My Drive" and go to the "Meet Recordings" folder.</h3>
  <img src="cid:img14">
  <br>
  <h3><strong>Step15 -</strong> Right-click on the file and select "Move". Choose the following location: <br> <a href="https://drive.google.com/drive/folders/1PXasGQRhE7dhh0Dn1XMaPKIuVsrY6GC-?usp=share_link">SHARING-GDRIVE > MS Team Meeting Recordings.</a></h3>
  <img src="cid:img15">
  <br>
  <h3><strong>Step16 -</strong> Click on "Move"</h3>
  <img src="cid:img16">
  <br>
  <h3><strong>Step17 -</strong> Once the file has been uploaded, right-click on it and select "Get shareable link".</h3>
  <img src="cid:img17">
  <br>
  <h3><strong>Step18 -</strong> You will also see a  Google Sheet  in the folder. Open the sheet.</h3>
  <img src="cid:img18">
  <br>
  <h3><strong>Step19 -</strong> Fill in the details for your meeting and paste the recording link into the sheet to easily and quickly access the recordings.</h3>
  <img src="cid:img19">
  <br>
  <p>SOP link- https://docs.google.com/document/d/1Y32EYK66j7148GDikQFy0KX3ILedLE1CMOPSCN7ZTHA/edit?usp=sharing</p>
  <p>Google sheet link- https://docs.google.com/spreadsheets/d/1Pn8VHnzqJkhaUp7oNy1q9Ww3jRa9dmLFaUMaxjVuOKA/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function chatAudioVideoCalls () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Enhance Collaboration: Get acquainted with Google Chat Audio & Video Calls";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>Google Chat Space, the chat application within our Google Workspace, offers built-in features for audio and video calls. This exciting functionality allows for seamless communication and collaboration, right from our familiar Workspace environment</p>
    <br>
    <p style="font-weight:bold">Benefits of Using Google Chat Audio & Video Calls:</p>
    <ul>
      <li><strong>Convenience: </strong> Initiate calls directly within ongoing chats, eliminating the need for switching platforms or scheduling meetings in advance.</li>
      <li><strong>Accessibility: </strong> Leverage Google Chat's accessibility across various devices, including desktops, laptops, tablets, and smartphones.</li>
      <li><strong>Enhanced Communication: </strong> Engage in face-to-face interactions through video calls, fostering a more personal and productive experience.</li>
      <li><strong>Improved Efficiency: </strong> Streamline communication workflows by seamlessly transitioning from text chat to audio or video discussions as needed.</li>
    </ul>
    <br>
    <p style="font-weight:bold">Guide to Using Google Chat Audio & Video Calls:</p>
    <ul>
      <li>Access Google Chat through your web browser or download the app on your device.</li>
      <li>Choose the individual you want to initiate a call with.</li>
      <li>Locate the video call icon (camera) or phone icon at the top right corner of the chat window.</li>
      <li>Click the desired icon (camera for video call, phone for audio call) to initiate the call.</li>
      <li>The recipient will receive a call notification. Once they join, your audio or video conversation will begin.</li>
    </ul>
    <br>
    <p><strong>Additional Tip: </strong> <em>You can share your screen during a video call for presentations or collaborative work or utilize chat functionalities like sending files and links alongside your audio or video call.<em></p>
    <br>
    <p>By incorporating Google Chat audio and video calls into our communication routine, we can foster a more collaborative and efficient work environment. We encourage you to explore this valuable feature and experience the benefits firsthand.</p>
    <p>For further assistance or detailed instructions, feel free to visit the Google Chat Help Center: https://support.google.com/messages/answer/7189714?hl=en</p>
    <p>We're confident this new way to connect will enhance our teamwork and communication within the organisation.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function travelBooking () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Travel Booking format";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head>
    <style>
      .lg {
        font-size: 1.2em;
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>As part of our ongoing efforts to enhance record-keeping & ease of working, we would like to inform you about our Travel Booking Procedure.</p>
    <p>To enhance the process, the Travel Booking Requisition form has been implemented for booking purposes.</p>
    <p>For your reference, please find the Standard Operating Procedure (SOP) and the Zoho form link provided below.</p>
    <p>If you have any questions, please do not hesitate to reach out to Glacier Team or Babita Kumari.</p>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">SOP for Travel Booking</h1>
    <br>
    <ol>
      <li class="lg">The Requester or the person traveling is required to complete the Travel Requisition Zoho form, providing all necessary details and then submit it.
        <ul>
          <li>LINK OF ZOHO FORM: https://zfrmz.com/jA18pDRu6yQgqsQo6BSf</li>
        </ul>
      </li>
      <li class="lg">An automated email will be generated and sent to the Requester, Creek Team, Business Head, and Top Management.</li>
      <li class="lg">The Business Head will review the booking and give approval by responding to all recipients in the email.</li>
      <li class="lg">If the travel expenses are to be reimbursed by the client, the Business Head should mark CC to the Info & Arctic team while approving the Travel Requisition.</li>
      <li class="lg">Glacier and Creek Teams will assess the fare, considering both Regular and Corporate rates for flight tickets, and proceed with the booking through the designated agent.</li>
      <li class="lg">Following the booking of tickets, the Creek Team will share the details in the same email chain and also communicate it on the Admin- Tour & Travels Google Chat group.</li>
      <li class="lg">Web check-in for VPs and above will be handled by the Creek Team; however, all others are required to perform their own web check-in.</li>
    </ol>
    <br>
    <div style="color:red">
      <h3>Note:</h3>
      <ol type="A">
        <li class="lg">Technical Associates assigned to each team are responsible to fill and submit the Travel Requisition form on behalf of the respective Business Heads.</li>
        <li class="lg">All details regarding tickets and hotel accommodations must be clearly mentioned in the Travel Requisition form to prevent any confusion or issues.</li>
        <li class="lg">Flight charges may vary between Regular and Corporate rates.</li>
        <li class="lg">If additional baggage needs to be booked, this should be mentioned in the remarks field of the Travel Requisition form.</li>
      </ol>
    </div>
    <br>
    <p>Link of the SOP: https://docs.google.com/document/d/1v94Kj_n9TubX4APPgNEIMlypNbtBHBIvgCIMuyNZzBE/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function msOfficeEIA () {
  const imgBlob1 = DriveApp.getFilesByName("online-msoffice-working-eiasection-1.jpg").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("online-msoffice-working-eiasection-2.jpg").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("online-msoffice-working-eiasection-3.jpg").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("online-msoffice-working-eiasection-4.jpg").next().getBlob();
  const imgBlob5 = DriveApp.getFilesByName("online-msoffice-working-eiasection-5.jpg").next().getBlob();
  const imgBlob6 = DriveApp.getFilesByName("online-msoffice-working-eiasection-6.jpg").next().getBlob();
  const imgBlob7 = DriveApp.getFilesByName("online-msoffice-working-eiasection-7.jpg").next().getBlob();
  const imgBlob8 = DriveApp.getFilesByName("online-msoffice-working-eiasection-8.jpg").next().getBlob();
  const imgBlob9 = DriveApp.getFilesByName("online-msoffice-working-eiasection-9.jpg").next().getBlob();
  const imgBlob10 = DriveApp.getFilesByName("online-msoffice-working-eiasection-10.jpg").next().getBlob();
  const imgBlob11 = DriveApp.getFilesByName("online-msoffice-working-eiasection-11.jpg").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
    'img5': imgBlob5,
    'img6': imgBlob6,
    'img7': imgBlob7,
    'img8': imgBlob8,
    'img9': imgBlob9,
    'img10': imgBlob10,
    'img11': imgBlob11
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] Online MS Office working- EIA Section";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the SOP for online MS Office working- EIA section:</p>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">SOP for online MS Office working- EIA Section</h1>
    <br>
    <h2 style="color:red">SECTION A- CREATION OF DOCUMENT</h2>
    <br>
    <h3><strong>Step1 -</strong> Login to your MS Teams account and click on Teams on the left side.</h3>
    <img src="cid:img1">
    <br>
    <h3><strong>Step2 -</strong> Click on the concerned Team channel under Your Teams.</h3>
    <img src="cid:img2">
    <br>
    <h3><strong>Step3 -</strong> Click on Files>>>> Word Document.</h3>
    <img src="cid:img3">
    <br>
    <h3><strong>Step4 -</strong> Name the document accordingly and click on create.</h3>
    <img src="cid:img4">
    <br>
    <h3><strong>Step5 -</strong> Select the File>>>> click on the 3 dots>>>>> click on Copy Link.</h3>
    <img src="cid:img5">
    <br>
    <h3><strong>Step6 -</strong> Click on Copy & paste the same in Google Chrome/ Microsoft Edge.</h3>
    <img src="cid:img6">
    <br><br>
    <h2 style="color:red">SECTION B- UPDATE THE ACCESS SETTINGS OF THE DOCUMENT</h2>
    <br>
    <h3><strong>Step7 -</strong> On the top right corner click on Share>>>>>> Manage Access.</h3>
    <img src="cid:img7">
    <br>
    <h3><strong>Step8 -</strong> Click on the icon shown and select the access needs to be given.</h3>
    <img src="cid:img8">
    <br><br>
    <h2 style="color:red">SECTION C- SHARING THE DOCUMENT</h2>
    <br>
    <h3><strong>Step9 -</strong> On the top right corner click on Share>>>>>> Share.</h3>
    <img src="cid:img9">
    <br>
    <h3><strong>Step10 -</strong> Click on the icon to change the access as required, enter the Email ID of the person to whom document needs to be shared and click on send.</h3>
    <img src="cid:img10">
    <br>
    <h3><strong>Step11 -</strong>  You can also send the document through the link, Click on Copy, the link will be copied to clipboard.</h3>
    <img src="cid:img11">
    <br>
    <h3><strong>Step12 -</strong> Paste the link in the Email and send it to the concerned person.</h3>
    <br>
    <p>Link to the SOP- https://docs.google.com/document/d/1MpotxBzSVRVTAQVfUHlKZ0q_nn6uDPc4KmeyXLKVSLk/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function gatePassPolicy() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Implementation of Gate Pass Policy";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>We hope this email finds you well. This is to inform you about the Gate Pass policy which is designed to enhance security aspects and monitor assets and keep a track of the same, streamline access control, and ensure the safety of assets.</p>
    <p>Detailed instructions and access to the gate pass policy is given below. If you have any questions or concerns regarding this policy, please do not hesitate to reach out to the Admin Department.</p>
    <p>Your cooperation and commitment to our security measures are greatly appreciated. We look forward to a smooth transition and a safer work environment for all.</p>
    <p>Link the doc- https://docs.google.com/document/d/1D94gnpXgysth7AlRB-vzA0AH73dxDiaICxcJJEkpMDE/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function presentPPTinMSTeams () {
  const imgBlob1 = DriveApp.getFilesByName("sop-present-ppt-ms-teams-1.jpg").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("sop-present-ppt-ms-teams-2.jpg").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("sop-present-ppt-ms-teams-3.jpg").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("sop-present-ppt-ms-teams-4.jpg").next().getBlob();
  const imgBlob5 = DriveApp.getFilesByName("sop-present-ppt-ms-teams-5.jpg").next().getBlob();
  const imgBlob6 = DriveApp.getFilesByName("sop-present-ppt-ms-teams-6.jpg").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
    'img5': imgBlob5,
    'img6': imgBlob6,
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] SOP to present a PPT in MS Teams";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Sharing a PowerPoint presentation directly in Microsoft Teams is a useful tool that can significantly enhance video call experience for clients and colleagues.</p>
    <br>
    <p><strong>Benefits: </strong></p>
    <ul>
      <li>Deliver professional and engaging presentations directly within the meeting window, eliminating the need for clunky screen sharing.</li>
      <li>Focus solely on the presentation without distractions from your desktop or other applications.</li>
      <li>Navigate through your slides seamlessly with keyboard shortcuts and presenter view features.</li>
    </ul>
    <br>
    <p><strong>Use Cases: </strong></p>
    <ul>
      <li>Client Presentations</li>
      <li>Daily Huddle Meetings</li>
    </ul>
    <br>
    <p>Please find below the SOP to share a Powerpoint Live presentation in MS Teams.</p>
    <br><br>
    <h1 style="font-weight:bold; text-align:center">SOP to share a Powerpoint Live presentation in MS Teams</h1>
    <br>
    <h3><strong>Step1 - </strong> if the PPT you want to present is on Google Drive, download it to your computer as a MS powerpoint file.</h3>
    <img src="cid:img1">
    <br>
    <h3><strong>Step2 - </strong> now open MS Teams and attach the file to the meeting chat area.</h3>
    <img src="cid:img2">
    <br>
    <h3><strong>Step3 - </strong> you'll get the option to "Upload from this device", select the downloaded file to share with the participants.</h3>
    <img src="cid:img3">
    <br>
    <h3><strong>Step4 - </strong> now start/join the meeting and click on "Share" button to share your content.</h3>
    <img src="cid:img4">
    <br>
    <h3><strong>Step5 - </strong> click on "Browse my computer" and select the ppt file.</h3>
    <img src="cid:img5">
    <br>
    <h3><strong>Step6 - </strong> use arrow keys on your keyboard to navigate through the slides, press the "B" key to show/hide speaker notes. Additionally you can highlight or point to specific areas during the presentation.</h3>
    <img src="cid:img6">
    <br><br>
    <p>SOP link- https://docs.google.com/document/d/1YFHi2FXZckeyR9sJWn2ujpUixoagYIqF6M0sIzSEg4o/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function laptopBag () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Laptop Bag Maintenance Policy";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>In our ongoing efforts to ensure the efficient use of company resources, we have implemented a Laptop Bag Maintenance Policy. This policy is designed to streamline the issuance and replacement of laptop bags for our team members who use them for various purposes, including site visits, fieldwork, and regular office use.</p>
    <h3><u>Policy Details:</u></h3>
    <ol>
      <li><strong>Eligibility for Laptop Bag Issuance:</strong><br>
      &nbsp; - Laptop bags will only be issued to employees who possess laptops and engage in site visits or other fieldwork as part of their job responsibilities.
      </li>
      <li><strong>Bag Replacement for Field and Site Personnel:</strong><br>
      &nbsp; - Laptop bags used exclusively for fieldwork and site visits will be eligible for replacement after a period of six (6) months from the date of issue. Replacement will be considered for bags that have undergone wear and tear due to work-related activities.
      </li>
      <li><strong>Bag Replacement for Regular Users:</strong><br>
      &nbsp; - Laptop bags used for regular office purposes will be eligible for replacement after a period of one (1) year from the date of issue. Replacement will be considered for bags that have experienced normal wear and tear during daily use.
      </li>
      <li><strong>Bag Repair Responsibility:</strong><br>
      &nbsp; - Prior to the eligible replacement period, employees are responsible for repairing their bags in case of any damage or issues encountered. The company will not issue new bags within the first six (6) months for field and site personnel, and within the first year for regular users unless there are exceptional circumstances.
      </li>
    </ol>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function masterPPTnGuidelines() {
  const imgBlob = DriveApp.getFilesByName("master-ppt-highlighter-tool-screenshot1.png").next().getBlob();
  let inlineImages = {'img1': imgBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Master PPT Template and Project Documentation Guidelines";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>As part of our ongoing efforts to streamline our processes and ensure consistency across all projects, we are implementing a new system for maintaining and presenting project information. Please take note of the following instructions and make the necessary adjustments to your workflow.</p>
    <br>
    <p style="font-weight:bold">Master PPT Template & Highlighting Instructions:</p>
    <ul>
      <li>We have created a Master PPT template that must be used for all internal and external Kick-Off Meetings (KOM). This template will help us maintain uniformity and readiness for approvals, management meetings, and client interactions.</li>
      <li>Link- https://docs.google.com/presentation/d/1yenQ1feIVU-pZkE_hyultwK-Wk4VjuuH8MU1MQkFHXs/edit?usp=sharing</li>
      <li>The linked Master PPT Template provides a user-friendly framework with clear instructions on required information. To ensure clarity and ownership of content, we recommend using the highlighter tool in Google Slides to highlight specific sections within the template according to the following guidelines:
      <ul>
        <li>Yellow Fields: To be filled internally.</li>
        <li>Blue Fields: To be assumed or gathered from the Project Proponent (PP).</li>
        <li>Green Fields: To be filled by the PP.</li>
        <li>Red Fields: Indicate any changes made.</li>
        <li>If all fields are transparent, it means the document is final.</li>
      </ul>
      </li>
    </ul>
    <img src="cid:img1">
    <br>
    <p style="font-weight:bold">Project Management & SOPs:</p>
    <ul>
      <li>Please make a habit of reviewing all admin emails, including those containing updates on Standard Operating Procedures (SOPs). Maintaining updated knowledge of SOPs is crucial for efficient project execution.</li>
      <li>For continuous improvement in project data management:
      <ul>
        <li>Project sheets should remain in use for internal calculations.</li>
        <li>While Master PPT, consolidating all relevant data into a client-friendly format to be used for PP approvals, management meetings etc.</li>
      </ul>
      </li>
    </ul>
    <br>
    <p style="font-weight:bold">Client Needs:</p>
    <ul>
      <li>To ensure we can promptly fulfill client requests, please be prepared to provide the following:
      <ul>
        <li><strong>NABET Accreditation Information: </strong> When needed, have readily available NABET accreditation letters and extensions for projects.</li>
        <li><strong>Project Chronology: </strong> Maintain a documented record within the master PPT outlining key project milestones, such as TOR application, grant, Public Hearing (PH), EIA submission, and Environmental Clearance (EC) letter.</li>
        <li><strong>Master PPT Links: </strong> Keep the link to the master PPT containing this information updated in the relevant chat space for each project.</li>
        <li><strong>Key Milestone Dates: </strong> For future projects, TF05 02 and 08 to document key milestone dates and also within the master PPT to facilitate faster response to client inquiries.</li>
      </ul>
      </li>
    </ul>
    <br>
    <p>A master list of all forms and formats is available for your reference.  Please refer to this list when needed to ensure you're using the most recent and appropriate documents.</p>
    <p>By adhering to these practices and utilizing the provided resources, we can achieve improved project management efficiency, streamlined client communication, and a more prepared team.</p>
    <p>If you have any questions, please don't hesitate to ask.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc,
    inlineImages: inlineImages
  });
}

function tg03 () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] TG03 - Technical Guidelines for EIA/EAC Applications";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the updated TG03 - Technical Guidelines for EIA/EAC applications in a single file format. This standardized format is designed to streamline the submission process for building construction, mining and industry projects.</p>
    <p>This standardized format ensures consistency and reduces the likelihood of errors.</p>
    <p>Please adhere to this format for all EIA/EAC applications.</p>
    <p>You can access the TG03 format using the following link: [https://docs.google.com/spreadsheets/d/13kqInum8kr5DXmr0bdQjObPARSiLJ7dENf5N-vjv-Yo/edit?usp=sharing]</p>
    <p>If you have any questions or require assistance, please don't hesitate to reach out.</p>
    <p>We appreciate your cooperation in adopting this new format.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function officeSupplies () {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Responsible Use and Care of Office Supplies";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>We invest heavily in providing you with the tools you need to succeed at your job. This includes laptops, chargers, mice, and other essential office supplies.</p>
    <p>These items are a significant company investment, and we expect everyone to treat them with care and responsibility. Unfortunately, lost or damaged equipment due to negligence will not be replaced at company expense.</p>
    <p>By taking accountability of our equipment, we ensure everyone has the resources they need to do their jobs effectively. It also helps us manage costs and avoid unnecessary delays due to equipment shortages.</p>
    <p>In instances where equipment is damaged or lost due to unforeseen circumstances, please report the incident immediately to the Admin team. They will assess the situation on a case-by-case basis.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
  });
}

function activeZohoForms() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Master list of all active ZOHO forms";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>Please find below the Google Sheet containing links and purposes of all ZOHO forms in circulation.</p>
    <br>
    <p>This sheet provides centralised resource for various types of forms, including:</p>
    <ul>
      <li>TFs</li>
      <li>BD Forms</li>
      <li>Admin Forms</li>
      <li>Accounting Forms</li>
      <li>HR Forms</li>
      <li>performance forms</li>
    </ul>
    <br>
    <p><strong>Please note: </strong> <em> This document will keep updating as and when a new form is finalized. So we encourage everyone to bookmark it for quick access to the forms they need.</em></p>
    <br>
    <p>Link to the sheet- https://docs.google.com/spreadsheets/d/1ScMYMmoUCCmHZlGX6VHCzlrUpkZMbrG-cd8yGiytCh4/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function conferenceRoom() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Streamlining VC meetings from conference room";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>To ensure efficient use of our conference room and maximize support for video conferencing, kindly follow the below mentioned points:</p>
    <ul>
      <li>Whenever you reserve the conference room for a meeting, please share all necessary details, such as meeting links and documents, to “vc@perfactgroup.in” via Email or Google Chat at least 2 hours prior to the meeting time.</li>
      <li>Download data on the desktop for easy access during the meeting.</li>
      <li>Remove all the downloaded data after the meeting.</li>
    </ul><br>
    <p><strong>Please note that the “vc@perfactgroup.in” email is logged in by default on the Mac Mini in Conference room. Please refrain from using any other ID for login purposes, and conduct all meetings using this default ID.</strong></p>
    <p>All essential settings, including Audio & Video configurations, have already been set up for MS Teams & Google Meet on the Mac Mini installed in the Conference room. Kindly avoid making any changes to these settings. If the meeting needs to be done through any other app, the Admin team will do the necessary settings.</p>
    <p>By sharing these details in advance, we can ensure a seamless and productive meeting experience for everyone.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}

function guidelinesSmoothCheckIn() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Guidelines for a Smooth Check-in Process for Train and Flight Travel";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>I hope this email finds you well, as many of our employees travel for work and to ensure everyone's safety, we have prepared some guidelines for those traveling by flight or train.</p>
    <p>These guidelines are designed to minimize risks and make your journey as hassle-free as possible.</p>
    <p>Please take the time to carefully review and follow these guidelines to ensure your safety and the safety of others.</p>
    <p>Your cooperation is greatly appreciated and will help ensure successful and safe business trips.</p>
    <br>
    <p><strong>Please see the attached PDFs for Travel related Guidelines:</strong></p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc,
    attachments:[
      DriveApp.getFilesByName("travel-guideline-for-employees-recurring.pdf").next().getBlob(),
      DriveApp.getFilesByName("train-travel-guideline-recurring.pdf").next().getBlob(),
      DriveApp.getFilesByName("flight-travel-guideline-recurring.pdf").next().getBlob()
      ]
    });
}

function purchaseForm () {
  const imgBlob1 = DriveApp.getFilesByName("sop-for-purchase-request-form-1.png").next().getBlob();
  const imgBlob2 = DriveApp.getFilesByName("sop-for-purchase-request-form-2.png").next().getBlob();
  const imgBlob3 = DriveApp.getFilesByName("sop-for-purchase-request-form-3.png").next().getBlob();
  const imgBlob4 = DriveApp.getFilesByName("sop-for-purchase-request-form-4.png").next().getBlob();
  const imgBlob5 = DriveApp.getFilesByName("sop-for-purchase-request-form-5.png").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1,
    'img2': imgBlob2,
    'img3': imgBlob3,
    'img4': imgBlob4,
    'img5': imgBlob5,
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "[Reminder] SOP for Office Essential Purchase Approval Process";
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>Please find the SOP below for the newly implemented approval process to purchase office essential items.</p>
    <br><br>
    <h2 style="font-weight:bold; text-align:center">SOP for Office Essential Purchase Approval Process</h2>
    <br>
    <p><strong>Step1 - </strong> Access the Purchase form and fill in the required details and information about the item for purchase.</p>
    <img src="cid:img1">
    <br>
    <p><strong>Step2 - </strong> On form submission an email will be sent to your respective Team Head with a button to respond to the request.</p>
    <img src="cid:img2">
    <br>
    <p><strong>Step3 - </strong> On clicking the response button a new page will open up where the TH can accept / reject the request after identifying the need for the new purchase. They can also add a comment to elaborate on their decision.</p>
    <img src="cid:img3">
    <br>
    <p><strong>Step4 - </strong> If rejected a mail will be sent to the concerned employee informing them that their request has been rejected with TH's comment stating the reason. Else on acceptance, a mail will be sent to the Admin Council- Purchase WG.</p>
    <img src="cid:img4">
    <br>
    <p><strong>Step5 - </strong> The Purchase WG will then finalise the vendor (refer approved vendor list- https://docs.google.com/spreadsheets/d/1wiPocloQtPVNuJWBAEbLX59jXUjCyvLb8Of_agPukeM/edit?usp=sharing) and costing for the item and make their decision after consulting the Admin Council or the Governing Council as applicable. Working Group will then respond by accepting / rejecting the request and adding their input as comment.</p>
    <img src="cid:img3">
    <br>
    <p><strong>Step6 - </strong> Finally the request will reach the Accounts department with all the details, who will then have the responsibility of placing PO, receiving the item and making payment.</p>
    <img src="cid:img5">
    <br><br>
    <p>Form link- https://forms.gle/rvJDo5kRUhrQYNBU7</p>
    <p>SOP link- https://docs.google.com/document/d/1Xaw75FoTtgZLIw8Xq4XFNVwwDu4whV9Gj96zsoe6ALQ/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function emailEtiquette () {
  const imgBlob1 = DriveApp.getFilesByName("importance-of-correct-email-etiquettes-hr-24-01-2023.png").next().getBlob();
  let inlineImages = {
    'img1': imgBlob1
  }
  const name = "IT Administrator / PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = `[Reminder] Importance of correct Email Etiquette`;
  const body = `
  <head>
    <style>
      @media screen and (min-width: 767px){
        img {
          width: 50%
        }
      }
    </style>
  </head>
  <body>
    <p>Dear all,</p>
    <p>I hope you are doing well. This mail is to reiterate the importance of writing professional and courteous emails to our clients and also among ourselves.</p>
    <p>In our work, writing mails is a major chunk of our communication process. There will be situations when we need to express our opinion and share our thoughts, how we frame and present them (positive and negative alike) ultimately matters. Thus, we should never lose out on basic courtesy and decency while conveying anything.</p>
    <p>Re-sharing the mail along with some useful tips mentioned in the link below.</p>
    <p>https://www.lawsociety.com.au/resources/resources/career-hub/10-rules-email-etiquette</p>
    <p><strong>Please go through the above link thoroughly.</strong></p>
    <p>Our organisation has always held communication and accountability in high regard and have always tried to ingrain these vital values in our people.</p>
    <p>This is therefore no surprise that we have a <strong>strong emailing culture</strong> in our organisation and that the need and use of email and google chat is well understood.</p>
    <p>We would like to take this opportunity to ask everyone to please use the tools provided to us like <strong>Google Chat and Gmail for communication</strong>.</p>
    <br>
    <img src="cid:img1">
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    cc: cc,
    htmlBody: body,
    inlineImages: inlineImages,
    name: name
  });
}

function maintainingHygieneInWorkplace() {
  const recipient = "family@perfactgroup.in";
  const subject = "[Reminder] Maintaining hygiene in the workplace";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>We hope this email finds you well. As a friendly reminder, we would like to bring to your attention the importance of maintaining hygiene in the workplace, especially in the washrooms.</p>
    <p>As a responsible organization, it is essential that we all take necessary steps to ensure the cleanliness and hygiene of our workplace, including the washrooms. Therefore, we would like to request everyone to adhere to the following washroom hygiene etiquettes:</p>
    <ul>
      <li>Always flush the toilet after use and make sure that there is no water left in the basin.</li>
      <li>Wash your hands thoroughly with soap and water after using the washroom.</li>
      <li>Dispose of all used tissues, paper towels, and sanitary products in the bins provided.</li>
      <li>Avoid throwing any non-disposable items like plastics, glass or metals into the toilet bowl.</li>
      <li>Report any maintenance issues like blocked toilets, leaking pipes or faucets to the concerned department immediately.</li>
    </ul>
    <p>Let's make sure that we all take our hygiene and cleanliness seriously and follow these simple washroom etiquette guidelines. This will help create a safe and healthy working environment for all of us.</p>
    <p>Thank you for your attention and cooperation.</p>
    <br>
    <p>नमस्कार सभी,</p>
    <p>हम आशा करते हैं, यह ईमेल आपको कुशलतापूर्वक कार्यरत पाए। एक अनुकूल स्मरण के तौर पर, हम आपको कार्यस्थल, विशेषकर वाशरूम में स्वच्छता बनाए रखने के महत्व की ओर ध्यान दिलाना चाहते हैं।</p>
    <p>एक जिम्मेदार संगठन के रूप में, यह आवश्यक है कि हम सभी अपने कार्यस्थल की साफ-सफाई और स्वच्छता सुनिश्चित करने के लिए आवश्यक कदम उठाएं, जिसमें वाशरूम भी शामिल हैं। इसलिए, हम सभी से अनुरोध करते हैं कि वे निम्नलिखित वाशरूम स्वच्छता शिष्टाचार का पालन करें:</p>
    <ul>
      <li>शौचालय का इस्तेमाल करने के बाद हमेशा फ्लश करें और सुनिश्चित करें कि बेसिन में पानी ना रुका रहे।</li>
      <li>वाशरूम का उपयोग करने के बाद साबुन और पानी से अपने हाथों को अच्छी तरह धो लें।</li>
      <li>उपयोग किए गए टिश्यू, पेपर टॉवल और सैनिटरी उत्पादों को दिए गए डिब्बों में ही फेंके।</li>
      <li>प्लास्टिक, कांच या धातु जैसी किसी भी गैर-अपघटनीय वस्तुओं को शौचालय के गड्ढे में फेंकने से बचें।</li>
      <li>अवरुद्ध शौचालय, टपकते पाइप या नल जैसी किसी भी रखरखाव समस्या को तुरंत संबंधित विभाग को रिपोर्ट करें।</li>
    </ul>
    <p>आइए सुनिश्चित करें कि हम सभी अपनी स्वच्छता को गंभीरता से लें और इन सरल वाशरूम शिष्टाचार दिशानिर्देशों का पालन करें। इससे हम सभी के लिए एक सुरक्षित और स्वस्थ कार्य वातावरण बनाने में मदद मिलेगी।</p>
    <p>आपके ध्यान और सहयोग के लिए धन्यवाद।</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc
    });
}
