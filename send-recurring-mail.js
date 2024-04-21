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
  // dates and months are numbers but months are zero indexed
  
  if(currentDate == 6){
    clientVisitChecklistZohoForm();
  }
  else if(currentDate == 8){
    setupGmailForOfflineUse();
  }
  else if(currentDate == 10){
    saveTeamsRecording();
  }
  else if(currentDate == 16){
    sopForNamingCalendarInvites();
  }
  else if(currentDate == 18){
    gatePassPolicy();
  }
  else if(currentDate == 20){
    activeZohoForms();
  }
  else if(currentDate == 24){
    maintainingHygieneInWorkplace();
  }
  else if(currentDate == 26){
    guidelinesSmoothCheckIn();
  }
}

function clientVisitChecklistZohoForm() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Client Visit Checklist Zoho Form";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>As part of the ongoing efforts to enhance our client engagement and hospitality, we are introducing a new initiative to streamline the arrangements for client visits at our Head Office.</p>
    <br>
    <p>Client Visit Checklist Zoho Form is implemented with immediate effect.</p>
    <br>
    <p>Link of the form: https://zfrmz.com/Obv7IM1JlNpv4BRMoeMw</p>
    <br>
    <p>Each team is required to fill & submit the form ahead of any scheduled client visit. This form is designed to gather essential information about the visit's requirements, ensuring that the Admin team can make the necessary arrangements to provide the best possible hospitality to our valued clients.</p>
    <br>
    <p>This form will enable us to anticipate and address all the necessary details, allowing us to create a positive and seamless experience for our clients during their visit to our Head Office.</p>
    <br>
    <p>Everyone's cooperation in adhering to this new process is greatly appreciated. For any further clarification, feel free to reach out to the Admin Team.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
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

function setupGmailForOfflineUse() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Set up Gmail for Offline Use";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>Hope this email finds you well. Shared below is a useful tool that allows you to use your email offline. This can be particularly helpful if you are traveling or do not have access to a reliable internet connection.</p>
    <br>
    <p>To use Gmail offline, you will need to first enable offline access in your Gmail settings. Here's how:</p>
    <ul>
      <li>Go to the gear icon in the top right corner of your Gmail account and click "Settings".</li>
      <li>In the "General" tab, scroll down to the "Offline" section.</li>
      <li>Click the "Enable offline mail" button.</li>
      <li>Follow the prompts to set up offline access.</li>
      <li>Once you have enabled offline access, you can use Gmail without an internet connection by going to https://mail.google.com/mail/u/0/?ui=2&zy=h in your web browser. You can read, search, and compose emails, as well as archive and delete messages. Any changes you make while offline will be synced with your account the next time you go online.</li>
    </ul>
    <br>
    <p>I hope this is helpful. If you have any questions or need further assistance, don't hesitate to reach out.</p>
    <br>
    <p>For further information, you can access the following link- https://binaryfork.com/use-gmail-offline-3846/</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
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

function saveTeamsRecording () {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
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
    'adminSign': signBlob,
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
  const name = "Administrator IT/ PERFACT";
  const recipient = "family@perfactgroup.in";
  const cc = "topmanagement@perfactgroup.in"
  const subject = "SOP to Save MS Teams Recording";
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
  <br>
  <p>Please find the below (SOP) to save MS team recording in Google Drive.</p>
  <br><br>
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
  <br><br>
  <p>SOP link- https://docs.google.com/document/d/1Y32EYK66j7148GDikQFy0KX3ILedLE1CMOPSCN7ZTHA/edit?usp=sharing</p>
  <p>Google sheet link- https://docs.google.com/spreadsheets/d/1Pn8VHnzqJkhaUp7oNy1q9Ww3jRa9dmLFaUMaxjVuOKA/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
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

function sopForNamingCalendarInvites() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Standard Operating Procedure (SOP) for Naming Calendar Invites";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>Please find the below format/ SOP for Naming Calendar Invites in PG Google Calendar.</p>
    <br>
    <p style="font-weight:bold">Purpose: To provide a consistent format for naming calendar invites that clearly communicates the purpose of the meeting to all attendees.</p>
    <br>
    <p>Components of Calendar invite should be:</p>
    <ul>
      <li><strong>Pcode: </strong> Include a unique identifier for the project or department. This helps to distinguish the meeting from others and makes it easier to search for relevant meetings in the future</li>
      <li><strong>External/Internal: </strong> Indicate whether the meeting is internal or external. This helps attendees know whether the meeting is with colleagues within the organization or with external stakeholders and helps set expectations accordingly</li>
      <li><strong>Objective: </strong> State the objective of the meeting, whether it's to discuss new ideas, resolve specific issues, review progress, or coordinate efforts. For example, "Review," "Initial Meeting," "Team Meeting," "Action Plan," or "Coordination Meeting"</li>
      <li><strong>Agenda: </strong> List the items to be discussed in order of priority or by topic area. This helps attendees prepare for the meeting and ensures that all relevant topics are covered.</li>
    </ul>
    <br>
    <p><strong>Format:</strong> <em>[INTERNAL/EXTERNAL] &lt;PCODE&gt; &lt;Objective&gt; &lt;Agenda&gt;</em></p>
    <p><strong>Example:</strong> <em>[INTERNAL] PE241834 - Review Meeting - FAE Report addition in EIA ; [EXTERNAL] PS241657 - Coordination Meeting - Client Comments on Report ; [INTERNAL] ADM0001 - Weekly Review - Open Points and Freshdesk</em></p>
    <br>
    <p>All other meeting-related details can be added in the meeting notes section of the calendar invite including VC App, old Minutes of meeting, document to be discussed etc.</p>
    <br>
    <p>This format should be used for all events/meetings on the calendar to ensure uniformity and clarity of purpose for all attendees.</p>
    <br>
    <p>Please refer to the attached image for reference.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc,
    inlineImages: inlineImages,
    attachments:[DriveApp.getFilesByName("SOP-for-Naming-Calendar-Invites.png").next().getBlob()]
    });
}

function gatePassPolicy() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Implementation of Gate Pass Policy";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>We hope this email finds you well. This is to inform you about the Gate Pass policy which is designed to enhance security aspects and monitor assets and keep a track of the same, streamline access control, and ensure the safety of assets.</p>
    <br>
    <p>Detailed instructions and access to the gate pass policy is given below. If you have any questions or concerns regarding this policy, please do not hesitate to reach out to the Admin Department - i.e., Vikas Madaan &/or Himanshu Kohli.</p>
    <br>
    <p>Your cooperation and commitment to our security measures are greatly appreciated. We look forward to a smooth transition and a safer work environment for all.</p>
    <br>
    <p>Link the doc- https://docs.google.com/document/d/1D94gnpXgysth7AlRB-vzA0AH73dxDiaICxcJJEkpMDE/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
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

function activeZohoForms() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Master list of all active ZOHO forms";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>Please find below the Google Sheet containing links and purposes of all ZOHO forms in circulation.</p>
    <br>
    <p>This sheet provides centralised resource for various types of forms, including:</p>
    <ul>
      <li>Technical Forms</li>
      <li>Business Development Forms</li>
      <li>Administrative Forms</li>
      <li>Accounting Forms</li>
    </ul>
    <br>
    <p><strong>Please note: </strong> <em> This document will keep updating as and when a new form is finalized. So we encourage everyone to bookmark it for quick access to the forms they need.</em></p>
    <br>
    <p>Link to the sheet- https://docs.google.com/spreadsheets/d/1ScMYMmoUCCmHZlGX6VHCzlrUpkZMbrG-cd8yGiytCh4/edit?usp=sharing</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
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

function maintainingHygieneInWorkplace() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Maintaining hygiene in the workplace";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>We hope this email finds you well. As a friendly reminder, we would like to bring to your attention the importance of maintaining hygiene in the workplace, especially in the washrooms.</p>
    <br>
    <p>As a responsible organization, it is essential that we all take necessary steps to ensure the cleanliness and hygiene of our workplace, including the washrooms. Therefore, we would like to request everyone to adhere to the following washroom hygiene etiquettes:</p>
    <ul>
      <li>Always flush the toilet after use and make sure that there is no water left in the basin.</li>
      <li>Wash your hands thoroughly with soap and water after using the washroom.</li>
      <li>Dispose of all used tissues, paper towels, and sanitary products in the bins provided.</li>
      <li>Avoid throwing any non-disposable items like plastics, glass or metals into the toilet bowl.</li>
      <li>Report any maintenance issues like blocked toilets, leaking pipes or faucets to the concerned department immediately.</li>
    </ul>
    <br>
    <p>Let's make sure that we all take our hygiene and cleanliness seriously and follow these simple washroom etiquette guidelines. This will help create a safe and healthy working environment for all of us.</p>
    <br>
    <p>Thank you for your attention and cooperation.</p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
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

function guidelinesSmoothCheckIn() {
  const signBlob = DriveApp.getFilesByName("admin-signature-for-recurring.jpg").next().getBlob();
  let inlineImages = {'adminSign': signBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Guidelines for a Smooth Check-in Process for Train and Flight Travel";
  const name = "Administrator IT/ PERFACT";
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
    <br>
    <p>I hope this email finds you well, as many of our employees travel for work and to ensure everyone's safety, we have prepared some guidelines for those traveling by flight or train.</p>
    <br>
    <p>These guidelines are designed to minimize risks and make your journey as hassle-free as possible.</p>
    <br>
    <p>Please take the time to carefully review and follow these guidelines to ensure your safety and the safety of others.</p>
    <br>
    <p>Your cooperation is greatly appreciated and will help ensure successful and safe business trips.</p>
    <br>
    <p><strong>Please see the attached PDFs for Travel related Guidelines:</strong></p>
    <br>
    <p>--------------------------</p>
    <p>Thanks & Regards</p>
    <img src="cid:adminSign">
    <br>
  </body>
  `;
  GmailApp.sendEmail(recipient, subject, body, {
    htmlBody: body,
    name: name,
    cc: cc,
    inlineImages: inlineImages,
    attachments:[
      DriveApp.getFilesByName("travel-guideline-for-employees-recurring.pdf").next().getBlob(),
      DriveApp.getFilesByName("train-travel-guideline-recurring.pdf").next().getBlob(),
      DriveApp.getFilesByName("flight-travel-guideline-recurring.pdf").next().getBlob()
      ]
    });
}

// 22. SOP for Recording EAC Meetings via MS Teams (discuss with Shweta ma'am)
// 02. PCODE list (remind Sakhsi Singhal for remaining PCODEs)
// 04. https://chat.google.com/room/AAAAg869g5g/XTVCW5TEB5k/XTVCW5TEB5k?cls=10
// 14.
