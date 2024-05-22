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

  if(currentDate == 8){
    clientVisitChecklistZohoForm();
  }
  else if(currentDate == 10){
    bookMeetingSlotsWithNBPB();
  }
  else if(currentDate == 12){
    setupGmailForOfflineUse();
  }
  else if(currentDate == 13 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    governingCouncilSystem();
  }
  // dates and months are numbers but months are zero indexed
  else if(currentDate == 14){
    saveTeamsRecording();
  }
  else if(currentDate == 15 && (currentMonth == 1 || currentMonth == 4 || currentMonth == 7 || currentMonth == 10)){
    chatAudioVideoCalls();
  }
  else if(currentDate == 18){
    sopForNamingCalendarInvites();
  }
  else if(currentDate == 20){
    gatePassPolicy();
  }
  else if(currentDate == 22){
    masterPPTnGuidelines();
  }
  else if(currentDate == 24){
    activeZohoForms();
  }
  else if(currentDate == 26){
    guidelinesSmoothCheckIn();
  }
  else if(currentDate == 28){
    maintainingHygieneInWorkplace();
  }
}

function clientVisitChecklistZohoForm() {
  const recipient = "family@perfactgroup.in";
  const subject = "Client Visit Checklist Zoho Form";
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

function bookMeetingSlotsWithNBPB() {
  const recipient = "family@perfactgroup.in";
  const subject = "Improved Scheduling for Internal Discussions, Client Meetings & Availability";
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
      <li>This eliminates the need for direct messages, calendar conflicts and forgotten meeting links.</li>
    </ul>
    <br>
    <p style="font-weight:bold">External Client Meetings with NB sir:</p>
    <ul>
      <li>A separate booking page has been created for scheduling external client meetings with NB sir</li>
      <li>Use this link and add required details: https://calendar.app.google/wWiARwVbsvRNd6wCA</li>
      <li>These links have also been added to the "Master List of Forms" sheet for everyone's ease.</li>
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

function setupGmailForOfflineUse() {
  const recipient = "family@perfactgroup.in";
  const subject = "Set up Gmail for Offline Use";
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

function governingCouncilSystem() {
  const recipient = "family@perfactgroup.in";
  const subject = "Governing Council and the updated Council System";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>We're always striving to be at the forefront of innovation in all aspects of business including corporate governance. That's why, after introducing the Council system last year to enhance our decision-making and operational efficiency, we're now excited to launch a Governing Council as the top-most body in the organisation. This move is set to amplify our organization and impact.</p>
    <p>This council will be the operational powerhouse, composed of key figures from across our organization, including top management and other council chairs. Meeting monthly, their goal is to ensure alignment with our objectives, address any challenges, and strategize for the future. For a deeper dive into the changes and the new system, please check out the attached PDF.</p>
    <p>The Governing Council will now oversee 8 major councils, streamlining processes and creating a more balanced environment for efficient resolution. Additionally, we're saying goodbye to sub-councils and task forces in favor of a leaner governance model. The introduction of working groups, taking over the roles of sub-councils, will bring enhanced powers and responsibilities, aiming for efficiency without the need for excessive oversight. With the introduction of this tremendous change, all major councils are requested to conduct a review of all active working groups and sub councils under their supervision, in order to reconstitute according to this new System alongwith their goals, objectives and portfolio division.</p>
    <p>This is a pivotal moment for us to drive change and march forward together. Your feedback is invaluable, and we're eager to hear your thoughts.</p>
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
    attachments:[DriveApp.getFilesByName("latest-council-system-v2.6-2024.pdf").next().getBlob()]
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
  const subject = "Enhance Collaboration: Get acquainted with Google Chat Audio & Video Calls";
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
function sopForNamingCalendarInvites() {
  const recipient = "family@perfactgroup.in";
  const subject = "Standard Operating Procedure (SOP) for Naming Calendar Invites";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
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
    <p>This format should be used for all events/meetings on the calendar to ensure uniformity and clarity of purpose for all attendees.</p>
    <p>Please refer to the attached image for reference.</p>
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
    attachments:[DriveApp.getFilesByName("SOP-for-Naming-Calendar-Invites.png").next().getBlob()]
    });
}

function gatePassPolicy() {
  const recipient = "family@perfactgroup.in";
  const subject = "Implementation of Gate Pass Policy";
  const name = "IT Administrator / PERFACT";
  const cc = "topmanagement@perfactgroup.in";
  const body = `
  <head></head>
  <body>
    <p>Dear all,</p>
    <p>We hope this email finds you well. This is to inform you about the Gate Pass policy which is designed to enhance security aspects and monitor assets and keep a track of the same, streamline access control, and ensure the safety of assets.</p>
    <p>Detailed instructions and access to the gate pass policy is given below. If you have any questions or concerns regarding this policy, please do not hesitate to reach out to the Admin Department - i.e., Vikas Madaan &/or Himanshu Kohli.</p>
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

function masterPPTnGuidelines() {
  const imgBlob = DriveApp.getFilesByName("master-ppt-highlighter-tool-screenshot1.png").next().getBlob();
  let inlineImages = {'img1': imgBlob};
  const recipient = "family@perfactgroup.in";
  const subject = "Master PPT Template and Project Documentation Guidelines";
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
        <li><strong>Key Milestone Dates: </strong> For future projects (TF05 02 and TF08 onward), document key milestone dates within the master PPT to facilitate faster response to client inquiries.</li>
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

function activeZohoForms() {
  const recipient = "family@perfactgroup.in";
  const subject = "Master list of all active ZOHO forms";
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

function guidelinesSmoothCheckIn() {
  const recipient = "family@perfactgroup.in";
  const subject = "Guidelines for a Smooth Check-in Process for Train and Flight Travel";
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

function maintainingHygieneInWorkplace() {
  const recipient = "family@perfactgroup.in";
  const subject = "Maintaining hygiene in the workplace";
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
