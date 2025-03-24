# Early Stacking Automation

## Overview
The **Early Stacking Automation** is a Google Apps Script solution designed to streamline container yard requests for early stacking at shipping terminals. It automates the processing and emailing of requests based on vessel schedules, ensuring timely submissions while reducing manual effort. The accompanying `SubsequentEarlyStackingUI.html` provides a user-friendly sidebar interface for manual subsequent sendings.

## Features
- **Automated 1st Sending**: Processes early stacking requests daily at 12 PM, sending emails with Excel attachments to terminal teams.
- **Subsequent Sending**: Handles follow-up requests for unprocessed containers, triggered daily or manually via UI.
- **Vessel Validation**: Checks requests against vessel schedules, rejecting those outside a 5-hour pre-opening window.
- **Email Customization**: Tailors email content based on terminal (JICT, KOJA, MAL) with formatted tables and attachments.
- **Logging**: Maintains a sending log to track sendings per vessel (up to 20 iterations).
- **UI Sidebar**: Allows manual selection of vessels and sending numbers for subsequent requests.

## Key Code Snippets
### Main Automation Logic
```javascript
function processEarlyStackingRequests() {
  // Process requests, validate against vessel data, and send emails
  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Early Data");
  const requestsByVessel = {};
  rows.forEach((row, index) => {
    if (requestTime < windowStart) {
      requestsByVessel[vesselName].requests.push({ bookingNumber, containerNo, type, pod, weight, gateIn });
      dataSheet.getRange(index + 2, 13).setValue("1st Sending");
    }
  });
  for (const vesselName in requestsByVessel) {
    sendEarlyStackingEmail(vesselName, terminal, requests, 1);
  }
}
```

### Email Sending with Attachment
```javascript
function sendEarlyStackingEmail(vesselName, terminal, newRequests, sendingNumber) {
  const subject = `(${sendingNumber}${ordinalSuffix.toUpperCase()} SENDING) REQUEST EARLY STACKING - ${vesselName}`;
  const excelBlob = createExcelAttachment(vesselName, allRequests, sendingNumber);
  MailApp.sendEmail({
    to: "dimasalif5@gmail.com",
    subject: subject,
    htmlBody: htmlBody,
    attachments: [excelBlob]
  });
}
```

### UI Integration
```html
<select id="vesselSelect" onchange="updateSending()"></select>
<button onclick="sendEmail()">Send All Unprocessed Requests</button>
<script>
  google.script.run.withSuccessHandler(vessels => {
    const vesselSelect = document.getElementById('vesselSelect');
    vessels.forEach(vessel => vesselSelect.innerHTML += `<option value="${vessel}">${vessel}</option>`);
  }).getVessels();
</script>
```

## Business Use Case
This tool is ideal for logistics and shipping companies managing container yard operations. It:
- **Improves Efficiency**: Automates repetitive tasks, reducing manual data entry and email drafting.
- **Ensures Compliance**: Validates requests against vessel schedules, minimizing errors and rejections.
- **Enhances Communication**: Provides clear, professional emails with attachments for terminal teams.
- **Scales Operations**: Handles multiple vessels and terminals (JICT, KOJA, MAL) with tailored workflows.

For a company like PT Ocean Network Express Indonesia, this reduces operational delays, ensures containers are stacked early when needed, and frees staff to focus on higher-value tasks.

## Setup
1. **Sheets Required**: `Early Data`, `Vessel`, `SendingLog`.
2. **Triggers**: Run `setupTriggers()` to schedule daily automation at 12 PM.
3. **Permissions**: Grant Apps Script access to Spreadsheet, Drive, and Mail services.

## Usage
- **Automatic**: 1st and subsequent sendings run daily at 12 PM.
- **Manual**: Use "Early Stacking Tools" menu or sidebar UI for on-demand sending.

This solution optimizes container yard coordination, saving time and improving reliability in logistics operations.
