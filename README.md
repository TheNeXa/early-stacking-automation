# Container Yard Request Automator

## What It Does
The **Container Yard Request Automator** is a Google Apps Script project that simplifies the process of submitting early stacking requests for shipping terminals. By tapping into vessel schedules, it handles the heavy lifting of preparing and emailing requests, complete with Excel attachments, to keep container yard operations running smoothly. Paired with a sleek HTML sidebar (`SubsequentEarlyStackingUI.html`), it offers a handy interface for manual follow-ups.

## Highlights
- **Daily Kickoff**: Fires off initial stacking requests every day at noon, bundled with Excel files for terminal crews.
- **Follow-Up Flexibility**: Tackles unprocessed containers with automated or manual resends via the sidebar.
- **Schedule Smarts**: Cross-checks requests against vessel timelines, flagging anything too early (outside a 5-hour pre-arrival window).
- **Tailored Outreach**: Crafts emails with formatted tables and attachments, customized for different terminals (e.g., JICT, KOJA, MAL).
- **Progress Tracking**: Logs each vessel’s send history, capping at 20 rounds.
- **User-Friendly Controls**: Lets you pick vessels and sending stages manually through a sidebar interface.

## Code Peek
### Core Workflow
```javascript
function handleStackingRequests() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Early Data");
  const vesselGroups = {};
  rows.forEach((row, idx) => {
    if (isWithinTimeWindow(requestTime)) {
      vesselGroups[vessel].requests.push({ booking, container, type, pod, weight, gate });
      sheet.getRange(idx + 2, 13).setValue("1st Sending");
    }
  });
  Object.keys(vesselGroups).forEach(vessel => sendRequestEmail(vessel, terminal, vesselGroups[vessel].requests, 1));
}
```

### Email Dispatch
```javascript
function sendRequestEmail(vessel, terminal, requests, sendCount) {
  const subject = `(${sendCount}${getOrdinal(sendCount)} Sending) Early Stacking Request - ${vessel}`;
  const attachment = generateExcelFile(vessel, requests, sendCount);
  MailApp.sendEmail({
    to: "terminal@example.com",
    subject: subject,
    htmlBody: buildEmailBody(requests),
    attachments: [attachment]
  });
}
```

### Sidebar Setup
```html
<select id="vesselDropdown" onchange="refreshOptions()"></select>
<button onclick="triggerSend()">Send Pending Requests</button>
<script>
  google.script.run.withSuccessHandler(data => {
    const dropdown = document.getElementById('vesselDropdown');
    data.forEach(v => dropdown.innerHTML += `<option value="${v}">${v}</option>`);
  }).fetchVesselList();
</script>
```

## Why It Matters
Perfect for logistics teams juggling container yard tasks, this tool:
- **Saves Time**: Cuts out repetitive grunt work like data entry and email drafting.
- **Keeps It Tight**: Matches requests to vessel schedules to dodge errors.
- **Boosts Clarity**: Delivers polished, attachment-ready emails to terminal staff.
- **Scales Up**: Manages multiple terminals and vessels with ease.

Think of it as a logistics sidekick—streamlining operations, dodging delays, and letting teams focus on the big picture.

## Getting Started
1. **Spreadsheets Needed**: Set up `Early Data`, `Vessel`, and `SendingLog`.
2. **Automation**: Run `setupTriggers()` for daily noon execution.
3. **Access**: Enable Spreadsheet, Drive, and Mail permissions in Apps Script.

## How to Use
- **Hands-Off**: Daily sends kick in at 12 PM.
- **Hands-On**: Use the custom menu or sidebar for manual control.

This project is all about making container yard logistics faster, smarter, and more reliable.
