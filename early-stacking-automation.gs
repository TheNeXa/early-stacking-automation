function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create "Early Stacking Tools" menu
  ui.createMenu('Early Stacking Tools')
    .addItem('Manual Send (1st Sending)', 'manualSendEarlyStackingRequests')
    .addItem('Subsequent Sending', 'openSubsequentEarlyStackingUI')
    .addToUi();

  // Create "Extend Closing Tools" menu
  ui.createMenu('Extend Closing Tools')
    .addItem('Manual Send (1st Sending)', 'manualSendExtendClosingRequests')
    .addItem('Subsequent Sending', 'openSubsequentSendingUI')
    .addToUi();
}

function processEarlyStackingRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Early Data");
  const vesselSheet = ss.getSheetByName("Vessel");
  const sendingLogSheet = ss.getSheetByName("SendingLog") || ss.insertSheet("SendingLog");

  // Initialize SendingLog sheet if it doesn't have headers
  const logHeaders = sendingLogSheet.getRange(1, 1, 1, 3).getValues()[0];
  if (logHeaders[0] !== "Vessel Name") {
    sendingLogSheet.getRange(1, 1, 1, 3).setValues([["Vessel Name", "Last Sending Number", "Last Sending Time"]]);
  }

  if (!dataSheet || !vesselSheet) {
    Logger.log("Error: Sheets not found");
    return { processed: 0, rejected: 0, emailsSent: 0 };
  }

  // Ensure "Sending" column exists (column 13, after "Approve Status")
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  if (headers[12] !== "Sending") {
    dataSheet.insertColumnAfter(12);
    dataSheet.getRange(1, 13).setValue("Sending");
  }

  const data = dataSheet.getDataRange().getDisplayValues();
  const vesselData = vesselSheet.getDataRange().getDisplayValues();
  const rows = data.slice(1);
  const vesselRows = vesselData.slice(1);

  const vesselMap = {};
  vesselRows.forEach(row => {
    vesselMap[row[0]] = {
      opening: new Date(`${row[1]} GMT+0700`),
      terminal: row[3]
    };
  });

  const requestsByVessel = {};
  let processedCount = 0;
  let rejectedCount = 0;

  rows.forEach((row, index) => {
    const [idRequest, timestamp, email1, email2, vesselName, bookingNumber, containerNo, type, pod, weight, gateIn, approveStatus, sending] = row;
    if (approveStatus || sending) return; // Skip if already approved or sent

    const requestTime = new Date(`${timestamp} GMT+0700`);
    const vesselInfo = vesselMap[vesselName];
    if (!vesselInfo) {
      dataSheet.getRange(index + 2, 13).setValue("Rejected - Vessel not found");
      rejectedCount++;
      return;
    }

    const { opening, terminal } = vesselInfo;
    const windowStart = new Date(opening.getTime() - 5 * 60 * 60 * 1000); // 5 hours before opening

    // Process requests BEFORE the 5-hour window, reject DURING or AFTER
    if (requestTime < windowStart) {
      if (!requestsByVessel[vesselName]) {
        requestsByVessel[vesselName] = { terminal, requests: [] };
      }
      requestsByVessel[vesselName].requests.push({
        bookingNumber: bookingNumber,
        containerNo: containerNo,
        type: type,
        pod: pod,
        weight: weight,
        gateIn: gateIn,
        rowIndex: index + 2
      });
      dataSheet.getRange(index + 2, 13).setValue("1st Sending");
      processedCount++;
    } else {
      dataSheet.getRange(index + 2, 13).setValue("Rejected");
      rejectedCount++;
    }
  });

  let emailsSent = 0;
  const currentTime = new Date();
  for (const vesselName in requestsByVessel) {
    const { terminal, requests } = requestsByVessel[vesselName];
    sendEarlyStackingEmail(vesselName, terminal, requests, 1);

    // Log the 1st Sending in SendingLog
    const logData = sendingLogSheet.getDataRange().getDisplayValues();
    const vesselLog = logData.find(row => row[0] === vesselName);
    if (!vesselLog) {
      sendingLogSheet.appendRow([vesselName, 1, currentTime.toISOString()]);
    } else {
      // Update existing entry (in case the vessel already exists)
      const logIndex = logData.findIndex(row => row[0] === vesselName);
      sendingLogSheet.getRange(logIndex + 1, 2).setValue(1);
      sendingLogSheet.getRange(logIndex + 1, 3).setValue(currentTime.toISOString());
    }

    emailsSent++;
  }

  Logger.log("Early Stacking Summary - Processed: %d, Rejected: %d, Emails Sent: %d", processedCount, rejectedCount, emailsSent);
  return { processed: processedCount, rejected: rejectedCount, emailsSent: emailsSent };
}

function manualSendEarlyStackingRequests() {
  const summary = processEarlyStackingRequests();
  SpreadsheetApp.getUi().alert(`Manual Send Complete\n\nProcessed: ${summary.processed}\nRejected: ${summary.rejected}\nEmails Sent: ${summary.emailsSent}`);
}

function processSubsequentEarlyStackingRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Early Data");
  const vesselSheet = ss.getSheetByName("Vessel");
  const sendingLogSheet = ss.getSheetByName("SendingLog");

  if (!dataSheet || !vesselSheet || !sendingLogSheet) {
    Logger.log("Error: Sheets not found");
    return;
  }

  const logData = sendingLogSheet.getDataRange().getDisplayValues().slice(1);
  const data = dataSheet.getDataRange().getDisplayValues();
  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);

  const vesselMap = {};
  vesselData.forEach(row => {
    vesselMap[row[0]] = { terminal: row[3] };
  });

  // Process vessels that have had at least one sending
  logData.forEach((logRow, logIndex) => {
    const vesselName = logRow[0];
    const lastSendingNumber = parseInt(logRow[1]) || 0;
    const lastSendingTime = logRow[2] ? new Date(logRow[2]) : null;

    // Skip if no previous sending or if we've reached the maximum sending number (20)
    if (!lastSendingNumber || !lastSendingTime || lastSendingNumber >= 20) return;

    // Check if it's the next day after the last sending
    const currentTime = new Date();
    const timeDiff = (currentTime - lastSendingTime) / (1000 * 60 * 60 * 24); // Difference in days
    if (timeDiff < 1) return; // Wait until the next day

    // Find new unprocessed requests for this vessel
    const newRequests = [];
    const rowsToUpdate = [];
    data.slice(1).forEach((row, index) => {
      const [idRequest, timestamp, , , rowVessel, bookingNumber, containerNo, type, pod, weight, gateIn, approveStatus, sending] = row;
      if (rowVessel === vesselName && !approveStatus && !sending) {
        newRequests.push({ bookingNumber, containerNo, type, pod, weight, gateIn });
        rowsToUpdate.push(index + 2);
      }
    });

    if (newRequests.length > 0) {
      // Calculate the next sending number
      const nextSendingNumber = lastSendingNumber + 1;

      // Collect all previously sent requests (exclude rejected)
      const allRequests = data.slice(1)
        .filter(row => row[4] === vesselName && row[12] && row[12] !== "Rejected")
        .map(row => ({
          bookingNumber: row[5],
          containerNo: row[6],
          type: row[7],
          pod: row[8],
          weight: row[9],
          gateIn: row[10]
        }))
        .concat(newRequests);

      const terminal = vesselMap[vesselName]?.terminal || "Unknown";
      sendEarlyStackingEmail(vesselName, terminal, newRequests, nextSendingNumber, allRequests);

      // Mark the requests with the next sending number
      rowsToUpdate.forEach(rowNum => {
        dataSheet.getRange(rowNum, 13).setValue(`${nextSendingNumber}${getOrdinalSuffix(nextSendingNumber)} Sending`);
      });

      // Update SendingLog with the new sending number and time
      sendingLogSheet.getRange(logIndex + 2, 2).setValue(nextSendingNumber);
      sendingLogSheet.getRange(logIndex + 2, 3).setValue(currentTime.toISOString());
      SpreadsheetApp.flush();

      Logger.log("%dth Sending processed for vessel %s: %d requests", nextSendingNumber, vesselName, newRequests.length);
    }
  });
}

function openSubsequentEarlyStackingUI() {
  const html = HtmlService.createHtmlOutputFromFile('SubsequentEarlyStackingUI')
    .setTitle('Subsequent Early Stacking');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getVessels() {
  const vesselSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vessel");
  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);
  return vesselData.map(row => row[0]);
}

function sendSubsequentEarlyStackingEmail(vessel, sendingNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Early Data");
  const vesselSheet = ss.getSheetByName("Vessel");

  const data = dataSheet.getDataRange().getDisplayValues();
  const newRequests = [];
  const rowsToUpdate = [];

  // Collect new unprocessed requests (not approved, not sent)
  data.slice(1).forEach((row, index) => {
    const [idRequest, , , , rowVessel, bookingNumber, containerNo, type, pod, weight, gateIn, approveStatus, sending] = row;
    if (rowVessel === vessel && !approveStatus && !sending) {
      newRequests.push({ bookingNumber, containerNo, type, pod, weight, gateIn });
      rowsToUpdate.push(index + 2);
    }
  });

  // Collect all previously sent or approved requests (exclude rejected)
  const allRequests = data.slice(1)
    .filter(row => row[4] === vessel && row[12] && row[12] !== "Rejected")
    .map(row => ({
      bookingNumber: row[5],
      containerNo: row[6],
      type: row[7],
      pod: row[8],
      weight: row[9],
      gateIn: row[10]
    }))
    .concat(newRequests);

  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);
  const terminal = vesselData.find(row => row[0] === vessel)[3];

  if (newRequests.length > 0) {
    sendEarlyStackingEmail(vessel, terminal, newRequests, sendingNumber, allRequests);

    rowsToUpdate.forEach(rowNum => {
      dataSheet.getRange(rowNum, 13).setValue(`${sendingNumber}${getOrdinalSuffix(sendingNumber)} Sending`);
    });
    SpreadsheetApp.flush();
  }

  return newRequests.length;
}

function sendEarlyStackingEmail(vesselName, terminal, newRequests, sendingNumber, allRequests = newRequests) {
  const recipient = "dimasalif5@gmail.com";
  const ordinalSuffix = getOrdinalSuffix(sendingNumber);
  const subject = `(${sendingNumber}${ordinalSuffix.toUpperCase()} SENDING) REQUEST EARLY STACKING - ${vesselName}`;

  // Create Excel attachment with all requests (sent or approved, excluding rejected)
  const tempSpreadsheet = SpreadsheetApp.create(`${vesselName}_${sendingNumber}${ordinalSuffix}_EarlyStacking_${Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd_HHmmss")}`);
  const sheet = tempSpreadsheet.getActiveSheet();

  const headers = ["Vessel Name", "Booking Number", "Container No", "Type", "POD", "Weight (Kg)", "Gate in Request"];
  sheet.appendRow(headers);

  allRequests.forEach(req => {
    sheet.appendRow([vesselName, req.bookingNumber || "", req.containerNo || "", req.type || "", req.pod || "", req.weight || "", req.gateIn || ""]);
  });

  const totalRows = allRequests.length + 1;
  const totalCols = headers.length;

  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange
    .setBackground("#BD0F72")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  const dataRange = sheet.getRange(2, 1, totalRows - 1, totalCols);
  dataRange
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");

  for (let col = 1; col <= totalCols; col++) {
    sheet.autoResizeColumn(col);
    const currentWidth = sheet.getColumnWidth(col);
    if (currentWidth < 100) sheet.setColumnWidth(col, 100);
  }
  sheet.setFrozenRows(1);

  const fileId = tempSpreadsheet.getId();
  const exportUrl = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet&alt=media`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const excelBlob = response.getBlob().setName(`${vesselName}_${sendingNumber}${ordinalSuffix}_EarlyStacking.xlsx`);

  // Email body (only new requests, excluding approved)
  let htmlBody = "";
  if (terminal === "JICT") {
    htmlBody = `
      <p><b>Dear Pak H. Bowo / Pak Endras / Pak Lajumadi / Pak Syawal / Pak Faizal & JICT team,</b><br>
         Yard & Ship Planning,<br>
         Billing & Gate team,</p>
      <p>Please assist to accept below ONE early stacking units on the subject vessel with details as follows:</p>`;
  } else if (terminal === "KOJA") {
    htmlBody = `
      <p><b>Dear Koja Planning,</b><br>
         <b>Billing Team,</b><br>
         <b>SSL Team,</b></p>
      <p>Please kindly assist to accept early stacking requests on the subject vessel.</p>`;
  } else if (terminal === "MAL") {
    htmlBody = `
      <p><b>Dear MAL Team,</b><br>
         <b>SPV Planning & Operations,</b></p>
      <p>Please assist to accept below ONE early stacking units on the subject vessel.</p>`;
  }

  htmlBody += `
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
      <tr style="background-color: #bd0f72; color: white;">
        <th><b>Vessel Name</b></th>
        <th><b>Booking Number</b></th>
        <th><b>Container No</b></th>
        <th><b>Type</b></th>
        <th><b>POD</b></th>
        <th><b>Weight (Kg)</b></th>
        <th><b>Gate in Request</b></th>
      </tr>
      ${newRequests.length > 0 ? newRequests.map(req => `
        <tr>
          <td>${vesselName}</td>
          <td>${req.bookingNumber || ""}</td>
          <td>${req.containerNo || ""}</td>
          <td>${req.type || ""}</td>
          <td>${req.pod || ""}</td>
          <td>${req.weight || ""}</td>
          <td>${req.gateIn || ""}</td>
        </tr>`).join('') : '<tr><td colspan="7">No new requests to display.</td></tr>'}
    </table>
    <p>For your convenience, the same data is also attached as an Excel file.</p>
    <p><span style="background-color: yellow;"><b>All costs incurred will be under shipper's responsibility.</b></span></p>
    <p>Appreciate your approval for the above early stacking request.<br>Thank you</p>
    <p>Best Regards,</p>
    <p style="color: #bd0f72; font-size: 18px;">AS ONE, WE CAN.</p>
    <p>□--------------------------------------------□<br>
    Product & Network<br>EQC & MnR | Vessel Operations<br>□--------------------------------------------□<br>
    <b>PT OCEAN NETWORK EXPRESS INDONESIA</b><br>AIA Central | 8th & 22nd Floor<br>Jl. Jenderal Sudirman Kav. 48A,<br>Jakarta Selatan 12930 - Indonesia<br>
    Phone number: (021) 50815150<br>Dialpad :+62-31-9920-6819<br>www.one-line.com</p>`;

  Logger.log("Sending email to %s - Subject: %s", recipient, subject);
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
    attachments: [excelBlob]
  });

  DriveApp.getFileById(fileId).setTrashed(true);
}

function getOrdinalSuffix(number) {
  const suffixes = ["th", "st", "nd", "rd"];
  const lastTwoDigits = number % 100;
  const lastDigit = number % 10;
  return (lastTwoDigits >= 11 && lastTwoDigits <= 13) ? "th" : suffixes[lastDigit] || "th";
}

function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Trigger for 1st Sending (daily at 12 PM)
  ScriptApp.newTrigger("processEarlyStackingRequests")
    .timeBased()
    .everyDays(1)
    .atHour(12) // 12 PM in Asia/Jakarta timezone
    .create();

  // Trigger for subsequent sendings (daily at 12 PM)
  ScriptApp.newTrigger("processSubsequentEarlyStackingRequests")
    .timeBased()
    .everyDays(1)
    .atHour(12) // 12 PM in Asia/Jakarta timezone
    .create();

  Logger.log("Triggers set up for 12 PM daily run: 1st and subsequent sendings");
}