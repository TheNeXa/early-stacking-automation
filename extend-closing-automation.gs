function processExtendRequests(triggerType = "manual") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extendSheet = ss.getSheetByName("Extend Data");
  const vesselSheet = ss.getSheetByName("Vessel");

  if (!extendSheet || !vesselSheet) {
    Logger.log("Error: Sheets not found");
    return { processed: 0, rejected: 0, emailsSent: 0 };
  }

  const headers = extendSheet.getRange(1, 1, 1, extendSheet.getLastColumn()).getValues()[0];
  if (headers[12] !== "Sending") {
    extendSheet.insertColumnAfter(12);
    extendSheet.getRange(1, 13).setValue("Sending");
  }

  const extendData = extendSheet.getDataRange().getDisplayValues();
  const rows = extendData.slice(1);

  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);
  const vesselMap = {};
  vesselData.forEach(row => {
    vesselMap[row[0]] = {
      opening: new Date(`${row[1]} GMT+0700`),
      closing: new Date(`${row[2]} GMT+0700`),
      terminal: row[3]
    };
  });

  const requestsByVessel = {};
  let processedCount = 0;
  let rejectedCount = 0;

  rows.forEach((row, index) => {
    const [idRequest, timestamp, email1, email2, gateIn, vessel, booking, container, type, pod, weight, , sending] = row;
    if (sending) return;

    const requestTime = new Date(`${timestamp} GMT+0700`);
    const vesselInfo = vesselMap[vessel];
    if (!vesselInfo) {
      extendSheet.getRange(index + 2, 13).setValue("Rejected - Vessel not found");
      rejectedCount++;
      return;
    }

    const { opening, closing, terminal } = vesselInfo;

    if (requestTime > opening && requestTime < closing) {
      if (!requestsByVessel[vessel]) {
        requestsByVessel[vessel] = { terminal, requests: [] };
      }
      requestsByVessel[vessel].requests.push({ 
        bookingNumber: booking,
        containerNo: container,
        type: type,
        pod: pod,
        weight: weight,
        gateIn: gateIn
      });
      extendSheet.getRange(index + 2, 13).setValue("1st Sending");
      processedCount++;
    } else {
      extendSheet.getRange(index + 2, 13).setValue("Rejected");
      rejectedCount++;
    }
  });

  let emailsSent = 0;
  for (const vessel in requestsByVessel) {
    const { terminal, requests } = requestsByVessel[vessel];
    sendExtendClosingEmail(vessel, terminal, requests, 1);
    emailsSent++;
  }

  Logger.log("Summary - Processed: %d, Rejected: %d, Emails Sent: %d", processedCount, rejectedCount, emailsSent);
  return { processed: processedCount, rejected: rejectedCount, emailsSent: emailsSent };
}

function manualSendExtendClosingRequests() {
  const summary = processExtendRequests("manual");
  SpreadsheetApp.getUi().alert(`Manual Send Complete\n\nProcessed: ${summary.processed}\nRejected: ${summary.rejected}\nEmails Sent: ${summary.emailsSent}`);
}

function openSubsequentSendingUI() {
  const html = HtmlService.createHtmlOutputFromFile('SubsequentExtendClosingUI')
    .setTitle('Subsequent Extend Closing');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getVessels() {
  const vesselSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vessel");
  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);
  return vesselData.map(row => row[0]);
}

function sendSubsequentEmail(vessel, sendingNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extendSheet = ss.getSheetByName("Extend Data");
  const vesselSheet = ss.getSheetByName("Vessel");

  const extendData = extendSheet.getDataRange().getDisplayValues();
  const newRequests = [];
  const rowsToUpdate = [];

  // Collect all unprocessed requests for the vessel
  extendData.slice(1).forEach((row, index) => {
    const [idRequest, , , , gateIn, rowVessel, booking, container, type, pod, weight, , sending] = row;
    if (rowVessel === vessel && !sending) {
      newRequests.push({ bookingNumber: booking, containerNo: container, type, pod, weight, gateIn });
      rowsToUpdate.push(index + 2);
    }
  });

  // All requests (no deduplication)
  const allRequests = extendData.slice(1)
    .filter(row => row[5] === vessel && row[12])
    .map(row => ({
      bookingNumber: row[6],
      containerNo: row[7],
      type: row[8],
      pod: row[9],
      weight: row[10],
      gateIn: row[4]
    }))
    .concat(newRequests);

  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);
  const terminal = vesselData.find(row => row[0] === vessel)[3];

  if (newRequests.length > 0) {
    sendExtendClosingEmail(vessel, terminal, newRequests, sendingNumber, allRequests);

    rowsToUpdate.forEach(rowNum => {
      extendSheet.getRange(rowNum, 13).setValue(`${sendingNumber}${getOrdinalSuffix(sendingNumber)} Sending`);
    });
    SpreadsheetApp.flush();
  }

  return newRequests.length;
}

function sendExtendClosingEmail(vessel, terminal, newRequests, sendingNumber, allRequests = newRequests) {
  const recipient = "dimasalif5@gmail.com";
  const ordinalSuffix = getOrdinalSuffix(sendingNumber);
  const subject = `(${sendingNumber}${ordinalSuffix.toUpperCase()} SENDING) REQUEST EXTEND CLOSING - ${vessel}`;
  
  const tempSpreadsheet = SpreadsheetApp.create(`${vessel}_${sendingNumber}${ordinalSuffix}_Sending_${Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd_HHmmss")}`);
  const sheet = tempSpreadsheet.getActiveSheet();
  
  const headers = ["Vessel Name", "Booking Number", "Container No", "Type", "POD", "Weight (Kg)", "Gate in Request"];
  sheet.appendRow(headers);
  
  allRequests.forEach(req => {
    sheet.appendRow([vessel, req.bookingNumber || "", req.containerNo || "", req.type || "", req.pod || "", req.weight || "", req.gateIn || ""]);
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
  const excelBlob = response.getBlob().setName(`${vessel}_${sendingNumber}${ordinalSuffix}_Sending.xlsx`);
  
  let htmlBody = "";
  if (terminal === "JICT") {
    htmlBody = `<p>Dear JICT Team,<br>Planning, Billing, & Gate Team,</p><p>Please kindly assist to accept below ONE late coming units to load onto the subject vessel.</p>`;
  } else if (terminal === "KOJA") {
    htmlBody = `<p><b>Dear Koja Planning,</b><br><b>Billing Team,</b><br><b>SSL Team,</b></p><p>Please kindly assist to accept extended closing requests on the subject vessel.</p>`;
  } else if (terminal === "MAL") {
    htmlBody = `<p><b>Dear MAL Team,</b><br><b>SPV Planning & Operations,</b></p><p>Please assist to accept below ONE late coming units to load on the subject vessel.</p>`;
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
      ${newRequests.map(req => `
        <tr>
          <td>${vessel}</td>
          <td>${req.bookingNumber || ""}</td>
          <td>${req.containerNo || ""}</td>
          <td>${req.type || ""}</td>
          <td>${req.pod || ""}</td>
          <td>${req.weight || ""}</td>
          <td>${req.gateIn || ""}</td>
        </tr>`).join('')}
    </table>
    <p>For your convenience, the same data is also attached as an Excel file.</p>
    <p><span style="background-color: yellow;"><b>All costs incurred will be under shipper's responsibility.</b></span></p>
    <p>Appreciate your approval for the above extend closing request.<br>Thank you</p>
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vesselSheet = ss.getSheetByName("Vessel");
  const vesselData = vesselSheet.getDataRange().getDisplayValues().slice(1);
  
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  vesselData.forEach(row => {
    const closingTime = new Date(`${row[2]} GMT+0700`);
    ScriptApp.newTrigger("processExtendRequests")
      .timeBased()
      .at(closingTime)
      .create();
    Logger.log("Scheduled 1st sending for %s at %s", row[0], closingTime);
  });
}