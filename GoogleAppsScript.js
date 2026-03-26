// ============================================================
// Paste this ENTIRE script into your Google Sheet Apps Script
// (Extensions > Apps Script > paste > Save > Deploy)
// ============================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ensure all sheets exist
    ensureSheets(ss);

    // Append to Raw Data
    var rawSheet = ss.getSheetByName("Raw Data");
    rawSheet.appendRow([
      data.ServiceType,
      data.TicketReference,
      data.Subject,
      data.HandledBy,
      data.InitiatedBy,
      data.TicketStatus,
      data.Hours,
      data.Date
    ]);

    // Rebuild filtered data sheets
    rebuildDeploymentReport(ss);
    rebuildTicketReport(ss);
    rebuildContractDetails(ss);
    rebuildHoursPerMonth(ss);

    return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function ensureSheets(ss) {
  var names = ["Raw Data", "Deployment Hours Report", "Ticket Hours Report", "Contract Details", "Hours Per Month"];
  var existing = ss.getSheets().map(function(s) { return s.getName(); });

  names.forEach(function(name) {
    if (existing.indexOf(name) === -1) {
      ss.insertSheet(name);
    }
  });

  // Add Raw Data headers if empty
  var raw = ss.getSheetByName("Raw Data");
  if (raw.getLastRow() === 0) {
    raw.appendRow(["Type Service", "Ticket Reference", "Subject", "Handled By", "Initiated By", "Ticket Status", "Hours", "Date"]);
    styleHeader(raw);
  }
  raw.hideSheet();
}

function rebuildDeploymentReport(ss) {
  var raw = ss.getSheetByName("Raw Data");
  var ws = ss.getSheetByName("Deployment Hours Report");

  // Clear everything
  ws.clear();

  // Headers
  ws.getRange(1, 1, 1, 7).setValues([["Type Service", "Ticket Reference", "Subject", "Handled By", "Initiated By", "Hours", "Date"]]);
  styleHeader(ws);

  // Filter non-Support entries
  var lastRow = raw.getLastRow();
  if (lastRow < 2) return;

  var data = raw.getRange(2, 1, lastRow - 1, 8).getValues();
  var row = 2;
  data.forEach(function(r) {
    if (r[0] !== "Support" && r[0] !== "") {
      var type = r[0] === "Engineer" ? "Deployment" : r[0];
      ws.getRange(row, 1, 1, 7).setValues([[type, r[1], r[2], r[3], r[4], r[6], r[7]]]);
      row++;
    }
  });

  // Data validation dropdown
  if (row > 2) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Management", "Deployment"])
      .setAllowInvalid(false)
      .build();
    ws.getRange(2, 1, 998, 1).setDataValidation(rule);
  }

  ws.setFrozenRows(1);
  ws.autoResizeColumns(1, 7);
}

function rebuildTicketReport(ss) {
  var raw = ss.getSheetByName("Raw Data");
  var ws = ss.getSheetByName("Ticket Hours Report");

  ws.clear();

  // Headers (7 columns - Total Hours in G)
  ws.getRange(1, 1, 1, 7).setValues([["Ticket", "Ticket Status", "Initiated By", "Handled By", "Date", "Hours", "Total Hours"]]);
  styleHeader(ws);

  var lastRow = raw.getLastRow();
  if (lastRow < 2) {
    ws.getRange(2, 7).setFormula("=SUM(F2:F2)");
    return;
  }

  var data = raw.getRange(2, 1, lastRow - 1, 8).getValues();
  var row = 2;
  data.forEach(function(r) {
    if (r[0] === "Support") {
      ws.getRange(row, 1, 1, 6).setValues([[r[1], r[5], r[4], r[3], r[7], r[6]]]);
      row++;
    }
  });

  var lastDataRow = Math.max(row - 1, 2);
  ws.getRange(2, 7).setFormula("=SUM(F2:F" + lastDataRow + ")");
  ws.getRange(2, 7).setFontWeight("bold");

  ws.setFrozenRows(1);
  ws.autoResizeColumns(1, 7);
}

function rebuildContractDetails(ss) {
  var ws = ss.getSheetByName("Contract Details");

  ws.clear();

  var data = [
    ["Service", "Total Hours"],
    ["Deployment", '=SUMIF(\'Deployment Hours Report\'!A2:A1000,"Deployment",\'Deployment Hours Report\'!F2:F1000)'],
    ["Management", '=SUMIF(\'Deployment Hours Report\'!A2:A1000,"Management",\'Deployment Hours Report\'!F2:F1000)'],
    ["Support", "='Ticket Hours Report'!G2"],
    ["Total", "=SUM(B2:B4)"]
  ];

  ws.getRange(1, 1, data.length, 2).setValues(data);
  styleHeader(ws);

  // Bold total row
  ws.getRange(5, 1, 1, 2).setFontWeight("bold");

  // Number format
  ws.getRange(2, 2, 4, 1).setNumberFormat("0.00");

  // Protect sheet
  var protection = ws.protect().setDescription("Read-only summary");
  protection.setWarningOnly(true);

  ws.autoResizeColumns(1, 2);
}

function rebuildHoursPerMonth(ss) {
  var raw = ss.getSheetByName("Raw Data");
  var ws = ss.getSheetByName("Hours Per Month");

  ws.clear();

  ws.getRange(1, 1, 1, 4).setValues([["Month", "Management Hours", "Support Hours", "Total"]]);
  styleHeader(ws);

  var lastRow = raw.getLastRow();
  if (lastRow < 2) return;

  var data = raw.getRange(2, 1, lastRow - 1, 8).getValues();

  // Get distinct months
  var months = {};
  data.forEach(function(r) {
    if (r[0] === "" || !r[7]) return;
    var d = new Date(r[7]);
    var key = Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/yyyy");
    months[key] = { year: d.getFullYear(), month: d.getMonth() };
  });

  // Sort months
  var sortedKeys = Object.keys(months).sort(function(a, b) {
    var ma = months[a], mb = months[b];
    return (ma.year - mb.year) || (ma.month - mb.month);
  });

  var row = 2;
  sortedKeys.forEach(function(monthStr) {
    ws.getRange(row, 1).setValue(monthStr);

    // Management Hours from Deployment Hours Report
    ws.getRange(row, 2).setFormula(
      '=SUMPRODUCT((\'Deployment Hours Report\'!$A$2:$A$1000="Management")*(TEXT(\'Deployment Hours Report\'!$G$2:$G$1000,"MM/YYYY")=A' + row + ')*\'Deployment Hours Report\'!$F$2:$F$1000)'
    );

    // Support Hours from Ticket Hours Report + Deployment entries
    ws.getRange(row, 3).setFormula(
      '=SUMPRODUCT((TEXT(\'Ticket Hours Report\'!$E$2:$E$1000,"MM/YYYY")=A' + row + ')*\'Ticket Hours Report\'!$F$2:$F$1000)' +
      '+SUMPRODUCT((\'Deployment Hours Report\'!$A$2:$A$1000="Deployment")*(TEXT(\'Deployment Hours Report\'!$G$2:$G$1000,"MM/YYYY")=A' + row + ')*\'Deployment Hours Report\'!$F$2:$F$1000)'
    );

    // Total
    ws.getRange(row, 4).setFormula("=B" + row + "+C" + row);

    row++;
  });

  var lastDataRow = row - 1;

  // Total row
  ws.getRange(row, 1).setValue("Total").setFontWeight("bold");
  ws.getRange(row, 2).setFormula("=SUM(B2:B" + lastDataRow + ")").setFontWeight("bold");
  ws.getRange(row, 3).setFormula("=SUM(C2:C" + lastDataRow + ")").setFontWeight("bold");
  ws.getRange(row, 4).setFormula("=SUM(D2:D" + lastDataRow + ")").setFontWeight("bold");

  // Number format
  ws.getRange(2, 2, row - 1, 3).setNumberFormat("0.00");

  // Protect sheet
  var protection = ws.protect().setDescription("Read-only summary");
  protection.setWarningOnly(true);

  ws.setFrozenRows(1);
  ws.autoResizeColumns(1, 4);
}

function styleHeader(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var header = sheet.getRange(1, 1, 1, lastCol);
  header.setFontWeight("bold");
  header.setFontColor("#FFFFFF");
  header.setBackground("#4472C4");
  header.setHorizontalAlignment("center");
}
