function generateDetailedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockInSheet = ss.getSheetByName("StockIn&Return");
  const stockOutSheet = ss.getSheetByName("Stock Out");

  const stockInData = stockInSheet.getRange(4, 1, stockInSheet.getLastRow() - 3, 10).getValues();
  const stockOutData = stockOutSheet.getRange(4, 1, stockOutSheet.getLastRow() - 3, 19).getValues();

  const PREFIX = "REP - ";
  const resultLogs = {};
  const resultSummary = {};

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith(PREFIX)) ss.deleteSheet(sheet);
  });

  function processRow(sku, item, person, allocation, action) {
    if (!item || !sku) return;
    const holder = allocation === "No Allocation Bag" ? person : allocation;
    if (!holder) return;

    if (!resultLogs[holder]) resultLogs[holder] = [];
    if (!resultSummary[holder]) resultSummary[holder] = {};

    resultLogs[holder].push([sku, item, action]);
    resultSummary[holder][item] = (resultSummary[holder][item] || 0) + (action === "Stock In" ? 1 : -1);
  }

  stockInData.forEach(row => processRow(row[6], row[9], row[2], row[4], row[3]));
  stockOutData.forEach(row => processRow(row[5], row[18], row[2], row[3], "Stock Out"));

  Object.keys(resultLogs).forEach(holder => {
    const sheet = ss.insertSheet(PREFIX + holder.toString().substring(0, 90));

    const logOutput = [["SKU", "Item Name", "Action"], ...resultLogs[holder]];
    sheet.getRange(1, 1, logOutput.length, 3).setValues(logOutput);

    const summaryOutput = [["Item", "Total"], ...Object.entries(resultSummary[holder])];
    sheet.getRange(1, 5, summaryOutput.length, 2).setValues(summaryOutput);
  });
}

function generateCompanySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const PREFIX = "COM - ";

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith(PREFIX)) ss.deleteSheet(sheet);
  });

  const restockSheet = ss.getSheetByName("Restock");
  const stockInSheet = ss.getSheetByName("StockIn&Return");
  const outrightSheet = ss.getSheetByName("Outright Purchase");
  const stockOutSheet = ss.getSheetByName("Stock Out");

  const restockLastRow = restockSheet.getLastRow();
  const stockInLastRow = stockInSheet.getLastRow();
  const outrightLastRow = outrightSheet.getLastRow();
  const stockOutLastRow = stockOutSheet.getLastRow();

  const restockData = restockLastRow >= 4
    ? restockSheet.getRange(4, 1, restockLastRow - 3, 12).getValues()
    : [];
  const stockInData = stockInLastRow >= 4
    ? stockInSheet.getRange(4, 1, stockInLastRow - 3, 11).getValues()
    : [];
  const outrightData = outrightLastRow >= 4
    ? outrightSheet.getRange(4, 1, outrightLastRow - 3, 16).getValues()
    : [];
  const stockOutData = stockOutLastRow >= 4
    ? stockOutSheet.getRange(4, 1, stockOutLastRow - 3, 20).getValues()
    : [];

  const companySummary = {};

  function addToSummary(company, item, delta) {
    if (!company || !item) return;
    if (!companySummary[company]) companySummary[company] = {};
    companySummary[company][item] = (companySummary[company][item] || 0) + delta;
  }

  restockData.forEach(row => {
    const company = row[11];
    const item = row[2];
    if (!company || !item) return;
    addToSummary(company, item, 1);
  });

  stockInData.forEach(row => {
    const item = row[9];
    const company = row[10];
    const reason = row[3];
    if (!company || !item || !reason) return;
    if (reason === "Stock In") addToSummary(company, item, -1);
    else if (reason === "Returned") addToSummary(company, item, 1);
  });

  outrightData.forEach(row => {
    const reason = row[4];
    const item = row[14];
    const company = row[15];
    if (!company || !item || !reason) return;
    if (reason === "Sold") addToSummary(company, item, -1);
    else if (reason === "Returned") addToSummary(company, item, 1);
  });

  stockOutData.forEach(row => {
    const company = row[19];
    const item = row[18];
    if (!company || !item) return;
    addToSummary(company, item, -1);
  });

  Object.keys(companySummary).forEach(company => {
    const sheet = ss.insertSheet(PREFIX + company.toString().substring(0, 90));

    const summaryOutput = [["Item", "Total"], ...Object.entries(companySummary[company])];
    sheet.getRange(1, 1, summaryOutput.length, 2).setValues(summaryOutput);
  });
}

/*
function generateDetailedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const stockInSheet = ss.getSheetByName("StockIn&Return");
  const stockOutSheet = ss.getSheetByName("Stock Out");

  const stockInData = stockInSheet.getRange(4, 1, stockInSheet.getLastRow() - 3, 10).getValues();
  const stockOutData = stockOutSheet.getRange(4, 1, stockOutSheet.getLastRow() - 3, 19).getValues();

  const resultLogs = {};
  const resultSummary = {};

  const PREFIX = "REP - ";

  // -------- DELETE ONLY GENERATED SHEETS --------
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith(PREFIX)) {
      ss.deleteSheet(sheet);
    }
  });

  // -------- STOCK IN & RETURN --------
  stockInData.forEach(row => {
    const person = row[2];
    const action = row[3];
    const allocation = row[4];
    const item = row[9];
    const sku = row[6];

    if (!item || !sku) return;

    const holder = allocation === "No Allocation Bag" ? person : allocation;
    if (!holder) return;

    if (!resultLogs[holder]) resultLogs[holder] = [];
    if (!resultSummary[holder]) resultSummary[holder] = {};

    resultLogs[holder].push([sku, item, action]);

    if (!resultSummary[holder][item]) resultSummary[holder][item] = 0;

    if (action === "Stock In") resultSummary[holder][item] += 1;
    else if (action === "Returned") resultSummary[holder][item] -= 1;
  });

  // -------- STOCK OUT --------
  stockOutData.forEach(row => {
    const person = row[2];
    const allocation = row[3];
    const item = row[18];
    const sku = row[5];

    if (!item || !sku) return;

    const holder = allocation === "No Allocation Bag" ? person : allocation;
    if (!holder) return;

    if (!resultLogs[holder]) resultLogs[holder] = [];
    if (!resultSummary[holder]) resultSummary[holder] = {};

    resultLogs[holder].push([sku, item, "Stock Out"]);

    if (!resultSummary[holder][item]) resultSummary[holder][item] = 0;

    resultSummary[holder][item] -= 1;
  });

  // -------- CREATE SHEETS --------
  Object.keys(resultLogs).forEach(holder => {
    const safeName = PREFIX + holder.toString().substring(0, 90);

    const sheet = ss.insertSheet(safeName);

    // LEFT SIDE
    let logOutput = [["SKU", "Item Name", "Action"]];
    logOutput.push(...resultLogs[holder]);
    sheet.getRange(1, 1, logOutput.length, 3).setValues(logOutput);

    // RIGHT SIDE (SUMMARY)
    let summaryOutput = [["Item", "Total"]];
    Object.keys(resultSummary[holder]).forEach(item => {
      summaryOutput.push([item, resultSummary[holder][item]]);
    });

    sheet.getRange(1, 5, summaryOutput.length, 2).setValues(summaryOutput);
  });
}
*/
