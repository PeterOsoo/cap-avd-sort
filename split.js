function splitAvidStrongroomCAP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("CAP List");
  const allData = sourceSheet.getDataRange().getValues();
  const headers = allData[0];
  const data = allData.slice(1);

  const neededCols = [
    "E-mail Address", "Throughput Week 1", "Quality Week 1", "TAT Week 1",
    "Throughput Week 2", "Quality Week 2", "TAT Week 2",
    "Throughput Week 3", "Quality Week 3", "TAT Week 3",
    "Throughput Week 4", "Quality Week 4", "TAT Week 4",
    "Throughput Week 5", "Quality Week 5", "TAT Week 5",
    "Recomendation", "Sync Status"
  ];

  const colIndexes = neededCols.map(c => headers.indexOf(c)).filter(i => i !== -1);
  const wsIndex = headers.indexOf("Workstream");

  const throughputCols = getExistingIndexes(headers, [
    "Throughput Week 1", "Throughput Week 2", "Throughput Week 3",
    "Throughput Week 4", "Throughput Week 5"
  ]);
  const qualityCols = getExistingIndexes(headers, [
    "Quality Week 1", "Quality Week 2", "Quality Week 3",
    "Quality Week 4", "Quality Week 5"
  ]);

  const filteredRows = [];
  const completedRows = [];

  data.forEach(row => {
    const ws = row[wsIndex];
    if (ws !== "AvidXchange Strongroom") return;

    const sync = (row[headers.indexOf("Sync Status")] || "").toString().trim();
    const picked = colIndexes.map(i => row[i]);
    const avgThroughput = average(throughputCols.map(i => parseFloat(row[i])));
    const avgQuality = average(qualityCols.map(i => parseFloat(row[i])));

    if (sync.toLowerCase().includes("cap final") || sync.toLowerCase().includes("cw resigned")) {
      completedRows.push([...picked, "", avgThroughput, avgQuality]);
    } else {
      const latestStatus = extractLatestStatus(sync);
      filteredRows.push([...picked, latestStatus, avgThroughput, avgQuality]);
    }
  });

  const finalHeaders = [...neededCols, "Latest Status", "Average Throughput", "Average Quality"];
  const allSheet = ss.getSheetByName("AvidSR_All") || ss.insertSheet("AvidSR_All");
  allSheet.clearContents();
  allSheet.clearFormats();

  // --- Header setup ---
  allSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  const headerRange = allSheet.getRange(1, 1, 1, finalHeaders.length);
  headerRange
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  allSheet.setFrozenRows(1);

  // --- Write Active CAPs ---
  let rowPointer = 2;
  if (filteredRows.length) {
    allSheet.getRange(rowPointer, 1, filteredRows.length, finalHeaders.length).setValues(filteredRows);
    rowPointer += filteredRows.length;
  }

  // --- 8 blank rows separator ---
  rowPointer += 8;

  // --- Label only in column A (no merge) ---
  const labelRow = rowPointer;
  const labelCell = allSheet.getRange(labelRow, 1);
  labelCell
    .setValue("These are CWS out of CAP / Completed")
    .setFontWeight("bold")
    .setBackground("#f2f2f2")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left");

  rowPointer++;

  // --- Write Completed CAPs ---
  if (completedRows.length) {
    allSheet.getRange(rowPointer, 1, completedRows.length, finalHeaders.length).setValues(completedRows);
  }

  // --- Styling and cleanup ---
  allSheet.autoResizeColumns(1, finalHeaders.length);
  const dataRange = allSheet.getDataRange();
  dataRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

  // --- Send summary email ---
  sendCompletionEmail(filteredRows.length, completedRows.length);
}

/** Helpers **/
function getExistingIndexes(headers, names) {
  return names.map(n => headers.indexOf(n)).filter(i => i !== -1);
}

function average(arr) {
  const valid = arr.filter(v => !isNaN(v));
  return valid.length ? (valid.reduce((a, b) => a + b, 0) / valid.length).toFixed(2) : "";
}

function extractLatestStatus(sync) {
  const matches = sync.match(/Week\s*\d+/g);
  if (!matches) return "";
  const latest = Math.max(...matches.map(m => parseInt(m.replace(/\D/g, ""))));
  return `Week ${latest}`;
}


