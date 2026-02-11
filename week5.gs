function appendWeek5RecentCWs() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("CAP List");
  const targetSheet = ss.getSheetByName("AvidSR_All");

  if (!sourceSheet || !targetSheet) {
    console.log("❌ Missing CAP List or AvidSR_All sheet.");
    return;
  }

  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData[0];
  const rows = sourceData.slice(1);

  const emailIndex = headers.indexOf("E-mail Address");
  const reasonIndex = headers.indexOf("Reason for CAP");
  const wsIndex = headers.indexOf("Workstream");
  const finalStageIndex = headers.indexOf("Final Stage Sent");
  const recIndex = headers.indexOf("Recomendation");
  const syncColIndex = headers.indexOf("Sync Status");
  const notifColIndex = headers.indexOf("Notification Email Sent");

  const today = new Date();
  const recentLimitDays = 7; // last 7 days

  // Column order: Reason for CAP right after Email
  const neededCols = [
    "E-mail Address",
    "Reason for CAP",
    "Throughput Week 1", "Quality Week 1",
    "Throughput Week 2", "Quality Week 2",
    "Throughput Week 3", "Quality Week 3",
    "Throughput Week 4", "Quality Week 4",
    "Throughput Week 5", "Quality Week 5",
    "Recomendation", "Sync Status"
  ];
  const colIndexes = neededCols.map(c => headers.indexOf(c));

  // Filter Week 5 CWs: must have Recommendation + completed recently
  const week5Updates = rows.filter(r => {
    if ((r[wsIndex] || "").toString().trim().toLowerCase() !== "avidxchange-strongroom") return false;

    // Must have Recommendation
    if (!r[recIndex] || r[recIndex].toString().trim() === "") return false;

    // Must have recent Final Stage Sent
    const finalStageRaw = r[finalStageIndex];
    if (!finalStageRaw) return false;
    const parsed = new Date(finalStageRaw);
    if (isNaN(parsed)) return false;

    const diffDays = (today - parsed) / (1000 * 60 * 60 * 24);
    return diffDays <= recentLimitDays;
  });

  // --- Find end of Week 4 section ---
  const targetData = targetSheet.getDataRange().getValues();
  let week4Row = targetData.findIndex(r => r[0].toString().includes("🟢 Week 4 CWs"));
  if (week4Row === -1) week4Row = targetData.length - 1;
  else {
    week4Row++;
    while (week4Row < targetData.length &&
           targetData[week4Row][0] &&
           !targetData[week4Row][0].toString().includes("🟢 Week") &&
           !targetData[week4Row][0].toString().includes("No data")) {
      week4Row++;
    }
  }

  // Remove "No data so far..." placeholder if present
  if (targetData[week4Row] && targetData[week4Row][0].toString().includes("No data so far")) {
    targetSheet.deleteRow(week4Row + 1);
  }

  // Insert 2 blank rows before Week 5
  targetSheet.insertRowsAfter(week4Row, 2);
  week4Row += 2;

  // Insert Week 5 header
  const week5Count = week5Updates.length;
  targetSheet.insertRowsAfter(week4Row, 1);
  const headerRange = targetSheet.getRange(week4Row + 1, 1);
  headerRange.setValue(`🟢 Week 5 CWs (✅ These CWs completed CAP recently — within the last ${recentLimitDays} days) (${week5Count} CWs)`)
             .setFontWeight("bold")
             .setBackground("#c9daf8"); // distinct light blue

  // Insert a blank row before data
  targetSheet.insertRowsAfter(week4Row + 1, 1);

  // Handle "No data" case
  if (week5Count === 0) {
    targetSheet.getRange(week4Row + 2, 1).setValue("No data so far...")
      .setFontStyle("italic")
      .setFontColor("#999999");
    SpreadsheetApp.getActive().toast(`⚠️ No Week 5 CWs found within last ${recentLimitDays} days.`, "Week 5", 5);
    return;
  }

  // Append Week 5 data
  week5Updates.forEach((row, i) => {
    const newRow = colIndexes.map(idx => row[idx] || "");

    // Last Communication Date
    let lastCommDate = "";
    if (row[syncColIndex]) {
      lastCommDate = extractLastDate(row[syncColIndex]);
    } else if (row[notifColIndex]) {
      lastCommDate = extractLastDate(row[notifColIndex]);
    }
    newRow.push(lastCommDate);

    // Latest Status
    newRow.push("Week 5");

    // Average Throughput
    const tpCols = ["Throughput Week 1","Throughput Week 2","Throughput Week 3","Throughput Week 4","Throughput Week 5"];
    const avgTp = tpCols.map(h => {
      const idx = headers.indexOf(h);
      return idx !== -1 ? parseFloat(row[idx]) || 0 : 0;
    }).reduce((a,b)=>a+b,0)/tpCols.length;
    newRow.push(avgTp.toFixed(2));

    // Average Quality
    const qCols = ["Quality Week 1","Quality Week 2","Quality Week 3","Quality Week 4","Quality Week 5"];
    const avgQ = qCols.map(h => {
      const idx = headers.indexOf(h);
      return idx !== -1 ? parseFloat(row[idx]) || 0 : 0;
    }).reduce((a,b)=>a+b,0)/qCols.length;
    newRow.push(avgQ.toFixed(2));

    // Insert row
    targetSheet.insertRowsAfter(week4Row + 2 + i, 1);
    const rowRange = targetSheet.getRange(week4Row + 3 + i, 1, 1, newRow.length);
    rowRange.setValues([newRow]);
    rowRange.setBackground("#f2f2f2"); // light green for Week 5 rows
  });

  SpreadsheetApp.getActive().toast(`✅ Week 5 appended with ${week5Count} CWs.`, "Week 5", 5);
}

/* ------------------- Helper ------------------- */
function extractLastDate(text) {
  if (!text) return "";
  const dateMatches = text.match(/\d{1,2}\/\d{1,2}\/\d{4}/g);
  if (!dateMatches) return "";
  return dateMatches[dateMatches.length - 1]; // Take the last date
}
