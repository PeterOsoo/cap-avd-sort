function filterStrongroom() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("CAP List");

  // --- Setup destination sheets ---
  const filteredName = "AvidSR_Filtered";
  const completedName = "AvidSR_Completed";

  // Always rebuild fresh
  [filteredName, completedName].forEach(name => {
    if (ss.getSheetByName(name)) ss.deleteSheet(ss.getSheetByName(name));
    ss.insertSheet(name);
  });

  const filteredSheet = ss.getSheetByName(filteredName);
  const completedSheet = ss.getSheetByName(completedName);

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift();

  // --- Split rows into active vs completed ---
  const filtered = [];
  const completed = [];

  data.forEach(row => {
    if (row[3] === "AvidXchange Strongroom") { // col D
      if (row[32]) {
        // Column AG (index 32, 0-based) has a value → Completed
        completed.push(row);
      } else {
        // No value in AG → Still Active
        filtered.push(row);
      }
    }
  });

  // --- Columns to remove ---
  const removeCols = [
    1, 3, 4, 5, 6, 7, 8, 9, 10,
    26, 27, 28, 29, 30, 31, 32, 33, 34
  ];

  const newHeaders = headers.filter((_, idx) => !removeCols.includes(idx + 1));

  const mapRow = row => row.filter((_, idx) => !removeCols.includes(idx + 1));

  const filteredData = filtered.map(mapRow);
  const completedData = completed.map(mapRow);

  // --- Write to Active sheet ---
  filteredSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  if (filteredData.length > 0) {
    filteredSheet.getRange(2, 1, filteredData.length, newHeaders.length).setValues(filteredData);
  }

  // --- Write to Completed sheet ---
  completedSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  if (completedData.length > 0) {
    completedSheet.getRange(2, 1, completedData.length, newHeaders.length).setValues(completedData);
  }

  ss.toast("✅ Strongroom filter completed.\n" +
           filteredData.length + " active rows → " + filteredName + "\n" +
           completedData.length + " completed rows → " + completedName,
           "Done", 7);
}



function createWednesdayTrigger() {
  // First, delete any old triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "filterStrongroom") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create a new time-based trigger for every Wednesday at 9am
  ScriptApp.newTrigger("filterStrongroom")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(9) // runs at 9 AM, adjust as you like (0–23)
    .create();
}
