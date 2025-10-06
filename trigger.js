function createDailyCAPTrigger() {
  // Delete any existing triggers for the same function (avoid duplicates)
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "splitAvidStrongroomCAP") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // ✅ Create new time-based triggers for Tuesday and Wednesday at 8:00 AM
  ScriptApp.newTrigger("splitAvidStrongroomCAP")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(8)
    .create();

  ScriptApp.newTrigger("splitAvidStrongroomCAP")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(8)
    .create();

  Logger.log("✅ Triggers created: splitAvidStrongroomCAP will run every Tue & Wed at 8:00 AM");
}
