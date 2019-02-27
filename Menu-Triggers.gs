function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Append GA Data')
      .addItem('Test Queries', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Scheduling')
      .addItem('Run Daily', 'menuItem2')
      .addItem('Run Weekly', 'menuItem3')
      .addItem('Run Monthly', 'menuItem4')
      .addItem('Stop Scheduling', 'menuItem5'))      
      .addToUi();
}

function menuItem1() {
  updateReport();
}

function menuItem2() {
  removeTriggers();
  createTrigger('Daily');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Report now scheduled to run daily.');
}

function menuItem3() {
  removeTriggers();
  createTrigger('Weekly');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Report now scheduled to run weekly.');
}

function menuItem4() {
  removeTriggers();
  createTrigger('Monthly');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Report now scheduled to run monthly.');
}

function menuItem5() {
  removeTriggers();
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('Report scheduling now stopped.');
}


function createTrigger(frequency) {
  switch(frequency) {      
    case 'Daily':
      // Trigger report update every day at 3am.
      ScriptApp.newTrigger('updateReport')
      .timeBased()
      .atHour(4)
      .everyDays(1)
      .create();
      break;
      
    case 'Weekly':  
      // Trigger report update at 4am on Monday of every week.
      ScriptApp.newTrigger('updateReport')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(4)
      .everyWeeks(1)
      .create();
      break;
      
    case 'Monthly':
      // Trigger report update at 4am on 1st day of every month.
      ScriptApp.newTrigger('updateReport')
      .timeBased()
      .onMonthDay(1)
      .atHour(4)
      .create();
      break;
  }
}


function removeTriggers() {
  // Deletes all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
