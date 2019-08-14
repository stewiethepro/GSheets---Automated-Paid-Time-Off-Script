// PTO Code

// global scope

// get spreadsheet
var ss = SpreadsheetApp.getActiveSpreadsheet();

// define store for global variables
var scriptProperties = PropertiesService.getScriptProperties();

// function to define PTO increment and day of the month
function initialiseScript() {

  // create log sheet if none already
  if (ss.getSheetByName('Log') == null) {

    var newSheet = ss.insertSheet();
    newSheet.setName("Log");

    var sheet = ss.getSheetByName("Log");

    var columnHeaders = [
      [ "Timestamp", "Log", "Old Value", "New Value" ]
    ];

    var range = sheet.getRange("A1:D1");
    range.setValues(columnHeaders);

    sheet.hideSheet();
  };
  // set script properties
  scriptProperties.setProperties({
    'increment': 1,
    'day': 1,
    'cell': 'A2'
  });

  // display message
  getGlobalVariables();

  // create new trigger
  newTrigger();
};


// utility functions

// resize columns on log sheet
function resizeLogColumns(){
  var sheet = ss.getSheetByName("Log");
  sheet.autoResizeColumns(1,4);
}


// primary functions

// increment function - add increment to PTO
function incrementPTO() {
  var cell = scriptProperties.getProperty('cell');
  var sheet = ss.getSheets()[0];

   var range = sheet.getRange(cell);
   var value = range.getValue();

   var increment = scriptProperties.getProperty('increment');

   var valueNum = parseFloat(value);
   var incrementNum = parseFloat(increment);
   var finalNum = valueNum + incrementNum;

   range.setValue(finalNum);

   var newValue = range.getValue();

   Logger.log('PTO incremented: ' + newValue + ' - ' + new Date());

   var logs = ss.getSheets()[1];
   logs.appendRow([new Date() , 'PTO incremented', value, newValue]);

   resizeLogColumns();
}

// timer function - run incrementPTO
function newTrigger() {
  if (scriptProperties != null) {
    if (ScriptApp.getProjectTriggers().length) {
      SpreadsheetApp.getUi()
      .alert('Trigger is already active');
    }
    else {
      var day = Math.round(scriptProperties.getProperty('day'));
      var increment = scriptProperties.getProperty('increment');
      var cell = scriptProperties.getProperty('cell');

      ScriptApp.newTrigger("incrementPTO")
      .timeBased()
      .onMonthDay(day)
      .create();

      Logger.log('Trigger set by ' + Session.getActiveUser().getEmail() + ': ' + new Date());

      var logs = ss.getSheets()[1];
      logs.appendRow([new Date() , 'Trigger set by ' + Session.getActiveUser().getEmail() + ' with increment: ' + increment + ' and day: ' + day + ' referencing cell: ' + cell]);

      resizeLogColumns();

      SpreadsheetApp.getUi()
      .alert('Trigger added with increment: ' + increment + ' and day: ' + day + ', referencing cell: ' + cell);
    };
  }
  else {
  SpreadsheetApp.getUi()
      .alert('Please initialise script - click the button in the PTO menu');
  };
};

// delete all triggers in the current project.
function deleteTrigger(){
  if (ScriptApp.getProjectTriggers().length) {
    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);

      Logger.log('Trigger deleted: ' + new Date());

      var logs = ss.getSheets()[1];
      logs.appendRow([new Date() , 'Trigger deleted by ' + Session.getActiveUser().getEmail()]);

      resizeLogColumns();

      SpreadsheetApp.getUi()
      .alert('Trigger Deleted');
      return;
  }
}
  else {
    SpreadsheetApp.getUi()
    .alert('There are no active triggers');
  };
};

// set increment function

function setIncrement() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      'Set increment',
      'Enter a number',
       ui.ButtonSet.OK_CANCEL);

  // Process the user's response
  var button = result.getSelectedButton();
  var inputText = result.getResponseText();
  var inputNum = parseFloat(inputText);

  if (button == ui.Button.OK) {
    // User clicked "OK"
    if (inputNum >= 0 && inputNum <= 5) {

      scriptProperties.setProperty('increment', inputNum);

      Logger.log('Increment set to ' + inputNum + '.')

      var logs = ss.getSheets()[1];
      logs.appendRow([new Date() , 'Increment set to ' + inputNum + ' by ' + Session.getActiveUser().getEmail()]);

      resizeLogColumns();

      ui.alert('Increment set to ' + inputNum + '.');
    }

    else {
      ui.alert('Please enter a number between 0 and 5.');
      setIncrement();
    }
  }

  else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel"
    return;
  }

  else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar
    return;
  };
};

// get increment function

function getIncrement() {
  var ui = SpreadsheetApp.getUi();
  var increment = scriptProperties.getProperty('increment');
  ui.alert('Increment is currently set to: ' + increment);
};

// set day function

function setDay() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      'Set day',
      'Enter an integer between 1 and 28 - This will start a new trigger on the chosen day',
       ui.ButtonSet.OK_CANCEL);

  // Process the user's response
  var button = result.getSelectedButton();
  var inputText = result.getResponseText();
  var inputInt = parseInt(inputText, 10);

  if (button == ui.Button.OK) {
    // User clicked "OK"
    if (inputInt >= 1 && inputInt <= 28) {

      scriptProperties.setProperty('day', inputInt);

      Logger.log('Day set to ' + inputInt + '.')

      // delete the trigger
      deleteTrigger();

      // create a new trigger with new day
      newTrigger();
    }

    else {
      ui.alert('Please enter an integer between 1 and 28.');
      setDay();
    }
  }

  else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel"
    return;
  }

  else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar
    return;
  };
};

// get global variables function

function getGlobalVariables() {
  var ui = SpreadsheetApp.getUi();

  var increment = scriptProperties.getProperty('increment');
  var day = Math.round(scriptProperties.getProperty('day'));
  var cell = scriptProperties.getProperty('cell');

  ui.alert('Increment is currently set to: ' + increment + ', Day is currently set to: ' + day + ' and Cell is currently set to: ' + cell);
};

// set cell function

function setCell() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
      'Set cell',
      'Enter the cell reference of the number to be incremented - a letter followed by one to two numbers in the format XX or XXX',
       ui.ButtonSet.OK_CANCEL);

  // Process the user's response
  var button = result.getSelectedButton();
  var inputText = result.getResponseText();

  if (button == ui.Button.OK) {
    // User clicked "OK"
    if (inputText.length == 2 || inputText.length == 3) {

     // check length of input and apply correct validation test
     if (inputText.length == 2) {
      var regTest = new RegExp("(^[a-zA-Z])([1-9])")
     }
     else {
      var regTest = new RegExp("(^[a-zA-Z])([1-9])([1-9])")
     }

     // run validation test
     if (regTest.test(inputText)) {

      var cell = inputText.toUpperCase();
      scriptProperties.setProperty('cell', cell);

      Logger.log('Cell set to ' + cell + '.')

      var logs = ss.getSheets()[1];
      logs.appendRow([new Date() , 'Cell set to ' + cell + ' by ' + Session.getActiveUser().getEmail()]);

      resizeLogColumns();

      ui.alert('Cell set to ' + cell + '.');

    }
    else
     ui.alert('Please enter a letter followed by one to two numbers in the format XX or XXX');
   }

   else {
     ui.alert('Please enter a letter followed by one to two numbers in the format XX or XXX');

     setCell();
    }
  }

  else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel"
    return;
  }

  else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar
    return;
  };
};


// UI functions

// initialise UI

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PTO Menu')
      .addItem('Initialise script', 'initialiseScript')
      .addSeparator()
      .addSubMenu(ui.createMenu('Set values')
      .addItem('Set cell', 'setCell')
      .addItem('Set increment', 'setIncrement')
      .addItem('Set day', 'setDay'))
      .addSeparator()
      .addItem('Get current values', 'getGlobalVariables')
      .addSeparator()
      .addItem('Delete trigger', 'deleteTrigger')
      .addToUi();
}
