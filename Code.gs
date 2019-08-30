var defaultValues = [
  ["Flat", "Share"],
  ["1", 0.1375],
  ["2", 0.1135],
  ["3", 0.2085],
  ["4", 0.1884],
  ["A", 0.1018],
  ["B", 0.0143],
  ["12", 0.2360],
];

var defaultValuesFormats = [
  ["@", "@"],
  ["@", "#.##%"],
  ["@", "#.##%"],
  ["@", "#.##%"],
  ["@", "#.##%"],
  ["@", "#.##%"],
  ["@", "#.##%"],
  ["@", "#.##%"],
];

function main(event) {
  if (event.changeType == 'INSERT_GRID') {
    flow()
   }
 }

function flow() {
   var ui = SpreadsheetApp.getUi();
   var newProject = dialogBoxWithText(ui, 'Create a new project...', 'What is the name of this project?');
   if (newProject && newProject != 'Cancelled') {
      var sheet = createNewSheet(newProject, ui);
      if (sheet) {
        var companies = [];
        var works = [];
        var costs = [];
        var cancelButtonTriggered = false;
        var closeButtonTriggered = false;
        var workCloseButtonTriggered = false;
        var i = 0;
        while (cancelButtonTriggered == false && closeButtonTriggered == false && workCloseButtonTriggered == false) {
          var createOnCancel = true;
          if (i == 0) {
            var createOnCancel = false;
          }
          var newCompany = getQuoteBox(ui, newProject);
          if (newCompany && newCompany != 'Cancelled') {
            var cancelButtonTriggered = false;

            var work = [];
            var cost = [];
            var workCancelButtonTriggered = false;
            var j = 0;
            while (workCancelButtonTriggered == false && workCloseButtonTriggered == false) {
              var createCompanyOnCancel = true;
              if (j == 0) {
                var createCompanyOnCancel = false;
              }
              var newWork = getWorkBox(ui, newCompany, newProject);
              if (newWork && newWork != 'Cancelled') {
                var newWorkCost = getCostBox(ui, newWork, newProject);
                if (newWorkCost && newWorkCost != 'Cancelled') {
                  work.push(newWork);
                  cost.push(newWorkCost);
                } else if (newWorkCost == undefined) {
                  var workCloseButtonTriggered = true;
                }
              } else if (newWork == 'Cancelled') {
                var workCancelButtonTriggered = true;
              } else if (newWork == undefined) {
                var workCloseButtonTriggered = true;
              }
              var j = j + 1;
            }
            if (createCompanyOnCancel && workCancelButtonTriggered && !workCloseButtonTriggered) {
              companies.push(newCompany);
              works.push(work);
              costs.push(cost);
            }
          } else if (newCompany == 'Cancelled') {
            var cancelButtonTriggered = true;
          } else if (newCompany == undefined) {
            var closeButtonTriggered = true;
          }
          var i = i + 1;
        }
        if (!closeButtonTriggered && !workCloseButtonTriggered && companies.length > 0) {
          buildNewSheetDefaults(sheet);
          var columnIndex = 3;
          var startIndexes = [];
          var endIndexes = [];
          for (c=0; c < companies.length; c++) {
            for (w=0; w < works[c].length; w++) {
              buildNewSheetColumn(sheet, w, columnIndex, companies[c], works[c][w], costs[c][w])
              if (w == 0) {
                var startColumnIndex = columnIndex;
              }
              if (w == works[c].length - 1) {
                var endColumnIndex = columnIndex;
              }
              var columnIndex = columnIndex + 1;
            }
            startIndexes.push(startColumnIndex);
            endIndexes.push(endColumnIndex);
          }
          
          for (c=0; c < companies.length; c++) {
            buildNewSheetTotalColumn(sheet, columnIndex, companies[c], startIndexes[c], endIndexes[c]);
            var adjustedTotal = getAdjustedTotal(sheet, columnIndex);
            var total = 0;
            for (ct=0; ct < costs[c].length; ct++) {
              var total = total + (costs[c][ct] * 100) // manipulate due to rounding errors
            }
            var total = total / 100;
            adjustValueForFlatA(ui, sheet, companies[c], total, adjustedTotal, columnIndex);
            setTotalForColumn(sheet, ui, columnIndex, total, companies[c]);
            var columnIndex = columnIndex + 1;
          }
        }
       }
   } else {
     deleteActiveSheet()
   }
 }

function deleteActiveSheet() {
 var activeSpreadsheet = SpreadsheetApp.getActive();
 var activeSheet = activeSpreadsheet.getActiveSheet();
 activeSpreadsheet.deleteSheet(activeSheet);
}

function createNewSheet(projectName, ui) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet.getSheetByName(projectName) != null) {
    ui.alert('project with name "' + projectName + '" already exists. project has not been created.')
    deleteActiveSheet(activeSpreadsheet);
    return;
  }
  deleteActiveSheet(activeSpreadsheet);
  var newSheet = activeSpreadsheet.insertSheet();
  newSheet.setName(projectName);
  return newSheet;
}

function buildNewSheetDefaults(sheet) {
  var range = sheet.getRange("A2:B9");
  range.setValues(defaultValues);
  var range = sheet.getRange("A2:B9");
  range.setNumberFormats(defaultValuesFormats);
  var range = sheet.getRange("A2:B2");
  range.setFontWeight('bold');
}

function insertCostIntoCell(sheet, columnIndex, rowIndex, cost) {
  var propCost = sheet.getRange(rowIndex, 2).getValue() * cost;
  var range = sheet.getRange(rowIndex, columnIndex);
  range.setValue(propCost);
  range.setNumberFormat("£#,##0.00");
}

function getTotalSumForFlat(sheet, rowIndex, columnIndex, firstColumnIndex, lastColumnIndex) {
  var total = 0;
  for (col = firstColumnIndex; col <= lastColumnIndex; col++) {
    var total = total  + (sheet.getRange(rowIndex, col).getValue() * 100); // manipulate due to rounding errors
  }
  var range = sheet.getRange(rowIndex, columnIndex);
  range.setFormula('=ROUND(' + (total / 100) + ', 2)');
  range.setNumberFormat("£#,##0.00");
  range.setFontWeight('bold');
}

function setTotalForColumn(sheet, ui, columnIndex, total, company) {
  var range = sheet.getRange(10, columnIndex);
  range.setValue(total);
  range.setFontWeight('bold');
  range.setNumberFormat("£#,##0.00");
  ui.alert('Total amount quoted for ' + company + ': ' + total);
}


function buildNewSheetColumn(sheet, workIndex, columnIndex, company, work, cost) {
  if (workIndex == 0) {
    var range = sheet.getRange(1, columnIndex);
    range.setValue(company);
    range.setFontWeight('bold');
  }
  var range = sheet.getRange(2, columnIndex);
  range.setValue(work);
  range.setFontWeight('bold');
  var startRowIndex = 3;
  for (f = 0; f < 7; f++) {
    insertCostIntoCell(sheet, columnIndex, startRowIndex + f, cost);
  }
}

function getAdjustedTotal(sheet, columnIndex) {
  var startRowIndex = 3;
  var adjustedTotal = 0;
  for (f = 0; f < 7; f++) {
    var adjustedTotal = adjustedTotal + sheet.getRange(startRowIndex + f, columnIndex).getValue();
  }
  return adjustedTotal;
}

function buildNewSheetTotalColumn(sheet, columnIndex, company, firstColumnIndex, lastColumnIndex) {
  var range = sheet.getRange(1, columnIndex);
  range.setValue(company);
  range.setFontWeight('bold');
  var range = sheet.getRange(2, columnIndex);
  range.setValue('Total');
  range.setFontWeight('bold');
  var startRowIndex = 3;
  for (f = 0; f < 7; f++) {
    getTotalSumForFlat(sheet, startRowIndex + f, columnIndex, firstColumnIndex, lastColumnIndex);
  }
}

function adjustValueForFlatA(ui, sheet, company, cost, roundedCost, totalCostColumnIndex) {
  var range = sheet.getRange(7, totalCostColumnIndex);
  var flatAValue = range.getValue();
  var newFlatAValue = flatAValue + (cost - roundedCost);
  range.setValue(newFlatAValue);
  if (flatAValue < newFlatAValue) {
    ui.alert('Flat A total amount for ' + company + ' quote adjusted upwards to ' + Math.round(newFlatAValue * 100)/ 100);
  }
  if (flatAValue > newFlatAValue) {
    ui.alert('Flat A total amount for ' + company + ' quote adjusted downwards to ' + Math.round(newFlatAValue * 100) / 100);
  }
}

function dialogBoxWithText(ui, title, message, sheetName) {
  var [projectTextButton, projectTextText] = getPromptButtonWithMessage(getTextOKResponse(ui, title, message));
  if (projectTextButton == ui.Button.OK) {
    if (projectTextText === '') {
      ui.alert('please enter a value.');
      return dialogBoxWithText(ui, title, message, sheetName);
    } else {
      return projectTextText;
    }
  } else if (projectTextButton == ui.Button.CANCEL) {
    return 'Cancelled';
  } else {
    return cancelledProjectSteps(ui, sheetName);
  }
}

function dialogBoxWithNumeric(ui, workType, sheetName) {
  var [projectNumericButton, projectNumericText] = getPromptButtonWithMessage(getNumericResponse(ui, workType));
  if (projectNumericButton == ui.Button.OK) {
    var projectNumeric = parseFloat(projectNumericText);
    if (isNaN(projectNumeric)) {
      ui.alert('please enter an integer or decimal number');
      return dialogBoxWithNumeric(ui, workType, sheetName);
    } else {
      return projectNumeric;
    }
  } else if (projectNumericButton == ui.Button.CANCEL) {
    return 'Cancelled';
  } else {
    return cancelledProjectSteps(ui, sheetName);
  }
}

function getQuoteBox(ui, sheetName) {
  return dialogBoxWithText(ui, 'Add a quote...', 'What is the name of the company?', sheetName);
}

function getWorkBox(ui, company, sheetName) {
  return dialogBoxWithText(ui, 'Add work for ' + company + '...', 'What does the work entail?', sheetName);
}

function getCostBox(ui, work, sheetName) {
  return dialogBoxWithNumeric(ui, work, sheetName);
}

function getTextOKResponse(ui, title, message) {
  return ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
}

function getNumericResponse(ui, workType) {
  return ui.prompt('How much will ' + workType + ' cost?', ui.ButtonSet.OK_CANCEL);
}

function getPromptButtonWithMessage(prompt) {
  return [prompt.getSelectedButton(), prompt.getResponseText()];
}

function cancelledProjectSteps(ui, sheetName) {
  ui.alert('project has not been created.')
  if (sheetName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(sheetName);
    ss.deleteSheet(sheet);
  }
}
