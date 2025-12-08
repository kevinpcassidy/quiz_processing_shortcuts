function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Formula Tools")
    .addItem("Fill Average Formulas", "fillAverageFormulaForSelectedColumn")
    .addItem("Check all Scores for Updates", "updateScoresFromSourceSheet")
    .addToUi();
}


function fillAverageFormulaForSelectedColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveCell();
  const targetCol = range.getColumn();

  if (targetCol <= 4) {
    SpreadsheetApp.getUi().alert("Please select a column at least 5 or later (needs 4 preceding columns).");
    return;
  }

  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Column A names
  const formulaTemplate = '=IF(COUNTA(%range%)=0, "", IFERROR(AVERAGE(LARGE(FILTER(%range%,%range%<>""),1), LARGE(FILTER(%range%,%range%<>""),2)), MAX(%range%)))';

  // Determine 4-column range immediately to the left
  const startCol = targetCol - 4;

  for (let i = 0; i < names.length; i++) {
    const row = i + 2;
    if (names[i][0]) { // only fill if Column A has a name
      const rangeA1 = sheet.getRange(row, startCol, 1, 4).getA1Notation();
      const formula = formulaTemplate.replace(/%range%/g, rangeA1);
      sheet.getRange(row, targetCol).setFormula(formula);
    }
  }

  SpreadsheetApp.getUi().alert("Formulas filled successfully!");
}

function updateScoresFromSourceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // Get all sheet names
  const sheets = ss.getSheets().map(sheet => sheet.getName());

  // Create a numbered selection prompt
  const response = ui.prompt(
    'Select source sheet by number:\n' +
    sheets.map((name, i) => `${i + 1}: ${name}`).join('\n'),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const choice = parseInt(response.getResponseText());
  if (isNaN(choice) || choice < 1 || choice > sheets.length) {
    ui.alert("Invalid selection.");
    return;
  }

  const sourceSheetName = sheets[choice - 1];
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  // Get source headers and data
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  const sourceNames = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 1).getValues().map(r => r[0]);

  // Get target headers and names
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const targetNames = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 1).getValues().map(r => r[0]);

  // Build a map of source data by student name
  const sourceMap = {};
  for (let i = 0; i < sourceNames.length; i++) {
    sourceMap[sourceNames[i]] = sourceData[i];
  }

  // Loop through each target column
  for (let tCol = 1; tCol <= targetHeaders.length; tCol++) {
    const header = targetHeaders[tCol - 1];
    const sCol = sourceHeaders.indexOf(header);
    if (sCol === -1) continue; // skip if header not found in source

    // Loop through each student
    for (let row = 0; row < targetNames.length; row++) {
      const studentName = targetNames[row];
      if (sourceMap[studentName] && sourceMap[studentName][sCol] !== "") {
        targetSheet.getRange(row + 2, tCol).setValue(sourceMap[studentName][sCol]);
      }
    }
  }

  ui.alert("Scores updated from " + sourceSheetName + "!");
}
