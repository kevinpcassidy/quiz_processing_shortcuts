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
  const formulaTemplate = '=IF(COUNTA(%range%)=0,"",IFERROR(AVERAGE(LARGE(%range%,{1}),LARGE(%range%,{2})),MAX(%range%)))';

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

  // ... [your existing selection code] ...

  const sourceSheetName = sheets[choice - 1];
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  // Get source headers and data
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
  const sourceNames = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 1).getValues().map(r => r[0]);

  // Get target headers and data (read entire sheet at once)
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  const targetData = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).getValues();
  const targetNames = targetData.map(r => r[0]);

  // Build a map of source data by student name
  const sourceMap = {};
  for (let i = 0; i < sourceNames.length; i++) {
    sourceMap[sourceNames[i]] = sourceData[i];
  }

  // Update the target data array (in memory)
  for (let tCol = 0; tCol < targetHeaders.length; tCol++) {
    const header = targetHeaders[tCol];
    const sCol = sourceHeaders.indexOf(header);
    if (sCol === -1) continue;

    for (let row = 0; row < targetNames.length; row++) {
      const studentName = targetNames[row];
      if (sourceMap[studentName] && sourceMap[studentName][sCol] !== "") {
        targetData[row][tCol] = sourceMap[studentName][sCol];
      }
    }
  }

  // Write everything back in ONE operation
  targetSheet.getRange(2, 1, targetData.length, targetData[0].length).setValues(targetData);

  ui.alert("Scores updated from " + sourceSheetName + "!");
}
