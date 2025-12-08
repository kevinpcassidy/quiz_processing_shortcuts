function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Formula Tools")
    .addItem("Fill Average Formulas", "fillAverageFormulaForSelectedColumn")
    .addItem("Check all Scores for Updates", "updateScoresFromSourceSheet")
    .addToUi();
}

/* ============================================================
   UNIVERSAL INSTANT-LOADING HTML ALERT + PROMPT
   ============================================================ */

function htmlAlert(message) {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial; font-size:12pt; padding:10px;">
      <div>${message}</div>
      <div style="margin-top:18px; text-align:right;">
        <button onclick="google.script.host.close()"
                style="font-size:12pt; padding:4px 12px;">OK</button>
      </div>
    </div>
  `)
    .setWidth(350)
    .setHeight(180);

  html.getContent(); // forces pre-render (stops blank window)
  SpreadsheetApp.getUi().showModalDialog(html, "Message");
}

//
// PROMPT: returns string by writing to a global variable
//

var __promptResponse = null;

function htmlPrompt(message) {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial; font-size:12pt; padding:10px;">
      <p>${message}</p>
      <input id="val" style="width:95%; font-size:12pt;"/>
      <div style="margin-top:16px; text-align:right;">
        <button style="font-size:12pt; padding:4px 12px;"
          onclick="
            const v=document.getElementById('val').value;
            google.script.run
              .withSuccessHandler(()=>google.script.host.close())
              ._storePromptValue(v);
          ">OK</button>

        <button style="font-size:12pt; padding:4px 12px;"
          onclick="google.script.host.close()">Cancel</button>
      </div>
    </div>
  `)
    .setWidth(380)
    .setHeight(220);

  html.getContent();
  SpreadsheetApp.getUi().showModalDialog(html, "Input Needed");
}

// Receives prompt values
function _storePromptValue(v) {
  __promptResponse = v;
}



/* ============================================================
   FILL AVERAGE FORMULA FOR SELECTED COLUMN
   ============================================================ */

function fillAverageFormulaForSelectedColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveCell();
  const targetCol = range.getColumn();

  if (targetCol <= 4) {
    htmlAlert("Please select a target column at least 5 or later (formula needs the 4 columns before it).");
    return;
  }

  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const startCol = targetCol - 4;

  const template =
    '=IF(COUNTA(%range%)=0,"",IFERROR(AVERAGE(LARGE(%range%,{1}),LARGE(%range%,{2})),MAX(%range%)))';

  for (let i = 0; i < names.length; i++) {
    const row = i + 2;
    if (names[i][0]) {
      const rangeA1 = sheet.getRange(row, startCol, 1, 4).getA1Notation();
      const formula = template.replace(/%range%/g, rangeA1);
      sheet.getRange(row, targetCol).setFormula(formula);
    }
  }

  htmlAlert("Formulas filled successfully!");
}


/* ============================================================
   UPDATE SCORES FROM SOURCE SHEET (SAFE DIALOG)
   ============================================================ */

function updateScoresFromSourceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial; font-size:12pt; padding:10px;">
      <h3 style="margin-top:0;">Select source sheet</h3>

      <form id="sheetForm">
        ${sheets
          .map(
            (name, i) => `
          <label style="display:block; margin:4px 0;">
            <input type="radio" name="sheetChoice" value="${i}" ${
              i === 0 ? "checked" : ""
            } />
            ${i + 1}: ${name}
          </label>
        `
          )
          .join("")}
      </form>

      <div style="margin-top:16px; text-align:right;">
        <button id="okBtn" style="font-size:12pt; padding:4px 12px;">OK</button>
        <button id="cancelBtn" style="font-size:12pt; padding:4px 12px;">Cancel</button>
      </div>

      <script>
        document.getElementById('okBtn').onclick = () => {
          const form = document.getElementById('sheetForm');
          const choice = form.sheetChoice.value;

          // Close fast:
          google.script.host.close();

          // Run update AFTER closing dialog:
          setTimeout(() => {
            google.script.run
              .withFailureHandler(err => alert('Error: ' + err.message))
              .updateScoresFromSourceSheet_withChoice(Number(choice));
          }, 10);
        };

        document.getElementById('cancelBtn').onclick = () => {
          google.script.host.close();
        };
      </script>
    </div>
  `)
    .setWidth(420)
    .setHeight(360);

  html.getContent();
  SpreadsheetApp.getUi().showModalDialog(html, "Choose source sheet");
}


function updateScoresFromSourceSheet_withChoice(choiceIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());

  if (choiceIndex < 0 || choiceIndex >= sheets.length) {
    htmlAlert("Invalid sheet choice.");
    return;
  }

  const targetSheet = ss.getActiveSheet();
  const sourceSheetName = sheets[choiceIndex];
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceLastRow = sourceSheet.getLastRow();

  if (sourceLastCol < 1 || sourceLastRow < 2) {
    htmlAlert("Source sheet is empty or missing data.");
    return;
  }

  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0];
  const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceLastCol).getValues();
  const sourceNames = sourceData.map(r => r[0]);

  const targetLastCol = targetSheet.getLastColumn();
  const targetLastRow = targetSheet.getLastRow();
  if (targetLastCol < 1 || targetLastRow < 2) {
    htmlAlert("Target sheet has no data to update.");
    return;
  }

  const targetHeaders = targetSheet.getRange(1, 1, 1, targetLastCol).getValues()[0];
  const targetNames = targetSheet.getRange(2, 1, targetLastRow - 1, 1).getValues().map(r => r[0]);

  // Map student → row
  const sourceMap = {};
  for (let i = 0; i < sourceNames.length; i++) {
    sourceMap[sourceNames[i]] = sourceData[i];
  }

  // Update matching headers
  for (let tCol = 0; tCol < targetHeaders.length; tCol++) {
    const header = targetHeaders[tCol];
    const sCol = sourceHeaders.indexOf(header);
    if (sCol === -1) continue;

    const colValues = [];
    for (let r = 0; r < targetNames.length; r++) {
      const name = targetNames[r];
      if (sourceMap[name] && sourceMap[name][sCol] !== "") {
        colValues.push([sourceMap[name][sCol]]);
      } else {
        colValues.push([
          targetSheet.getRange(r + 2, tCol + 1).getValue(),
        ]);
      }
    }

    targetSheet.getRange(2, tCol + 1, colValues.length, 1).setValues(colValues);
  }

  htmlAlert("Scores updated from " + sourceSheetName + "!");
}
