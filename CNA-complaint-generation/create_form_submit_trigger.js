function createFormSubmitTrigger() {
  deleteTrigger("setFormulasInInstertedRow");

  ScriptApp.newTrigger("setFormulasInInstertedRow")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();
}

function deleteTrigger(name) {
  var triggers = getProjectTriggersByName(name);
  for (var i = 0; i < triggers.length; i++)
    ScriptApp.deleteTrigger(triggers[i]);
}

function getProjectTriggersByName(name) {
  return ScriptApp.getProjectTriggers().filter(
    (trigger) => trigger.getHandlerFunction() === name
  );
}

/**
 * Set formulas in the new row inserted by form submission.
 */
function setFormulasInInstertedRow(event) {
  var sheet = event.range.getSheet();
  var newRowIndex = event.range.getRow();
  Logger.log("NEW ROW: " + newRowIndex);
  sheet
    .getRange("N" + newRowIndex)
    .setFormula(
      "VLOOKUP(G" + newRowIndex + ";'pt formule'!$B$2:$C$11;2;FALSE)"
    );
  Logger.log("SET FORMULA TO CELL: " + "N" + newRowIndex);
  sheet
    .getRange("O" + newRowIndex)
    .setFormula(
      "VLOOKUP(I" + newRowIndex + ";'pt formule'!$B$2:$C$11;2;FALSE)"
    );
  Logger.log("SET FORMULA TO CELL: " + "O" + newRowIndex);
  sheet
    .getRange("P" + newRowIndex)
    .setFormula(
      "VLOOKUP(K" + newRowIndex + ";'pt formule'!$B$2:$C$11;2;FALSE)"
    );
  Logger.log("SET FORMULA TO CELL: " + "P" + newRowIndex);
  sheet
    .getRange("Q" + newRowIndex)
    .setFormula(
      "VLOOKUP(M" + newRowIndex + ";'pt formule'!$B$2:$C$11;2;FALSE)"
    );
  Logger.log("SET FORMULA TO CELL: " + "Q" + newRowIndex);
}
