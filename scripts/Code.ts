function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Functionalities")
      .addSubMenu(ui.createMenu("Database")
        .addItem("Update ID", "updateIdNumber")
        .addItem("Remove Entry", "removeEntry"))
      .addSubMenu(ui.createMenu("Results")
        .addItem("Clear Page", "clearPage"))
      .addToUi(); 
}

function updateIdNumber() {
  const db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");

  if (db !== null) {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(db);

    const ui = SpreadsheetApp.getUi();
    let result = ui.alert("Confirmation", 
                "Are you sure you want to update the ID numbers?", 
                ui.ButtonSet.YES_NO);

    if (result == ui.Button.YES) {
      const range = db?.getRange(`A2:A${db?.getLastRow()}`);

      let IdNumbers: any = [];
      for (let i:number=1; i<db?.getLastRow(); i++) {
        const arr: any[] = [];
        arr.push(i)
        IdNumbers.push(arr);
      }
      range.setValues(IdNumbers);
    }
  }
};

function removeEntry() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("Delete an Entry", 
              "Please input the ID of the entry that you wish to delete", 
              ui.ButtonSet.OK_CANCEL);
  
  const button = result.getSelectedButton();
  const text = result.getResponseText();

  if (button == ui.Button.OK) {
    const result1 = ui.alert("Confirmation", `Are you sure that you want to delete entry with ID ${text}?
                            This action cannot be undone.`, ui.ButtonSet.YES_NO);
    if (result1 == ui.Button.YES) {
      const db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
      if (db !== null && Number(text) > 0 && Number(text) <= db?.getLastRow()) {
        // Find the right entry
        let id: number = 0;
        let flag: boolean = false;

        for (let i:number=2; i<=db?.getLastRow(); i++) {
          id = db?.getRange(`A${i}`).getValue();
          
          if (id === Number(text)) {
            db.deleteRow(i);
            flag = true;
            break;
          }
        }
        if (!flag) {
          ui.alert("Notice", `Sorry, we are unable to find the entry with ID ${text}`, ui.ButtonSet.OK);
        }
      }
      else {
        ui.alert("Notice", `Sorry, we are unable to find the entry with ID ${text}`, ui.ButtonSet.OK);
      }
    }
  }
};

function clearPage() {
  const results = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");

  if (results !== null) {
    const lastRow: number = results?.getLastRow();

    // Delete all the rows to start with a "clean slate"
    if (lastRow > 1) {
      results.deleteRows(2, (lastRow-1));
    }
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(results);
  }
};