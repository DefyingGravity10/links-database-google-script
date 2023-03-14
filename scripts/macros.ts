/* Submission of a new entry */
class Entry {
  entryName: string;
  entryLink: string;
  entryTags: string;
}

function Submit() {
  const entrySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  const entry = new Entry();
  
  if (entrySheet !== null) {
    entry.entryName = entrySheet?.getRange("EntryName").getValue();
    entry.entryLink = entrySheet?.getRange("EntryLink").getValue();
    entry.entryTags = entrySheet?.getRange("EntryTags").getValue();
    
    // Both Name and Link are not supplied
    if (entry.entryName.length === 0 || entry.entryLink.length === 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Warning", "Please supply both the Name and Link.", ui.ButtonSet.OK); 
    }

    // The link supplied is not a valid url
    else if (!(entry.entryLink.includes("http://") || entry.entryLink.includes("https://"))){
      const ui = SpreadsheetApp.getUi();
      ui.alert("Warning", "Please supply a valid link", ui.ButtonSet.OK); 
    }
    else {
      addEntry(entry);
      entrySheet?.getRange("EntryAttr").clearContent(); 
    }
    
  }
};

function addEntry(entry: Entry) {
  const databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");

  if (databaseSheet !== null) {
    // Add a new row of cells
    const lastRow: number = databaseSheet?.getLastRow();
    let lastID: number = 0;

    if (lastRow !== 1) {
      lastID = databaseSheet?.getRange(`A${lastRow}`).getValue();
    }
    const newRow: number = lastRow + 1;
    databaseSheet.insertRowAfter(lastRow);

    const range = databaseSheet?.getRange(`A${newRow}:D${newRow}`);
    const values = [[lastID+1, entry.entryName, entry.entryLink, entry.entryTags]];
    range.setValues(values);
    range.setBackground("#fff2cc");
    range.setFontWeight("normal");
  }
};

/* Searching for specific entry(ies) */
class SearchItem {
  searchName: string;
  tags: string;
}

function Search() {
  const searchSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main");
  const searchItem = new SearchItem();
  
  if (searchSheet !== null) {
    searchItem.searchName = searchSheet?.getRange("SearchName").getValue();
    searchItem.tags = searchSheet?.getRange("SearchTags").getValue();

    if (searchItem.searchName.length === 0 && searchItem.tags.length === 0) {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Warning", "Please supply either a Name or some tags.", ui.ButtonSet.OK); 
    }
    else {
      updateResults(searchItem);

      // Move to the results tab
      const results = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
      if (results !== null) {
        SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(results);
      }
    }
  }
};

function updateResults(searchItem: SearchItem) {
  const db = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  const results = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");

  if (db !== null && results !== null) {
    const lastRowResults: number = results?.getLastRow();
    const lastRow: number = db?.getLastRow();

    // Delete all the rows to start with a "clean slate"
    if (lastRowResults > 1) {
      results.deleteRows(2, (lastRowResults-1));
    }

    let currRow: number = 1;
    for (let i:number = 2; i<=lastRow; i++) {
      if (searchItem.searchName.length !== 0 &&
          !db.getRange(`B${i}`).getValue().trim().toLowerCase().includes(searchItem.searchName.trim().toLowerCase())) {
        continue;
      }

      let dbArray = db.getRange(`D${i}`).getValue().split(",").map(item => item.trim().toLowerCase());
      let searchArray = searchItem.tags.split(",").map(item => item.trim().toLowerCase());

      dbArray.sort();
      searchArray.sort();

      if (searchArray.length > dbArray.length) {
        continue;
      }

      let count: number = 0;
      for (let j: number = 0; j < dbArray.length; j++) {
        for (let k: number = 0; k < searchArray.length; k++) {
          if (dbArray[j] === searchArray[k]) {
            ++count;
            break;
          }
        }
      }
      if ((count < searchArray.length && searchArray[0] !== "")) {
        continue;
      }
      
      // We can add the entry if we reached this point
      results.insertRowAfter(currRow);
      ++currRow;

      const range = results.getRange(`A${currRow}:D${currRow}`);
      const values = [[db.getRange(`A${i}`).getValue(), db.getRange(`B${i}`).getValue(), 
                      db.getRange(`C${i}`).getValue(), db.getRange(`D${i}`).getValue()]];
      range.setValues(values);
      range.setBackground("#fff2cc");
      range.setFontWeight("normal");
    }
  }
};