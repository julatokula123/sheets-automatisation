function sortujKolumnyWedlugTak() {
  const arkusz = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sheetname'); // Zmień nazwę arkusza jeśli inna
  const startCol = 16; // Kolumna P = 16
  const endCol = 120;  // Kolumna DP = 120
  const startRow = 7;  // nagłówki w wierszu 7
  const endRow = 38;   // dane do wiersza 38

  const liczbaWierszy = endRow - startRow + 1;
  const liczbaKolumn = endCol - startCol + 1;

  
  const zakres = arkusz.getRange(startRow, startCol, liczbaWierszy, liczbaKolumn);
  const dane = zakres.getValues();

  
  let kolumny = [];
  for(let c = 0; c < liczbaKolumn; c++) {
    let liczTak = 0;
    for(let r = 1; r < liczbaWierszy; r++) {
      if (String(dane[r][c]).toLowerCase() === "tak") liczTak++;
    }
    kolumny.push({index: c, liczTak: liczTak});
  }

  
  kolumny.sort((a, b) => b.liczTak - a.liczTak);

  
  let noweDane = [];
  for(let r = 0; r < liczbaWierszy; r++) {
    let nowyWiersz = [];
    for(let i = 0; i < liczbaKolumn; i++) {
      nowyWiersz.push(dane[r][kolumny[i].index]);
    }
    noweDane.push(nowyWiersz);
  }

  
  zakres.setValues(noweDane);
}
