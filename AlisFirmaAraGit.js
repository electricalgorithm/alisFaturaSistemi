function main(workbook: ExcelScript.Workbook) {
  const alisFirmaAdiCellCode = "G10"
  const anaPanelSheetName = "Ana Panel"
  
  // Get the main form sheet.
  const anaPanelSheet = workbook.getWorksheet(anaPanelSheetName);
  // Get the value of alis firma adi.
  let alisFirmaAdiCell = anaPanelSheet.getRange(alisFirmaAdiCellCode);
  let alisFirmaAdiData: string = alisFirmaAdiCell.getValues()[0][0].toString().toUpperCase();
  
  let worksheet = workbook.getWorksheet(alisFirmaAdiData);
  worksheet.activate()

  // Clear the field.
  alisFirmaAdiCell.setValue("")
}
