function main(workbook: ExcelScript.Workbook) {
  const santiyeAdiCellCode = "G12"
  const anaPanelSheetName = "Ana Panel"

  // Get the main form sheet.
  const anaPanelSheet = workbook.getWorksheet(anaPanelSheetName);
  // Get the value of alis firma adi.
  let santiyeAdiCell = anaPanelSheet.getRange(santiyeAdiCellCode);
  let santiyeAdiData: string = santiyeAdiCell.getValues()[0][0].toString().toUpperCase();

  let worksheet = workbook.getWorksheet(santiyeAdiData);
  worksheet.activate()

  // Clear the field.
  santiyeAdiCell.setValue("")
}
