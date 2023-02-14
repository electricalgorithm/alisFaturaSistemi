function main(workbook: ExcelScript.Workbook) {
  const alisFirmaCell = "B2"
  const santiyeFirmaCell = "B3"
  const tarihCell = "B4"
  const faturaNoCell = "B5"
  const aciklamaCell = "B6"
  const kdvTipiCell = "B7"
  const tutarCell = "B8"
  const odenenCell = "B9"

  // Get data from the form.
  let anaPanelSheet = workbook.getWorksheet("Ana Panel");
  anaPanelSheet.getRange(alisFirmaCell).setValue("");
  anaPanelSheet.getRange(santiyeFirmaCell).setValue("");
  anaPanelSheet.getRange(tarihCell).setValue("");
  anaPanelSheet.getRange(faturaNoCell).setValue("");
  anaPanelSheet.getRange(aciklamaCell).setValue("");
  anaPanelSheet.getRange(kdvTipiCell).setValue("");
  anaPanelSheet.getRange(tutarCell).setValue("");
  anaPanelSheet.getRange(odenenCell).setValue("");
}
