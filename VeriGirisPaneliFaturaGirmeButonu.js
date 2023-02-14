function main(workbook: ExcelScript.Workbook) {
  // The cell range information.
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
  let alisFirmaData: string = anaPanelSheet
                      .getRange(alisFirmaCell).getValues()[0][0]
                      .toString();
  let santiyeFirmaData: string = anaPanelSheet
                      .getRange(santiyeFirmaCell).getValues()[0][0]
                      .toString();
  let tarihData = anaPanelSheet
                      .getRange(tarihCell).getValues()[0][0];
  let faturaNoData: string = anaPanelSheet
                      .getRange(faturaNoCell).getValues()[0][0]
                      .toString();
  let aciklamaData: string = anaPanelSheet
                      .getRange(aciklamaCell).getValues()[0][0]
                      .toString();
  let kdvTipiData = anaPanelSheet
                      .getRange(kdvTipiCell).getValues()[0][0];
  let tutarData = anaPanelSheet
                      .getRange(tutarCell).getValues()[0][0];
  let odenenData = anaPanelSheet.getRange(odenenCell).getValues()[0][0];

  // Get the number of row for the next row in the spesific sheets.
  const alinanSheet = workbook.getWorksheet(alisFirmaData);
  const santiyeSheet = workbook.getWorksheet(santiyeFirmaData);
  const alinanSheetNextRow = alinanSheet
                              .getUsedRange()
                              .getValues().length + 1;
  const santiyeSheetNextRow = santiyeSheet
                              .getUsedRange()
                              .getValues().length + 1;

  let kdvliTutar: number; 
  let kdvsizTutar: number;
  if (tutarData == "") {
    kdvliTutar = 0
    kdvsizTutar = 0
  } else {
    kdvliTutar = tutarData;
    kdvsizTutar = kdvliTutar / (1 + kdvTipiData)
  }

  let newDataForAlinan = [
    [
      tarihData,
      faturaNoData,
      aciklamaData,
      tutarData,
      odenenData
    ]
  ]

  // Save into the alinan table.
  const numberOfRowsAlinan: number = newDataForAlinan.length;
  const numberOfColumnsAlinan: number = newDataForAlinan[0].length;
  const usedRangeAlinan = alinanSheet.getUsedRange();
  const newRangeAlinan = usedRangeAlinan
    .getOffsetRange(usedRangeAlinan.getRowCount(), 0)
    .getAbsoluteResizedRange(numberOfRowsAlinan, numberOfColumnsAlinan);
  newRangeAlinan.setValues(newDataForAlinan)
  const formulaCellRowCount = usedRangeAlinan.getRowCount() + 1
  const formulaCellColName = "F"
  let formulaCellRangeName = formulaCellColName + formulaCellRowCount.toString()
  let formula = "=$D$" + formulaCellRowCount.toString() + "-$E$" + formulaCellRowCount.toString()
  alinanSheet.getRange(formulaCellRangeName).setFormula(formula)

  if (tutarData != "") {
    let newDataForSantiye = [
      [
        tarihData,
        faturaNoData,
        aciklamaData,
        kdvsizTutar,
        kdvliTutar
      ]
    ]

    // Save into the santiye table.
    const numberOfRowsSantiye: number = newDataForSantiye.length;
    const numberOfColumnsSantiye: number = newDataForSantiye[0].length;
    const usedRangeSantiye = santiyeSheet.getUsedRange();
    const newRangeSantiye = usedRangeSantiye
      .getOffsetRange(usedRangeSantiye.getRowCount(), 0)
      .getAbsoluteResizedRange(numberOfRowsSantiye, numberOfColumnsSantiye);
    newRangeSantiye.setValues(newDataForSantiye)
  }
}
