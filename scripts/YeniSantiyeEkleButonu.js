function main(workbook: ExcelScript.Workbook) {
  const santiyeAdiCellCode = "G6"
  const anaPanelSheetName = "Ana Panel"
  const santiyelerSheetName = "Şantiyeler"
  // In AlisFirmalar sheet,
  const kdvliColName = "B";
  const kdvsizColName = "C";

  // Get the main form sheet.
  const anaPanelSheet = workbook.getWorksheet(anaPanelSheetName);
  // Get the value of alis firma kodu.
  let santiyeAdiCell = anaPanelSheet.getRange(santiyeAdiCellCode);
  let santiyeAdiData: string = santiyeAdiCell.getValues()[0][0]
                                  .toString().toUpperCase();

  /* CREATE NEW WORKSHEET */
  // Check if the "Data" worksheet already exists.
  let worksheet = workbook.getWorksheet(santiyeAdiData)
  if (worksheet) {
      console.log("İlgili şantiye zaten daha önce yaratılmış.");
  } else {
      // Add a new worksheet.
      worksheet = workbook.addWorksheet(santiyeAdiData);
  }

  // Prettify new worksheet.
  worksheet.getRange("A1:E1").merge(false);
  worksheet.getRange("H1:I1").merge(false);
  worksheet.getRange("H2:I2").merge(false);
  worksheet.getRange("A1").setValue(santiyeAdiData + " Hesap");
  worksheet.getRange("A1").getFormat().setRowHeight(50);
  worksheet.getRange("A2").setValue("Tarih");
  worksheet.getRange("A2").getFormat().setColumnWidth(100);
  worksheet.getRange("A2").getFormat().setRowHeight(50);
  worksheet.getRange("B2").setValue("Fatura No");
  worksheet.getRange("B2").getFormat().setColumnWidth(150);
  worksheet.getRange("C2").setValue("Açıklama");
  worksheet.getRange("C2").getFormat().setColumnWidth(150);
  worksheet.getRange("D2").setValue("KDV'siz");
  worksheet.getRange("D2").getFormat().setColumnWidth(100);
  worksheet.getRange("E2").setValue("KDV'li");
  worksheet.getRange("E2").getFormat().setColumnWidth(100);
  worksheet.getRange("H1").setValue("Toplam KDV'siz:");
  worksheet.getRange("H1").getFormat().setColumnWidth(100);
  worksheet.getRange("H2").setValue("Toplam KDV'li:");
  worksheet.getRange("H2").getFormat().setColumnWidth(100);

  const headerName = worksheet.getRange("A1:E1").getFormat().getFont();
  headerName.setName("Calibri");
  headerName.setSize(18);
  headerName.setBold(true);

  const headerTabs = worksheet.getRange("A2:E2").getFormat().getFont();
  headerTabs.setName("Calibri");
  headerTabs.setSize(14);
  headerTabs.setBold(true);

  // Assign the formulas.
  worksheet.getRange("J1").setFormulaLocal("=TOPLA($D$3:$D$100)")
  worksheet.getRange("J2").setFormulaLocal("=TOPLA($E$3:$E$100)")

  /* APPEND THE NEW RECORD TO ALİSFIRMALAR SHEET */
  let newData = [[santiyeAdiData]]
  const alinanSheet = workbook.getWorksheet(santiyelerSheetName);
  const numberOfRowsData = newData.length;
  const numberOfColumnsData = newData[0].length;
  const usedRangeForAlinan = alinanSheet.getUsedRange();
  const newRangeAlinan = usedRangeForAlinan
      .getOffsetRange(usedRangeForAlinan.getRowCount(), 0)
      .getAbsoluteResizedRange(
          numberOfRowsData, numberOfColumnsData
      );
  newRangeAlinan.setValues(newData)
  // Add the formula to B
  const formulaCellRowCount = (usedRangeForAlinan.getRowCount() + 1).toString();
  let kdvliFormulaRange = kdvliColName + formulaCellRowCount;
  let kdvliFormula = "='" + santiyeAdiData + "'!J1"
  alinanSheet.getRange(kdvliFormulaRange).setFormula(kdvliFormula)
  // Add the formula to C
  let kdvsizFormulaRange = kdvsizColName + formulaCellRowCount;
  let kdvsizFormula = "='" + santiyeAdiData + "'!J2"
  alinanSheet.getRange(kdvsizFormulaRange).setFormula(kdvsizFormula)
  // Clear the field.
  santiyeAdiCell.setValue("")
}
