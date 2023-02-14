function main(workbook: ExcelScript.Workbook) {
  // Cell codes in Ana Panel
  const alisFirmaAdiCellCode = "G2";
  const alisFirmaIBANCellCode = "G3";
  const alisFirmaUnvanCellCode = "G4";
  // Cell codes in own alis table
  const borcCellCodeOwnTable = "H2";
  const odenenCellCodeOwnTable = "I2";
  const bakiyeCellCodeOwnTable = "J2";
  const unvanCellCodeOwnTable = "L2";
  const ibanCellCodeOwnTable = "M2";
  // Sheet Names
  const anaPanelSheetName = "Ana Panel";
  const alisFirmalarSheetName = "AlışFirmalar";
  // In AlisFirmalar sheet,
  const unvanColName = "C";
  const borcColName = "D";
  const odenenColName = "E";
  const bakiyeColName = "F";
  // MAXs
  const maxRowSize = "250";
  
  // Get the main form sheet.
  const anaPanelSheet = workbook.getWorksheet(anaPanelSheetName);
  // Get the value of alis firma kodu.
  let alisFirmaAdiCell = anaPanelSheet.getRange(alisFirmaAdiCellCode);
  let alisFirmaAdiData: string = alisFirmaAdiCell.getValues()[0][0].toString().toUpperCase();
  
  // Get the value of alis firma IBAN.
  let alisFirmaIBANCell = anaPanelSheet.getRange(alisFirmaIBANCellCode);
  let alisFirmaIBANData: string = alisFirmaIBANCell.getValues()[0][0].toString().toUpperCase();

  // Get the value of alis firma unvan.
  let alisFirmaUnvanCell = anaPanelSheet.getRange(alisFirmaUnvanCellCode);
  let alisFirmaUnvanData: string = alisFirmaUnvanCell.getValues()[0][0].toString().toUpperCase();
  

  /* CREATE NEW WORKSHEET */
  // Check if the "Data" worksheet already exists.
  let worksheet = workbook.getWorksheet(alisFirmaAdiData)
  if (worksheet) {
      console.log("İlgili girdi zaten daha önce yaratılmış.");
  } else {
      // Add a new worksheet.
    worksheet = workbook.addWorksheet(alisFirmaAdiData);
  }

  // Prettify new worksheet.
  worksheet.getRange("A1:F1").merge(false);
  worksheet.getRange("A1").setValue(alisFirmaAdiData + " Cari Hesap");
  worksheet.getRange("A1").getFormat().setRowHeight(50);
  
  worksheet.getRange("A2").setValue("Tarih");
  worksheet.getRange("A2").getFormat().setColumnWidth(75);
  worksheet.getRange("A2").getFormat().setRowHeight(50);
  
  worksheet.getRange("B2").setValue("Fatura No");
  worksheet.getRange("B2").getFormat().setColumnWidth(75);
  
  worksheet.getRange("C2").setValue("Açıklama");
  worksheet.getRange("C2").getFormat().setColumnWidth(150);
  
  worksheet.getRange("D2").setValue("Borç");
  worksheet.getRange("D2").getFormat().setColumnWidth(75);
  
  worksheet.getRange("E2").setValue("Ödenen");
  worksheet.getRange("E2").getFormat().setColumnWidth(75);
 
  worksheet.getRange("F2").setValue("Bakiye");
  worksheet.getRange("F2").getFormat().setColumnWidth(75);
  
  worksheet.getRange("H1").setValue("T. Borç");
  worksheet.getRange("H2").getFormat().setColumnWidth(75);
  
  worksheet.getRange("I1").setValue("T. Ödenen");
  worksheet.getRange("I2").getFormat().setColumnWidth(90);
  
  worksheet.getRange("J1").setValue("T. Bakiye");
  worksheet.getRange("J2").getFormat().setColumnWidth(75);
  
  worksheet.getRange("L1").setValue("Ünvan");
  worksheet.getRange("L2").getFormat().setColumnWidth(150);
  worksheet.getRange("L2").setValue(alisFirmaUnvanData);

  worksheet.getRange("M1").setValue("IBAN");
  worksheet.getRange("M2").getFormat().setColumnWidth(150);
  worksheet.getRange("M2").setValue(alisFirmaIBANData);

  const headerTabs = worksheet.getRange("A2:J2").getFormat().getFont();
  headerTabs.setName("Calibri");
  headerTabs.setSize(14);
  headerTabs.setBold(true);

  const headerName = worksheet.getRange("A1:M1").getFormat().getFont();
  headerName.setName("Calibri");
  headerName.setSize(18);
  headerName.setBold(true);

  const totalOdenenDataFormatting = worksheet.getRange("I2").getFormat();
  totalOdenenDataFormatting.getFont().setColor("#FF0000")
  totalOdenenDataFormatting.getFill().setColor("#D0CECE")


  // Assign the formulas.
  worksheet.getRange("H2").setFormulaLocal("=TOPLA($D$3:$D$" + maxRowSize + ")")
  worksheet.getRange("I2").setFormulaLocal("=TOPLA($E$3:$E$" + maxRowSize + ")")
  worksheet.getRange("J2").setFormulaLocal("=TOPLA($F$3:$F$" + maxRowSize + ")")

  /* APPEND THE NEW RECORD TO ALİSFIRMALAR SHEET */
  let newData = [
    [
    alisFirmaAdiData,
    alisFirmaIBANData,
    alisFirmaUnvanData
    ]
  ]
  const alinanSheet = workbook.getWorksheet(alisFirmalarSheetName);
  const numberOfRowsData = newData.length;
  const numberOfColumnsData = newData[0].length;
  const usedRangeForAlinan = alinanSheet.getUsedRange();
  const newRangeAlinan = usedRangeForAlinan
    .getOffsetRange(usedRangeForAlinan.getRowCount(), 0)
    .getAbsoluteResizedRange(
      numberOfRowsData, numberOfColumnsData
    );
  newRangeAlinan.setValues(newData)
  // Add the formula to C
  const formulaCellRowCount = (usedRangeForAlinan.getRowCount() + 1).toString();
  
  let borcFormulaRange = borcColName + formulaCellRowCount;
  let borcFormula = "='" + alisFirmaAdiData + "'!" + borcCellCodeOwnTable
  alinanSheet.getRange(borcFormulaRange).setFormula(borcFormula)
  // Add the formula to D
  let odenenFormulaRange = odenenColName + formulaCellRowCount;
  let odenenFormula = "='" + alisFirmaAdiData + "'!" + odenenCellCodeOwnTable
  alinanSheet.getRange(odenenFormulaRange).setFormula(odenenFormula)
  // Add the formula to D
  let bakiyeFormulaRange = bakiyeColName + formulaCellRowCount;
  let bakiyeFormula = "='" + alisFirmaAdiData + "'!" + bakiyeCellCodeOwnTable
  alinanSheet.getRange(bakiyeFormulaRange).setFormula(bakiyeFormula)


  // Clear the field.
  alisFirmaAdiCell.setValue("");
  alisFirmaIBANCell.setValue("");
  alisFirmaUnvanCell.setValue("");
}
