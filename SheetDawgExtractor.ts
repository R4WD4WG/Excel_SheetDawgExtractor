// R4WD4WG on GitHub

// Office Scripts for Excel
// Change extension to .osts if this does not work with Microsoft Excel.

// This script is used to collect data from multiple worksheets in an Excel workbook and write the data to a new worksheet.
// The script is designed to be run from the Excel UI, but can also be run from the command line using the Office Scripts CLI tool.


function main(workbook: ExcelScript.Workbook) {
    const sheetSuffix = 'TEXT ENDING HERE';
    const sheetNameLength = 8;
  
    // Create a new worksheet to store the collected data
    const outputWorksheet = workbook.addWorksheet('Collected Data');
    // Sels values of cell in (Row, Column) starting at 0
    outputWorksheet.getCell(0, 0).setValue('Value1');
    outputWorksheet.getCell(0, 1).setValue('Value2');
    // etc etc
  
    let rowIndex = 1;
  
    // Iterate through all worksheets in the workbook
    // For loop that is created based on # of sheets in excel file
    workbook.getWorksheets().forEach((worksheet) => {
      // Gets the worksheet name
      const sheetName = worksheet.getName();
      // Boolean for "target sheet" - does it fit the parameters set below? If so, goes true.
      let isTargetSheet = false;
  
      // Check if the sheet name ends with the target suffix
      if (sheetName.endsWith(sheetSuffix)) { // Does it with the suffix?
        isTargetSheet = true;
      } else if (sheetName.length === sheetNameLength && sheetName.endsWith('_NAV')) {
        // Check if the sheet name matches the parameter naming pattern
        const cellB2Value = worksheet.getRange('B2').getValue() as string;
        if (cellB2Value.includes('Parameter') || cellB2Value.includes('Parameter2')) { 
          // Matches length of Parameter names & has this contained in the name of the length? (Cell B2)
          isTargetSheet = true;
        }
      }
  
      if (isTargetSheet) {  // if isTargetSheet = True
        const columnB = worksheet.getRange('B:B');
        const columnF = worksheet.getRange('F:F');
        let value2Row: number | null = null;
        let value3Row: number | null = null;
        // getUsedRange - the range we are looking at (B:B for example)
        // getValues - gets the values of cell
        // forEach - initializes the for loop, for example:
        // For each cell in the range we select, look to see if the cell matches either values below
        columnB.getUsedRange().getValues().forEach((cellValue: (string | number)[], cellIndex: number) => {
          const cellText = (cellValue[0] as string).trim();
          // looks for TNA & Liabilities lines
          if (cellText === 'Value2') {
            value2Row = cellIndex;
          } else if (cellText === 'Value3') {
            value3row = cellIndex;
          }
        });
  
        if (value2Row !== null && value3row !== null) {
          const cellValue1 = worksheet.getRange('B2').getValues();
          const cellValue2 = columnF.getCell(value2Row, 0).getValues();
          const cellValue3 = columnF.getCell(value3row, 0).getValues();
  
          const cellValue4 = (cellValue2[0][0] as number) + (cellValue3[0][0] as number);
  
          // Extract  DataCode and Data Name from cell B2
          const DataInfo = (cellValue1[0][0] as string).split('Name:');
          const DataCode = DataInfo[0].replace('Word:', '').trim();
          const DataName = DataInfo[1].trim();
  
          // Write the collected data to the output sheet
          outputWorksheet.getCell(rowIndex, 0).setValue(sheetName);
          outputWorksheet.getCell(rowIndex, 1).setValue(cellValue2[0][0]);
          outputWorksheet.getCell(rowIndex, 1).setNumberFormat('#,##0.00');  // comma formatting with decimals
          outputWorksheet.getCell(rowIndex, 2).setValue(cellValue3[0][0]);
          outputWorksheet.getCell(rowIndex, 2).setNumberFormat('#,##0.00');  // comma formatting with decimals
          outputWorksheet.getCell(rowIndex, 3).setValue(cellValue4);
          outputWorksheet.getCell(rowIndex, 3).setNumberFormat('#,##0.00');  // comma formatting with decimals
          outputWorksheet.getCell(rowIndex, 4).setValue(DataCode);
          outputWorksheet.getCell(rowIndex, 5).setValue(DataName);
  
          rowIndex++;
        }
      }
    });
  }
