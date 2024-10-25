function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Transaction Data Verifier')
    .setFaviconUrl('https://www.google.com/images/spreadsheets/sheets_app_icon.png');
}

function processSpreadsheet(spreadsheetId) {
  try {
    // Get the spreadsheet by ID
    const ss = SpreadsheetApp.openById(spreadsheetId);
    
    // Get both sheets
    const transactionsSheet = ss.getSheetByName('Transactions');
    const sfRawSheet = ss.getSheetByName('SF_RAW');
    
    if (!transactionsSheet || !sfRawSheet) {
      return {
        success: false,
        message: "Error: Could not find sheets named 'Transactions' and 'SF_RAW'. Please verify sheet names."
      };
    }

    // Get all data from both sheets (starting from row 2)
    const transactionsData = transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, transactionsSheet.getLastColumn()).getValues();
    const sfRawData = sfRawSheet.getRange(2, 1, sfRawSheet.getLastRow() - 1, sfRawSheet.getLastColumn()).getValues();
    
    // Create green border style
    const greenBorder = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const greenColor = '#00FF00';

    // Process each row in the transactions sheet
    transactionsData.forEach((transactionRow, rowIndex) => {
      // Get relevant values from transactions sheet and trim whitespace where needed
      const beneficiaryName = (transactionRow[22] || '').toString().toLowerCase().trim(); // Column W
      const transactionNumber = (transactionRow[0] || '').toString().trim(); // Column A
      const destinationCurrency = transactionRow[1]; // Column B - No trim (as specified)
      const beneficiaryAccountNumber = (transactionRow[37] || '').toString().trim(); // Column AL
      const routingCodeValue1 = (transactionRow[42] || '').toString().trim(); // Column AQ
      const columnCValue = (transactionRow[2] || '').toString().trim(); // Column C
      const columnFValue = (transactionRow[5] || '').toString().trim(); // Column F
      const columnEValue = (transactionRow[4] || '').toString().trim(); // Column E
      
      // Find matching row in SF_RAW
      sfRawData.forEach((sfRow) => {
        // Trim whitespace from SF_RAW values except specified columns
        const contractorName = (sfRow[1] || '').toString().toLowerCase().trim(); // Column B
        const invoiceNumber = (sfRow[2] || '').toString().trim(); // Column C
        const columnHValue = (sfRow[7] || '').toString().trim(); // Column H
        const bankInfo = sfRow[9]; // Column J
        const columnIValue = (sfRow[8] || '').toString().trim(); // Column I
        
        // Skip if either name is empty
        if (!beneficiaryName || !contractorName) {
          return;
        }

        // Parse JSON in bank info column
        let bankInfoObj;
        try {
          if (bankInfo) {
            bankInfoObj = JSON.parse(bankInfo);
            // Trim whitespace from JSON values if they exist
            if (bankInfoObj.beneficiaryAccountNumber) {
              bankInfoObj.beneficiaryAccountNumber = bankInfoObj.beneficiaryAccountNumber.trim();
            }
            if (bankInfoObj.routingCodeValue1) {
              bankInfoObj.routingCodeValue1 = bankInfoObj.routingCodeValue1.trim();
            }
            // Don't trim destinationCurrency as specified
          } else {
            return; // Skip if bankInfo is empty
          }
        } catch (e) {
          Logger.log('Error parsing JSON for row: ' + (rowIndex + 2) + '. Error: ' + e.message);
          return;
        }
        
        // Check if names match (case-insensitive and trimmed)
        if (beneficiaryName === contractorName) {
          // Set green border for beneficiary name cell
          const nameCell = transactionsSheet.getRange(rowIndex + 2, 23); // Column W
          nameCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          
          // 1. Verify transaction number (Tab1.A = Tab2.C)
          if (invoiceNumber === transactionNumber) {
            const transactionCell = transactionsSheet.getRange(rowIndex + 2, 1); // Column A
            transactionCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          }
          
          // 2. Verify destination currency from JSON (Tab1.B matches destinationCurrency in Tab2.J)
          if (bankInfoObj.destinationCurrency === destinationCurrency) {
            const currencyCell = transactionsSheet.getRange(rowIndex + 2, 2); // Column B
            currencyCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          }

          // 2. Verify Invoice Currency
          if (columnEValue === columnIValue) {
          const columnECell = transactionsSheet.getRange(rowIndex + 2, 5); // Column E
          columnECell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          }
          
          // 4. Verify beneficiary account number (Tab1.AL matches beneficiaryAccountNumber in Tab2.J)
          if (bankInfoObj.beneficiaryAccountNumber) {
            // Remove spaces from both strings for comparison
            const cleanBankInfoAccount = bankInfoObj.beneficiaryAccountNumber.replace(/\s+/g, '');
            const cleanTransactionAccount = beneficiaryAccountNumber.replace(/\s+/g, '');
            
            if (cleanBankInfoAccount === cleanTransactionAccount) {
              const accountCell = transactionsSheet.getRange(rowIndex + 2, 38); // Column AL
              accountCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
            }
          }
          
          // 5. Verify routing code value (Tab1.AQ matches routingCodeValue1 in Tab2.J)
          if (bankInfoObj.routingCodeValue1 && routingCodeValue1) {
            const cleanBankInfoRouting = bankInfoObj.routingCodeValue1.replace(/\s+/g, '');
            const cleanTransactionRouting = routingCodeValue1.replace(/\s+/g, '');
            
            if (cleanBankInfoRouting === cleanTransactionRouting) {
              const routingCell = transactionsSheet.getRange(rowIndex + 2, 43); // Column AQ
              routingCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
            }
          }

          // 6. Verify if Column C or F matches Column H from SF_RAW
          if (columnCValue === columnHValue) {
            const columnCCell = transactionsSheet.getRange(rowIndex + 2, 3); // Column C
            columnCCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          }
          
          if (columnFValue === columnHValue) {
            const columnFCell = transactionsSheet.getRange(rowIndex + 2, 6); // Column F
            columnFCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          }
        }
      });
    });

    return {
      success: true,
      message: "Processing completed successfully! The matching cells have been highlighted in green."
    };
    
  } catch (error) {
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}

function clearVerification(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const transactionsSheet = ss.getSheetByName('Transactions');
    const lastRow = transactionsSheet.getLastRow();
    
    // Clear borders for all relevant columns
    const columnsToCheck = [1, 2, 3, 5, 6, 23, 38, 43]; // A, B, C, F, W, AL, AQ
    columnsToCheck.forEach(col => {
      const range = transactionsSheet.getRange(2, col, lastRow - 1, 1);
      range.setBorder(false, false, false, false, false, false);
    });

    return {
      success: true,
      message: "All verification highlights have been cleared."
    };
  } catch (error) {
    return {
      success: false,
      message: "Error clearing verification: " + error.toString()
    };
  }
}
