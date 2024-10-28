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
    
    // Create border styles
    const greenBorder = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const greenColor = '#00FF00';
    const redBorder = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const redColor = '#FF0000';

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
        const invoiceNumber = (sfRow[2] || '').toString().trim(); // Column C
        const columnHValue = (sfRow[7] || '').toString().trim(); // Column H
        const bankInfo = sfRow[9]; // Column J
        const columnIValue = (sfRow[8] || '').toString().trim(); // Column I
        
        // Skip if either name is empty
        if (!beneficiaryName) {
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
          } else {
            return; // Skip if bankInfo is empty
          }
        } catch (e) {
          Logger.log('Error parsing JSON for row: ' + (rowIndex + 2) + '. Error: ' + e.message);
          return;
        }
        
        // Check if names match (case-insensitive and trimmed)
        if (beneficiaryName === bankInfoObj.beneficiaryName.toLowerCase().trim()) {
          // Set green border for beneficiary name cell
          const nameCell = transactionsSheet.getRange(rowIndex + 2, 23); // Column W
          nameCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          
          // 1. Verify transaction number (Tab1.A = Tab2.C)
          if (invoiceNumber === transactionNumber) {
            const transactionCell = transactionsSheet.getRange(rowIndex + 2, 1); // Column A
            transactionCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          } else {
            const transactionCell = transactionsSheet.getRange(rowIndex + 2, 1); // Column A
            transactionCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
          }
          
          // 2. Verify destination currency
          if (bankInfoObj.destinationCurrency === destinationCurrency) {
            const currencyCell = transactionsSheet.getRange(rowIndex + 2, 2); // Column B
            currencyCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          } else {
            const currencyCell = transactionsSheet.getRange(rowIndex + 2, 2); // Column B
            currencyCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
          }

          // 3. Verify Invoice Currency (Column E - only green border, no red)
          if (columnEValue === columnIValue) {
            const columnECell = transactionsSheet.getRange(rowIndex + 2, 5); // Column E
            columnECell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          }
          
          // 4. Verify beneficiary account number
          if (bankInfoObj.beneficiaryAccountNumber) {
            const cleanBankInfoAccount = bankInfoObj.beneficiaryAccountNumber.replace(/\s+/g, '');
            const cleanTransactionAccount = beneficiaryAccountNumber.replace(/\s+/g, '');
            
            if (cleanBankInfoAccount === cleanTransactionAccount) {
              const accountCell = transactionsSheet.getRange(rowIndex + 2, 38); // Column AL
              accountCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
            } else {
              const accountCell = transactionsSheet.getRange(rowIndex + 2, 38); // Column AL
              accountCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
            }
          }
          
          // 5. Verify routing code value
          if (bankInfoObj.routingCodeValue1 && routingCodeValue1) {
            const cleanBankInfoRouting = bankInfoObj.routingCodeValue1.replace(/\s+/g, '');
            const cleanTransactionRouting = routingCodeValue1.replace(/\s+/g, '');
            
            if (cleanBankInfoRouting === cleanTransactionRouting) {
              const routingCell = transactionsSheet.getRange(rowIndex + 2, 43); // Column AQ
              routingCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
            } else {
              const routingCell = transactionsSheet.getRange(rowIndex + 2, 43); // Column AQ
              routingCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
            }
          }

          // 6. Verify if Column C or F matches Column H from SF_RAW
            const cMatches = columnCValue === columnHValue;
            const fMatches = columnFValue === columnHValue;

            // Apply green border to Column C if it matches
            if (cMatches) {
              const columnCCell = transactionsSheet.getRange(rowIndex + 2, 3); // Column C
              columnCCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
            } else if (!cMatches && !fMatches) {
              // Only apply red border to Column C if neither C nor F matches
              const columnCCell = transactionsSheet.getRange(rowIndex + 2, 3); // Column C
              columnCCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
            }

            // Apply green border to Column F if it matches
            if (fMatches) {
              const columnFCell = transactionsSheet.getRange(rowIndex + 2, 6); // Column F
              columnFCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
            } else if (!cMatches && !fMatches) {
              // Only apply red border to Column F if neither C nor F matches
              const columnFCell = transactionsSheet.getRange(rowIndex + 2, 6); // Column F
              columnFCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
            }
        }
      });
    });

    return {
      success: true,
      message: "Processing completed successfully! Matching cells are highlighted in green, mismatches in red."
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
    const columnsToCheck = [1, 2, 3, 5, 6, 23, 38, 43]; // A, B, C, E, F, W, AL, AQ
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

function moveFileToPaymentUploads(spreadsheetId) {
  try {
    // Get the file by ID
    const file = DriveApp.getFileById(spreadsheetId);
    const originalName = file.getName();
    
    // Get current date and format it
    const currentDate = new Date();
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const formattedDate = `${monthNames[currentDate.getMonth()]} ${currentDate.getDate()}, ${currentDate.getFullYear()}`;
    
    // Rename file based on the pattern: everything before "-" + formatted date
    const nameParts = originalName.split('-');
    if (nameParts.length > 1) {
      const newName = `${nameParts[0].trim()} - ${formattedDate}`;
      file.setName(newName);
    }
    
    // Get or create month folder (format: MM.YYYY Payment Uploads)
    const parentFolder = DriveApp.getFolderById('15V7lVpj5kaF1YjzhZkXvicb6diSV9thF');
    const monthFolderName = `${(currentDate.getMonth() + 1).toString().padStart(2, '0')}.${currentDate.getFullYear()} Payment Uploads`;
    
    // Check if month folder exists, if not create it
    let monthFolder;
    const monthFolders = parentFolder.getFoldersByName(monthFolderName);
    if (monthFolders.hasNext()) {
      monthFolder = monthFolders.next();
    } else {
      monthFolder = parentFolder.createFolder(monthFolderName);
    }
    
    // Get or create date folder (format: Month DD, YYYY)
    const dateFolderName = formattedDate;
    let dateFolder;
    const dateFolders = monthFolder.getFoldersByName(dateFolderName);
    if (dateFolders.hasNext()) {
      dateFolder = dateFolders.next();
    } else {
      dateFolder = monthFolder.createFolder(dateFolderName);
    }
    
    // Move file to the appropriate date folder
    file.moveTo(dateFolder);
    
    return {
      success: true,
      message: `File successfully moved to ${monthFolderName}/${dateFolderName}`
    };
  } catch (error) {
    return {
      success: false,
      message: "Error moving file: " + error.toString()
    };
  }
}
