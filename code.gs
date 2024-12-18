function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Transaction Data Verifier')
    .setFaviconUrl('https://www.google.com/images/spreadsheets/sheets_app_icon.png');
}

function processSpreadsheet(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const transactionsSheet = ss.getSheetByName('Transactions');
    const sfRawSpreadsheet = SpreadsheetApp.openById('19rkSHP7fkOl4aFJ6MtUx7EVclsKGCCb7PzWNvNPDjIc');
const sfRawSheet = sfRawSpreadsheet.getSheetByName('SF_RAW');
    
    if (!transactionsSheet || !sfRawSheet) {
        return { 
            success: false, 
            message: "Error: Could not find 'Transactions' sheet in current spreadsheet or 'SF_RAW' sheet in reference spreadsheet." 
        };
    }

    const transactionsData = transactionsSheet.getRange(2, 1, transactionsSheet.getLastRow() - 1, transactionsSheet.getLastColumn()).getValues();
    const sfRawData = sfRawSheet.getRange(2, 1, sfRawSheet.getLastRow() - 1, sfRawSheet.getLastColumn()).getValues();
    
    const greenBorder = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const greenColor = '#00FF00';
    const redBorder = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    const redColor = '#FF0000';

    // Create a map to store SF_RAW entries by beneficiary name
    const sfRawMap = new Map();
    sfRawData.forEach((sfRow, index) => {
      const bankInfo = sfRow[9];
      if (!bankInfo) return;
      
      try {
        const bankInfoObj = JSON.parse(bankInfo);
        if (bankInfoObj.beneficiaryAccountNumber) {
          bankInfoObj.beneficiaryAccountNumber = bankInfoObj.beneficiaryAccountNumber.trim();
        }
        if (bankInfoObj.routingCodeValue1) {
          bankInfoObj.routingCodeValue1 = bankInfoObj.routingCodeValue1.trim();
        }
        
        const name = bankInfoObj.beneficiaryName.toLowerCase().trim();
        if (!sfRawMap.has(name)) {
          sfRawMap.set(name, []);
        }
        sfRawMap.get(name).push({
          row: sfRow,
          bankInfoObj: bankInfoObj
        });
      } catch (e) {
        Logger.log('Error parsing JSON in SF_RAW: ' + e.message);
      }
    });

    // Process each transaction row
    transactionsData.forEach((transactionRow, rowIndex) => {
      const beneficiaryName = (transactionRow[22] || '').toString().toLowerCase().trim();
      const transactionNumber = (transactionRow[0] || '').toString().trim();
      const destinationCurrency = transactionRow[1];
      const beneficiaryAccountNumber = (transactionRow[37] || '').toString().trim();
      const routingCodeValue1 = (transactionRow[42] || '').toString().trim();
      const columnCValue = (transactionRow[2] || '').toString().trim();
      const columnFValue = (transactionRow[5] || '').toString().trim();
      const columnEValue = (transactionRow[4] || '').toString().trim();

      if (!beneficiaryName) return;

      const matchingSfRows = sfRawMap.get(beneficiaryName) || [];

      if (matchingSfRows.length > 0) {
        // Set green border for beneficiary name
        const nameCell = transactionsSheet.getRange(rowIndex + 2, 23);
        nameCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);

        // Try to find exact transaction number match first
        let bestMatch = matchingSfRows.find(match => 
          (match.row[2] || '').toString().trim() === transactionNumber
        ) || matchingSfRows[0]; // If no match, use first entry

        const sfRow = bestMatch.row;
        const bankInfoObj = bestMatch.bankInfoObj;
        const invoiceNumber = (sfRow[2] || '').toString().trim();
        const columnHValue = (sfRow[7] || '').toString().trim();
        const columnIValue = (sfRow[8] || '').toString().trim();

        // Transaction number verification
        if (invoiceNumber === transactionNumber) {
          const transactionCell = transactionsSheet.getRange(rowIndex + 2, 1);
          transactionCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
        } else {
          const transactionCell = transactionsSheet.getRange(rowIndex + 2, 1);
          transactionCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
        }

        // Currency verification
        if (bankInfoObj.destinationCurrency === destinationCurrency) {
          const currencyCell = transactionsSheet.getRange(rowIndex + 2, 2);
          currencyCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
        } else {
          const currencyCell = transactionsSheet.getRange(rowIndex + 2, 2);
          currencyCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
        }

        // Invoice Currency verification
        if (columnEValue === columnIValue) {
          const columnECell = transactionsSheet.getRange(rowIndex + 2, 5);
          columnECell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
        }

        // Account number verification
        if (bankInfoObj.beneficiaryAccountNumber) {
          const cleanBankInfoAccount = bankInfoObj.beneficiaryAccountNumber.replace(/\s+/g, '');
          const cleanTransactionAccount = beneficiaryAccountNumber.replace(/\s+/g, '');
          
          if (cleanBankInfoAccount === cleanTransactionAccount) {
            const accountCell = transactionsSheet.getRange(rowIndex + 2, 38);
            accountCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          } else {
            const accountCell = transactionsSheet.getRange(rowIndex + 2, 38);
            accountCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
          }
        }

        // Routing code verification
        if (bankInfoObj.routingCodeValue1 && routingCodeValue1) {
          const cleanBankInfoRouting = bankInfoObj.routingCodeValue1.replace(/\s+/g, '');
          const cleanTransactionRouting = routingCodeValue1.replace(/\s+/g, '');
          
          if (cleanBankInfoRouting === cleanTransactionRouting) {
            const routingCell = transactionsSheet.getRange(rowIndex + 2, 43);
            routingCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
          } else {
            const routingCell = transactionsSheet.getRange(rowIndex + 2, 43);
            routingCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
          }
        }

        // Column C and F verification
        const cMatches = columnCValue === columnHValue;
        const fMatches = columnFValue === columnHValue;

        if (cMatches) {
          const columnCCell = transactionsSheet.getRange(rowIndex + 2, 3);
          columnCCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
        } else if (!cMatches && !fMatches) {
          const columnCCell = transactionsSheet.getRange(rowIndex + 2, 3);
          columnCCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
        }

        if (fMatches) {
          const columnFCell = transactionsSheet.getRange(rowIndex + 2, 6);
          columnFCell.setBorder(true, true, true, true, null, null, greenColor, greenBorder);
        } else if (!cMatches && !fMatches) {
          const columnFCell = transactionsSheet.getRange(rowIndex + 2, 6);
          columnFCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
        }

      } else {
        // No matching beneficiary name found
        const nameCell = transactionsSheet.getRange(rowIndex + 2, 23);
        nameCell.setBorder(true, true, true, true, null, null, redColor, redBorder);
      }
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
