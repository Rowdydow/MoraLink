// Moraware Integration Google Apps Script
// Configure the MORAWARE_BASE_URL constant below with your company's Moraware URL

// CONFIGURATION - UPDATE THIS WITH YOUR COMPANY'S MORAWARE URL
const MORAWARE_BASE_URL = 'https://yourcompany.moraware.net'; // Replace 'yourcompany' with your actual Moraware database name

function onOpen(e) {
  scrollToToday();
  addHyperlinksToJobs();
  scrollToTodayMobile();
}

function createTodayButton() {
  // Create a custom menu with a "Go to Today" button
  SpreadsheetApp.getUi()
    .createMenu('Schedule Tools')
    .addItem('Go to Today', 'scrollToTodayMobile')
    .addToUi();
}

function scrollToTodayMobile() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AUTOFAB');
  if (!sheet) return;
  
  const today = new Date();
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const dayOfWeek = today.getDay();
  
  // Skip weekends
  if (dayOfWeek === 0 || dayOfWeek === 6) return;
  
  // Format we're looking for
  const todayFormatted = `${days[dayOfWeek]} ${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear().toString().slice(-2)}`;
  
  // Create a temporary named range to help with navigation
  const lastRow = sheet.getLastRow();
  
  // Scan the entire sheet starting from row 1
  for (let i = 1; i <= lastRow; i++) {
    // Check column A and column C
    for (let col of [1, 3]) {
      const cellValue = sheet.getRange(i, col).getValue().toString().trim();
      
      if (cellValue === todayFormatted) {
        // Create a named range for today
        try {
          const namedRanges = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
          const existingRange = namedRanges.find(range => range.getName() === 'TodaySection');
          if (existingRange) existingRange.remove();
          
          // Create new named range
          SpreadsheetApp.getActiveSpreadsheet().setNamedRange('TodaySection', sheet.getRange(i, 1, 1, 1));
          
          // Set the active range
          sheet.setActiveRange(sheet.getRange(i, 1));
          
          // Try multiple approaches to ensure scrolling
          sheet.getRange(i, 1).activate();
          SpreadsheetApp.setActiveRange(sheet.getRange(i, 1));
          
          // Force sheet to redraw
          SpreadsheetApp.flush();
          
          return;
        } catch (e) {
          console.error('Error while creating named range: ' + e.toString());
        }
      }
    }
  }
}

function scrollToToday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AUTOFAB');
  if (!sheet) {
    Logger.log("Sheet not found");
    return;
  }

  const today = new Date();
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const dayOfWeek = today.getDay();
  
  // Format we're looking for
  const todayFormatted = `${days[dayOfWeek]} ${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear().toString().slice(-2)}`;
  Logger.log(`Looking for date: ${todayFormatted}`);

  // Don't run on weekends
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log("Weekend detected, stopping");
    return;
  }

  const lastRow = sheet.getLastRow();
  let targetRow = null;
  
  // Start searching from row 1
  const startRow = 1;

  // Look directly for the formatted date string in column C, starting from row 1
  for (let i = startRow; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 3).getValue().toString().trim();
    Logger.log(`Row ${i}: Found "${cellValue}"`);
    
    if (cellValue === todayFormatted) {
      targetRow = i;
      Logger.log(`Match found at row ${targetRow}`);
      break;
    }
  }

  if (targetRow) {
    Logger.log(`Scrolling to row ${targetRow}`);
    if (targetRow > 1) {
      // First hide everything above
      sheet.hideRows(1, targetRow - 1);
      
      // Set active selection to the target row
      sheet.setActiveRange(sheet.getRange(targetRow, 1));
      
      // Force a redraw
      SpreadsheetApp.flush();
      
      // Small delay to ensure the UI updates
      Utilities.sleep(100);
      
      // Show the rows again
      sheet.showRows(1, targetRow - 1);
      
      // Force another redraw
      SpreadsheetApp.flush();
      
      // Set the active selection again to ensure it stays in view
      sheet.setActiveRange(sheet.getRange(targetRow, 1));
      
      // Scroll to ensure the row is at the top
      sheet.getRange(targetRow, 1).activate();
    }
  } else {
    Logger.log("No target row found");
  }
}

function isSameDay(date1, date2) {
  return date1.getDate() === date2.getDate() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getFullYear() === date2.getFullYear();
}

function findFirstEmptyRow(sheet, startRow, column) {
  const lastRow = sheet.getLastRow();
  for (let i = startRow; i <= lastRow; i++) {
    const range = sheet.getRange(i, column, 1, 2);
    const values = range.getValues()[0];
    const backgrounds = range.getBackgrounds()[0];
    
    if ((!values[0] || values[0] === '') && (!values[1] || values[1] === '')) {
      const bgColor = backgrounds[0].toLowerCase();
      if (bgColor !== '#ffa500' && bgColor !== '#ff9900') {
        return i;
      }
    }
  }
  return null;
}

function findTodayDateRow(sheet) {
  const today = new Date();
  const dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return null;
  
  const lastRow = sheet.getLastRow();
  let dateRows = [];
  
  for (let i = 1; i <= lastRow; i++) {
    const bgColor = sheet.getRange(i, 2).getBackground().toLowerCase();
    if (bgColor === '#ffa500' || bgColor === '#ff9900') {
      dateRows.push(i);
    }
  }
  
  return dateRows[dayOfWeek - 1] || null;
}

function addHyperlinksToJobs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoFabSheet = ss.getSheetByName('AUTOFAB');
  const jobIndexSheet = ss.getSheetByName('JobIndex');
  
  if (!autoFabSheet || !jobIndexSheet) {
    console.error('Required sheets not found');
    return;
  }

  // Get both Moraware IDs and job numbers from JobIndex
  const morawareRange = jobIndexSheet.getRange('B:B').getValues();
  const jobNumberRange = jobIndexSheet.getRange('C:C').getValues();
  
  // Create a mapping between job numbers and Moraware IDs
  const jobMapping = new Map();
  
  for (let i = 0; i < jobNumberRange.length; i++) {
    const jobNumber = String(jobNumberRange[i][0]).trim();
    const morawareId = String(morawareRange[i][0]).trim();
    
    // Skip empty cells
    if (!jobNumber || !morawareId) continue;
    
    // Remove any non-numeric characters from morawareId
    const cleanMorawareId = morawareId.replace(/\D/g, '');
    
    // Only add if we have a valid 5-digit Moraware ID
    if (cleanMorawareId.length === 5) {
      jobMapping.set(jobNumber, cleanMorawareId);
    }
  }

  // Define columns to process
  const columnsToProcess = [2, 6, 9];  // B, F, I
  
  // Get the full range of the AutoFab sheet
  const lastRow = autoFabSheet.getLastRow();
  
  // Start processing from row 1
  const startRow = 1;
  
  // Process each column
  for (const col of columnsToProcess) {
    // Get the range starting from row 1
    const range = autoFabSheet.getRange(startRow, col, lastRow - startRow + 1);
    const values = range.getValues();
    
    values.forEach((value, rowIndex) => {
      const cellValue = String(value[0]).trim();
      
      // Skip empty cells
      if (!cellValue) return;
      
      // Adjust rowIndex to account for starting at row 1
      const actualRow = startRow + rowIndex;
      const jobCell = autoFabSheet.getRange(actualRow, col);
      const richTextValue = jobCell.getRichTextValue();
      
      // Skip if cell already has a hyperlink
      if (richTextValue && richTextValue.getLinkUrl()) return;
      
      // Check if this job number exists in our mapping
      if (jobMapping.has(cellValue)) {
        const morawareId = jobMapping.get(cellValue);
        // Use the configurable base URL
        const url = `${MORAWARE_BASE_URL}/sys/job/${morawareId}`;
        
        // Get the font color from the description cell
        const descCell = autoFabSheet.getRange(actualRow, col + 1);
        const fontColor = descCell.getFontColor();
        
        try {
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(cellValue)
            .setLinkUrl(0, cellValue.length, url)
            .build();
          
          jobCell.setRichTextValue(richText);
          jobCell.setFontColor(fontColor);  // Match the description cell's font color
        } catch (error) {
          console.error(`Error setting hyperlink for ${cellValue} at row ${actualRow}: ${error}`);
        }
      }
    });
  }
}

function onChange(e) {
  try {
    if (!e) return;
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AUTOFAB');
    if (!sheet) return;
    
    const range = sheet.getActiveRange();
    const row = range.getRow();
    const col = range.getColumn();
    const value = range.getValue();
    const oldValue = e.oldValue || '';
    
    // Check for "installs tomorrow" in Saw schedule first
    if (col === 3) { // Column C
      const lowercaseValue = value.toString().toLowerCase();
      if (lowercaseValue.includes('installs tomorrow') || 
          lowercaseValue.includes('installs tom') || 
          lowercaseValue.includes('inst tom')) {
        const todayRow = findTodayDateRow(sheet);
        if (todayRow) {
          const jobNumber = sheet.getRange(row, 2).getValue();
          const jobDesc = sheet.getRange(row, 3).getValue();
          const firstEmptyRow = findFirstEmptyRow(sheet, todayRow + 1, 9);
          
          if (firstEmptyRow) {
            sheet.getRange(firstEmptyRow, 9).setValue(jobNumber);
            sheet.getRange(firstEmptyRow, 10).setValue(jobDesc);
            sheet.getRange(firstEmptyRow, 9, 1, 2).setFontColor('#000000');
            sheet.getRange(firstEmptyRow, 9, 1, 2).setBackground('#00ff00');
            sheet.getRange(firstEmptyRow, 9, 1, 2).setFontSize(19);
          }
        }
      }
    }
    
    const isSawSection = (col === 2 || col === 3);
    const isCNCSection = (col === 6 || col === 7);
    
    if (!isSawSection && !isCNCSection) return;
    
    const activeCell = sheet.getRange(row, col);
    const cellColor = activeCell.getFontColor();
    const currentBgColor = activeCell.getBackground().toLowerCase();
    
    // Handle CNC to Polish movement when text is red
    if (isCNCSection && cellColor && cellColor.toLowerCase() === '#ff0000') {
      // Get the job details
      const jobNumber = sheet.getRange(row, 6).getValue();
      const jobDesc = sheet.getRange(row, 7).getValue();

      // Start from current row and find the next orange row
      let targetDateRow = null;
      for (let i = row + 1; i <= sheet.getLastRow(); i++) {
          const range = sheet.getRange(i, 2, 1, 10);
          const backgrounds = range.getBackgrounds()[0];
          if (backgrounds.every(color => color.toLowerCase() === '#ffa500' || color.toLowerCase() === '#ff9900')) {
              targetDateRow = i;
              break;
          }
      }

      // If we found the next date row, add to its Polish section
      if (targetDateRow) {
          // Find first empty row after the date row
          let targetRow = targetDateRow + 1;
          while (targetRow <= sheet.getLastRow()) {
              const checkRange = sheet.getRange(targetRow, 9, 1, 2);
              const values = checkRange.getValues()[0];
              if (!values[0] && !values[1]) {  // If both cells are empty
                  // Add the job here
                  sheet.getRange(targetRow, 9).setValue(jobNumber);
                  sheet.getRange(targetRow, 10).setValue(jobDesc);
                  sheet.getRange(targetRow, 9, 1, 2).setFontColor('#000000');
                  sheet.getRange(targetRow, 9, 1, 2).setBackground('#00ff00');
                  sheet.getRange(targetRow, 9, 1, 2).setFontSize(19);
                  break;
              }
              targetRow++;
          }
      }

      // Set backgrounds to white for original cells
      if (currentBgColor !== '#ffa500' && currentBgColor !== '#ff9900') {
          sheet.getRange(row, 6).setBackground('white');
          sheet.getRange(row, 7).setBackground('white');
      }
    }

    // Handle saw to CNC movement
    if (isSawSection && cellColor && cellColor.toLowerCase() === '#ff0000') {
      if (currentBgColor !== '#ffa500' && currentBgColor !== '#ff9900') {
        sheet.getRange(row, 2).setBackground('white');
        sheet.getRange(row, 3).setBackground('white');
        
        let currentRow = row;
        let dateHeaderRow = null;
        
        while (currentRow > 0) {
          const bgColor = sheet.getRange(currentRow, 2).getBackground();
          if (bgColor === '#ffa500' || bgColor === '#ff9900') {
            dateHeaderRow = currentRow;
            break;
          }
          currentRow--;
        }
        
        if (dateHeaderRow) {
          const jobNumber = sheet.getRange(row, 2).getValue();
          const jobDesc = sheet.getRange(row, 3).getValue();
          const jobNotes = sheet.getRange(row, 4).getValue();
          
          const isNoCNC = (jobDesc && jobDesc.toString().toUpperCase().includes('NO CNC')) || 
                         (jobNotes && jobNotes.toString().toUpperCase().includes('NO CNC'));
          
          const cncTargetRow = moveJob(sheet, dateHeaderRow, jobNumber, jobDesc, 6, false);
          
          if (isNoCNC && cncTargetRow) {
            sheet.getRange(cncTargetRow, 6).setFontColor('#ff0000').setBackground('white');
            sheet.getRange(cncTargetRow, 7).setFontColor('#ff0000').setBackground('white');
            
            moveJob(sheet, dateHeaderRow, jobNumber, jobDesc, 9, true);
          }
        }
      }
    }
    
  } catch (error) {
    console.error(error);
  }
}

function moveJob(sheet, dateHeaderRow, jobNumber, jobDesc, targetColumn, nextDay) {
  let targetDateRow = dateHeaderRow;
  if (nextDay) {
    const lastRow = sheet.getLastRow();
    for (let i = dateHeaderRow + 1; i <= lastRow; i++) {
      const range = sheet.getRange(i, 2, 1, 10);
      const backgrounds = range.getBackgrounds()[0];
      if (backgrounds.every(color => 
          color.toLowerCase() === '#ffa500' || 
          color.toLowerCase() === '#ff9900')) {
          targetDateRow = i;
          break;
      }
    }
  }
  
  const startSearchRow = targetDateRow + 1;
  const lastRow = sheet.getLastRow();
  let isDuplicate = false;
  
  for (let i = startSearchRow; i <= lastRow; i++) {
    const existingJobNumber = sheet.getRange(i, targetColumn).getValue();
    const existingJobDesc = sheet.getRange(i, targetColumn + 1).getValue();
    
    if (existingJobNumber === jobNumber && existingJobDesc === jobDesc) {
      isDuplicate = true;
      break;
    }
    
    const range = sheet.getRange(i, targetColumn);
    const bgColor = range.getBackground().toLowerCase();
    if (bgColor === '#ffa500' || bgColor === '#ff9900') {
      break;
    }
  }
  
  if (!isDuplicate) {
    let targetRow = null;
    for (let i = startSearchRow; i <= lastRow; i++) {
      const range = sheet.getRange(i, targetColumn, 1, 2);
      const values = range.getValues()[0];
      
      if ((!values[0] || values[0] === '') && (!values[1] || values[1] === '')) {
        const bgColor = range.getBackgrounds()[0][0].toLowerCase();
        if (bgColor !== '#ffa500' && bgColor !== '#ff9900') {
          targetRow = i;
          break;
        }
      }
      
      const bgColor = range.getBackgrounds()[0][0].toLowerCase();
      if (bgColor === '#ffa500' || bgColor === '#ff9900') {
        break;
      }
    }
    
    if (targetRow) {
      const targetCell = sheet.getRange(targetRow, targetColumn);
      targetCell.setValue(jobNumber);
      targetCell.setFontColor('#000000');
      targetCell.setBackground('#00ff00');
      targetCell.setFontSize(19);
      
      const descCell = sheet.getRange(targetRow, targetColumn + 1);
      descCell.setValue(jobDesc);
      descCell.setFontColor('#000000');
      descCell.setFontSize(19);
      
      // Add hyperlinks after moving the job
      addHyperlinksToJobs();
      
      return targetRow;
    }
  }
  
  return null;
}

function addTomorrowInstallsToPolish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jobIndexSheet = ss.getSheetByName('JobIndex');
  const autoFabSheet = ss.getSheetByName('AUTOFAB');
  
  if (!jobIndexSheet || !autoFabSheet) {
    Logger.log('Required sheets not found');
    return;
  }

  // Get tomorrow's date (skipping weekends)
  const tomorrow = getNextWorkday();
  const tomorrowFormatted = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "MM/dd/yyyy");
  
  Logger.log(`Looking for installs scheduled for: ${tomorrowFormatted}`);
  
  // Get install dates from JobIndex
  const installDates = jobIndexSheet.getRange('F:F').getValues(); // Assuming install dates are in column F
  const jobNumbers = jobIndexSheet.getRange('C:C').getValues();
  const jobDescs = jobIndexSheet.getRange('D:D').getValues();
  
  // Find today's row and tomorrow's row in AUTOFAB
  const sheet = autoFabSheet;
  const lastRow = sheet.getLastRow();
  
  // Get today and tomorrow in the expected format
  const today = new Date();
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const diasEspanol = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  
  const todayDayOfWeek = today.getDay();
  const tomorrowDayOfWeek = tomorrow.getDay();
  
  const todayFormatted = `'${days[todayDayOfWeek]} ${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear().toString().slice(-2)}`;
  const tomorrowFormattedSpanish = `'${diasEspanol[tomorrowDayOfWeek]} ${tomorrow.getMonth() + 1}/${tomorrow.getDate()}/${tomorrow.getFullYear().toString().slice(-2)}`;
  
  let tomorrowRow = null;
  
  // Find tomorrow's row by looking for the Spanish date in column J
  for (let i = 1; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 10).getValue().toString().trim(); // Column J is index 10
    if (cellValue === tomorrowFormattedSpanish || cellValue === tomorrowFormattedSpanish.substring(1)) {
      tomorrowRow = i;
      break;
    }
  }
  
  if (!tomorrowRow) {
    Logger.log('Could not find tomorrow\'s row');
    return;
  }
  
  Logger.log(`Found tomorrow's row at: ${tomorrowRow}`);
  
  let jobsAdded = 0;

  // Check each install date
  installDates.forEach((dateCell, index) => {
    // Skip empty cells
    if (!dateCell[0]) return;
    
    const dateStr = String(dateCell[0]);
    // Try to parse the date (handle different formats)
    let installDate;
    try {
      if (dateStr.includes('/')) {
        installDate = Utilities.formatDate(new Date(dateStr), Session.getScriptTimeZone(), "MM/dd/yyyy");
      } else {
        // Handle Excel date serial numbers
        installDate = Utilities.formatDate(new Date(dateCell[0]), Session.getScriptTimeZone(), "MM/dd/yyyy");
      }
    } catch (e) {
      Logger.log(`Error parsing date: ${dateStr}`);
      return;
    }
    
    if (installDate === tomorrowFormatted) {
      const jobNumber = jobNumbers[index][0];
      const jobDesc = jobDescs[index][0];
      
      // Skip if either jobNumber or jobDesc is empty
      if (!jobNumber || !jobDesc) return;
      
      Logger.log(`Found job installing tomorrow - Job: ${jobNumber}, Desc: ${jobDesc}`);
      
      // Find first empty row in Polish section starting after tomorrow's date row
      const firstEmptyRow = findFirstEmptyRow(sheet, tomorrowRow + 1, 9);
      if (firstEmptyRow) {
        // Double check the row is actually empty
        const checkRange = sheet.getRange(firstEmptyRow, 9, 1, 2);
        const checkValues = checkRange.getValues()[0];
        if (!checkValues[0] && !checkValues[1]) {
          sheet.getRange(firstEmptyRow, 9).setValue(jobNumber);
          sheet.getRange(firstEmptyRow, 10).setValue(jobDesc);
          sheet.getRange(firstEmptyRow, 9, 1, 2)
            .setBackground('#00ff00')
            .setFontSize(19)
            .setFontColor('#000000');
          
          jobsAdded++;
          Logger.log(`Added job to Polish at row ${firstEmptyRow}`);
        }
      } else {
        Logger.log('No empty row found in Polish section');
      }
    }
  });
  
  Logger.log(`Process complete. Added ${jobsAdded} jobs to Polish.`);
}

// Helper function to get next workday (skipping weekends)
function getNextWorkday() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  
  // If tomorrow is Saturday (6), move to Monday
  if (tomorrow.getDay() === 6) {
    tomorrow.setDate(tomorrow.getDate() + 2);
  }
  // If tomorrow is Sunday (0), move to Monday
  else if (tomorrow.getDay() === 0) {
    tomorrow.setDate(tomorrow.getDate() + 1);
  }
  
  return tomorrow;
}
