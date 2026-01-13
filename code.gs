/**
 * LinkedIn Profile Finder Add-on
 * Uses LinkFinder AI API to find LinkedIn profile URLs from names and companies
 */

// Add menu when add-on is installed
function onInstall(e) {
  onOpen(e);
}

// Add menu when document is opened
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Find LinkedIn Profiles', 'showSidebar')
    .addItem('Settings', 'showSettings')
    .addItem('Help', 'showHelp')
    .addToUi();
}

// Show sidebar for finding profiles
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('LinkedIn Profile Finder')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Show settings dialog
function showSettings() {
  var html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'API Settings');
}

// Show help dialog
function showHelp() {
  var html = HtmlService.createHtmlOutputFromFile('Help')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Help & Documentation');
}

// Save API key to user properties
function saveApiKey(apiKey) {
  if (!apiKey || apiKey.trim() === '') {
    throw new Error('API key cannot be empty');
  }
  PropertiesService.getUserProperties().setProperty('LINKFINDER_API_KEY', apiKey.trim());
  return { success: true, message: 'API key saved successfully' };
}

// Get API key from user properties
function getApiKey() {
  return PropertiesService.getUserProperties().getProperty('LINKFINDER_API_KEY');
}

// Check if API key is configured
function isApiKeyConfigured() {
  var apiKey = getApiKey();
  return apiKey && apiKey.trim() !== '';
}

// Start column selection mode
function startColumnSelectionMode() {
  return true;
}

// Find LinkedIn profiles using selected columns
function findLinkedInProfilesFromSelection(nameColumn, companyColumn, outputColumn) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var apiKey = getApiKey();
  
  if (!apiKey) {
    throw new Error('API key not configured. Please set your API key in Settings.');
  }
  
  // Convert column letters to numbers if needed
  nameColumn = columnLetterToNumber(nameColumn);
  outputColumn = columnLetterToNumber(outputColumn);
  if (companyColumn) {
    companyColumn = columnLetterToNumber(companyColumn);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Sheet needs at least 2 rows (header + data)');
  }
  
  // Start from row 2 (skip header)
  var startRow = 2;
  var numRows = lastRow - startRow + 1;
  
  var names = sheet.getRange(startRow, nameColumn, numRows, 1).getValues();
  var companies = companyColumn ? sheet.getRange(startRow, companyColumn, numRows, 1).getValues() : [];
  
  var results = [];
  var processedCount = 0;
  var errorCount = 0;
  
  for (var i = 0; i < names.length; i++) {
    var name = names[i][0];
    var company = companies.length > 0 ? companies[i][0] : '';
    
    if (!name || name.toString().trim() === '') {
      results.push(['']);
      continue;
    }
    
    var inputData = name.toString().trim();
    if (company && company.toString().trim() !== '') {
      inputData += ' ' + company.toString().trim();
    }
    
    try {
      var linkedInUrl = callLinkFinderApi(apiKey, inputData);
      results.push([linkedInUrl]);
      processedCount++;
      Utilities.sleep(500); // Rate limiting
    } catch (error) {
      results.push(['ERROR: ' + error.message]);
      errorCount++;
    }
  }
  
  sheet.getRange(startRow, outputColumn, results.length, 1).setValues(results);
  
  return {
    success: true,
    processed: processedCount,
    errors: errorCount,
    total: names.length
  };
}

// Call LinkFinder AI API
function callLinkFinderApi(apiKey, inputData) {
  var url = 'https://api.linkfinderai.com';
  
  var payload = {
    'type': 'lead_full_name_to_linkedin_url',
    'input_data': inputData
  };
  
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    Logger.log('Calling API with data: ' + inputData);
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log('Response code: ' + responseCode);
    Logger.log('Response text: ' + responseText);
    
    if (responseCode !== 200) {
      throw new Error('API request failed with status ' + responseCode + ': ' + responseText);
    }
    
    var result = JSON.parse(responseText);
    
    if (result.status === 'success' && result.result) {
      return result.result;
    } else if (result.status === 'error') {
      Logger.log('API returned error: ' + JSON.stringify(result));
      return 'Not found';
    } else {
      Logger.log('Unexpected response: ' + JSON.stringify(result));
      return 'Not found';
    }
  } catch (error) {
    Logger.log('API Error: ' + error.message);
    throw new Error('API request failed: ' + error.message);
  }
}

// Convert column letter to number
function columnLetterToNumber(column) {
  // If already a number, return it
  if (typeof column === 'number') {
    return column;
  }
  
  // If it's a string that's actually a number
  if (!isNaN(column)) {
    return parseInt(column);
  }
  
  // Convert letter(s) to number
  column = column.toUpperCase();
  var result = 0;
  for (var i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

// Convert column number to letter
function columnNumberToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
