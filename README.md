# Automail-from-googlesheet
function sendConsolidatedEmails() {
  // Get active spreadsheet and sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Assume first row is headers
  var headers = data[0];
  
  // Find index of email column (modify if different)
  var emailColumnIndex = headers.indexOf('Email');
  
  // Skip email column when creating table
  var dataColumns = headers.filter((h, idx) => idx !== emailColumnIndex);
  
  // Group data by unique email
  var emailGroups = {};
  
  // Start from second row (index 1) to skip headers
  for (var i = 1; i < data.length; i++) {
    var rowData = data[i];
    var email = rowData[emailColumnIndex];
    
    // Skip rows with empty email
    if (!email) continue;
    
    // Group data for this email
    if (!emailGroups[email]) {
      emailGroups[email] = [];
    }
    
    // Add row data excluding email column
    var rowDataWithoutEmail = rowData.filter((_, idx) => idx !== emailColumnIndex);
    emailGroups[email].push(rowDataWithoutEmail);
  }
  
  // Send consolidated emails
  for (var email in emailGroups) {
    var groupData = emailGroups[email];
    
    // Create HTML table
    var htmlTable = '<table border="1" cellpadding="5" style="border-collapse: collapse;">';
    
    // Add headers
    htmlTable += '<tr>' + 
      dataColumns.map(col => `<th style="background-color: #f2f2f2;">${col}</th>`).join('') + 
      '</tr>';
    
    // Add data rows
    groupData.forEach(row => {
      htmlTable += '<tr>' + 
        row.map(cell => `<td>${cell || ''}</td>`).join('') + 
        '</tr>';
    });
    
    htmlTable += '</table>';
    
    // Send email
    MailApp.sendEmail({
      to: email,
      subject: 'Invoice Pending',
      htmlBody: 'Invoices are pending, kindly send invoices ASAP.<br><br>' + htmlTable
    });
  }
}

// Add a menu item to trigger the script
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Email Automation')
    .addItem('Send Consolidated Emails', 'sendConsolidatedEmails')
    .addToUi();
}
