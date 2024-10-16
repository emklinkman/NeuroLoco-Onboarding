// NeuroLoco Automated Onboarding Email script
// University of Michigan Robotics Department
// EK Klinkman
// Version 1.1 10.15.2024
// Version 1.2 10.16.2024 Moved global variables to beginning


// Populate global variables
var spreadsheetId = '1d5nM5QMJWwudZskRu_OyLfEBrwZsk75yHI5kKpbcN-o'; // Replace with spreadsheet ID***
var sheetName = 'Form Responses 1'; // Specific sheet name***
var labManagerEmail = "emilykk@umich.edu"; // Person to be notified of new hire***      
var htmlTemplate = HtmlService.createTemplateFromFile('welcome_email.html'); // insert .html file of email to be sent to new hire***
var emailBody = "Welcome to the NeuroLoco team! Important onboarding info inside" // fill with desired email text

// Apps script code

function onFormSubmit(e) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  
  var lastRow = sheet.getLastRow(); // Get the last row with data
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()); // Get the data starting from row 2 (assuming row 1 has headers)
  var data = dataRange.getValues(); // Get all data in the range
  
  // Loop through each row of the sheet
  data.forEach(function(rowData, rowIndex) {
    var processed = rowData[11]; // Assuming column L (index 11) is the "Processed" column***
  
    // Check if the row has been processed
    if (processed !== "Yes") {
      var newHireName = rowData[2]; // Adjust column as needed (e.g., Name in column B)***
      var newHireEmail = rowData[9].toString().trim(); // Column J for email, clean up spaces with trim()***
      var mentorEmail = rowData[10].toString().trim(); // Column K for mentor email, clean up spaces with trim()***
      var startDate = rowData[6]; // Adjust column for start date (e.g., column G)***

    // Log the email address for debugging
      Logger.log('New Hire Email: ' + newHireEmail);
      Logger.log('Mentor Email: ' + mentorEmail);
      
      // Check if the email address is valid
     if (validateEmail(newHireEmail) && validateEmail(mentorEmail)) {
        
        // Load the HTML content from welcome-email.html and pass the newHireName to the template
        htmlTemplate.newHireName = newHireName; // Pass the new hire's name into the HTML template
        var htmlBody = htmlTemplate.evaluate().getContent();
        
        // Send HTML email to the new hire
        GmailApp.sendEmail(newHireEmail, emailBody, "", {
          htmlBody: htmlBody,
          cc: mentorEmail // CC the mentor email
        });
        
        // Notify lab manager
        var managerSubject = "New Hire: NeuroLoco: " + newHireName; // change quote text with desired email subject
        var managerBody = "A new hire has been added to the lab via Google Form. Please add them to: Mcommunity, Slack, NeuroLoco shared calendar, and GitHub.\n\nName: " + newHireName + "\nStart Date: " + startDate + "\nUM Email Address: " + newHireEmail; // change quote text with desired email subject + variables to include in report email

        // Send email to the lab manager
        GmailApp.sendEmail(labManagerEmail, managerSubject, managerBody);
        
        // Mark the row as processed in the "Processed" column (column L)
        sheet.getRange(rowIndex + 2, 12).setValue("Yes"); // Write "Yes" in column L (index 11)
      } else {
        Logger.log('Invalid email address: ' + newHireEmail); // Log invalid email
      }
    }
  });
}

// Function to validate email format
function validateEmail(email) {
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}
