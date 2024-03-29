const typeformFormId = "<INSERT FORM ID>";
const questionNumber = 2;
const personalAccessToken = "<INSERT_ACCESS_TOKEN>";
const refreshResultsIntervalMinutes = 5; 
//My Google Sheet for this project: https://docs.google.com/spreadsheets/d/1txsefS4sJ55nogtc-jxUhXtd1r_cze9zbjRqauBPrq4/edit#gid=115302059

function updateTodayClasses(){
  let todaysClasses =  getTodaysClasses();
  let form = getTypeformForm(typeformFormId, personalAccessToken)

  var classOptions = todaysClasses.map(function(classItem) {
    return { label: classItem.className + ' @ ' + classItem.startTime + ' w/ '+  classItem.staff };
  });

  form.fields[questionNumber-1].properties.choices = classOptions;
  updateTypeformForm(typeformFormId, personalAccessToken,form)
}

function refreshResults(){
 var lastRefreshDate = new Date(new Date().getTime() - ((refreshResultsIntervalMinutes+5)  * 60 * 1000));

  let formResponses = fetchFormResponses(typeformFormId, "since="+lastRefreshDate);
  // console.log(JSON.stringify(formResponses))

  let formattedResponses = formResponses.map(function(formResponse){
      return  [formResponse.submitted_at,...formResponse.answer_values]
  })
  // console.log(formattedResponses)
  saveUniqueToSheet("Feedback Responses",formattedResponses, 0)
}

function fetchFormResponses(formId, qsp) {
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + personalAccessToken
    },
    'muteHttpExceptions': true
  };

  let query = ""
  if(qsp){
    query = "?="+qsp
  }
  
  // Fetch responses from the form's API
  var FORM_API_URL = `https://api.typeform.com/forms/${formId}/responses${query}`; // Example API URL
  var response = UrlFetchApp.fetch(FORM_API_URL, options);
  var jsonResponse = JSON.parse(response.getContentText());
  let formResponses = jsonResponse.items;

  let formattedAnswers = formResponses.map(function(response){
      let answers = response.answers.map(function(answer){
      if(answer.type === "number"){
        return answer.number;
      }
      else if(answer.type === "choice"){
        return answer.choice.label
      }else{
        return "unknown answer type";
      }
    });
    return {
        ...response,
        answer_values: answers
      }
  });
  return formattedAnswers;
}


function saveUniqueToSheet(sheetName, values2D, uniqueKeyIndex) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    // If the sheet doesn't exist, create it
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  var existingDataRange = sheet.getDataRange();
  var existingValues = existingDataRange.getValues();
  
  // Assuming the first value in each row is a unique identifier for the response
  var existingIds = existingValues.map(function(row) { return row[uniqueKeyIndex]; });
  
  var uniqueValues = values2D.filter(function(row) {
    return existingIds.indexOf(row[0]) === -1; // Filter out rows that already exist
  });
  
  if (uniqueValues.length > 0) {
    // Append new unique rows to the sheet
    uniqueValues.forEach(function(row) {
      sheet.appendRow(row);
    });
  }
}

function getTodaysClasses() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Class Session Export"); // Target specific sheet
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // Get today's date formatted to match your spreadsheet's date format
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "d-MMM-yy");

  console.log(dateString,values[250][2])
  var todaysClasses = values.filter(function(row) {
    return Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "d-MMM-yy") === dateString; // Assuming the 'Date' is in the third column
  });

  // Sort today's classes by time
  todaysClasses.sort(function(a, b) {
    return new Date('1970/01/01 ' + a[3]) - new Date('1970/01/01 ' + b[3]); // Assuming the 'Time' is in the fourth column
  });
  
  // Prepare the data to be logged or manipulated
  var classInfo = todaysClasses.map(function(row) {
    var formattedTime = Utilities.formatDate(row[3], Session.getScriptTimeZone(), "h:mma").toLowerCase(); // Formats the time

    return {
      className: row[0].split('(')[0], // Class name is in the first column
      staff: row[7],     // Staff name is in the eighth column
      startTime: formattedTime  // Start time is in the fourth column
    };
  });
  
  // Example of logging the results
  classInfo.forEach(function(classItem) {
    Logger.log(classItem.className + ' @ ' + classItem.startTime + ' w/ '+  classItem.staff );
  });
  return classInfo;
}

function getTypeformForm(formId, personalAccessToken) {
  var options = {
    'method' : 'get',
    'headers' : {
      'Authorization': 'Bearer ' + personalAccessToken
    },
    'muteHttpExceptions': true
  };
  
  var url = 'https://api.typeform.com/forms/' + formId;
  
  var response = UrlFetchApp.fetch(url, options);
  var responseData = response.getContentText();
  var responseJson = JSON.parse(responseData);
  
  // Logger.log(responseJson); // This will log the entire form object
  
  // You can also process the form data as needed, e.g., extract specific information
  // For example, to log the form title:
  if (responseJson && responseJson.title) {
    Logger.log("Form Title: " + responseJson.title);
    return responseJson
  }
}

function updateTypeformForm(formId, personalAccessToken, updatedFormData) {
  var options = {
    method: 'put',
    headers: {
      Authorization: 'Bearer ' + personalAccessToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(updatedFormData),
    muteHttpExceptions: true
  };

  var url = 'https://api.typeform.com/forms/' + formId;

  var response = UrlFetchApp.fetch(url, options);
  var responseData = response.getContentText();
  var responseJson = JSON.parse(responseData);

  Logger.log(responseJson); // Log the updated form data or handle as needed
}

