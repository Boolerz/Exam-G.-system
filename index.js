// This function serves the web form to the user
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

// This function processes the data submitted from the form
function processFormData(data) {
  Logger.log("Processing form data...");
  Logger.log(JSON.stringify(data)); // Log the entire data object
  
  if (!data || !data.subject || !data.students) {
    Logger.log("Invalid data received.");
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var broadSheet = ss.getSheetByName('Broad Sheet');

  if (!broadSheet) {
    Logger.log("Broad Sheet not found.");
    return;
  }

  var subject = data.subject;
  var students = data.students;

  Logger.log("Subject: " + subject);
  Logger.log("Students: " + JSON.stringify(students));

  // Map subjects to specific columns in the Broad Sheet
  var subjectColumnMap = {
    "Maths": 4,  // Example: Column D
    "Eng": 5,    // Example: Column E
    "Kisw": 6,   // Example: Column F
    "Chem": 7,   // Example: Column G
    "Phy": 8,    // Example: Column H
    "Bio": 9     // Example: Column I
  };

  var subjectColumn = subjectColumnMap[subject];
  if (!subjectColumn) {
    Logger.log("Invalid subject: " + subject);
    return;
  }

  // Iterate through each student and update their marks
  students.forEach(function(student) {
    var admissionNumber = student.admissionNumber;
    var mark = student.mark;

    Logger.log("Processing student " + admissionNumber + " with mark " + mark);

    // Find the row of the student using their admission number
    var studentRow = findStudentRow(admissionNumber, broadSheet);
    if (studentRow !== -1) {
      broadSheet.getRange(studentRow, subjectColumn).setValue(mark);
      Logger.log("Mark updated for student " + admissionNumber + " in row " + studentRow);
    } else {
      Logger.log("Student with admission number " + admissionNumber + " not found.");
    }
  });
}

// Function to find the student's row in the sheet based on their admission number
function findStudentRow(admissionNumber, sheet) {
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1);  // Assume admission numbers are in Column A
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim() === admissionNumber.trim()) {
      Logger.log("Found student with admission number " + admissionNumber + " at row " + (i + 2));
      return i + 2;  // Return the row number (index + 2 because of the 0-based index and header row)
    }
  }

  Logger.log("Student with admission number " + admissionNumber + " not found.");
  return -1;
}

// Function to map the subject name to the correct column in the sheet
function getSubjectColumn(subject) {
  var subjectMap = {
    "Maths": 4,  // Column D
    "Eng": 5,    // Column E
    "Kisw": 6,   // Column F
    "Chem": 7,   // Column G
    "Phy": 8,    // Column H
    "Bio": 9     // Column I
  };

  var column = subjectMap[subject] || -1;
  if (column === -1) {
    Logger.log("Subject " + subject + " not found in subject map.");
  } else {
    Logger.log("Subject " + subject + " maps to column " + column);
  }

  return column;
}
