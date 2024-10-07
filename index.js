// This function serves the web form to the user
function doGet() {
    return HtmlService.createHtmlOutputFromFile('index');
  }
  
  // This function processes the data submitted from the form
  function processFormData(data) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Broad Sheet");
  
    // Log if the sheet was found or not
    if (!sheet) {
      Logger.log("Error: Broad Sheet not found.");
      return;
    }
    
    Logger.log("Processing form data...");
  
    var subject = data.subject;
    var students = data.students;
  
    // Log the subject and students array being processed
    Logger.log("Subject: " + subject);
    Logger.log("Students: " + JSON.stringify(students));
  
    students.forEach(function(student) {
      var admissionNumber = student.admissionNumber;
      var mark = student.mark;
  
      // Log each student's admission number and mark
      Logger.log("Admission Number: " + admissionNumber + ", Mark: " + mark);
  
      if (!admissionNumber || !mark) {
        Logger.log("Skipping student with missing data.");
        return;
      }
  
      // Find the student row and subject column
      var studentRow = findStudentRow(admissionNumber, sheet);
      var subjectColumn = getSubjectColumn(subject);
  
      // Log the result of finding the student row and subject column
      Logger.log("Student Row: " + studentRow + ", Subject Column: " + subjectColumn);
  
      // If row and column are valid, add the mark to the sheet
      if (studentRow !== -1 && subjectColumn !== -1) {
        sheet.getRange(studentRow, subjectColumn).setValue(mark);
        Logger.log("Mark added for student " + admissionNumber + ": " + mark);
      } else {
        Logger.log("Unable to add mark for student " + admissionNumber);
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
  