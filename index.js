// Function to serve the web form
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

// Function to process form data
function processFormDataForClass(data) {
  Logger.log("Processing form data...");
  Logger.log(JSON.stringify(data)); // Log the entire data object

  if (!data || !data.selectedClass || !data.subject || !data.students) {
    Logger.log("Invalid data received.");
    return;
  }

  // Get the folder named 'EXAM CENTER'
  var folder = DriveApp.getFoldersByName('EXAM CENTER').next();
  if (!folder) {
    Logger.log("Folder EXAM CENTER not found.");
    return;
  }

  // Define the spreadsheet files for Form 1 and Form 2
  var classMap = {
    "Form 1": folder.getFilesByName('Form 1').next(),
    "Form 2": folder.getFilesByName('Form 2').next()
  };

  var spreadsheet = classMap[data.selectedClass];
  if (!spreadsheet) {
    Logger.log("Spreadsheet for " + data.selectedClass + " not found.");
    return;
  }

  // Open the Broad Sheet in the selected spreadsheet
  var ss = SpreadsheetApp.open(spreadsheet);
  var broadSheet = ss.getSheetByName('Broad Sheet');
  if (!broadSheet) {
    Logger.log("Broad Sheet not found.");
    return;
  }

  // Map subjects to specific columns in the Broad Sheet
  var subjectColumnMap = {
    "Maths": 4,   // Column D
    "Eng": 5,     // Column E
    "Kisw": 6,    // Column F
    "Chem": 7,    // Column G
    "Phy": 8,     // Column H
    "Bio": 9,     // Column I
    "Hist": 10,   // Column J
    "Geo": 11,    // Column K
    "C.R.E": 12,  // Column L
    "I.R.E": 13,  // Column M
    "Agr": 14,    // Column N
    "B.std": 15   // Column O
  };

  var subjectColumn = subjectColumnMap[data.subject];
  if (!subjectColumn) {
    Logger.log("Invalid subject: " + data.subject);
    return;
  }

  // Iterate through each student and update their marks
  data.students.forEach(function(student) {
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

// Function to find the student's row based on admission number
function findStudentRow(admissionNumber, sheet) {
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1);  // Assuming admission numbers are in Column A
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
