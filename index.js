function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function submitMarks(data) {
    var subject = data.subject;
    var classSelected = data.class;
    var marks = data.marks;

    // Spreadsheet IDs for Form 1 and Form 2
    var form1SpreadsheetId = '15NKbo7KZdhEK9f2MUGXZ2Raj2QeUUSJ9HIUjnuzn6W0';
    var form2SpreadsheetId = '1vDTRpIJwVym8B76DR3KSv0fO3EEylfe1oPYp-EepF3E';

    // Select the correct spreadsheet based on class
    var spreadsheetId = classSelected === 'Form 1' ? form1SpreadsheetId : form2SpreadsheetId;
    var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Broad Sheet');

    // Find the column for the selected subject
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var subjectColumn = headers.indexOf(subject) + 1;  // Add 1 because index starts at 0

    // Check if the subject column exists
    if (subjectColumn === 0) {
        throw new Error('Subject "' + subject + '" not found in headers.');
    }

    // Loop through the students and fill in the marks
    marks.forEach(function(student) {
        var admissionNumber = student.admissionNumber;
        var mark = student.mark;

        // Find the row for the student's admission number
        var admissionNumbers = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        var rowIndex = admissionNumbers.findIndex(function(row) {
            return row[0] == admissionNumber;
        });

        // If student is found, add the mark to the correct subject column
        if (rowIndex > -1) {
            sheet.getRange(rowIndex + 2, subjectColumn).setValue(mark);
        }
    });
}

