function sendReminder() {
  var idSpreadSheet = 'your-spreadsheet-id'
  var spreadSheet = SpreadsheetApp.openById(idSpreadSheet)
  var sheet = spreadSheet.getSheetByName('your-sheet-name')
  var rangeValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  for (var i in rangeValues) {
    var employeeName = sheet.getRange(2 + Number(i), 1).getValue()
    var activity = sheet.getRange(2 + Number(i), 2).getValue()
    var phoneNumber = sheet.getRange(2 + Number(i), 4).getValue()
    var place = sheet.getRange(2 + Number(i), 5).getValue()

    var todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy"); // 21 April 2023

    var trainingDate = new Date(sheet.getRange(2 + Number(i), 3).getValue());
    var formattedTrainingDate = Utilities.formatDate(trainingDate, Session.getScriptTimeZone(), "dd MMMM yyyy"); // 23 April 2023

    var reminderDate = new Date(trainingDate - (0 * 24 * 60 * 60 * 1000)); // Change day reminder
    var formattedReminderDate = Utilities.formatDate(reminderDate, Session.getScriptTimeZone(), "dd MMMM yyyy"); // 21 April 2023

    const requestBody = {
      'target': String(phoneNumber),
      'message':
        '*_This is an auto generated message, please do not reply._*\r\n\r\n' +
        ''+
        'Dear ' + employeeName + ',\r\n' +
        'Ini adalah pengingat tentang pelatihan Anda pada :\r\n\r\n' +
        'Tanggal : ' + formattedTrainingDate + '\r\n' +
        'Subject : ' + activity + '\r\n' +
        'Lokasi : ' + place + '\r\n\r\n' +
        'Mohon untuk datang tepat waktu dan berpakaian dengan Sopan.'
    };

    var result = sheet.getRange(2 + Number(i), 6);
    var remark = sheet.getRange(2 + Number(i), 7);

    try {
      if (compareDates(new Date(todayDate), new Date(formattedReminderDate)) == 0 && (result.isBlank() || result.getValue() === 'FAILED')) {
        const headers = {
          'Authorization': 'YOUR-AUTHORIZATION-CODE',
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        };
        var url = 'https://api.fonnte.com/send'
        var params = {
          method: 'POST',
          payload: JSON.stringify(requestBody),
          headers: headers,
          contentType: "application/json"
        };
        var response = UrlFetchApp.fetch(url, params)

        switch (response.getResponseCode()) {
          case 200:
            result.setValue('SUCCESSFUL').setBackground('#b7e1cd');
            remark.setValue(JSON.parse(response).detail);
            break;
          case 303:
            result.setValue('FAILED').setBackground('#ea4335');
            remark.setValue('See Other');
            break;
          case 400:
            result.setValue('FAILED').setBackground('#ea4335');
            remark.setValue('Bad request');
            break;
          case 404:
            result.setValue('FAILED').setBackground('#ea4335');
            remark.setValue('Resource not found');
            break;
          case 500:
            result.setValue('FAILED').setBackground('#ea4335');
            remark.setValue('Internal server error');
            break;
        }
      }
    } catch (err) {
      result.setValue('FAILED').setBackground('#ea4335');
      remark.setValue(String(err).replace('\n', ''));
    }
  }

}


// Helper Function
function compareDates(date1, date2) {
  if (date1.getTime() === date2.getTime()) {
    return 0; // dates are equal
  } else if (date1.getTime() < date2.getTime()) {
    return -1; // date1 is before date2
  } else {
    return 1; // date1 is after date2
  }
}

function moveToCompleteSheet() {
  var idSpreadSheet = '1znvdH-4Lggn_z8VCHHDOno_VNgoiiqkQW-Gfb3dG1y4'
  var sourceSheetName = "cobaya"; // Replace with the name of the source sheet
  var destinationSheetName = "hasilcoba"; // Replace with the name of the destination sheet

  var spreadSheet = SpreadsheetApp.openById(idSpreadSheet)

  var sourceSheet = spreadSheet.getSheetByName(sourceSheetName);
  var destinationSheet = spreadSheet.getSheetByName(destinationSheetName);

  var sourceData = sourceSheet.getDataRange().getValues();

  for (var i = sourceData.length - 1; i > 0; i--) {
    var row = sourceData[i];
    var dueDate = row[2]; // due date in column C, because index start from 0, so C = 2

    if (new Date(dueDate) < new Date()) {
      destinationSheet.appendRow(row);
      sourceSheet.deleteRow(1 + Number(i));
    }
  }
}