```js
const Header = {
  Timestamp: 'Time-stamp',
  EmailAddress: 'Email',
  Name: 'Name',
  StartDate: 'Start date',
  EndDate: 'End date',
  Approval: 'Approval',
};

const Approval = {
  Approved: 'Approved',
  NotApproved: 'Not approved',
};


function formSetup() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  if (sheet.getFormUrl()) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(
      'ℹ️ A Form already exists',
      'Unlink the form and try again.\n\n' +
      'From the top menu:\n' +
      'Click "Form" > "Unlink form"',
      ui.ButtonSet.OK
    );
    return;
  }

  // Create the form.
  let form = FormApp.create('Request time off')
      .setCollectEmail(true)
      .setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId())
      .setLimitOneResponsePerUser(false);

  form.addTextItem().setTitle(Header.Name).setRequired(true);
  form.addDateItem().setTitle(Header.StartDate).setRequired(true);
  form.addDateItem().setTitle(Header.EndDate).setRequired(true);
  form.addListItem().setTitle(Header.Reason).setChoiceValues(Object.values(Reason)).setRequired(true);
  columnSetup()
}

function columnSetup() {
  let sheet = SpreadsheetApp.getActiveSheet();

  appendColumn(sheet, Header.Approval, Object.values(Approval));
}

function appendColumn(sheet, headerName, maybeChoices) {
  let range = sheet.getRange(1, sheet.getLastColumn() + 1);

  // Create the header header name.
  range.setValue(headerName);

  // If we pass choices to the function, create validation rules.
  if (maybeChoices) {
    let rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(maybeChoices)
        .build();

    range.offset(sheet.getFrozenRows(), 0, sheet.getMaxRows())
        .setDataValidation(rule);
  }
}

function asObject(headers, rowArray, rowIndex) {
  return headers.reduce(
    (row, header, i) => {
      row[header] = rowArray[i];
      return row;
    }, {rowNumber: rowIndex + 1});
}

function data(){
  let sheet = SpreadsheetApp.getActiveSheet();
  let dataRange = sheet.getDataRange().getValues();
  let headers = dataRange.shift();

  dataRange
      .map((row, i) => asObject(headers, row, i))
      .map(process)
      
}


function process(row) {
  let name = row[Header.Name];
  let email = row[Header.EmailAddress];
  let startDate = row[Header.StartDate];
  let endDate = row[Header.EndDate];
  let approval = row[Header.Approval];
  let message = `Your vacation time request from `
      + `${startDate.toDateString()} to `
      + `${endDate.toDateString()}: ${approval}`;

  if (approval == Approval.NotApproved) {
    // If not approved, send an email.

    let subject = 'Your vacation time request was NOT approved';
    MailApp.sendEmail(email, subject, message);

    Logger.log(`Not approved, email sent, row=${JSON.stringify(row)}`);
  }

  else if (approval == Approval.Approved) {
    // If approved, create a calendar event.

    CalendarApp.getCalendarById("c_phbb5q70vllatui1at9kfrs3os@group.calendar.google.com")
      .createAllDayEvent(
          name + ' on Vacation',
          startDate,
          endDate,
          {
            description: message,
          });

    // Send a confirmation email.
    let subject = 'Confirmed, your vacation time request has been approved!';
    MailApp.sendEmail(email, subject, message);

    Logger.log(`Approved, calendar event created, row=${JSON.stringify(row)}`);
  }

  return row;
}
```
