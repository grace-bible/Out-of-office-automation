// Copyright 2020 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * Several constants are defined using JavaScript objects to store 
 * header names, reasons for time off, campus locations, email addresses, 
 * and various approval statuses. These constants are used 
 * throughout the script.
 */

const Header = {
  Timestamp: 'Timestamp',
  EmailAddress: 'Email Address',
  FullName: 'Name',
  Campus: 'Campus',
  StartDate: 'Start date',
  EndDate: 'End date',
  Reason: 'Reason',
  Description: 'Brief description',
  SuperAddress: 'Supervisor email',
  SupervisorApproval: 'My supervisor has already approved this request',
  HRApproval: 'HR approval',
  EventCreated: 'Calendar event status',
};

const Reason = {
  Personal: 'Personal',
  Professional: 'Professional',
  DWTL: 'DWTL',
};

const Campus = {
  AND: 'Anderson',
  SW: 'Southwood',
  CRK: 'Creekside',
  MT: 'Midtown',
  SYS: 'System',
};

/**
 * AND = grace-bible.org_qpq142rs3q8ujjovg633e5uhlg@group.calendar.google.com
 * CRK = grace-bible.org_nviveqkhsmbdqtiasj2nokl1pg@group.calendar.google.com
 * MT = c_uh7mlh14u22ui24sncqmrm3rrs@group.calendar.google.com
 * SW = grace-bible.org_4rtpbu8ot1fdsf5i7sl3dvkl2k@group.calendar.google.com
 */

const OOOcal = 'grace-bible.org_323330343338383235@resource.calendar.google.com'

const OOOemail = 'hr@grace-bible.org'

const SupervisorApproval = {
  Approved: 'Approved',
  NotApproved: 'Not approved',
};

const HRApproval = {
  Approved: 'Approved',
  NotApproved: 'Not approved',
};

const EventCreated = {
  NotCreated: 'Event not created',
  Created: 'Event created',
};

/**
* This function sets up custom menu items that appear when the Google Sheet is
* opened. The menu includes options for "Form setup," "Column setup," and 
* "Create calendar events." These menu items allow users to trigger specific 
* functions.
*/
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Approval functions')
      .addItem('Form setup', 'formSetup')
      .addItem('Column setup', 'columnSetup')
      .addItem('Create calendar events', 'eventSetup')
      .addToUi();
}

/**
 * This function is responsible for creating time-driven triggers. 
 * It uses `ScriptApp.newTrigger` to set up triggers that can execute functions
 * at specified times. In the code, there are two examples, one for triggering 
 * the `eventSetup` function every hour and another for triggering it every 
 * Monday.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
function createTimeDrivenTriggers() {
  // Trigger every 4 hours.
  ScriptApp.newTrigger('eventSetup')
      .timeBased()
      .everyHours(1)
      .create();
  // Trigger every Monday at 09:00.
  // ScriptApp.newTrigger('eventSetup')
  //     .timeBased()
  //     .onWeekDay(ScriptApp.WeekDay.MONDAY)
  //     .atHour(7)
  //     .create();
}

/**
   * This function sets up a Google Form for requesting time off. It creates 
   * form items such as text inputs, date pickers, and checkboxes. The form is 
   * linked to the Google Sheet to capture responses. If a form already exists,
   * it displays a message to unlink the existing form.
   */
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
  let form = FormApp.create('Out of office (OOO) request')
      .setCollectEmail(true)
      .setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId())
      .setLimitOneResponsePerUser(false);

  form.addTextItem().setTitle(Header.FullName).setRequired(true);
  form.addListItem().setTitle(Header.Campus).setChoiceValues(Object.values(Campus)).setRequired(false);
  form.addDateItem().setTitle(Header.StartDate).setRequired(true);
  form.addDateItem().setTitle(Header.EndDate).setRequired(true);
  form.addListItem().setTitle(Header.Reason).setChoiceValues(Object.values(Reason)).setRequired(true);
  form.addTextItem().setTitle(Header.Description).setRequired(false);
  form.addTextItem().setTitle(Header.SuperAddress).setRequired(true);
  
  let item = form.addCheckboxItem();

  item.setTitle(Header.SupervisorApproval)
  .setChoices([
    item.createChoice('Approved'),
  ])
  .showOtherOption(false)
  .setRequired(true);
}

/**
 * This function adds columns to the active Google Sheet. It uses the 
 * `appendColumn` function to add columns for "HR Approval" and "Calendar 
 * event status." Optional choices for these columns can be defined for data 
 * validation.
 */
function columnSetup() {
  let sheet = SpreadsheetApp.getActiveSheet();

  appendColumn(sheet, Header.HRApproval, Object.values(HRApproval));
  appendColumn(sheet, Header.EventCreated, Object.values(EventCreated));
}

/**
 * Appends a new column. It allows for optional choices to be defined for data validation if provided.
 * 
 *  @param {SpreadsheetApp.Sheet} sheet - tab in sheet.
 *  @param {string} headerName - name of column.
 *  @param {(string[] | null)} maybeChoices - optional drop down values for validation.
 */
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

/**
* This function checks the creation status of each entry in the Google Sheet 
* and, if a calendar event has not been created, it triggers the process 
* function gets the data range, validate headers, and process rows accordingly.
*/
function eventSetup() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let dataRange = sheet.getDataRange().getValues();
  let headers = dataRange.shift();

  validateSheetHeaders(headers, Header);

  let rows = dataRange
      .map((row, i) => asObject(headers, row, i))
      .filter(row => row[Header.EventCreated] != EventCreated.Created)
      .map(process)
      .map(row => writeRowToSheet(sheet, headers, row));
}

/**
* This function validates that the sheet headers match a predefined schema. 
* If a header is missing, it throws an error.
*/
function validateSheetHeaders(headers, schema) {
  for (let header of Object.values(schema)) {
    if (!headers.includes(header)) {
      throw `🦕 ⚠️ Header "${header}" not found in sheet: ${JSON.stringify(headers)}`;
    }
  }
}

/**
* This function converts rows (represented as arrays) into objects with named 
* properties. It uses the headers as keys to create objects from row data.
* 
* @param {string[]} headers - list of column names.
* @param {any[]} rowArray - values of a row as an array.
* @param {int} rowIndex - index of the row.
*/
function asObject(headers, rowArray, rowIndex) {
  return headers.reduce(
    (row, header, i) => {
      row[header] = rowArray[i];
      return row;
    }, {rowNumber: rowIndex + 1});
}

/**
* This function processes each row to check if it's "approved." If approved, 
* it creates a calendar event using the CalendarApp service and sends 
* confirmation emails. If not approved, it sends an email notification.
* 
* @param {Object} row - values in a row.
* @returns {Object} the row with a "notified status" column populated.
*/
function process(row) {
  let email = row[Header.EmailAddress];
  let name = row[Header.FullName];
  let supervisor = row[Header.SuperAddress];
  let replyall = `${supervisor}, ${OOOemail}`;
  let campus = row[Header.Campus];
  let guestEmails = `${OOOcal}, ${email}`;
  let today = new Date();
    let day = today.getDate();
    let month = today.getMonth() + 1;
    let year = today.getFullYear();
  let startDate = row[Header.StartDate];
    // Create a new variable to store the incremented date.
    let incrementStartDate = new Date(startDate);
    // Increment the `incrementStartDate` variable by 1 day.
    incrementStartDate.setDate(incrementStartDate.getDate() + 1);
  let endDate = row[Header.EndDate];
    // Create a new variable to store the incremented date.
    let incrementEndDate = new Date(endDate);
    // Increment the `incrementStartDate` variable by 1 day.
    incrementEndDate.setDate(incrementEndDate.getDate() + 2);
  let reason = row[Header.Reason];
  let description = row[Header.Description];
  let superApproval = (row[Header.SupervisorApproval]);
  let hrApproval = (row[Header.HRApproval]);
  let eventName = `${name} - ${reason}`;
  let eventDescription = `${superApproval} by ${supervisor}\n`
      + `Submitted on ${month}-${day}-${year}\n\n`
      + `${description}`;
  let message = `${name} has requested supervisor-approved ${reason} out-of-office from ${startDate.toDateString()} until ${endDate.toDateString()}\n\n`
      + `Reason: ${reason}\n\n`
      + `${description}\n\n`
      + `To discuss this request, "Reply All" to include Human Resources, the Supervisor, and the Employee on the thread.`;

  // /* Check if the user has a calendar. */
  // const calendar = CalendarApp.getCalendarById(email);
  // if (!calendar) {
  //     // The user does not have a calendar.
  //     Logger.log(`User does not have a calendar: ${email}`);
  //     // Display a message to the user.
  //     UiApp.alert(`User does not have a calendar. Please contact IT support before using this script.`);
  //     // Skip creating the calendar event.
  //     return row;
  // }

  /* Confirm that the supervisor approved. */
  if (superApproval == SupervisorApproval.NotApproved) {
    // If not approved, send an email.
    let subject = `[OOO] Your vacation time request failed, contact Josh McKenna`;
    MailApp.sendEmail(email, subject, message, {name: 'Out of office (OOO) automation', cc: replyall, bcc: 'joshmckenna@grace-bible.org'});
    row[Header.EventCreated] = EventCreated.Created;

    Logger.log(`Not approved, email sent, row=${JSON.stringify(row)}`);
  }

  else if (superApproval == SupervisorApproval.Approved) {
    // If approved, create a calendar event.
    CalendarApp.getCalendarById('ooo@grace-bible.org')
        .createAllDayEvent(
            eventName,
            incrementStartDate,
            incrementEndDate,
            {
              description: eventDescription,
              guests: guestEmails,
              sendInvites: true,
            })
        .setGuestsCanModify(true);
  
      // Send a confirmation email.
      let subject = `[OOO] New request for ${name} starting on ${startDate.toDateString()}`;
      MailApp.sendEmail(email, subject, message, {name: 'Out of office (OOO) automation', cc: replyall});
  
      row[Header.EventCreated] = EventCreated.Created;
  
      Logger.log(`Approved calendar event created, row=${JSON.stringify(row)}`);
  }

  else {
    row[Header.EventCreated] = EventCreated.NotCreated;

    Logger.log(`No action taken, row=${JSON.stringify(row)}`);
  }

  return row;
}

/**
* This function rewrites a row into the Google Sheet. It takes a sheet, 
* headers, and a row object as input and sets the values in the 
* appropriate rows and columns.
* 
* @param {SpreadsheetApp.Sheet} sheet - tab in sheet.
* @param {string[]} headers - list of column names.
* @param {Object} row - values in a row.
*/
function writeRowToSheet(sheet, headers, row) {
let rowArray = headers.map(header => row[header]);
let rowNumber = sheet.getFrozenRows() + row.rowNumber;
sheet.getRange(rowNumber, 1, 1, rowArray.length).setValues([rowArray]);
}