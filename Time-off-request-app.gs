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

const Header = {
  Timestamp: "Timestamp",
  EmailAddress: "Email Address",
  FullName: "Name",
  Campus: "Campus",
  StartDate: "Start date",
  EndDate: "End date",
  Reason: "Reason",
  Description: "Brief description",
  SuperAddress: "Supervisor email",
  SupervisorApproval: "My supervisor has already approved this request",
  HRApproval: "HR approval",
  EventCreated: "Calendar event status",
};

const Reason = {
  Personal: "Personal",
  Professional: "Professional",
  DWTL: "DWTL",
};

const Campus = {
  AND: "Anderson",
  SW: "Southwood",
  CRK: "Creekside",
  MT: "Midtown",
  SYS: "System",
};

/**
 * AND = grace-bible.org_qpq142rs3q8ujjovg633e5uhlg@group.calendar.google.com
 * CRK = grace-bible.org_nviveqkhsmbdqtiasj2nokl1pg@group.calendar.google.com
 * MT = c_uh7mlh14u22ui24sncqmrm3rrs@group.calendar.google.com
 * SW = grace-bible.org_4rtpbu8ot1fdsf5i7sl3dvkl2k@group.calendar.google.com
 */

const OOOcal =
  "grace-bible.org_323330343338383235@resource.calendar.google.com";

const OOOemail = "janineford@grace-bible.org, madelineechols@grace-bible.org";

const SupervisorApproval = {
  Approved: "Approved",
  NotApproved: "Not approved",
};

const HRApproval = {
  Approved: "Approved",
  NotApproved: "Not approved",
};

const EventCreated = {
  NotCreated: "Event not created",
  Created: "Event created",
  Canceled: "Request canceled",
};

/**
 * Add custom menu items when opening the sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Approval functions")
    .addItem("Form setup", "formSetup")
    .addItem("Column setup", "columnSetup")
    .addItem("Create calendar events", "eventSetup")
    .addToUi();
}

/**
 * Creates time-driven trigger(s).
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
function createTimeDrivenTriggers() {
  // Trigger every 1 hour.
  ScriptApp.newTrigger("eventSetup").timeBased().everyHours(1).create();
}

/**
 * Set up the "Out of office (OOO) request" Google Form, and link the form's trigger to
 * optionally send an email to an additional address (like a manager).
 */
function formSetup() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  if (sheet.getFormUrl()) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(
      "ℹ️ A Form already exists",
      "Unlink the form and try again.\n\n" +
        "From the top menu:\n" +
        'Click "Form" > "Unlink form"',
      ui.ButtonSet.OK
    );
    return;
  }

  // Create the form.
  let form = FormApp.create("Out of office (OOO) request")
    .setCollectEmail(true)
    .setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId())
    .setLimitOneResponsePerUser(false);

  form.addTextItem().setTitle(Header.FullName).setRequired(true);
  form
    .addListItem()
    .setTitle(Header.Campus)
    .setChoiceValues(Object.values(Campus))
    .setRequired(false);
  form.addDateItem().setTitle(Header.StartDate).setRequired(true);
  form.addDateItem().setTitle(Header.EndDate).setRequired(true);
  form
    .addListItem()
    .setTitle(Header.Reason)
    .setChoiceValues(Object.values(Reason))
    .setRequired(true);
  form.addTextItem().setTitle(Header.Description).setRequired(false);
  form.addTextItem().setTitle(Header.SuperAddress).setRequired(true);

  let item = form.addCheckboxItem();

  item
    .setTitle(Header.SupervisorApproval)
    .setChoices([item.createChoice("Approved")])
    .showOtherOption(false)
    .setRequired(true);
}

/**
 * Creates an "HR Approved" and "Calendar event status" column
 */
function columnSetup() {
  let sheet = SpreadsheetApp.getActiveSheet();

  appendColumn(sheet, Header.HRApproval, Object.values(HRApproval));
  appendColumn(sheet, Header.EventCreated, Object.values(EventCreated));
}

/**
 * Appends a new column.
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

    range
      .offset(sheet.getFrozenRows(), 0, sheet.getMaxRows())
      .setDataValidation(rule);
  }
}

/**
 * Checks the creation status of each entry and, if not created,
 * creates a new calendar item accordingly.
 */
function eventSetup() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let dataRange = sheet.getDataRange().getValues();
  let headers = dataRange.shift();

  validateSheetHeaders(headers, Header);

  let rows = dataRange
    .map((row, i) => asObject(headers, row, i))
    .filter(
      (row) =>
        row[Header.EventCreated] != EventCreated.Created &&
        row[Header.EventCreated] != EventCreated.Canceled &&
        row[Header.HRApproval] != HRApproval.NotApproved
    )
    .map(process)
    .map((row) => writeRowToSheet(sheet, headers, row));
}

/**
 * Validate that the sheet headers match a schema.
 */
function validateSheetHeaders(headers, schema) {
  for (let header of Object.values(schema)) {
    if (!headers.includes(header)) {
      throw `🦕 ⚠️ Header "${header}" not found in sheet: ${JSON.stringify(
        headers
      )}`;
    }
  }
}

/**
 * Validate that supervisor email(s) match a schema.
 */
function validateEmails(input) {
  let regex =
    /[-A-Za-z0-9!#$%&'*+\/=?^_`{|}~]+(?:\.[-A-Za-z0-9!#$%&'*+\/=?^_`{|}~]+)*@(?:[A-Za-z0-9](?:[-A-Za-z0-9]*[A-Za-z0-9])?\.)+[A-Za-z0-9](?:[-A-Za-z0-9]*[A-Za-z0-9])?,\s+[-A-Za-z0-9!#$%&'*+\/=?^_`{|}~]+(?:\.[-A-Za-z0-9!#$%&'*+\/=?^_`{|}~]+)*@(?:[A-Za-z0-9](?:[-A-Za-z0-9]*[A-Za-z0-9])?\.)+[A-Za-z0-9](?:[-A-Za-z0-9]*[A-Za-z0-9])?/i;
  return regex.test(input);
}

/**
 * Convert the row arrays into objects.
 * Start with an empty object, then create a new field
 * for each header name using the corresponding row value.
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
    },
    { rowNumber: rowIndex + 1 }
  );
}

/**
 * Checks if a row is marked as "approved". If approved a calendar
 * event is created for the user. If not approved, an email
 * notification is sent.
 *
 * @param {Object} row - values in a row.
 * @returns {Object} the row with a "notified status" column populated.
 */
function process(row) {
  let email = row[Header.EmailAddress].split(" ")[0];
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
  // Increment the `incrementStartDate` variable by 0 days.
  incrementStartDate.setDate(incrementStartDate.getDate() + 0);
  let endDate = row[Header.EndDate];
  // Create a new variable to store the incremented date.
  let incrementEndDate = new Date(endDate);
  // Increment the `incrementEndDate` variable by 1 day.
  incrementEndDate.setDate(incrementEndDate.getDate() + 1);
  let reason = row[Header.Reason];
  let description = row[Header.Description];
  let superApproval = row[Header.SupervisorApproval];
  let hrApproval = row[Header.HRApproval];
  let eventName = `${name} - ${reason}`;
  let eventDescription =
    `${superApproval} by ${supervisor}\n` +
    `Submitted on ${month}-${day}-${year}\n\n` +
    `${description}`;
  let message =
    `${name} has requested supervisor-approved ${reason} out-of-office from ${startDate.toDateString()} until ${endDate.toDateString()}\n\n` +
    `Reason: ${reason}\n\n` +
    `${description}\n\n` +
    `To discuss this request, "Reply All" to include Human Resources, the Supervisor, and the Employee on the thread.`;

  /* Confirm that the supervisor approved. */
  if (superApproval == SupervisorApproval.NotApproved) {
    // If not approved, send an email and cancel the request.
    let subject = `[OOO] Request ERROR: Supervisor approval required`;
    MailApp.sendEmail(
      email,
      subject,
      "Please secure your supervisor's approval before resubmitting your request.",
      {
        name: "Out of office (OOO) automation ERROR",
        cc: replyall,
        // bcc: "joshmckenna+error@grace-bible.org",
      }
    );
    row[Header.EventCreated] = EventCreated.Canceled;

    Logger.log(
      `ERROR: Approval denied, email sent, row=${JSON.stringify(row)}`
    );
    SpreadsheetApp.getUi().alert(
      `ERROR: ${name} must secure Superivsor approval before requesting time OOO. See row row=${JSON.stringify(
        row
      )} for the canceled request. ${email} was notified to resubmit the request after securing Supervisor approval.`
    );
  } else if (HRApproval == HRApproval.NotApproved) {
    // If not denied, send an email and cancel the request.
    let subject = `[OOO] Request ERROR: HR denied your request`;
    MailApp.sendEmail(
      email,
      subject,
      "Please contact HR immediately for more details. Do NOT resubmit your request without contacting HR.",
      {
        name: "Out of office (OOO) automation ERROR",
        cc: replyall,
        // bcc: "joshmckenna+error@grace-bible.org",
      }
    );
    row[Header.EventCreated] = EventCreated.Canceled;

    Logger.log(
      `ERROR: Approval denied, email sent, row=${JSON.stringify(row)}`
    );
    SpreadsheetApp.getUi().alert(
      `ERROR: HR has denied this request for ${name}. See row row=${JSON.stringify(
        row
      )} for the canceled request. ${email} was notified to contact HR for more information.`
    );
  } else if (incrementEndDate.getTime() < incrementStartDate.getTime()) {
    // If startDate after endDate, send an email and cancel the request.
    let subject = `[OOO] Request ERROR: Only God transcends time`;
    MailApp.sendEmail(
      email,
      subject,
      "Please check the dates you entered and resubmit your request with a valid start date that precedes the end date.",
      {
        name: "Out of office (OOO) automation ERROR",
        cc: replyall,
        // bcc: "joshmckenna+error@grace-bible.org",
      }
    );
    row[Header.EventCreated] = EventCreated.Canceled;

    Logger.log(
      `ERROR: Requested dates invalid, email sent, row=${JSON.stringify(row)}`
    );
    SpreadsheetApp.getUi().alert(
      `ERROR: ${name} has requested to time travel without a proper permit. See row row=${JSON.stringify(
        row
      )} for the canceled request. ${email} was notified to resubmit the request with valid dates.`
    );
  } else if (
    superApproval == SupervisorApproval.Approved &&
    hrApproval != HRApproval.NotApproved &&
    incrementEndDate.getTime() > incrementStartDate.getTime()
  ) {
    // If approved, create a calendar event.
    CalendarApp.getCalendarById("ooo@grace-bible.org")
      .createAllDayEvent(eventName, incrementStartDate, incrementEndDate, {
        description: eventDescription,
        guests: guestEmails,
        sendInvites: true,
      })
      .setGuestsCanModify(true);

    // Send a confirmation email.
    let subject = `[OOO] New request for ${name} starting on ${startDate.toDateString()}`;
    MailApp.sendEmail(email, subject, message, {
      name: "Out of office (OOO) automation",
      cc: replyall,
    });

    row[Header.EventCreated] = EventCreated.Created;

    Logger.log(
      `Approved calendar event for ${name} created, row=${JSON.stringify(row)}`
    );
  } else {
    // For any other error, send an email and cancel the request.
    let subject = `[OOO] ERROR: Unexpected error occurred`;
    MailApp.sendEmail(
      "joshmckenna+error@grace-bible.org",
      subject,
      "Please contact joshmckenna+error@grace-bible.org for more details. Please pay close attention to your typing and submitted details when you resubmit your request.",
      { name: "Out of office (OOO) automation ERROR", cc: email, bcc: replyall }
    );
    row[Header.EventCreated] = EventCreated.Canceled;

    Logger.log(`No action taken, row=${JSON.stringify(row)}`);
    SpreadsheetApp.getUi().alert(
      `ERROR: Unexpexted fatal error at row row=${JSON.stringify(
        row
      )} and joshmckenna+error@grace-bible.org has been notified to investigate.`
    );
  }

  return row;
}

/**
 * Rewrites a row into the sheet.
 *
 * @param {SpreadsheetApp.Sheet} sheet - tab in sheet.
 * @param {string[]} headers - list of column names.
 * @param {Object} row - values in a row.
 */
function writeRowToSheet(sheet, headers, row) {
  let rowArray = headers.map((header) => row[header]);
  let rowNumber = sheet.getFrozenRows() + row.rowNumber;
  sheet.getRange(rowNumber, 1, 1, rowArray.length).setValues([rowArray]);
}
