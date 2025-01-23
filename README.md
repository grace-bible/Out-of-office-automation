<!-- @format -->

# Time-off Request App for Google Sheets

This is a time-off request application built for Google Sheets. It allows employees to submit vacation requests through a form and provides functionality for managing and approving those requests. The app utilizes Google Sheets, Google Forms, CalendarApp, and MailApp services in Google Apps Script.

1. **Form setup**: Creates an intake form that you can share with users to submit time off requests. Their responses will be automatically added to the sheet.
2. **Column setup**: Creates the columns needed to manage approvals in the sheet.
3. **Create**: Creates a calendar reservation request on the specified out-of-office (OOO) calendar, and sends email notifications to both staff and HR with each created calendar request.

## Customization

The expected workflow is that employees request time off by submitting calendar invites to a shared "Out of office" calendar managed by HR, rather than a spreadsheet-based approval process.

Prior to running the "Form setup" and "Column setup" functions of the [Time-off-request-app.gs](Time-off-request-app.gs) script, you can customize certain elements for your own operating environment in the global constant declarations preceding the various functions. This script is configured to receive an employee's email address, and to use the form submitted email to dynamically `getCalendarByID(email)`.

1. Customize the header and approval options:<br>
   Modify the `Header` and `Reason` and `Campus` and `SupervisorApproval` and `HRApproval` objects to match your desired header names and reasons for requesting time off. You can change the field names and values according to your requirements.
2. Set your calendar ID:<br>
   Replace the `OOOcal` constant with the desired calendar ID. This ID specifies the calendar where approved time-off events will be added.
3. Enable necessary services:<br>
   From the Google Sheets menu, go to "Extensions" > "Apps Script" to open the Apps Script editor. If prompted, click on "Enable" to enable Google Apps Script.

## How to use

1. Create a new Google Sheet.
2. From the menu, click <kbd>Extensions</kbd> > <kbd>Apps Script</kbd>.
3. Copy the contents of [Time-off-request-app.gs](Time-off-request-app.gs) and paste it over the boilerplate `Code.gs` in the Apps Script Editor
4. Re-open the Google Sheet you created and wait for a few seconds for a custom menu called <kbd>Approval functions</kbd> to appear at the top of the sheet.
5. Click <kbd>Approval functions</kbd> > <kbd>Form setup</kbd>, then wait for the dialog to indicate completion of the script.
6. Next, use the <kbd>Column setup</kbd> function to create columns for tracking the status of requests.
7. Click <kbd>Tools</kbd> > <kbd>Manage form</kbd> > <kbd>Send form</kbd> to open the form sharing dialogue to share it with users who need to submit time off requests. As users submit requests, they will be automatically added to the sheet.

> [!NOTE]
> The current version of this script only runs by default when triggered manually by clicking "Approval functions" > "Create calendar events". A time-based trigger can be added from the Apps Script Triggers editor by triggering the `create()` function.

### Benefits

- This app is a simple and easy way to manage vacation requests.
- It is free to use.
- It is customizable to meet the specific needs of your organization.
- It can be used to track the status of requests.
- It can be used to send email notifications to users.

### Limitations

- This app is not a comprehensive vacation management system.
- It does not support features such as accrual tracking or blackout dates.
- It is not designed to be used by a large number of users.

## Conclusion

This Apps Script project is a simple and effective way to manage vacation requests. It is a good option for _small_ organizations that are looking for a free and easy-to-use solution.

## Usage

1. Submitting time-off requests:<br>
   Employees can access the form by clicking on <kbd>Form</kbd> > <kbd>Go to live form</kbd> from the Google Sheets menu. They can then fill out the form with their name, start date, and end date. Once submitted, the request will be added to the Google Sheets document.
2. Approving or rejecting requests:<br>
   From the Google Sheets menu, click on <kbd>Approval functions</kbd> > <kbd>Create calendar events</kbd> to process the time-off requests and place them on the specificed `OOOcal`. Based on the success, the corresponding actions will be taken:<br>
   _ If the calendar event has not been created or fails, an email will be sent to the employee notifying them of the rejection.
   _ If the calendar event is created a calendar event will be created for the requested time-off, and a confirmation email will be sent to the employee.

## Notes

- The app assumes that the Google Sheets document has the required headers as specified in the `Header` object.
- The app uses the MailApp service to send emails. Ensure that the email sending capabilities are enabled in your Google Workspace account.
- Customize the email content and subject in the `process(row)` function according to your needs.
- Use the `onOpen` function to add a custom menu option to your Google Sheets document for easy access to the app.

## Disclaimer

This application was created to demonstrate the basic functionality of a time-off request system using Google Sheets, Google Forms, and Google Apps Script. It is recommended to review and modify the code as per your specific requirements and ensure compliance with your organization's policies and guidelines.

For more information and support, please refer to the [Google Apps Script documentation](https://developers.google.com/apps-script/reference) and the Google Sheets Help Center.
