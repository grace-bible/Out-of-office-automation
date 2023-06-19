# Time-off Request App for Google Sheets

This is a time-off request application built for Google Sheets. It allows employees to submit vacation requests through a form and provides functionality for managing and approving those requests. The app utilizes Google Sheets, Google Forms, CalendarApp, and MailApp services in Google Apps Script.

## Setup

1. Open your Google Sheets document.
2. Set your calendar ID:<br>
   Replace ```CalendarApp.getCalendarById("")``` with the desired calendar ID in the ```process``` function. This ID specifies the calendar where approved time-off events will be added. For example, ``` CalendarApp.getCalendarById("your-calendar-id")```
3. Customize the header and approval options:<br>
   Modify the ```Header``` and ```Approval``` objects to match your desired header names and approval options. You can change the field names and values according to your requirements.
4. Enable necessary services:<br>
   From the Google Sheets menu, go to "Extensions" > "Apps Script" to open the Apps Script editor. If prompted, click on "Enable" to enable Google Apps Script.
5. Set up the form:<br>
   In the Apps Script editor, run the ```formSetup```function. This will create a form titled "Request time off" and link it to your Google Sheets document. The form will have fields for name, start date, end date, and     
   approval status.

## Usage

1. Submitting time-off requests:<br>
  Employees can access the form by clicking on "Form" > "Go to live form" from the Google Sheets menu. They can then fill out the form with their name, start date, and end date. Once submitted, the request will be added to     the Google Sheets document.
2. Approving or rejecting requests:<br>
  From the Google Sheets menu, click on "🏝 Vacation" > "Notify employees" to process the time-off requests. Based on the approval status, the corresponding actions will be taken:<br>
      * If the request is not approved, an email will be sent to the employee notifying them of the rejection.
      * If the request is approved, a calendar event will be created for the requested time-off, and a confirmation email will be sent to the employee.

## Notes

* Make sure to set the necessary calendar IDs in the process function and formSetup function.
* The app assumes that the Google Sheets document has the required headers as specified in the ```Header``` object.
* The app uses the MailApp service to send emails. Ensure that the email sending capabilities are enabled in your Google Workspace account.
* Customize the email content and subject in the ```process``` function according to your needs.
* Use the ```onOpen``` function to add a custom menu option to your Google Sheets document for easy access to the app.

## Disclaimer

This application was created to demonstrate the basic functionality of a time-off request system using Google Sheets, Google Forms, and Google Apps Script. It is recommended to review and modify the code as per your specific requirements and ensure compliance with your organization's policies and guidelines.

For more information and support, please refer to the Google Apps Script documentation and the Google Sheets Help Center.
