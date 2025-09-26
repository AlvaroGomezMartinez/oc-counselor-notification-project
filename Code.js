// The following script works along with a Google Form that students submit when they request to meet with a counselor.
// There is a trigger that is set to send the emails out when a form is submitted to this sheet.
// Point of contact is: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-397-9408, 210-363-1577
// If you want to watch a Screencastify talking through this script follow this link: https://app.screencastify.com/v3/watch/ue27Z389KgH2KnefoM0r

function onFormSubmit(e) {
  try {
    // Checks if values array exists
    if (!e.values) {
      Logger.log("Form values are undefined.");
      return;
    }

    // Get the responses submitted in the form
    let responses = e.values;

    // Log the form responses
    Logger.log("Form Responses: " + JSON.stringify(responses));

    // Specify the column index of the counselor selection question
    let counselorColumnIndex = 7;

    // This is a map of counselor names to their corresponding emails. The keys have to match exactly how the options are in the Google Form.
    // let counselorEmails = {
    //   '(A - C) Appleby': 'janelle.appleby@nisd.net',
    //   '(D - Ha) Hewgley': 'shanna.hewgley@nisd.net',
    //   '(He - Mi) Ramos': 'elizabeth.ramos@nisd.net',
    //   '(Mo - R) Clarke': 'orlando.matta@nisd.net',
    //   '(S - Z) Pearson': 'samantha.pearson@nisd.net',
    //   '(ASTA-All Ag Students) Zablocki': 'pamela.zablocki@nisd.net',
    //   '(Head Counselor) Matta': 'Orlando.Matta@nisd.net'
    // };

    // Used for testing. Remember to comment out the CC in the MailApp.sendEmail function below.
    let counselorEmails = {
      '(A - C) Appleby': 'alvaro.gomez@nisd.net',
      '(D - Ha) Hewgley': 'alvaro.gomez@nisd.net',
      '(He - Mi) Ramos': 'alvaro.gomez@nisd.net',
      '(Mo - R) Clarke': 'alvaro.gomez@nisd.net',
      '(S - Z) Pearson': 'alvaro.gomez@nisd.net',
      '(ASTA-All Ag Students) Zablocki': 'alvaro.gomez@nisd.net',
      '(Head Counselor) Matta': 'alvaro.gomez@nisd.net'
    };

    // Get the selected counselor's name from the response
    let counselorName = responses[counselorColumnIndex - 1].trim();

    // Get the counselor's email based on the selected name
    let counselorEmail = counselorEmails[counselorName];
    // Checks if the counselor's email was found
    if (!counselorEmail) {
      throw new Error(`Could not find an email for counselor. Please check for a typo or mismatch between the Google Form dropdown option selected: "${responses[6]}" and the counselorEmails map value: "${counselorName}" in the script. Both have to match exactly (e.g. spaces, punctuation, etc.) for the email to send.`);
    }

    // Get other relevant information from the form responses
    let studentEmail = responses[1]; // Col B
    let lastName = responses[2];     // Col C
    let firstName = responses[3];    // Col D
    let studentId = responses[4];    // Col E
    let grade = responses[5];        // Col F

    // Compose the email message with conditionally included parts
    let subject = 'REQUEST TO SEE COUNSELOR';
    let body = `${firstName} requested to meet with you.\n\nFollowing are the details that they provided:\n${lastName}, ${firstName} ${studentId}, Grade: ${grade}, Email: ${studentEmail}\n`;

    // Send the email to the selected counselor
    MailApp.sendEmail({
      to: counselorEmail,
      // cc: "ydette.cothran@nisd.net",
      subject: subject,
      body: body
    });

  } catch (error) {
    // Log detailed error information
    Logger.log("Error occurred:");
    Logger.log("Error message: " + error.message);
    Logger.log("Error stack: " + error.stack);
    Logger.log("Error toString: " + error.toString());

    // Send an error notification email
    sendErrorNotificationEmail(error);
  }
}

function sendErrorNotificationEmail(error) {
  try {
    // Specify the recipient email address for error notifications
    let adminEmail = 'orlando.matta@nisd.net';

    // Compose the email message for error notification with detailed error info
    let errorSubject = 'Error Notification: Script Execution Issue';
    let errorBody = `An error occurred in the script:\n\n` +
                   `Error Message: ${error.message}\n` +
                   `Error Stack: ${error.stack}\n` +
                   `Full Error: ${error.toString()}\n\n` +
                   `Time of Error: ${new Date().toLocaleString()}`;

    // Log details about the error notification email
    Logger.log("Error Notification Email Details:" + JSON.stringify({
      to: adminEmail,
      subject: errorSubject,
      body: errorBody
    }));

    // Send the error notification email
    MailApp.sendEmail({
      to: adminEmail,
      subject: errorSubject,
      body: errorBody
    });
  } catch (e) {
    // Log any additional error that might occur during the error notification process
    Logger.log("Error in sending error notification email:" + e);
  }
}