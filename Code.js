// The following script works along with a Google Form that students submit when they request to meet with a counselor.
// There is a trigger that is set to send the emails out when a form is submitted to this sheet.
// Point of contact is: Alvaro Gomez, Special Campuses Academic Technology Coach, 210-363-1577
// If you want to watch a Screencastify talking through this script follow this link: https://app.screencastify.com/v3/watch/ue27Z389KgH2KnefoM0r

function onFormSubmit(e) {
  try {
    // Checks if 'e' is undefined
    if (!e || !e.values) {
      Logger.log("Form submission event or values are undefined.");
      return;
    }

    // Log details about the event
    Logger.log("Form Submission Event:", e);

    // Get the responses submitted in the form
    let responses = e.values;

    // Log the form responses
    Logger.log("Form Responses: ", responses);

    // Specify the column index of the counselor selection question
    let counselorColumnIndex = 7;

    // Map counselor names to their corresponding emails
    let counselorEmails = {
      '(A-CAR) Appleby': 'janelle.appleby@nisd.net',
      '(CAS-FL) Hewgley': 'shanna.hewgley@nisd.net',
      '(FO-H) Ramos': 'elizabeth.ramos@nisd.net',
      '(I-MA) Matta': 'orlando.matta@nisd.net',
      '(MC-P) Clarke': 'darrell.clarke@nisd.net',
      '(Q-SO) Pearson': 'samantha.pearson@nisd.net',
      '(SP-Z) Holder': 'kristina.holder@nisd.net',
      '(ASTA-All Ag Students) Zablocki': 'pamela.zablocki@nisd.net',
      '(Head Counselor) Schmidt': 'kimberly.schmidt@nisd.net'
    };

    // Used for testing
    // let counselorEmails = {
    //   '(A-CAR) Appleby': 'alvaro.gomez@nisd.net',
    //   '(CAS-FL) Hewgley': 'alvaro.gomez@nisd.net',
    //   '(FO-H) Ramos': 'alvaro.gomez@nisd.net',
    //   '(I-MA) Matta': 'orlando.matta@nisd.net',
    //   '(MC-P) Clarke': 'alvaro.gomez@nisd.net',
    //   '(Q-SO) Pearson': 'alvaro.gomez@nisd.net',
    //   '(SP-Z) Holder': 'alvaro.gomez@nisd.net',
    //   '(ASTA-All Ag Students) Zablocki': 'alvaro.gomez@nisd.net',
    //   '(Head Counselor) Schmidt': 'alvaro.gomez@nisd.net'
    // };

    // Get the selected counselor's name from the response
    let counselorName = responses[counselorColumnIndex - 1];

    // Get the counselor's email based on the selected name
    let counselorEmail = counselorEmails[counselorName];

    // Get other relevant information from the form responses
    let studentEmail = responses[1]
    let lastName = responses[2];
    let firstName = responses[3];
    let studentId = responses[4];
    let grade = responses[5];
    let reason = responses[7];
    let urgent = responses[8];
    let dropCourse = responses[9];
    let addCourse = responses[10];
    let description = responses[11];
    let period = responses[12];

    // Compose the email message with conditionally included parts
    let subject = 'REQUEST TO SEE COUNSELOR';
    let body = `${firstName} requested to meet with you.\n\nFollowing are the details that they provided:\n${lastName}, ${firstName} ${studentId}, Grade: ${grade}, Email: ${studentEmail}\n`;

    // Customized parts of the body of the email that are appended to the body above. The rest of the body will be different depending
    // on the reason they requested to see their counselor.
    if (reason === 'Personal Matter (Please indicate URGENT in the notes)') {
      body += `They would like to see you for a personal matter.\nThey indicated the following regarding urgency: ${urgent}\nThe best period to meet is: ${period}`;
    }
    if (reason === 'College/Scholarship Information') {
      body += `They would like to see you for college/scholarship information.\nThe best period to meet is: ${period}`;
    }
    if (reason === 'Credit Information') {
      body += `They would like to see you regarding credit information.\nThe best period to meet is: ${period}`;
    }
    if (reason === 'Schedule Change') {
      body += `They would like to meet for a schedule change.\nCourse requesting to drop: ${dropCourse}\nCourse requesting to add: ${addCourse}\nThe best period to meet is: ${period}`;
    }
    if (reason === 'Other') {
      body += `They indicated that they have an "Other" request that isn't listed in the dropdown and provided this brief description: ${description}\nThe best period to meet is: ${period}`;
    }

    // Log details about the email
    Logger.log("Email Details (Counselor):", {
      to: counselorEmail,
      subject: subject,
      body: body
    });

    // Send the email to the selected counselor
    MailApp.sendEmail({
      to: counselorEmail,
      cc: "ydette.cothran@nisd.net",
      subject: subject,
      body: body
    });

    // A line break for better readability when checking the exectution log.
    Logger.log("----------------------------------------------------------")

    // Compose the email message for the student
    let studentSubject = 'Confirmation: Request to See Counselor';
    let studentBody = `Dear ${firstName},\n\nYour request to see a counselor was submited and they should be in contact with you soon. Reason you requested to see them: ${reason}\n`;

    // Log details about the email to the student
    Logger.log("Email Details (Student):", {
      to: studentEmail,
      subject: studentSubject,
      body: studentBody
    });

  } catch (error) {
    Logger.log("Error:", error);

    // Send an error notification email
    sendErrorNotificationEmail(error);
  }
}

function sendErrorNotificationEmail(error) {
  try {
    // Specify the recipient email address for error notifications
    let adminEmail = 'orlando.matta@nisd.net';

    // Compose the email message for error notification
    let errorSubject = 'Error Notification: Script Execution Issue';
    let errorBody = `An error occurred in the script:\n\n${error}`;

    // Log details about the error notification email
    Logger.log("Error Notification Email Details:", {
      to: adminEmail,
      subject: errorSubject,
      body: errorBody
    });

    // Send the error notification email
    MailApp.sendEmail({
      to: adminEmail,
      subject: errorSubject,
      body: errorBody
    });
  } catch (e) {
    // Log any additional error that might occur during the error notification process
    Logger.log("Error in sending error notification email:", e);
  }
}
