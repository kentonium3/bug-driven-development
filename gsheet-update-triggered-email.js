/**
 * This script automates sending an email with data from a Google Sheet.
 * It is designed to be triggered by a new form submission.
 */
function sendDailyUpdate() {
  
  // 1. Get the active spreadsheet and the specific sheets with the data.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = spreadsheet.getSheetByName("Daily Tracker");
  const formSheet = spreadsheet.getSheetByName("Form Responses 1");
  
  // Ensure the sheets exist before proceeding
  if (!dataSheet) {
    Logger.log("Sheet 'Daily Tracker' not found. Please check the sheet name.");
    return;
  }
  if (!formSheet) {
    Logger.log("Sheet 'Form Responses 1' not found. Please check the sheet name.");
    return;
  }
  
  // 2. Get the most recent comment from the form submission sheet.
  // The comment is in column C of the last row.
  const lastRow = formSheet.getLastRow();
  const comment = formSheet.getRange(lastRow, 3).getValue();
  
  // 3. Get the data from the specified range (A1:D33) as it is displayed in the sheet.
  // Using getDisplayValues() instead of getValues() to preserve the cell's formatting.
  const dataRange = dataSheet.getRange("A1:D33");
  const data = dataRange.getDisplayValues();
  
  // 4. Create the email subject. The date has been removed to keep the thread consistent.
  const subject = "5:00a rise tracking update";
  
  // 5. Set the recipient email address for testing. Uncomment/comment below to use the test email address.
  // The recipient for testing purposes.
  // const recipientEmail = "kentgale@gmail.com"; 

  // The actual Google Group email address.
  const recipientEmail = "inyourface@googlegroups.com";
  
  // 6. Create the HTML table for the email body.
  // The comment from the form submission is now included at the top.
  let htmlBody = `<p><b>Comment:</b> ${comment}</p><h2>5:00a Rise Tracker</h2><table style="border-collapse: collapse;">`;
  
  // Loop through the data to create table rows and cells.
  data.forEach((row, rowIndex) => {
    // Alternate row colors for better readability
    const backgroundColor = rowIndex % 2 === 0 ? '#f0f0f0' : '#ffffff';
    htmlBody += `<tr style="background-color: ${backgroundColor};">`;
    row.forEach(cellData => {
      // Add a border to each cell for a clear grid look
      htmlBody += `<td style="border: 1px solid #cccccc; padding: 8px;">${cellData}</td>`;
    });
    htmlBody += '</tr>';
  });
  htmlBody += '</table>';
  
  // 7. Send the email. This logic now handles replying to an existing thread or creating a new one.
  try {
    // Retrieve the stored thread ID from script properties.
    const properties = PropertiesService.getScriptProperties();
    const threadId = properties.getProperty('riseTrackerThreadId');

    if (threadId) {
      // A thread ID exists, so try to reply to the existing thread.
      const thread = GmailApp.getThreadById(threadId);
      if (thread) {
        // Fix: Use the more reliable sendEmail method to ensure the message goes to the group.
        // Get the messages in the thread to build the 'references' header.
        const messages = thread.getMessages();
        const lastMessageId = messages[messages.length - 1].getId();
        const messageIds = messages.map(msg => msg.getId());

        GmailApp.sendEmail(recipientEmail, subject, "", {
          htmlBody: htmlBody,
          inReplyTo: lastMessageId,
          references: messageIds.join(" ")
        });
        
        Logger.log(`Replied to existing thread: ${threadId}`);
      } else {
        // The thread was not found (e.g., deleted), so we start a new one.
        Logger.log("Thread not found. Starting a new thread.");
        
        // Fix: Use createDraft and send() to reliably get the thread ID.
        const draft = GmailApp.createDraft(recipientEmail, subject, "", {htmlBody: htmlBody});
        const newThread = draft.send().getThread();
        properties.setProperty('riseTrackerThreadId', newThread.getId());
        Logger.log(`Created new thread: ${newThread.getId()}`);
      }
    } else {
      // No thread ID exists, so this is the first run.
      // Fix: Use createDraft and send() to reliably get the thread ID.
      const draft = GmailApp.createDraft(recipientEmail, subject, "", {htmlBody: htmlBody});
      const newThread = draft.send().getThread();
      properties.setProperty('riseTrackerThreadId', newThread.getId());
      Logger.log(`Created first thread: ${newThread.getId()}`);
    }
    
    Logger.log("Email update sent successfully!");
    
  } catch (e) {
    // Log any errors that occur during the email sending process
    Logger.log(`Error sending email: ${e.toString()}`);
  }
  
}
