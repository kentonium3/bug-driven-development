/**
 * This script automates sending an email with data from a Google Sheet.
 * It is designed to be triggered by a new form submission.
 */

// =====================================================================
// SCRIPT CONFIGURATION
// =====================================================================
const CONFIG = {
  // The name of the sheet containing the data to be copied and emailed.
  dataSheetName: "Daily Tracker",
  
  // The name of the sheet where the form responses are saved.
  formSheetName: "Form Responses 1",
  
  // The range of data to copy from the dataSheet (e.g., "A1:D33").
  dataRangeToCopy: "A1:D33",
  
  // The column index (1-based) of the comment in the form responses sheet.
  // Column C is the 3rd column.
  commentColumnIndex: 3,
  
  // The subject line for the email.
  emailSubject: "5:00a rise tracking update",
  
  // The email address or Google Group to send the update to.
  recipientEmail: "inyourface@googlegroups.com",
};

// =====================================================================
// SCRIPT LOGIC
// =====================================================================
function sendDailyUpdate() {
  
  // 1. Get the active spreadsheet and the specific sheets with the data.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = spreadsheet.getSheetByName(CONFIG.dataSheetName);
  const formSheet = spreadsheet.getSheetByName(CONFIG.formSheetName);
  
  // Ensure the sheets exist before proceeding
  if (!dataSheet) {
    Logger.log(`Sheet '${CONFIG.dataSheetName}' not found. Please check the sheet name in the CONFIG.`);
    return;
  }
  if (!formSheet) {
    Logger.log(`Sheet '${CONFIG.formSheetName}' not found. Please check the sheet name in the CONFIG.`);
    return;
  }
  
  // 2. Get the most recent comment from the form submission sheet.
  // The comment is in the specified column of the last row.
  const lastRow = formSheet.getLastRow();
  const comment = formSheet.getRange(lastRow, CONFIG.commentColumnIndex).getValue();
  
  // 3. Get the data from the specified range as it is displayed in the sheet.
  // Using getDisplayValues() instead of getValues() to preserve the cell's formatting.
  const dataRange = dataSheet.getRange(CONFIG.dataRangeToCopy);
  const data = dataRange.getDisplayValues();
  
  // 4. Set the email subject and recipient from the CONFIG.
  const subject = CONFIG.emailSubject;
  const recipientEmail = CONFIG.recipientEmail;
  
  // 5. Create the HTML table for the email body.
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
  
  // 6. Send the email. This logic now handles replying to an existing thread or creating a new one.
  try {
    const properties = PropertiesService.getScriptProperties();
    const currentThreadId = properties.getProperty('riseTrackerThreadId');
    let thread = null;
    
    // Check if a thread ID exists and try to find the thread using a robust search.
    if (currentThreadId) {
      // Use the search function to find the thread by ID, regardless of its labels.
      // This is a more reliable method than getThreadById().
      const threads = GmailApp.search("thread:" + currentThreadId);
      if (threads.length > 0) {
        thread = threads[0];
      }
    }

    if (thread) {
      // A thread ID exists and is valid, so reply to the existing thread.
      Logger.log(`Found existing thread. Replying to thread ID: ${currentThreadId}`);
      
      const messages = thread.getMessages();
      const lastMessageId = messages[messages.length - 1].getId();
      const messageIds = messages.map(msg => msg.getId());

      GmailApp.sendEmail(recipientEmail, subject, "", {
        htmlBody: htmlBody,
        inReplyTo: lastMessageId,
        references: messageIds.join(" ")
      });
      
    } else {
      // The thread was not found, so create a new one.
      if (currentThreadId) {
        properties.setProperty('lastKnownRiseTrackerThreadId', currentThreadId);
        Logger.log(`Previous thread ID (${currentThreadId}) stored for debugging. Starting a new thread.`);
      } else {
        Logger.log("No previous thread ID found. Starting a new thread.");
      }
      
      // Use createDraft and send() to reliably get the thread ID.
      const draft = GmailApp.createDraft(recipientEmail, subject, "", {htmlBody: htmlBody});
      const newThread = draft.send().getThread();
      properties.setProperty('riseTrackerThreadId', newThread.getId());
      Logger.log(`Created new thread with ID: ${newThread.getId()}`);
    }
    
    Logger.log("Email update sent successfully!");
    
  } catch (e) {
    // Log any errors that occur during the email sending process
    Logger.log(`Error sending email: ${e.toString()}`);
  }
  
}
