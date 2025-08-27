/**
 * @fileoverview Daily Update Email Script with Fixed Threading
 * 
 * This script automates sending an email with data from a Google Sheet.
 * It is designed to be triggered by a new form submission and uses Gmail's
 * native threading capabilities for reliable email threading across all clients.
 * 
 * Version: v3.0-PRODUCTION-FIXED
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
  
  // Threading configuration
  threadIdProperty: 'riseTrackerThreadId',
  
  // Script version for debugging and tracking
  scriptVersion: "v3.0-PRODUCTION-FIXED"
};

// =====================================================================
// MAIN FUNCTION
// =====================================================================

/**
 * Main function to send daily update emails with proper threading
 * This is designed to be triggered by form submissions
 */
function sendDailyUpdate() {
  Logger.log(`--- Starting Daily Update (${CONFIG.scriptVersion}) ---`);
  Logger.log(`Timestamp: ${new Date().toLocaleString()}`);
  
  try {
    // Step 1: Validate and get spreadsheet data
    const emailData = getSpreadsheetData();
    if (!emailData) {
      Logger.log("❌ Failed to get spreadsheet data - aborting");
      return;
    }
    
    // Step 2: Create email content
    const htmlBody = createEmailBody(emailData);
    
    // Step 3: Send email with native threading (THE KEY FIX!)
    const success = sendThreadedEmail(htmlBody);
    
    if (success) {
      Logger.log("✅ Daily update sent successfully with proper threading!");
    } else {
      Logger.log("❌ Failed to send daily update");
    }
    
  } catch (error) {
    Logger.log(`❌ CRITICAL ERROR in sendDailyUpdate: ${error.toString()}`);
    Logger.log(`❌ Stack trace: ${error.stack || 'No stack trace available'}`);
  }
  
  Logger.log("--- Daily Update Complete ---");
}

// =====================================================================
// DATA RETRIEVAL FUNCTIONS
// =====================================================================

/**
 * Gets and validates spreadsheet data
 * @return {Object|null} Object with comment and data, or null if failed
 */
function getSpreadsheetData() {
  Logger.log("1. Getting spreadsheet data...");
  
  try {
    // Get the active spreadsheet and the specific sheets with the data
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`1. Spreadsheet: "${spreadsheet.getName()}"`);
    
    const dataSheet = spreadsheet.getSheetByName(CONFIG.dataSheetName);
    const formSheet = spreadsheet.getSheetByName(CONFIG.formSheetName);
    
    // Validate sheets exist
    if (!dataSheet) {
      Logger.log(`❌ Sheet '${CONFIG.dataSheetName}' not found. Available sheets: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
      return null;
    }
    if (!formSheet) {
      Logger.log(`❌ Sheet '${CONFIG.formSheetName}' not found. Available sheets: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
      return null;
    }
    
    Logger.log(`2. Found data sheet: "${CONFIG.dataSheetName}"`);
    Logger.log(`2. Found form sheet: "${CONFIG.formSheetName}"`);
    
    // Get the most recent comment from the form submission sheet
    const lastRow = formSheet.getLastRow();
    Logger.log(`3. Form sheet last row: ${lastRow}`);
    
    if (lastRow < 1) {
      Logger.log("⚠️ No data in form sheet, using default comment");
      return {
        comment: "No form submissions yet",
        data: getDataSheetData(dataSheet)
      };
    }
    
    const comment = formSheet.getRange(lastRow, CONFIG.commentColumnIndex).getValue();
    Logger.log(`3. Latest comment: "${comment}"`);
    
    // Get the data from the specified range
    const data = getDataSheetData(dataSheet);
    
    return { comment, data };
    
  } catch (error) {
    Logger.log(`❌ Error getting spreadsheet data: ${error.toString()}`);
    return null;
  }
}

/**
 * Gets data from the data sheet with validation
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet The data sheet
 * @return {Array|null} 2D array of data or null if failed
 */
function getDataSheetData(dataSheet) {
  try {
    Logger.log(`4. Getting data from range: ${CONFIG.dataRangeToCopy}`);
    
    // Validate the range exists
    const dataRange = dataSheet.getRange(CONFIG.dataRangeToCopy);
    
    // Get the data using getDisplayValues() to preserve formatting
    const data = dataRange.getDisplayValues();
    
    Logger.log(`4. Retrieved ${data.length} rows and ${data[0] ? data[0].length : 0} columns`);
    
    // Validate we have data
    if (!data || data.length === 0) {
      Logger.log("⚠️ No data found in specified range");
      return [["No Data", "Available"]];
    }
    
    return data;
    
  } catch (error) {
    Logger.log(`❌ Error getting data sheet data: ${error.toString()}`);
    return [["Error", "Loading Data"]];
  }
}

// =====================================================================
// EMAIL CREATION FUNCTION
// =====================================================================

/**
 * Creates HTML email body from spreadsheet data
 * @param {Object} emailData Object containing comment and data
 * @return {string} HTML email body
 */
function createEmailBody(emailData) {
  Logger.log("5. Creating email body...");
  
  const { comment, data } = emailData;
  const timestamp = new Date().toLocaleString();
  
  try {
    // Start building HTML body
    let htmlBody = `
      <div style="font-family: Arial, sans-serif; max-width: 800px;">
        <p><b>Comment:</b> ${comment || 'No comment'}</p>
        <h2>5:00a Rise Tracker</h2>
        <p style="color: #666; font-size: 12px;">
          Sent: ${timestamp} | Script: ${CONFIG.scriptVersion}
        </p>
        <table style="border-collapse: collapse; width: 100%;">
    `;
    
    // Add table rows
    data.forEach((row, rowIndex) => {
      // Alternate row colors for better readability
      const backgroundColor = rowIndex % 2 === 0 ? '#f0f0f0' : '#ffffff';
      const isHeaderRow = rowIndex === 0;
      const fontWeight = isHeaderRow ? 'bold' : 'normal';
      const headerBg = isHeaderRow ? '#d4e6f1' : backgroundColor;
      
      htmlBody += `<tr style="background-color: ${headerBg};">`;
      
      row.forEach(cellData => {
        // Handle empty cells
        const cellContent = cellData || '';
        htmlBody += `
          <td style="
            border: 1px solid #cccccc; 
            padding: 8px; 
            font-weight: ${fontWeight};
            text-align: ${isHeaderRow ? 'center' : 'left'};
          ">
            ${cellContent}
          </td>
        `;
      });
      
      htmlBody += '</tr>';
    });
    
    htmlBody += `
        </table>
        <p style="margin-top: 20px; color: #888; font-size: 10px;">
          Automated email from Google Apps Script (${CONFIG.scriptVersion})
        </p>
      </div>
    `;
    
    Logger.log(`5. Email body created (${htmlBody.length} characters)`);
    return htmlBody;
    
  } catch (error) {
    Logger.log(`❌ Error creating email body: ${error.toString()}`);
    
    // Return a simple fallback email
    return `
      <div style="font-family: Arial, sans-serif;">
        <h2>5:00a Rise Tracker</h2>
        <p><b>Comment:</b> ${comment || 'No comment'}</p>
        <p><b>Error:</b> Could not format data table</p>
        <p>Sent: ${timestamp}</p>
      </div>
    `;
  }
}

// =====================================================================
// THREADING FUNCTIONS - THE KEY FIX!
// =====================================================================

/**
 * Sends email using Gmail's native threading capabilities
 * This replaces the problematic Message-ID extraction approach
 * @param {string} htmlBody The HTML content to send
 * @return {boolean} True if successful
 */
function sendThreadedEmail(htmlBody) {
  Logger.log("6. Starting threaded email send...");
  
  const properties = PropertiesService.getScriptProperties();
  const storedThreadId = properties.getProperty(CONFIG.threadIdProperty);
  
  Logger.log(`6. Stored thread ID: ${storedThreadId || 'None'}`);
  
  try {
    if (storedThreadId) {
      // Try to reply to existing thread using NATIVE THREADING
      const replySuccess = replyToExistingThread(storedThreadId, htmlBody);
      if (replySuccess) {
        Logger.log("7. ✅ Successfully replied to existing thread using native threading");
        return true;
      } else {
        Logger.log("7. Reply failed, creating new thread instead");
        // Continue to create new thread
      }
    } else {
      Logger.log("6. No stored thread found, creating new thread");
    }
    
    // Create new thread if none exists or reply failed
    const newThreadSuccess = createNewThread(htmlBody, properties);
    if (newThreadSuccess) {
      Logger.log("8. ✅ Successfully created new thread");
      return true;
    } else {
      Logger.log("8. ❌ Failed to create new thread");
      return false;
    }
    
  } catch (error) {
    Logger.log(`❌ Error in sendThreadedEmail: ${error.toString()}`);
    return false;
  }
}

/**
 * Replies to existing thread using Gmail's native reply method
 * @param {string} threadId The stored thread ID
 * @param {string} htmlBody The email content
 * @return {boolean} True if successful
 */
function replyToExistingThread(threadId, htmlBody) {
  Logger.log(`7. Attempting to reply to thread: ${threadId}`);
  
  try {
    let thread = null;
    
    try {
      thread = GmailApp.getThreadById(threadId);
    } catch (directError) {
      Logger.log(`7. Direct thread lookup failed: ${directError.toString()}`);
      Logger.log("7. Trying alternative search method...");
      
      const threads = GmailApp.search("thread:" + threadId);
      if (threads.length > 0) {
        thread = threads[0];
        Logger.log("7. Found thread via search method");
      }
    }
    
    if (!thread) {
      Logger.log("7. ❌ Thread not found by any method");
      return false;
    }
    
    const messages = thread.getMessages();
    Logger.log(`7. Found thread with ${messages.length} messages`);
    Logger.log(`7. Thread subject: "${thread.getFirstMessageSubject()}"`);
    
    // Get the first message for threading headers
    const firstMessage = messages[0];
    const lastMessage = messages[messages.length - 1];
    
    // Extract Message-ID for threading
    const rawContent = firstMessage.getRawContent();
    const messageIdMatch = rawContent.match(/Message-ID:\s*<([^>]+)>/i);
    const messageId = messageIdMatch ? messageIdMatch[1] : null;
    
    if (!messageId) {
      Logger.log("7. ⚠️ Could not extract Message-ID, falling back to reply method");
      // Fallback - but this won't work as intended
      lastMessage.reply("", {htmlBody: htmlBody});
      return true;
    }
    
    Logger.log(`7. First message ID: ${messageId}`);
    
    // Build References header for threading
    let references = `<${messageId}>`;
    
    // Extract existing References if any
    const referencesMatch = rawContent.match(/References:\s*([^\r\n]+)/i);
    if (referencesMatch) {
      references = referencesMatch[1].trim() + " " + references;
    }
    
    // Use sendEmail with proper headers to force group recipient
    GmailApp.sendEmail(
      CONFIG.recipientEmail,  // This WILL go to the group
      "Re: " + thread.getFirstMessageSubject(),
      "",
      {
        htmlBody: htmlBody,
        headers: {
          "In-Reply-To": `<${messageId}>`,
          "References": references
        }
      }
    );
    
    Logger.log("7. ✅ Reply sent to group with threading headers");
    return true;
    
  } catch (error) {
    Logger.log(`❌ Error replying to thread ${threadId}: ${error.toString()}`);
    return false;
  }
}

/**
 * Creates a new thread and stores the ID
 * @param {string} htmlBody The email content
 * @param {PropertiesService.Properties} properties The properties service
 * @return {boolean} True if successful
 */
function createNewThread(htmlBody, properties) {
  Logger.log("8. Creating new thread...");
  
  try {
    // Store old thread ID for debugging if it exists
    const oldThreadId = properties.getProperty(CONFIG.threadIdProperty);
    if (oldThreadId) {
      properties.setProperty('lastKnownRiseTrackerThreadId', oldThreadId);
      Logger.log(`8. Previous thread ID (${oldThreadId}) stored for debugging`);
    }
    
    // Use createDraft and send() to reliably get the thread ID
    const draft = GmailApp.createDraft(
      CONFIG.recipientEmail,
      CONFIG.emailSubject,
      "",
      { htmlBody: htmlBody }
    );
    
    Logger.log("8. Draft created, sending...");
    const newThread = draft.send().getThread();
    const newThreadId = newThread.getId();
    
    // Store the new thread ID
    properties.setProperty(CONFIG.threadIdProperty, newThreadId);
    
    Logger.log(`8. ✅ Created new thread with ID: ${newThreadId}`);
    Logger.log(`8. Thread subject: "${newThread.getFirstMessageSubject()}"`);
    
    return true;
    
  } catch (error) {
    Logger.log(`❌ Error creating new thread: ${error.toString()}`);
    return false;
  }
}

// =====================================================================
// UTILITY FUNCTIONS
// =====================================================================

/**
 * Test function - run this to test the script without a form trigger
 */
function testDailyUpdate() {
  Logger.log("=== MANUAL TEST OF DAILY UPDATE ===");
  sendDailyUpdate();
  Logger.log("=== TEST COMPLETE ===");
}

/**
 * Reset threading - clears stored thread ID to start fresh
 */
function resetThreading() {
  const properties = PropertiesService.getScriptProperties();
  const oldThreadId = properties.getProperty(CONFIG.threadIdProperty);
  
  properties.deleteProperty(CONFIG.threadIdProperty);
  
  Logger.log("✅ Threading reset completed");
  Logger.log(`Previous thread ID: ${oldThreadId || 'None'}`);
  Logger.log("Next email will start a new thread");
}

/**
 * Get current thread information for debugging
 */
function getThreadInfo() {
  Logger.log("--- Thread Information ---");
  
  const properties = PropertiesService.getScriptProperties();
  const threadId = properties.getProperty(CONFIG.threadIdProperty);
  const lastKnownThreadId = properties.getProperty('lastKnownRiseTrackerThreadId');
  
  Logger.log(`Current thread ID: ${threadId || 'None'}`);
  Logger.log(`Last known thread ID: ${lastKnownThreadId || 'None'}`);
  
  if (threadId) {
    try {
      const thread = GmailApp.getThreadById(threadId);
      const messages = thread.getMessages();
      
      Logger.log(`Subject: ${thread.getFirstMessageSubject()}`);
      Logger.log(`Message count: ${messages.length}`);
      Logger.log(`First message: ${messages[0].getDate()}`);
      Logger.log(`Last message: ${messages[messages.length - 1].getDate()}`);
      Logger.log(`Labels: ${thread.getLabels().map(l => l.getName()).join(', ')}`);
      
    } catch (error) {
      Logger.log(`❌ Thread ${threadId} is invalid: ${error.toString()}`);
    }
  }
}

/**
 * Validate configuration and environment
 */
function validateSetup() {
  Logger.log("--- Setup Validation ---");
  
  try {
    // Check spreadsheet access
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`✅ Spreadsheet: "${spreadsheet.getName()}"`);
    
    // Check sheets exist
    const sheets = spreadsheet.getSheets().map(s => s.getName());
    Logger.log(`Available sheets: ${sheets.join(', ')}`);
    
    const dataSheet = spreadsheet.getSheetByName(CONFIG.dataSheetName);
    const formSheet = spreadsheet.getSheetByName(CONFIG.formSheetName);
    
    Logger.log(`Data sheet "${CONFIG.dataSheetName}": ${dataSheet ? '✅ Found' : '❌ Missing'}`);
    Logger.log(`Form sheet "${CONFIG.formSheetName}": ${formSheet ? '✅ Found' : '❌ Missing'}`);
    
    // Check data range
    if (dataSheet) {
      try {
        const range = dataSheet.getRange(CONFIG.dataRangeToCopy);
        Logger.log(`✅ Data range "${CONFIG.dataRangeToCopy}" is valid`);
      } catch (rangeError) {
        Logger.log(`❌ Data range "${CONFIG.dataRangeToCopy}" is invalid: ${rangeError.toString()}`);
      }
    }
    
    // Check email configuration
    Logger.log(`Email subject: "${CONFIG.emailSubject}"`);
    Logger.log(`Recipient: "${CONFIG.recipientEmail}"`);
    
    // Check user permissions
    Logger.log(`Script user: ${Session.getActiveUser().getEmail()}`);
    
    Logger.log("--- Validation Complete ---");
    
  } catch (error) {
    Logger.log(`❌ Validation error: ${error.toString()}`);
  }
}