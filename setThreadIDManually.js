/**
 * A one-time helper script to manually set the thread ID.
 * This should only be run once to fix a missing ID.
 */
function setThreadIdManually() {
  const properties = PropertiesService.getScriptProperties();
  
  // PASTE THE THREAD ID FROM YOUR GMAIL URL HERE
  const threadId = "FMfcgzQcpKlPvJmlwZBmdSwkwSWczRJJ"; 
  
  properties.setProperty('riseTrackerThreadId', threadId);
  Logger.log("Thread ID has been manually set to: " + threadId);
}