/**
 * Main entry point for the Strava-to-Sheets Sync.
 */

// Global constant for sheet headers if needed
const HEADERS = ['Activity ID', 'Name', 'Type', 'Distance (m)', 'Moving Time (s)', 'Start Date', 'Photo'];

/**
 * Creates a custom menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Strava Tools')
      .addItem('Sync Now (Batch Mode)', 'runBatchSync') // Renamed to reflect batch nature
      .addItem('Test Sync (First 5)', 'testSync')
      .addToUi();
}

/**
 * Initiates the sync process. Managed via Triggers.
 * Fetches activities in batches to respect rate limits.
 */
function runBatchSync() {
  const props = PropertiesService.getScriptProperties();
  
  // Get current page cursor, default to 1
  let currentPage = parseInt(props.getProperty('SYNC_CURSOR_PAGE')) || 1;
  
  console.log(`Starting batch sync for page ${currentPage}`);
  
  // Fetch activities for the current page
  // Using BATCH_SIZE from Config.gs
  const activities = fetchActivities(currentPage, BATCH_SIZE);
  
  if (!activities || activities.length === 0) {
    console.log('No more activities found. Sync complete.');
    // Clean up
    props.deleteProperty('SYNC_CURSOR_PAGE');
    deleteTrigger('runBatchSync');
    SpreadsheetApp.getUi().alert('Sync Complete!');
    return;
  }
  
  // Process this batch
  const newRows = processActivities(activities);
  
  if (newRows > 0) {
    console.log(`Processed page ${currentPage}. Moving to next page.`);
    props.setProperty('SYNC_CURSOR_PAGE', String(currentPage + 1));
    
    // Schedule next batch if we processed a full batch (implying there might be more)
    // Or just always schedule if we got activities? safest is if we got results, try next page.
    createTrigger('runBatchSync', 20); // Schedule for 20 minutes later
    SpreadsheetApp.getUi().alert(`Batch ${currentPage} complete. Next batch scheduled in 20 mins.`);
  } else {
    // If we fetched activities but filtered them all out (duplicates), we still need to check next page?
    // Yes, because older activities might be in the sheet but we are paginating generally from newest to oldest 
    // (Strava default). 
    // IF default is newest->oldest, and we encounter duplicates, we might be caught up.
    // BUT user might have gaps. Safer to continue paginating until empty result if doing a full backfill.
    // However, for incremental updates, we usually stop when we find the first duplicate.
    // Let's assume FULL BACKFILL mode logic: continue even if duplicates found in this batch.
    
    console.log(`Page ${currentPage} processed (all duplicates?). Moving to next page.`);
    props.setProperty('SYNC_CURSOR_PAGE', String(currentPage + 1));
    createTrigger('runBatchSync', 20);
  }
}

/**
 * Test function to sync only the first 5 activities.
 * Does not set triggers.
 */
function testSync() {
  console.log('Starting Test Sync...');
  const activities = fetchActivities(1, TEST_PAGE_SIZE);
  
  if (activities && activities.length > 0) {
    const count = processActivities(activities);
    SpreadsheetApp.getUi().alert(`Test Sync Complete. Processed ${count} new activities.`);
  } else {
    SpreadsheetApp.getUi().alert('Test Sync: No activities found.');
  }
}

/**
 * Helper to process a list of activities and write valid ones to sheet.
 * @param {Array} activities 
 * @return {number} Count of new rows added.
 */
function processActivities(activities) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
    existingIds = values.flat().map(String);
  }

  const newRows = [];
  
  for (const activity of activities) {
    const activityId = String(activity.id);
    
    if (existingIds.includes(activityId)) {
      continue;
    }
    
    let photoUrl = fetchActivityPhotos(activity.id);
    let photoCellFormula = photoUrl ? `=HYPERLINK("${photoUrl}", "View Photo")` : 'No Photo';

    newRows.push([
      activityId,
      activity.name,
      activity.type,
      activity.distance,
      activity.moving_time,
      activity.start_date_local,
      photoCellFormula
    ]);
  }

  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
    console.log(`Added ${newRows.length} new activities.`);
  }
  
  return newRows.length;
}

/**
 * Creates a time-based trigger.
 * Deletes any existing trigger for the function first to avoid duplicates.
 * @param {string} funcName 
 * @param {number} minutesAfter 
 */
function createTrigger(funcName, minutesAfter) {
  deleteTrigger(funcName); // Ensure only one trigger exists
  ScriptApp.newTrigger(funcName)
      .timeBased()
      .after(minutesAfter * 60 * 1000)
      .create();
}

/**
 * Deletes all triggers for a specific function.
 * @param {string} funcName 
 */
function deleteTrigger(funcName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}
