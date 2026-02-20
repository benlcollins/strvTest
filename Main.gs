/**
 * Main entry point for the Strava-to-Sheets Sync.
 */

// Expanded Headers
const HEADERS = [
  'Activity ID', 
  'Name', 
  'Type', 
  'Distance (mi)', 
  'Elevation (ft)', 
  'Moving Time', 
  'Elapsed Time', 
  'Start Date', 
  'Max Speed (mph)', 
  'Avg Speed (mph)', 
  'Calories',
  'Description',
  'Temp (C)',
  'Achievement Count', 
  'Kudos Count', 
  'Comment Count', 
  'Bike', 
  'Shoes', 
  'Results',
  'Photos',
  'Strava Link'
];

/**
 * Creates a custom menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Strava Tools')
      .addItem('Sync Now (Fast Parallel)', 'runBatchSync') 
      .addItem('Test Sync (First 5)', 'testSync')
      .addSeparator()
      .addItem('Debug: Recent Activity JSON', 'debugActivityJSON')
      .addToUi();
}

/**
 * DEBUG TOOL: Fetches the most recent activity and logs filtered JSON.
 */
function debugActivityJSON() {
  console.log("Fetching detailed activity for debug...");
  
  const activities = fetchActivities(1, 1); 
  if (!activities || activities.length === 0) {
    alertUser("No activities found.");
    return;
  }
  
  const activityId = activities[0].id;
  const detail = fetchActivityDetails(activityId); 
  
  if (!detail) {
    alertUser("Failed to fetch details.");
    return;
  }
  
  delete detail.segment_efforts;
  delete detail.map;
  delete detail.splits_metric;
  delete detail.splits_standard;
  delete detail.laps;
  delete detail.best_efforts;
  
  let debugSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Debug');
  if (!debugSheet) {
    debugSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Debug');
  }
  debugSheet.clear();
  const jsonStr = JSON.stringify(detail, null, 2);
  
  if (jsonStr.length > 50000) {
     debugSheet.getRange(1, 1).setValue(jsonStr.substring(0, 50000) + "... [TRUNCATED]");
  } else {
     debugSheet.getRange(1, 1).setValue(jsonStr);
  }
  
  SpreadsheetApp.setActiveSheet(debugSheet);
  alertUser(`Debug JSON (filtered) written to 'Debug' sheet.`);
}

/**
 * Initiates the sync process. Managed via Triggers.
 * Handles rate limits by pausing execution without losing place.
 */
/**
 * Initiates the sync process. Managed via Triggers.
 * Optimized to skip through existing activities efficiently.
 */
function runBatchSync() {
  const props = PropertiesService.getScriptProperties();
  let currentPage = parseInt(props.getProperty('SYNC_CURSOR_PAGE')) || 1;
  const MAX_PAGES_PER_EXEC = 8; 
  
  console.log(`Starting batch sync starting at page ${currentPage}`);
  
  try {
    let pagesProcessedInThisRun = 0;
    let totalNewActivitiesFound = 0;

    while (pagesProcessedInThisRun < MAX_PAGES_PER_EXEC) {
      console.log(`Fetching page ${currentPage} (Batch Size: ${BATCH_SIZE})...`);
      const activities = fetchActivities(currentPage, BATCH_SIZE);
      
      if (!activities || activities.length === 0) {
        console.log('No more activities found from API. Sync complete.');
        props.deleteProperty('SYNC_CURSOR_PAGE');
        deleteTrigger('runBatchSync');
        alertUser('Sync Complete! No more history found on Strava.');
        return;
      }
      
      const result = processActivities(activities);
      const count = result.processed;
      totalNewActivitiesFound += count;
      pagesProcessedInThisRun++;
      
      if (!result.hasMoreOnPage) {
        currentPage++; 
      }

      if (count > 0) {
        console.log(`Found ${count} new activities. Stopping loop to process next batch.`);
        break;
      } else {
        console.log(`Page ${currentPage - (result.hasMoreOnPage ? 0 : 1)} skipped (Duplicates).`);
      }
    }
    
    props.setProperty('SYNC_CURSOR_PAGE', String(currentPage));
    const waitTime = (totalNewActivitiesFound === 0) ? 1 : 15;
    createTrigger('runBatchSync', waitTime); 
    
    const msg = (totalNewActivitiesFound === 0) 
      ? `Skipped ${pagesProcessedInThisRun} pages of history. Continuing in 1 min...`
      : `Added ${totalNewActivitiesFound} new activities. Next batch in 15 mins.`;
    
    alertUser(msg);

  } catch (e) {
    if (e.message === 'RATE_LIMIT_EXCEEDED') {
      console.warn('Rate Limit Hit. Sync paused.');
      deleteTrigger('runBatchSync');
      alertUser('Rate Limit Hit. Sync paused for 24 hours. Cursor saved.');
    } else {
      console.error('Error in runBatchSync: ' + e);
      deleteTrigger('runBatchSync');
      alertUser('Error during sync: ' + e);
    }
  }
}

function testSync() {
  console.log('Starting Test Sync...');
  try {
    const activities = fetchActivities(1, 5); 
    
    if (activities && activities.length > 0) {
      const result = processActivities(activities);
      alertUser(`Test Sync Complete. Processed ${result.processed} new activities.`);
    } else {
      alertUser('Test Sync: No activities found.');
    }
  } catch (e) {
    alertUser('Test Sync Error: ' + e);
  }
}

/**
 * Helper to safely alert the user.
 * swallows alerts if running in background (TimeBased trigger).
 */
function alertUser(message) {
  try {
    // Check if we can access UI (we can't in time-based triggers)
    // There isn't a direct "Am I in trigger?" check, but getUi() throws if not available?
    // actually, DocumentApp.getUi() / SpreadsheetApp.getUi() implies bound script active user.
    // Better way: Check if we have an active user context? 
    // Usually catching the exception is the way.
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // We are likely in a background trigger
    console.log(`[Background Alert]: ${message}`);
  }
}

/**
 * Helper to process a list of activities and write valid ones to sheet.
 * Uses DYNAMIC COLUMN MAPPING for robustness.
 */
function processActivities(activities) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return 0;
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Create a map of Header Name -> Column Index (0-based)
  const headerMap = {};
  rawHeaders.forEach((h, i) => headerMap[h.trim()] = i);

  // Determine Activity ID column
  const idColIndex = headerMap['Activity ID'];
  if (idColIndex === undefined) {
    console.error("Critical Error: 'Activity ID' column not found.");
    return 0;
  }

  // Determine Link column (Handle user's 'URL ID' or our 'Strava Link')
  let linkColIndex = headerMap['Strava Link'];
  if (linkColIndex === undefined) linkColIndex = headerMap['URL ID'];
  
  // If still missing, we will add it later if we have new rows
  
  let existingIds = [];
  if (lastRow > 1) {
    const values = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues(); 
    existingIds = values.flat().map(String);
  }

  let newActivities = activities.filter(a => !existingIds.includes(String(a.id)));
  
  if (newActivities.length === 0) {
    return { processed: 0, hasMoreOnPage: false };
  }

  const MAX_ACTIVITIES_PER_RUN = 40;
  const hasMoreOnPage = newActivities.length > MAX_ACTIVITIES_PER_RUN;
  if (hasMoreOnPage) {
    console.log(`Found ${newActivities.length} total new activities, limiting this run to ${MAX_ACTIVITIES_PER_RUN} to avoid rate limits.`);
    newActivities = newActivities.slice(0, MAX_ACTIVITIES_PER_RUN);
  }

  const activityIds = newActivities.map(a => a.id);
  console.log(`Fetching details for ${activityIds.length} activities...`);
  const detailedDataMap = fetchActivitiesDetailsParallel(activityIds);
  
  const newRows = [];
  
  for (const summary of newActivities) {
    const data = detailedDataMap[summary.id];
    const activity = data && data.details ? data.details : summary;
    const photoUrls = data ? data.photos : [];
    
    // Map activity data to the specific columns we have in the sheet
    const row = new Array(lastCol).fill('');
    
    const set = (header, val) => {
      const idx = headerMap[header];
      if (idx !== undefined) row[idx] = val;
    };

    set('Activity ID', String(activity.id));
    set('Name', activity.name);
    set('Type', activity.type);
    set('Distance (mi)', (activity.distance * 0.000621371).toFixed(2));
    set('Elevation (ft)', activity.total_elevation_gain ? (activity.total_elevation_gain * 3.28084).toFixed(0) : 0);
    set('Moving Time', formatTime(activity.moving_time));
    set('Elapsed Time', formatTime(activity.elapsed_time));
    set('Start Date', activity.start_date_local);
    set('Max Speed (mph)', (activity.max_speed * 2.23694).toFixed(1));
    set('Avg Speed (mph)', (activity.average_speed * 2.23694).toFixed(1));
    set('Calories', activity.calories || '');
    set('Description', activity.description || '');
    set('Temp (C)', activity.average_temp || '');
    set('Achievement Count', activity.achievement_count || 0);
    set('Kudos Count', activity.kudos_count || 0);
    set('Comment Count', activity.comment_count || 0);
    
    // Gear Logic
    const gearName = activity.gear ? activity.gear.name : '';
    const bikeTypes = ['Ride', 'VirtualRide', 'EBikeRide', 'GravelRide', 'MountainBikeRide'];
    if (bikeTypes.includes(activity.type)) set('Bike', gearName);
    else set('Shoes', gearName);

    // Results/Photos
    set('Photos', photoUrls.join(', \n'));
    
    // Results extraction
    let results = '';
    if (activity.segment_efforts && Array.isArray(activity.segment_efforts)) {
      const achievements = [];
      activity.segment_efforts.forEach(effort => {
        if (effort.achievements) {
          effort.achievements.forEach(a => achievements.push(`${a.type_descr || a.type}: ${effort.name}`));
        }
      });
      results = achievements.join('; ');
    }
    set('Results', results);

    // Link
    const link = `https://www.strava.com/activities/${activity.id}`;
    if (headerMap['Strava Link'] !== undefined) set('Strava Link', link);
    else if (headerMap['URL ID'] !== undefined) set('URL ID', link);
    else {
      // If neither exists, we'll append it to the row and update headerMap for subsequent rows
      row.push(link);
    }

    newRows.push(row);
  }

  if (newRows.length > 0) {
    // Check if we grew the row length (added a missing Link column)
    if (newRows[0].length > lastCol) {
       console.log("Adding missing Strava Link column header...");
       sheet.getRange(1, lastCol + 1).setValue('Strava Link');
    }
    sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
  }
  
  return { processed: newRows.length, hasMoreOnPage: hasMoreOnPage };
}

function formatTime(seconds) {
  if (!seconds) return '00:00:00';
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return [h, m, s].map(v => v.toString().padStart(2, '0')).join(':');
}

function createTrigger(funcName, minutesAfter) {
  deleteTrigger(funcName); 
  ScriptApp.newTrigger(funcName)
      .timeBased()
      .after(minutesAfter * 60 * 1000)
      .create();
}

function deleteTrigger(funcName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * BACKFILL TOOL: Adds Strava Links to existing records.
 * Specifically handles the ~400 records already in the sheet.
 */
function backfillStravaLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    alertUser(`Sheet '${TARGET_SHEET_NAME}' not found.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    alertUser("No data to backfill.");
    return;
  }

  // Detect Columns
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idColIndex = headers.indexOf('Activity ID');
  let linkColIndex = headers.indexOf('Strava Link');
  if (linkColIndex === -1) linkColIndex = headers.indexOf('URL ID'); // Handle manual column

  if (idColIndex === -1) {
    alertUser("Could not find 'Activity ID' column.");
    return;
  }
  
  if (linkColIndex === -1) {
    console.log("Adding 'Strava Link' header...");
    linkColIndex = lastCol; 
    sheet.getRange(1, linkColIndex + 1).setValue('Strava Link');
  }

  // 2. Read IDs and Existing Links
  const idData = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues(); 
  const linkData = sheet.getRange(2, linkColIndex + 1, lastRow - 1, 1).getValues();
  
  const updates = [];
  let count = 0;

  for (let i = 0; i < idData.length; i++) {
    const activityId = String(idData[i][0]);
    const currentLink = linkData[i][0];

    if (activityId && activityId !== "" && !currentLink) {
      updates.push([`https://www.strava.com/activities/${activityId}`]);
      count++;
    } else {
      updates.push([currentLink]); 
    }
  }

  if (count > 0) {
    sheet.getRange(2, linkColIndex + 1, updates.length, 1).setValues(updates);
    alertUser(`Successfully backfilled ${count} Strava Links.`);
  } else {
    alertUser("No links needed backfilling.");
  }
}
