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
  'Achievement Count', // New
  'Kudos Count', // New
  'Comment Count', // New
  'Bike', 
  'Shoes', 
  'Results',
  'Photos'
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
    SpreadsheetApp.getUi().alert("No activities found.");
    return;
  }
  
  const activityId = activities[0].id;
  const detail = fetchActivityDetails(activityId); 
  
  if (!detail) {
    SpreadsheetApp.getUi().alert("Failed to fetch details.");
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
  SpreadsheetApp.getUi().alert(`Debug JSON (filtered) written to 'Debug' sheet.`);
}

/**
 * Initiates the sync process. Managed via Triggers.
 */
function runBatchSync() {
  const props = PropertiesService.getScriptProperties();
  let currentPage = parseInt(props.getProperty('SYNC_CURSOR_PAGE')) || 1;
  
  console.log(`Starting batch sync for page ${currentPage}`);
  
  const activities = fetchActivities(currentPage, BATCH_SIZE);
  
  if (!activities || activities.length === 0) {
    console.log('No more activities found. Sync complete.');
    props.deleteProperty('SYNC_CURSOR_PAGE');
    deleteTrigger('runBatchSync');
    SpreadsheetApp.getUi().alert('Sync Complete!');
    return;
  }
  
  const newRows = processActivities(activities);
  
  if (newRows > 0) {
    console.log(`Processed page ${currentPage}. Moving to next page.`);
    props.setProperty('SYNC_CURSOR_PAGE', String(currentPage + 1));
    createTrigger('runBatchSync', 10); 
    SpreadsheetApp.getUi().alert(`Batch ${currentPage} complete. Next batch scheduled in 10 mins.`);
  } else {
    console.log(`Page ${currentPage} processed (all duplicates?). Moving to next page.`);
    props.setProperty('SYNC_CURSOR_PAGE', String(currentPage + 1));
    createTrigger('runBatchSync', 10);
  }
}

function testSync() {
  console.log('Starting Test Sync...');
  const activities = fetchActivities(1, 5); 
  
  if (activities && activities.length > 0) {
    const count = processActivities(activities);
    SpreadsheetApp.getUi().alert(`Test Sync Complete. Processed ${count} new activities.`);
  } else {
    SpreadsheetApp.getUi().alert('Test Sync: No activities found.');
  }
}

/**
 * Helper to process a list of activities and write valid ones to sheet.
 * Uses PARALLEL fetching for speed.
 * @param {Array} activities - Summary activity objects
 * @return {number} Count of new rows added.
 */
function processActivities(activities) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) return 0;
  
  // Update header check to match new length if needed, or user clears sheet
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  const lastRow = sheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
    existingIds = values.flat().map(String);
  }

  const newActivities = activities.filter(a => !existingIds.includes(String(a.id)));
  
  if (newActivities.length === 0) {
    return 0;
  }

  const activityIds = newActivities.map(a => a.id);
  console.log(`Fetching details for ${activityIds.length} activities in parallel...`);
  const detailedDataMap = fetchActivitiesDetailsParallel(activityIds);
  
  const newRows = [];
  
  for (const summary of newActivities) {
    const data = detailedDataMap[summary.id];
    const activity = data && data.details ? data.details : summary;
    const photoUrls = data ? data.photos : [];
    
    let photosCell = '';
    if (photoUrls.length > 0) {
      photosCell = photoUrls.join(', \n');
    }

    const distanceMi = (activity.distance * 0.000621371).toFixed(2);
    const elevationFt = activity.total_elevation_gain ? (activity.total_elevation_gain * 3.28084).toFixed(0) : 0;
    const movingTime = formatTime(activity.moving_time);
    const elapsedTime = formatTime(activity.elapsed_time);
    const maxSpeedMph = (activity.max_speed * 2.23694).toFixed(1);
    const avgSpeedMph = (activity.average_speed * 2.23694).toFixed(1);
    
    const temp = activity.average_temp || '';
    const desc = activity.description || '';
    
    // Counts
    const achieveCount = activity.achievement_count || 0;
    const kudosCount = activity.kudos_count || 0;
    const commentCount = activity.comment_count || 0;

    // Gear
    let bike = '';
    let shoes = '';
    const gearName = activity.gear ? activity.gear.name : '';
    const bikeTypes = ['Ride', 'VirtualRide', 'EBikeRide', 'GravelRide', 'MountainBikeRide'];
    if (bikeTypes.includes(activity.type)) {
      bike = gearName;
    } else {
      shoes = gearName;
    }

    // Results
    let results = '';
    if (activity.segment_efforts && Array.isArray(activity.segment_efforts)) {
      const achievementsList = [];
      activity.segment_efforts.forEach(effort => {
        if (effort.achievements && effort.achievements.length > 0) {
          effort.achievements.forEach(a => {
            let typeName = a.type_descr;
            if (!typeName) {
                if (a.type === 'pr') typeName = 'PR';
                else if (a.type === 'overall') typeName = 'KoM/QoM';
                else if (a.type === 'year_overall') typeName = 'Year Best';
                else typeName = a.type || 'Achievement';
            }
            if (a.rank && a.rank > 1) {
                typeName += ` (${a.rank})`;
            }
            achievementsList.push(`${typeName}: ${effort.name}`);
          });
        }
      });
      results = achievementsList.join('; ');
    } else if (activity.achievements && activity.achievements.length > 0) {
       results = activity.achievements.map(a => a.type || 'Award').join(', ');
    }

    newRows.push([
      String(activity.id),
      activity.name,
      activity.type,
      distanceMi,
      elevationFt,
      movingTime,
      elapsedTime,
      activity.start_date_local,
      maxSpeedMph,
      avgSpeedMph,
      activity.calories || '',
      desc,
      temp,
      achieveCount, // New
      kudosCount, // New
      commentCount, // New
      bike,
      shoes,
      results,
      photosCell
    ]);
  }

  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
    console.log(`Added ${newRows.length} new activities.`);
  }
  
  return newRows.length;
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
