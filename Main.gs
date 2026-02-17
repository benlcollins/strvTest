/**
 * Main entry point for the Strava-to-Sheets Sync.
 */

// Global constant for sheet headers if needed, but logic handles appending.
const HEADERS = ['Activity ID', 'Name', 'Type', 'Distance (m)', 'Moving Time (s)', 'Start Date', 'Photo'];

/**
 * Creates a custom menu in the Google Sheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Strava Tools')
      .addItem('Sync Now', 'syncActivities')
      .addToUi();
}

/**
 * Main function to sync Strava activities to the sheet.
 * Handles deduplication and photo fetching.
 */
function syncActivities() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Target sheet "${TARGET_SHEET_NAME}" not found. Please create it.`);
    return;
  }

  // Get existing Activity IDs to prevent duplicates
  const lastRow = sheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) { // Assuming headers are in row 1
    // Column A contains Activity IDs
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
    existingIds = values.flat().map(String); // Ensure comparison as strings
  }

  // Fetch recent activities
  const activities = fetchActivities();
  if (!activities || activities.length === 0) {
    console.log('No activities found or error fetching.');
    return;
  }

  const newRows = [];

  // Sort activities by date ascending if API returns descending, to keep sheet chronological?
  // Strava API usually returns descending. Let's process valid ones.
  // We want to add only NEW activities. simple check:
  
  for (const activity of activities) {
    const activityId = String(activity.id); // Ensure string comparison
    
    if (existingIds.includes(activityId)) {
      continue; // Skip duplicates
    }
    
    // It's a new activity! Fetch photo.
    let photoUrl = fetchActivityPhotos(activity.id);
    let photoCellFormula = '';
    
    if (photoUrl) {
      photoCellFormula = `=HYPERLINK("${photoUrl}", "View Photo")`;
    } else {
      photoCellFormula = 'No Photo';
    }

    /*
     * Columns Mapping:
     * A: Activity ID
     * B: Name
     * C: Type
     * D: Distance (meters)
     * E: Moving Time (seconds)
     * F: Start Date
     * G: Photo (Formula)
     */
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
    // Append all new rows at once
    // Note: Since we have formulas, using setValues might treat them as strings if not careful,
    // but usually setValues handles formula strings correctly if they start with =.
    sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
    console.log(`Added ${newRows.length} new activities.`);
    SpreadsheetApp.getUi().alert(`Synced ${newRows.length} new activities.`);
  } else {
    console.log('No new activities to add.');
    SpreadsheetApp.getUi().alert('No new activities found.');
  }
}
