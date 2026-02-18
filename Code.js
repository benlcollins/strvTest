function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Strava Tools')
        .addItem('Fetch Data', 'myStravaFunction')
        .addToUi();
}

function myStravaFunction() {
    Browser.msgBox("Hello from Clasp!");
}

function clearTokens() {
    PropertiesService.getScriptProperties().deleteAllProperties();
    console.log("Old tokens cleared. Now try running your Sync again.");
}