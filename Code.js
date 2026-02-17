function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Strava Tools')
        .addItem('Fetch Data', 'myStravaFunction')
        .addToUi();
}

function myStravaFunction() {
    Browser.msgBox("Hello from Clasp!");
}