function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Run script')
        .addItem('Merge Data', 'mergeData')
        .addToUi();
}
