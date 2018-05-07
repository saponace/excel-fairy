var EXPORT_FOLDER_ID = '#####';
var INTEREST_STATEMENT_SHEET = {
    name: 'Interest statement',
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interest statement'),
    dateCell: 'H1',
    totalCell: 'H35',
    entityCell: 'C3',
    pdfExportRange: {
        r1: 5,
        r2: 47,
        c1: 1,
        c2: 8
    }
};

var ENTITIES_SHEET = {
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entity"),
    entityNameColumn: letterToColumnStart0('A'),
    emailAddressColumn: letterToColumnStart0('G'),
    emailSubjectColumn: letterToColumnStart0('M'),
    emailBodyColumn: letterToColumnStart0('N'),
    entitiesListRange:{
        r1: 3,
        r2: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entity").getLastRow(),
        c1: letterToColumn('A'),
        c2: letterToColumn('N')
    }
};

function exportInterestStatements() {
    var entitiesNames = getEntitiesNames();
    var sheetUpdateInterval = 500; // Interval in ms between two entities switch. To let the spreadsheet to update itself. Not sure if needed
    var gSpreadSheetRateLimitingMinInterval = 6000; // Interval in ms between two exports. Google spreadsheet API (used to export sheet to PDF.
    // Returns HTTP 429 for rate limiting if too many requests are sent simultaneously
    for(var i=0; i < entitiesNames.length; i++){
        INTEREST_STATEMENT_SHEET.sheet.getRange(INTEREST_STATEMENT_SHEET.entityCell).setValue(entitiesNames[i]);
        var totalValue = INTEREST_STATEMENT_SHEET.sheet.getRange(INTEREST_STATEMENT_SHEET.totalCell).getValue();
        Utilities.sleep(sheetUpdateInterval);
        if(totalValue !== 0){
            exportInterestStatementForCurrentEntity();
            Utilities.sleep(gSpreadSheetRateLimitingMinInterval-sheetUpdateInterval);
        }
    }
}

function exportInterestStatementForCurrentEntity(){
    var dateStr = INTEREST_STATEMENT_SHEET.sheet.getRange(INTEREST_STATEMENT_SHEET.dateCell).getValue();
    var entity = INTEREST_STATEMENT_SHEET.sheet.getRange(INTEREST_STATEMENT_SHEET.entityCell).getValue();
    var fileName = entity + ' - ' + INTEREST_STATEMENT_SHEET.name + ' - ' + dateStr;
    var exportFolderId = getFolderToExportPdfTo(EXPORT_FOLDER_ID, dateStr).getId();

    var exportOptions = {
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: INTEREST_STATEMENT_SHEET.pdfExportRange
    };
    var exportedFile = savePdf(exportOptions);
    sendEmail(exportedFile);
}

function sendEmail(attachment) {
    var entityName = INTEREST_STATEMENT_SHEET.sheet.getRange(INTEREST_STATEMENT_SHEET.entityCell).getValue();
    var entity = getEntityFromName(entityName);
    if(!entity)
        SpreadsheetApp.getActiveSpreadsheet().toast('Entity ' + entityName + ' not found in entities list. No email sent');
    else {
        var recipient = entity[ENTITIES_SHEET.emailAddressColumn];
        var subject = entity[ENTITIES_SHEET.emailSubjectColumn];
        var message = entity[ENTITIES_SHEET.emailBodyColumn];
        var emailOptions = {
            attachments: [attachment.getAs(MimeType.PDF)],
            name: 'Automatic loan tracker mail sender'
        };
        MailApp.sendEmail(recipient, subject, message, emailOptions);
    }
}

function getEntitiesNames(){
    var entities = ENTITIES_SHEET.sheet.getRange(ENTITIES_SHEET.entitiesListRange.r1,
        ENTITIES_SHEET.entitiesListRange.c1,
        ENTITIES_SHEET.entitiesListRange.r2 - ENTITIES_SHEET.entitiesListRange.r1,
        ENTITIES_SHEET.entitiesListRange.c2 - ENTITIES_SHEET.entitiesListRange.c1+1).getValues();
    return entities.map(function(entity){return entity[ENTITIES_SHEET.entityNameColumn];});
}

function getEntityFromName(entityName){
    var entities = ENTITIES_SHEET.sheet.getRange(ENTITIES_SHEET.entitiesListRange.r1,
        ENTITIES_SHEET.entitiesListRange.c1,
        ENTITIES_SHEET.entitiesListRange.r2 - ENTITIES_SHEET.entitiesListRange.r1,
        ENTITIES_SHEET.entitiesListRange.c2 - ENTITIES_SHEET.entitiesListRange.c1+1).getValues();

    for (var i=0; i < entities.length; i++){
        if(entities[i][ENTITIES_SHEET.entityNameColumn] === entityName)
            return entities[i];
    }
    return null;
}
