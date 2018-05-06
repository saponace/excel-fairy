var EXPORT_FOLDER_ID = '#####';
var INTEREST_STATEMENT_SHEET = {
    name: 'Interest statement',
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interest statement'),
    dateCell: 'H1',
    pdfExportRange: {
        r1: 5,
        r2: 47,
        c1: 1,
        c2: 8
    }
};
var EMAIL_LIST_SHEET = {
    sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email list"),
    recipientsListRange:{
        r1: 3,
        r2: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email list").getLastRow() - 2,
        c1: 1,
        c2: 3
    }
};

function exportToPdf() {
    var dateStr = INTEREST_STATEMENT_SHEET.sheet.getRange(INTEREST_STATEMENT_SHEET.dateCell).getValue();
    var exportFolderId = getFolderToExportPdfTo(EXPORT_FOLDER_ID, dateStr).getId();
    var fileName = INTEREST_STATEMENT_SHEET.name + ' ' + dateStr;

    var exportOptions = {
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: INTEREST_STATEMENT_SHEET.pdfExportRange
    };
    var exportedFile = savePdf(exportOptions);
    sendEmails(exportedFile);
}

function sendEmails(attachment) {
    var emails = EMAIL_LIST_SHEET.sheet.getRange(EMAIL_LIST_SHEET.recipientsListRange.r1,
        EMAIL_LIST_SHEET.recipientsListRange.c1,
        EMAIL_LIST_SHEET.recipientsListRange.r2,
        EMAIL_LIST_SHEET.recipientsListRange.c2).getValues();

    for (var i=0; i < emails.length; i++){
        var email = emails[i];
        var recipient = email[0];
        var subject = email[1];
        var message = email[2];
        var emailOptions = {
            attachments: [attachment.getAs(MimeType.PDF)],
            name: 'Automatic loan tracker mail sender'
        };
        MailApp.sendEmail(recipient, subject, message, emailOptions);
    }
}


