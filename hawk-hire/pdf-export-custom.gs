/**
 * v0.1 - custom
 * Export one or all sheets in a spreadsheet as PDF files on user's Google Drive,
 * in same folder that contained original spreadsheet.
 *
 * Adapted from https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c25
 *
 * @param {Object}  opts       (optional) Export options
 *                                Can contain any combination of fields
 *                                in following example:
 *                                {
 *                                    spreadSheetId: 'spreadSheetId',
 *                                    sheetId: 'sheetId',
 *                                    exportFolderId: 'folderId',
 *                                    exportFileName: 'file',
 *                                    range: {
 *                                        r1: 0,
 *                                        r2: 0,
 *                                        c1: 0,
 *                                        c2: 0
 *                                    }
 *                                }
 */
function save(opts) {

    opts = !!opts ? opts : {};

    // If a sheet ID was provided, open that sheet, otherwise assume script is
    // sheet-bound, and open the active spreadsheet.
    var ss = (opts.spreadSheetId) ? SpreadsheetApp.openById(opts.spreadSheetId) : SpreadsheetApp.getActiveSpreadsheet();

    // Get URL of spreadsheet, and remove the trailing 'edit'
    var url = ss.getUrl().replace(/edit$/,'');

    // Get folder containing spreadsheet, for later export
    // If folder ID is provided, use it. Otherwise export to
    // same folder as the one containing the spreadsheet
    var folder;
    if(opts.exportFolderId){
        folder = DriveApp.getFolderById(opts.exportFolderId);
    }
    else {
        var parents = DriveApp.getFileById(ss.getId()).getParents();
        if (parents.hasNext()) {
            folder = parents.next();
        }
        else {
            folder = DriveApp.getRootFolder();
        }
    }

    // Set range url parameters
    var rangeParameters = '';
    if(opts.range
        && opts.range.r1 && opts.range.r1 === parseInt(opts.range.r1, 10)
        && opts.range.r2 && opts.range.r2 === parseInt(opts.range.r2, 10)
        && opts.range.c1 && opts.range.c1 === parseInt(opts.range.c1, 10)
        && opts.range.c2 && opts.range.c2 === parseInt(opts.range.c2, 10))
        rangeParameters = '&r1=' + opts.range.r1 +
            '&r2=' + opts.range.r2 +
            '&c1=' + opts.range.c1 +
            '&c2=' + opts.range.c2;


    // If provided a sheetId, save it, otherwise save active sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    if(opts.sheetId){
        var sheets = ss.getSheets();
        for (var i=0; i<sheets.length; i++) {
            var currentSheet = sheets[i];
            if (opts.sheetId === currentSheet.getSheetId())
                sheet = currentSheet;
        }
    }

//additional parameters for exporting the sheet as a pdf
    var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf
        + '&gid=' + sheet.getSheetId()   //the sheet's Id
        // following parameters are optional...
        + '&size=letter'      // paper size
        + '&portrait=true'    // orientation, false for landscape
        + '&fitw=true'        // fit to width, false for actual size
        + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
        + '&gridlines=false'  // hide gridlines
        + rangeParameters     // range
        + '&fzr=true';       // do not repeat row headers (frozen rows) on each page

    var options = {
        headers: {
            'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
        }
    };

    var response = UrlFetchApp.fetch(url + url_ext, options);

    var fileName;
    if(opts.exportFileName)
        fileName = opts.exportFileName + '.pdf';
    else
        fileName = ss.getName() + ' - ' + sheet.getName() + '.pdf';

    var blob = response.getBlob().setName(fileName);
    folder.createFile(blob);
}
