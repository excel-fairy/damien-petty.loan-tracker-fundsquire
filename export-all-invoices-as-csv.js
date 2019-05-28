/**
 * Main function
 * Called by custom menu
 */
function exportInvoicesAsCsv(){
    var dateStr = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.dateCell).getValue();
    var fileName = INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.name + ' - ' + dateStr;
    var exportFolderId = getFolderToExportPdfTo(EXPORT_FOLDER_ID, dateStr).getId();

    var invoices = INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.sheet.getRange(
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c1,
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.r2 - INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c2 - INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c1 + 1
    ).getValues();

    var i = invoices.length;
    var lastInvoiceNumber = null;
    while (i--){
        var invoice = invoices[i];
        if (invoice[INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.descriptionColumn] === '')
            invoices.splice(i, 1);
        else if (!lastInvoiceNumber)
            lastInvoiceNumber = invoice[INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.invoiceNumberColumn];
    }

    var tempFilterSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('temp filter sheet');
    tempFilterSheet.getRange(
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c1,
        invoices.length,
        INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c2 - INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c1 + 1
    ).setValues(invoices);


    var exportOptions = {
        sheetId: tempFilterSheet.getSheetId(),
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: {
            r1: INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.r1 - 1,
            r2: tempFilterSheet.getLastRow(),
            c1: INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c1 - 1,
            c2: INTEREST_STATEMENT_SPREADSHEET.invoicesSheet.exportRange.c2
        },
        fileFormat: 'csv'
    };
    ExportSpreadsheet.export(exportOptions);
    INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.lastInvoiceNumberCell).setValue(lastInvoiceNumber);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempFilterSheet);
}