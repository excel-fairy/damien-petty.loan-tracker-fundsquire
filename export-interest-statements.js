/**
 * Main function
 * Called by custom menu
 */
function exportInterestStatements() {
    var entitiesNames = getEntitiesNames();
    var sheetUpdateInterval = 500; // Interval in ms between two entities switch. To let the spreadsheet to update itself. Not sure if needed
    var gSpreadSheetRateLimitingMinInterval = 6000; // Interval in ms between two exports. Google spreadsheet API (used to export sheet to PDF.
                                                    // Returns HTTP 429 for rate limiting if too many requests are sent simultaneously
    var currentlyExportingEntity = getCurrentlyExportingEntityFromCalc();
    var currentlyExportingBorrower = getCurrentlyExportingBorrowerFromCalc();
    var entitiesStartIndex = 0;
    var exportExecutionStartDate = new Date();
    if(currentlyExportingEntity !== '')
        entitiesStartIndex = entitiesNames.indexOf(currentlyExportingEntity);
    for(var i = entitiesStartIndex; i < entitiesNames.length; i++){
        var entityName = entitiesNames[i];
        setCurrentlyExportingEntity(entityName);

        var borrowersOfEntity = getBorrowersOfEntity(entityName);
        if(borrowersOfEntity.length !== 0) {
            var borrowersStartIndex = 0;
            if(currentlyExportingBorrower !== '')
                borrowersStartIndex = borrowersOfEntity.indexOf(currentlyExportingBorrower);
            for(var j = borrowersStartIndex; j < borrowersOfEntity.length; j++) {
                var borrower = borrowersOfEntity[j];
                setCurrentlyExportingBorrower(borrower);

                // Stops if script execution is becoming too close from the GAS limit per script (limit is 5min, stops at 4m30s)
                if (isTimeUp(exportExecutionStartDate)) {
                    var currentlyExportingEntityIndex = entitiesNames.indexOf(getCurrentlyExportingEntityFromCalc());
                    var lastExportedEntity = currentlyExportingEntityIndex > 0 ? currentlyExportingEntityIndex - 1 : 0;
                    updateExportStatus(false);
                    SpreadsheetApp.getActiveSpreadsheet().toast('Script execution is too long and had to stop. The last ' +
                        'exported entity is ' + lastExportedEntity + '. Next execution will start from here');
                    return;
                }
                else
                    updateExportStatus(true);

                var totalOfCurrentMonth = getTotalOfCurrentMonthForCurrentEntityAndBorrower();
                Utilities.sleep(sheetUpdateInterval);
                if (totalOfCurrentMonth !== 0) {
                    exportInterestStatementForCurrentEntityAndBorrower();
                    Utilities.sleep(gSpreadSheetRateLimitingMinInterval - sheetUpdateInterval);
                }
            }
        }
    }
    setCurrentlyExportingEntity('');
    setCurrentlyExportingBorrower('');
    updateExportStatus(false)
}

/**
 * Get the total amount for the entity currently selected at the selected month
 */
function getTotalOfCurrentMonthForCurrentEntityAndBorrower(){
    var allTransactions = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(
        INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.transactionsRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.transactionsRange.c1,
        INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.transactionsRange.r2 - INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.transactionsRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.transactionsRange.c2 - INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.transactionsRange.c1
    ).getValues();
    // Get the total of the last line that is not empty
    retVal = 0;
    for (var i = 0; i < allTransactions.length; i++){
        var loopTransaction = allTransactions[i];
        var loopTransactionTotal = loopTransaction[INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.totalColumn - 1]; // "-1" because transactions range starts at column B
        if(typeof(loopTransactionTotal) === "number")
            var retVal = loopTransactionTotal;
        else
            break;
    }
    return retVal;
}

/**
 * Update the GUI display of the export progress
 * @param executionOnGoing Is the export ongoing (true) or has it stopped (false) ?
 */
function updateExportStatus(executionOnGoing) {
    var textToWrite;
    if(executionOnGoing)
        textToWrite = "Script in progress";
    else {
        textToWrite = 'All Entities exported!';
        var lastExportedEntity = getCurrentlyExportingEntityFromCalc();
        var lastExportedBorrower = getCurrentlyExportingBorrowerFromCalc();
        if(lastExportedEntity !== '')
            textToWrite = 'Exported until entity ' + lastExportedEntity + ' for borrower ' + lastExportedBorrower + ', hit "Export & Send" again to proceed with the export';
    }
    INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.exportStatusCell).setValue(textToWrite);
}

function exportInterestStatementForCurrentEntityAndBorrower(){
    var dateStr = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.dateCell).getValue();
    var entity = getCurrentlyExportingEntityFromInterestStatement();
    var borrower = getCurrentlyExportingBorrowerFromInterestStatement();
    var fileName = entity + ' - ' + borrower + ' - ' + INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.name + ' - ' + dateStr;
    var exportFolderId = getFolderToExportPdfTo(EXPORT_FOLDER_ID, dateStr).getId();

    var exportOptions = {
        exportFolderId: exportFolderId,
        exportFileName: fileName,
        range: INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.pdfExportRange
    };
    var exportedFile = ExportSpreadsheet.export(exportOptions);
    sendEmail(exportedFile);
}

function sendEmail(attachment) {
    var entityName = getCurrentlyExportingEntityFromInterestStatement();
    var borrowerName = getCurrentlyExportingBorrowerFromInterestStatement();
    var entity = getEntityFromName(entityName);
    if(!entity)
        SpreadsheetApp.getActiveSpreadsheet().toast('Entity ' + entityName + ' not found in entities list. No email sent');
    else {
        var recipient = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailAddressColumn];
        var subjectTemplate = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailSubjectColumn];
        var subject = fillBorrower(subjectTemplate, borrowerName);
        var messageTemplate = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailBodyColumn];
        var message = fillBorrower(messageTemplate, borrowerName);
        var carbonCopyEmailAddresses = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.carbonCopyEmailAddressesColumn];
        var emailOptions = {
            attachments: [attachment.getAs(MimeType.PDF)],
            name: borrowerName,
            cc: carbonCopyEmailAddresses
        };
        MailApp.sendEmail(recipient, subject, message, emailOptions);
    }
}


/**
 * Get all the borrowers of a given entity
 * @param entity
 * @return Array of borrowers
 */
function getBorrowersOfEntity(entity){
    var entityNameColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn);
    var borrowerCol0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.borrowerColumn);
    var loans = getAllLoans();
    return loans.filter(function (loan) { // Keep only loans for the entity
        return loan[entityNameColS0] === entity;
    }).map(function (loan) { // Keep only the borrowers of the entities and ditch the rest of data
        return loan[borrowerCol0];
    }).filter(function onlyUnique(value, index, self) { // Deduplicate array
            return self.indexOf(value) === index;
        }
    );
}

function getEntityFromName(entityName){
    var entities = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.r1 + 1,
        INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c2 - INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entitiesListRange.c1 + 1).getValues();

    for (var i=0; i < entities.length; i++){
        if(entities[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entityNameColumn)] === entityName)
            return entities[i];
    }
    return null;
}

function getAllLoans() {
    var loansOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.loansSheet.sheet;
    var loansRange = loansOriginalSheet.getRange(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoanRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn),
        loansOriginalSheet.getLastRow(),
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.lastLoansColumn) -
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.loansSheet.firstLoansColumn)+1);
    return loansRange.getValues();
}

function getCurrentlyExportingEntityFromInterestStatement() {
    return INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.entityCell).getValue()
}

function getCurrentlyExportingEntityFromCalc() {
    return INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingEntityCell).getValue();
}

function setCurrentlyExportingEntity(entityName) {
    INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.entityCell).setValue(entityName);
    INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingEntityCell).setValue(entityName);
}

function getCurrentlyExportingBorrowerFromInterestStatement() {
    return INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.borrowerCell).getValue();
}

function getCurrentlyExportingBorrowerFromCalc() {
    return INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingBorrowerCell).getValue();
}

function setCurrentlyExportingBorrower(borrower) {
    INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.borrowerCell).setValue(borrower);
    INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingBorrowerCell).setValue(borrower);
}

/**
 * Replace the string "{borrower}" in a template string
 * @param templateString template
 * @param borrower Borrower to inject in the template
 */
function fillBorrower(templateString, borrower){
    return templateString.replace("{borrower}", borrower);
}