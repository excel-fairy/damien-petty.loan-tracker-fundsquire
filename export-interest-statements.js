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
    var currentlyExportingCurrency = getCurrentlyExportingCurrencyFromCalc();
    var entitiesStartIndex = 0;
    var exportExecutionStartDate = new Date();
    if(currentlyExportingEntity !== '')
        entitiesStartIndex = entitiesNames.indexOf(currentlyExportingEntity);
    for(var i = entitiesStartIndex; i < entitiesNames.length; i++){
        var entityName = entitiesNames[i];
        setCurrentlyExportingEntity(entityName);

        var currenciesOfEntity = getCurrenciesOfEntity(entityName);
        if(currenciesOfEntity.length !== 0) {
            var currenciesStartIndex = 0;
            if(currentlyExportingCurrency !== '')
                currenciesStartIndex = currenciesOfEntity.indexOf(currentlyExportingCurrency);
            for(var j = currenciesStartIndex; j < currenciesOfEntity.length; j++) {
                var currency = currenciesOfEntity[j];
                setCurrentlyExportingCurrency(currency);

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

                var totalOfCurrentMonth = getTotalOfCurrentMonthForCurrentEntityAndCurrency();
                Utilities.sleep(sheetUpdateInterval);
                if (totalOfCurrentMonth !== 0) {
                    exportInterestStatementForCurrentEntityAndCurrency();
                    Utilities.sleep(gSpreadSheetRateLimitingMinInterval - sheetUpdateInterval);
                }
            }
        }
    }
    setCurrentlyExportingEntity('');
    setCurrentlyExportingCurrency('');
    updateExportStatus(false)
}

/**
 * Get the total amount for the entity currently selected at the selected month
 */
function getTotalOfCurrentMonthForCurrentEntityAndCurrency(){
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
        var lastExportedCurrency = getCurrentlyExportingCurrencyFromCalc();
        if(lastExportedEntity !== '')
            textToWrite = 'Exported until entity ' + lastExportedEntity + ' for currency ' + lastExportedCurrency + ', hit "Export & Send" again to proceed with the export';
    }
    INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.exportStatusCell).setValue(textToWrite);
}

function exportInterestStatementForCurrentEntityAndCurrency(){
    var dateStr = INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.dateCell).getValue();
    var entity = getCurrentlyExportingEntityFromInterestStatement();
    var currency = getCurrentlyExportingCurrencyFromInterestStatement();
    var fileName;
    if(currency === "AUD")
        fileName = entity + ' - ' + INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.name + ' - ' + dateStr;
    else
        fileName = entity + ' - ' + INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.name + ' (' + currency + ') - ' + dateStr;
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
    var currencyName = getCurrentlyExportingCurrencyFromInterestStatement();
    var entity = getEntityFromName(entityName);
    if(!entity)
        SpreadsheetApp.getActiveSpreadsheet().toast('Entity ' + entityName + ' not found in entities list. No email sent');
    else {
        var recipient = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailAddressColumn];
        var subjectTemplate = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailSubjectColumn];
        var subject = fillCurrency(subjectTemplate, currencyName);
        var messageTemplate = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.emailBodyColumn];
        var message = fillCurrency(messageTemplate, currencyName);
        var carbonCopyEmailAddresses = entity[INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.carbonCopyEmailAddressesColumn];
        var emailOptions = {
            attachments: [attachment.getAs(MimeType.PDF)],
            name: currencyName,
            cc: carbonCopyEmailAddresses
        };
        MailApp.sendEmail(recipient, subject, message, emailOptions);
    }
}


/**
 * Get all the currencies of a given entity
 * @param entity
 * @return Array of currencies
 */
function getCurrenciesOfEntity(entity){
    var entityNameColS0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.entityNameColumn);
    var currencyCol0 = ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.loansSheet.currencyColumn);
    var loans = getAllLoans();
    return loans.filter(function (loan) { // Keep only loans for the entity
        return loan[entityNameColS0] === entity;
    }).map(function (loan) { // Keep only the currencies of the entities and ditch the rest of data
        return loan[currencyCol0];
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

function getCurrentlyExportingCurrencyFromInterestStatement() {
    return INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.currencyCell).getValue();
}

function getCurrentlyExportingCurrencyFromCalc() {
    return INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingCurrencyCell).getValue();
}

function setCurrentlyExportingCurrency(currency) {
    INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.interestStatementSheet.currencyCell).setValue(currency);
    INTEREST_STATEMENT_SPREADSHEET.calcSheet.sheet.getRange(INTEREST_STATEMENT_SPREADSHEET.calcSheet.currentlyExportingCurrencyCell).setValue(currency);
}

/**
 * Replace the string "{currency}" in a template string
 * @param templateString template
 * @param currency Currency to inject in the template
 */
function fillCurrency(templateString, currency){
    return templateString.replace("{currency}", currency);
}