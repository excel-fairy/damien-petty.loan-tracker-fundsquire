
/**
 * Called by custom menu
 */
function openCreateEntityPopup() {
    var htmlTemplate = HtmlService.createTemplateFromFile('createentity');
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Import entity')
        .setWidth(705)
        .setHeight(420);
    SpreadsheetApp.getUi().showDialog(htmlOutput);
}



// noinspection JSUnusedGlobalSymbols
/**
 * Main function
 * Called by HTML button in popup
 */
function createEntity(data) {
    SpreadsheetApp.getUi().alert ('Entity is being created. It will appear in the "Entities" tab shortly');
    insertEntityInEntitiesSheet(data);
}

function insertEntityInEntitiesSheet(data){
    var rowToInsert = buildEntityToInsert(data);
    var entitiesOriginalSheet = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet;
    var entityBeforeEntityToInsertRow = getOffsetOfEntityBeforeEntityToInsertAlphabeticalOrder(data.entityName);
    var rangeRowToSet = entitiesOriginalSheet.getRange(entityBeforeEntityToInsertRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn),
        1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.lastEntityColumn)
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn) + 1);


    duplicateEntityRow(entityBeforeEntityToInsertRow);
    rangeRowToSet.setValues([rowToInsert]);

    duplicateCellFromRowAbove('B', entitiesOriginalSheet, entityBeforeEntityToInsertRow + 1);
    duplicateCellFromRowAbove('C', entitiesOriginalSheet, entityBeforeEntityToInsertRow + 1);
    duplicateCellFromRowAbove('M', entitiesOriginalSheet, entityBeforeEntityToInsertRow + 1);
    duplicateCellFromRowAbove('N', entitiesOriginalSheet, entityBeforeEntityToInsertRow + 1);
}

function duplicateCellFromRowAbove(columnLetter, sheet, newEntityRow) {
    var rangeToDuplicateFrom = sheet.getRange(newEntityRow - 1,
        ColumnNames.letterToColumn(columnLetter), 1, 1);
    var rangeToDuplicateTo = sheet.getRange(newEntityRow,
        ColumnNames.letterToColumn(columnLetter), 1, 1);
    rangeToDuplicateFrom.copyTo(rangeToDuplicateTo);
}

function getOffsetOfEntityBeforeEntityToInsertAlphabeticalOrder(entityToInsertName) {
    var entitiesSheet = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet;
    var entitiesRange = entitiesSheet.getRange(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn),
        entitiesSheet.getLastRow(),
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.lastEntityColumn) -
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn)+1);
    var allEntities = entitiesRange.getValues();
    var firstEntityName = allEntities[0][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entityNameColumn)];
    for(var i=0; i < allEntities.length; i++){
        var currentEntityName = allEntities[i][ColumnNames.letterToColumnStart0(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.entityNameColumn)];
        if(currentEntityName.localeCompare(entityToInsertName) > 0)
            return i + (INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityRow - 1);
    }
    return allEntities.length;
}

// Duplicate row to get all the data that won't be overwritten
function duplicateEntityRow(entityRow){
    var entitiesSheet = INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.sheet;
    entitiesSheet.insertRowAfter(entityRow);
    var rangeRowOfEntity = entitiesSheet.getRange(entityRow,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn),
        1,
        entitiesSheet.getLastColumn()
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn) + 1);
    var rangeRowToCopyDestination = entitiesSheet.getRange(entityRow + 1,
        ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn),
        1,
        entitiesSheet.getLastColumn()
        - ColumnNames.letterToColumn(INTEREST_STATEMENT_SPREADSHEET.entitiesSheet.firstEntityColumn) + 1);
    rangeRowOfEntity.copyTo(rangeRowToCopyDestination);
}

function buildEntityToInsert(data) {
    var row = [];
    row[ColumnNames.letterToColumnStart0('A')] = data.entityName;
    row[ColumnNames.letterToColumnStart0('B')] = null;
    row[ColumnNames.letterToColumnStart0('C')] = null;
    row[ColumnNames.letterToColumnStart0('D')] = 0;
    row[ColumnNames.letterToColumnStart0('E')] = data.abnAbc;
    row[ColumnNames.letterToColumnStart0('F')] = data.primaryContact;
    row[ColumnNames.letterToColumnStart0('G')] = data.emailAddress;
    row[ColumnNames.letterToColumnStart0('H')] = data.phoneNumber;
    row[ColumnNames.letterToColumnStart0('I')] = data.accountName;
    row[ColumnNames.letterToColumnStart0('J')] = data.bsbNumber;
    row[ColumnNames.letterToColumnStart0('K')] = data.accountNumber;
    row[ColumnNames.letterToColumnStart0('L')] = data.firstName;
    row[ColumnNames.letterToColumnStart0('M')] = null;
    row[ColumnNames.letterToColumnStart0('N')] = null;
    row[ColumnNames.letterToColumnStart0('O')] = data.carbonCopyEmailAddress;
    return row;
}