function onEdit(event) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    //Script Last Update Timming

    var actSht = event.source.getActiveSheet();
    var actRng = event.source.getActiveRange();

    var activeCell = actSht.getActiveCell();
    var row = activeCell.getRow();
    var column = activeCell.getColumn();

    // get last column index
    var lastColumn = actSht.getLastColumn();
    var colNums = [];
    for (var i = 0; i < lastColumn; i++) {
        colNums.push(i);
    }

    if (row < 2) return; //If header row then return
    //var colNums  = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]; //Coulmns, whose edit is considered
    if (colNums.indexOf(column) == -1) return; //If column other than considered then return


    var index = actRng.getRowIndex();
    var dateCol = actSht.getLastColumn();
    var lastCell = actSht.getRange(index, dateCol);
    var date = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy HH:mm");

    // Note: Insert the Date when someone update the row in the last coulmn
    lastCell.setValue(date);

    // Nota: Insert a comment in the lastCell with the user who modify that row
    lastCell.setComment(Session.getEffectiveUser());
}
