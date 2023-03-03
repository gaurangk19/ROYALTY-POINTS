function onEdit(e) {
    var ss = e.source;
    var range = e.range;
    var sheetname = ss.getActiveSheet().getName();
    var lastrow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Points').getLastRow();
    var lastcolumn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Points').getLastColumn();
    if (range.getColumn() == 6 && range.getRow() == 10 && sheetname == 'Points' && SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Points').getRange(lastrow - 1, lastcolumn).getValue()) {
        var result = SpreadsheetApp.getUi().alert("Phone Number not present", "Entered Number not in database. Do you want to add new contact?", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
        if (result === SpreadsheetApp.getUi().Button.OK) {
            goToSheet("New Customer", 12, 6);
        }
    }
}

function addnewcontact() {
    var FILE = SpreadsheetApp.openById("1OpbUnYSG7eq6-VXhUgNBDnszFn55st0RTPgZPY6kZDM");
    var Points = FILE.getSheetByName("Points");
    var Customers = FILE.getSheetByName("Customers");
    var New_Customer = FILE.getSheetByName("New Customer");

    if (New_Customer.getRange(12, 6).isBlank() || New_Customer.getRange(13, 6).isBlank()) {
        var result = SpreadsheetApp.getUi().alert("There was a Problem", "Complete details not entered.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    else {
        Customers.getRange(New_Customer.getRange('X1000:X1000').getValue()).setValue(New_Customer.getRange('F11:F11').getValue());
        Customers.getRange(New_Customer.getRange('Y1000:Y1000').getValue()).setValue(New_Customer.getRange('F12:F12').getValue());
        Customers.getRange(New_Customer.getRange('Z1000:Z1000').getValue()).setValue(New_Customer.getRange('F13:F13').getValue());

        New_Customer.getRange('F12:F12').clearContent();
        New_Customer.getRange('F13:F13').clearContent();
    }
}

function goToSheet(sheetName, row, col) {
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    SpreadsheetApp.setActiveSheet(sheet);
    var range = sheet.getRange(row, col)
    SpreadsheetApp.setActiveRange(range);
}

function myLoadButton() {
    var FILE = SpreadsheetApp.openById("1OpbUnYSG7eq6-VXhUgNBDnszFn55st0RTPgZPY6kZDM");
    var Points = FILE.getSheetByName("Points");
    var Customers = FILE.getSheetByName("Customers");

    if (Points.getRange(13, 6).isBlank() || Points.getRange(14, 6).isBlank()) {
        SpreadsheetApp.getUi().alert("There was a Problem", "Final bill is not entered.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    else {
        var cell = Points.getRange(10, 6);
        var lastcolumn = Customers.getLastColumn();
        if (cell.isBlank() == false) {
            if (Customers.getRange(1, lastcolumn).getValue() != "#N/A" && Customers.getRange(2, lastcolumn).getValue() != "#N/A") {
                var oldpoints = Customers.getRange(Customers.getRange(2, lastcolumn).getValue()).getValue();
                Customers.getRange(Customers.getRange(1, lastcolumn).getValue()).setValue(oldpoints);
                Customers.getRange(Customers.getRange(2, lastcolumn).getValue()).setValue(Points.getRange('Z1000:Z1000').getValue());
            }
            Points.getRange(10, 6).clearContent();
            Points.getRange(13, 6, 2, 1).clearContent();
        }
    }
}
