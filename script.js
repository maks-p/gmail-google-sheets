function getMessage() {
  
    // Search term that returns only relevant emails in Gmail
    var searchTerm = 'from:noreply@messaging.squareup.com AND Daily Sales Summary Report'
    
    var thread = GmailApp.search(searchTerm, 0, 10)[0];
    var message = thread.getMessages()[0];
    Logger.log(message.getSubject());
    
    return message;
}
  
  
function connectSpreadsheet() {

    var url = 'https://docs.google.com/spreadsheets/d/1SMHd310dIH2DkCW6UTh4xOsfAssZ6ozkfFYZrV-cB78/edit#gid=0';
    var ss = SpreadsheetApp.openByUrl(url);

    return ss;
}
    
  
function appendHTMLarray(){

    var ss = connectSpreadsheet();
    var testData = ss.getSheetByName('test_data');

    var message = getMessage();
    var messageBody = message.getBody();
    var line_array = messageBody.split("\n");

    // Append HTML lines to Spreadsheet
    for (i = 0; i < line_array.length; i++){
        testData.appendRow([line_array[i]]);
        };
}


function getLastDate(){

    var ss = connectSpreadsheet();
    var salesData = ss.getSheetByName('sales_data');

    var lastRow = salesData.getMaxRows();
    var values = salesData.getRange("A1:" + 'A' + lastRow).getValues();

    for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {}
    var date = new Date(values[lastRow - 1]);

    return date;
}
  
  
function parseEmailPOS() {

    var ss = connectSpreadsheet();
    var salesData = ss.getSheetByName('sales_data');

    var message = getMessage();
    var messageBody = message.getBody();
    var line_array = messageBody.split("\n");
    
    // POS Specific Data Here
    var reportDate = new Date(message.getSubject().split(" for ")[1]);
    var BusinessDay = line_array[212].split(">")[1].split(' ')[0];
    var adjustedGrossSales = line_array[299].split('$')[1];
    var discounts = line_array[380].split('$')[1];
    var salesTax = line_array[437].split('$')[1];
    var netSales = line_array[413].split('$')[1];
    var ccTips = line_array[470].split('$')[1];
   
    // Check Date function 
    function checkDate() {
      if (reportDate.getTime() === getLastDate().getTime()) {
        return true;
        } else {
        return false; 
       }
    };
    
    if (!checkDate()) {
      
      // POS Specific Data Here
      if (line_array[399].trim() === '<!-- Net Sales Partial -->' &&
          line_array[456].trim() === '<!-- Tips Partial -->') {
        
         salesData.appendRow([reportDate, 
                              BusinessDay, 
                              adjustedGrossSales,
                              discounts,
                              salesTax,
                              netSales,
                              ccTips]);
        
        } else {
        salesData.appendRow(['ERROR - Confirm Value Indexes']);
        };
    } else {
      Logger.log('Date Already Exists');
      }
};
  
  
  