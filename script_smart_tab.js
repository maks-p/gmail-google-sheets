function getMessage() {
  
    // Search term that returns only relevant emails in Gmail
    var searchTerm = 'from:admin@jupiterdisco.smarttab.com AND Daily Sales Report'
    
    var thread = GmailApp.search(searchTerm, 0, 10)[0];
    var message = thread.getMessages()[0];
    Logger.log(message.getSubject());
    
    return message;
}
  
  
function connectSpreadsheet() {

    var url = 'https://docs.google.com/spreadsheets/d/<YOUR_KEY_HERE>/edit#gid=0';
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
    var reportDate = new Date(line_array[13].split(">")[1].replace('</td', '').replace(/\./g,'/'))
    var BusinessDay = line_array[12].split(">")[1].replace('</td', '')
    var adjustedGrossSales = line_array[42].split("$")[1].replace(/[^\d.]/g, '')
    var discounts = line_array[46].split('$')[1].replace(/[^\d.]/g, '')
    var salesTax = line_array[58].split('$')[1].replace(/[^\d.]/g, '')
    var netSales = line_array[54].split('$')[1].replace(/[^\d.]/g, '')
    var ccTips = line_array[73].split('$')[1].replace(/[^\d.]/g, '')
   
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
      if (line_array[36].trim() === '<!-- REVENUE -->' &&
          line_array[72].trim() === '<td style="font-family:arial;font-size:14px;">+ TIPS:</td>') {
        
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
  
  
  