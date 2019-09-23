
function getMessageDate(message) {
    
    var messageDate = new Date(message.getDate());
    var effectiveDate = new Date(messageDate - (24*60*60*1000))
    //Logger.log(effectiveDate)
    
    return effectiveDate 
}


function getLastDate(){

    var ss = connectSpreadsheet();
    var salesData = ss.getSheetByName('raw_pmix');

    var lastRow = salesData.getMaxRows();
    var values = salesData.getRange("A1:" + 'A' + lastRow).getValues();

    for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--) {}
    var date = new Date(values[lastRow - 1]);

    return date;
}


function getMessage(n) {
    
    // Search term that returns only relevant emails in Gmail
    var searchTerm = 'from:admin@jupiterdisco.smarttab.com AND Daily Sales Report'
    
    var thread = GmailApp.search(searchTerm, 0, 365)[n];
    var message = thread.getMessages()[0];
    Logger.log(message.getSubject())
    return message;
}



function getAttachment(message) {
  
  var attachments = message.getAttachments({includeInlineImages: false,
                                           includeAttachments: true});
  
  return attachments;
      
}


function findPMix(message) {
  
  var attachments = getAttachment(message);
  
  for (i = 0; i < attachments.length; i++) {
    if (attachments[i].getName() === 'ProductReport_Jupiter Disco.htm'); {
      return attachments[i];
    }
  }
}


function connectSpreadsheet() {

    var url = 'https://docs.google.com/spreadsheets/<YOUR_KEY_HERE>/edit#gid=0';
    var ss = SpreadsheetApp.openByUrl(url);

    return ss;
}
  

function filterLogic(item) {
  if (item.replace(/\s+/g, '').length !== 0){
    return true
  } else {
    return false
  }
}


function cleanArray(arr) {
  
  var tempArr = []
  
  for (i = 6; i < arr.length; i++) {
    if (filterLogic(arr[i])) {
        tempArr.push(arr[i].replace(/<[^>]+>/g, ""))
    }
  }
  return tempArr
}


function makeNested(arr, items, message) {
  
    var date = getMessageDate(message)
    var newArr = []
    var arrayLength = Math.ceil(arr.length / items)

    for (i = 0; i < arrayLength; i++) {

        var tempArr = []

        for (j = i * items; j < (i+1) * items; j++) {
            if (arr[j]) {
            tempArr.push(arr[j].trim())
            }
        }
        tempArr.unshift(date)
        newArr.push(tempArr)
    } 
    return newArr
}


function getArray(message) {
  
    var attachment = findPMix(message);
    var attachmentString = attachment.getDataAsString();
    
    return attachmentString.split("\n");  
  
}


function appendPMix(message) {
    
    const cols = 9
  
    const ss = connectSpreadsheet();
    const testData = ss.getSheetByName('raw_pmix');
    
    var arr = getArray(message)
    var cleanArr = cleanArray(arr)
    var newArr = makeNested(cleanArr, cols, message)
    Logger.log(message.getSubject())
   
    for (var i = 0; i < newArr.length; i++) {
      testData.appendRow(newArr[i])
    }
    SpreadsheetApp.flush();
}


function appendOne() {
  
   var message = getMessage(0)
   
   // Check Date function 
    function checkDate() {
      if (getMessageDate(message).getTime() === getLastDate().getTime()) {
        return true;
        } else {
        return false; 
       }
    };
    
    if (!checkDate()) {
      appendPMix(message)
    } else {
      Logger.log('Date Already Exists');
    }
}


function appendAll() {
    
    const batchSize = 100;
    var loop = 1;

    for (var i = batchSize * loop; i < (batchSize * (1 + loop)); i++) {
      Logger.log(i);
      appendOne(i);
    }
 
}

