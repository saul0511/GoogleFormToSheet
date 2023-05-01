// template sheet url
const tUrl = '';  

// set default ID cards pixel width
const idCardWidth = 400;

function onFormSubmit(e) {

  // get responses from google form
 
  //var form = FormApp.openById('1FE_eUVA1Nng6ZR-YuFCvBgX_foxUaRB1WgnUon9WFGg'); // Form ID
  //var formResponses = form.getResponses();

  var formResponses = FormApp.getActiveForm().getResponses();  
  var formCount = formResponses.length;
  var formResponse = formResponses[formCount - 1];
  var itemResponses = formResponse.getItemResponses();


  // set src and dst sheets
  var srcSpreadsheet = SpreadsheetApp.openByUrl(tUrl);
  var dstSpreadsheet = SpreadsheetApp.create("ResultSheet" );

  // declare arrays to save responses
  var titleArray = [];
  var responseArray = [];
  var imageTitleArray = [];
  var imageIdArray = [];

  for (var i = 0; i < itemResponses.length; i++) {
    var itemResponse = itemResponses[i];
    
    // if the response is an image
    if (itemResponse.getItem().getType() == FormApp.ItemType.FILE_UPLOAD) {
     
      var imageId = itemResponse.getResponse();
      
      imageTitleArray.push(itemResponse.getItem().getTitle());
      imageIdArray.push(imageId);
      
    } else {

      // not image 
      titleArray.push(itemResponse.getItem().getTitle());
      responseArray.push(itemResponse.getResponse());
    }
  }
  
  // fullfill the dst sheets
  // replace answers to relative <<title>>
  copyAndReplaceItem(srcSpreadsheet, dstSpreadsheet,  titleArray, responseArray, imageTitleArray, imageIdArray);
  Logger.log("replace complete.");


  // delete one sheet from dstSpreadsheet
  //var ds = dstSpreadsheet.getActive();
  var sheet = dstSpreadsheet.getSheetByName('工作表1');
  dstSpreadsheet.deleteSheet(sheet);

  // set output directory
  var folderName = "Generate Documents";
  var folder;
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    Logger.log("folder exist");
    folder = folders.next();
  } else {
    Logger.log("generate file folder");
    folder = DriveApp.createFolder(folderName);
  }

  var file = DriveApp.getFileById(dstSpreadsheet.getId());  
  
  // move files
  folder.addFile(file);
  Logger.log("write files complete.");

  // delete temp image files
  for(var it = 0 ; it < imageTitleArray.length ; it++){
    var imageId = imageIdArray[it];    
    var imageFile = DriveApp.getFileById(imageId); // failed to get
    imageFile.setTrashed(true);  
  }
}

function copyAndReplaceItem(srcSpreadsheet, dstSpreadsheet, titleArray, responseArray, imageTitleArray, imageIdArray){

  // get current data for file name
  var currentDate = new Date();
  var currentDayOfMonth = currentDate.getDate();
  var currentMonth = currentDate.getMonth(); 
  var currentYear = currentDate.getFullYear();
  var dateString = currentDayOfMonth + "-" + (currentMonth + 1) + "-" + currentYear;



  var sSheets = srcSpreadsheet.getSheets();
  var name;

  // copy all sheets from srcspreadsheet to datspreadsheet
  for (var i = 0; i < sSheets.length; i++) { 
    var sheet = sSheets[i];
    sheet.copyTo(dstSpreadsheet);
  }

  var dSheets = dstSpreadsheet.getSheets();
  for (var i = 0; i < dSheets.length; i++) { // traverse all sheets
    var dSheet = dSheets[i];
    var dataRange = dSheet.getDataRange(); // select currnt sheet's full ranges
    var values = dataRange.getValues(); // select currnt sheet's full values(2D arrays)
    for (var row = 0; row < values.length; row++) {
      for (var col = 0; col < values[0].length; col++) { // traverse all values
        
        // compare all titles and insert all responses
        for(var t = 0 ; t < titleArray.length ; t++){ 
          if (values[row][col] === "<<" + titleArray[t] + ">>") {
            dSheet.getRange(row + 1, col + 1).setValue(responseArray[t]);

            if(titleArray[t] === "姓名"){
              name = responseArray[t] + dateString;
            }
          }
        }

        // compare all image titles and insert all images
        for(var it = 0 ; it < imageTitleArray.length ; it++){
          
          if (values[row][col] === "<<" + imageTitleArray[it] + ">>") {

            dSheet.getRange(row + 1, col + 1).setValue("");
          
            var imageId = imageIdArray[it];

            // adjust image blob
            var imgTempBlob = DriveApp.getFileById(imageId).getBlob();
            var fileId = DriveApp.createFile(imgTempBlob).getId();
            var link = Drive.Files.get(fileId).thumbnailLink.replace(/\=s.+/, "=s" + idCardWidth);
            var imgBlob = UrlFetchApp.fetch(link).getBlob();
            Drive.Files.remove(fileId);
            
            var img = dSheet.insertImage(imgBlob, col+1, row+1);

            /*var imgBlob = DriveApp.getFileById(imageId).getBlob();
            var img = dSheet.insertImage(imgBlob, col+1, row+1);

            // adjust image width to fit current border
            var imgWidth = img.getWidth();
            var imgHeight = img.getHeight();

            var scale = idCardWidth / imgWidth;

            img.setWidth(imgWidth * scale);
            img.setHeight(imgHeight * scale);*/
   
          }        
        }    
      }
    }
  }
  dstSpreadsheet.rename("勞務單 " + name);
}
