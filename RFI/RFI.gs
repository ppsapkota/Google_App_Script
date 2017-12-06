/*
 * Global Variables
 */

// Form URL
var formURL = 'https://docs.google.com/forms/d/1HR4qT7i26D4U4oSSQ6ooOoGRpjH5_o6d-4QlJtFTU28/viewform';
// Sheet name used as destination of the form responses
var sheetName = 'Form Responses 1';
/*
 * Name of the column to be used to hold the response edit URLs 
 * It should match exactly the header of the related column, 
 * otherwise it will do nothing.
 */
var columnName = 'Edit Url';
var columnName_ID = 'RFI Number';
var columnName_increment = 'Increment';

var prefillname_ID = 'entry.357851482';
// Responses starting row
var startRow = 2;

//--------------------------------------------------------------------------
function getEditResponseUrls(){
  //Logger.log('sendEmailsapp ran!');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues(); 
  var columnIndex = headers[0].indexOf(columnName);
  var data = sheet.getDataRange().getValues();
  var form = FormApp.openByUrl(formURL);
  //response ID
  var columnIndex_ID = headers[0].indexOf(columnName_ID);
  var responses = form.getResponses();
  var responses_c = responses.length;
  //
  for(var i = startRow-1; i < data.length; i++) {
    if(data[i][0] != '' && data[i][columnIndex] == '') {
    //if(data[i][0] != '') {
      var timestamp = data[i][0];
      var formSubmitted = form.getResponses(timestamp);
      if(formSubmitted.length < 1) continue;
      var editResponseUrl = formSubmitted[0].getEditResponseUrl();   
      // response ID -Punya   
      var id_existing = data[i][columnIndex_ID];
        
        //make the ID
        var yymmdd = Utilities.formatDate(timestamp,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyMMdd");
        var p_id = "0000" + responses_c;
        var id_new = "H-"+ yymmdd + "-" + p_id.substr(p_id.length-4);        
        if (id_existing == ''){
            sheet.getRange(i+1, columnIndex_ID+1).setValue(id_new);
            editResponseUrl = editResponseUrl + "&"+prefillname_ID+"="+id_new;
        } else
        {
          editResponseUrl = editResponseUrl + "&"+prefillname_ID+"="+id_existing;
        }
      //make response URL
      //https://docs.google.com/forms/d/e/1FAIpQLSfsyPYxu-QBjWI9G9xVvfXG_kTWo5QrzQttiEzNH2KwumYiBQ/viewform?edit2=2_ABaOnud8RvP7GhRnj6UGis4MHqYI9_Xo6XOhXfWxJkmYe7bbKHiHLYUCUJb3Jw&entry.357851482=H30
      sheet.getRange(i+1, columnIndex+1).setValue(editResponseUrl);
      //Logger.log('sendEmailsapp ran!'+editResponseUrl);
      //  
    }      
  } //for
} //function

//--------------------------------------------------------------------------
function getEditResponseUrl_forActiveRow(){
  var activeSheet = SpreadsheetApp.getActiveSheet();
  //var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  //var data = activeSheet.getDataRange().getValues();
  var form = FormApp.openByUrl(formURL);
  //Response count
  var responses = form.getResponses();
  var responses_c = responses.length;
  //
  var numberOfColumns = activeSheet.getLastColumn();
  var activeRowIndex = activeSheet.getActiveRange().getRowIndex();
  var activeRow = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues();
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues();
  var columnIndex = 0; //timestamp
  //
  //response ID
  //var columnIndex_ID = headers[0].indexOf(columnName_ID);//RFI Number
  var columnIndex_ID = headerRow[0].indexOf(columnName_ID);
  var columnIndex_EditUrl = headerRow[0].indexOf(columnName);//Edit Url
  var timestamp_val = activeRow[0][0];
  var editurl_val = activeRow[0][columnIndex_EditUrl];
  //check timestamp data and edit url data
    if (timestamp_val != '' && editurl_val == '') {
      var formSubmitted = form.getResponses(timestamp_val);
      var editResponseUrl = formSubmitted[0].getEditResponseUrl();   
      // response ID - RFI Number
      var id_existing = activeRow[0][columnIndex_ID];
      //make the ID
      var yymmdd = Utilities.formatDate(timestamp_val,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyMMdd");
      var pad_id = "0000" + responses_c;
      var id_new = "H-"+ yymmdd + "-" + pad_id.substr(pad_id.length-4);        
      if (id_existing == ''){
            activeSheet.getRange(activeRowIndex, columnIndex_ID+1).setValue(id_new);
            editResponseUrl = editResponseUrl + "&"+prefillname_ID+"="+id_new;
        } else
        {
          editResponseUrl = editResponseUrl + "&"+prefillname_ID+"="+id_existing;
        }
      //make response URL
      //https://docs.google.com/forms/d/e/1FAIpQLSfsyPYxu-QBjWI9G9xVvfXG_kTWo5QrzQttiEzNH2KwumYiBQ/viewform?edit2=2_ABaOnud8RvP7GhRnj6UGis4MHqYI9_Xo6XOhXfWxJkmYe7bbKHiHLYUCUJb3Jw&entry.357851482=H30
      activeSheet.getRange(activeRowIndex, columnIndex_EditUrl+1).setValue(editResponseUrl);
      //
    } //if     
  
} //function



//------------------------PDF Creation-------------
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('RFI Records')
    .addItem('Create PDF (selected row)', 'createPdf')
    .addItem('Get Edit Url (selected row)', 'getEditResponseUrl_forActiveRow')
    .addToUi();
} // onOpen()

/**  
 * Take the fields from the active row in the active sheet
 * and, using a Google Doc template, create a PDF doc with these
 * fields replacing the keys in the template. The keys are identified
 * by having a % either side, e.g. %Name%.
 *
 * @return {Object} the completed PDF file
 */
// dev: andrewroberts.net

// Replace this with ID of your template document.
var TEMPLATE_ID = '1P4QQyS8MuyLaUH4BAstuaMG4PkVxl71-K6Hul7gL4x8';
//https://docs.google.com/document/d/1P4QQyS8MuyLaUH4BAstuaMG4PkVxl71-K6Hul7gL4x8/edit
//https://docs.google.com/document/d/1P4QQyS8MuyLaUH4BAstuaMG4PkVxl71-K6Hul7gL4x8/edit?usp=drive_web - 
// var TEMPLATE_ID = '1wtGEp27HNEVwImeh2as7bRNw-tO4HkwPGcAsTrSNTPc' // Demo template
// Demo script - http://bit.ly/createPDF 
// You can specify a name for the new PDF file here, or leave empty to use the 
// name of the template.
var PDF_FILE_NAME = ''; //dynamically generated
var PDF_FOLDER_NAME = 'RFI_PDF_Output';
var PDF_dir_id = '0B9RbVk3syuU0ZGQ2YUdka0k0RGc';
//https://drive.google.com/drive/folders/0B9RbVk3syuU0ZGQ2YUdka0k0RGc?usp=sharing
var dir = DriveApp.getFoldersByName(PDF_FOLDER_NAME).next();

//SpreadsheetApp.getUi().alert('dir--'+dir)
/**
 * Eventhandler for spreadsheet opening - add a menu.
 */

function createPdf() {

  if (TEMPLATE_ID == '') {    
    SpreadsheetApp.getUi().alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }

  // Set up the docs and the spreadsheet access
  var template_file = DriveApp.getFileById(TEMPLATE_ID);
  var copyFile = template_file.makeCopy(dir);  
    
  var copyId = copyFile.getId();
  var copyDoc = DocumentApp.openById(copyId);
  var copyBody = copyDoc.getActiveSection();
  
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var numberOfColumns = activeSheet.getLastColumn();
  var activeRowIndex = activeSheet.getActiveRange().getRowIndex();
  var activeRow = activeSheet.getRange(activeRowIndex, 1, 1, numberOfColumns).getValues();
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues();
  var columnIndex = 0; //timestamp
   
  //check if selection is empty row
////  if (activeRow[0][0]===''){   
//      SpreadsheetApp.getActiveSpreadsheet().toast('Select the row that has data.','Something went wrong...');
//      return;
//      }
//  
  
  // Replace the keys with the spreadsheet values 
//  for (;columnIndex < headerRow[0].length; columnIndex++) {    
//    copyBody.replaceText('%' + headerRow[0][columnIndex] + '%', 
//                         activeRow[0][columnIndex])                         
//  }
    //SpreadsheetApp.getUi().alert(copyId);
  //Punya  
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);   
  //Column Names to export
  
  //Increment
  var col_increment = headerRow[0].indexOf(columnName_increment); 
  var increment = activeRow[0][col_increment]+1;
  activeSheet.getRange(activeRowIndex, col_increment+1).setValue(increment);
  //process for the rest
  var col_index = headerRow[0].indexOf('Timestamp'); 
  var val_timestamp = activeRow[0][col_index];
  var yymmdd = Utilities.formatDate(val_timestamp,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd HH:mm:ss");
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',yymmdd); 
  //SpreadsheetApp.getUi().alert(yymmdd);
  
  //
  var col_index = headerRow[0].indexOf('RFI Number'); 
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]);
  PDF_FILE_NAME = activeRow[0][col_index];
  
  //SpreadsheetApp.getUi().alert(PDF_FILE_NAME);
  
  var col_index = headerRow[0].indexOf('Priority');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]);  
   
  var col_index = headerRow[0].indexOf('Organisation');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Initiator');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  PDF_FILE_NAME = PDF_FILE_NAME+'-'+activeRow[0][col_index]+'-'+increment;
  
  var col_index = headerRow[0].indexOf('Others asked');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Purpose');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Category');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Background');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Request for information');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Response by');
  val_timestamp = activeRow[0][col_index];
  //check if datee is provided or not
  if (val_timestamp != ''){
      yymmdd = Utilities.formatDate(val_timestamp,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
  }else {
    yymmdd='';
  }
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',yymmdd); 
  
  var col_index = headerRow[0].indexOf('Response Details');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Response Status');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  
  var col_index = headerRow[0].indexOf('Edit Url');
  copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  //SpreadsheetApp.getUi().alert(col_index);
    
  // Create the PDF file, rename it if required and delete the doc copy
  copyDoc.saveAndClose();
  
  //SpreadsheetApp.getUi().alert(dir);
  //var file = dir.createFile(name, content);
  var newFile = dir.createFile(copyFile.getAs('application/pdf'));  
  
  //response file name  
  if (PDF_FILE_NAME != '') {  
    newFile.setName(PDF_FILE_NAME);
  } 
  // pdf url
  var PDF_URL = 'https://drive.google.com/file/d/'+newFile.getId()+'/view?usp=drivesdk'
  var col_index = headerRow[0].indexOf('View PDF');
  //SpreadsheetApp.getUi().alert(newFile.getId());  
  //copyBody.replaceText('%' + headerRow[0][col_index] + '%',activeRow[0][col_index]); 
  activeSheet.getRange(activeRowIndex, col_index+1).setValue(PDF_URL);
  copyFile.setTrashed(true);
  
  //SpreadsheetApp.getUi().alert('New PDF file::"'+PDF_FILE_NAME+'" created in "' + PDF_FOLDER_NAME + '" folder of your Google Drive');
  
} // createPdf()


//--------END PDF Creation--------------------------------------







function pad(num, size) {
    var s = "0000" + num;
    return s.substr(s.length-size);
}

//function GenerateID(){
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
//  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues(); 
//  var columnIndex_ID = headers[0].indexOf(columnName_ID);
//  var data = sheet.getDataRange().getValues();
//  var form = FormApp.openByUrl(formURL);  
//  var responses = form.getResponses();
//  var responses_c = responses.length;
//  // get active row id or cursor row id
//  //
//  
//  var r = sheet.getActiveCell().getRow();
//  var id_existing = data[r-1][columnIndex_ID];
//  var timestamp = data[r-1][0];
//  //
//  var yymmdd = Utilities.formatDate(timestamp,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyMMdd");
//  var id_new = responses_c;
//  SpreadsheetApp.getActiveSpreadsheet().toast(yymmdd," Old Total submissions",-1); 
////  //SpreadsheetApp.getActiveSpreadsheet().toast(id_new,"NEW Total submissions",-1); 
////  // check whether value is already there or not
////  if (id_existing == '')
////  {
////      sheet.getRange(r, columnIndex_ID+1).setValue(id_new);
////  }  
//}

/*
function onEdit_getEditResponseUrls()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues(); 
  var columnIndex1 = headers[0].indexOf(columnName_TEST);
  var data = sheet.getDataRange().getValues();
  var form = FormApp.openByUrl(formURL);
  //response ID
  var columnIndex_ID = headers[0].indexOf(columnName_ID);
  var responses = form.getResponses();
  var responses_c = responses.length;
  //get active row index
  var r = sheet.getActiveCell().getRow();
  //SpreadsheetApp.getActiveSpreadsheet().toast(r," Old Total submissions",-1);
  var timestamp = data[r-1][0];
  var formSubmitted = form.getResponses(timestamp);
  var editResponseUrl = formSubmitted[0].getEditResponseUrl();
  sheet.getRange(r, columnIndex1+1).setValue(editResponseUrl);
  //SpreadsheetApp.getActiveSpreadsheet().toast(editResponseUrl," Old Total submissions",-1); 
} //function

function OnForm(){
    var sh = SpreadsheetApp.getActiveSheet()
    var startcell = sh.getRange('A2').getValue();
      if(! startcell){sh.getRange('A2').setValue('PQOT-14-0001');return}; // this is only to       handle initial situation when A1 is empty.
    var colValues = sh.getRange('A2:A').getValues();// get all the values in column A in an array
    var year = Utilities.formatDate(new Date, "PST", "yy")
    var max=0;// define the max variable to a minimal value
        for(var r in colValues){ // iterate the array
        var vv=colValues[r][0].toString().replace(/[^0-9]/g,'');// remove the letters from the string to convert to number
          if(Number(vv)>max){max=vv};// get the highest numeric value in th column, no matter what happens in the column... this runs at array level so it is very fast
      }
    max++ ; // increment to be 1 above max value
    sh.getRange(sh.getLastRow(), 1).setValue(Utilities.formatString('PQOT-'+year+'-%04d',max));// and write it back to sheet's last row.
     
    
}
var vv=colValues[r][0].toString().substr(-4).replace(/\D/g,'')
*/


//
//pad.zeros = new Array(5).join('0');
//function pad_alternate(num, len) {
//    var str = String(num),
//        diff = len - str.length;
//    if(diff <= 0) return str;
//    if(diff > pad.zeros.length)
//        pad.zeros = new Array(diff + 1).join('0');
//    return pad.zeros.substr(0, diff) + str;
//}

