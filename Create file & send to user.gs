function getdata() {

var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1bVEZVuvOtw......./edit#gid=0");
var tab = ss.getSheetByName("D1");
var datavalues = tab.getRange(1, 2,tab.getLastRow(),11).getDisplayValues();
var html = "<br><br>Factures à valider: <br><br> " + '<table>';

if(datavalues.length>0) {
  for(var i=0; i<datavalues.length; i++) {
    html+= '<tr style = bgcolor="Blue">';
    for(var j=0; j<datavalues[i].length; j++) {
      if(i==0) {
        html+= Utilities.formatString('<td bgcolor="LightBlue"; style = "border:1px solid black" <th>%s</th></td>',datavalues[i][j]);
      }else {
        html+= Utilities.formatString('<td style = "border:1px solid black" >%s</td>',datavalues[i][j]);
      }
    }
  }
  html+= '<table>';
}
 return html; 
}



function sendmail() {

  var startRow = 2; 
  var numRows = 700; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("P1"); 
  var dataRange = sheet.getRange(startRow, 1, numRows, 3);
  var data = dataRange.getValues();
  
  
  for (var j = 0; j < data.length; j++) {
    var row = data[j]; 
    var emailAddress = row[2]; 
    if (emailAddress != "") 
    {
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1bVEZVuvOtwWZ9R2eac8wBMjoEU8oY7WuSNVcsj2J7Sc/edit#gid=0");
    var tab = ss.getSheetByName("DATA");
    var tab2 = ss.getSheetByName("D1");
    var sheet = ss.getSheetByName("P1");
    var lastSourceRow = sheet.getLastRow();//get last row    
    var file = DriveApp.getFilesByName('SUPORT_test.pdf');
    
      var tab3 = SpreadsheetApp.openById("1bVEZVu........").getSheetByName("D1").getRange("A2:AI");
      tab3.clearContent();
      var range = tab.getRange(2, 1,tab.getLastRow()-1,35).getValues();
      var currentName = row[0];
      var fdata = range.filter(function(item){ return item[27] === currentName; });
      tab2.getRange(2, 1,fdata.length,fdata[0].length).setValues(fdata);

      
      
      var destination2 = DriveApp.getFolderById('1XpNXeL_W9ekgc........');
      var newFile = SpreadsheetApp.create(currentName).getId();
      destination2.addFile(DriveApp.getFileById(newFile));
      
      var sheet2 = ss.getSheets()[4];
      var destination = SpreadsheetApp.openById(newFile);
      sheet2.copyTo(destination).setName(currentName);
      destination.deleteActiveSheet();
      var newFile_share = DriveApp.getFileById(newFile);
      newFile_share.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      
      var url = destination.getUrl();
      
      
      var alias = GmailApp.getAliases();
      var subject = "Subject";
      var message = "Hello, \r\n\r\n..... ";
      

     GmailApp.sendEmail(emailAddress, subject, message, {attachments: file.next().getAs('application/pdf'), from: alias[0], name: "NAME Name", htmlBody: "Hello, <br><br>Please find .......... <br>TEXT...... <br>TEXT...... <br><br> " + url + "<br><br>Thank you! <br><br> " + getdata()}); 

    }
 
    }

  }
