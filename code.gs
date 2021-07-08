//=====================================
// Custom menu
//=====================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Discount Request')
  .addItem('Send e-mail', 'send_discount_email')
      .addItem('About', 'about')
      .addToUi();
}

function about(){
  SpreadsheetApp.getUi().alert("Discount E-mail Request Script\nLast updated July, 2021\n\nEduardo Lima ©\nlima.vaz.eduardo@gmail.com");
}

//=====================================
// Main script
//=====================================
function send_discount_email() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Send_mail');
  var ts = new Date().toLocaleString(undefined, {
    day:   'numeric',
    month: 'short',
    year:  'numeric',
    hour:   '2-digit',
    minute: '2-digit',
    second: '2-digit',
});
  
  var row = sheet.getRange(1, 11).getValue(); //L2
  
  var to  = sheet.getRange('AA1').getValue();
  var cc1 = sheet.getRange('AA2').getValue();
  var cc2 = sheet.getRange('AA3').getValue();
  var cc3 = sheet.getRange('AA4').getValue();
  // cc1 = T, cc2 = T, cc3 = T
  if(cc1 != "" && cc2 != "" && cc3 != ""){
    var cc_list = cc1 + "," + cc2 + "," + cc3;
  }
  // cc1 = T, cc2 = F, cc3 = T
  if(cc1 != "" && cc2 == "" && cc3 != ""){
    var cc_list = cc1 + "," + cc3;
  }
  // cc1 = T, cc2 = T, cc3 = F
  if(cc1 != "" && cc2 != "" && cc3 == ""){
    var cc_list = cc1 + "," + cc2;
  }
  // cc1 = F, cc2 = T, cc3 = T
  if(cc1 == "" && cc2 != "" && cc3 != ""){
    var cc_list = cc2 + "," + cc3;;
  }
  // cc1 = F, cc2 = F, cc3 = T
  if(cc1 == "" && cc2 == "" && cc3 != ""){
    var cc_list = cc3;
  }
  // cc1 = F, cc2 = F, cc3 = F
  if(cc1 == "" && cc2 == "" && cc3 == ""){
    var cc_list = "";
  } 
  
  var subject = sheet.getRange('AA5').getValue();
  var body1 = sheet.getRange('AA6').getValue();
  var body2 = sheet.getRange('AA7').getValue();
  var body3 = sheet.getRange('AA8').getValue();
  var signature1 = sheet.getRange('AA9').getValue();
  var signature2 = sheet.getRange('AA10').getValue();
  
  var student = sheet.getRange(row, 1).getValue(); //A2
  var parent =  sheet.getRange(row, 2).getValue(); //B2
  var grade =  sheet.getRange(row, 3).getValue(); //C2
  var percent =  sheet.getRange(row, 4).getValue(); //D2
  var installment =  sheet.getRange(row, 5).getValue(); //E2
  var school_supplies =  sheet.getRange(row, 6).getValue(); //F2
  var total_amount =  sheet.getRange(row, 7).getValue(); //G2
  Logger.log("cc_list: ", cc_list);
  
  
   MailApp.sendEmail({
    to: to,
    cc: cc_list,
    subject: subject + " " + student,
    htmlBody: "<img src='https://cdn.pixabay.com/photo/2017/05/23/19/42/seal-2338306__480.png' alt='E-mail' style='width:130px;height:60px;'>" + 
     "<br>" +
     body1 + "<br>" + 
     "<pre><ul>" +
     "Student: " + student + "<br>" +
     "Parent:  " + parent + "<br>" +
     grade +"º grade" +  "<br>" +
     "</ul></pre>" +
     body2 + " " + percent*100 + "%" + "<br>" + 
     body3 +  "<br>" + 
     "<pre><ul>" +
     "Payment:    $ " + installment.toFixed(2)  + "<br>" +
     "Materials:  $ " + school_supplies.toFixed(2) + "<br>" +
     "Total:      $ " + total_amount.toFixed(2) + "<br>" +
     "</ul></pre>" + 
     
     "<br>" + 
     "<b>" + signature1 + "</b>" +  "<br>" +
     "<i>" + signature2 + "</i>"
  });
  
  
  //MailApp.sendEmail(emailAddress, subject, message);
  sheet.getRange(row, 8).setValue("Sent: " + ts).setFontColor('red'); // H
  // Make sure the cell is updated right away in case the script is interrupted
  SpreadsheetApp.flush();
}
