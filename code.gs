//=====================================
// Custom menu
//=====================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Discount Request')
  .addItem('Send e-mail', 'send_discount_email')
      .addItem('About', 'about')
      .addToUi();
  //Insert 22 Checkboxes at column I (9)
  for(var i = 2; i <= 22; i++){
    var cb = SpreadsheetApp.getActive().getSheetByName('Send_mail');
    cb.getRange(i,9).insertCheckboxes();
  }
}

function about(){
  SpreadsheetApp.getUi().alert("Discount Request Script\nLast updated July, 2021\n\nEduardo Lima ©\nlima.vaz.eduardo@gmail.com");
}

//=====================================
// Main script
//=====================================
function send_discount_email() {
  var ss       = SpreadsheetApp.getActive();
  var sheet    = ss.getSheetByName('Send_mail');
  var settings = ss.getSheetByName('Settings');
  var ts = new Date().toLocaleString(undefined, {
    day:    'numeric',
    month:  'short',
    year:   'numeric',
    hour:   '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
  
  // Column I(9): Checkboxes
  for(var i = 2; i <= 22; i++){
    // Checkbox is FALSE
    if(sheet.getRange(i, 9).getValue() == false){ 
      // Last FALSE Checkbox row #22
      if(i == 22){
        SpreadsheetApp.getUi().alert("Select 01 student on column 'I'");
        return;
      }
    }
    // Checkbox is TRUE
    if((sheet.getRange(i, 9).getValue() == true)){  
      // Check if column A is empty
      if(sheet.getRange(i, 1).getValue() == ""){ 
        SpreadsheetApp.getUi().alert("Error, row: " + i + " must not be empty");
        // Clear Checkbox of column I.
        sheet.getRange(i, 9).setValue(false);
        return;
      }
       var row = i;
       break;
      } 
  console.log('i: ', i);
  console.log('row: ', row);
  } 
    
  var to = settings.getRange('B1').getValue();
  var cc = settings.getRange('B2').getValue();
    
  var subject    = settings.getRange('B3').getValue();
  var body1      = settings.getRange('B4').getValue();
  var body2      = settings.getRange('B5').getValue();
  var body3      = settings.getRange('B6').getValue();
  var signature1 = settings.getRange('B7').getValue();
  var signature2 = settings.getRange('B8').getValue();
  
  var student         = sheet.getRange(row, 1).getValue(); //A2
  var parent          = sheet.getRange(row, 2).getValue(); //B2
  var grade           = sheet.getRange(row, 3).getValue(); //C2
  var percent         = sheet.getRange(row, 4).getValue(); //D2
  var installment     = sheet.getRange(row, 5).getValue(); //E2
  var school_supplies = sheet.getRange(row, 6).getValue(); //F2
  var total_amount    = sheet.getRange(row, 7).getValue(); //G2
  console.log("cc: ", cc);
  
  var myObj = {
    style: "currency",
    currency: "BRL"
  }
  
   MailApp.sendEmail({
    to: to,
    cc: cc,
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
     "Payment:    " + installment.toLocaleString("pt-BR", myObj)  + "<br>" +
     "Materials:  " + school_supplies.toLocaleString("pt-BR", myObj) + "<br>" +
     "Total:      " + total_amount.toLocaleString("pt-BR", myObj) + "<br>" +
     "</ul></pre>" + 
     
     "<br>" + 
     "<b>" + signature1 + "</b>" +  "<br>" +
     "<i>" + signature2 + "</i>"
  });
  
  
  //MailApp.sendEmail(emailAddress, subject, message);
  sheet.getRange(row, 8).setValue("Sent: " + ts).setFontColor('red'); // H
  // Make sure the cell is updated right away in case the script is interrupted
  SpreadsheetApp.flush();
  
  // Clear Check Box of column I.
  sheet.getRange(row, 9).setValue(false);
  
}
