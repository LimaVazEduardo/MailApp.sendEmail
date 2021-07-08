# MailApp.sendEmail
Cool script to send e-mails from a Google Spreadsheet

Greetings!

Here is a cool script to send standard emails using Google Spreadsheet.

### Problem to solve

Let's say you work for a local school and frequently need to send e-mails to the account department, authorizing some discounts.
You could use a spreadsheet to create a standard output mail like the one below, in your email system:

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/mail_body.png)

### Suggested solution

First let's set up a new spreadsheet just like the one below:  
Note we will also create a new drop down menu called: "*Discount Request*"

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/sheet.png)

At the same spreadsheet, navigate to column **Z** and create the e-mail settings just like figure below:

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/email_setup.png)

Go to menu *Tools* and choose *Script editor*

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/script_editor.png)

You will see a screen like this one:  
Give the project a name: *Discount_request_email*

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/script.png)

Copy and paste the script `code.gs` of this repository to your `code.gs` file.

### How the script works

In the first part of the script, we will create a menu in the spreadsheet:

```
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

```

Then, we need to create variables to store the components of a traditional e-mail, like:
  - To
  - Cc
  - Subject
  - Body
  - Signature

All those values will be fetched from the spreadsheet and could carry any value you may need.

```
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

```

I know it is a nightmare to work with *dates* in JavaScript.  
Hope these lines will ease your pain.

```
var ts = new Date().toLocaleString(undefined, {
    day:   'numeric',
    month: 'short',
    year:  'numeric',
    hour:   '2-digit',
    minute: '2-digit',
    second: '2-digit',
}
```

Since we may have more than one carbon copy recipients, we need to test the existing of all 3 possibilities:  
*If you know a better way to do this, let me know in the comments* :)

```
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
```

Fetching the others important variables:

```
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

```

Now, comes the best part, actually send off the email:  
We will use the htmlBody option in this script.

```
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
```
See Google documentation about send email
[Class MailApp](https://developers.google.com/apps-script/reference/mail/mail-app)

And at the last part, we would like to have a timestamp signature of the email sent.  
Here is how to do it.

```
  //MailApp.sendEmail(emailAddress, subject, message);
  sheet.getRange(row, 8).setValue("Sent: " + ts).setFontColor('red'); // H
  // Make sure the cell is updated right away in case the script is interrupted
  SpreadsheetApp.flush();
}
```
### How to send and e-mail

After filling out the columns **A** until **G**, type the row number of the corresponding student name with the approved discount, at cell **K1**.  
Go to menu *Discount Request* and choose *Send e-mail* option.

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/send_email.png)

*The first time you run this script, you will need to allow the script to access your spreadsheet.*

### Final notes:
This is what I learned using Google scripts.  
It is possible to use Google spreadsheets to send standardized emails for enhanced communication.  
You could also insert a logic to send automatically e-mails, case some conditions are met.

Hope this script helps you at work.

Let me know your comments.

Best regards,

[Eduardo Lima](https://www.linkedin.com/in/eduardo1lima/)
