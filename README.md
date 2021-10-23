Second version:  
# MailApp.sendEmail
Cool script to send e-mails from a Google Spreadsheet

Greetings!

Here is a cool script to send standard emails using Google Spreadsheet.

### Problem to solve

Let's say you work for a local school and frequently need to send e-mails to the account department, 
authorizing some discounts.  
Every month you need to calculate a discount value and assembly an email to be sent to your organization 
with student name, parent name and other pieces of information.   

You also need to keep a log of all discounts approved.

#### This is what you do every month

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/mail_body.png)

### Suggested solution

Create a simple system to send standardized emails to pre-difened addresses.

First let's set up a new Google spreadsheet just like the one below:  
Note we will also create a new drop down menu called: "*Discount Request*"

![sheet](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/sheet.png)

\
\
At the same spreadsheet, create a new tab and call it **Settings**:  
Insert the following values in column A.
 - To:
 - Cc:   *you may separate e-mail addresses using "," commas*
 - Subject:
 - Body1:
 - Body2:
 - Body3:
 - Signature1:
 - Signature2: 

`The cells Body1, Body2 and Body3 will help you to update the email body without having to edit the script.`

\
\
![email_setup](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/email_setup.png)

\
\
Go to menu *Tools* and choose *Script editor*

![script_editor](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/script_editor.png)

\
\
You will see a screen like this one:  
Give the project a name: **Discount_request_email**

![script](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/script.png)

Copy and paste the script `code.gs` of this repository to your `code.gs` file.

### How the script works

In the first part of the script, we will create a menu in the spreadsheet and also the checkboxes:

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
  //Insert 22 Checkboxes at column I (9)
  for(var i = 2; i <= 22; i++){
    var cb = SpreadsheetApp.getActive().getSheetByName('Send_mail');
    cb.getRange(i,9).insertCheckboxes();
  }
}

function about(){
  SpreadsheetApp.getUi().alert("Discount E-mail Request Script\nLast updated July, 2021\n\nEduardo Lima ©\nlima.vaz.eduardo@gmail.com");
}

```
\
\
Then, we need to create variables to access the components of both tabs, *Send_mail* and *Settings*:

```
//=====================================
// Main script
//=====================================
function send_discount_email() {
  var ss       = SpreadsheetApp.getActive();
  var sheet    = ss.getSheetByName('Send_mail');
  var settings = ss.getSheetByName('Settings');

```
![tabs](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/tabs.png)

\
\
After sending the email, the script will log a timestamp in the spreadsheet to indicate the e-mail was sent.

I know it is a nightmare to work with *dates* in JavaScript.  
Hope these lines will ease your pain.

```
var ts = new Date().toLocaleString(undefined, {
    day:    'numeric',
    month:  'short',
    year:   'numeric',
    hour:   '2-digit',
    minute: '2-digit',
    second: '2-digit',
}
```

\
\
At this part of the script, we are checking which *Checkbox* is marked as TRUE or FALSE.
 - You may extend this check over the first 22 cells.

If none of the checkboxes at column *I* are marked, a pop up window will alert you to select one student.
If one checkboxe at column *I* is marked, the script will check if there is a student name assigned to column *A*.
  If not, a pop up window will alert you of an error.
  If yes, variable *row* will be assigned with the row number of the student you have selected.

```
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
```
\
\
Fetching email values from tab "Settings":

```
  var to = settings.getRange('B1').getValue();
  var cc = settings.getRange('B2').getValue();
    
  var subject    = settings.getRange('B3').getValue();
  var body1      = settings.getRange('B4').getValue();
  var body2      = settings.getRange('B5').getValue();
  var body3      = settings.getRange('B6').getValue();
  var signature1 = settings.getRange('B7').getValue();
  var signature2 = settings.getRange('B8').getValue();
  
```
\
\
Fetching student details:

```
 var student         = sheet.getRange(row, 1).getValue(); //A2
 var parent          = sheet.getRange(row, 2).getValue(); //B2
 var grade           = sheet.getRange(row, 3).getValue(); //C2
 var percent         = sheet.getRange(row, 4).getValue(); //D2
 var installment     = sheet.getRange(row, 5).getValue(); //E2
 var school_supplies = sheet.getRange(row, 6).getValue(); //F2
 var total_amount    = sheet.getRange(row, 7).getValue(); //G2
 console.log("cc: ", cc);
 
```
\
\
In order to have the correct currency format for your country, we are going to set variable *myObj*.

Example: `R$ 1.000,00`

```
  var myObj = {
    style: "currency",
    currency: "BRL"
  }
  
```
See more details at: [W3Schools](https://www.w3schools.com/jsref/jsref_tolocalestring_number.asp)

\
\
Now, comes the best part, actually send off the email:  
We will use the htmlBody option in this script.

```
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
  
```

\
\
See Google documentation about send email
[Class MailApp](https://developers.google.com/apps-script/reference/mail/mail-app)

\
\
And at the last part, we would like to have a timestamp signature of the email sent.  
Here is how to do it.

```
  //MailApp.sendEmail(emailAddress, subject, message);
  sheet.getRange(row, 8).setValue("Sent: " + ts).setFontColor('red'); // H
  // Make sure the cell is updated right away in case the script is interrupted
  SpreadsheetApp.flush();
}

```

\
\
Finally, it is a good idea to clear the checkbox so we avoid confusion for the user:

```
 // Clear Check Box of column I.
  sheet.getRange(row, 9).setValue(false);

```

### How to send and e-mail

After filling out the columns **A** until **G**, with the student details, check the corresponding Checkbox at column **I**.

Go to menu *Discount Request* and choose *Send e-mail* option.

![send_email](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/send_email.png)

*The first time you run this script, you will need to allow the script to access your spreadsheet.*


### E-mail log

After sending the e-mail, a timestamp will be logged at column *H* and the checkbox will be cleared.

![email_rachel](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/email_rachel.png)



### Final notes:
This is what I have learned using Google scripts.  
It is possible to use Google spreadsheets to send standardized emails for enhanced communication.  
You could also insert a logic to send automatically e-mails, case some conditions are met.

Hope this script helps you at work.

Let me know your comments.

Best regards,

[Eduardo Lima](https://www.linkedin.com/in/eduardo1lima/)

[Twitter](https://twitter.com/Eduardo69564454)
