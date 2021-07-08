# MailApp.sendEmail
Cool script to send e-mails from a Google Spreadsheet

Greetings!

Here is a cool script to send standard emails using Google Spreadsheet.

### Problem to solve

Let's say you work for a local school and frequently need to send e-mails to the account department, authorizing some discounts.
You could use a spreadsheet to create a standard output like the one below, in your email system:

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

![mail body](https://github.com/LimaVazEduardo/MailApp.sendEmail/blob/main/script.png)

Copy and paste the script code.gs

### How it works

The first part of the script we will create a menu in the spreadsheet:

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





