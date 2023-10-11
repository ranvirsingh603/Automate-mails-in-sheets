# Sending Emails Automatically in Google Sheets 📊

![Google Sheets Logo](https://www.gstatic.com/images/branding/product/1x/sheets_48dp.png)

## **Step 1: Open Google Sheets 📝**

Begin by launching Google Sheets on your computer.

## **Step 2: Create a Sheet with Email Information 📬**
1. In your Google Sheets, dedicate the first column for **email addresses**.
2. Use the second column for **names**.
3. Reserve the third column for **mail subject lines**.

## **Step 3: Create a Drawing or Icon 🎨**
1. Select any cell within your spreadsheet.
2. Go to the **Insert** menu and choose **Drawing**.
3. Craft a unique and imaginative icon or image.
4. When you're done, click **Save and Close** to insert it into your sheet.
<div style="float: right; padding-left: 20px;">
  <img src="https://github.com/ranvirsingh603/Automate-mails-in-sheets/blob/main/Screenshot%202023-10-11%20125541.png" alt="Google Sheets Logo" height="150">
</div> 

## **Step 4: Assign a Script to the Icon 📜**
1. Click on the drawing icon you've just created.
2. In the top-right corner, click on the three dots (`...`).
3. Choose **Assign Script** from the dropdown menu.
4. Name your script; let's call it "sendmail."
<div style="float: right; padding-left: 20px;">
  <img src="https://github.com/ranvirsingh603/Automate-mails-in-sheets/blob/main/Screenshot%202023-10-11%20125713.png" alt="Google Sheets Logo" height="150">
</div> 
That's it! You've successfully assigned the "sendmail script, which will run when you open the sheet. ✨

## **Step 5: Google Apps Script - Sending Automated Emails 📃**

This line defines a JavaScript function named sendmail(). Functions in JavaScript are blocks of reusable code.
```javascript
function sendmail() {
```
***
This line declares a constant variable named sheet and assigns it the value of the active spreadsheet using SpreadsheetApp.getActiveSheet(). It represents the currently open Google Sheets spreadsheet.
```javascript
const sheet = SpreadsheetApp.getActiveSheet();
```
***
Here, a variable named last is declared and set to the last row number that contains data in the sheet using getLastRow(). This is used to determine how many rows to process.
```javascript
  var last = sheet.getLastRow();
```
---
This line declares a variable data2 and retrieves the values in the range starting from row 2 and column 1 (A) to the last row and column 3 (C) in the spreadsheet. This range includes email addresses, names, and email subjects for each recipient.
```javascript
  var data2 = sheet.getRange(2, 1, last, 3).getValues();
```
---

Here, a variable len is calculated. It filters the data2 array to count the number of rows where the value in the first column (index 0) is non-empty or truthy. This gives you the number of recipients to process.
```javascript
  var len = data2.filter(function(value) { return value[0]; }).length;
```
---
This line initializes a for loop. It starts a loop where i represents the current iteration from 0 to len - 1.
```javascript
  for (var i = 0; i < len; ++i) {
```
---
Inside the loop, a variable row is created, which contains the data (email, name, and subject) for the current recipient, based on the ith row from data2.
```javascript
    var row = data2[i];
```
---
This section extracts the specific data from the row variable. It assigns the email address to the email variable, the recipient's name to the name variable, and the email subject to the sub variable.
```javascript
    var email = row[0];
    var name = row[1];
    var sub = row[2];
```
---
Here, a variable message is created, which is a string containing the email message. It greets the recipient by name, asks how they are, and provides a test message.
```javascript
    var message = "Hello " + name + ",\n\nHow are you.\n\nThis is the test for sending automatic mails";
```
---
This section sends the email using GmailApp.sendEmail(). It uses the email variable for the recipient's email address, the sub variable for the email subject, and the message variable for the email content.
A try...catch block is used to handle any errors that may occur during email sending. If an error occurs, it is caught and not explicitly handled in this script.
In summary, this script is designed to send personalized emails to a list of recipients stored in a Google Sheets spreadsheet. It extracts the email, name, and subject from the spreadsheet and sends the email message with a test content to each recipient. It uses a loop to iterate through the list of recipients and catches any potential errors that occur during email sending.
```javascript
    try {
      GmailApp.sendEmail(email, sub, message);
    } catch (e) {
    }
```
---
Save the Script

## **Step 6: Sending mails 📨**
Click on the **SEND MAIL** icon we created earlier this will run the script and automatically send mails.

### **CONGRATULATIONS!**
**Ranvir Singh Matharu**

