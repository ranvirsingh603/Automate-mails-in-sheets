function sendmail() {
  const sheet = SpreadsheetApp.getActiveSheet()
  var last= sheet.getLastRow();
  var data2=sheet.getRange(2,1,last,3).getValues();
  var len = data2.filter(function(value){return value[0]}).length
  for (var i = 0; i < len; ++i){
    var row = data2[i]
    var email = row[0]
    var name = row[1]
    var sub = row[2]
    var message = "Hello "+name+",\n\nHow are you.\n\nThis is the test for sending automatic mails"
    try{
      GmailApp.sendEmail(email, sub, message);
    } catch(e){}
  
  }
}
