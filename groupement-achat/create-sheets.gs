

function getEmail(contact){
  return contact.getEmails()[0].getAddress();
}

function copySpreadsheet(contact, spreadSheet, folderId, dryRun) {
  
  Logger.log("Creating spreadsheet for user %s",  contact.getFullName());
  
  if(dryRun){
    Logger.log("DRY_RUN: would create spreadSheet")
    return "fakeUrl"
  }
  
  var userBDC = originalBDC.copy("Bon de commande "+contact.getFullName())
  userBDC.getRange("E2").setValue(contact.getFamilyName())
  userBDC.getRange("E3").setValue(contact.getGivenName())
  userBDC.getRange("E6").setValue(getEmail(contact))
  var fileOfUser = DriveApp.getFileById(userBDC.getId())
  fileOfUser.moveTo(DriveApp.getFolderById(folderId))
  fileOfUser.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT)
  Logger.log(fileOfUser.getUrl())
  
  Logger.log("Created spreadsheet %s",  fileOfUser.getId());
  return fileOfUser.getUrl()
}

function sendEmail(contact, url, dryRun) {
 
  var email = getEmail(contact)
  Logger.log("Sending email with url to %s", email)
  
  var message = ""
  message += "Bonjour, \n"
  message += "ci dessous vous trouverez le lien vers votre bon de commande personnel.\n"
  message += "\n"
  message += url+"\n"
  message += "\n"
  message += "Remplissez simplement le bon de commande avant la date limite et procedez au paiement.\n"
  message += "\n"
  message += "Merci!\n"
  message += "\n"
      
  if(dryRun){
    Logger.log("DRY_RUN: would send email: %s", message)
  } else {
    MailApp.sendEmail(email, 
                      "Commande groupement d'achat la Marm'Hotte", 
                      message
                     )
  }
  
}


function run(){

  var dryRunSpreasheet = true
  var dryRunEmail = true
  
  var contactGroupName = "[TEMP] Groupement achat"
  var originalSpreadsheetId = "enter spreadsheet id here"
  var folderForSpreadsheetStorageId = "enter folder id here"
  
  
  var originalSpreadsheet = SpreadsheetApp.openById(originalSpreadsheetId)
  var contactGroup = ContactsApp.getContactGroup(contactGroupName)
  
  var allContacts = contactGroup.getContacts()
  
  Logger.log("Will send email for %s users, quota is: %s", allContacts.length, MailApp.getRemainingDailyQuota())
  allContacts.forEach(function(c){
    
    
    var url = copySpreadsheet(c, originalSpreadsheet, folderForSpreadsheetStorageId, dryRunSpreasheet)
    sendEmail(c, url, dryRunEmail)
    
  })
  
  
}