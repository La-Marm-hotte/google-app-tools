

function getEmail(contact){
  return contact.getEmails()[0].getAddress();
}

function copySpreadsheet(contact, spreadSheet, folderId, dryRun) {
  
  Logger.log("Creating spreadsheet for user %s",  contact.getFullName());
  if(dryRun){
    Logger.log("DRY_RUN: would create spreadSheet")
    return "fakeUrl"
  }
  
  var numeroAdherant = contact.getAddresses()[0].getLabel()
  var userBDC = spreadSheet.copy("Bon de commande - "+numeroAdherant+" - "+contact.getFullName())
  userBDC.getRange("E2").setValue(contact.getFamilyName())
  userBDC.getRange("E3").setValue(contact.getGivenName())
  userBDC.getRange("E4").setValue(numeroAdherant)
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
  message += "Après avoir consulté le catalogue dans l'onglet correspondant pour repérer les produits, allez sur le bon de commande et modifiez la colonne \"Reference\" à l'aide de la petite flèche sur la droite.\n"
  message += "Pour que la commande soit traitée, merci d'indiquer dans le bon de commande quelle est le moyen de paiement utilisé et d'effectuer le règlement avant la date butoir par carte bancaire (de préférence), virement bancaire ou chèque.\n"
  message += "\n"
  message += "L'équipe de La Marm'Hotte\n"
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

  var dryRunSpreasheet = false
  var dryRunEmail = false
  
  var contactGroupName = "Adhérents"
  var originalSpreadsheetId = "<To be replaced>"
  var folderForSpreadsheetStorageId = "<To be replaced>"
  
  
  var originalSpreadsheet = SpreadsheetApp.openById(originalSpreadsheetId)
  var contactGroup = ContactsApp.getContactGroup(contactGroupName)
  
  var allContacts = contactGroup.getContacts()
  
  Logger.log("Will send email for %s users, quota is: %s", allContacts.length, MailApp.getRemainingDailyQuota())
  allContacts.forEach(function(c){
    
    
    var url = copySpreadsheet(c, originalSpreadsheet, folderForSpreadsheetStorageId, dryRunSpreasheet)
    sendEmail(c, url, dryRunEmail)
    
  })
  
  
}