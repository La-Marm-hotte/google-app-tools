
// [START groupement_achat_create_spreadsheet]
/**
 * Create a spreadsheet for all user in the specified group and sent it's link by email
 */
function createSpreadSheetForEachUser() {
  var contactGroup = ContactsApp.getContactGroup('LaMarmHotte');
  var originalBDC = SpreadsheetApp.openById('1o0SZMKMcfdxjAd7dynMSzivKlTJwedxX4mlCzLEgbJ8');

  Logger.log('Daily quota: %s', MailApp.getRemainingDailyQuota());
  contactGroup.getContacts().forEach(function(c) {
    var email = c.getEmails()[0].getAddress();
    Logger.log('found user %s with email %s', c.getFullName(), email);
      var userBDC = originalBDC.copy('Bon de commande '+c.getFullName());
      userBDC.getRange('E2').setValue(c.getFamilyName());
      userBDC.getRange('E3').setValue(c.getGivenName());
      userBDC.getRange('E6').setValue(email);
      var fileOfUser = DriveApp.getFileById(userBDC.getId());
      fileOfUser.moveTo(DriveApp.getFolderById('1sOVYH81t5htAcKBizMNlUAxiVyqw247t'));
      fileOfUser.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
      Logger.log(fileOfUser.getUrl());
    MailApp.sendEmail(email, 'Bon de commande la Marm\'Hotte', ''+
                      'Bonjour, voici le lien vers le bon de commande: '+
                     fileOfUser.getUrl());
  Logger.log('Daily quota: %s', MailApp.getRemainingDailyQuota());
  });
}
// [END groupement_achat_create_spreadsheet]
