function addEditortoAllProtectedRanges(EmailsInStringFormat) {

  var ss = SpreadsheetApp.getActive()
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  for (var i = 0; i < protections.length; i++) {

    var protection = protections[i];
    if (protection.canEdit()) {
      protection.addEditors(EmailsInStringFormat)

    }
  }

  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  for (var i = 0; i < protections.length; i++) {

    var protection = protections[i];

    if (protection.canEdit()) {

      protection.addEditors(EmailsInStringFormat)
    }
  }

}