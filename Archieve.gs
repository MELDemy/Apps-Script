//    عمل قايمه جديده
function menuItem(menuName1, funName1) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('My Menu')
    .addItem(menuName1, funName1)
    // .addSubMenu(ui.createMenu('حدث')
    //     .addItem('menuName2', 'funName2'))

    // .addSeparator()

    // .addItem('menuName3', 'FunName3')

    .addToUi(); ///////////////
}