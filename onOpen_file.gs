function onOpen() {
    ShowSideBar();
}

function ShowSideBar() {
    // var html = HtmlService.createHtmlOutputFromFile('Index')
    var html = HtmlService.createTemplateFromFile('Index').evaluate()
        .setTitle('Console de gestion')
        .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showSidebar(html);
}
// new comment
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();

}

function onOpen_referentiel() {
    ShowSideBar_referentiel();
}

function ShowSideBar_referentiel() {
    // var html = HtmlService.createHtmlOutputFromFile('Index')
    var html = HtmlService.createTemplateFromFile('Index-referentiel').evaluate()
        .setTitle('Console référentiel')
        .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showSidebar(html);
}


function onOpen_appel() {
    ShowSideBar_appel();
}

function ShowSideBar_appel() {
    // var html = HtmlService.createHtmlOutputFromFile('Index')
    var html = HtmlService.createTemplateFromFile('Index-appel').evaluate()
        .setTitle('Console d\'appel')
        .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showSidebar(html);
}


