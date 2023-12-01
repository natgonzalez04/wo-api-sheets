var ui = SpreadsheetApp.getUi();

function onOpen(){
  ui.createMenu('World Office')
  .addItem('Configuración', 'setConfiguration')
  .addItem('Ver API Key', 'mostrarPaginaPrincipal')
  .addToUi();

}

function setConfiguration(){
  ui.alert('Hola, Ventana de configuracion');
}

//function apiKey(){
//  ui.alert('Hola, Ventana de apiKey');
//}


function mostrarPaginaPrincipal() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Index.html')
        .setTitle('Mi Complemento');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
}