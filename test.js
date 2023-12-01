// var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
// var last = sh.getLastRow();

// var ui = SpreadsheetApp.getUi();

// function onOpen() {

//   ui.createMenu('World Office')
//     .addItem('ConfigurarTest', 'formGuardarClaveAPI')
//     .addItem('Ver API key', 'usarClaveAPI')

// }


// function formGuardarClaveAPI(){

//   var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');
//   claveAPI = 'DKAJSDLKAJS'

//   var htmlOutput = HtmlService.createHtmlOutput(
//     '<!DOCTYPE html>' +
//     '<html>' +
//     '<head>' +
//     '<meta charset="utf-8">' +
//     '<title>Configuración de API</title>' +
//     '<link rel="stylesheet" href="https://descargaswo.com/cloud/css/styles.css">' +
//     '</head>' +
//     '<body>' +
//     '<form id="apiConfigForm">' +
//     '<label for="apiKey">Clave API de WOCloud:</label>' +
//     '<input type="text" value="' + claveAPI + '" id="apiKey" name="apiKey"><br><br>' +
//     '<input type="button" value="Guardar" onclick="guardarClaveAPI()">' +
//     '</form>' +
//     '<script>' +
//     'function guardarClaveAPI() {' +
//     'var claveAPI = document.getElementById("apiKey").value;' +
//     'google.script.run.guardarClaveAPI(claveAPI);' +
//     'google.script.host.close();' +
//     '}' +
//     '</script>' +
//     '</body>' +
//     '</html>'
//   )
//   .setWidth(300)
//   .setTitle('Configurar API');

// }



// function guardarClaveAPI(clave){
//   PropertiesService.getDocumentProperties().setProperty('claveAPI', clave);
// }
