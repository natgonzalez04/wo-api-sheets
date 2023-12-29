var ui = SpreadsheetApp.getUi();


function onOpen(){
  ui.createMenu('World Office')
  .addItem('Configuración', 'mostrarPaginaPrincipal')
  .addToUi();
}

function mostrarPaginaPrincipal() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('views/login.html')
        .setWidth(400)
        .setHeight(330);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bienvenido a World Office Cloud Api');
}

function guardarClaveAPI(clave) {
  PropertiesService.getDocumentProperties().setProperty('claveAPI', clave);
}

function validacionClaveAPI() {
    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');
    var apiUrl = 'https://apidev.worldoffice.cloud/validaTokenAPI';

    var headers = {
      'Content-Type': 'application/json',
      'Authorization': 'WO ' + claveAPI,
    };

    var options = {
      'method': 'get',
      'headers': headers,
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);
    if (response == 'true'){
      activarOpciones();
      ui.alert('Tu clave API ha sido confirmada con exito.');
      return Boolean(claveAPI);
    }else {
      ui.alert('Tu clave API esta inactiva y/o no ha sido configurada.');
      return Boolean(claveAPI);
    }

}

function activarOpciones(){
    var ui = SpreadsheetApp.getUi();
    var Catalogos = ui.createMenu('Catálogos')
    .addItem('Unidades de medida', 'listarUnidadesMedida')
    .addItem('Monedas', 'listarMonedas')
    .addItem('Formas de pago', 'listarFormasPago')
    .addItem('Listar empresas', 'listarEmpresas')
    .addItem('Tipos de Documento', 'listarTiposDocumento')
    .addItem('Ciudades', 'listarCiudades')
    .addItem('Centro de costos', 'listarCentroCostos')

    var Terceros = ui.createMenu('Terceros')
    .addItem('Consultar todos los terceros', 'consumeAndDisplayAPI')
    .addItem('Consultar direcciones', 'direccionesListar');
    
    var Documentos = ui.createMenu('Documentos')
    .addItem('Consultar documentos de ventas', 'documentosVenta')
    .addItem('Consultar documentos de compra', 'direccionesListar');

    var Inventarios = ui.createMenu('Inventarios')
    .addItem('Consultar todos los inventarios', 'inventariosListar')
    .addItem('Consultar bodegas', 'bodegasListar');

    var CuentasContables = ui.createMenu('Cuentas Contables')
    .addItem('Consultar cuentas contables', 'inventariosListar');

    var Contabilidad = ui.createMenu('Contabilidad')
    .addItem('Consultar documentos contables', 'documentosVenta');

    var menu = ui.createMenu('World Office');
    menu.addItem('Ver API key', 'mostrarClaveAPI')
    menu.addItem('Cerrar Sesión', 'cerrarSesion')
    .addSeparator()
    .addSubMenu(Catalogos)
    .addSubMenu(Documentos)
    .addSubMenu(Terceros)
    .addSubMenu(Inventarios)
    .addSubMenu(CuentasContables)
    .addSubMenu(Contabilidad)
    .addToUi();
}

function mostrarClaveAPI(){
    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');
    var claveAPIFormateada = claveAPI.match(/.{1,50}/g).join("\n");
    ui.alert('Tu clave API es: \n' + claveAPIFormateada);
}

////////////////////// funciones de documentos //////////////////////


function documentosVenta(){

}

function functiondireccionesListar(){

}

function inventariosListar(){

}

function bodegasListar(){

}

function cerrarSesion(){
  PropertiesService.getDocumentProperties().deleteAllProperties();
  onOpen();
}




// funciones de consulta antiguas

// function inicializarCatalogo() {
//   var opcionesCatalogo = [
//       { nombre: 'Seleccione', id: 0},
//       { nombre: 'Unidades de Medida', id: 1},
//       { nombre: 'Tipos de Moneda', id: 2},
//       { nombre: 'Formas de Pago', id: 3},
//       { nombre: 'Empresas', id: 4},
//       { nombre: 'Tipos de Documento', id: 5},
//       { nombre: 'Ciudades', id: 6},  
//       { nombre: 'Centro de Costos', id: 7}
//   ]

//     var htmlOutput = HtmlService.createHtmlOutput(
//       '<form id="catalogoForm">' +
//       '  <label for="opcionesCatalogo">Selecciona una opcion</label>' +
//       '  <select id="opcionesCatalogo" name="opcionesCatalogo" onchange="guardarSelectionCatalogo(this.value)">' +
//       opcionesCatalogo.map(function(opcion) {
//         return '<option value="' + opcion.id + '">' + opcion.nombre + '</option>';
//       }).join('') +
//       '  </select><br><br>' +
//       '  <p id="seleccionActual"></p>' + 
//       '</form>' +
//       '<script>' +
//       '  function guardarSelectionCatalogo(opcion) {' +
//       '    var opcionSeleccionada = ' + JSON.stringify(opcionesCatalogo) + '.find(function(op) { return op.id.toString() === opcion; });' +
//       '    document.getElementById("seleccionActual").innerText = "Selección actual: " + (opcionSeleccionada ? opcionSeleccionada.nombre : "Ninguna");' +
//       '  }' +
//       '</script>'
//     )
//     .setWidth(300)
  
//     SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Catalogo API');
// }


// function listarEmpresas(){
//   var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

//   var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/empresas/listarEmpresas';
  
//   var headers = {
//     'Content-Type': 'application/json',
//     'Authorization': 'WO ' + claveAPI,
//   };
  
//   var payload = {
//     "columnaOrdenar": "id",
//     "pagina": 0,
//     "registrosPorPagina": 10,
//     "orden": "DESC",
//     "filtros": [],
//     "canal": 0,
//     "registroInicial": 0
//   };

//   var options = {
//     'method': 'post',
//     'headers': headers,
//     'payload': JSON.stringify(payload),
//     'muteHttpExceptions': true
//   };
  
//   var response = UrlFetchApp.fetch(apiUrl, options);
//   Logger.log(response);

//   if (response.getResponseCode() === 200) {
//     var responseData = response.getContentText();
//     var jsonData = JSON.parse(responseData);
//     var content = jsonData.content;
//     Logger.log(content);
//   } 
//   else 
//   {
//     var errorResponse = response.getContentText();
//     Logger.log("Error response: " + errorResponse);
//   }

// }

// function listarUnidadMoneda(){
//   var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

//   var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/monedas/listarMonedas';
  
//   var headers = {
//     'Content-Type': 'application/json',
//     'Authorization': 'WO ' + claveAPI,
//   };
  
//   var payload = {
//     "columnaOrdenar": "id",
//     "pagina": 0,
//     "registrosPorPagina": 10,
//     "orden": "DESC",
//     "filtros": [],
//     "canal": 0,
//     "registroInicial": 0
//   };

//   var options = {
//     'method': 'post',
//     'headers': headers,
//     'payload': JSON.stringify(payload),
//     'muteHttpExceptions': true
//   };
  
//   var response = UrlFetchApp.fetch(apiUrl, options);
//   Logger.log(response);

//   if (response.getResponseCode() === 200) {
//     var responseData = response.getContentText();
//     var jsonData = JSON.parse(responseData);
//     var content = jsonData.content;
//     Logger.log(content);
//   } 
//   else 
//   {
//     var errorResponse = response.getContentText();
//     Logger.log("Error response: " + errorResponse);
//   }

// }

// function listarUnidadMoneda(){
//   var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

//   var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/unidadesDeMedida/listarUnidadMedida';
  
//   var headers = {
//     'Content-Type': 'application/json',
//     'Authorization': 'WO ' + claveAPI,
//   };
  
//   var payload = {
//     "columnaOrdenar": "id",
//     "pagina": 0,
//     "registrosPorPagina": 10,
//     "orden": "DESC",
//     "filtros": [],
//     "canal": 0,
//     "registroInicial": 0
//   };

//   var options = {
//     'method': 'post',
//     'headers': headers,
//     'payload': JSON.stringify(payload),
//     'muteHttpExceptions': true
//   };
  
//   var response = UrlFetchApp.fetch(apiUrl, options);
//   Logger.log(response);

//   if (response.getResponseCode() === 200) {
//     var responseData = response.getContentText();
//     var jsonData = JSON.parse(responseData);
//     var content = jsonData.content;
//     Logger.log(content);
//   } 
//   else 
//   {
//     var errorResponse = response.getContentText();
//     Logger.log("Error response: " + errorResponse);
//   }

// }


