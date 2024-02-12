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
    .addItem('Centro de costos', 'listarCentroCostos')
    .addItem('Ciudades', 'listarCiudades')
    .addItem('Estado civil', 'listarEstadoCivil')
    .addItem('Formas de pago', 'listarFormasPago')
    .addItem('Generos', 'listarGenero')
    .addItem('Listar bancos', 'listarBancos')
    .addItem('Listar empresas', 'listarEmpresas')
    .addItem('Monedas', 'listarMonedas')
    .addItem('Paises', 'listarPaises')
    .addItem('Prefijos', 'listarPrefijos')
    .addItem('Tipos de cuenta', 'listarTipoCuenta')
    .addItem('Tipos de documento', 'listarTiposDocumento')
    .addItem('Unidades de medida', 'listarUnidadesMedida')

    var Terceros = ui.createMenu('Terceros')
    .addItem('Consultar todos los terceros', 'consumeAndDisplayAPI')
    .addItem('Consultar direcciones', 'direccionesListar');
    
    var Documentos = ui.createMenu('Documentos')
    .addItem('Consultar documentos de ventas', 'listardocumentosVenta')
    .addItem('Consultar documentos de compra', 'listardocumentosCompra');

    var Inventarios = ui.createMenu('Inventarios')
    .addItem('Consultar todos los inventarios', 'inventariosListar')
    .addItem('Consultar bodegas', 'bodegasListar');

    var CuentasContables = ui.createMenu('Cuentas Contables')
    .addItem('Consultar cuentas contables', 'inventariosListar');

    var Contabilidad = ui.createMenu('Contabilidad')
    .addItem('Consultar documentos contables', 'documentosVenta');

    var menu = ui.createMenu('World Office');
    menu.addItem('Ver Token', 'mostrarClaveAPI')
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


function listardocumentosCompra(){

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



