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
  PropertiesService.getUserProperties().setProperty('claveAPI', clave);
}

//almacenamiento de la clave api global

function almacenamientoClave() {
  return PropertiesService.getUserProperties().getProperty('claveAPI');
}

function validacionClaveAPI() {
    var claveAPI = almacenamientoClave();
    var apiUrl = 'https://api.worldoffice.cloud/validaTokenAPI';

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
      ui.alert('Tu token API ha sido confirmado con exito.');
      return Boolean(claveAPI);
    }else {
      ui.alert('Tu token API esta inactivo y/o no ha sido configurado.');
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
    .addItem('Listar terceros', 'listarTerceros')
    .addItem('Consultar tercero', 'consultarTerceros')
    .addItem('Consultar dirección tercero', 'consultarTercerosDireccion')
    .addItem('Listar clasificación impuestos', 'listarClasificacionImpuestos')
    .addItem('Tipos de contribuyente', 'listarTiposContribuyente')
    .addItem('Tipos de identificación', 'listarTiposIdentificacion')
    .addItem('Tipos de tercero', 'listarTiposTercero')


    var Documentos = ui.createMenu('Documentos')
    .addItem('Consultar documentos de ventas', 'listarDocumentosVenta')
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
    var claveAPI = PropertiesService.getUserProperties().getProperty('claveAPI');
    var claveAPIFormateada = claveAPI.match(/.{1,50}/g).join("\n");
    ui.alert('Tu token API en uso es: \n\n' + claveAPIFormateada);
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
  PropertiesService.getUserProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().deleteAllProperties();
  onOpen();
}



