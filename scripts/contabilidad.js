////////////////////// Funciones documentos contabilidad //////////////////////

var ui = SpreadsheetApp.getUi();

var payload = {
    "columnaOrdenar": "id",
    "pagina": 0,
    "registrosPorPagina": 1000,
    "orden": "DESC",
    "filtros": [],
    "canal": 0,
    "registroInicial": 0
};


////////////////////// Listar formas de Pago //////////////////////

function listarFormasPagoContabilidad(){
    var selectFormaPago = 'formaPagoContabilidad';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/contabilidad/listarFormasPagoContable';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);

    if (response.getResponseCode() === 202) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneralDocumentosContabilidad(selectFormaPago);
        var scriptPropertiesFormaPago = PropertiesService.getScriptProperties();
        scriptPropertiesFormaPago.setProperty('contentFormaPagoContabilidad', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
        mapeoErroresDocumentoContabilidad(errorResponse);
    }
}

function guardarSeleccionFormaPagoContabilidad(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaFormaPagoContabilidad', seleccionString);
}

function guardarSeleccionOpcionFormaPagoContabilidad(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataFormaPagoContabilidad', seleccionOptionData);
}

function getDataFormaPagoContabilidad() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataformaPago = scriptProperties.getProperty('contentFormaPagoContabilidad');
    return JSON.parse(contentDataformaPago);
}

function mostrarDatosCeldaFormaPagoContabilidad() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataFormaPagoContabilidad');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaFormaPagoContabilidad');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaFormaPagoContabilidad', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataFormaPagoContabilidad', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaFormaPagoContabilidad', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataFormaPagoContabilidad', '');
    }
}

////////////////////// Consultar contabilizaciones //////////////////////

function consultarContabilizaciones(){
    var selectFormaPago = 'consultarContabilizaciones';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/contabilidad/contabilizaciones';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        // viewGeneralDocumentosContabilidad(selectFormaPago);
        // var scriptPropertiesFormaPago = PropertiesService.getScriptProperties();
        // scriptPropertiesFormaPago.setProperty('contentFormaPagoContabilidad', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }
}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneralDocumentosContabilidad(select){
    if(select == 'formaPagoContabilidad'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/contabilidad/contabilidad-forma-pago.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(750)
        .setHeight(490);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar Forma de Pago Contabilidad');
    }
}

////////////////////// Errores de respuesta //////////////////////

function mapeoErroresDocumentoContabilidad(response){
    if(response == '403'){
        var htmlOutput = HtmlService.createHtmlOutput('<p style="font-family: Raleway, sans-serif; font-size:14px; text-align: center; margin:-25px 0 10px 0;"><span style="color:#2196F3; font-size: 48px;">&#9888;</span><br><br>No tienes permisos de consulta para este servicio, puedes activarlos desde tu cuenta de World Office cloud</p>')
        .setWidth(430)
        .setHeight(120);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
    } else if(response == '404'){
        var htmlOutput = HtmlService.createHtmlOutput('<p style="font-family: Raleway, sans-serif; font-size:14px; text-align: center; margin:-25px 0 10px 0;"><span style="color:#2196F3; font-size: 48px;">&#9888;</span><br><br>No se encontraron elementos con los criterios de busqueda seleccionados.</p>')
        .setWidth(430)
        .setHeight(120);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
    }    
}
