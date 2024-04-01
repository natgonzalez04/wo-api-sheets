////////////////////// Funciones cuentas Contables //////////////////////

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

////////////////////// Listar Cuentas Contables //////////////////////

function listarCuentasContables(){

    var selectListarCuentasContables = 'listarCuentasContables';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/cuentasContables/listarCuentaContable';
    
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
    Logger.log(response);
    Logger.log(response.length);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content.slice(0,200);
        Logger.log(content);
        viewGeneralCuentasContables(selectListarCuentasContables);
        var scriptPropertiesCuentasContables = PropertiesService.getDocumentProperties();
        scriptPropertiesCuentasContables.setProperty('contentCuentaContable', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }
}

function guardarSeleccionCuentaContable(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getUserProperties().setProperty('seleccionCeldaCuentaContable', seleccionString);
}

function guardarSeleccionOpcionCuentaContable(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getUserProperties().setProperty('optionDataCuentaContable', seleccionOptionData);
}

function getDataCuentaContable() {
    var scriptProperties = PropertiesService.getDocumentProperties();
    var contentDataCuentaContable = scriptProperties.getProperty('contentCuentaContable');
    return JSON.parse(contentDataCuentaContable);
}

function mostrarDatosCeldaCuentaContable() {
    var selectOption = PropertiesService.getUserProperties().getProperty('optionDataCuentaContable');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getUserProperties().getProperty('seleccionCeldaCuentaContable');
    var datos = JSON.parse(datosString);
    datos = datos.flat(); 
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();
    if(option == 'listar'){
        datos = datos.filter(function(item) {
            return item !== '';
        });
        for (var i = 0; i < datos.length; i+= 2) {
            var celda1 = celdaActiva.offset(i/2, 0);
            var celda2 = celdaActiva.offset(i/2, 1);
            celda1.setValue(datos[i]);
            if (i + 1 < datos.length) {
                celda2.setValue(datos[i + 1]);
            }
        }
        PropertiesService.getUserProperties().setProperty('seleccionCeldaCuentaContable', '');
        PropertiesService.getUserProperties().setProperty('optionDataCuentaContable', '');
    }else if(option == 'menu'){
        datos = datos.filter(function(item) {
            return item !== '';
        });
        
        var datosPareados = [];
        for (var i = 0; i < datos.length; i += 2) {
            var par = datos[i];
            if (i + 1 < datos.length) {
                par += ' ' + datos[i + 1];
            }
            datosPareados.push(par);
        }
        
        var regla = SpreadsheetApp.newDataValidation()
            .requireValueInList(datosPareados)
            .build();
        
        celdaActiva.setDataValidation(regla);
        PropertiesService.getUserProperties().setProperty('seleccionCeldaCuentaContable', '');
        PropertiesService.getUserProperties().setProperty('optionDataCuentaContable', '');

    }

}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneralCuentasContables(select){
    if(select == 'listarCuentasContables'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/cuentasContables/listar-cuentas-contables.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(500);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Cuentas Contables');
    }

}

////////////////////// Errores de respuesta //////////////////////

function mapeoErroresDocumento(response){
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