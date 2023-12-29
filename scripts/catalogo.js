////////////////////// Funciones de catalogo //////////////////////

var ui = SpreadsheetApp.getUi();

var payload = {
    "columnaOrdenar": "id",
    "pagina": 0,
    "registrosPorPagina": 10,
    "orden": "DESC",
    "filtros": [],
    "canal": 0,
    "registroInicial": 0
};

////////////////////// Listar unidades de medida //////////////////////

function listarUnidadesMedida() {
    var selectUnidad = 'unidades';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/unidadesDeMedida/listarUnidadMedida';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };
    
    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectUnidad);
        var scriptPropertiesUnidad = PropertiesService.getScriptProperties();
        scriptPropertiesUnidad.setProperty('contentUnidad', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }
    
}

function guardarSeleccionUnidad(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaUnidad', seleccionString);
}

function getDataUnidad() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataUnidad = scriptProperties.getProperty('contentUnidad');
    return JSON.parse(contentDataUnidad);
}

function guardarSeleccionOpcionUnidad(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataUnidad', seleccionOptionData);
}

function mostrarDatosCeldaUnidad() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataUnidad');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaUnidad');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaUnidad', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataUnidad', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaUnidad', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataUnidad', '');
    }
}

////////////////////// Listar monedas //////////////////////

function listarMonedas(){
    var selectMoneda = 'monedas';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/monedas/listarMonedas';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.content;
        Logger.log(content);
        viewGeneral(selectMoneda);
        var scriptPropertiesMoneda = PropertiesService.getScriptProperties();
        scriptPropertiesMoneda.setProperty('contentMoneda', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionMoneda(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaMoneda', seleccionString);
}

function guardarSeleccionOpcionMoneda(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataMoneda', seleccionOptionData);
}

function getDataMoneda() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataMoneda = scriptProperties.getProperty('contentMoneda');
    return JSON.parse(contentDataMoneda);
}

function mostrarDatosCeldaMoneda() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataMoneda');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaMoneda');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaMoneda', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataMoneda', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaMoneda', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataMoneda', '');
    }


}


////////////////////// Listar formas de Pago //////////////////////

function listarFormasPago(){
    var selectFormaPago = 'formaPago';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/formasDePago/listarFormaPago';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 202) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.content;
        Logger.log(content);
        viewGeneral(selectFormaPago);
        var scriptPropertiesFormaPago = PropertiesService.getScriptProperties();
        scriptPropertiesFormaPago.setProperty('contentFormaPago', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }
}

function guardarSeleccionFormaPago(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaFormaPago', seleccionString);
}

function guardarSeleccionOpcionFormaPago(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataFormaPago', seleccionOptionData);
}

function getDataFormaPago() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataformaPago = scriptProperties.getProperty('contentFormaPago');
    return JSON.parse(contentDataformaPago);
}

function mostrarDatosCeldaFormaPago() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataFormaPago');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaFormaPago');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaFormaPago', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataFormaPago', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaFormaPago', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataFormaPago', '');
    }
}

////////////////////// Listar empresas //////////////////////

function listarEmpresas(){
    var selectEmpresas = 'empresas';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/empresas/listarEmpresas';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.content;
        Logger.log(content);
        viewGeneral(selectEmpresas);
        var scriptPropertiesEmpresas = PropertiesService.getScriptProperties();
        scriptPropertiesEmpresas.setProperty('contentEmpresa', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionEmpresa(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaEmpresa', seleccionString);
}

function guardarSeleccionOpcionEmpresa(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataEmpresa', seleccionOptionData);
}

function getDataEmpresa() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataEmpresa = scriptProperties.getProperty('contentEmpresa');
    return JSON.parse(contentDataEmpresa);
}

function mostrarDatosCeldaEmpresa() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataEmpresa');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaEmpresa');
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
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaEmpresa', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataEmpresa', '');
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
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaEmpresa', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataEmpresa', '');

    }

}

////////////////////// Listar tipos Documento //////////////////////

function listarTiposDocumento(){
    var selectTipoDoc = 'tipoDocumento';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/documentosTipos/listarTipoDocumento';
    

    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectTipoDoc);
        var scriptPropertiesTipoDoc = PropertiesService.getScriptProperties();
        scriptPropertiesTipoDoc.setProperty('contentTipoDoc', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }
}

function guardarSeleccionTipoDoc(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoDoc', seleccionString);
}

function guardarSeleccionOpcionTipoDoc(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataTipoDoc', seleccionOptionData);
}

function getTipoDocumento() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataTipoDoc = scriptProperties.getProperty('contentTipoDoc');
    return JSON.parse(contentDataTipoDoc);
}

function mostrarDatosCeldaTipoDoc() {
    
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataTipoDoc');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaTipoDoc');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoDoc', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoDoc', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoDoc', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoDoc', '');
    }
}

////////////////////// Listar Ciudades //////////////////////

function listarCiudades(){
    var selectCiudades = 'ciudades';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/ciudad/listarCiudades';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.content;
        Logger.log(content);
        viewGeneral(selectCiudades);
        var scriptPropertiesCiudades = PropertiesService.getScriptProperties();
        scriptPropertiesCiudades.setProperty('contentCiudades', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionCiudades(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaCiudades', seleccionString);
}

function guardarSeleccionOpcionCiudades(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataCiudades', seleccionOptionData);
}

function getCiudades() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataCiudades = scriptProperties.getProperty('contentCiudades');
    return JSON.parse(contentDataCiudades);
}

function mostrarDatosCeldaCiudades() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataCiudades');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaCiudades');
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
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaCiudades', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataCiudades', '');
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
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaCiudades', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataCiudades', '');

    }
}

////////////////////// Listar CentroCostos //////////////////////

function listarCentroCostos(){
    var selectCentroCostos = 'centroCostos';

    var claveAPI = PropertiesService.getDocumentProperties().getProperty('claveAPI');

    var apiUrl = 'https://apidev.worldoffice.cloud/api/v1/centrosDeCosto/listarCentroCosto';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': 'WO ' + claveAPI,
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectCentroCostos);
        var scriptPropertiesCentroCosto = PropertiesService.getScriptProperties();
        scriptPropertiesCentroCosto.setProperty('contentCentroCosto', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionCentroCosto(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaCentroCosto', seleccionString);
}

function guardarSeleccionOpcionCentroCosto(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataCentroCosto', seleccionOptionData);
}

function getDataCentroCosto() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataCentroCosto = scriptProperties.getProperty('contentCentroCosto');
    return JSON.parse(contentDataCentroCosto);
}

function mostrarDatosCeldaCentroCosto() {
    
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataCentroCosto');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaCentroCosto');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaCentroCosto', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataCentroCosto', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaCentroCosto', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataCentroCosto', '');
    }
}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneral(select){
    if(select  == 'unidades'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-unidades.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Selecciona una o mas unidades de medida');
    }else if(select == 'monedas'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-monedas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Selecciona una o mas monedas');
    }else if(select == 'formaPago'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-forma-pago.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Selecciona una o mas formas de pago');
    }else if(select == 'empresas'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-empresas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Selecciona una o mas empresas');
    }else if(select == 'tipoDocumento'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-tipo-documento.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Selecciona uno o mas tipos de documento');
    }else if(select == 'ciudades'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-ciudades.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete
            , 'Selecciona uno o mas ciudades');
    }else if(select == 'centroCostos'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo-centro-costo.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(330);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Selecciona uno o mas centros de costos');
    }
    
}