////////////////////// Funciones de catalogo //////////////////////

// var ui = SpreadsheetApp.getUi();

var payload = {
    "columnaOrdenar": "id",
    "pagina": 0,
    "registrosPorPagina": 1000,
    "orden": "ASC",
    "filtros": [],
    "canal": 0,
    "registroInicial": 0
};

////////////////////// Listar unidades de medida //////////////////////

function listarUnidadesMedida() {
    var selectUnidad = 'unidades';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/unidadesDeMedida/listarUnidadMedida';
    
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

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/monedas/listarMonedas';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
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

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/formasDePago/listarFormaPagoDocumento';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
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

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/empresas/listarEmpresas';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
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

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentosTipos/listarTipoDocumento';
    

    
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

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/ciudad/listarCiudades';
    
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
    Logger.log(response.getResponseCode());
    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content.length);
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
        
        datosPareados = datosPareados.slice(0, 500);
        Logger.log(datosPareados);
        
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

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/centrosDeCosto/listarCentroCosto';
    
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

////////////////////// Listar Bancos //////////////////////

function listarBancos(){
    var selectBancos = 'bancos';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/listarBancos';
    
    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };

    var payloadBanco = {
        "columnaOrdenar": "nombreCompleto",
        "pagina": 0,
        "registrosPorPagina": 1000,
        "orden": "ASC",
        "filtros": [
            {
                "atributo": "codigo",
                "valor": "",
                "valor2": null,
                "tipoFiltro": 0,
                "tipoDato": 5,
                "nombreColumna": null,
                "valores": [
                    1
                ],
                "clase": "terceroTipos",
                "operador": 0,
                "subGrupo": "filtro"
            },
            {
                "atributo": "senActivo",
                "valor": "true",
                "valor2": null,
                "tipoFiltro": 0,
                "tipoDato": 1,
                "nombreColumna": null,
                "valores": null,
                "clase": "terceroTipos",
                "operador": 0,
                "subGrupo": "filtro"
            },
            {
                "atributo": "senDisponible",
                "valor": "true",
                "valor2": null,
                "tipoFiltro": 0,
                "tipoDato": 1,
                "nombreColumna": null,
                "valores": null,
                "clase": "terceroTipos",
                "operador": 0,
                "subGrupo": "filtro"
            }
        ],
        "canal": 0,
        "registroInicial": 0
    };

    var options = {
        'method': 'post',
        'headers': headers,
        'payload': JSON.stringify(payloadBanco),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectBancos);
        var scriptPropertiesBancos = PropertiesService.getScriptProperties();
        scriptPropertiesBancos.setProperty('contentBancos', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionBancos(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaBancos', seleccionString);
}

function guardarSeleccionOpcionBancos(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataBancos', seleccionOptionData);
}

function getDataBancos() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataBancos = scriptProperties.getProperty('contentBancos');
    return JSON.parse(contentDataBancos);
}

function mostrarDatosCeldaBancos() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataBancos');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaBancos');
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
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaBancos', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataBancos', '');
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
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaBancos', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataBancos', '');

    }

}

////////////////////// Listar Tipos de cuenta //////////////////////

function listarTipoCuenta(){
    var selectTipoCuenta = 'tipoCuenta';
    
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/listarTipoCuenta';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectTipoCuenta);
        var scriptPropertiesTipoCuenta = PropertiesService.getScriptProperties();
        scriptPropertiesTipoCuenta.setProperty('contentTipoCuenta', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getResponseCode();
        if(errorResponse == 403){
            ui.alert('Error', 'No tiene permisos para acceder a este recurso', ui.ButtonSet.OK);
        }
        Logger.log("Error response: " + errorResponse);
    }
    
}

function guardarSeleccionTipoCuenta(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoCuenta', seleccionString);
}

function getDataTipoCuenta() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataTipoCuenta = scriptProperties.getProperty('contentTipoCuenta');
    return JSON.parse(contentDataTipoCuenta);
}

function guardarSeleccionOpcionTipoCuenta(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataTipoCuenta', seleccionOptionData);
}

function mostrarDatosCeldaTipoCuenta() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataTipoCuenta');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaTipoCuenta');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoCuenta', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoCuenta', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoCuenta', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoCuenta', '');
    }
}

////////////////////// Listar Genero //////////////////////

function listarGenero(){
    var selectGenero = 'genero';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/listarGenero';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectGenero);
        var scriptPropertiesGenero = PropertiesService.getScriptProperties();
        scriptPropertiesGenero.setProperty('contentGenero', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionGenero(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaGenero', seleccionString);
}

function getDataGenero() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataGenero = scriptProperties.getProperty('contentGenero');
    return JSON.parse(contentDataGenero);
}

function guardarSeleccionOpcionGenero(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataGenero', seleccionOptionData);
}

function mostrarDatosCeldaGenero() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataGenero');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaGenero');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaGenero', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataGenero', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaGenero', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataGenero', '');
    }
}

////////////////////// Listar Estado Civil //////////////////////

function listarEstadoCivil(){
    var selectEstadoCivil = 'estadoCivil';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/listarEstadoCivil';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectEstadoCivil);
        var scriptPropertiesEstadoCivil = PropertiesService.getScriptProperties();
        scriptPropertiesEstadoCivil.setProperty('contentEstadoCivil', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }
}

function guardarSeleccionEstadoCivil(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaEstadoCivil', seleccionString);
}

function getDataEstadoCivil() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataEstadoCivil = scriptProperties.getProperty('contentEstadoCivil');
    return JSON.parse(contentDataEstadoCivil);
}

function guardarSeleccionOpcionEstadoCivil(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataEstadoCivil', seleccionOptionData);
}

function mostrarDatosCeldaEstadoCivil() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataEstadoCivil');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaEstadoCivil');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaEstadoCivil', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataEstadoCivil', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaEstadoCivil', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataEstadoCivil', '');
    }
}

////////////////////// Listar Prefijos //////////////////////

function listarPrefijos(){
    var selectPrefijos = 'prefijos';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentosTipos/listarPrefijoDocumento';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectPrefijos);
        var scriptPropertiesPrefijos = PropertiesService.getScriptProperties();
        scriptPropertiesPrefijos.setProperty('contentPrefijos', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionPrefijos(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaPrefijos', seleccionString);
}

function getDataPrefijos() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataPrefijos = scriptProperties.getProperty('contentPrefijos');
    return JSON.parse(contentDataPrefijos);
}

function guardarSeleccionOpcionPrefijos(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataPrefijos', seleccionOptionData);
}

function mostrarDatosCeldaPrefijos() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataPrefijos');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaPrefijos');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaPrefijos', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataPrefijos', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaPrefijos', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataPrefijos', '');
    }
}

////////////////////// Listar Paises //////////////////////

function listarPaises(){
    var selectPaises = 'paises';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/ciudad/listarPaises';
    
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

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content);
        viewGeneral(selectPaises);
        var scriptPropertiesPaises = PropertiesService.getScriptProperties();
        scriptPropertiesPaises.setProperty('contentPaises', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionPaises(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaPaises', seleccionString);
}

function guardarSeleccionOpcionPaises(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataPaises', seleccionOptionData);
}

function getDataPaises() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataPaises = scriptProperties.getProperty('contentPaises');
    return JSON.parse(contentDataPaises);
}

function mostrarDatosCeldaPaises() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataPaises');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaPaises');
    var datos = JSON.parse(datosString);
    datos = datos.flat(); 
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();
    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaPaises', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataPaises', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaPaises', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataPaises', '');
    }
}


////////////////////// Llamado de vistas general //////////////////////

function viewGeneral(select){
    if(select  == 'unidades'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-unidades.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Unidades de medida');
    }else if(select == 'monedas'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-monedas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Monedas');
    }else if(select == 'formaPago'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-forma-pago.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Formas de pago');
    }else if(select == 'empresas'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-empresas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Empresas');
    }else if(select == 'bancos'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-bancos.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Bancos');
    }else if(select == 'tipoDocumento'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-tipo-documento.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Tipos de documento');
    }else if(select == 'ciudades'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-ciudades.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Ciudades');
    }else if(select == 'centroCostos'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-centro-costo.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Centros de costos');
    }else if(select == 'tipoCuenta'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-tipo-cuenta.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(320);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Tipos de cuenta');
    }else if(select == 'genero'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-genero.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(320);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Generos');
    }else if(select == 'estadoCivil'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-estado-civil.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Estado civil');
    }else if(select == 'prefijos'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-prefijos.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Prefijos');
    }else if(select == 'paises'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/catalogo/catalogo-paises.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Paises');
    }
}


