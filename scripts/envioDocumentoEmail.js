////////////////////// Funciones iniciales //////////////////////

// var ui = SpreadsheetApp.getUi();

var payload = {
    "columnaOrdenar": "id",
    "pagina": 0,
    "registrosPorPagina": 1000,
    "orden": "DESC",
    "filtros": [],
    "canal": 0,
    "registroInicial": 0
};


////////////////////// Enviar documentos por email //////////////////////

function guardarSeleccionDocumentoGeneral(seleccion) {
    var documentoGeneralEmail = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', documentoGeneralEmail);
}

function consultarDocumentoGeneralSelect(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('dataDocumentoGeneralEmail');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
     
    var categoriaConsulta = datos.categoriaDocumento;
    var numeroDocumento = datos.numeroDocumento;
    var tipoDocumento = datos.tipoDocumento;

    if(categoriaConsulta == 'VENTAS'){
        consultaInicialDocumentoVenta(tipoDocumento, numeroDocumento);
    } else if(categoriaConsulta == 'COMPRAS'){
        consultaInicialDocumentoCompra(tipoDocumento, numeroDocumento);
    } else if(categoriaConsulta == 'CONTABLES'){
        consultaInicialDocumentoContables(tipoDocumento, numeroDocumento);
    }
}

function enviarDocumentoEmailVenta(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('dataDocumentoVentaEmail');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idDocumento = datos.idDocumento;
    var email = datos.email;
    var nombreDestinatario = datos.nombreDestinatario;

    var claveAPI = almacenamientoClave();
        
    var apiUrl = `https://api.worldoffice.cloud/api/v1/documentos/enviaDocumentoMail/${idDocumento}/${email}/${nombreDestinatario}`;

    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };
    
    var options = {
        'method': 'get',
        'headers': headers,
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = [jsonData.data];
        Logger.log(content);
    }else{
        var errorResponse = response.getResponseCode();
        Logger.log("Error response: " + errorResponse);
        mapeoErroresDocumentosEmail(errorResponse);
    }

    PropertiesService.getDocumentProperties().setProperty('dataDocumentoVentaEmail', '');

}

function consultaInicialDocumentoVenta(tipoDoc, numDoc) {

    if(tipoDoc && numDoc) {
        var claveAPI = almacenamientoClave();
        var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'

        var payloadFV = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 2000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": tipoDoc,
                    "valor2": null,
                    "tipoFiltro": 0,
                    "tipoDato": 0,
                    "nombreColumna": null,
                    "valores": null,
                    "clase": null,
                    "operador": 0,
                    "subGrupo": "filtro"
                },
                {
                    "atributo": "numero",
                    "tipoDato": 4,
                    "nombreColumna": "Número",
                    "tipoFiltro": 0,
                    "valor": numDoc,
                    "operador": 0
                }
            ],
            "canal": 0,
            "registroInicial": 0
        };


        var headers = {
            'Content-Type': 'application/json',
            'Authorization': claveAPI,
        };

        var options = {
            'method': 'post',
            'headers': headers,
            'payload': JSON.stringify(payloadFV),
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);

        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content);
            Logger.log(content.length);

            var scriptPropertiesDocEmail = PropertiesService.getScriptProperties();
            scriptPropertiesDocEmail.setProperty('contentDataRetorno', JSON.stringify(content));

            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
        }else {
            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);

            mapeoErroresDocumentosEmail(errorResponse);
        }

    }

    PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');

}

function consultaInicialDocumentoCompra(tipoDoc, numDoc) {

    if(tipoDoc && numDoc) {
        var claveAPI = almacenamientoClave();
        var apiUrl = 'https://api.worldoffice.cloud/api/v1/compra/listarDocumentoCompra'

        var payloadFV = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 2000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": tipoDoc,
                    "valor2": null,
                    "tipoFiltro": 0,
                    "tipoDato": 0,
                    "nombreColumna": null,
                    "valores": null,
                    "clase": null,
                    "operador": 0,
                    "subGrupo": "filtro"
                },
                {
                    "atributo": "numero",
                    "tipoDato": 4,
                    "nombreColumna": "Número",
                    "tipoFiltro": 0,
                    "valor": numDoc,
                    "operador": 0
                }
            ],
            "canal": 0,
            "registroInicial": 0
        };


        var headers = {
            'Content-Type': 'application/json',
            'Authorization': claveAPI,
        };

        var options = {
            'method': 'post',
            'headers': headers,
            'payload': JSON.stringify(payloadFV),
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);

        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content);
            Logger.log(content.length);

            var scriptPropertiesDocEmail = PropertiesService.getScriptProperties();
            scriptPropertiesDocEmail.setProperty('contentDataRetorno', JSON.stringify(content));

            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
        }else {
            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);

            mapeoErroresDocumentosEmail(errorResponse);
        }

    }

    PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
}

function consultaInicialDocumentoContables(tipoDoc, numDoc) {

    if(tipoDoc && numDoc) {
        var claveAPI = almacenamientoClave();
        var apiUrl = 'https://api.worldoffice.cloud/api/v1/contabilidad/listarDocContable'

        var payloadFV = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 2000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": tipoDoc,
                    "valor2": null,
                    "tipoFiltro": 0,
                    "tipoDato": 0,
                    "nombreColumna": null,
                    "valores": null,
                    "clase": null,
                    "operador": 0,
                    "subGrupo": "filtro"
                },
                {
                    "atributo": "numero",
                    "tipoDato": 4,
                    "nombreColumna": "Número",
                    "tipoFiltro": 0,
                    "valor": numDoc,
                    "operador": 0
                }
            ],
            "canal": 0,
            "registroInicial": 0
        };


        var headers = {
            'Content-Type': 'application/json',
            'Authorization': claveAPI,
        };

        var options = {
            'method': 'post',
            'headers': headers,
            'payload': JSON.stringify(payloadFV),
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);

        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content);
            Logger.log(content.length);

            var scriptPropertiesDocEmail = PropertiesService.getScriptProperties();
            scriptPropertiesDocEmail.setProperty('contentDataRetorno', JSON.stringify(content));

            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
        }else {
            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);

            mapeoErroresDocumentosEmail(errorResponse);
        }

    }

    PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralEmail', '');
}

function getDataRetornoDocumentos() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataRetorno = scriptProperties.getProperty('contentDataRetorno');
    return JSON.parse(contentDataRetorno);
}

function guardarDocumentoEmailFinal(dataSeleccion) {
    var documentoFinalEmail = JSON.stringify(dataSeleccion);
    PropertiesService.getDocumentProperties().setProperty('dataEnvioEmailFinal', documentoFinalEmail);
}

function enviarDocumentoEmailVenta(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('dataEnvioEmailFinal');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idDocumento = datos.idDocumento;
    var email = datos.email;
    var nombreDestinatario = datos.nombreDestinatario;

    var claveAPI = almacenamientoClave();
        
    var apiUrl = `https://api.worldoffice.cloud/api/v1/documentos/enviaDocumentoMail/${idDocumento}/${email}/${nombreDestinatario}`;

    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };
    
    var options = {
        'method': 'get',
        'headers': headers,
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = [jsonData.data];
        Logger.log(content);
    }else{
        var errorResponse = response.getResponseCode();
        Logger.log("Error response: " + errorResponse);
        mapeoErroresDocumentosEmail(errorResponse);
    }
    PropertiesService.getScriptProperties().setProperty('dataDocumentoGeneralEmail', '');
    PropertiesService.getDocumentProperties().deleteProperty('contentDataRetorno');

}

function enviarDocumentoEmailCompra(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('dataEnvioEmailFinal');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idDocumento = datos.idDocumento;
    var email = datos.email;
    var nombreDestinatario = datos.nombreDestinatario;

    var claveAPI = almacenamientoClave();
        
    var apiUrl = `https://api.worldoffice.cloud/api/v1/compra/enviaCompraMail/${idDocumento}/${email}/${nombreDestinatario}`;

    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };
    
    var options = {
        'method': 'get',
        'headers': headers,
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = [jsonData.data];
        Logger.log(content);
    }else{
        var errorResponse = response.getResponseCode();
        Logger.log("Error response: " + errorResponse);
        mapeoErroresDocumentosEmail(errorResponse);
    }

    PropertiesService.getScriptProperties().setProperty('dataDocumentoGeneralEmail', '');
    PropertiesService.getDocumentProperties().deleteProperty('contentDataRetorno');

}

function enviarDocumentoEmailContable(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('dataEnvioEmailFinal');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idDocumento = datos.idDocumento;
    var email = datos.email;
    var nombreDestinatario = datos.nombreDestinatario;

    var claveAPI = almacenamientoClave();
        
    var apiUrl = `https://api.worldoffice.cloud/api/v1/contabilidad/enviarDocumentoEmail/${idDocumento}/${email}/${nombreDestinatario}`;

    var headers = {
        'Content-Type': 'application/json',
        'Authorization': claveAPI,
    };
    
    var options = {
        'method': 'get',
        'headers': headers,
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = [jsonData.data];
        Logger.log(content);
    }else{
        var errorResponse = response.getResponseCode();
        Logger.log("Error response: " + errorResponse);
        mapeoErroresDocumentosEmail(errorResponse);
    }

    PropertiesService.getScriptProperties().setProperty('dataDocumentoGeneralEmail', '');
    PropertiesService.getDocumentProperties().deleteProperty('contentDataRetorno');
}


////////////////////// Errores de respuesta //////////////////////

function mapeoErroresDocumentosEmail(response){
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

