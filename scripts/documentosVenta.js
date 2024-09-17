////////////////////// Funciones documentos Venta //////////////////////

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

////////////////////// Listar Documento Ventas //////////////////////

function listarDocumentosVenta(){
    var selectListarDocVentas = 'listarDocVentas';
    viewGeneralDocumentos(selectListarDocVentas);
}

function guardarSeleccionPaginadoDocumentos(seleccion) {
    var seleccionPaginadoDoc = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', seleccionPaginadoDoc);
}

function validacionInicial(){
    var tipoDocumentoSeleccionado = PropertiesService.getDocumentProperties().getProperty('seleccionPaginadoDocumentos');
    var datosSelect = JSON.parse(tipoDocumentoSeleccionado);
    Logger.log(datosSelect.tipoDocumento);
    if(datosSelect != null){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'

        var payloadInicial = {
            "columnaOrdenar": "id",
            "pagina": 0,
            "registrosPorPagina": 1,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": datosSelect.tipoDocumento,
                    "valor2": null,
                    "tipoFiltro": 0,
                    "tipoDato": 0,
                    "nombreColumna": null,
                    "valores": null,
                    "clase": null,
                    "operador": 0,
                    "subGrupo": "filtro"
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
            'payload': JSON.stringify(payloadInicial),
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var totalRegistros = jsonData.data.totalElements;
            Logger.log(totalRegistros);
            var scriptPropertiesTotalRegistros = PropertiesService.getScriptProperties();
            scriptPropertiesTotalRegistros.setProperty('contentTotalRegistros', JSON.stringify(totalRegistros));
        }
    }

}

function getDataTotalRegistros() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataTotalRegistros = scriptProperties.getProperty('contentTotalRegistros');
    return JSON.parse(contentDataTotalRegistros);
}

function mostraDocumentosVenta(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionPaginadoDocumentos');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var tipoDoc = datos.tipoDocumento;
    var totalRegistros = getDataTotalRegistros();
    Logger.log(totalRegistros);
    var paginas = totalRegistros >= 2000 ? Math.ceil(totalRegistros / 2000) : 1;
    var pag = 0;
    var procesoTerminado = false;
    var fechaInicial = datos.registroInicial;
    var fechaFin = datos.registroFinal;
    var numeroInicial = datos.numeroInicial;
    var numeroFinal = datos.numeroFinal;
    var prefijoSelect = datos.prefijo;
    var tercero = datos.tercero;
    var tipoTercero = datos.terceroTipo;
    
    if(datos.registrosCompletos == true){
        for (var i = 0; i < paginas; i++) {
            var datoInicial = i * 2000;
            listarRegistroCompletos(tipoDoc, datoInicial,pag)
            pag++;
            if (pag == paginas) {
                procesoTerminado = true;
                break;
            }
        }
    }else if(datos.registroInicial && datos.registroFinal){
        listarRegistroFecha(tipoDoc, fechaInicial, fechaFin);
        procesoTerminado = true;
    }else if(datos.numeroInicial && datos.numeroFinal){
        listarRegistroNumero(tipoDoc, numeroInicial, numeroFinal);
        procesoTerminado = true;
    }else if(datos.prefijo){
        listarRegistroPrefijo(tipoDoc, prefijoSelect);
        procesoTerminado = true;
    }else if(datos.tercero && tipoTercero == 'empresa'){
        listarRegistroTerceroEmpresa(tipoDoc, tercero);
        procesoTerminado = true;
    }else if(datos.tercero && tipoTercero == 'cliente'){
        listarRegistroTerceroCliente(tipoDoc, tercero);
        procesoTerminado = true;
    }else if(datos.tercero && tipoTercero == 'vendedor'){
        listarRegistroTerceroVendedor(tipoDoc, tercero);
        procesoTerminado = true;
    }
    

    return procesoTerminado;
}

function listarRegistroCompletos(documento, inicial, pagina){

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
                "valor": documento,
                "valor2": null,
                "tipoFiltro": 0,
                "tipoDato": 0,
                "nombreColumna": null,
                "valores": null,
                "clase": null,
                "operador": 0,
                "subGrupo": "filtro"
            }
        ],
        "canal": 0,
        "registroInicial": inicial
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

    var registrosContador = pagina * 2000;

    if (response.getResponseCode() === 200) {
        var responseData = response.getContentText();
        var jsonData = JSON.parse(responseData);
        var content = jsonData.data.content;
        Logger.log(content.length);
        var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var celdaActiva;

        if (pagina >= 1) {
            celdaActiva = hojaActiva.getActiveCell().offset(registrosContador + 1, 0);
        } else {
            celdaActiva = hojaActiva.getActiveCell();
        }

        var keys = {
            "id": "Id",
            "fecha": "Fecha",
            "prefijo": "Prefijo",
            "numero": "Número",
            "empresa": "Empresa",
            "terceroExterno": "Cliente",
            "terceroInterno": "Vendedor",
            "formaPago": "Forma de Pago",
            "concepto": "Concepto",
        };

        var j = 0;
        for (var key in keys) {
            var headerCell = celdaActiva.offset(0, j);
            headerCell.setValue(keys[key]);
            j++;
        }

        for (var i = 0; i < content.length; i++) {
            j = 0;
            for (var key in keys) {
                var cell = celdaActiva.offset(i + 1, j);
                if(key == 'id'){
                    cell.setValue(String(content[i][key]));
                }else{
                    cell.setValue(content[i][key]);
                }
                j++;
            }
        }

        PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
    } else {
        PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        var errorResponse = response.getResponseCode();
        Logger.log("Error response: " + errorResponse);
        mapeoErroresDocumento(errorResponse);
    }
}

function listarRegistroFecha(documento, fechaInicial, fechaFin){
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'
    
    var payloadDoc = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 5000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": documento,
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
                    "atributo": "fecha",
                    "tipoDato": 3,
                    "nombreColumna": "Fecha",
                    "tipoFiltro": 8,
                    "valor": fechaInicial,
                    "valor2": fechaFin,
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
            'payload': JSON.stringify(payloadDoc),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
    
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content.length);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
    
            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };
    
            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }
    
            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell = celdaActiva.offset(i + 1, j);
                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErroresDocumento(errorResponse);
        }
    
    
}

function listarRegistroNumero(documento, numeroInicial,numeroFin){
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'
    
    var payloadDoc = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 5000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": documento,
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
                    "tipoFiltro": 8,
                    "valor": numeroInicial,
                    "valor2": numeroFin,
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
            'payload': JSON.stringify(payloadDoc),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
    
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content.length);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
    
            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };
    
            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }
    
            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell = celdaActiva.offset(i + 1, j);
                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErroresDocumento(errorResponse);
        }
    
    
}

function listarRegistroPrefijo(documento, prefijo){
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'
    
    var payloadDoc = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 5000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": documento,
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
                    "atributo": "prefijo.nombre",
                    "tipoDato": 0,
                    "nombreColumna": "Prefijo",
                    "tipoFiltro": 1,
                    "valor": prefijo,
                    "operador": 0
                },
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
            'payload': JSON.stringify(payloadDoc),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
    
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content.length);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
    
            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };
    
            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }
    
            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell = celdaActiva.offset(i + 1, j);
                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErroresDocumento(errorResponse);
        }
    
    
}

function listarRegistroTerceroEmpresa(documento, empresa){
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'
    
    var payloadDoc = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 5000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": documento,
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
                    "atributo": "empresa.nombre",
                    "tipoDato": 0,
                    "nombreColumna": "Empresa",
                    "tipoFiltro": 1,
                    "valor": empresa,
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
            'payload': JSON.stringify(payloadDoc),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
    
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content.length);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
    
            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };
    
            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }
    
            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell = celdaActiva.offset(i + 1, j);
                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErroresDocumento(errorResponse);
        }
    
    
}

function listarRegistroTerceroCliente(documento, cliente){
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'
    
    var payloadDoc = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 5000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": documento,
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
                    "atributo": "terceroExterno.nombreCompleto",
                    "tipoDato": 0,
                    "nombreColumna": "Cliente",
                    "tipoFiltro": 1,
                    "valor": cliente,
                    "operador": 0
                },
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
            'payload': JSON.stringify(payloadDoc),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
    
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content.length);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
    
            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };
    
            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }
    
            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell = celdaActiva.offset(i + 1, j);
                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErroresDocumento(errorResponse);
        }
    
    
}

function listarRegistroTerceroVendedor(documento, vendedor){
    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'
    
    var payloadDoc = {
            "columnaOrdenar": "fecha,id",
            "pagina": 0,
            "registrosPorPagina": 5000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "documentoTipo.codigoDocumento",
                    "valor": documento,
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
                    "atributo": "terceroInterno.nombreCompleto",
                    "tipoDato": 0,
                    "nombreColumna": "Vendedor",
                    "tipoFiltro": 1,
                    "valor": vendedor,
                    "operador": 0
                },
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
            'payload': JSON.stringify(payloadDoc),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);
        Logger.log(response);
        Logger.log(response.length);
    
        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content.length);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
    
            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };
    
            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }
    
            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell = celdaActiva.offset(i + 1, j);
                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErroresDocumento(errorResponse);
        }
    
    
}

// function mostraDocumentosVenta(){
//     var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionPaginadoDocumentos');
//     var datos = JSON.parse(datosString);
//     Logger.log(datos);
//     var totalRegistros = getDataTotalRegistros();
//     Logger.log(totalRegistros);
//     var formatoTotalRegistros = Number(totalRegistros);
//     Logger.log(formatoTotalRegistros);
//     // var claveAPI = almacenamientoClave();

//     // var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'

//     // var payloadFV = {
//     //     "columnaOrdenar": "fecha,id",
//     //     "pagina": 0,
//     //     "registrosPorPagina": 2000,
//     //     "orden": "DESC",
//     //     "filtros": [
//     //         {
//     //             "atributo": "documentoTipo.codigoDocumento",
//     //             "valor": datos.tipoDocumento,
//     //             "valor2": null,
//     //             "tipoFiltro": 0,
//     //             "tipoDato": 0,
//     //             "nombreColumna": null,
//     //             "valores": null,
//     //             "clase": null,
//     //             "operador": 0,
//     //             "subGrupo": "filtro"
//     //         },
//     //         {
//     //             "atributo": "fecha",
//     //             "tipoDato": 3,
//     //             "nombreColumna": "Fecha",
//     //             "tipoFiltro": 8,
//     //             "valor": datos.registroInicial,
//     //             "valor2": datos.registroFinal,
//     //             "operador": 0
//     //         }
//     //     ],
//     //     "canal": 0,
//     //     "registroInicial": 0
//     // };


//     // var headers = {
//     //     'Content-Type': 'application/json',
//     //     'Authorization': claveAPI,
//     // };

//     // var options = {
//     //     'method': 'post',
//     //     'headers': headers,
//     //     'payload': JSON.stringify(payloadFV),
//     //     'muteHttpExceptions': true
//     // };

//     // var response = UrlFetchApp.fetch(apiUrl, options);
//     // Logger.log(response);
//     // Logger.log(response.length);

//     if (response.getResponseCode() === 200) {
//         // var responseData = response.getContentText();
//         // var jsonData = JSON.parse(responseData);
//         // var content = jsonData.data.content;
//         // Logger.log(content.length);
//         // var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//         // var celdaActiva = hojaActiva.getActiveCell();

//         // var keys = {
//         //     "id": "Id",
//         //     "fecha": "Fecha",
//         //     "prefijo": "Prefijo",
//         //     "numero": "Número",
//         //     "empresa": "Empresa",
//         //     "terceroExterno": "Cliente",
//         //     "terceroInterno": "Vendedor",
//         //     "formaPago": "Forma de Pago",
//         //     "concepto": "Concepto",
//         // };

//         // var j = 0;
//         // for (var key in keys) {
//         //     var headerCell = celdaActiva.offset(0, j);
//         //     headerCell.setValue(keys[key]);
//         //     j++;
//         // }

//         // for (var i = 0; i < content.length; i++) {
//         //     j = 0;
//         //     for (var key in keys) {
//         //         var cell = celdaActiva.offset(i + 1, j);
//         //         if(key == 'id'){
//         //             cell.setValue(String(content[i][key]));
//         //         }else{
//         //             cell.setValue(content[i][key]);
//         //         }
//         //         j++;
//         //     }
//         // }
//         PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
//     } else {
//         PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', '');
//         var errorResponse = response.getResponseCode();
//         Logger.log("Error response: " + errorResponse);
//         mapeoErroresDocumento(errorResponse);
//     }


// }



////////////////////// Consultar documento venta por identificacion //////////////////////

function consultarDocumentoVenta(){
    var selectConsultarDocVenta = 'consultarDocumentoVenta';
    viewGeneralDocumentos(selectConsultarDocVenta);
}

function guardarSeleccionDocumentoVenta(seleccion) {
    var seleccionDocumentoVenta = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoVenta', seleccionDocumentoVenta);
}

function mostrarDatosConsultaDocumentoVenta() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionDocumentoVenta');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idDocumento = datos.idDocumento;
    var validacionImportarEncabezado = datos.importarEncabezados;
    var tipoDocumento = datos.tipoDocumento;

    if(idDocumento && tipoDocumento) {
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
                    "valor": datos.tipoDocumento,
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
                    "valor": datos.idDocumento,
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
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "fecha": "Fecha",
                "prefijo": "Prefijo",
                "numero": "Número",
                "empresa": "Empresa",
                "terceroExterno": "Cliente",
                "terceroInterno": "Vendedor",
                "formaPago": "Forma de Pago",
                "concepto": "Concepto",
            };


            if(validacionImportarEncabezado){
                var j = 0;
                for (var key in keys) {
                    var headerCell = celdaActiva.offset(0, j);
                    headerCell.setValue(keys[key]);
                    j++;
                }
            }


            for (var i = 0; i < content.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell;
                    if(validacionImportarEncabezado == true){
                        cell = celdaActiva.offset(i + 1, j);
                    } else {
                        cell = celdaActiva.offset(i, j);
                    }
                    if(key == 'senPrincipal'){
                        if(content[i][key] == true){
                            cell.setValue('Si');
                        }else{
                            cell.setValue('No');
                        }
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoVenta', '');
        }else {
            PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoVenta', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErroresDocumento(errorResponse);
        }

    }

    PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoVenta', '');

}

////////////////////// Enviar documento por Email //////////////////////

function consultarEnvioEmailDocumento(){
    var enviarEmailDocumento = 'enviarEmailDocumento';
    viewGeneralDocumentos(enviarEmailDocumento);
}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneralDocumentos(select){
    if(select == 'listarDocVentas'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/documentos/documento-listado-ventas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(820)
        .setHeight(500);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Documentos de Venta');
    }else if(select == 'consultarDocumentoVenta'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/documentos/documento-consulta-ventas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(490);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar Documentos de Venta');
    }else if(select == 'enviarEmailDocumento'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/documentos/documento-enviar-email.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(620);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Enviar documento por email');
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