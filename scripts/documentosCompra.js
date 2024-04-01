////////////////////// Funciones documentos Venta //////////////////////

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

////////////////////// Listar Documento Ventas //////////////////////

function listarDocumentosCompra(){
    var selectListarDocVentas = 'listarDocCompras';
    viewGeneralDocumentosCompra(selectListarDocVentas);
}

function guardarSeleccionPaginadoDocumentosCompra(seleccion) {
    var seleccionPaginadoDoc = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentosCompra', seleccionPaginadoDoc);
}

function mostraDocumentosCompra(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionPaginadoDocumentosCompra');
    var datos = JSON.parse(datosString);
    Logger.log(datos.registroInicial);

    var claveAPI = almacenamientoClave();
    
    var apiUrl = 'https://api.worldoffice.cloud/api/v1/compra/listarDocumentoCompra'

    var payloadFC = {
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
                "atributo": "fecha",
                "tipoDato": 3,
                "nombreColumna": "Fecha",
                "tipoFiltro": 8,
                "valor": datos.registroInicial,
                "valor2": datos.registroFinal,
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
        'payload': JSON.stringify(payloadFC),
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
            "terceroExterno": "Proveedor",
            "terceroInterno": "Comprador",
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
        mapeoErroresDocumentoCompra(errorResponse);
    }


}

function pruebaCompra(){
    var claveAPI = almacenamientoClave();
    
    var apiUrl = 'https://api.worldoffice.cloud/api/v1/compra/listarDocumentoCompra'

    var payloadFC = {
        "columnaOrdenar": "fecha,id",
        "pagina": 0,
        "registrosPorPagina": 2000,
        "orden": "DESC",
        "filtros": [
            {
                "atributo": "documentoTipo.codigoDocumento",
                "valor": "NDC",
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
                "valor":"2024-02-09",
                "valor2": "2024-02-10",
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
        'payload': JSON.stringify(payloadFC),
        'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log(response);
    Logger.log(response.length);
    var errorResponse = response.getResponseCode();
    Logger.log(errorResponse);
}


////////////////////// Llamado de vistas general //////////////////////

function viewGeneralDocumentosCompra(select){
    if(select == 'listarDocCompras'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/documentos/documento-listado-compras.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(750)
        .setHeight(490);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Documentos de Compra');
    }
}

////////////////////// Errores de respuesta //////////////////////

function mapeoErroresDocumentoCompra(response){
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