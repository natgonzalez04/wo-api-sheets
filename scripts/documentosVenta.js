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

function listarDocumentosVenta(){
    var selectListarDocVentas = 'listarDocVentas';
    viewGeneralDocumentos(selectListarDocVentas);
}

function guardarSeleccionPaginadoDocumentos(seleccion) {
    var seleccionPaginadoDoc = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionPaginadoDocumentos', seleccionPaginadoDoc);
}

function mostraDocumentosVenta(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionPaginadoDocumentos');
    var datos = JSON.parse(datosString);
    Logger.log(datos.registroInicial);

    var claveAPI = almacenamientoClave();
    
    var apiUrl = 'https://api.worldoffice.cloud/api/v1/documentos/listarDocumentoVenta'

    var payloadFV = {
        "columnaOrdenar": "fecha,id",
        "pagina": 0,
        "registrosPorPagina": 50,
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


////////////////////// Llamado de vistas general //////////////////////

function viewGeneralDocumentos(select){
    if(select == 'listarDocVentas'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/documentos/documento-listado-ventas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(490);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Documentos de Venta');
    }
}

////////////////////// Errores de respuesta //////////////////////

function mapeoErroresDocumento(response){
    if(response == '403'){
        var htmlOutput = HtmlService.createHtmlOutput('<p style="font-family: Raleway, sans-serif; font-size:14px; text-align: center; margin:-25px 0 10px 0;"><span style="color:#2196F3; font-size: 48px;">&#9888;</span><br><br>No tienes permisos de consulta para este servicio, puedes activarlos desde tu cuenta de World Office cloud</p>')
        .setWidth(430)
        .setHeight(120);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
    }    
}