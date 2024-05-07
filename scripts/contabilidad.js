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

////////////////////// Consultar documento compra por identificacion //////////////////////

function consultarDocumentoContabilidad(){
    var selectConsultarDocContabilidad = 'consultarDocumentoContabilidad';
    viewGeneralDocumentosContabilidad(selectConsultarDocContabilidad);
}

function guardarSeleccionDocumentoContabilidad(seleccion) {
    var seleccionDocumentoCompra = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoContabilidad', seleccionDocumentoCompra);
}

function mostrarDatosConsultaDocumentoContabilidad() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionDocumentoContabilidad');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idDocumento = datos.idDocumento;
    var validacionImportarEncabezado = datos.importarEncabezados;
    var tipoDocumento = datos.tipoDocumento;

    if(idDocumento && tipoDocumento) {
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
                "terceroExterno": tipoDocumento == 'CE' ? "Beneficiario" : "Recibido de",
                "terceroInterno": tipoDocumento == 'NC' ? "Tercero" : "Elaborado por",
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
            PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoContabilidad', '');
        }else {
            PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoContabilidad', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);

            mapeoErroresDocumentoContabilidad(errorResponse);
        }

    }

    PropertiesService.getDocumentProperties().setProperty('seleccionDocumentoContabilidad', '');

}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneralDocumentosContabilidad(select){
    if(select == 'consultarDocumentoContabilidad'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/contabilidad/documento-consulta-contabilidad.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(750)
        .setHeight(490);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar Documentos de Contabilidad');
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
