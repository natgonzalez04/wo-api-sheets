////////////////////// Funciones cuentas Contables //////////////////////

var ui = SpreadsheetApp.getUi();

var payload = {
    "columnaOrdenar": "id",
    "pagina": 0,
    "registrosPorPagina": 4000,
    "orden": "DESC",
    "filtros": [],
    "canal": 0,
    "registroInicial": 0
};

////////////////////// Listar Cuentas Contables //////////////////////

function listarCuentasContables(){

    var selectListarCuentasContables = 'listarCuentasContables';
    viewGeneralCuentasContables(selectListarCuentasContables);

}

function consultaCuentasInicial(){
    var datosString = PropertiesService.getDocumentProperties().getProperty('dataDocumentoGeneralCuenta');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var categoria = datos.categoriaDocumento
    var numeroCuenta = datos.numeroCuenta

    if(categoria && !numeroCuenta){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/cuentasContables/listarCuentaContable';

        var payloadContable = {
            "columnaOrdenar": "codigo",
            "pagina": 0,
            "registrosPorPagina": 2000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "cuentaContableTipo.nombre",
                    "valor": categoria,
                    "valor2": null,
                    "tipoFiltro": 1,
                    "tipoDato": 0,
                    "nombreColumna": "Tipo",
                    "valores": null,
                    "clase": null,
                    "operador": 0,
                    "subGrupo": "filtro"
                },
                {
                    "atributo": "senActivo",
                    "tipoDato": 1,
                    "nombreColumna": "Activo",
                    "tipoFiltro": 0,
                    "valor": true,
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
            'payload': JSON.stringify(payloadContable),
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);

        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content);

            var scriptPropertiesCuentasContables = PropertiesService.getDocumentProperties();
            scriptPropertiesCuentasContables.setProperty('contentCuentaContable', JSON.stringify(content));

            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');
        }
        else
        {
            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');
            var errorResponse = response.getContentText();
            Logger.log("Error response: " + errorResponse);
        }

        PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');
    }else if(numeroCuenta && !categoria){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/cuentasContables/listarCuentaContable';

        var payloadContable = {
            "columnaOrdenar": "codigo",
            "pagina": 0,
            "registrosPorPagina": 2000,
            "orden": "DESC",
            "filtros": [
                {
                    "atributo": "codigo",
                    "valor": numeroCuenta,
                    "valor2": null,
                    "tipoFiltro": 1,
                    "tipoDato": 0,
                    "nombreColumna": "Código",
                    "operador": 0,
                },
                {
                    "atributo": "senActivo",
                    "tipoDato": 1,
                    "nombreColumna": "Activo",
                    "tipoFiltro": 0,
                    "valor": true,
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
            'payload': JSON.stringify(payloadContable),
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(apiUrl, options);

        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = jsonData.data.content;
            Logger.log(content);

            var scriptPropertiesCuentasContables = PropertiesService.getDocumentProperties();
            scriptPropertiesCuentasContables.setProperty('contentCuentaContable', JSON.stringify(content));

            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');
        }
        else
        {
            PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');
            var errorResponse = response.getContentText();
            Logger.log("Error response: " + errorResponse);
        }

        PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');
    }

    PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', '');

}

function guardarSeleccionGeneralCuenta(seleccion) {
    var documentoGeneralCuenta = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('dataDocumentoGeneralCuenta', documentoGeneralCuenta);
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

function deletePropertyCuentaContable() {
    var scriptProperties = PropertiesService.getDocumentProperties();
    scriptProperties.deleteProperty('contentCuentaContable');
}

function mostrarDatosCeldaCuentaContable() {

    var datosString = PropertiesService.getUserProperties().getProperty('seleccionCeldaCuentaContable');
    var datos = JSON.parse(datosString);
    datos = datos.map(subarray => subarray.filter(item => item !== '' && item !== null));
    datosRecibidos = datos.flat();

    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    var codigosCuentaSelect = datosRecibidos.filter(function(item, index) {
        return index % 2 === 0;
    });

    var dataCuentaContable = getDataCuentaContable();

    var dataCuentaContableFiltrada = dataCuentaContable.filter(function(item) {
        return codigosCuentaSelect.includes(item.codigo);
    })


            var keys = {
                "id": "Id",
                "codigo": "Codigo",
                "nombre": "Nombre",
                "subCuentaContable": "Subcuenta",
                "cuentaContableTipo": "Tipo",
                "cuentaContableGrupo":"Grupo",
                "senManejaCentroCosto": "Centro Costo",
                "senActivo": "Estado",
                "senVisible": "Visible",
                "senAjustePorInflacion":"Ajuste Por Inflación"
            };

            var j = 0;
            for (var key in keys) {
                var headerCell = celdaActiva.offset(0, j);
                headerCell.setValue(keys[key]);
                j++;
            }

            for (var i = 0; i < dataCuentaContableFiltrada.length; i++) {
                j = 0;
                for (var key in keys) {
                    var cell;
                    cell = celdaActiva.offset(i + 1, j);

                    if(key == 'subCuentaContable'){
                        cell.setValue(dataCuentaContableFiltrada[i][key].codigo);
                    }else if(key == 'cuentaContableTipo'|| key == 'cuentaContableGrupo'){
                        cell.setValue(dataCuentaContableFiltrada[i][key].nombre);
                    }else if(key == 'senActivo'){
                        if(dataCuentaContableFiltrada[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'senVisible' || key == 'senAjustePorInflacion' || key == 'senManejaCentroCosto'){
                        if(dataCuentaContableFiltrada[i][key] == true){
                            cell.setValue('Si');
                        }else{
                            cell.setValue('No');
                        }
                    }else if(key == 'id'){
                        cell.setValue(String(dataCuentaContableFiltrada[i][key]));
                    }else{
                        cell.setValue(dataCuentaContableFiltrada[i][key]);
                    }
                    j++;
                }
            }

        PropertiesService.getUserProperties().setProperty('seleccionCeldaCuentaContable', '');
        PropertiesService.getUserProperties().setProperty('optionDataCuentaContable', '');

}

////////////////////// Consultar Cuentas Contables //////////////////////

function consultarCuentasContables(){
    var selectConsultarCuentasContables = 'consultarCuentasContables';
    viewGeneralCuentasContables(selectConsultarCuentasContables);
}

function guardarSeleccionCuentasContables(seleccion) {
    var seleccionCuentasContables = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionConsultaCuentaCriterio', seleccionCuentasContables);
}

function mostrarDatosConsultaCuentasContables() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionConsultaCuentaCriterio');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idCuentaContable = datos.idCuentaContable;
    var validacionImportarEncabezado = datos.importarEncabezados;
    var codigoCuentaContable = datos.codigoCuentaContable;

    if(idCuentaContable&& !codigoCuentaContable) {
        cuentasContablesId(idCuentaContable, validacionImportarEncabezado);

    } else if(codigoCuentaContable && !idCuentaContable){
        cuentasContablesCodigo(codigoCuentaContable, validacionImportarEncabezado);
    }

    PropertiesService.getDocumentProperties().setProperty('seleccionConsultaCuentaCriterio', '');

}

function cuentasContablesId(id, validacion){

        var claveAPI = almacenamientoClave();

        var apiUrl = `https://api.worldoffice.cloud/api/v1/cuentasContables/consultar/${id}`;

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

        if (response.getResponseCode() === 202) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = [jsonData.data];
            Logger.log(content);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "codigo": "Codigo",
                "nombre": "Nombre",
                "subCuentaContable": "Subcuenta",
                "cuentaContableTipo": "Tipo",
                "cuentaContableGrupo":"Grupo",
                "senManejaCentroCosto": "Centro Costo",
                "senActivo": "Estado",
                "senVisible": "Visible",
                "senAjustePorInflacion":"Ajuste Por Inflación"
            };


            if(validacion){
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
                    if(validacion == true){
                        cell = celdaActiva.offset(i + 1, j);
                    } else {
                        cell = celdaActiva.offset(i, j);
                    }

                    if(key == 'subCuentaContable'){
                        cell.setValue(content[i][key].codigo);
                    }else if(key == 'cuentaContableTipo'|| key == 'cuentaContableGrupo'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'senActivo'){
                        if(content[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'senVisible' || key == 'senAjustePorInflacion' || key == 'senManejaCentroCosto'){
                        if(content[i][key] == true){
                            cell.setValue('Si');
                        }else{
                            cell.setValue('No');
                        }
                    }else if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionConsultaCuentaCriterio', '');
        }else {
            PropertiesService.getDocumentProperties().setProperty('seleccionConsultaCuentaCriterio', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErrores(errorResponse);
        }
}

function cuentasContablesCodigo(codigo, validacion){

    var claveAPI = almacenamientoClave();

    var apiUrl = `https://api.worldoffice.cloud/api/v1/cuentasContables/consultaCodigo/${codigo}`;

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

        if (response.getResponseCode() === 202) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var content = [jsonData.data];
            Logger.log(content);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "codigo": "Codigo",
                "nombre": "Nombre",
                "subCuentaContable": "Subcuenta",
                "cuentaContableTipo": "Tipo",
                "cuentaContableGrupo":"Grupo",
                "senManejaCentroCosto": "Centro Costo",
                "senActivo": "Estado",
                "senVisible": "Visible",
                "senAjustePorInflacion":"Ajuste Por Inflación"
            };


            if(validacion){
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
                    if(validacion == true){
                        cell = celdaActiva.offset(i + 1, j);
                    } else {
                        cell = celdaActiva.offset(i, j);
                    }

                    if(key == 'subCuentaContable'){
                        cell.setValue(content[i][key].codigo);
                    }else if(key == 'cuentaContableTipo'|| key == 'cuentaContableGrupo'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'senActivo'){
                        if(content[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'senVisible' || key == 'senAjustePorInflacion' || key == 'senManejaCentroCosto'){
                        if(content[i][key] == true){
                            cell.setValue('Si');
                        }else{
                            cell.setValue('No');
                        }
                    }else if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]);
                    }
                    j++;
                }
            }
            PropertiesService.getDocumentProperties().setProperty('seleccionConsultaCuentaCriterio', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionConsultaCuentaCriterio', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErrores(errorResponse);
        }
}


////////////////////// Llamado de vistas general //////////////////////

function viewGeneralCuentasContables(select){
    if(select == 'listarCuentasContables'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/cuentasContables/cuentasContables-listar-cuentas-contables.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(610);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Cuentas Contables');
    }else if(select == 'consultarCuentasContables'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/cuentasContables/cuentasContables-consultar.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(470);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar Cuenta Contable');
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