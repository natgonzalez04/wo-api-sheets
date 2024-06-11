////////////////////// Funciones Inventarios //////////////////////

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

////////////////////// Listar Inventarios //////////////////////

function listarInventarios(){

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/inventarios/listarInventarios';
    
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
}

////////////////////// Consultar Inventarios //////////////////////

function consultarInventarios(){
    var selectConsultarInventarios = 'consultarInventarios';
    viewGeneralInventarios(selectConsultarInventarios);
}

function guardarSeleccionInventarios(seleccion) {
    var seleccionInventario = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionInventarioCriterio', seleccionInventario);
}

function mostrarDatosConsultaInventarios() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionInventarioCriterio');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idInventario = datos.idInventario;
    var validacionImportarEncabezado = datos.importarEncabezados;
    var codigoInventario = datos.codigoInventario;

    if(idInventario && !codigoInventario) {
        inventarioId(idInventario, validacionImportarEncabezado);

    } else if(codigoInventario && !idInventario){
        inventarioCodigo(codigoInventario, validacionImportarEncabezado);
    }
    
    PropertiesService.getDocumentProperties().setProperty('seleccionInventarioCriterio', '');

}

function inventarioId(id, validacion){

        var claveAPI = almacenamientoClave();
        
        var apiUrl = `https://api.worldoffice.cloud/api/v1/inventarios/consultaId/${id}`;
    
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
                "codigo": "Código",
                "descripcion": "Descripción",
                "inventarioClasificacion": "Clasificación",
                "unidadMedida": "Unidad Medida", 
                "senActivo": "Estado",
                "impuestosValor": "% Impuesto Ventas",
                "impuestosInventarioTipoImpuestoVenta": "Tipo Impuesto Venta",
                "senFacturarSinExistencias": "Facturar Sin Existencias",
                "senManejaLotes": "Maneja Lotes",
                "senManejaTallaColor": "Maneja Talla Color",
                "senManejaSeriales": "Maneja Seriales"
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
                    if(key == 'inventarioClasificacion'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'unidadMedida'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'senActivo'){
                        if(content[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'impuestosValor'){
                        cell.setValue(content[i]['impuestos'][0].valor);
                    }else if(key == 'impuestosInventarioTipoImpuestoVenta'){
                        cell.setValue(content[i]['impuestos'][0].inventarioTipoImpuestoVenta.nombre);
                    }else if(key == 'senFacturarSinExistencias' || key == 'senManejaLotes' || key == 'senManejaTallaColor' || key == 'senManejaSeriales'){
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
            PropertiesService.getDocumentProperties().setProperty('seleccionInventarioCriterio', '');
        }else {   
            PropertiesService.getDocumentProperties().setProperty('seleccionInventarioCriterio', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
        }
}

function inventarioCodigo(codigo, validacion){

    var claveAPI = almacenamientoClave();

    var apiUrl = `https://api.worldoffice.cloud/api/v1/inventarios/consultaCodigo/${codigo}`;
    
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
                "descripcion": "Descripción",
                "inventarioClasificacion": "Clasificación",
                "unidadMedida": "Unidad Medida", 
                "senActivo": "Estado",
                "impuestosValor": "% Impuesto Ventas",
                "impuestosInventarioTipoImpuestoVenta": "Tipo Impuesto Venta",
                "senFacturarSinExistencias": "Facturar Sin Existencias",
                "senManejaLotes": "Maneja Lotes",
                "senManejaTallaColor": "Maneja Talla Color",
                "senManejaSeriales": "Maneja Seriales"
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
            
                    if(key == 'inventarioClasificacion'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'unidadMedida'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'senActivo'){
                        if(content[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'impuestosValor'){
                        cell.setValue(content[i]['impuestos'][0].valor);
                    }else if(key == 'impuestosInventarioTipoImpuestoVenta'){
                        cell.setValue(content[i]['impuestos'][0].inventarioTipoImpuestoVenta.nombre);
                    }else if(key == 'senFacturarSinExistencias' || key == 'senManejaLotes' || key == 'senManejaTallaColor' || key == 'senManejaSeriales'){
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
            PropertiesService.getDocumentProperties().setProperty('seleccionInventarioCriterio', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionInventarioCriterio', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            mapeoErrores(errorResponse);
        }
}

//////////////////////// Listar Clasificación Inventarios //////////////////////

function listarClasificacionInventarios(){
    var selectListarClasificacion = 'listarClasificacionInventarios';
    viewGeneralInventarios(selectListarClasificacion);
}

function mostrarDatosClasificacionInventarios(){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/inventarios/clasificaciones';

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
            var dataEntrance = jsonData.data;
            var content = dataEntrance.content;
            Logger.log(content);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "codigo": "Código",
                "nombre": "Nombre",
                "senAfectaCostoPromedio": "Afecta Costo Promedio",
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
                    var cell;
                    cell = celdaActiva.offset(i + 1, j); 

                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else if(key == 'senAfectaCostoPromedio'){
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

        }else{
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
            
        }
}

//////////////////////// Listar Bodegas Inventarios //////////////////////

function listarBodegasInventarios(){
    var selectListarBodegas = 'listarBodegasInventarios';
    viewGeneralInventarios(selectListarBodegas);
}

function mostrarDatosBodegasInventarios(){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/bodegas/listarBodega';

        var headers = {
            'Content-Type': 'application/json',
            'Authorization': claveAPI,
        };

        var payloadBodega = {
            "columnaOrdenar": "id",
            "pagina": 0,
            "registrosPorPagina": 1000,
            "orden": "ASC",
            "filtros": [],
            "canal": 0,
            "registroInicial": 0
        };

        var options = {
            'method': 'post',
            'headers': headers,
            'payload': JSON.stringify(payloadBodega),
            'muteHttpExceptions': true
        };
    
        var response = UrlFetchApp.fetch(apiUrl, options);

        if (response.getResponseCode() === 200) {
            var responseData = response.getContentText();
            var jsonData = JSON.parse(responseData);
            var dataEntrance = jsonData.data;
            var content = dataEntrance.content;
            Logger.log(content);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "nombre": "Nombre",
                "senActiva": "Activo",
                "senPredeterminada": "Predeterminada",
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
                    var cell;
                    cell = celdaActiva.offset(i + 1, j); 

                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else if(key == 'senPredeterminada' || key == 'senActiva'){
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

        }else{
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
            
        }
}

//////////////////////// Listar Grupos Inventarios //////////////////////

function listarGruposInventarios(){
    var selectListarGrupos = 'listarGruposInventarios';
    viewGeneralInventarios(selectListarGrupos);
}

function mostrarDatosGruposInventarios(){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/inventarios/grupos';

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
            var dataEntrance = jsonData.data;
            var content = dataEntrance.content;
            Logger.log(content);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "codigo": "Código",
                "nombreGrupo": "Nombre",
                "senActivo": "Activo",
                "cantidadHijos": "Subniveles"
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
                    var cell;
                    cell = celdaActiva.offset(i + 1, j); 

                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else if(key == 'senActivo'){
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

        }else{
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
            
        }
}

//////////////////////// Listar Impúestos de venta por Inventario  //////////////////////

function listarImpuestosVentaInventarios(){
    var selectListarImpuestosVenta = 'listarImpuestosVentaInventarios';
    viewGeneralInventarios(selectListarImpuestosVenta);
}

function mostrarDatosImpuestosVentaInventarios(){
        var claveAPI = almacenamientoClave();

        var apiUrl = 'https://api.worldoffice.cloud/api/v1/inventarios/impuestosDeVentas';

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
            var dataEntrance = jsonData.data;
            var content = dataEntrance.content;
            Logger.log(content);
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();

            var keys = {
                "id": "Id",
                "nombre": "Nombre", 
                "tipo": "Tipo"
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
                    var cell;
                    cell = celdaActiva.offset(i + 1, j); 

                    if(key == 'id'){
                        cell.setValue(String(content[i][key]));
                    }else{
                        cell.setValue(content[i][key]); 
                    }
                    j++;
                }
            }

        }else{
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
            
        }
}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneralInventarios(select){
    if(select == 'listarInventarios'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/documentos/documento-listado-ventas.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(720)
        .setHeight(500);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Inventario');
    } else if(select == 'consultarInventarios'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/inventarios/inventarios-consultar-inventarios.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(820)
        .setHeight(500);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar Inventario');
    }else if(select == 'listarClasificacionInventarios'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/inventarios/inventarios-clasificacion-inventarios.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(620)
        .setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Clasificación Inventario');
    }else if(select == 'listarBodegasInventarios'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/inventarios/inventarios-bodegas-inventarios.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(620)
        .setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Bodegas Inventario');
    }else if(select == 'listarGruposInventarios'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/inventarios/inventarios-grupos-inventarios.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(620)
        .setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Grupos Inventario');
    }else if(select == 'listarImpuestosVentaInventarios'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/inventarios/inventarios-impuestos-ventas-inventarios.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(620)
        .setHeight(300);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar Impuestos Venta por Inventario');
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