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

////////////////////// Listar tipos de identificacion //////////////////////

function listarTiposIdentificacion(){
    var selectTiposIdentificacion = 'listarTiposIdentificacion';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/tiposIdentificacion';
    
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
        viewGeneralTerceros(selectTiposIdentificacion);
        var scriptPropertiesTiposId= PropertiesService.getScriptProperties();
        scriptPropertiesTiposId.setProperty('contentTipoIdentificacion', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionTipoIdentificacion(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoIdentificacion', seleccionString);
}

function getDataTipoIdentificacion() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataTipoIdentificacion = scriptProperties.getProperty('contentTipoIdentificacion');
    return JSON.parse(contentDataTipoIdentificacion);
}

function guardarSeleccionOpcionTipoIdentificacion(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataTipoIdentificacion', seleccionOptionData);
}

function mostrarDatosCeldaTipoIdentificacion() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataTipoIdentificacion');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaTipoIdentificacion');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoIdentificacion', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoIdentificacion', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoIdentificacion', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoIdentificacion', '');
    }
}

////////////////////// Listar tipos de contribuyente //////////////////////

function listarTiposContribuyente(){
    var selectTipoContribuyente = 'listarTipoContribuyente';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/tiposContribuyente';
    
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
        viewGeneralTerceros(selectTipoContribuyente);
        var scriptPropertiesTipoContribuyente = PropertiesService.getScriptProperties();
        scriptPropertiesTipoContribuyente.setProperty('contentTipoContribuyente', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionTipoContribuyente(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoContribuyente', seleccionString);
}

function getDataTipoContribuyente() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataTipoContribuyente = scriptProperties.getProperty('contentTipoContribuyente');
    return JSON.parse(contentDataTipoContribuyente);
}

function guardarSeleccionOpcionTipoContribuyente(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataTipoContribuyente', seleccionOptionData);
}

function mostrarDatosCeldaTipoContribuyente() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataTipoContribuyente');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaTipoContribuyente');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoContribuyente', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoContribuyente', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoContribuyente', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoContribuyente', '');
    }
}

////////////////////// Listar tipos de tercero //////////////////////

function listarTiposTercero(){
    var selectListarTiposTerceros = 'listarTiposTerceros';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/terceroTipos';
    
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
        viewGeneralTerceros(selectListarTiposTerceros);
        var scriptPropertiesTipoTercero = PropertiesService.getScriptProperties();
        scriptPropertiesTipoTercero.setProperty('contentTipoTercero', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionTipoTercero(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoTercero', seleccionString);
}

function getDataTipoTercero() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataTipoTercero = scriptProperties.getProperty('contentTipoTercero');
    return JSON.parse(contentDataTipoTercero);
}

function guardarSeleccionOpcionTipoTercero(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataTipoTercero', seleccionOptionData);
}

function mostrarDatosCeldaTipoTercero() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataTipoTercero');
    var option = JSON.parse(selectOption);
    Logger.log(option);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaTipoTercero');
    var datos = JSON.parse(datosString);
    Logger.log(datos);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoTercero', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoTercero', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaTipoTercero', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataTipoTercero', '');
    }
}



////////////////////// listar clasificacion impuestos //////////////////////

function listarClasificacionImpuestos(){
    var selectClasificacionImpuestos = 'listarClasificacionImpuestos';

    var claveAPI = almacenamientoClave();

    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/clasificacionImpuestos';
    
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
        viewGeneralTerceros(selectClasificacionImpuestos);
        var scriptPropertiesClasificacionImpuestos = PropertiesService.getScriptProperties();
        scriptPropertiesClasificacionImpuestos.setProperty('contentClasificacionImpuestos', JSON.stringify(content));
    } 
    else 
    {
        var errorResponse = response.getContentText();
        Logger.log("Error response: " + errorResponse);
    }

}

function guardarSeleccionClasificacionImpuestos(seleccion) {
    var seleccionString = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionCeldaClasificacionImpuestos', seleccionString);
}

function getDataClasificacionImpuestos() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var contentDataClasificacionImpuestos = scriptProperties.getProperty('contentClasificacionImpuestos');
    return JSON.parse(contentDataClasificacionImpuestos);
}

function guardarSeleccionOpcionClasificacionImpuestos(option) {
    var seleccionOptionData = JSON.stringify(option);
    PropertiesService.getDocumentProperties().setProperty('optionDataClasificacionImpuestos', seleccionOptionData);
}

function mostrarDatosCeldaClasificacionImpuestos() {
    var selectOption = PropertiesService.getDocumentProperties().getProperty('optionDataClasificacionImpuestos');
    var option = JSON.parse(selectOption);
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionCeldaClasificacionImpuestos');
    var datos = JSON.parse(datosString);
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaActiva = hojaActiva.getActiveCell();

    if(option == 'listar'){
        for (var i = 0; i < datos.length; i++) {
            var celda = celdaActiva.offset(i,0);
            celda.setValue(datos[i]);
        }
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaClasificacionImpuestos', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataClasificacionImpuestos', '');
    }else if(option == 'menu'){
        var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(datos)
        .build();

        celdaActiva.setDataValidation(regla);
        PropertiesService.getDocumentProperties().setProperty('seleccionCeldaClasificacionImpuestos', '');
        PropertiesService.getDocumentProperties().setProperty('optionDataClasificacionImpuestos', '');
    }
}
////////////////////// Consultar tercero //////////////////////

function consultarTerceros(){
    var selectConsultarTerceros = 'consultarTerceros';
    viewGeneralTerceros(selectConsultarTerceros);
}

function guardarSeleccionTercero(seleccion) {
    var seleccionTercero = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionTerceroCriterio', seleccionTercero);
}

function mostrarDatosConsultaTercero() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionTerceroCriterio');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idUsuario = datos.idTercero;
    var validacionImportarEncabezado = datos.importarEncabezados;
    var idIdentificacion = datos.identificacion;

    if(idUsuario && !idIdentificacion) {
        terceroId(idUsuario, validacionImportarEncabezado);

    } else if(idIdentificacion && !idUsuario){
        terceroIdentificacion(idIdentificacion, validacionImportarEncabezado);
    }
    
    PropertiesService.getDocumentProperties().setProperty('seleccionTerceroCriterio', '');

}

function terceroId(id, validacion){

        var claveAPI = almacenamientoClave();
        
        var apiUrl = `https://api.worldoffice.cloud/api/v1/terceros/consultar/${id}`;
    
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
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
            
            var keys = {
                "id": "Id",
                "terceroTipoIdentificacion": "Tipo Identificación",
                "identificacion": "Identificación",
                "nombreCompleto": "Nombre Completo",
                "ciudad": "Ubicación", 
                "terceroTipos": "Tipo Tercero",
                "senActivo": "Estado",
                "aplicaICAVentas": "Aplica ICA" 
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
                    if(key == 'terceroTipoIdentificacion'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'terceroTipos'){
                        cell.setValue(content[i][key][0].nombre);
                    }else if(key == 'ciudad'){
                        cell.setValue(content[i][key].ciudadNombre + ' - ' + content[i][key].ubicacionDepartamento.nombre);
                    }else if(key == 'senActivo'){
                        if(content[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'aplicaICAVentas'){
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
            PropertiesService.getDocumentProperties().setProperty('seleccionTerceroCriterio', '');
        }else {   
            PropertiesService.getDocumentProperties().setProperty('seleccionTerceroCriterio', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
        }
}

function terceroIdentificacion(identificacion, validacion){

    var claveAPI = almacenamientoClave();

    var apiUrl = `https://api.worldoffice.cloud/api/v1/terceros/identificacion/${identificacion}`;
    
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
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
            
            var keys = {
                "id": "Id",
                "terceroTipoIdentificacion": "Tipo Identificación",
                "identificacion": "Identificación",
                "nombreCompleto": "Nombre Completo",
                "ciudad": "Ubicación", 
                "terceroTipos": "Tipo Tercero",
                "senActivo": "Estado",
                "aplicaICAVentas": "Aplica ICA" 
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
            
                    if(key == 'terceroTipoIdentificacion'){
                        cell.setValue(content[i][key].nombre);
                    }else if(key == 'terceroTipos'){
                        cell.setValue(content[i][key][0].nombre);
                    }else if(key == 'ciudad'){
                        cell.setValue(content[i][key].ciudadNombre + ' - ' + content[i][key].ubicacionDepartamento.nombre);
                    }else if(key == 'senActivo'){
                        if(content[i][key] == true){
                            cell.setValue('Activo');
                        }else{
                            cell.setValue('Inactivo');
                        }
                    }else if(key == 'aplicaICAVentas'){
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
            PropertiesService.getDocumentProperties().setProperty('seleccionTerceroCriterio', '');
        } else {
            PropertiesService.getDocumentProperties().setProperty('seleccionTerceroCriterio', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
        }
}

////////////////////// Consultar tercero direccion //////////////////////

function consultarTercerosDireccion(){
    var selectConsultarTerceros = 'consultarTercerosDireccion';
    viewGeneralTerceros(selectConsultarTerceros);
}

function guardarSeleccionTerceroDireccion(seleccion) {
    var seleccionTerceroDireccion = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionTerceroDireccion', seleccionTerceroDireccion);
}

function mostrarDatosConsultaTerceroDireccion() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionTerceroDireccion');
    var datos = JSON.parse(datosString);
    Logger.log(datos);

    var idUsuario = datos.idTercero;
    var validacionImportarEncabezado = datos.importarEncabezados;

    if(idUsuario) {
        var claveAPI = almacenamientoClave();
        var apiUrl = `https://api.worldoffice.cloud/api/v1/terceros/direccion/${idUsuario}`;

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
            var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
            var celdaActiva = hojaActiva.getActiveCell();
            
            var keys = {
                "direccion": "Dirección",
                "ciudad": "Ciudad", 
                "telefonoPrincipal": "Telefono principal",
                "emailPrincipal": "Email principal",
                "sucursal": "Nombre de dirección o sucursal",
                "senPrincipal":"Principal",
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
            PropertiesService.getDocumentProperties().setProperty('seleccionTerceroDireccion', '');
        }else {   
            PropertiesService.getDocumentProperties().setProperty('seleccionTerceroDireccion', '');
            var errorResponse = response.getResponseCode();
            Logger.log("Error response: " + errorResponse);
            Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
            mapeoErrores(errorResponse);
        }

    } 
    
    PropertiesService.getDocumentProperties().setProperty('seleccionTerceroDireccion', '');

}

////////////////////// listar Tercero //////////////////////

function listarTerceros(){
    var selectListarTerceros = 'listarTerceros';
    viewGeneralTerceros(selectListarTerceros);
}

function guardarSeleccionPaginado(seleccion) {
    var seleccionPaginado = JSON.stringify(seleccion);
    PropertiesService.getDocumentProperties().setProperty('seleccionPaginado', seleccionPaginado);
}

function mostrarDatosPaginadoTercero() {
    var datosString = PropertiesService.getDocumentProperties().getProperty('seleccionPaginado');
    var datos = JSON.parse(datosString);
    Logger.log(datos.registroInicial);
    
    var claveAPI = almacenamientoClave();
    
    var apiUrl = 'https://api.worldoffice.cloud/api/v1/terceros/listarTerceros'

    var payload = {
        "columnaOrdenar": "id",
        "pagina": 0,
        "registrosPorPagina": datos.registroFinal,
        "orden": "DESC",
        "filtros": [],
        "canal": 0,
        "registroInicial": datos.registroInicial
    };
    
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
        var content = jsonData.data.content;
        Logger.log(content.length);
        var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        var celdaActiva = hojaActiva.getActiveCell();
        
        var keys = {
            "id": "Id",
            "tipoIdentificacion": "Tipo Identificación",
            "identificacion": "Identificación",
            "nombreCompleto": "Nombre Completo",
            "ubicacionCiudad": "Ubicación", 
            "codigo": "Código",
            "terceroTipos": "Tipo Tercero",
            "senActivo": "Estado",
            "aplicaICAVentas": "Aplica ICA" 
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
                cell.setValue(content[i][key]);
                j++;
            }
        }
    } 
    else 
    {
        var errorResponse = response.getResponseCode();
        Logger.log("Error response: " + errorResponse);
        Logger.log(errorResponse.errorCode,'asdasdasdasdsad');
        mapeoErrores(errorResponse);
    }
}

////////////////////// Llamado de vistas general //////////////////////

function viewGeneralTerceros(select){
    if(select == 'listarTiposIdentificacion'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-tipos-identificacion.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Tipos identificación');
    }else if(select == 'listarTipoContribuyente'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-tipos-contribuyente.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Tipos contribuyente');
    }else if(select == 'listarTiposTerceros'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-tipos-terceros.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Tipos terceros');
    }else if(select == 'listarClasificacionImpuestos'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-clasificacion-impuestos.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(400)
        .setHeight(370);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Clasificación impuestos');
    }else if(select  == 'listarTerceros'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-listar-terceros.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(700)
        .setHeight(470);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Listar terceros');
    } else if(select  == 'consultarTerceros'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-consultar-terceros.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(700)
        .setHeight(470);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar tercero');
    }  else if(select  == 'consultarTercerosDireccion'){
        var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/terceros/tercero-direccion-terceros.html').getContent();
        var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
        var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
        .setWidth(700)
        .setHeight(470);
        SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar dirección tercero');
    } 
}

////////////////////// Errores de respuesta //////////////////////

function mapeoErrores(response){
    if(response == '403'){
        var htmlOutput = HtmlService.createHtmlOutput('<p style="font-family: Raleway, sans-serif; font-size:14px; text-align: center; margin:-25px 0 10px 0;"><span style="color:#2196F3; font-size: 48px;">&#9888;</span><br><br>No tienes permisos de consulta para este servicio, puedes activarlos desde tu cuenta de World Office cloud</p>')
        .setWidth(430)
        .setHeight(120);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Error');
    }    
}