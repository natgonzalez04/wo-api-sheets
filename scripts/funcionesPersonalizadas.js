////////////////////////////////////////// funciones personalizadas contables //////////////////////////////////////////////

/**
 * Calcula el IVA de un precio dado.
 *
 * @param {number} precio El precio sin IVA.
 * @return {number} El precio con el IVA incluido.
 * @customfunction
 */
function CALCULARIVA(precio) {
    var tasaIVA = 0.19; // Tasa del IVA del 21%
    return precio * (1 + tasaIVA);
}

/**
 * Devuelve el saldo de una cuenta en la fecha especificada, para la empresa y centro de costos seleccionado.
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaMovimiento Fecha del movimiento DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionTercero Número de identificación del tercero formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del saldo en cuenta.
 * @customfunction
 */
function SALDOCUENTA(codigoCuentaContable, identificacionTercero, fechaMovimiento, centroCostos) {
    // Convertir la fecha a un formato legible (p.ej., DD/MM/YYYY). Ajusta el formato según necesites.
    var fechaFormateada = Utilities.formatDate(new Date(fechaMovimiento), Session.getScriptTimeZone(), "dd/MM/yyyy");
    // Concatenar y devolver la información
    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Tercero: ${identificacionTercero}, Fecha Movimiento: ${fechaFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 1.000.000`;
    // var saldo = 1000000;
    // return saldo;
}


//////////////////////////////////////// funcion cuentas por cobrar /////////////////////////////////////////

/**
 * Devuelve el saldo de cuentas por cobrar de un tercero y centro de costos especifico.
 *
 * @param {890963456} identificacionTercero Número de identificación del tercero formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del saldo por cobrar en cuenta.
 * @customfunction
 */
function CUENTASXCOBRAR(identificacionTercero, centroCostos) {
    // Convertir la fecha a un formato legible (p.ej., DD/MM/YYYY). Ajusta el formato según necesites.
    // Concatenar y devolver la información
    return `Número Identificación Tercero: ${identificacionTercero}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 2.000.000`;
    // var saldo = 1000000;
    // return saldo;
}

//////////////////////////////////////// funcion cuentas por pagar /////////////////////////////////////////

/**
 * Devuelve el saldo de cuentas por pagar de un tercero y centro de costos especifico.
 *
 * @param {890963456} identificacionTercero Número de identificación del tercero formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del saldo por pagar en cuenta.
 * @customfunction
 */
function CUENTASXPAGAR(identificacionTercero, centroCostos) {
    // Convertir la fecha a un formato legible (p.ej., DD/MM/YYYY). Ajusta el formato según necesites.

    // Concatenar y devolver la información
    return `Número Identificación Tercero: ${identificacionTercero}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 3.000.000`;
    // var saldo = 1000000;
    // return saldo;
}


///////////////////////////////////////// funcion de movimiento debito /////////////////////////////////////////

/**
 * Devuelve el movimiento debito de una cuenta contable entre una fecha inicial y final para la empresa y centro de costos seleccionado.
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento debito en cuenta.
 * @customfunction
 */
function MOVIMIENTODEBITO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 7.500.000`;

}

/**
 * Devuelve el movimiento crédito de una cuenta contable entre una fecha inicial y final para la empresa y centro de costos seleccionado
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento crédito DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento crédito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento crédito en cuenta.
 * @customfunction
 */
function MOVIMIENTOCREDITO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 8.500.000`;

}

/**
 * Devuelve el total de movimiento debito-crédito de una cuenta contable entre una fecha inicial y final para la empresa y centro de costos seleccionado
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento en cuenta.
 * @customfunction
 */
function MOVIMIENTOSALDO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Tercero: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 9.500.000`;

}

/**
 * Devuelve el movimiento debito de una cuenta para un tercero especifico entre una fecha inicial y final para la empresa y centro de costos seleccionado.
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {890963456} identificacionTercero Número de identificación del tercero formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento debito en cuenta.
 * @customfunction
 */
function MOVIMIENTOTERCERODEBITO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, identificacionTercero, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Tercero: ${identificacionTercero}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 7.500.000`;

}

/**
 * Devuelve el movimiento crédito de una cuenta para un tercero especifico entre una fecha inicial y final para la empresa y centro de costos seleccionado
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento crédito DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento crédito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {890963456} identificacionTercero Número de identificación del tercero formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento crédito en cuenta.
 * @customfunction
 */
function MOVIMIENTOTERCEROCREDITO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, identificacionTercero, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Tercero: ${identificacionTercero}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 8.500.000`;
    
}

/**
 * Devuelve el total de movimiento debito-crédito de una cuenta para un tercero especifico entre una fecha inicial y final para la empresa y centro de costos seleccionado
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {890963456} identificacionTercero Número de identificación del tercero formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento en cuenta.
 * @customfunction
 */
function MOVIMIENTOTERCEROSALDO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, identificacionTercero, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Tercero: ${identificacionTercero}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 9.500.000`;

}


/**
 * Devuelve el movimiento presupuestado de una cuenta contable entre una fecha inicial y final para la empresa seleccionada.
 *
 * @param {110505} codigoCuentaContable Código de la cuenta contable formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento debito en cuenta.
 * @customfunction
 */
function MOVIMIENTOPRESUPUESTO(codigoCuentaContable, fechaInicio, fechaFin, identificacionEmpresa, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Código Cuenta Contable: ${codigoCuentaContable}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 2.500.000`;

}

/**
 * Devuelve el valor de ventas de un producto entre una fecha inicial y final para la empresa y centro de costos seleccionado. Puede elegir entre el valor antes de IVA o incluido.
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {15/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {iva} tipoIVA tipo IVA formato Texto
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento debito en cuenta.
 * @customfunction
 */
function VENTASPRODUCTO(codigoInventario, fechaInicio, fechaFin, identificacionEmpresa, tipoIVA , centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Codigo Inventario: ${codigoInventario}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Tipo IVA: ${tipoIVA}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 5.500.000`;

}

/**
 * Devuelve el valor de ventas de un producto sin incluir devoluciones, entre una fecha inicial y final para la empresa y centro de costos seleccionado. Puede elegir entre el valor antes de IVA o incluido
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {17/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {iva} tipoIVA tipo IVA formato Texto
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve el valor del movimiento debito en cuenta.
 * @customfunction
 */
function VENTASPRODUCTOSINDEVOLUCIONES(codigoInventario, fechaInicio, fechaFin, identificacionEmpresa, tipoIVA , centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Codigo Inventario: ${codigoInventario}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Tipo IVA: ${tipoIVA}, Centro de Costos: ${centroCostos}, Saldo Cuenta: 3.500.000`;

}

/**
 * Devuelve las cantidades vendidas de un producto, entre una fecha inicial y final para la empresa y centro de costos seleccionado.
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {18/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve las cantidades vendidas de un producto.
 * @customfunction
 */
function CANTIDADESVENDIDASPRODUCTO(codigoInventario, fechaInicio, fechaFin, identificacionEmpresa , centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Codigo Inventario: ${codigoInventario}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada}, Centro de Costos: ${centroCostos}, cantidades: 2.500`;

}

/**
 * Devuelve las cantidades vendidas de un producto sin incluir devoluciones, entre una fecha inicial y final para la empresa y centro de costos seleccionado
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @param {15/06/2024} fechaInicio Fecha inicio del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {16/06/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {Bogota-Centro} centroCostos Centro de costos formato Texto
 * @return {string} Devuelve las cantidades vendidas de un producto sin incluir devoluciones.
 * @customfunction
 */
function CANTIDADESVENDIDASINDEVOLUCIONES(codigoInventario, fechaInicio, fechaFin, identificacionEmpresa, centroCostos) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Codigo Inventario: ${codigoInventario}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Inicio: ${fechaInicioFormateada}, Fecha Fin: ${fechaFinFormateada},Centro de Costos: ${centroCostos}, cantidades: 1.500`;

}

/**
 * Devuelve la cantidad de existencias de un producto a una fecha de corte y empresa especifica.
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @param {16/08/2024} fechaFin Fecha fin del movimiento debito DD/MM/AAAA formato Fecha.
 * @param {890963456} identificacionEmpresa Número de identificación de la empresa formato Texto. 
 * @param {PRINCIPAL1} bodega Código de la bodega formato Texto. 
 * @param {APT123} lote Código del lote formato Texto. 
 * @return {string} Devuelve la cantidad de existencias de un producto.
 * @customfunction
 */
function CANTIDADESEXISTENCIASPRODUCTO(codigoInventario, fechaFin, identificacionEmpresa, bodega, lote) {

    var fechaInicioFormateada = Utilities.formatDate(new Date(fechaInicio), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var fechaFinFormateada = Utilities.formatDate(new Date(fechaFin), Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Codigo Inventario: ${codigoInventario}, Número Identificación Empresa: ${identificacionEmpresa}, Fecha Fin: ${fechaFinFormateada},Bodega: ${bodega}, Lote: ${lote}, existencias: 500`;

}

/**
 * Devuelve el precio de venta de un producto de acuerdo a su lista de precios 1, 2, 3 y 4.
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @param {CONTADO} listaprecio Nombre lista de precio formato Texto.
 * @return {string} Devuelve el precio de venta de un producto.
 * @customfunction
 */
function PRECIOVENTAPRODUCTO(codigoInventario, listaprecio) {

    return `Codigo Inventario: ${codigoInventario}, Listaprecio: ${listaprecio}, precio: 4.100`;

}

/**
 * Devuelve el costo promedio de un producto
 *
 * @param {110505} codigoInventario Código del inventario formato Texto.
 * @return {string} Devuelve el costo promedio de un producto
 * @customfunction
 */
function COSTOPROMEDIO(codigoInventario) {

    return `Codigo Inventario: ${codigoInventario}, precio: 3.444`;

}

/**
 * Convierte un valor numérico en un valor en letras - español
 *
 * @param {4.200.000} numero valor numerico en celda formato Texto - Numero.
 * @return {string} Devuelve el valor numérico en un valor en letras - español.
 * @customfunction
 */
function NUMEROALETRASESPANOL(numero) {

    const units = ['Cero', 'Uno', 'Dos', 'Tres', 'Cuatro', 'Cinco', 'Seis', 'Siete', 'Ocho', 'Nueve'];
    const teens = ['Once', 'Doce', 'Trece', 'Catorce', 'Quince', 'Dieciséis', 'Diecisiete', 'Dieciocho', 'Diecinueve'];
    const tens = ['Diez', 'Veinte', 'Treinta', 'Cuarenta', 'Cincuenta', 'Sesenta', 'Setenta', 'Ochenta', 'Noventa'];
    const hundreds = ['Cien', 'Doscientos', 'Trescientos', 'Cuatrocientos', 'Quinientos', 'Seiscientos', 'Setecientos', 'Ochocientos', 'Novecientos'];
    const thousands = ['Mil', 'Millón', 'Mil Millones', 'Billón'];

    if (numero < 10) return units[numero];
    if (numero < 20) return teens[numero - 11];
    if (numero < 100) return tens[Math.floor(numero / 10) - 1] + (numero % 10 !== 0 ? ' y ' + units[numero % 10] : '');
    if (numero < 1000) return (numero < 200 ? 'Ciento' : hundreds[Math.floor(numero / 100) - 1]) + (numero % 100 !== 0 ? ' ' + NUMEROALETRASESPANOL(numero % 100) : '');

    for (let i = 0; i < thousands.length; i++) {
        const unit = 1000 ** (i + 1);
        if (numero < unit * 1000) {
            return NUMEROALETRASESPANOL(Math.floor(numero / unit)) + ' ' + thousands[i] + (numero % unit !== 0 ? ' ' + NUMEROALETRASESPANOL(numero % unit) : '');
        }
    }
}

/**
 * Convierte un valor numérico en un valor en letras - ingles
 *
 * @param {3.200.000} numero valor numerico en celda formato Texto - Numero.
 * @return {string} Devuelve el valor numérico en un valor en letras - ingles.
 * @customfunction
 */
function NUMEROALETRASINGLES(numero) {

    const units = ['Zero', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['Ten', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    const thousands = ['Thousand', 'Million', 'Billion', 'Trillion'];

    if (numero < 10) return units[numero];
    if (numero < 20) return teens[numero - 11];
    if (numero < 100) return tens[Math.floor(numero / 10) - 1] + (numero % 10 !== 0 ? ' ' + units[numero % 10] : '');
    if (numero < 1000) return units[Math.floor(numero / 100)] + ' Hundred' + (numero % 100 !== 0 ? ' ' + NUMEROALETRASINGLES(numero % 100) : '');

    for (let i = 0; i < thousands.length; i++) {
        const unit = 1000 ** (i + 1);
        if (numero < unit * 1000) {
            return NUMEROALETRASINGLES(Math.floor(numero / unit)) + ' ' + thousands[i] + (numero % unit !== 0 ? ' ' + NUMEROALETRASINGLES(numero % unit) : '');
        }
    }

}

/**
 * Devuelve el valor de un campo del catalogo de terceros
 *
 * @param {890456789} identificacionTercero Identificación del tercero formato Texto.
 * @param {nombreTercero} campoTercero Nombre del campo a consultar (nombreTercero, codigo, tipo, estado, ciudad, tipoId, nomina) formato Texto.
 * @return {string} Devuelve el valor del campo solicitado.
 * @customfunction
 */
function INFORMACIONCAMPOTERCERO(identificacionTercero, campoTercero) {

    return `identificacion Tercero: ${identificacionTercero}, campoTercero: ${campoTercero}`;

}

/**
 * Devuelve el valor de un campo del catalogo de inventarios
 *
 * @param {ADT1254} codigoInventario Código del inventario formato Texto.
 * @param {descripcion} campoInventario Nombre del campo a consultar (descripcion, unidad, tipoImpuesto, clasificacion, estado, ciudad, porcentajeImpuesto, facturarExisitencias, manejaLotes, manejaSeriales) formato Texto.
 * @return {string} Devuelve el valor del campo solicitado.
 * @customfunction
 */
function INFORMACIONCAMPOINVENTARIO(codigoInventario, campoInventario) {

    return `Codigo Inventario: ${codigoInventario}, campoInventario: ${campoInventario}`;

}

/**
 * Devuelve el valor de un campo del catalogo de cuentas contables
 *
 * @param {110254} codigoCuenta Código de la cuenta formato Texto.
 * @param {nombreCuenta} campoCuenta Nombre del campo a consultar (nombreCuenta, subcuenta, tipo, grupo, estado) formato Texto.
 * @return {string} Devuelve el valor del campo solicitado.
 * @customfunction
 */
function INFORMACIONCAMPOCUENTA(codigoCuenta, campoCuenta) {

    return `Codigo Cuenta: ${codigoCuenta}, campoCuenta: ${campoCuenta}`;

}

/**
 * Devuelve el valor de un campo del catalogo de empresas
 *
 * @param {890962789} identificacionEmpresa Identificación de la empresa formato Texto.
 * @param {nombreEmpresa} campoEmpresa Nombre del campo a consultar (nombreEmpresa, codigo, tipo, estado, ciudad, tipoId, nomina) formato Texto.
 * @return {string} Devuelve el valor del campo solicitado.
 * @customfunction
 */
function INFORMACIONCAMPOEMPRESA(identificacionEmpresa, campoEmpresa) {

    return `identificacion Empresa: ${identificacionEmpresa}, campoEmpresa: ${campoEmpresa}`;

}

/**
 * Devuelve el valor de un campo del catalogo de centro de costos
 *
 * @param {RTY125} codigoCentrocosto Código del inventario formato Texto.
 * @param {nombre} campoCentrocosto Nombre del campo a consultar (nombre, estado) formato Texto.
 * @return {string} Devuelve el valor del campo solicitado.
 * @customfunction
 */
function INFORMACIONCAMPOCENTROCOSTO(codigoCentrocosto, campoCentrocosto ) {

    return `Codigo Centrocosto: ${codigoCentrocosto}, campoCentrocosto: ${campoCentrocosto}`;

}


///////////////////////////// Funciones de Apertura //////////////////////////////////////////////////////


function abrirSaldoCuenta() {
    // var infoCelda = obtenerInformacionCeldaActiva();
    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-saldo-cuenta.html').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(690);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar saldo cuenta');
}

function abrirSaldoCuentasPagarCobrar() {
    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/saldo-cuentas-por-pagar-cobrar.html').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(290);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar saldo cuenta x cobrar y pagar');
}


function abrirMovimientos() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-movimientos.html').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(690);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar movimientos');
}


function abrirMovimientosTercero() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-movimientos-tercero.html').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(690);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar movimientos tercero');
}


function abrirMovimientosPresupuesto() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-movimientos-presupuesto.html').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(690);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar movimientos presupuesto');
}


function abrirVentasProducto() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-ventas-producto.html').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(490);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar ventas producto');
}

function abrirCantidadesVentasProducto() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-cantidades-ventas-producto').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(420);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar cantidades vendidas producto');
}

function abrirExistenciasProducto() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-existencias-producto').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(320);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar existencias producto');
}

function abrirPrecioVentaProducto() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-precio-venta-producto').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(240);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar precio venta producto');
}

function abrirCostoPromedioProducto() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/consultar-costo-promedio-producto').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Consultar costo promedio producto');
}

function abrirConvertirValorNumerico() {

    var htmlOutputView = HtmlService.createHtmlOutputFromFile('views/funcionesPersonalizadas/convertir-valor-numerico').getContent();
    var htmlOutputStyle = HtmlService.createHtmlOutputFromFile('styles/style.html').getContent();
    var htmlOutputComplete = HtmlService.createHtmlOutput(htmlOutputView + htmlOutputStyle)
    .setWidth(750)
    .setHeight(330);
    SpreadsheetApp.getUi().showModalDialog(htmlOutputComplete, 'Transformar valor numerico a letras');
}

function abrirInformacionCampo(){

}





///////////////////////////// funcition edit /////////////////////////////////

// /**
//  * Se ejecuta automáticamente cada vez que una celda es editada.
//  *
//  * @param {Object} e El objeto de evento que contiene información sobre la celda editada.
//  */
// function onEdit(e) {
//     // Set a comment on the edited cell to indicate when it was changed.
//     const range = e.range;
//     const formulaAccion = range.getFormula();
//     const valueEdit = range.getValue();
//     // range.setNote('Last modified: ' + valueEdit);
//     if (formulaAccion.includes('SALDOCUENTA(')) {
//         // La celda contiene la fórmula "calcularSaldo()", realiza acciones específicas aquí
//         range.setNote(valueEdit);
//         // abrirSaldoCuenta();
//     }
// }