// 🔐 SISTEMA DE SEGURIDAD CON PROPIEDADES DEL SCRIPT
function obtenerPropiedadesEcosistema() {
  const props = PropertiesService.getScriptProperties();
  return {
    SHEET_ID:       props.getProperty('SHEET_ID'),
    FOLDER_ID:      props.getProperty('FOLDER_ID'),
    TELEGRAM_TOKEN: props.getProperty('TELEGRAM_TOKEN'),
    CHAT_ID:        props.getProperty('CHAT_ID'),
    PASSWORD:       props.getProperty('PASSWORD'), // 🔑 Añadido para el Login
    ML_ACCESS_TOKEN:  props.getProperty('ML_ACCESS_TOKEN'),
    ML_REFRESH_TOKEN: props.getProperty('ML_REFRESH_TOKEN'),
    ML_CLIENT_ID:     props.getProperty('ML_CLIENT_ID'),
    ML_CLIENT_SECRET: props.getProperty('ML_CLIENT_SECRET'),
    API_TOKEN:        props.getProperty('API_TOKEN'),
    RESCIND_FOLDER_ID: props.getProperty('RESCIND_FOLDER_ID')
  };
}

// 🔐 VALIDA EL TOKEN DE API RECIBIDO CONTRA EL CONFIGURADO EN PROPERTIES SERVICE
// Sin fallback hardcodeado: si API_TOKEN no está configurado, SIEMPRE rechaza.
function validarTokenApi(credenciales, tokenRecibido) {
  if (!credenciales.API_TOKEN) {
    console.error("API_TOKEN no está configurado en las propiedades del script. Rechazando todas las peticiones.");
    return false;
  }
  return tokenRecibido === credenciales.API_TOKEN;
}

function doGet(e) {
  const credenciales = obtenerPropiedadesEcosistema();
  if (!credenciales.SHEET_ID) {
    return ContentService.createTextOutput(JSON.stringify({error: "Falta configurar la propiedad SHEET_ID en el Script"})).setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.openById(credenciales.SHEET_ID);
  
  // --- ACCIÓN PÚBLICA: OBTENER ESTADO DE CUENTA CLIENTE ---
  if (e.parameter.action === 'obtenerEstadoCuentaCliente') {
    const rawImei = String(e.parameter.imei || "").trim();
    const rawTelefono = String(e.parameter.telefono || "").trim();
    
    const normalizarVal = (val) => {
      if (!val) return "";
      let str = String(val).trim();
      if (str.toLowerCase().includes('e')) {
        let num = Number(val);
        if (!isNaN(num)) str = num.toFixed(0);
      }
      return str.replace(/[^0-9]/g, "");
    };
    
    const imeiNormalizado = normalizarVal(rawImei);
    const telNormalizado = normalizarVal(rawTelefono);
    
    if (!imeiNormalizado) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Falta proporcionar el IMEI" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const sheetClientes = ss.getSheetByName('Clientes');
    const sheetPagos = ss.getSheetByName('Pagos');
    const sheetInv = ss.getSheetByName('Inventario');
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    
    if (!sheetClientes) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Pestaña Clientes no encontrada" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const ultFilaClientes = sheetClientes.getLastRow();
    let clienteEncontrado = null;
    
    if (ultFilaClientes > 1) {
      const datosClientes = sheetClientes.getRange(2, 1, ultFilaClientes - 1, 21).getValues();
      for (let i = 0; i < datosClientes.length; i++) {
        const fila = datosClientes[i];
        const imeiFila = normalizarVal(fila[5]);
        const telFila = normalizarVal(fila[2]);
        
        if (imeiFila === imeiNormalizado) {
          if (telNormalizado && telFila !== telNormalizado) {
            return ContentService.createTextOutput(JSON.stringify({ error: "El número de teléfono no coincide con el registrado en tu cuenta" })).setMimeType(ContentService.MimeType.JSON);
          }
          
          clienteEncontrado = {
            id: String(fila[19] || "").trim(),
            cliente: String(fila[1] || "").trim(),
            telefono: String(fila[2] || "").trim(),
            modelo: String(fila[4] || "").trim(),
            imei: imeiNormalizado,
            cuota: parseFloat(String(fila[10] || "0").replace(/[\$,]/g, "")) || 0,
            totalFinanciado: parseFloat(String(fila[15] || "0").replace(/[\$,]/g, "")) || 0,
            diaRaya: String(fila[18] || "").trim().toUpperCase(),
            tipoPeriodo: String(fila[20] || "SEMANAL").trim().toUpperCase(),
            fechaInicio: fila[0],
            engancheFila: parseFloat(String(fila[6] || "0").replace(/[\$,]/g, "")) || 0
          };
          break;
        }
      }
    }
    
    if (!clienteEncontrado) {
      return ContentService.createTextOutput(JSON.stringify({ error: "No se encontró ningún equipo activo con el IMEI proporcionado" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Obtener precio de contado desde Inventario (columna G / índice 6)
    let precioContado = clienteEncontrado.totalFinanciado;
    if (sheetInv) {
      const dataInv = sheetInv.getDataRange().getValues();
      for (let i = 1; i < dataInv.length; i++) {
        if (normalizarVal(dataInv[i][0]) === imeiNormalizado) {
          precioContado = parseFloat(String(dataInv[i][6]).replace(/[\$,]/g, "")) || precioContado;
          break;
        }
      }
    }
    
    // Obtener configuración global
    const filaConfig = sheetConfig ? sheetConfig.getRange("A2:N2").getValues()[0] : [];
    let interesTasaRaw = parseFloat(filaConfig[0]);
    if (isNaN(interesTasaRaw)) {
      interesTasaRaw = 5;
    } else if (interesTasaRaw > 0 && interesTasaRaw <= 1) {
      interesTasaRaw = interesTasaRaw * 100;
    }
    const interesTasa = interesTasaRaw;
    const semTol = parseFloat(filaConfig[2]) || 3;
    
    // Obtener historial de pagos para este IMEI
    const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    const abonos = [];
    let totalPagado = 0;
    
    for (let j = 1; j < dataPagos.length; j++) {
      const imeiPago = normalizarVal(dataPagos[j][0]);
      if (imeiPago === imeiNormalizado) {
        const fechaPago = dataPagos[j][2];
        const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
        const metodoPago = String(dataPagos[j][4] || "Efectivo").trim();
        
        let fechaFormateada = "";
        if (fechaPago instanceof Date) {
          fechaFormateada = Utilities.formatDate(fechaPago, Session.getScriptTimeZone() || "America/Mexico_City", "dd/MM/yyyy");
        } else if (fechaPago) {
          fechaFormateada = String(fechaPago);
        }
        
        totalPagado += montoPago;
        abonos.push({
          fecha: fechaFormateada,
          monto: montoPago,
          metodo: metodoPago
        });
      }
    }
    
    const saldoPendiente = Math.max(0, clienteEncontrado.totalFinanciado - totalPagado);
    const totalCuotasPactadas = Math.round(clienteEncontrado.totalFinanciado / clienteEncontrado.cuota);
    const cuotasPagadas = Math.min(totalCuotasPactadas, Math.round(totalPagado / clienteEncontrado.cuota));
    const plazoTexto = `${totalCuotasPactadas} PAGOS ${clienteEncontrado.tipoPeriodo}ES`;
    
    const diasPorPeriodo = (clienteEncontrado.tipoPeriodo === "QUINCENAL") ? 15 : 7;
    const startTimestamp = clienteEncontrado.fechaInicio;
    
    const normalizarDia = (diaStr) => {
      const diasIngles = {
        "monday": "LUNES", "tuesday": "MARTES", "wednesday": "MIERCOLES",
        "thursday": "JUEVES", "friday": "VIERNES", "saturday": "SABADO", "sunday": "DOMINGO"
      };
      const normalized = diaStr.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      if (diasIngles[normalized]) return diasIngles[normalized];
      return normalized.toUpperCase();
    };
    
    let proxVencimientoStr = "N/A";
    let semanasAtraso = 0;
    let interesMoratorioAcumulado = 0;
    
    if (startTimestamp instanceof Date) {
      const diasSemana = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
      let indexRaya = diasSemana.indexOf(normalizarDia(clienteEncontrado.diaRaya));
      if (indexRaya === -1) indexRaya = 0;
      
      let primerDiaRaya = new Date(startTimestamp.getTime());
      while (primerDiaRaya.getDay() !== indexRaya) {
        primerDiaRaya.setDate(primerDiaRaya.getDate() + 1);
      }
      primerDiaRaya.setHours(0, 0, 0, 0);
      
      const vencimientoDate = new Date(primerDiaRaya.getTime() + cuotasPagadas * diasPorPeriodo * 24 * 60 * 60 * 1000);
      
      const hoy = new Date();
      hoy.setHours(0, 0, 0, 0);
      
      const diffTime = hoy.getTime() - vencimientoDate.getTime();
      const diasAtraso = Math.max(0, Math.floor(diffTime / (24 * 60 * 60 * 1000)));
      semanasAtraso = Math.floor(diasAtraso / 7);
      interesMoratorioAcumulado = diasAtraso * clienteEncontrado.cuota * (interesTasa / 100);
      
      const diasNombresEs = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
      const diaSemanaEs = diasNombresEs[vencimientoDate.getDay()];
      const scriptTimeZone = Session.getScriptTimeZone() || "America/Mexico_City";
      proxVencimientoStr = `${diaSemanaEs} ${Utilities.formatDate(vencimientoDate, scriptTimeZone, "dd/MM/yyyy")}`;
    }
    
    let metodosPago = "";
    try {
      metodosPago = PropertiesService.getScriptProperties().getProperty('METODOS_PAGO') || "";
      metodosPago = metodosPago.replace(/\\n/g, '\n').replace(/\\r/g, '\r');
    } catch(e) {}
    
    const montoEnganche = clienteEncontrado.engancheFila;
    const saldoPendienteRescision = Math.max(0, Math.round(precioContado - montoEnganche - totalPagado + interesMoratorioAcumulado));
    
    // Determinar estado de cuenta
    let estadoActual = "ACTIVO";
    if (saldoPendiente <= 0) {
      estadoActual = "LIQUIDADO";
    } else if (semanasAtraso >= semTol) {
      estadoActual = "RESCINDIDO";
    }
    
    // Verificar si ya existe un PDF de rescisión en Google Drive para este IMEI (dentro de carpeta Contratos Rescindidos fija)
    let pdfUrlExistente = "";
    if (estadoActual === "RESCINDIDO") {
      try {
        const folder = DriveApp.getFolderById(credenciales.RESCIND_FOLDER_ID);
        const files = folder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          const name = file.getName();
          if (name.startsWith("Rescision_Contrato_") && name.includes(imeiNormalizado) && name.endsWith(".pdf")) {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            pdfUrlExistente = "https://drive.google.com/uc?export=download&id=" + file.getId();
            break;
          }
        }
      } catch(eDriveCheck) {
        console.error("Error buscando PDF rescision existente en Drive: " + eDriveCheck.toString());
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      pdfUrlRescision: pdfUrlExistente,
      id: clienteEncontrado.id,
      cliente: clienteEncontrado.cliente,
      telefono: clienteEncontrado.telefono,
      modelo: clienteEncontrado.modelo,
      imei: clienteEncontrado.imei,
      cuota: Math.round(clienteEncontrado.cuota),
      saldoPendiente: Math.round(saldoPendiente),
      totalFinanciado: Math.round(clienteEncontrado.totalFinanciado),
      totalPagado: Math.round(totalPagado),
      diaRaya: clienteEncontrado.diaRaya,
      tipoPeriodo: clienteEncontrado.tipoPeriodo,
      plazo: plazoTexto,
      progresoTexto: `${cuotasPagadas} de ${totalCuotasPactadas} pagos realizados`,
      fechaProximoPago: proxVencimientoStr,
      metodosPago: metodosPago,
      estadoCuenta: estadoActual,
      
      // Datos adicionales para desglose financiero de rescisión
      precioContado: Math.round(precioContado),
      montoEnganche: Math.round(montoEnganche),
      totalPagosCapital: Math.round(totalPagado), // Lo ya aportado a capital
      interesMoratorioAcumulado: Math.round(interesMoratorioAcumulado),
      semanasAtraso: semanasAtraso,
      interesTasa: interesTasa,
      saldoPendienteRescision: Math.round(saldoPendienteRescision),
      
      abonos: abonos.reverse()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // --- DIAGNOSTIC: OBTENER INVENTARIO COMPLETO ---
  if (e.parameter.action === 'diagnoseInventario') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const sheetInv = ss.getSheetByName('Inventario');
    if (!sheetInv) return ContentService.createTextOutput("Sheet Inventario not found").setMimeType(ContentService.MimeType.TEXT);
    const data = sheetInv.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }
  
  // --- ACCIÓN: REFRESCAR EL ACCESS TOKEN DE MERCADO LIBRE ---
  if (e.parameter.action === 'refrescarTokenAhora') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const token = refrescarTokenML();
    if (token) {
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', token: token })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se pudo refrescar el token' })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- ACCIÓN: DESACTIVAR TRIGGERS ANTIGUOS MIGRADOS A N8N ---
  if (e.parameter.action === 'desactivarTriggersLocales') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    desactivarTriggersAntiguos();
    return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Triggers locales desactivados correctamente' })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: OBTENER ITEMS DEL CATÁLOGO PARA SINCRONIZAR CON MERCADO LIBRE ---
  if (e.parameter.action === 'obtenerItemsCatalogoML') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const sheetCat = ss.getSheetByName('Catalogo');
    const items = [];
    if (sheetCat) {
      const ultFila = sheetCat.getLastRow();
      if (ultFila > 1) {
        const dataCat = sheetCat.getRange(2, 1, ultFila - 1, sheetCat.getLastColumn()).getValues();
        const formulasLinks = sheetCat.getRange(2, 7, ultFila - 1, 1).getFormulas();
        const linksRicos = sheetCat.getRange(2, 7, ultFila - 1, 1).getRichTextValues();
        
        for (let j = 0; j < dataCat.length; j++) {
          const filaReal = j + 2;
          const modelo = String(dataCat[j][1] || '').trim();
          
          const formula = formulasLinks[j] && formulasLinks[j][0] ? formulasLinks[j][0] : "";
          let formulaLink = "";
          if (formula.startsWith("=HYPERLINK")) {
            const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
            if (match) formulaLink = match[1];
          }
          const richLink = linksRicos[j] && linksRicos[j][0] ? linksRicos[j][0].getLinkUrl() : "";
          const link = richLink || formulaLink || String(dataCat[j][6] || '').trim();
          
          // Sólo enviamos aquellos que tienen enlace a Mercado Libre
          if (link && link.indexOf("mercadolibre") !== -1) {
            items.push({
              fila: filaReal,
              modelo: modelo,
              link: link
            });
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify(items)).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: ACTUALIZAR PRECIOS DEL CATÁLOGO DESDE MERCADO LIBRE ---
  if (e.parameter.action === 'sincronizarCatalogoML') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const res = sincronizarPreciosCatalogoCompletoML();
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: VERIFICAR STOCK EN PARALELO PARA EL DASHBOARD ---
  if (e.parameter.action === 'verificarStockGlobal') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const res = verificarStockGlobal();
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: OBTENER CLIENTES PROPENSOS A ATRASO CON PAGO MAÑANA (COBRANZA PREVENTIVA) [ACTUALIZADO] ---
  if (e.parameter.action === 'obtenerClientesCobranzaPreventiva') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }

    const normalizarImei = (val) => {
      if (!val) return "";
      let str = String(val).trim();
      if (str.toLowerCase().includes('e')) {
        let num = Number(val);
        if (!isNaN(num)) {
          str = num.toFixed(0);
        }
      }
      return str.replace(/[^0-9]/g, "");
    };

    const forzarPrueba = e.parameter.forzar === "true";

    const sheetClientes = ss.getSheetByName('Clientes');
    const sheetPagos = ss.getSheetByName('Pagos');
    if (!sheetClientes) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Pestaña Clientes no encontrada" })).setMimeType(ContentService.MimeType.JSON);
    }

    // Calcular día de la semana de mañana en CDMX
    // Usamos desplazamiento UTC de CDMX (-6 horas o -5 en verano, pero getDay() local de Google es UTC por defecto)
    // Para ser 100% seguros, formateamos la fecha en la zona horaria del script
    const timestampHoy = new Date();
    const timestampManana = new Date(timestampHoy.getTime() + 24 * 60 * 60 * 1000);
    const scriptTimeZone = Session.getScriptTimeZone() || "America/Mexico_City";
    
    // Obtenemos el día de mañana (LUNES, MARTES, etc.) mapeando getDay() en la zona horaria destino
    const stringDiaManana = Utilities.formatDate(timestampManana, scriptTimeZone, "EEEE");
    
    const normalizarDia = (diaStr) => {
      const diasIngles = {
        "monday": "LUNES", "tuesday": "MARTES", "wednesday": "MIERCOLES",
        "thursday": "JUEVES", "friday": "VIERNES", "saturday": "SABADO", "sunday": "DOMINGO"
      };
      const normalized = diaStr.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      if (diasIngles[normalized]) return diasIngles[normalized];
      return normalized.toUpperCase();
    };
    
    const diaMananaNorm = normalizarDia(stringDiaManana);

    // Cargar pagos agrupados por IMEI
    const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    const pagosPorImei = {};
    for (let j = 1; j < dataPagos.length; j++) {
      const imeiPago = normalizarImei(dataPagos[j][0]);
      const fechaPago = dataPagos[j][2];
      const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
      if (imeiPago && fechaPago) {
        if (!pagosPorImei[imeiPago]) pagosPorImei[imeiPago] = [];
        pagosPorImei[imeiPago].push({ fecha: new Date(fechaPago), monto: montoPago });
      }
    }

    const ultFilaClientes = sheetClientes.getLastRow();
    const clientesCobranza = [];

    if (ultFilaClientes > 1) {
      const datosClientes = sheetClientes.getRange(2, 1, ultFilaClientes - 1, 21).getValues();
      
      for (let i = 0; i < datosClientes.length; i++) {
        const fila = datosClientes[i];
        const imei = normalizarImei(fila[5]);
        const clienteNombre = String(fila[1] || '').trim();
        const cuotaNum = parseFloat(String(fila[10] || "0").replace(/[\$,]/g, "")) || 0;
        const totalFinanciadoVal = parseFloat(String(fila[15] || "0").replace(/[\$,]/g, "")) || 0;
        const startTimestamp = fila[0];
        const rawDiaRaya = String(fila[18] || '').trim();
        const diaRayaNorm = normalizarDia(rawDiaRaya);
        const tipoPeriodo = String(fila[20] || "SEMANAL").toUpperCase();

        if (!imei || !clienteNombre || cuotaNum <= 0 || !startTimestamp) continue;

        // Sumar todos los abonos
        const abonos = pagosPorImei[imei] || [];
        let totalPagado = 0;
        abonos.forEach(a => totalPagado += a.monto);

        const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);
        if (saldoPendienteBase <= 0) continue; // Cliente ya liquidó

        // Si mañana no es su día de pago, lo omitimos
        if (diaRayaNorm !== diaMananaNorm) continue;

        // Si es quincenal, solo le toca pagar cada 2 semanas (semanas pares desde el inicio del contrato)
        if (tipoPeriodo === "QUINCENAL" && startTimestamp instanceof Date) {
          const semanasTranscurridas = Math.floor((timestampManana - startTimestamp) / (1000 * 60 * 60 * 24 * 7));
          if (semanasTranscurridas % 2 !== 0) {
            continue; // Es la semana de descanso (puente), no le toca pago
          }
        }

        // Calcular días de atraso actuales
        let diasAtrasoVal = 0;
        if (startTimestamp instanceof Date) {
          const diasTranscurridos = Math.floor((timestampHoy - startTimestamp) / (1000 * 60 * 60 * 24));
          const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
          const periodosTranscurridos = Math.floor(diasTranscurridos / diasPorPeriodo);
          const totalEsperado = periodosTranscurridos * cuotaNum;
          if (totalPagado < totalEsperado) {
            const atrasoMonto = totalEsperado - totalPagado;
            diasAtrasoVal = Math.ceil((atrasoMonto / cuotaNum) * diasPorPeriodo);
          }
        }

        // Determinar si "suele pagar tarde"
        let suelePagarTarde = false;
        if (diasAtrasoVal > 0) {
          suelePagarTarde = true; // Si va atrasado hoy, califica
        } else if (abonos.length > 0 && startTimestamp instanceof Date) {
          // Excluir el enganche (pagos en las primeras 24h desde startTimestamp)
          const cuotasReales = abonos.filter(a => {
            const diffHoras = (a.fecha.getTime() - startTimestamp.getTime()) / (1000 * 60 * 60);
            return diffHoras > 24; // Más de 24 horas después del inicio
          });

          if (cuotasReales.length > 0) {
            const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
            
            // Calcular el primer día de raya teórico después del inicio
            const diasSemana = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
            let indexRaya = diasSemana.indexOf(diaRayaNorm);
            if (indexRaya === -1) indexRaya = 0; // Default domingo
            
            let primerDiaRaya = new Date(startTimestamp.getTime());
            while (primerDiaRaya.getDay() !== indexRaya) {
              primerDiaRaya.setDate(primerDiaRaya.getDate() + 1);
            }
            primerDiaRaya.setHours(0, 0, 0, 0);

            // Ordenar cuotas por fecha
            cuotasReales.sort((a, b) => a.fecha - b.fecha);
            
            let abonosTarde = 0;
            for (let k = 0; k < cuotasReales.length; k++) {
              const fechaAbono = cuotasReales[k].fecha;
              // Fecha límite teórica para esta cuota semanal/quincenal
              const fechaLimiteTeorica = new Date(primerDiaRaya.getTime() + k * diasPorPeriodo * 24 * 60 * 60 * 1000);
              
              // Normalizar horas a 0 para comparar fechas puras
              const dReal = new Date(fechaAbono.getTime());
              dReal.setHours(0, 0, 0, 0);
              const dLimite = new Date(fechaLimiteTeorica.getTime());
              dLimite.setHours(0, 0, 0, 0);
              
              // Si se pagó después de la fecha límite
              if (dReal > dLimite) {
                abonosTarde++;
              }
            }
            const porcentajeTarde = abonosTarde / cuotasReales.length;
            if (porcentajeTarde > 0.25) {
              suelePagarTarde = true;
            }
          }
        }

        // Solo reportar los que suelen pagar tarde (o si se fuerza la prueba)
        if (suelePagarTarde || forzarPrueba) {
          const modelo = String(fila[4] || 'tu equipo CelYa').trim();
          const diasSemanaNombres = {
            "LUNES": "lunes", "MARTES": "martes", "MIERCOLES": "miércoles",
            "JUEVES": "jueves", "VIERNES": "viernes", "SABADO": "sábado", "DOMINGO": "domingo"
          };
          const nombreDiaManana = diasSemanaNombres[diaMananaNorm] || diaMananaNorm.toLowerCase();
          
          const mensajeCobranza = `Hola ${clienteNombre}, ¡qué bueno que disfrutas tu equipo ${modelo}! Mañana ${nombreDiaManana} es tu fecha de pago. Recuerda depositar tu cuota de $${cuotaNum} MXN a tiempo para evitar recargos o suspensión.`;

          clientesCobranza.push({
            id: fila[19] ? String(fila[19]).trim() : "C-Y" + (i + 1),
            cliente: clienteNombre,
            telefono: String(fila[2] || '').trim(),
            modelo: modelo,
            cuota: cuotaNum,
            diaRaya: rawDiaRaya,
            diasAtraso: diasAtrasoVal,
            mensajeCobranza: mensajeCobranza
          });
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      diaManana: diaMananaNorm,
      totalClientes: clientesCobranza.length,
      clientes: clientesCobranza,
      telegramChatId: credenciales.CHAT_ID || PropertiesService.getScriptProperties().getProperty("CHAT_ID"),
      telegramBotToken: credenciales.TELEGRAM_TOKEN || PropertiesService.getScriptProperties().getProperty("TELEGRAM_TOKEN")
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: OBTENER EL ID SIGUIENTE PARA LA WEB AL CARGAR ---
  if (e.parameter.action === 'obtenerSiguienteId') {
    const sheetClientes = ss.getSheetByName('Clientes');
    const siguienteId = generarSiguienteIdCelYa(sheetClientes);
    return ContentService.createTextOutput(JSON.stringify({ siguienteId: siguienteId })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: OBTENER COTIZACIÓN POR ID ---
  if (e.parameter.action === 'obtenerCotizacion') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const id = parseInt(e.parameter.id) || 0;
    const sheetCot = ss.getSheetByName('Cotizaciones');
    if (!sheetCot || id <= 1 || id > sheetCot.getLastRow()) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Cotización no encontrada" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const headers = sheetCot.getRange(1, 1, 1, Math.max(25, sheetCot.getLastColumn())).getValues()[0].map(h => String(h || '').toLowerCase().trim());
    const rowData = sheetCot.getRange(id, 1, 1, sheetCot.getLastColumn()).getValues()[0];
    
    const result = {};
    headers.forEach((h, idx) => {
      if (h) {
        result[h] = rowData[idx];
      }
    });
    
    return ContentService.createTextOutput(JSON.stringify({ result: 'success', data: result })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN RÁPIDA: EXTRAER DATOS DE TEXTO VIP ---
  if (e.parameter.action === 'buscarDatosVip') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const sheetClientes = ss.getSheetByName('Clientes');
    const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];
    const imeiBuscado = e.parameter.imeiAnterior ? e.parameter.imeiAnterior.toString().trim() : "";
    
    let respuesta = { success: false };
    
    for (let i = dataClientes.length - 1; i >= 1; i--) {
      if (dataClientes[i][5] && dataClientes[i][5].toString().trim() === imeiBuscado) { // Columna F (IMEI)
        respuesta = {
          success:         true,
          cliente:         dataClientes[i][1], // Columna B (Nombre)
          telefono:        dataClientes[i][2], // Columna C
          fechaNacimiento: dataClientes[i][3], // Columna D
          direccion:       dataClientes[i][7], // Columna H
          telefonoAval:    dataClientes[i][8]  // Columna I
        };
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify(respuesta)).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: VERIFICACIÓN PREVIA DE IMEI ---
  if (e.parameter.action === 'verificarImei') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ existe: false, error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const sheetClientes = ss.getSheetByName('Clientes');
    const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];
    const imeiBuscado = e.parameter.imei ? e.parameter.imei.toString().trim() : "";
    let yaExiste = false;
    let cliente = "";
    
    for (let i = 1; i < dataClientes.length; i++) {
      if (dataClientes[i][5] && dataClientes[i][5].toString().trim() === imeiBuscado) {
        yaExiste = true;
        cliente = dataClientes[i][1] || ""; // Columna B (Nombre del Cliente)
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ existe: yaExiste, cliente: cliente })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: REGISTRAR LOG DESDE EL FRONTEND EN DEBUG_LOGS ---
  if (e.parameter.action === 'registrarLogFrontend') {
    const mensaje = e.parameter.mensaje || "";
    const origen = e.parameter.origen || "Frontend";
    escribirLogDebug("[" + origen + "] " + mensaje);
    return ContentService.createTextOutput(JSON.stringify({ result: 'success' })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: OBTENER CONFIGURACIÓN ACTUAL EN DETALLE ---
  if (e.parameter.action === 'obtenerConfiguracionActual') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    if (!sheetConfig) {
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se encontró la pestaña Configuración" })).setMimeType(ContentService.MimeType.JSON);
    }
    const filaConfig = sheetConfig.getRange("A2:Q2").getValues()[0];
    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      config: {
        interes:           filaConfig[0] !== undefined ? parseFloat(filaConfig[0]) : 6,
        horasTolerancia:   filaConfig[1] !== undefined ? parseFloat(filaConfig[1]) : 48,
        semanasTolerancia: filaConfig[2] !== undefined ? parseFloat(filaConfig[2]) : 6,
        planRapido:        filaConfig[3] !== undefined && filaConfig[3] !== "" ? parseFloat(filaConfig[3]) : 50,
        planComodo:        filaConfig[4] !== undefined && filaConfig[4] !== "" ? parseFloat(filaConfig[4]) : 75,
        gananciaContado:   filaConfig[5] !== undefined && filaConfig[5] !== "" ? parseFloat(filaConfig[5]) : 15,
        engancheSemanal:   filaConfig[6] !== undefined && filaConfig[6] !== "" ? parseFloat(filaConfig[6]) : 20,

        engancheComodo:    filaConfig[8] !== undefined && filaConfig[8] !== "" ? parseFloat(filaConfig[8]) : 30,
        descuentoLiquidacion: filaConfig[9] !== undefined && filaConfig[9] !== "" ? parseFloat(filaConfig[9]) : 15,
        telegramBotToken:  filaConfig[10] !== undefined ? String(filaConfig[10] || '').trim() : "",
        telegramChatId:    filaConfig[11] !== undefined ? String(filaConfig[11] || '').trim() : "",
        tasaInteresQuincenalRapido: filaConfig[12] !== undefined && filaConfig[12] !== "" ? parseFloat(filaConfig[12]) : 45,
        tasaInteresQuincenalComodo: filaConfig[13] !== undefined && filaConfig[13] !== "" ? parseFloat(filaConfig[13]) : 65,
        engancheMinimoQuincenalRapido: filaConfig[14] !== undefined && filaConfig[14] !== "" ? parseFloat(filaConfig[14]) : 25,
        engancheMinimoQuincenalComodo: filaConfig[15] !== undefined && filaConfig[15] !== "" ? parseFloat(filaConfig[15]) : 35,
        codigoDescuento:   filaConfig[16] !== undefined ? String(filaConfig[16] || '').trim() : "",
        metodosPago:       PropertiesService.getScriptProperties().getProperty('METODOS_PAGO') || ""
      }
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: LEER TASAS DESDE CELDAS EXACTAS (D2 Y E2) ---
  if (e.parameter.action === 'obtenerTasasConfig') {
    // 🔄 MIGRACIÓN TEMPORAL DE PRECIO EN INVENTARIO PARA IMEI 493813629210568
    try {
      const tempSheetInv = ss.getSheetByName('Inventario');
      if (tempSheetInv) {
        const tempValues = tempSheetInv.getDataRange().getValues();
        for (let i = 1; i < tempValues.length; i++) {
          const rowImei = String(tempValues[i][0] || '').trim();
          if (rowImei === "493813629210568") {
            const currentPrice = parseFloat(tempValues[i][6]) || 0;
            if (currentPrice === 2199 || currentPrice === 0) {
              tempSheetInv.getRange(i + 1, 7).setValue(2557);
              escribirLogDebug("MIGRACION: Se actualizó precio de IMEI 493813629210568 a 2557");
            }
          }
        }
      }
    } catch (migErr) {
      Logger.log("Error en migración temporal: " + migErr.toString());
    }

    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    if (!sheetConfig) {
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se encontró la pestaña Configuración" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 📊 D2 = Plan Rápido (50) | E2 = Plan Cómodo (75)
    const tasaAnual26 = parseFloat(sheetConfig.getRange("D2").getValue()) || 50; 
    const tasaAnual52 = parseFloat(sheetConfig.getRange("E2").getValue()) || 75;
    
    // 📊 NUEVO: Leemos F2 (Ganancia de contado). Si está vacío, por defecto usa 15%
    const gananciaContadoVal = parseFloat(sheetConfig.getRange("F2").getValue()) || 15;
    
    const engancheSemanalVal = parseFloat(sheetConfig.getRange("G2").getValue()) || 20;

    const engancheComodoVal = parseFloat(sheetConfig.getRange("I2").getValue()) || 30;

    // Leemos las nuevas celdas M2, N2, O2, P2 para quincenales y Q2 para código de descuento
    const tasaQuinRapidoVal = parseFloat(sheetConfig.getRange("M2").getValue()) || 45;
    const tasaQuinComodoVal = parseFloat(sheetConfig.getRange("N2").getValue()) || 65;
    const engQuinRapidoVal = parseFloat(sheetConfig.getRange("O2").getValue()) || 25;
    const engQuinComodoVal = parseFloat(sheetConfig.getRange("P2").getValue()) || 35;
    const codigoDescuentoVal = String(sheetConfig.getRange("Q2").getValue() || '').trim();
    
    // Empaquetamos y enviamos los factores matemáticos reales
    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      interes26: tasaAnual26 / 100,        // 0.50
      interes52: tasaAnual52 / 100,        // 0.75
      enganche26: (tasaAnual26 / 2) / 100, // 0.25 (Mitad de 50% para el enganche)
      enganche52: (tasaAnual52 / 2) / 100, // 0.375 (Mitad de 75% para el enganche)
      gananciaContadoFactor: 1 + (gananciaContadoVal / 100), // Convierte 15 a 1.15
      gananciaContadoPorcentaje: gananciaContadoVal, // Mandamos el "15" limpio para el input del dashboard
      engancheSemanal: engancheSemanalVal / 100,

      engancheComodo: engancheComodoVal / 100,
      interesQuin26: tasaQuinRapidoVal / 100,
      interesQuin52: tasaQuinComodoVal / 100,
      engancheQuin26: engQuinRapidoVal / 100,
      engancheQuin52: engQuinComodoVal / 100,
      codigoDescuento: codigoDescuentoVal
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: OBTENER PLACEHOLDERS DESDE LA PESTAÑA 'Cotizaciones' ---
  if (e.parameter.action === 'obtenerCatalogoPublico') {
    const sheetCot = ss.getSheetByName('Cotizaciones');
    if (!sheetCot) {
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', resultados: [] })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    let gananciaContadoVal = 15;
    let tasaAnual26 = 66;
    let tasaAnual52 = 90;
    if (sheetConfig) {
      const valConfig = sheetConfig.getRange("F2").getValue();
      if (valConfig !== undefined && valConfig !== "") gananciaContadoVal = parseFloat(valConfig) || 15;
      const val26 = sheetConfig.getRange("D2").getValue();
      if (val26 !== undefined && val26 !== "") tasaAnual26 = parseFloat(val26) || 66;
      const val52 = sheetConfig.getRange("E2").getValue();
      if (val52 !== undefined && val52 !== "") tasaAnual52 = parseFloat(val52) || 90;
    }
    const factorGananciaContado = 1 + (gananciaContadoVal / 100);
    const factorTotal26 = 1 + (tasaAnual26 / 100);
    const factorTotal52 = 1 + (tasaAnual52 / 100);

    const sheetCat = ss.getSheetByName('Catalogo');
    let catData = [];
    let catFormulas = [];
    let catRichTexts = [];
    if (sheetCat) {
      const catRange = sheetCat.getDataRange();
      catData = catRange.getValues();
      catFormulas = catRange.getFormulas();
      catRichTexts = catRange.getRichTextValues();
    }

    const cotRange = sheetCot.getDataRange();
    const headersCot = sheetCot.getRange(1, 1, 1, Math.max(25, sheetCot.getLastColumn())).getValues()[0].map(h => String(h || '').toLowerCase().trim());
    const allCot = cotRange.getValues();
    const allCotFormulas = cotRange.getFormulas();
    const allCotRichTexts = cotRange.getRichTextValues();
    const resultados = [];

    const getValFromRow = (row, headerName) => {
      const idx = headersCot.indexOf(headerName.toLowerCase().trim());
      return idx !== -1 ? row[idx] : "";
    };

    const getLinkFromRow = (rowIdx, headerName) => {
      const idx = headersCot.indexOf(headerName.toLowerCase().trim());
      if (idx === -1) return "";
      return extraerLinkDeArrays(allCot, allCotFormulas, allCotRichTexts, rowIdx, idx);
    };

    for (let i = 1; i < allCot.length; i++) {
      const row = allCot[i];
      const cliente = getValFromRow(row, "Cliente") || "";
      const link = getLinkFromRow(i, "Enlace Mercado Libre") || "";
      const color = getValFromRow(row, "Color") || "";
      const memoria = getValFromRow(row, "Capacidad") || "";
      const planElegido = getValFromRow(row, "Plan Elegido") || "";
      const frecuencia = getValFromRow(row, "Frecuencia") || "";
      const idCliente = getValFromRow(row, "ID Cliente") || "";
      const estadoCredito = getValFromRow(row, "Estado de Crédito") || getValFromRow(row, "Estado") || "";
      const estadoCotizacion = getValFromRow(row, "Estado de Cotización") || "";
      const fecha = getValFromRow(row, "Fecha") || "";
      const plazoSolicitado = getValFromRow(row, "Plazo Solicitado") || "";
      const precioContadoCot = parseFloat(getValFromRow(row, "Precio Contado")) || 0;

      if (!cliente) continue;

      // Ignorar cotizaciones ya asignadas
      if (estadoCotizacion.toUpperCase() === "ASIGNADA") continue;

      let modelo = "";
      let precioCosto = 0;
      let foto = "";
      let precioContadoFinalFromSheet = 0;

      if (link && typeof link === "string") {
        let targetMlm = "";
        const match = link.match(/(MLM\-?\d+)/i);
        if (match) targetMlm = match[0].replace("-", "").toUpperCase();

        for (let j = 1; j < catData.length; j++) {
          const catLink = extraerLinkDeArrays(catData, catFormulas, catRichTexts, j, 6);
          let catMlm = "";
          if (catLink) {
            const match2 = catLink.match(/(MLM\-?\d+)/i);
            if (match2) catMlm = match2[0].replace("-", "").toUpperCase();
          }
          if ((targetMlm && catMlm && targetMlm === catMlm) || link.trim() === catLink.trim()) {
            modelo = catData[j][1] || "";
            precioCosto = parseFloat(catData[j][3]) || 0;
            foto = extraerLinkDeArrays(catData, catFormulas, catRichTexts, j, 7) || catData[j][7] || "";
            precioContadoFinalFromSheet = parseFloat(catData[j][11]) || 0;
            break;
          }
        }
      }

      if (!modelo) {
        if (plazoSolicitado) {
          const matchPlazo = plazoSolicitado.match(/\(PEDIDO ESPECIAL:\s*([^)]+)\)/i) || plazoSolicitado.match(/\(([^)]+)\)$/);
          if (matchPlazo) {
            modelo = matchPlazo[1];
          }
        }
      }
      if (!modelo) {
        modelo = "Celular Cotizado";
      }

      let precioContado = precioContadoFinalFromSheet > 0 
        ? precioContadoFinalFromSheet 
        : (precioContadoCot > 0 ? precioContadoCot : Math.round(precioCosto * factorGananciaContado));
      const factorTotal = (planElegido === "52" || planElegido === 52) ? factorTotal52 : factorTotal26;
      const precioFinanciado = Math.round(precioContado * factorTotal);

      resultados.push({
        fila: i + 1,
        cliente: cliente,
        idCliente: idCliente,
        modelo: modelo,
        color: color,
        memoria: memoria,
        link: link,
        foto: foto,
        planElegido: planElegido,
        frecuencia: frecuencia,
        estadoCredito: estadoCredito,
        fecha: fecha,
        precioPlaceholder: precioFinanciado,
        precioContado: precioContado
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ result: 'success', resultados: resultados })).setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: PROXY PARA OBTENER DATOS DE MERCADO LIBRE DESDE GOOGLE SERVERS ---
  if (e.parameter.action === 'obtenerVariantesML') {
    const cleanId = e.parameter.id || "";
    const isProduct = e.parameter.isProduct === "true";
    if (!cleanId) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Falta id" })).setMimeType(ContentService.MimeType.JSON);
    }
    try {
      let response = null;
      if (isProduct) {
        response = fetchConToken("https://api.mercadolibre.com/products/" + cleanId);
      } else {
        response = fetchConToken("https://api.mercadolibre.com/items/" + cleanId);
      }

      let responseText = "";
      let responseCode = 0;
      
      if (response && response.getResponseCode() === 200) {
        responseText = response.getContentText();
        responseCode = 200;
      } else {
        let fallbackResponse = null;
        if (isProduct) {
          fallbackResponse = fetchConToken("https://api.mercadolibre.com/items/" + cleanId);
        } else {
          fallbackResponse = fetchConToken("https://api.mercadolibre.com/products/" + cleanId);
        }
        if (fallbackResponse && fallbackResponse.getResponseCode() === 200) {
          responseText = fallbackResponse.getContentText();
          responseCode = 200;
        }
      }

      if (responseCode === 200) {
        try {
          const parsed = JSON.parse(responseText);
          if (parsed && parsed.catalog_product_id) {
            const prodRes = fetchConToken("https://api.mercadolibre.com/products/" + parsed.catalog_product_id);
            if (prodRes && prodRes.getResponseCode() === 200) {
              const prodData = JSON.parse(prodRes.getContentText());
              if (!prodData.price || !prodData.buy_box_winner) {
                const itemsUrl = "https://api.mercadolibre.com/products/" + parsed.catalog_product_id + "/items";
                const itemsRes = fetchConToken(itemsUrl);
                if (itemsRes && itemsRes.getResponseCode() === 200) {
                  const itemsData = JSON.parse(itemsRes.getContentText());
                  const results = itemsData.results || [];
                  if (results.length > 0) {
                    results.sort(function(a, b) { return (a.price || 0) - (b.price || 0); });
                    const bestItem = results[0];
                    prodData.price = bestItem.price;
                    prodData.buy_box_winner = {
                      item_id: bestItem.item_id || bestItem.id,
                      price: bestItem.price,
                      shipping: bestItem.shipping,
                      seller_id: bestItem.seller_id,
                      condition: bestItem.condition
                    };
                  }
                }
              }
              responseText = JSON.stringify(prodData);
            }
          } else if (parsed && isProduct && (!parsed.price || !parsed.buy_box_winner)) {
            const itemsUrl = "https://api.mercadolibre.com/products/" + cleanId + "/items";
            const itemsRes = fetchConToken(itemsUrl);
            if (itemsRes && itemsRes.getResponseCode() === 200) {
              const itemsData = JSON.parse(itemsRes.getContentText());
              const results = itemsData.results || [];
              if (results.length > 0) {
                results.sort(function(a, b) { return (a.price || 0) - (b.price || 0); });
                const bestItem = results[0];
                parsed.price = bestItem.price;
                parsed.buy_box_winner = {
                  item_id: bestItem.item_id || bestItem.id,
                  price: bestItem.price,
                  shipping: bestItem.shipping,
                  seller_id: bestItem.seller_id,
                  condition: bestItem.condition
                };
                responseText = JSON.stringify(parsed);
              }
            }
          }
        } catch (e) {
          Logger.log("Error auto-resolving catalog product in proxy: " + e.toString());
        }
        return ContentService.createTextOutput(responseText).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ error: "API returned status " + (response ? response.getResponseCode() : "unknown") })).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (err) {
      Logger.log("Error en obtenerVariantesML: " + err.toString());
      return ContentService.createTextOutput(JSON.stringify({ error: "No se pudo obtener la información del producto." })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- ACCIÓN: GENERAR INVENTARIO EN VIVO PARA TU CATALOGO.HTML ---
  if (e.parameter.action === 'obtenerInventarioPublico') {
    const sheetInv = ss.getSheetByName('Inventario');
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    if (!sheetInv || !sheetConfig) {
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Falta pestaña Inventario o Configuración" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 1. Jalamos tus tasas reales de las celdas D2 y E2 para hacer la matemática al vuelo
    const tasaAnual26 = parseFloat(sheetConfig.getRange("D2").getValue()) || 66;
    const tasaAnual52 = parseFloat(sheetConfig.getRange("E2").getValue()) || 90;
    
    const factorEnganche26 = (tasaAnual26 / 2) / 100; // 0.33
    const factorEnganche52 = (tasaAnual52 / 2) / 100; // 0.45
    const factorTotal26 = 1 + (tasaAnual26 / 100);    // 1.66
    const factorTotal52 = 1 + (tasaAnual52 / 100);    // 1.90

    const planElegido = e.parameter.plan || "26";

    const ultFila = sheetInv.getLastRow();
    let listaEquipos = [];
    
    if (ultFila > 1) {
      // Leemos las columnas de tu inventario (7 columnas): A (IMEI), B (Modelo), C (Color), D (Memoria), E (Estado), F (Notas), G (Precio Final)
      const datos = sheetInv.getRange(2, 1, ultFila - 1, 7).getValues(); 
      
      datos.forEach(fila => {
        const imei = fila[0] ? fila[0].toString().trim() : "";
        const modelo = fila[1] ? fila[1].toString().trim() : "";
        const color = fila[2] ? fila[2].toString().trim() : "";
        const memoria = fila[3] ? fila[3].toString().trim() : "";
        const estado = fila[4] ? fila[4].toString().toUpperCase().trim() : "";
        const precioContado = parseFloat(fila[6]) || 0;
        
        // 🔒 FILTRO DEL PLAN MAESTRO: Solo mandamos a la web lo que esté físicamente libre,
        // descartamos placeholders temporales que traigan la palabra "ML" y solo marcas permitidas
        if (estado === "DISPONIBLE" && !imei.includes("ML") && esMarcaAutorizada(modelo)) {
          
          // Calculamos de forma exacta y matemática en base a tus tasas de Configuracion
          const engancheCalculado = Math.ceil((precioContado * (planElegido === "26" ? factorEnganche26 : factorEnganche52)) / 10) * 10; 
          const engancheFinal = engancheCalculado;

          const cuota26 = Math.round(((precioContado * factorTotal26) - engancheFinal) / 26);
          const cuota52 = Math.round(((precioContado * factorTotal52) - engancheFinal) / 52);

          // Concatenamos color y memoria para que aparezcan completos en el contrato
          const modeloCompleto = color && memoria ? `${modelo} (${color} - ${memoria})` : modelo;

          listaEquipos.push({
            modelo: modeloCompleto,
            precio: precioContado,
            precioContado: precioContado,
            enganche: engancheFinal,
            cuota26: cuota26,
            cuota52: cuota52,
            imei: imei
          });
        }
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({ result: "success", equipos: listaEquipos }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // --- ACCIÓN: BUSCADOR DE PRODUCTOS INTELIGENTE CON INTERESES INTEGRADOS ---
  if (e.parameter.action === 'buscarModelosCatalogo') {
    try {
      const palabraBuscar = e.parameter.q || "";
      if (palabraBuscar === "") {
        try {
          const cachedCatalog = CacheService.getScriptCache().get("celya_public_catalog");
          if (cachedCatalog) {
            return ContentService.createTextOutput(cachedCatalog).setMimeType(ContentService.MimeType.JSON);
          }
        } catch(eCacheGet) {}
      }
      const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
      const sheetInv = ss.getSheetByName('Catalogo') || ss.getSheetByName('Inventario');
      if (!sheetConfig || !sheetInv) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Falta pestaña Inventario o Configuración" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      let resultadosFormateados = [];
      
      // 📊 PARSE DE TASAS Y GANANCIAS GLOBALES (Lectura agrupada súper rápida)
      const configRow = sheetConfig.getRange("A2:N2").getValues()[0] || [];
      const tasaAnual26 = parseFloat(configRow[3]) || 66; // Columna D (índice 3)
      const tasaAnual52 = parseFloat(configRow[4]) || 90; // Columna E (índice 4)
      const gananciaContadoVal = parseFloat(configRow[5]) || 36; // Columna F (índice 5)
      const engancheSemanalVal = parseFloat(configRow[6]) || 16.5; // Columna G (índice 6)
      const engancheComodoVal = parseFloat(configRow[8]) || 11; // Columna I (índice 8)
      
      const factorGananciaContado = 1 + (gananciaContadoVal / 100); 
      const factorEnganche26 = (tasaAnual26 / 2) / 100; 
      const factorEngancheSemanal = engancheSemanalVal / 100;
      const factorEngancheComodo = engancheComodoVal / 100;
      const factorTotal26 = 1 + (tasaAnual26 / 100);
      const factorTotal52 = 1 + (tasaAnual52 / 100);
      const isCatalogo = (sheetInv.getName() === 'Catalogo');
      const minColsNeeded = isCatalogo ? 13 : 7;

      // 🔄 ESCENARIO A: CARGA INICIAL DEL CATÁLOGO (Query Vacío)
      if (palabraBuscar === "") {
        const ultFila = sheetInv.getLastRow();
        if (ultFila > 1) {
          const maxCols = sheetInv.getLastColumn();
          if (maxCols < minColsNeeded) {
            sheetInv.insertColumnsAfter(maxCols, minColsNeeded - maxCols);
          }
          const totalCols = sheetInv.getLastColumn();
          const datosInv = sheetInv.getRange(2, 1, ultFila - 1, totalCols).getValues();
          
          let linksRicos = [];
          let formulasLinks = [];
          if (isCatalogo) {
            linksRicos = sheetInv.getRange(2, 7, ultFila - 1, 5).getRichTextValues();
            formulasLinks = sheetInv.getRange(2, 7, ultFila - 1, 5).getFormulas();
          }
          
          let peticionesRealizadas = 0;
          const MAX_PETICIONES = 3; 
          
          datosInv.forEach((fila, index) => {
            const numeroFilaReal = index + 2; 
            const imei = fila[0].toString().toUpperCase().trim(); 
            let modeloTitulo = fila[1] ? fila[1].toString().trim() : ""; 

            // 🚫 FILTRAR POR MARCAS FINANCIADAS CELYA
            if (modeloTitulo !== "" && !esMarcaAutorizada(modeloTitulo)) return;

            let precioContadoManual = 0;
            let precioContadoFinalFromSheet = 0;
            let estado = "";
            let linkMercadoLibre = "";
            let cachedFoto = "";
            let cachedMemoria = "";
            let cachedColor = "";
            let cachedVariantes = "";
            let coloresDisponibles = "";
            let esImportacion = false;
            let esUsado = false;

            if (isCatalogo) {
              precioContadoManual = parseFloat(fila[3]) || 0;
              precioContadoFinalFromSheet = (fila.length > 11 && fila[11] !== undefined && fila[11] !== null && String(fila[11]).trim() !== "") ? parseFloat(fila[11]) || 0 : 0;
              estado = fila[5] ? fila[5].toString().toUpperCase().trim() : "";
              const notasFila = fila.length > 4 && fila[4] ? fila[4].toString().toUpperCase() : "";
              esImportacion = notasFila.includes("IMPORTACION");
              esUsado = notasFila.includes("USADO");
              
              const formula = formulasLinks[index] && formulasLinks[index][0] ? formulasLinks[index][0] : "";
              let formulaLink = "";
              if (formula.startsWith("=HYPERLINK")) {
                const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
                if (match) formulaLink = match[1];
              }
              const richLink = linksRicos[index] && linksRicos[index][0] ? linksRicos[index][0].getLinkUrl() : "";
              linkMercadoLibre = richLink || formulaLink || (fila[6] ? fila[6].toString().trim() : "");
              
              const formulaFoto = formulasLinks[index] && formulasLinks[index][1] ? formulasLinks[index][1] : "";
              let formulaFotoLink = "";
              if (formulaFoto.startsWith("=HYPERLINK")) {
                const match = formulaFoto.match(/HYPERLINK\("([^"]+)"/i) || formulaFoto.match(/HYPERLINK\('([^']+)'/i);
                if (match) formulaFotoLink = match[1];
              }
              const richFotoLink = linksRicos[index] && linksRicos[index][1] ? linksRicos[index][1].getLinkUrl() : "";
              cachedFoto = richFotoLink || formulaFotoLink || (fila[7] ? fila[7].toString().trim() : "");
              if (cachedFoto.toLowerCase().trim() === "foto") cachedFoto = "";
              
              cachedMemoria = fila[8] ? fila[8].toString().trim() : "";
              cachedColor = fila[9] ? fila[9].toString().trim() : "";
              const richVariantesLink = linksRicos[index] && linksRicos[index][4] ? linksRicos[index][4].getLinkUrl() : "";
              cachedVariantes = richVariantesLink || (fila[10] ? fila[10].toString().trim() : "");
              coloresDisponibles = (fila.length > 12 && fila[12] !== undefined && fila[12] !== null) ? fila[12].toString().trim() : "";
            } else {
              precioContadoManual = parseFloat(fila[6]) || 0;
              estado = fila[4] ? fila[4].toString().toUpperCase().trim() : "";
              cachedMemoria = fila[3] ? fila[3].toString().trim() : "";
              cachedColor = fila[2] ? fila[2].toString().trim() : "";
            }
            
            if (estado === "DISPONIBLE" || estado === "MAS VENDIDO" || estado === "BAJO DEMANDA" || imei.startsWith("999999")) {
              let itemId = "";
              if (linkMercadoLibre.includes("mercadolibre")) {
                const regex = /(MLM\-?\d+)/i;
                const coincidencia = linkMercadoLibre.match(regex);
                if (coincidencia) itemId = coincidencia[0].replace("-", "").toUpperCase();
              }
              
              let fotoProducto = cachedFoto || "https://images.unsplash.com/photo-1511707171634-5f897ff02aa9?w=150&auto=format&fit=crop"; 
              
              if ((precioContadoManual === 0 || cachedFoto === "") && itemId !== "") {
                if (peticionesRealizadas < MAX_PETICIONES) {
                  try {
                    const apiUrl = `https://api.mercadolibre.com/items/${itemId}`;
                    const respuestaApi = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
                    if (respuestaApi.getResponseCode() === 200) {
                      const itemData = JSON.parse(respuestaApi.getContentText());
                      const precioCostoML = parseFloat(itemData.price) || 0;
                      
                      if (precioCostoML > 0 && precioContadoManual === 0) {
                        precioContadoManual = Math.round(precioCostoML);
                        if (isCatalogo) {
                          sheetInv.getRange(numeroFilaReal, 4).setValue(precioContadoManual);
                        }
                      }
                      if (modeloTitulo === "") {
                        modeloTitulo = itemData.title;
                        if (!esMarcaAutorizada(modeloTitulo)) return;
                        sheetInv.getRange(numeroFilaReal, 2).setValue(modeloTitulo);
                      }
                      if (itemData.pictures && itemData.pictures.length > 0 && cachedFoto === "") {
                        fotoProducto = itemData.pictures[0].secure_url;
                        if (isCatalogo) {
                          sheetInv.getRange(numeroFilaReal, 8).setFormula(`=HYPERLINK("${fotoProducto}", "foto")`);
                        }
                      }
                      peticionesRealizadas++;
                    }
                  } catch (err) {
                    Logger.log("Error en API ML: " + err.toString());
                  }
                }
              }
              
              let precioContado = 0;
              if (precioContadoFinalFromSheet > 0) {
                precioContado = Math.round(precioContadoFinalFromSheet);
              } else {
                precioContado = Math.round(precioContadoManual * factorGananciaContado);
              }
              
              if (precioContado > 0) {
                const engMin = parseFloat(fila[2]) || (precioContado * factorEngancheSemanal);
                const cuota26 = ((precioContado * (1 + (tasaAnual26/100))) - engMin) / 26;
                const cuota52 = ((precioContado * (1 + (tasaAnual52/100))) - engMin) / 52;
                
                // Parsear variantesJson para evitar doble-escape en JSON.stringify
                let variantesObj = null;
                if (cachedVariantes && cachedVariantes.trim() !== "") {
                  try {
                    variantesObj = JSON.parse(cachedVariantes);
                  } catch (e) {
                    Logger.log("Error parseando variantesJson: " + e.toString());
                    variantesObj = null;
                  }
                }
                
                resultadosFormateados.push({
                  id: itemId || imei,
                  modelo: modeloTitulo || "Celular CelYa",
                  foto: fotoProducto, 
                  memoria: cachedMemoria,
                  color: cachedColor,
                  engancheMinimo: Math.ceil(engMin / 10) * 10,
                  cuota26: Math.round(cuota26),
                  cuota52: Math.round(cuota52),
                  esFijo: true, // Booleano nativo para el render
                  entregaInmediata: (estado === "DISPONIBLE" || estado === "MAS VENDIDO"),
                  estado: estado,
                  precioContado: Math.round(precioContado),
                  link: linkMercadoLibre,
                  variantesJson: variantesObj || {},
                  coloresDisponibles: coloresDisponibles,
                  esImportacion: esImportacion,
                  esUsado: esUsado,
                  filaNum: numeroFilaReal
                });
              }
            }
          });
        }
        return ContentService.createTextOutput(JSON.stringify({ result: "success", resultados: resultadosFormateados })).setMimeType(ContentService.MimeType.JSON);
      
      } else {
        // 🔄 ESCENARIO B: EL CLIENTE ESTÁ USANDO EL BUSCADOR ACTIVAMENTE
        const ultFila = sheetInv.getLastRow();
        if (ultFila > 1) {
          const maxCols = sheetInv.getLastColumn();
          if (maxCols < minColsNeeded) {
            sheetInv.insertColumnsAfter(maxCols, minColsNeeded - maxCols);
          }
          const totalCols = sheetInv.getLastColumn();
          const datosInv = sheetInv.getRange(2, 1, ultFila - 1, totalCols).getValues();
          
          let linksRicos = [];
          let formulasLinks = [];
          if (isCatalogo) {
            linksRicos = sheetInv.getRange(2, 7, ultFila - 1, 5).getRichTextValues();
            formulasLinks = sheetInv.getRange(2, 7, ultFila - 1, 5).getFormulas();
          }
          const queryNormalizada = palabraBuscar.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

          let peticionesRealizadas = 0;
          const MAX_PETICIONES = 3;

          datosInv.forEach((fila, index) => {
            const numeroFilaReal = index + 2; 
            let imei = typeof fila[0] === "number" ? fila[0].toFixed(0) : (fila[0] || "").toString().trim();
            imei = imei.toUpperCase();

            let modeloTitulo = fila[1] ? fila[1].toString().trim() : ""; 

            if (modeloTitulo !== "" && !esMarcaAutorizada(modeloTitulo)) return;

            let precioContadoManual = 0;
            let precioContadoFinalFromSheet = 0;
            let estado = "";
            let linkMercadoLibre = "";
            let cachedFoto = "";
            let cachedMemoria = "";
            let cachedColor = "";
            let cachedVariantes = "";
            let coloresDisponibles = "";
            let esImportacion = false;

            if (isCatalogo) {
              precioContadoManual = parseFloat(fila[3]) || 0;
              precioContadoFinalFromSheet = (fila.length > 11 && fila[11] !== undefined && fila[11] !== null && String(fila[11]).trim() !== "") ? parseFloat(fila[11]) || 0 : 0;
              estado = fila[5] ? fila[5].toString().toUpperCase().trim() : "";
              esImportacion = (fila.length > 4 && fila[4] === "IMPORTACION");
              
              const formula = formulasLinks[index] && formulasLinks[index][0] ? formulasLinks[index][0] : "";
              let formulaLink = "";
              if (formula.startsWith("=HYPERLINK")) {
                const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
                if (match) formulaLink = match[1];
              }
              const richLink = linksRicos[index] && linksRicos[index][0] ? linksRicos[index][0].getLinkUrl() : "";
              linkMercadoLibre = richLink || formulaLink || (fila[6] ? fila[6].toString().trim() : "");
              
              const formulaFoto = formulasLinks[index] && formulasLinks[index][1] ? formulasLinks[index][1] : "";
              let formulaFotoLink = "";
              if (formulaFoto.startsWith("=HYPERLINK")) {
                const match = formulaFoto.match(/HYPERLINK\("([^"]+)"/i) || formulaFoto.match(/HYPERLINK\('([^']+)'/i);
                if (match) formulaFotoLink = match[1];
              }
              const richFotoLink = linksRicos[index] && linksRicos[index][1] ? linksRicos[index][1].getLinkUrl() : "";
              cachedFoto = richFotoLink || formulaFotoLink || (fila[7] ? fila[7].toString().trim() : "");
              if (cachedFoto.toLowerCase().trim() === "foto") cachedFoto = "";
              
              cachedMemoria = fila[8] ? fila[8].toString().trim() : "";
              cachedColor = fila[9] ? fila[9].toString().trim() : "";
              const richVariantesLink = linksRicos[index] && linksRicos[index][4] ? linksRicos[index][4].getLinkUrl() : "";
              cachedVariantes = richVariantesLink || (fila[10] ? fila[10].toString().trim() : "");
              coloresDisponibles = (fila.length > 12 && fila[12] !== undefined && fila[12] !== null) ? fila[12].toString().trim() : "";
            } else {
              precioContadoManual = parseFloat(fila[6]) || 0;
              estado = fila[4] ? fila[4].toString().toUpperCase().trim() : "";
              cachedMemoria = fila[3] ? fila[3].toString().trim() : "";
              cachedColor = fila[2] ? fila[2].toString().trim() : "";
            }

            let coincideQuery = true;
            if (queryNormalizada !== "") {
              const modeloNormalizado = modeloTitulo.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
              coincideQuery = modeloNormalizado.includes(queryNormalizada) || imei.includes(queryNormalizada);
            }

            if (coincideQuery && (estado === "DISPONIBLE" || estado === "MAS VENDIDO" || estado === "BAJO DEMANDA" || imei.startsWith("999999"))) {
              let itemId = "";
              if (linkMercadoLibre.includes("mercadolibre")) {
                const regex = /(MLM\-?\d+)/i;
                const coincidencia = linkMercadoLibre.match(regex);
                if (coincidencia) itemId = coincidencia[0].replace("-", "").toUpperCase();
              }

              let fotoProducto = cachedFoto || "https://images.unsplash.com/photo-1511707171634-5f897ff02aa9?w=150&auto=format&fit=crop"; 
              const esEntregaInmediata = ((estado === "DISPONIBLE" || estado === "MAS VENDIDO") && !imei.startsWith("999999"));
              
              let precioContado = 0;
              if (precioContadoFinalFromSheet > 0) {
                precioContado = Math.round(precioContadoFinalFromSheet);
              } else {
                precioContado = Math.round(precioContadoManual * factorGananciaContado);
              }

              if (precioContado > 0) {
                const engMin = (precioContado * factorEngancheSemanal);
                const engComodoMin = (precioContado * factorEngancheComodo);
                const cuota26 = ((precioContado * factorTotal26) - engMin) / 26;
                const cuota52 = ((precioContado * factorTotal52) - engComodoMin) / 52;
                
                // Parsear variantesJson para evitar doble-escape en JSON.stringify
                let variantesObj = null;
                if (cachedVariantes && cachedVariantes.trim() !== "") {
                  try {
                    variantesObj = JSON.parse(cachedVariantes);
                  } catch (e) {
                    Logger.log("Error parseando variantesJson en búsqueda: " + e.toString());
                    variantesObj = null;
                  }
                }
                
                resultadosFormateados.push({
                  id: itemId || imei,
                  modelo: modeloTitulo || "Celular CelYa",
                  foto: fotoProducto,
                  memoria: cachedMemoria,
                  color: cachedColor,
                  engancheMinimo: Math.ceil(Math.min(engMin, engComodoMin) / 10) * 10,
                  cuota26: Math.round(cuota26),
                  cuota52: Math.round(cuota52),
                  precioContado: Math.round(precioContado),
                  esFijo: true, // Booleano nativo
                  entregaInmediata: esEntregaInmediata,
                  estado: estado,
                  link: linkMercadoLibre,
                  variantesJson: variantesObj || {},
                  coloresDisponibles: coloresDisponibles,
                  esImportacion: esImportacion,
                  filaNum: numeroFilaReal
                });
              }
            }
          });
        }

        // 🚀 FALLBACK: SI NO HUBO RESULTADOS LOCALES, BUSCAR EN MERCADOLIBRE GLOBAL VIVO
        if (palabraBuscar.trim() !== "" && resultadosFormateados.length === 0) {
          try {
            const palabraClean = palabraBuscar.toLowerCase().trim();
            const urlBuscarML = "https://api.mercadolibre.com/sites/MLM/search?q=" + encodeURIComponent(palabraClean) + "&limit=5";
            const resML = UrlFetchApp.fetch(urlBuscarML, { muteHttpExceptions: true });
            
            if (resML.getResponseCode() === 200) {
              const dataML = JSON.parse(resML.getContentText());
              if (dataML.results && dataML.results.length > 0) {
                dataML.results.forEach(item => {
                  if (!esMarcaAutorizada(item.title)) return;

                  const precioCostoML = parseFloat(item.price) || 0;
                  if (precioCostoML > 0) {
                    const precioVentaSugerido = Math.round(precioCostoML * factorGananciaContado);
                    const engMin = Math.ceil((precioVentaSugerido * factorEngancheSemanal) / 10) * 10;
                    const engComodoMin = Math.ceil((precioVentaSugerido * factorEngancheComodo) / 10) * 10;
                    const cuota26 = Math.round(((precioVentaSugerido * factorTotal26) - engMin) / 26);
                    const cuota52 = Math.round(((precioVentaSugerido * factorTotal52) - engComodoMin) / 52);
                    
                    let isCbtLive = false;
                    if (item.shipping && (item.shipping.logistic_type === "remote" || (item.shipping.tags && (item.shipping.tags.includes("cbt_fulfillment") || item.shipping.tags.includes("cbt"))))) {
                      isCbtLive = true;
                    }

                    resultadosFormateados.push({
                      id: item.id,
                      modelo: item.title,
                      foto: item.thumbnail ? item.thumbnail.replace("-I.jpg", "-O.jpg") : "https://images.unsplash.com/photo-1511707171634-5f897ff02aa9?w=150&auto=format&fit=crop",
                      memoria: "",
                      color: "",
                      engancheMinimo: Math.min(engMin, engComodoMin),
                      cuota26: cuota26,
                      cuota52: cuota52,
                      precioContado: precioVentaSugerido,
                      esFijo: false, // Indica de forma limpia que es externo
                      entregaInmediata: false,
                      estado: "BAJO DEMANDA",
                      link: item.permalink || ("https://articulo.mercadolibre.com.mx/MLM-" + item.id.replace("MLM","")),
                      esImportacion: isCbtLive
                    });
                  }
                });
              }
            }
          } catch (err) {
            escribirLogDebug("Error buscando en la API global de ML: " + err.toString());
          }
        }
        const jsonResponse = JSON.stringify({ result: "success", resultados: resultadosFormateados });
        if (palabraBuscar === "") {
          try {
            CacheService.getScriptCache().put("celya_public_catalog", jsonResponse, 3600);
          } catch(eCachePut) {}
        }
        return ContentService.createTextOutput(jsonResponse).setMimeType(ContentService.MimeType.JSON);
      }
                           
    } catch (error) {
      Logger.log("Error en buscarModelosCatalogo: " + error.toString());
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Ocurrió un error procesando el catálogo." })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- ACCIÓN: OBTENER TODOS LOS DATOS PARA EL DASHBOARD ADMINISTRATIVO ---
  if (e.parameter.action === 'obtenerDatosDashboard') {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    const sheetClientes = ss.getSheetByName('Clientes');
    const sheetInventario = ss.getSheetByName('Inventario');
    const sheetPagos = ss.getSheetByName('Pagos');
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');

    // Cargar parámetros de configuración
    let configActual = { interes: "6", horastol: "24", semanastol: "3", interesPlanComodo: "75", interesPlanRapido: "50", gananciaContado: "15", engancheSemanal: "20", engancheComodo: "30", descuentoLiquidacion: "15", telegramBotToken: "", telegramChatId: "", tasaInteresQuincenalRapido: "45", tasaInteresQuincenalComodo: "65", engancheMinimoQuincenalRapido: "25", engancheMinimoQuincenalComodo: "35" };
    if (sheetConfig) {
      const filaConfig = sheetConfig.getRange("A2:P2").getValues()[0];
      configActual = {
        interes:           filaConfig[0] !== undefined ? String(filaConfig[0]) : "6",
        horastol:          filaConfig[1] !== undefined ? String(filaConfig[1]) : "24",
        semanastol:        filaConfig[2] !== undefined ? String(filaConfig[2]) : "3",
        interesPlanRapido: filaConfig[3] !== undefined && filaConfig[3] !== "" ? String(filaConfig[3]) : "50",
        interesPlanComodo: filaConfig[4] !== undefined && filaConfig[4] !== "" ? String(filaConfig[4]) : "75",
        gananciaContado:   filaConfig[5] !== undefined && filaConfig[5] !== "" ? String(filaConfig[5]) : "15",
        engancheSemanal:   filaConfig[6] !== undefined && filaConfig[6] !== "" ? String(filaConfig[6]) : "20",

        engancheComodo:    filaConfig[8] !== undefined && filaConfig[8] !== "" ? String(filaConfig[8]) : "30",
        descuentoLiquidacion: filaConfig[9] !== undefined && filaConfig[9] !== "" ? String(filaConfig[9]) : "15",
        telegramBotToken:  filaConfig[10] !== undefined ? String(filaConfig[10]) : "",
        telegramChatId:    filaConfig[11] !== undefined ? String(filaConfig[11]) : "",
        tasaInteresQuincenalRapido: filaConfig[12] !== undefined && filaConfig[12] !== "" ? String(filaConfig[12]) : "45",
        tasaInteresQuincenalComodo: filaConfig[13] !== undefined && filaConfig[13] !== "" ? String(filaConfig[13]) : "65",
        engancheMinimoQuincenalRapido: filaConfig[14] !== undefined && filaConfig[14] !== "" ? String(filaConfig[14]) : "25",
        engancheMinimoQuincenalComodo: filaConfig[15] !== undefined && filaConfig[15] !== "" ? String(filaConfig[15]) : "35"
      };
    }

    // Cargar inventario vinculando precios desde Catalogo
    let inventario = [];
    if (sheetInventario) {
      const dataInv = sheetInventario.getDataRange().getValues();
      const preciosCatalogo = obtenerPreciosDeCatalogo(ss);
      const gananciaContadoVal = parseFloat(configActual.gananciaContado) || 15;
      const factorGananciaContado = 1 + (gananciaContadoVal / 100);
      const tasaAnual26 = parseFloat(configActual.interesPlanRapido) || 66;
      const factorEnganche26 = (tasaAnual26 / 2) / 100;

      for (let i = 1; i < dataInv.length; i++) {
        if (dataInv[i][0]) {
          const imei = String(dataInv[i][0] || '').trim();
          const modelo = String(dataInv[i][1] || '').trim();
          const color = String(dataInv[i][2] || '').trim();
          const memoria = String(dataInv[i][3] || '').trim();
          const estado = String(dataInv[i][4] || '').trim();
          const notas = String(dataInv[i][5] || '').trim();
          
          const precioFinalInv = dataInv[i][6] ? parseFloat(dataInv[i][6]) || 0 : 0;
          const catInfo = preciosCatalogo[modelo.toUpperCase()];
          const precioCosto = catInfo ? catInfo.costo : 0;
          const precioContadoFinal = catInfo ? catInfo.precioContadoFinal : 0;
          let precioContadoCalculado = 0;
          if (precioFinalInv > 0) {
            precioContadoCalculado = precioFinalInv;
          } else if (precioContadoFinal > 0) {
            precioContadoCalculado = precioContadoFinal;
          } else {
            precioContadoCalculado = Math.round(precioCosto * factorGananciaContado);
          }
          const engancheCalculado = Math.ceil((precioContadoCalculado * factorEnganche26) / 10) * 10;
          
          const factorTotal26 = 1 + (tasaAnual26 / 100);
          const totalFinanciadoCalculado = Math.round((precioContadoCalculado * factorTotal26) - engancheCalculado);

          inventario.push({
            modelo:          modelo,
            imei:            imei,
            color:           color,
            memoria:         memoria,
            estado:          estado,
            notas:           notas,
            enganche:        engancheCalculado,
            precioContado:   precioContadoCalculado,
            planRapido:      "",
            totalFinanciado: totalFinanciadoCalculado
          });
        }
      }
    }

    // Agrupar pagos por IMEI
    const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    const pagosPorImei = {};
    for (let j = 1; j < dataPagos.length; j++) {
      const imeiPago = String(dataPagos[j][0]).trim();
      const fechaPago = dataPagos[j][2];
      const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
      const metodoPago = String(dataPagos[j][4] || "");
      
      if (imeiPago) {
        if (!pagosPorImei[imeiPago]) {
          pagosPorImei[imeiPago] = [];
        }
        let fechaFormateada = "";
        if (fechaPago instanceof Date) {
          fechaFormateada = Utilities.formatDate(fechaPago, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (fechaPago) {
          fechaFormateada = String(fechaPago);
        }
        pagosPorImei[imeiPago].push({
          fecha: fechaFormateada,
          monto: montoPago,
          chofer: metodoPago
        });
      }
    }

    // Cargar clientes activos (con corrección para leer hasta la columna T/index 19 y U/index 20)
    const ultFilaClientes = sheetClientes.getLastRow();
    let clientesActivos = [];
    if (ultFilaClientes > 1) {
      // 🔄 CORRECCIÓN: Leemos hasta la columna 21 (Columna U) para traer Saldo, Días de Atraso y ID
      const datosClientes = sheetClientes.getRange(2, 1, ultFilaClientes - 1, 21).getValues();
      
      clientesActivos = datosClientes.map((fila, index) => {
        const imeiKey = fila[5] ? fila[5].toString().trim() : "";
        const clienteNombre = fila[1] ? fila[1].toString().trim() : "";
        const planSemanas = fila[14] ? fila[14].toString().trim() : "26";
        const cuotaNum = parseFloat(String(fila[10] || "0").replace(/[\$,]/g, "")) || 0;
        const totalFinanciadoVal = parseFloat(String(fila[15] || "0").replace(/[\$,]/g, "")) || 0;
        
        let numeroPlan = "52";
        if (planSemanas.includes("26") || planSemanas.includes("13")) {
          numeroPlan = "26";
        }

        // Sumar todos los abonos de este IMEI
        let totalPagado = 0;
        const abonos = pagosPorImei[imeiKey] || [];
        abonos.forEach(function(a) { totalPagado += a.monto; });

        const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);

        // Calcular Días de Atraso
        const startTimestamp = fila[0];
        const tipoPeriodo = String(fila[20] || "SEMANAL").toUpperCase();
        let diasAtrasoVal = 0;
        if (startTimestamp instanceof Date && cuotaNum > 0 && saldoPendienteBase > 0) {
          const diasTranscurridos = Math.floor((new Date() - startTimestamp) / (1000 * 60 * 60 * 24));
          const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
          const periodosTranscurridos = Math.floor(diasTranscurridos / diasPorPeriodo);
          const totalEsperado = periodosTranscurridos * cuotaNum;
          if (totalPagado < totalEsperado) {
            const atrasoMonto = totalEsperado - totalPagado;
            diasAtrasoVal = Math.ceil((atrasoMonto / cuotaNum) * diasPorPeriodo);
          }
        }

        // Calcular penalidad por cada semana de atraso
        let semanasAtrasoVal = Math.floor(diasAtrasoVal / 7);
        let penalidadMonto = 0;
        if (semanasAtrasoVal > 0) {
          const tasaInteres = parseFloat(configActual.interes) || 6;
          penalidadMonto = semanasAtrasoVal * (cuotaNum * (tasaInteres / 100));
        }

        const saldoPendienteVal = saldoPendienteBase + penalidadMonto;

        return {
          id:              fila[19] ? fila[19].toString().trim() : "C-Y" + (index + 1), // Columna T (ID)
          imei:            imeiKey,                                                     // Columna F (IMEI)
          cliente:         clienteNombre,                                               // Columna B (Nombre)
          modelo:          fila[4] ? fila[4].toString().trim() : "Genérico",            // Columna E (Modelo)
          plan:            numeroPlan,                                                  // Columna O (Plazo)
          cuota:           cuotaNum,
          etiquetaCompleta: (clienteNombre || "S/N") + " (" + (imeiKey || "S/IMEI") + ")",
          saldoPendiente:  saldoPendienteVal,
          diasAtraso:      diasAtrasoVal,
          tipoPeriodo:     tipoPeriodo,
          historialAbonos: abonos,
          diaRaya:         fila[18] ? fila[18].toString().trim() : "",
          interesCliente:  fila[11] !== undefined && fila[11] !== "" ? parseFloat(fila[11]) : null,
          horasToleranciaCliente: fila[12] !== undefined && fila[12] !== "" ? parseInt(fila[12]) : null,
          semanasToleranciaCliente: fila[13] !== undefined && fila[13] !== "" ? parseInt(fila[13]) : null,
          fechaInicio:     startTimestamp instanceof Date ? Utilities.formatDate(startTimestamp, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX") : String(startTimestamp || ""),
          totalFinanciado: totalFinanciadoVal,
          totalPagado:     totalPagado
        };
      }).filter(c => c.imei !== ""); // Filtramos filas vacías
    }

    // Cargar promesas de pago activas (PENDIENTES)
    const sheetPromesas = ss.getSheetByName('Promesas');
    let promesasActivas = [];
    if (sheetPromesas) {
      const dataProm = sheetPromesas.getDataRange().getValues();
      for (let k = 1; k < dataProm.length; k++) {
        const est = String(dataProm[k][4] || '').trim().toUpperCase();
        if (est === "PENDIENTE") {
          let fProm = dataProm[k][2];
          let fPromStr = "";
          if (fProm instanceof Date) {
            fPromStr = Utilities.formatDate(fProm, Session.getScriptTimeZone(), "dd/MM/yyyy");
          } else if (fProm) {
            fPromStr = String(fProm);
          }
          promesasActivas.push({
            imei:           String(dataProm[k][0] || '').trim(),
            cliente:        String(dataProm[k][1] || '').trim(),
            fechaPromesa:   fPromStr,
            montoPrometido: parseFloat(dataProm[k][3]) || 0,
            notas:          String(dataProm[k][6] || '').trim()
          });
        }
      }
    }

    const metodosPagoVal = PropertiesService.getScriptProperties().getProperty('METODOS_PAGO') || "";
    return ContentService.createTextOutput(JSON.stringify({ 
      inventario: inventario, 
      config: configActual,
      clientesActivos: clientesActivos,
      promesasActivas: promesasActivas,
      metodosPago: metodosPagoVal
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    if (!validarTokenApi(credenciales, e.parameter.token)) {
      return ContentService.createTextOutput(JSON.stringify({ error: "Acceso denegado: Token de seguridad requerido o inválido" })).setMimeType(ContentService.MimeType.JSON);
    }
    // 1. OBTENER INVENTARIO CON MAPEO CORRECTO 📱
    const sheetInv = ss.getSheetByName('Inventario');
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    const sheetClientes = ss.getSheetByName('Clientes');
    const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];

    const imeisVendidos = {};
    for (let j = 1; j < dataClientes.length; j++) {
      const imeiCliente = String(dataClientes[j][5] || '').trim();
      if (imeiCliente) {
        imeisVendidos[imeiCliente] = true;
      }
    }

    let configActual = { 
      interes: "6", 
      horastol: "24", 
      semanastol: "3", 
      interesPlanComodo: "75", 
      interesPlanRapido: "50", 
      gananciaContado: "15",
      engancheSemanal: "20",

      engancheComodo: "30",
      descuentoLiquidacion: "15",
      telegramBotToken: "",
      telegramChatId: "",
      tasaInteresQuincenalRapido: "45",
      tasaInteresQuincenalComodo: "65",
      engancheMinimoQuincenalRapido: "25",
      engancheMinimoQuincenalComodo: "35"
    };
    if (sheetConfig) {
      const filaConfig = sheetConfig.getRange("A2:P2").getValues()[0];
      configActual = {
        interes:           filaConfig[0] !== undefined ? String(filaConfig[0]) : "6",
        horastol:          filaConfig[1] !== undefined ? String(filaConfig[1]) : "24",
        semanastol:        filaConfig[2] !== undefined ? String(filaConfig[2]) : "3",
        interesPlanRapido: filaConfig[3] !== undefined && filaConfig[3] !== "" ? String(filaConfig[3]) : "50",
        interesPlanComodo: filaConfig[4] !== undefined && filaConfig[4] !== "" ? String(filaConfig[4]) : "75",
        gananciaContado:   filaConfig[5] !== undefined && filaConfig[5] !== "" ? String(filaConfig[5]) : "15",
        engancheSemanal:   filaConfig[6] !== undefined && filaConfig[6] !== "" ? String(filaConfig[6]) : "20",

        engancheComodo:    filaConfig[8] !== undefined && filaConfig[8] !== "" ? String(filaConfig[8]) : "30",
        descuentoLiquidacion: filaConfig[9] !== undefined && filaConfig[9] !== "" ? String(filaConfig[9]) : "15",
        telegramBotToken:  filaConfig[10] !== undefined ? String(filaConfig[10]) : "",
        telegramChatId:    filaConfig[11] !== undefined ? String(filaConfig[11]) : "",
        tasaInteresQuincenalRapido: filaConfig[12] !== undefined && filaConfig[12] !== "" ? String(filaConfig[12]) : "45",
        tasaInteresQuincenalComodo: filaConfig[13] !== undefined && filaConfig[13] !== "" ? String(filaConfig[13]) : "65",
        engancheMinimoQuincenalRapido: filaConfig[14] !== undefined && filaConfig[14] !== "" ? String(filaConfig[14]) : "25",
        engancheMinimoQuincenalComodo: filaConfig[15] !== undefined && filaConfig[15] !== "" ? String(filaConfig[15]) : "35"
      };
    }

    const dataInv = sheetInv ? sheetInv.getDataRange().getValues() : [];
    const inventario = [];

    let colImei = 0;
    let colModelo = 1;
    let colColor = 2;
    let colMemoria = 3;
    let colEstado = 4;
    let colNotas = 5;
    let colPrecioFinal = 6;

    if (dataInv.length > 0) {
      const headers = dataInv[0].map(function(h) { return h.toString().toLowerCase().trim(); });
      const findIndex = function(keywords) {
        return headers.findIndex(function(h) {
          return keywords.some(function(k) { return h.indexOf(k) !== -1; });
        });
      };

      const idxImei = findIndex(["imei"]);
      if (idxImei !== -1) colImei = idxImei;
      const idxModelo = findIndex(["modelo", "equipo"]);
      if (idxModelo !== -1) colModelo = idxModelo;
      const idxColor = findIndex(["color"]);
      if (idxColor !== -1) colColor = idxColor;
      const idxMemoria = findIndex(["memoria", "capacidad"]);
      if (idxMemoria !== -1) colMemoria = idxMemoria;
      const idxEstado = findIndex(["estado", "status"]);
      if (idxEstado !== -1) colEstado = idxEstado;
      const idxNotas = findIndex(["nota", "comentario"]);
      if (idxNotas !== -1) colNotas = idxNotas;
      const idxPrecioFinal = findIndex(["precio", "final", "precio final"]);
      if (idxPrecioFinal !== -1) colPrecioFinal = idxPrecioFinal;
    }

    const preciosCatalogo = obtenerPreciosDeCatalogo(ss);

    for (let i = 1; i < dataInv.length; i++) {
      const estado = dataInv[i][colEstado] ? dataInv[i][colEstado].toString().trim().toUpperCase() : "";
      const modelo = dataInv[i][colModelo] ? dataInv[i][colModelo].toString().trim() : "";
      const imei = dataInv[i][colImei] ? dataInv[i][colImei].toString().trim() : "";
      const precioFinalInv = dataInv[i][colPrecioFinal] ? parseFloat(dataInv[i][colPrecioFinal]) || 0 : 0;
      
      let esApto = (estado === "DISPONIBLE" || estado === "MAS VENDIDO" || estado === "BAJO DEMANDA" || (estado === "SIENDO PAGADO" && !imeisVendidos[imei] && !imeisVendidos[imei.toUpperCase()]));
      if (esApto && !imei.toUpperCase().includes("ML") && esMarcaAutorizada(modelo)) {
        // Usar precio final del inventario si está disponible, si no calcular desde catálogo
        let precioContadoCalculado = 0;
        if (precioFinalInv > 0) {
          precioContadoCalculado = precioFinalInv;
        } else {
          const catInfo = preciosCatalogo[modelo.toUpperCase()];
          const precioCosto = catInfo ? catInfo.costo : 0;
          const precioContadoFinal = catInfo ? catInfo.precioContadoFinal : 0;
          if (precioContadoFinal > 0) {
            precioContadoCalculado = precioContadoFinal;
          } else {
            const gananciaContadoVal = parseFloat(configActual.gananciaContado) || 15;
            const factorGananciaContado = 1 + (gananciaContadoVal / 100);
            precioContadoCalculado = Math.round(precioCosto * factorGananciaContado);
          }
        }

        const tasaAnual26 = parseFloat(configActual.interesPlanRapido) || 66;
        const engancheSemanalVal = parseFloat(configActual.engancheSemanal) || 16.5;
        const factorEngancheSemanal = engancheSemanalVal / 100;
        const engancheFinal = Math.ceil((precioContadoCalculado * factorEngancheSemanal) / 10) * 10;

        const factorTotal26 = 1 + (tasaAnual26 / 100);
        const totalFinanciadoCalculado = Math.round((precioContadoCalculado * factorTotal26) - engancheFinal);

        inventario.push({
          imei:            imei, 
          modelo:          modelo, 
          enganche:        String(engancheFinal), 
          precioContado:   String(precioContadoCalculado),
          planRapido:      "",           
          planComodo:      "",           
          totalFinanciado: String(totalFinanciadoCalculado),           
          totalLetra:      numeroALetras(totalFinanciadoCalculado) 
        });
      }
    }

    // 2. EXTRAER CLIENTES ACTIVOS Y CALCULAR COBRANZA DINÁMICA
    
    const sheetPagos = ss.getSheetByName('Pagos');
    const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    
    // Agrupar pagos por IMEI
    const pagosPorImei = {};
    for (let j = 1; j < dataPagos.length; j++) {
      const imeiPago = String(dataPagos[j][0]).trim();
      const fechaPago = dataPagos[j][2]; // Date object
      const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
      const metodoPago = String(dataPagos[j][4] || "");
      
      if (imeiPago) {
        if (!pagosPorImei[imeiPago]) {
          pagosPorImei[imeiPago] = [];
        }
        
        let fechaFormateada = "";
        if (fechaPago instanceof Date) {
          fechaFormateada = Utilities.formatDate(fechaPago, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (fechaPago) {
          fechaFormateada = String(fechaPago);
        }
        
        pagosPorImei[imeiPago].push({
          fecha: fechaFormateada,
          monto: montoPago,
          chofer: metodoPago
        });
      }
    }

    const clientesActivos = [];
    for (let i = 1; i < dataClientes.length; i++) {
      const clienteNombre = dataClientes[i][1]; // Columna B
      const modeloCel = dataClientes[i][4];     // Columna E
      const imeiCel = dataClientes[i][5];       // Columna F
      const planSemanas = dataClientes[i][14];  // Columna O
      
      if (clienteNombre && imeiCel) {
        const imeiKey = imeiCel.toString().trim();
        let numeroPlan = "52";
        if (planSemanas && planSemanas.toString().includes("26")) {
          numeroPlan = "26";
        }
        
        // Sumar todos los abonos de este IMEI
        let totalPagado = 0;
        const abonos = pagosPorImei[imeiKey] || [];
        abonos.forEach(function(a) { totalPagado += a.monto; });
        
        const cuotaNum = parseFloat(String(dataClientes[i][10] || "0").replace(/[\$,]/g, "")) || 0;
        
        const totalFinanciadoVal = parseFloat(String(dataClientes[i][15] || "0").replace(/[\$,]/g, "")) || 0;
        const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);
        
        // Calcular Días de Atraso
        const startTimestamp = dataClientes[i][0];
        const tipoPeriodo = String(dataClientes[i][20] || "SEMANAL").toUpperCase();
        let diasAtrasoVal = 0;
        if (startTimestamp instanceof Date && cuotaNum > 0 && saldoPendienteBase > 0) {
          const diasTranscurridos = Math.floor((new Date() - startTimestamp) / (1000 * 60 * 60 * 24));
          const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
          const periodosTranscurridos = Math.floor(diasTranscurridos / diasPorPeriodo);
          const totalEsperado = periodosTranscurridos * cuotaNum;
          if (totalPagado < totalEsperado) {
            const atrasoMonto = totalEsperado - totalPagado;
            diasAtrasoVal = Math.ceil((atrasoMonto / cuotaNum) * diasPorPeriodo);
          }
        }

        // Calcular penalidad por cada semana de atraso
        let semanasAtrasoVal = Math.floor(diasAtrasoVal / 7);
        let penalidadMonto = 0;
        if (semanasAtrasoVal > 0) {
          const tasaInteres = parseFloat(configActual.interes) || 6;
          penalidadMonto = semanasAtrasoVal * (cuotaNum * (tasaInteres / 100));
        }

        const saldoPendienteVal = saldoPendienteBase + penalidadMonto;

        clientesActivos.push({
          id: String(dataClientes[i][19] || ("C-Y" + i)).trim(),
          imei: imeiKey,
          modelo: modeloCel ? modeloCel.toString().trim() : "Genérico",
          cliente: clienteNombre.toString().trim(),
          plan: numeroPlan,
          cuota: cuotaNum,
          saldoPendiente: saldoPendienteVal,
          diasAtraso: diasAtrasoVal,
          historialAbonos: abonos,
          etiquetaCompleta: `${clienteNombre.toString().trim()} - ${modeloCel ? modeloCel.toString().trim() : "Sin Modelo"} (${imeiKey})`
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ 
      inventario: inventario, 
      config: configActual,
      clientesActivos: clientesActivos 
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    console.error("Error en doGet: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({error: "Ocurrió un error procesando la solicitud."})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const credenciales = obtenerPropiedadesEcosistema();
  
  // 🤖 DETECTAR SI LA PETICIÓN ES UN WEBHOOK DE TELEGRAM (JSON raw)
  if (e && e.postData && e.postData.contents) {
    try {
      const update = JSON.parse(e.postData.contents);
      if (update && (update.message || update.callback_query)) {
        procesarWebhookTelegram(update, credenciales);
        return HtmlService.createHtmlOutput("OK");
      }
    } catch (parseErr) {
      // Ignorar error si no es un JSON válido o no viene de Telegram
    }
  }
  try {
    const ss = SpreadsheetApp.openById(credenciales.SHEET_ID);

    // --- ACCIÓN: FIRMAR RESCISIÓN DE CONTRATO ---
    if (e.parameter.action === 'firmarRescisionCliente') {
      const rawImei = String(e.parameter.imei || "").trim();
      const firmaBase64 = e.parameter.firmaBase64 || "";
      
      const normalizarVal = (val) => {
        if (!val) return "";
        let str = String(val).trim();
        if (str.toLowerCase().includes('e')) {
          let num = Number(val);
          if (!isNaN(num)) str = num.toFixed(0);
        }
        return str.replace(/[^0-9]/g, "");
      };
      
      const imeiNormalizado = normalizarVal(rawImei);
      if (!imeiNormalizado || !firmaBase64) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Faltan parámetros requeridos (IMEI o firma)" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const sheetClientes = ss.getSheetByName('Clientes');
      const sheetPagos = ss.getSheetByName('Pagos');
      const sheetInv = ss.getSheetByName('Inventario');
      const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
      
      if (!sheetClientes) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Pestaña Clientes no encontrada" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Buscar cliente por IMEI
      const ultFilaClientes = sheetClientes.getLastRow();
      let clienteEncontrado = null;
      let filaIndexCliente = -1;
      
      if (ultFilaClientes > 1) {
        const datosClientes = sheetClientes.getRange(2, 1, ultFilaClientes - 1, 21).getValues();
        for (let i = 0; i < datosClientes.length; i++) {
          const fila = datosClientes[i];
          const imeiFila = normalizarVal(fila[5]);
          
          if (imeiFila === imeiNormalizado) {
            clienteEncontrado = {
              id: String(fila[19] || "").trim(),
              cliente: String(fila[1] || "").trim(),
              telefono: String(fila[2] || "").trim(),
              modelo: String(fila[4] || "").trim(),
              cuota: parseFloat(String(fila[10] || "0").replace(/[\$,]/g, "")) || 0,
              totalFinanciado: parseFloat(String(fila[15] || "0").replace(/[\$,]/g, "")) || 0,
              diaRaya: String(fila[18] || "").trim().toUpperCase(),
              tipoPeriodo: String(fila[20] || "SEMANAL").trim().toUpperCase(),
              fechaInicio: fila[0],
              engancheFila: parseFloat(String(fila[6] || "0").replace(/[\$,]/g, "")) || 0
            };
            filaIndexCliente = i + 2; // Fila real en el excel
            break;
          }
        }
      }
      
      if (!clienteEncontrado) {
        return ContentService.createTextOutput(JSON.stringify({ error: "No se encontró ningún equipo activo con el IMEI proporcionado" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Obtener precio de contado desde Inventario
      let precioContado = clienteEncontrado.totalFinanciado;
      if (sheetInv) {
        const dataInv = sheetInv.getDataRange().getValues();
        for (let i = 1; i < dataInv.length; i++) {
          if (normalizarVal(dataInv[i][0]) === imeiNormalizado) {
            precioContado = parseFloat(String(dataInv[i][6]).replace(/[\$,]/g, "")) || precioContado;
            break;
          }
        }
      }
      
      // Obtener parámetros globales de Configuración
      const filaConfig = sheetConfig ? sheetConfig.getRange("A2:N2").getValues()[0] : [];
      let interesTasaRaw = parseFloat(filaConfig[0]);
      if (isNaN(interesTasaRaw)) {
        interesTasaRaw = 5;
      } else if (interesTasaRaw > 0 && interesTasaRaw <= 1) {
        interesTasaRaw = interesTasaRaw * 100;
      }
      const interesTasa = interesTasaRaw;
      const semTol = parseFloat(filaConfig[2]) || 3;
      
      // Calcular pagos ordinarios y saldo
      const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
      let totalPagosCapital = 0;
      for (let j = 1; j < dataPagos.length; j++) {
        if (normalizarVal(dataPagos[j][0]) === imeiNormalizado) {
          totalPagosCapital += parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
        }
      }
      
      const totalCuotasPactadas = Math.round(clienteEncontrado.totalFinanciado / clienteEncontrado.cuota);
      const cuotasPagadas = Math.min(totalCuotasPactadas, Math.round(totalPagosCapital / clienteEncontrado.cuota));
      const diasPorPeriodo = (clienteEncontrado.tipoPeriodo === "QUINCENAL") ? 15 : 7;
      const startTimestamp = clienteEncontrado.fechaInicio;
      
      const normalizarDia = (diaStr) => {
        const diasIngles = {
          "monday": "LUNES", "tuesday": "MARTES", "wednesday": "MIERCOLES",
          "thursday": "JUEVES", "friday": "VIERNES", "saturday": "SABADO", "sunday": "DOMINGO"
        };
        const normalized = diaStr.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        if (diasIngles[normalized]) return diasIngles[normalized];
        return normalized.toUpperCase();
      };
      
      let proxVencimientoStr = "N/A";
      let semanasAtraso = 0;
      let interesMoratorioAcumulado = 0;
      
      if (startTimestamp instanceof Date) {
        const diasSemana = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
        let indexRaya = diasSemana.indexOf(normalizarDia(clienteEncontrado.diaRaya));
        if (indexRaya === -1) indexRaya = 0;
        
        let primerDiaRaya = new Date(startTimestamp.getTime());
        while (primerDiaRaya.getDay() !== indexRaya) {
          primerDiaRaya.setDate(primerDiaRaya.getDate() + 1);
        }
        primerDiaRaya.setHours(0, 0, 0, 0);
        
        const vencimientoDate = new Date(primerDiaRaya.getTime() + cuotasPagadas * diasPorPeriodo * 24 * 60 * 60 * 1000);
        
        const hoy = new Date();
        hoy.setHours(0, 0, 0, 0);
        
        const diffTime = hoy.getTime() - vencimientoDate.getTime();
        const diasAtraso = Math.max(0, Math.floor(diffTime / (24 * 60 * 60 * 1000)));
        semanasAtraso = Math.floor(diasAtraso / 7);
        interesMoratorioAcumulado = diasAtraso * clienteEncontrado.cuota * (interesTasa / 100);
      }
      
      const montoEnganche = clienteEncontrado.engancheFila;
      const saldoPendienteRescision = Math.max(0, Math.round(precioContado - montoEnganche - totalPagosCapital + interesMoratorioAcumulado));
      
      // 1. Guardar la firma del cliente en Google Drive
      let firmaUrlLarga = "";
      let firmaLinkAutocrat = "";
      try {
        const base64Data = firmaBase64.split(',')[1] || firmaBase64;
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', `${clienteEncontrado.cliente}_Firma_Rescision.png`);
        const file = DriveApp.getFolderById(credenciales.RESCIND_FOLDER_ID).createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        firmaUrlLarga = file.getUrl();
        firmaLinkAutocrat = "https://drive.google.com/uc?export=download&id=" + file.getId();
      } catch (errFirma) {
        console.error("Error guardando firma rescision en Drive: " + errFirma.toString());
      }
      
      // 2. Enviar el webhook a n8n para generar el PDF de Rescisión
      let n8nRescisionUrl = "http://107.175.122.33:5678/webhook/6/webhook/generar-rescision";
      
      const formatedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "America/Mexico_City", "dd/MM/yyyy");
      
      const payloadRescision = {
        cliente: clienteEncontrado.cliente,
        idCliente: clienteEncontrado.id,
        modelo: clienteEncontrado.modelo,
        imei: imeiNormalizado,
        plazo: `${totalCuotasPactadas} PAGOS ${clienteEncontrado.tipoPeriodo}ES`,
        tipoPeriodo: clienteEncontrado.tipoPeriodo,
        semanasAtraso: semanasAtraso,
        precioContado: Math.round(precioContado),
        montoEnganche: Math.round(montoEnganche),
        totalPagosCapital: Math.round(totalPagosCapital),
        interesMoratorioAcumulado: Math.round(interesMoratorioAcumulado),
        interes: Math.round(interesTasa),
        saldoPendiente: Math.round(saldoPendienteRescision),
        fechaNotificacion: formatedDate,
        firmaBase64: firmaBase64
      };
      
      try {
        const response = UrlFetchApp.fetch(n8nRescisionUrl, {
          method: "POST",
          contentType: "application/json",
          payload: JSON.stringify(payloadRescision),
          muteHttpExceptions: true
        });
        
        const responseCode = response.getResponseCode();
        if (responseCode === 200) {
          const pdfBlob = Utilities.newBlob(response.getContent(), 'application/pdf', `Rescision_Contrato_${clienteEncontrado.cliente}_${imeiNormalizado}.pdf`);
          
           // Guardar PDF en Drive (dentro de carpeta Contratos Rescindidos fija) y obtener enlace
          let pdfUrlDrive = "";
          try {
            const folder = DriveApp.getFolderById(credenciales.RESCIND_FOLDER_ID);
            const filePdf = folder.createFile(pdfBlob);
            filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            pdfUrlDrive = "https://drive.google.com/uc?export=download&id=" + filePdf.getId();
          } catch(eDrive) {
            console.error("Error guardando PDF rescision en Drive: " + eDrive.toString());
          }
          
          // Enviar por correo al administrador
          try {
            const subject = `Rescisión de Contrato de ${clienteEncontrado.cliente} (Firmado)`;
            const body = `Se adjunta el documento de Rescisión de Contrato firmado digitalmente por el cliente.\n\n` +
                         `Cliente: ${clienteEncontrado.cliente}\n` +
                         `IMEI: ${imeiNormalizado}\n` +
                         `ID Cliente: ${clienteEncontrado.id}\n` +
                         `Fecha: ${formatedDate}\n\n` +
                         `Equipo: ${clienteEncontrado.modelo}`;
            
            GmailApp.sendEmail("celyamex@gmail.com", subject, body, {
              attachments: [pdfBlob]
            });
          } catch(eMail) {
            console.error("Error enviando email rescision: " + eMail.toString());
          }
          
          return ContentService.createTextOutput(JSON.stringify({ 
            result: "success", 
            message: "Rescisión firmada y PDF generado con éxito",
            pdfUrl: pdfUrlDrive
          })).setMimeType(ContentService.MimeType.JSON);
        } else {
          return ContentService.createTextOutput(JSON.stringify({ error: "Error del servidor n8n: " + response.getContentText() })).setMimeType(ContentService.MimeType.JSON);
        }
      } catch(errN8n) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Error de red con n8n: " + errN8n.toString() })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- ACCIÓN: ACTUALIZAR PRECIOS DE COSTO DEL CATÁLOGO EN BATCH (MERCADO LIBRE) ---
    if (e.parameter.action === 'actualizarPreciosCatalogoMasivo') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const postData = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : null;
      if (!postData || !postData.actualizaciones || !Array.isArray(postData.actualizaciones)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Datos de actualización inválidos" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const sheetCat = ss.getSheetByName('Catalogo');
      if (!sheetCat) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se encontró pestaña 'Catalogo'" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const actualizaciones = postData.actualizaciones;
      let contador = 0;
      
      // Realizar las actualizaciones
      for (let i = 0; i < actualizaciones.length; i++) {
        const item = actualizaciones[i];
        const fila = parseInt(item.fila);
        const nuevoPrecio = parseFloat(item.precio);
        
        if (fila > 1 && !isNaN(nuevoPrecio)) {
          // Columna D: ML Costo (Columna 4)
          sheetCat.getRange(fila, 4).setValue(nuevoPrecio);
          contador++;
        }
      }
      
      SpreadsheetApp.flush();
      escribirLogDebug("Sincronización Catálogo ML: " + contador + " precios actualizados por n8n.");
      return ContentService.createTextOutput(JSON.stringify({ result: "success", message: contador + " precios actualizados correctamente" })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- ACCIÓN: LOGIN DE SEGURIDAD ---
    if (e.parameter.action === 'login') {
      if (e.parameter.password === credenciales.PASSWORD) {
        return ContentService.createTextOutput(JSON.stringify({ result: "success", token: credenciales.API_TOKEN })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Contraseña incorrecta" })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // --- ACCIÓN: PROCESAR COTIZACIÓN DE CRÉDITO DESDE CATALOGO.HTML ---
    if (e.parameter.action === 'registrarCotizacionPublica') {
      let sheetCot = ss.getSheetByName('Cotizaciones');
      
      // Si es la primera cotización de la historia y no existe la pestaña, la creamos limpia
      if (!sheetCot) {
        sheetCot = ss.insertSheet('Cotizaciones');
        sheetCot.appendRow(["Fecha", "Cliente", "WhatsApp / Teléfono", "Enlace Mercado Libre", "Plazo Solicitado", "Estado de Crédito", "Sueldo", "Préstamos Activos", "Color", "Capacidad", "Link Prellenado"]);
      }
      
      const nombre = e.parameter.nombre || "";
      const telefono = e.parameter.telefono || "";
      const link = e.parameter.link || "";
      const plazo = e.parameter.plazo || "";
      const sueldoStr = e.parameter.sueldo || ""; // weekly income option
      const prestamosStr = e.parameter.prestamos || ""; // weekly deudas option
      const cuota = parseFloat(e.parameter.cuota) || 150; // weekly CelYa payment
      const color = e.parameter.color || "";
      const capacidad = e.parameter.capacidad || "";
      const planElegido = e.parameter.planElegido || "";
      const frecuencia = e.parameter.frecuencia || "";
      const linkPrellenado = e.parameter.linkPrellenado || "";
      
      const precioCostoTg     = parseFloat(e.parameter.precioCosto)   || 0;
      const precioContadoTg   = parseFloat(e.parameter.precioContado) || 0;
      const engancheTg        = parseFloat(e.parameter.engancheCalc)  || 0;
      
      // Mapeo de Ingresos semanales a valores representativos (medias)
      let ingresoSemanal = 1000;
      if (sueldoStr.includes("$1,500 a $3,000") || sueldoStr.includes("$3,000 a $6,000")) {
        ingresoSemanal = 2250;
      } else if (sueldoStr.includes("$1,500 a $2,500") || sueldoStr.includes("$3,000 a $5,000 quincenales")) {
        ingresoSemanal = 2000;
      } else if (sueldoStr.includes("$2,500 a $3,500") || sueldoStr.includes("$5,000 a $7,000")) {
        ingresoSemanal = 3000;
      } else if (sueldoStr.includes("$3,500 a $5,000") || sueldoStr.includes("$7,000 a $10,000")) {
        ingresoSemanal = 4250;
      } else if (sueldoStr.includes("$3,000 a $5,000")) {
        ingresoSemanal = 4000;
      } else if (sueldoStr.includes("Más de $5,000") || sueldoStr.includes("Mas de $5,000") || sueldoStr.includes("Más de $10,000") || sueldoStr.includes("Mas de $10,000")) {
        ingresoSemanal = 6000;
      }
      
      // Mapeo de Deudas semanales a valores representativos (medias)
      let deudasSemanal = 0;
      if (prestamosStr.includes("Menos de $500") || prestamosStr.includes("Menos de $1,000")) {
        deudasSemanal = 250;
      } else if (prestamosStr.includes("$500 a $1,000") || prestamosStr.includes("$1,000 a $2,000")) {
        deudasSemanal = 750;
      } else if (prestamosStr.includes("Más de $1,000") || prestamosStr.includes("Mas de $1,000") || prestamosStr.includes("Más de $2,000") || prestamosStr.includes("Mas de $2,000")) {
        deudasSemanal = 1500;
      }
      
      const margenLibre = ingresoSemanal - deudasSemanal;
      
      // Convertir a cuota semanal equivalente para la lógica de pre-calificación semanal si es quincenal
      let cuotaSemanalParaValidar = cuota;
      if (plazo.toUpperCase().includes("QUINCENA") || plazo.toUpperCase().includes("QUINCENAL")) {
        cuotaSemanalParaValidar = cuota / 2;
      }

      const esContado = plazo.toUpperCase().includes("CONTADO") || plazo.toUpperCase().includes("PAGO DE CONTADO");

      // Lógica de Pre-calificación / Regla del 50% y Margen Libre
      let esViable = true;
      let estadoCredito = "APROBADO AUTOMATICO";
      
      if (esContado) {
        estadoCredito = "COMPRA DE CONTADO";
      } else {
        if (deudasSemanal >= (ingresoSemanal * 0.5)) {
          esViable = false;
          estadoCredito = "RECHAZADO AUTOMATICO - MOROSO SEGURO (DEUDAS >= 50%)";
        } else if (margenLibre < cuotaSemanalParaValidar) {
          esViable = false;
          estadoCredito = "RECHAZADO AUTOMATICO - MARGEN INSUFICIENTE";
        } else if (margenLibre < (cuotaSemanalParaValidar * 2)) {
          esViable = true;
          estadoCredito = "REVISION MANUAL";
        }
      }
      
      let fechaHoy = "";
      try {
        fechaHoy = Utilities.formatDate(new Date(), "America/Mexico_City", "dd/MM/yyyy HH:mm:ss");
      } catch(errDate) {
        fechaHoy = new Date().toLocaleString('es-MX');
      }
      
      // Generar el ID Único anticipado para este cliente
      const sheetClientes = ss.getSheetByName('Clientes');
      const idCliente = generarSiguienteIdCelYa(sheetClientes);

      // Asegurarse de tener las columnas necesarias en la cabecera
      const numCols = Math.max(25, sheetCot.getLastColumn());
      const cabecera = sheetCot.getRange(1, 1, 1, numCols).getValues()[0];
      
      const cabeceraLower = cabecera.map(h => String(h || '').toLowerCase().trim());

      let colColor = cabeceraLower.indexOf("color") + 1;
      let colCapacidad = cabeceraLower.indexOf("capacidad") + 1;
      let colPlanElegido = cabeceraLower.indexOf("plan elegido") + 1;
      let colFrecuencia = cabeceraLower.indexOf("frecuencia") + 1;
      let colIdCliente = cabeceraLower.indexOf("id cliente") + 1;
      let colLinkPrellenado = cabeceraLower.indexOf("link prellenado") + 1;
      let colEstadoCotizacion = cabeceraLower.indexOf("estado de cotización") + 1;
      if (colEstadoCotizacion === 0) colEstadoCotizacion = cabeceraLower.indexOf("estado de cotizacion") + 1;
      
      let colPrecioCosto = cabeceraLower.indexOf("precio costo") + 1;
      let colPrecioContado = cabeceraLower.indexOf("precio contado") + 1;
      let colEngancheCotizado = cabeceraLower.indexOf("enganche cotizado") + 1;

      let nextCol = numCols + 1;
      if (colColor === 0) {
        colColor = nextCol++;
        sheetCot.getRange(1, colColor).setValue("Color");
      }
      if (colCapacidad === 0) {
        colCapacidad = nextCol++;
        sheetCot.getRange(1, colCapacidad).setValue("Capacidad");
      }
      if (colPlanElegido === 0) {
        colPlanElegido = nextCol++;
        sheetCot.getRange(1, colPlanElegido).setValue("Plan Elegido");
      }
      if (colFrecuencia === 0) {
        colFrecuencia = nextCol++;
        sheetCot.getRange(1, colFrecuencia).setValue("Frecuencia");
      }
      if (colIdCliente === 0) {
        colIdCliente = nextCol++;
        sheetCot.getRange(1, colIdCliente).setValue("ID Cliente");
      }
      if (colLinkPrellenado === 0) {
        colLinkPrellenado = nextCol++;
        sheetCot.getRange(1, colLinkPrellenado).setValue("Link Prellenado");
      }
      if (colEstadoCotizacion === 0) {
        colEstadoCotizacion = nextCol++;
        sheetCot.getRange(1, colEstadoCotizacion).setValue("Estado de Cotización");
      }
      if (colPrecioCosto === 0) {
        colPrecioCosto = nextCol++;
        sheetCot.getRange(1, colPrecioCosto).setValue("Precio Costo");
      }
      if (colPrecioContado === 0) {
        colPrecioContado = nextCol++;
        sheetCot.getRange(1, colPrecioContado).setValue("Precio Contado");
      }
      if (colEngancheCotizado === 0) {
        colEngancheCotizado = nextCol++;
        sheetCot.getRange(1, colEngancheCotizado).setValue("Enganche Cotizado");
      }

      // Insertamos la fila en tu base de datos para auditoría e historial
      const nuevaFila = sheetCot.getLastRow() + 1;
      
      let finalPrefilledUrl = linkPrellenado;
      if (finalPrefilledUrl) {
        if (finalPrefilledUrl.indexOf("?") === -1) {
          finalPrefilledUrl += `?pre_cotizacion_id=${nuevaFila}&pre_client_id=${idCliente}`;
        } else {
          finalPrefilledUrl += `&pre_cotizacion_id=${nuevaFila}&pre_client_id=${idCliente}`;
        }
      }

      sheetCot.getRange(nuevaFila, 1).setValue(fechaHoy);
      sheetCot.getRange(nuevaFila, 2).setValue(nombre);
      sheetCot.getRange(nuevaFila, 3).setValue(telefono);
      if (link && (link.indexOf("http://") === 0 || link.indexOf("https://") === 0)) {
        const safeLinkVal = link.replace(/"/g, '%22');
        sheetCot.getRange(nuevaFila, 4).setFormula(`=HYPERLINK("${safeLinkVal}", "link")`);
      } else {
        sheetCot.getRange(nuevaFila, 4).setValue(link);
      }
      sheetCot.getRange(nuevaFila, 5).setValue(plazo);
      sheetCot.getRange(nuevaFila, 6).setValue(estadoCredito);
      sheetCot.getRange(nuevaFila, 7).setValue(sueldoStr);
      sheetCot.getRange(nuevaFila, 8).setValue(prestamosStr);
      
      sheetCot.getRange(nuevaFila, colColor).setValue(color);
      sheetCot.getRange(nuevaFila, colCapacidad).setValue(capacidad);
      sheetCot.getRange(nuevaFila, colPlanElegido).setValue(planElegido);
      sheetCot.getRange(nuevaFila, colFrecuencia).setValue(frecuencia);
      if (colIdCliente > 0) {
        sheetCot.getRange(nuevaFila, colIdCliente).setValue(idCliente);
      }
      if (colPrecioCosto > 0) {
        sheetCot.getRange(nuevaFila, colPrecioCosto).setValue(parseFloat(e.parameter.precioCosto) || 0);
      }
      if (colPrecioContado > 0) {
        sheetCot.getRange(nuevaFila, colPrecioContado).setValue(parseFloat(e.parameter.precioContado) || 0);
      }
      if (colEngancheCotizado > 0) {
        sheetCot.getRange(nuevaFila, colEngancheCotizado).setValue(engancheTg);
      }
      
      if (colLinkPrellenado > 0) {
        if (finalPrefilledUrl && (finalPrefilledUrl.indexOf("http://") === 0 || finalPrefilledUrl.indexOf("https://") === 0)) {
          const safePrefilledVal = finalPrefilledUrl.replace(/"/g, '%22');
          sheetCot.getRange(nuevaFila, colLinkPrellenado).setFormula(`=HYPERLINK("${safePrefilledVal}", "Abrir Contrato")`);
        } else {
          sheetCot.getRange(nuevaFila, colLinkPrellenado).setValue(finalPrefilledUrl);
        }
      }

      // Marcar la cotización como Pendiente
      if (colEstadoCotizacion > 0) {
        sheetCot.getRange(nuevaFila, colEstadoCotizacion).setValue("Pendiente");
      }
      
      // 🤖 ALERTA AL ASISTENTE DE TELEGRAM (Solo se envía si el cliente es Viable/Revisión Manual/Contado)
      if (esViable && credenciales.TELEGRAM_TOKEN && credenciales.CHAT_ID) {
        let advertenciaJusto = "";
        if (estadoCredito === "REVISION MANUAL") {
          advertenciaJusto = `⚠️ <b>OJO:</b> Este cliente viene condicionado por ingresos justos.\n\n`;
        }
        
        let linkPrellenadoHTML = "";
        if (finalPrefilledUrl) {
          linkPrellenadoHTML = `📝 <b>Contrato Prellenado:</b>\n<a href="${finalPrefilledUrl}">Abrir Contrato</a>\n\n`;
        }

        // Sanitización y escape de variables para HTML de Telegram
        const nombreEscaped = escapeHTML(nombre);
        const telefonoEscaped = escapeHTML(telefono);
        const idClienteEscaped = escapeHTML(idCliente || "");
        const sueldoStrEscaped = escapeHTML(sueldoStr);
        const prestamosStrEscaped = escapeHTML(prestamosStr);
        const estadoCreditoEscaped = escapeHTML(estadoCredito);
        const plazoEscaped = escapeHTML(plazo);
        const colorEscaped = escapeHTML(color || "No especificado");
        const capacidadEscaped = escapeHTML(capacidad || "No especificada");

        // Construcción limpia del enlace del equipo (HTML)
        let linkTelegram = "";
        if (link && (link.indexOf("http://") === 0 || link.indexOf("https://") === 0)) {
          const celName = extraerNombreCelular(plazo, color, capacidad);
          const celSanitizado = sanitizarNombreCelular(celName);
          linkTelegram = `<a href="${link}">${escapeHTML(celSanitizado)}</a>`;
        } else {
          linkTelegram = escapeHTML(link);
        }

        // ── Desglose de precio enviado directamente desde el frontend ─────────

        let bloquePrecio = "";
        if (precioContadoTg > 0 && precioCostoTg > 0) {
          let sheetConfigTgVal = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
          const gananciaPct = sheetConfigTgVal
            ? (parseFloat(sheetConfigTgVal.getRange("F2").getValue()) || 15)
            : 15;
          const gananciaAmt = (precioContadoTg - precioCostoTg).toLocaleString('es-MX');
          bloquePrecio = `\n💡 <b>Desglose de Precio CelYa:</b>\n` +
                         `   • Costo ML: $${precioCostoTg.toLocaleString('es-MX')} MXN\n` +
                         `   • Ganancia CelYa +${gananciaPct}%: $${gananciaAmt} MXN\n` +
                         `   • Precio Contado: $${precioContadoTg.toLocaleString('es-MX')} MXN\n` +
                         `   • Enganche cotizado: $${engancheTg.toLocaleString('es-MX')} MXN\n`;
        }
        // ─────────────────────────────────────────────────────────────────────

        let mensajeTelegram = "";
        if (esContado) {
          mensajeTelegram = `💰 <b>¡NUEVA COMPRA DE CONTADO!</b> 💰\n\n` +
                            `👤 <b>Cliente:</b> ${nombreEscaped}\n` +
                            `📱 <b>WhatsApp:</b> ${telefonoEscaped}\n` +
                            `🆔 <b>ID Cliente:</b> ${idClienteEscaped}\n` +
                            `📝 <b>Estado:</b> COMPRA DE CONTADO\n` +
                            `🗓️ <b>Detalles:</b> ${plazoEscaped}\n` +
                            `🎨 <b>Color:</b> ${colorEscaped}\n` +
                            `💾 <b>Capacidad:</b> ${capacidadEscaped}\n` +
                            `${bloquePrecio}\n` +
                            `${linkPrellenadoHTML}` +
                            `🔗 <b>Link / Equipo:</b> ${linkTelegram}`;
        } else {
          mensajeTelegram = `🚨 <b>¡NUEVA SOLICITUD DE CRÉDITO!</b> 🚨\n\n` +
                            `${advertenciaJusto}` +
                            `👤 <b>Cliente:</b> ${nombreEscaped}\n` +
                            `📱 <b>WhatsApp:</b> ${telefonoEscaped}\n` +
                            `🆔 <b>ID Cliente:</b> ${idClienteEscaped}\n` +
                            `💰 <b>Ingresos:</b> ${sueldoStrEscaped}\n` +
                            `🏦 <b>Deudas:</b> ${prestamosStrEscaped}\n` +
                            `💵 <b>Cuota CelYa:</b> $${cuota.toLocaleString('es-MX')}\n` +
                            `📊 <b>Margen Libre Estimado:</b> $${margenLibre.toLocaleString('es-MX')}/sem\n` +
                            `📝 <b>Estado:</b> ${estadoCreditoEscaped}\n` +
                            `🗓️ <b>Plazo elegido:</b> ${plazoEscaped}\n` +
                            `🎨 <b>Color:</b> ${colorEscaped}\n` +
                            `💾 <b>Capacidad:</b> ${capacidadEscaped}\n` +
                            `${bloquePrecio}\n` +
                            `${linkPrellenadoHTML}` +
                            `🔗 <b>Link / Equipo:</b> ${linkTelegram}`;
        }
                                
        enviarMensajeTelegram(credenciales, mensajeTelegram);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ result: esViable ? "success" : "rejected", prefilledUrl: finalPrefilledUrl, linkPrellenado: finalPrefilledUrl })).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === "actualizarConfiguracionGlobal" || e.parameter.action === "updateConfig") {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
      if (!sheetConfig) throw new Error("Falta pestaña Configuracion");

      const interes = e.parameter.interes;
      const horastol = e.parameter.horasTolerancia !== undefined ? e.parameter.horasTolerancia : e.parameter.horastol;
      const semanastol = e.parameter.semanasTolerancia !== undefined ? e.parameter.semanasTolerancia : e.parameter.semanastol;
      const interesPlanRapido = e.parameter.planRapido !== undefined ? e.parameter.planRapido : e.parameter.interesPlanRapido;
      const interesPlanComodo = e.parameter.planComodo !== undefined ? e.parameter.planComodo : e.parameter.interesPlanComodo;
      const gananciaContado = e.parameter.gananciaContado;
      const engancheSemanal = e.parameter.engancheSemanal;

      const engancheComodo = e.parameter.engancheComodo;
      const descuentoLiquidacion = e.parameter.descuentoLiquidacion;
      const telegramBotToken = e.parameter.telegramBotToken;
      const telegramChatId = e.parameter.telegramChatId;
      const tasaInteresQuincenalRapido = e.parameter.tasaInteresQuincenalRapido;
      const tasaInteresQuincenalComodo = e.parameter.tasaInteresQuincenalComodo;
      const engancheMinimoQuincenalRapido = e.parameter.engancheMinimoQuincenalRapido;
      const engancheMinimoQuincenalComodo = e.parameter.engancheMinimoQuincenalComodo;

      sheetConfig.getRange("A2:E2").setValues([[
        interes, 
        horastol, 
        semanastol, 
        interesPlanRapido, 
        interesPlanComodo
      ]]);
      
      if (gananciaContado !== undefined) {
        sheetConfig.getRange("F2").setValue(gananciaContado);
      }
      if (engancheSemanal !== undefined) {
        sheetConfig.getRange("G2").setValue(engancheSemanal);
      }

      if (engancheComodo !== undefined) {
        sheetConfig.getRange("I2").setValue(engancheComodo);
      }
      if (descuentoLiquidacion !== undefined) {
        sheetConfig.getRange("J2").setValue(descuentoLiquidacion);
      }
      if (telegramBotToken !== undefined) {
        sheetConfig.getRange("K2").setValue(telegramBotToken);
      }
      if (telegramChatId !== undefined) {
        sheetConfig.getRange("L2").setValue(telegramChatId);
      }
      if (tasaInteresQuincenalRapido !== undefined) {
        sheetConfig.getRange("M2").setValue(tasaInteresQuincenalRapido);
      }
      if (tasaInteresQuincenalComodo !== undefined) {
        sheetConfig.getRange("N2").setValue(tasaInteresQuincenalComodo);
      }
      if (engancheMinimoQuincenalRapido !== undefined) {
        sheetConfig.getRange("O2").setValue(engancheMinimoQuincenalRapido);
      }
      if (engancheMinimoQuincenalComodo !== undefined) {
        sheetConfig.getRange("P2").setValue(engancheMinimoQuincenalComodo);
      }
      const codigoDescuento = e.parameter.codigoDescuento;
      if (codigoDescuento !== undefined) {
        sheetConfig.getRange("Q2").setValue(codigoDescuento);
      }
      return ContentService.createTextOutput(JSON.stringify({result: "success", message: "Configuración actualizada"})).setMimeType(ContentService.MimeType.JSON);
    }

    // REGISTRAR UN PAGO NUEVO
    if (e.parameter.action === "registrarPago") {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetPagos = ss.getSheetByName('Pagos');
      if (!sheetPagos) throw new Error("No se encontró la pestaña 'Pagos'");
      
      const timestamp = new Date();
      const imeiPago = e.parameter.imei || "";
      const nombreCliente = e.parameter.nombreCliente || "";
      const montoPago = e.parameter.monto ? parseFloat(e.parameter.monto.replace(/[\$,]/g, "").trim()) : 0;
      const metodoPago = e.parameter.metodo || "";
      const semanaPago = e.parameter.semana || "";
      const folioPago = e.parameter.folio || "";
      const notasPago = e.parameter.notas || "";
      
      sheetPagos.appendRow([imeiPago, nombreCliente, timestamp, montoPago, metodoPago, semanaPago, folioPago, notasPago]);
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({result: "success", message: "Pago guardado correctamente"})).setMimeType(ContentService.MimeType.JSON);
    }

    // REGISTRAR PROMESA DE PAGO
    if (e.parameter.action === 'registrarPromesa') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      let sheetPromesas = ss.getSheetByName('Promesas');
      if (!sheetPromesas) {
        sheetPromesas = ss.insertSheet('Promesas');
        const headersPromesas = ["IMEI", "Cliente", "Fecha Promesa", "Monto Prometido", "Estatus", "Fecha Registro", "Notas"];
        sheetPromesas.getRange(1, 1, 1, headersPromesas.length).setValues([headersPromesas]);
      }
      
      const imeiPromesa = String(e.parameter.imei || '').trim();
      const clientePromesa = String(e.parameter.cliente || '').trim();
      const fechaPromesa = String(e.parameter.fechaPromesa || '').trim();
      const montoPrometido = parseFloat(e.parameter.montoPrometido) || 0;
      const notasPromesa = String(e.parameter.notes || e.parameter.notas || '').trim();

      sheetPromesas.appendRow([
        imeiPromesa,
        clientePromesa,
        fechaPromesa,
        montoPrometido,
        "PENDIENTE",
        new Date(),
        notasPromesa
      ]);
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Promesa de pago registrada exitosamente' })).setMimeType(ContentService.MimeType.JSON);
    }

    // ACTUALIZAR ESTATUS DE PROMESA DE PAGO
    if (e.parameter.action === 'actualizarEstatusPromesa') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetPromesas = ss.getSheetByName('Promesas');
      if (!sheetPromesas) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró la pestaña de promesas' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const imeiPromesa = String(e.parameter.imei || '').trim();
      const nuevoEstatus = String(e.parameter.estatus || '').trim().toUpperCase();
      
      if (!imeiPromesa) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'IMEI requerido' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const dataProm = sheetPromesas.getDataRange().getValues();
      let encontrada = false;
      
      // Buscamos la última promesa PENDIENTE para este IMEI y la actualizamos
      for (let k = dataProm.length - 1; k >= 1; k--) {
        const rowImei = String(dataProm[k][0] || '').trim();
        const rowEst = String(dataProm[k][4] || '').trim().toUpperCase();
        if (rowImei === imeiPromesa && rowEst === "PENDIENTE") {
          sheetPromesas.getRange(k + 1, 5).setValue(nuevoEstatus); // Columna E (Estatus) es columna 5
          encontrada = true;
          break; // Solo actualizamos la última
        }
      }
      
      SpreadsheetApp.flush();
      
      if (encontrada) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Estatus de promesa actualizado correctamente' })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró una promesa pendiente para este IMEI' })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // CAMBIAR ESTADO DE EQUIPO EN INVENTARIO
    if (e.parameter.action === 'cambiarEstadoEquipo') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetInv = ss.getSheetByName('Inventario');
      if (!sheetInv) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró la pestaña Inventario' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const imeiKey = String(e.parameter.imei || '').trim();
      const nuevoEstado = String(e.parameter.estado || '').trim().toUpperCase();
      
      if (!imeiKey || !nuevoEstado) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'IMEI y estado requeridos' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const dataInv = sheetInv.getDataRange().getValues();
      let encontrado = false;
      for (let i = 1; i < dataInv.length; i++) {
        const rowImei = String(dataInv[i][0] || '').trim();
        if (rowImei === imeiKey) {
          sheetInv.getRange(i + 1, 5).setValue(nuevoEstado); // Columna E (Estado) es columna 5
          encontrado = true;
          break;
        }
      }
      
      SpreadsheetApp.flush();
      if (encontrado) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Estado del equipo actualizado exitosamente' })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró el IMEI en el Inventario' })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ENVIAR NOTIFICACIÓN A TELEGRAM
    if (e.parameter.action === 'enviarNotificacionTelegram') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
      if (!sheetConfig) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró la pestaña Configuración' })).setMimeType(ContentService.MimeType.JSON);
      }
      const filaConfig = sheetConfig.getRange("A2:L2").getValues()[0];
      const botToken = String(filaConfig[10] || '').trim();
      const chatId = String(filaConfig[11] || '').trim();
      
      const texto = e.parameter.mensaje || "";
      if (!botToken || !chatId) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Configuración de Telegram incompleta en la hoja Configuración.' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      enviarMensajeTelegram(botToken, chatId, texto);
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Notificación enviada a Telegram' })).setMimeType(ContentService.MimeType.JSON);
    }

    // PROBAR RESUMEN MATUTINO TELEGRAM
    if (e.parameter.action === 'probarResumenMatutino') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      enviarResumenMatutinoTelegram();
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Resumen matutino de prueba enviado a Telegram' })).setMimeType(ContentService.MimeType.JSON);
    }

    // REGISTRAR WEBHOOK DE TELEGRAM
    if (e.parameter.action === 'registrarWebhookTelegram') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
      const filaConfig = sheetConfig ? sheetConfig.getRange("A2:L2").getValues()[0] : null;
      const botToken = filaConfig ? String(filaConfig[10] || '').trim() : "";
      
      if (!botToken) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Introduce primero el Telegram Bot Token en la configuración y guárdalo." })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const webAppUrl = ScriptApp.getService().getUrl();
      if (!webAppUrl || webAppUrl.indexOf("macros/s/") === -1) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se pudo obtener la URL del script. Asegúrate de haber publicado tu Apps Script como Web App (Cualquier persona puede acceder)." })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const telegramUrl = "https://api.telegram.org/bot" + botToken + "/setWebhook?url=" + encodeURIComponent(webAppUrl) + "&drop_pending_updates=true";
      const response = UrlFetchApp.fetch(telegramUrl, { muteHttpExceptions: true });
      const resData = JSON.parse(response.getContentText());
      
      if (resData.ok) {
        return ContentService.createTextOutput(JSON.stringify({ result: "success", message: "¡Webhook de Telegram configurado con éxito!" })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Telegram rechazó la configuración: " + resData.description })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // REGISTRAR CIERRE DE CAJA DIARIO
    if (e.parameter.action === 'registrarCierreCaja') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const sheetCierres = ss.getSheetByName('Cierres') || ss.insertSheet('Cierres');
      const headersCierres = ["Fecha Cierre", "Efectivo Cobrado", "Transferencia Cobrada", "Tarjeta Cobrada", "Total Cobrado", "Diferencia Efectivo", "Notas Cierre"];
      if (sheetCierres.getLastRow() === 0) {
        sheetCierres.getRange(1, 1, 1, headersCierres.length).setValues([headersCierres]);
      }
      
      const fechaCierre = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      const efectivo = parseFloat(e.parameter.efectivo) || 0;
      const transferencia = parseFloat(e.parameter.transferencia) || 0;
      const tarjeta = parseFloat(e.parameter.tarjeta) || 0;
      const total = efectivo + transferencia + tarjeta;
      const diferencia = parseFloat(e.parameter.diferencia) || 0;
      const notas = e.parameter.notas || "";
      
      sheetCierres.appendRow([
        fechaCierre,
        efectivo,
        transferencia,
        tarjeta,
        total,
        diferencia,
        notas
      ]);
      
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Cierre de caja registrado con éxito' })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- ACCIÓN: AGREGAR CELULAR AL CATÁLOGO ---
    if (e.parameter.action === 'agregarCelularCatalogo') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const link = e.parameter.link || "";
      let imei = e.parameter.imei || "";
      const estado = e.parameter.estado || "DISPONIBLE";
      
      if (!link) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Falta el enlace" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      const sheetCat = ss.getSheetByName('Catalogo') || ss.insertSheet('Catalogo');
      if (!sheetCat) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se encontró ni se pudo crear la pestaña 'Catalogo'" })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Si el IMEI está vacío, autogeneramos uno secuencial basado en 999999
      if (!imei) {
        const ultFilaCat = sheetCat.getLastRow();
        let maxSeq = 0;
        if (ultFilaCat > 1) {
          const range = sheetCat.getRange(2, 1, ultFilaCat - 1, 1).getValues();
          range.forEach(r => {
            const val = r[0] ? r[0].toString().trim() : "";
            if (val.startsWith("999999")) {
              const seq = parseInt(val.substring(6)) || 0;
              if (seq > maxSeq) {
                maxSeq = seq;
              }
            }
          });
        }
        const nextSeq = maxSeq + 1;
        const paddedSeq = nextSeq.toString().padStart(9, "0");
        imei = "999999" + paddedSeq;
      }
      
      // Consultar Mercado Libre
      const condicionParam = e.parameter.condicion || "new";
      const resML = obtenerDatosDeMercadoLibre(link, condicionParam);
      if (!resML) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se pudieron obtener los datos de Mercado Libre. Recuerda usar enlaces oficiales de catálogo (/p/) de preferencia." })).setMimeType(ContentService.MimeType.JSON);
      }
      
       const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
       
       const costoImport = parseFloat(e.parameter.costoImportacion) || 0;
       const precioCostoML = Math.round(resML.price) + Math.round(costoImport);
       const esImportParam = e.parameter.esImportacion;
       
       let notasValor = "";
       if (esImportParam === "true" || esImportParam === true || costoImport > 0) {
         notasValor = "IMPORTACION";
       }
       if (condicionParam === "used" || condicionParam === "refurbished") {
         notasValor = notasValor ? notasValor + ", USADO" : "USADO";
       }
       
       const foto = resML.pictures && resML.pictures.length > 0 ? resML.pictures[0].secure_url : "";
       const memoria = resML.memoria || "";
       const color = resML.color || "";
       const modelo = limpiarNombreModelo(resML.title || "");
       
       // Obtener variantes completas de inmediato
       const minifiedML = obtenerDatosVariantesML(link, condicionParam);
       const variantesJson = minifiedML ? JSON.stringify(minifiedML) : "";
       const colDisponibles = extraerColoresDeDatosML(minifiedML);

        // Añadir fila al Catalogo
        sheetCat.appendRow([
          imei,
          modelo,
          "",
          precioCostoML,
          notasValor,
          estado,
          link,
          foto,
          memoria,
          color,
          "informacion", // Relleno inicial para Columna K
          "",            // Columna L: manual override
          colDisponibles // Columna M: coloresDisponibles
        ]);
        
        const ultimaFila = sheetCat.getLastRow();
        
        // Aplicar hipervínculo enriquecido para la columna G
        const celdaLink = sheetCat.getRange(ultimaFila, 7);
        const richText = SpreadsheetApp.newRichTextValue()
                          .setText("Ver en ML")
                          .setLinkUrl(link)
                          .build();
        celdaLink.setRichTextValue(richText);

        // Guardar variantes JSON en la columna K
        if (variantesJson) {
          sheetCat.getRange(ultimaFila, 11).setValue(variantesJson);
        }
       
       SpreadsheetApp.flush();
       
       return ContentService.createTextOutput(JSON.stringify({
         result: "success",
         modelo: modelo,
         precioContado: precioCostoML,
         imei: imei
       })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- ACCIÓN: ACTUALIZAR VARIANTES JSON EN EL CATÁLOGO ---
    if (e.parameter.action === 'actualizarVariantesJson') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ error: "Token inválido" })).setMimeType(ContentService.MimeType.JSON);
      }
      const row = parseInt(e.parameter.row);
      const variantesStr = e.parameter.variantes;
      if (!row || !variantesStr) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "Falta fila o variantes" })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetCat = ss.getSheetByName('Catalogo');
      if (!sheetCat) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "No se encontró pestaña Catalogo" })).setMimeType(ContentService.MimeType.JSON);
      }
      sheetCat.getRange(row, 11).setValue(variantesStr);
      
      try {
        const parsed = JSON.parse(variantesStr);
        const colDisponibles = extraerColoresDeDatosML(parsed);
        if (colDisponibles) {
          sheetCat.getRange(row, 13).setValue(colDisponibles);
        }
      } catch(err) {
        Logger.log("Error al extraer colores para actualizarVariantesJson: " + err.toString());
      }
      
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- ACCIÓN: ACTUALIZAR EL LINK DE UN REGISTRO EN EL CATÁLOGO ---
    if (e.parameter.action === 'actualizarLinkCatalogo') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const row = parseInt(e.parameter.row);
      const nuevoLink = String(e.parameter.link || '').trim();
      if (!row || !nuevoLink) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Fila y link requeridos' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetCat = ss.getSheetByName('Catalogo');
      if (!sheetCat) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró la pestaña Catalogo' })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // 1. Actualizar el link en la columna G (7)
      const celdaLink = sheetCat.getRange(row, 7);
      const richText = SpreadsheetApp.newRichTextValue()
                        .setText("Ver en ML")
                        .setLinkUrl(nuevoLink)
                        .build();
      celdaLink.setRichTextValue(richText);
      
      // 2. Intentar actualizar foto, variantesJson y coloresDisponibles en base al nuevo link
      let successMessage = "Enlace actualizado con éxito.";
      try {
        const minifiedML = obtenerDatosVariantesML(nuevoLink);
        if (minifiedML) {
          const variantesJson = JSON.stringify(minifiedML);
          sheetCat.getRange(row, 11).setValue(variantesJson);
          
          const colDisponibles = extraerColoresDeDatosML(minifiedML);
          if (colDisponibles) {
            sheetCat.getRange(row, 13).setValue(colDisponibles);
          }
          
          if (minifiedML.pictures && minifiedML.pictures.length > 0) {
            const fotoUrl = minifiedML.pictures[0].url || minifiedML.pictures[0].secure_url || "";
            if (fotoUrl) {
              sheetCat.getRange(row, 8).setFormula(`=HYPERLINK("${fotoUrl}", "foto")`);
            }
          }
          successMessage = "Enlace y variantes actualizados con éxito.";
        }
      } catch(err) {
        successMessage = "Enlace actualizado, pero falló la actualización de variantes: " + err.toString();
      }
      
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: successMessage })).setMimeType(ContentService.MimeType.JSON);
    }

    // --- ACCIÓN: CAMBIAR ESTADO DE UN REGISTRO EN EL CATÁLOGO ---
    if (e.parameter.action === 'cambiarEstadoCatalogo') {
      if (!validarTokenApi(credenciales, e.parameter.token)) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
      }
      const row = parseInt(e.parameter.row);
      const nuevoEstado = String(e.parameter.estado || '').trim().toUpperCase();
      if (!row || !nuevoEstado) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Fila y estado requeridos' })).setMimeType(ContentService.MimeType.JSON);
      }
      const sheetCat = ss.getSheetByName('Catalogo');
      if (!sheetCat) {
        return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'No se encontró la pestaña Catalogo' })).setMimeType(ContentService.MimeType.JSON);
      }
      sheetCat.getRange(row, 6).setValue(nuevoEstado);
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Estado del catálogo actualizado exitosamente' })).setMimeType(ContentService.MimeType.JSON);
    }

     // --- ACCIÓN: OBTENER PLACEHOLDERS DESDE LA PESTAÑA 'Cotizaciones' ---
     if (e.parameter.action === 'obtenerCatalogoPublico') {
       const sheetCot = ss.getSheetByName('Cotizaciones');
       if (!sheetCot) {
         return ContentService.createTextOutput(JSON.stringify({ result: 'success', resultados: [] })).setMimeType(ContentService.MimeType.JSON);
       }
       
       const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
       let gananciaContadoVal = 15;
       let tasaAnual26 = 66;
       let tasaAnual52 = 90;
       if (sheetConfig) {
         const valConfig = sheetConfig.getRange("F2").getValue();
         if (valConfig !== undefined && valConfig !== "") gananciaContadoVal = parseFloat(valConfig) || 15;
         const val26 = sheetConfig.getRange("D2").getValue();
         if (val26 !== undefined && val26 !== "") tasaAnual26 = parseFloat(val26) || 66;
         const val52 = sheetConfig.getRange("E2").getValue();
         if (val52 !== undefined && val52 !== "") tasaAnual52 = parseFloat(val52) || 90;
       }
       const factorGananciaContado = 1 + (gananciaContadoVal / 100);
       const factorTotal26 = 1 + (tasaAnual26 / 100);
       const factorTotal52 = 1 + (tasaAnual52 / 100);

       const sheetCat = ss.getSheetByName('Catalogo');
       let catData = [];
       let catFormulas = [];
       let catRichTexts = [];
       if (sheetCat) {
         const catRange = sheetCat.getDataRange();
         catData = catRange.getValues();
         catFormulas = catRange.getFormulas();
         catRichTexts = catRange.getRichTextValues();
       }

       const cotRange = sheetCot.getDataRange();
       const headersCot = sheetCot.getRange(1, 1, 1, Math.max(25, sheetCot.getLastColumn())).getValues()[0].map(h => String(h || '').toLowerCase().trim());
       const allCot = cotRange.getValues();
       const allCotFormulas = cotRange.getFormulas();
       const allCotRichTexts = cotRange.getRichTextValues();
       const resultados = [];

       const getValFromRow = (row, headerName) => {
         const idx = headersCot.indexOf(headerName.toLowerCase().trim());
         return idx !== -1 ? row[idx] : "";
       };

       const getLinkFromRow = (rowIdx, headerName) => {
         const idx = headersCot.indexOf(headerName.toLowerCase().trim());
         if (idx === -1) return "";
         return extraerLinkDeArrays(allCot, allCotFormulas, allCotRichTexts, rowIdx, idx);
       };

       for (let i = 1; i < allCot.length; i++) {
         const row = allCot[i];
         const cliente = getValFromRow(row, "Cliente") || "";
         const link = getLinkFromRow(i, "Enlace Mercado Libre") || "";
         const color = getValFromRow(row, "Color") || "";
         const memoria = getValFromRow(row, "Capacidad") || "";
         const planElegido = getValFromRow(row, "Plan Elegido") || "";
         const frecuencia = getValFromRow(row, "Frecuencia") || "";
         const idCliente = getValFromRow(row, "ID Cliente") || "";
         const estadoCredito = getValFromRow(row, "Estado de Crédito") || getValFromRow(row, "Estado") || "";
         const estadoCotizacion = getValFromRow(row, "Estado de Cotización") || "";
          let fecha = getValFromRow(row, "Fecha") || "";
          if (fecha instanceof Date) {
            fecha = Utilities.formatDate(fecha, "America/Mexico_City", "dd/MM/yyyy HH:mm:ss");
          } else if (fecha) {
            fecha = String(fecha);
          }
         const plazoSolicitado = getValFromRow(row, "Plazo Solicitado") || "";
         const precioContadoCot = parseFloat(getValFromRow(row, "Precio Contado")) || 0;

         if (!cliente) continue;

         // Ignorar cotizaciones ya asignadas
         if (estadoCotizacion.toUpperCase() === "ASIGNADA") continue;

         let modelo = "";
         let precioCosto = 0;
         let foto = "";
         let precioContadoFinalFromSheet = 0;

         if (link && typeof link === "string") {
           let targetMlm = "";
           const match = link.match(/(MLM\-?\d+)/i);
           if (match) targetMlm = match[0].replace("-", "").toUpperCase();

           for (let j = 1; j < catData.length; j++) {
             const catLink = extraerLinkDeArrays(catData, catFormulas, catRichTexts, j, 6);
             let catMlm = "";
             if (catLink) {
               const match2 = catLink.match(/(MLM\-?\d+)/i);
               if (match2) catMlm = match2[0].replace("-", "").toUpperCase();
             }
             if ((targetMlm && catMlm && targetMlm === catMlm) || link.trim() === catLink.trim()) {
               modelo = catData[j][1] || "";
               precioCosto = parseFloat(catData[j][3]) || 0;
               foto = extraerLinkDeArrays(catData, catFormulas, catRichTexts, j, 7) || catData[j][7] || "";
               precioContadoFinalFromSheet = parseFloat(catData[j][11]) || 0;
               break;
             }
           }
         }

         if (!modelo) {
           if (plazoSolicitado) {
             const matchPlazo = plazoSolicitado.match(/\(PEDIDO ESPECIAL:\s*([^)]+)\)/i) || plazoSolicitado.match(/\(([^)]+)\)$/);
             if (matchPlazo) {
               modelo = matchPlazo[1];
             }
           }
         }
         if (!modelo) {
           modelo = "Celular Cotizado";
         }

         let precioContado = precioContadoFinalFromSheet > 0 
           ? precioContadoFinalFromSheet 
           : (precioContadoCot > 0 ? precioContadoCot : Math.round(precioCosto * factorGananciaContado));
         const factorTotal = (planElegido === "52" || planElegido === 52) ? factorTotal52 : factorTotal26;
         const precioFinanciado = Math.round(precioContado * factorTotal);

         resultados.push({
           fila: i + 1,
           cliente: cliente,
           idCliente: idCliente,
           modelo: modelo,
           color: color,
           memoria: memoria,
           link: link,
           foto: foto,
           planElegido: planElegido,
           frecuencia: frecuencia,
           estadoCredito: estadoCredito,
           fecha: fecha,
           precioPlaceholder: precioFinanciado,
           precioContado: precioContado
         });
       }
       return ContentService.createTextOutput(JSON.stringify({ result: 'success', resultados: resultados })).setMimeType(ContentService.MimeType.JSON);
     }
 
     // --- ACCIÓN: CONVERTIR PLACEHOLDER DE 'Cotizaciones' A 'Inventario' ---
     if (e.parameter.action === 'convertirPlaceholderAInventario') {
       const tokenValido = validarTokenApi(credenciales, e.parameter.token);
       escribirLogDebug("convertirPlaceholderAInventario - params: tokenValid=" + tokenValido + ", fila=" + e.parameter.fila + ", imei=" + e.parameter.imei + ", modelo=" + e.parameter.modelo + ", precioFinal=" + e.parameter.precioFinal);
       if (!tokenValido) {
         return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Token inválido' })).setMimeType(ContentService.MimeType.JSON);
       }
 
       const sheetCot = ss.getSheetByName('Cotizaciones');
       const sheetInv = ss.getSheetByName('Inventario');
       if (!sheetCot || !sheetInv) {
         return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Faltan pestañas Cotizaciones o Inventario' })).setMimeType(ContentService.MimeType.JSON);
       }
 
       const fila = parseInt(e.parameter.fila, 10); // fila 1-based en la hoja Cotizaciones
       if (!fila || fila <= 1) {
         return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'Fila inválida' })).setMimeType(ContentService.MimeType.JSON);
       }
 
       const headersCot = sheetCot.getRange(1, 1, 1, Math.max(25, sheetCot.getLastColumn())).getValues()[0].map(h => String(h || '').toLowerCase().trim());
       const rowValues = sheetCot.getRange(fila, 1, 1, sheetCot.getLastColumn()).getValues()[0];
       const rowFormulas = sheetCot.getRange(fila, 1, 1, sheetCot.getLastColumn()).getFormulas()[0];
       const rowRichTexts = sheetCot.getRange(fila, 1, 1, sheetCot.getLastColumn()).getRichTextValues()[0];
 
       const getValFromRow = (headerName) => {
         const idx = headersCot.indexOf(headerName.toLowerCase().trim());
         return idx !== -1 ? rowValues[idx] : "";
       };

       const getLinkFromRow = (headerName) => {
         const idx = headersCot.indexOf(headerName.toLowerCase().trim());
         if (idx === -1) return "";
         if (rowRichTexts && rowRichTexts[idx]) {
           const url = rowRichTexts[idx].getLinkUrl();
           if (url) return url;
         }
         if (rowFormulas && rowFormulas[idx]) {
           const formula = rowFormulas[idx];
           if (formula && formula.startsWith("=")) {
             if (formula.toUpperCase().indexOf("HYPERLINK") !== -1) {
               const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
               if (match) return match[1];
             }
           }
         }
         const val = String(rowValues[idx] || '');
         if (val.indexOf("http://") === 0 || val.indexOf("https://") === 0) {
           return val;
         }
         return "";
       };
 
       const cliente = getValFromRow("Cliente");
       const link = getLinkFromRow("Enlace Mercado Libre");
       const color = e.parameter.color || getValFromRow("Color") || "";
       const memoria = e.parameter.memoria || getValFromRow("Capacidad") || "";
 
       // Buscar modelo en Catalogo usando el link (o usar el modelo editado enviado)
       let modelo = e.parameter.modelo || "";
       const sheetCat = ss.getSheetByName('Catalogo');
       if (!modelo && sheetCat && link) {
         const catRange = sheetCat.getDataRange();
         const catData = catRange.getValues();
         const catFormulas = catRange.getFormulas();
         const catRichTexts = catRange.getRichTextValues();
         let targetMlm = "";
         const match = link.match(/(MLM\-?\d+)/i);
         if (match) targetMlm = match[0].replace("-", "").toUpperCase();
 
         for (let j = 1; j < catData.length; j++) {
           const catLink = extraerLinkDeArrays(catData, catFormulas, catRichTexts, j, 6);
           let catMlm = "";
           if (catLink) {
             const match2 = catLink.match(/(MLM\-?\d+)/i);
             if (match2) catMlm = match2[0].replace("-", "").toUpperCase();
           }
           if ((targetMlm && catMlm && targetMlm === catMlm) || link.trim() === catLink.trim()) {
             modelo = catData[j][1] || "";
             break;
           }
         }
       }
 
       if (!modelo) {
         modelo = "Celular Cotizado";
       }
 
       // Parámetros editables enviados desde el modal
       const imeiNuevo = e.parameter.imei || '';
       const estadoNuevo = e.parameter.estado || 'SIENDO PAGADO';
       const notasNuevo = e.parameter.notas || '';
       const precioFinal = parseFloat(e.parameter.precioFinal) || 0;
 
       sheetInv.appendRow([
         imeiNuevo,        // A (1): IMEI
         modelo,           // B (2): Modelo
         color,            // C (3): Color
         memoria,          // D (4): Memoria
         estadoNuevo,      // E (5): Estado
         notasNuevo,       // F (6): Notas
         precioFinal       // G (7): Precio Final
       ]);
 
        // Marcar la cotización como Asignada en la hoja Cotizaciones
        const headersCotConv = sheetCot.getRange(1, 1, 1, Math.max(25, sheetCot.getLastColumn())).getValues()[0].map(h => String(h || '').toLowerCase().trim());
        let colEstCot = headersCotConv.indexOf("estado de cotización") + 1;
        if (colEstCot === 0) {
          colEstCot = headersCotConv.indexOf("estado de cotizacion") + 1;
        }
        if (colEstCot === 0) {
          // Crear la columna si no existe
          colEstCot = sheetCot.getLastColumn() + 1;
          sheetCot.getRange(1, colEstCot).setValue("Estado de Cotización");
        }
        sheetCot.getRange(fila, colEstCot).setValue("Asignada");
        SpreadsheetApp.flush();
 
       return ContentService.createTextOutput(JSON.stringify({ result: 'success', message: 'Cotización convertida a Inventario' })).setMimeType(ContentService.MimeType.JSON);
     }

    // REGISTRO DE CONTRATO ORIGINAL (Venta Nueva)
    const sheetClientes = ss.getSheetByName('Clientes');
    if (!sheetClientes) throw new Error("No se encontró la pestaña 'Clientes'");
    
    const cliente = e.parameter.cliente ? e.parameter.cliente : "SinNombre";
    const clienteMayusculas = cliente.toUpperCase();
    const direccionMayusculas = e.parameter.direccion ? e.parameter.direccion.toString().toUpperCase() : "";
    const modelo = e.parameter.modelo ? e.parameter.modelo : "SinModelo";
    const imei = e.parameter.imei ? e.parameter.imei : "SinIMEI";

    // --- CANDADO DE VALIDACIÓN DE IMEI ÚNICO ---
    const imeiNuevo = e.parameter.imei ? e.parameter.imei.toString().trim() : "";

    if (imeiNuevo !== "" && imeiNuevo !== "SinIMEI") {
      const dataClientesValidar = sheetClientes.getDataRange().getValues();
      // Columna F (Índice 5) es donde guardas el IMEI en tu appendRow histórico
      for (let i = 1; i < dataClientesValidar.length; i++) {
        if (dataClientesValidar[i][5] && dataClientesValidar[i][5].toString().trim() === imeiNuevo) {
          const clienteComprador = dataClientesValidar[i][1] ? dataClientesValidar[i][1].toString().trim().toUpperCase() : "DESCONOCIDO";
          return ContentService.createTextOutput(JSON.stringify({
            result: "error", 
            error: "🚫 EL IMEI YA FUE VENDIDO A " + clienteComprador + ". Este equipo cuenta con un contrato activo."
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    // --------------------------------------------
    const timestamp = new Date();
    let firmaUrlLarga = "";
    let firmaLinkAutocrat = "";
    
    // Obtener o Generar el ID Único con formato C-Y000
    let nuevoIdCliente = e.parameter.idCliente ? e.parameter.idCliente.toString().trim() : "";
    if (!nuevoIdCliente.startsWith("C-Y")) {
      nuevoIdCliente = generarSiguienteIdCelYa(sheetClientes);
    }

    const mensajeTelegram = `👤 Cliente: ${cliente}\n📱 Modelo: ${modelo}\n🆔 IMEI: ${imei}\n🆔 ID Cliente: ${nuevoIdCliente}`;
    
    try {
      if (e.parameter.firmaBase64 && e.parameter.firmaBase64 !== "") {
        const base64Data = e.parameter.firmaBase64.split(',')[1];
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', `${cliente}_Firma.png`);
        const file = DriveApp.getFolderById(credenciales.FOLDER_ID).createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        firmaUrlLarga = file.getUrl();
        // 🛠️ FIX DE FIRMA: Corregido para forzar la descarga e impresión transparente en Autocrat
        firmaLinkAutocrat = "https://drive.google.com/uc?export=download&id=" + file.getId();
      }
    } catch(errFirma) {}

    try {
      if (e.parameter.ineDelanteraBase64 && e.parameter.ineDelanteraBase64.includes(",")) {
        const base64Data = e.parameter.ineDelanteraBase64.split(',')[1];
        enviarBlobATelegram(credenciales, Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', `${cliente}_INE_Frontal.jpg`), `🪪 INE FRONTAL\n\n${mensajeTelegram}`);
      }
      if (e.parameter.ineTraseraBase64 && e.parameter.ineTraseraBase64.includes(",")) {
        const base64Data = e.parameter.ineTraseraBase64.split(',')[1];
        enviarBlobATelegram(credenciales, Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', `${cliente}_INE_Reverso.jpg`), `🪪 INE REVERSO\n\n${mensajeTelegram}`);
      }
    } catch(errTel) {}

    try {
      if (e.parameter.fotoEvidenciaBase64 && e.parameter.fotoEvidenciaBase64.includes(",")) {
        const base64DataEv = e.parameter.fotoEvidenciaBase64.split(',')[1];
        const blobEv = Utilities.newBlob(Utilities.base64Decode(base64DataEv), 'image/jpeg', `${cliente}_Evidencia.jpg`);
        DriveApp.getFolderById(credenciales.FOLDER_ID).createFile(blobEv);
        enviarBlobATelegram(credenciales, blobEv, `📸 CELULAR ENTREGADO ENCENDIDO\n\n👤 Cliente: ${cliente}\n📱 Modelo: ${modelo}`);
      }
    } catch(errEv) {}

    const formatearMoneda = (valor) => {
      if (!valor || valor.toString().trim() === "") return "";
      let numero = valor.toString().replace(/[\$,]/g, "").trim();
      if (isNaN(numero)) return valor; 
      return "$" + parseFloat(numero).toLocaleString('es-MX', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    };
    const formatearPorcentaje = (valor) => {
      if (!valor || valor.toString().trim() === "") return "";
      let numero = valor.toString().replace(/%/g, "").trim();
      if (isNaN(numero)) return valor;
      return numero + "%";
    };

    const proximaFila = sheetClientes.getLastRow() + 1;
    const formulaVencimiento = `=IF(OR(A${proximaFila}="", O${proximaFila}=""), "", A${proximaFila} + (VALUE(REGEXEXTRACT(O${proximaFila}, "[0-9]+")) * 7))`;
    
    sheetClientes.appendRow([
      timestamp,                                                         // A
      clienteMayusculas,                                                 // B
      e.parameter.telefono ? e.parameter.telefono : "",                 // C
      e.parameter.fechaNacimiento ? e.parameter.fechaNacimiento : "",    // D
      modelo,                                                            // E
      imeiNuevo,                                                         // F
      e.parameter.monto ? formatearMoneda(e.parameter.monto) : "",      // G
      direccionMayusculas,                                               // H
      e.parameter.telefonoAval ? e.parameter.telefonoAval : "",          // I
      firmaLinkAutocrat ? `=HYPERLINK("${firmaLinkAutocrat}", "${firmaLinkAutocrat}")` : "", // J ✍️ ¡AQUÍ COLOCAMOS EL FORMATO DE AUTOCRAT (uc?export=download)!
      formatearMoneda(e.parameter.cuota ? Math.round(parseFloat(e.parameter.cuota.toString().replace(/[\$,]/g, "").trim()) || 0) : ""), // K
      formatearPorcentaje(e.parameter.interes),                          // L
      e.parameter.horastol ? e.parameter.horastol + " hrs" : "",         // M
      e.parameter.semanastol ? e.parameter.semanastol + " sem" : "",     // N
      (function() {
        const numSemanas = parseInt(e.parameter.plazoSemanas) || 26;
        const frecuencia = e.parameter.tipoPeriodo ? e.parameter.tipoPeriodo.toLowerCase() : "semanal";
        if (frecuencia === "quincenal") {
          const pagosQuincenales = numSemanas / 2;
          return pagosQuincenales + " PAGOS QUINCENALES";
        } else if (frecuencia === "contado") {
          return "1 PAGO DE CONTADO";
        } else {
          return numSemanas + " PAGOS SEMANALES";
        }
      })(), // O
      formatearMoneda(e.parameter.totalFinanciado),                      // P
      (function() {
        const totalFin = e.parameter.totalFinanciado ? parseFloat(e.parameter.totalFinanciado.toString().replace(/[\$,]/g, "").trim()) || 0 : 0;
        return numeroALetras(totalFin);
      })(),                                                               // Q
      formulaVencimiento,                                                // R
      e.parameter.diaRaya ? e.parameter.diaRaya.toUpperCase() : "",      // S 💰 Columna 19: Día de Pago
      nuevoIdCliente,                                                     // T 🆔 Columna 20: ID de Cliente (C-Y000)
      e.parameter.tipoPeriodo ? e.parameter.tipoPeriodo.toUpperCase() : "SEMANAL" // U 📆 Columna 21: Tipo de Periodo
    ]);
    
    // --- LLAMADA AL WEBHOOK DE n8n PARA GENERAR CONTRATO PDF Y ENVIAR EMAIL ---
    try {
      let n8nWebhookUrl = credenciales.N8N_CONTRATO_WEBHOOK || PropertiesService.getScriptProperties().getProperty("N8N_CONTRATO_WEBHOOK");
      if (n8nWebhookUrl === "https://n8n.estrenacelya.com/webhook/contrato-pagare" || n8nWebhookUrl === "https://n8n.estrenacelya.com/webhook/3/webhook/contrato-pagare" || (n8nWebhookUrl && !n8nWebhookUrl.includes("/3/webhook/"))) {
        n8nWebhookUrl = "http://107.175.122.33:5678/webhook/3/webhook/contrato-pagare";
        PropertiesService.getScriptProperties().setProperty("N8N_CONTRATO_WEBHOOK", n8nWebhookUrl);
      }
      if (n8nWebhookUrl) {
        const numSemanasVal = parseInt(e.parameter.plazoSemanas) || 26;
        const frecuenciaVal = e.parameter.tipoPeriodo ? e.parameter.tipoPeriodo.toLowerCase() : "semanal";
        let diasTotales = numSemanasVal * 7;
        if (frecuenciaVal === "quincenal") {
          diasTotales = (numSemanasVal / 2) * 15;
        } else if (frecuenciaVal === "contado") {
          diasTotales = 0;
        }
        const fechaVencimientoObj = new Date(timestamp.getTime() + diasTotales * 24 * 60 * 60 * 1000);
        const fechaVencimientoStr = Utilities.formatDate(fechaVencimientoObj, "America/Mexico_City", "dd/MM/yyyy");
        const fechaEntregaStr = Utilities.formatDate(timestamp, "America/Mexico_City", "dd/MM/yyyy");

        const payloadWebhook = {
          timestamp: Utilities.formatDate(timestamp, "America/Mexico_City", "yyyy-MM-dd HH:mm:ss"),
          fechaEntrega: fechaEntregaStr,
          fechaVencimiento: fechaVencimientoStr,
          cliente: clienteMayusculas,
          telefono: e.parameter.telefono ? e.parameter.telefono : "",
          fechaNacimiento: e.parameter.fechaNacimiento ? e.parameter.fechaNacimiento : "",
          modelo: modelo,
          imei: imeiNuevo,
          monto: e.parameter.monto ? formatearMoneda(e.parameter.monto) : "",
          direccion: direccionMayusculas,
          telefonoAval: e.parameter.telefonoAval ? e.parameter.telefonoAval : "",
          firmaLinkAutocrat: firmaLinkAutocrat,
          firmaBase64: e.parameter.firmaBase64 || "",
          cuota: e.parameter.cuota ? Math.round(parseFloat(e.parameter.cuota.toString().replace(/[\$,]/g, "").trim()) || 0) : "",
          interes: e.parameter.interes ? e.parameter.interes : "",
          horastol: e.parameter.horastol ? e.parameter.horastol : "",
          semanastol: e.parameter.semanastol ? e.parameter.semanastol : "",
          plazoSemanas: numSemanasVal,
          plazoText: (function() {
            if (frecuenciaVal === "quincenal") {
              const pagosQuincenales = numSemanasVal / 2;
              return pagosQuincenales + " PAGOS QUINCENALES";
            } else if (frecuenciaVal === "contado") {
              return "1 PAGO DE CONTADO";
            } else {
              return numSemanasVal + " PAGOS SEMANALES";
            }
          })(),
          totalFinanciado: e.parameter.totalFinanciado ? formatearMoneda(e.parameter.totalFinanciado) : "",
          totalFinanciadoLetras: (function() {
            const totalFin = e.parameter.totalFinanciado ? parseFloat(e.parameter.totalFinanciado.toString().replace(/[\$,]/g, "").trim()) || 0 : 0;
            return numeroALetras(totalFin);
          })(),
          diaRaya: e.parameter.diaRaya ? e.parameter.diaRaya.toUpperCase() : "",
          idCliente: nuevoIdCliente,
          tipoPeriodo: e.parameter.tipoPeriodo ? e.parameter.tipoPeriodo.toUpperCase() : "SEMANAL"
        };
        
        const response = UrlFetchApp.fetch(n8nWebhookUrl, {
          method: "POST",
          contentType: "application/json",
          payload: JSON.stringify(payloadWebhook),
          muteHttpExceptions: true
        });
        
        const responseCode = response.getResponseCode();
        if (responseCode === 200) {
          const contentType = response.getHeaders()["Content-Type"] || response.getHeaders()["content-type"] || "";
          const blobContentType = response.getBlob().getContentType() || "";
          
          if (contentType.toLowerCase().indexOf("pdf") !== -1 || blobContentType.toLowerCase().indexOf("pdf") !== -1) {
            const pdfBlob = response.getBlob().setName("Contrato_y_Pagare_" + clienteMayusculas + "_" + imeiNuevo + ".pdf");
            
            // 1. Guardar en Google Drive (Carpeta especificada por el usuario)
            try {
              const folder = DriveApp.getFolderById("1PoPjWmHMePfW8VsMn6D7ORvGCOYZFrWY");
              folder.createFile(pdfBlob);
            } catch (errDrive) {
              console.error("Error guardando PDF en Google Drive: " + errDrive.toString());
              escribirLogDebug("Error guardando PDF en Drive: " + errDrive.toString());
            }
            
            // 2. Enviar por correo
            try {
              const subject = "Contrato y Pagare para " + clienteMayusculas;
              const body = "Se adjunta el Contrato y Pagaré digital.\n\n" +
                           "Cliente: " + clienteMayusculas + "\n" +
                           "Modelo: " + modelo + "\n" +
                           "IMEI: " + imeiNuevo + "\n" +
                           "ID Cliente: " + nuevoIdCliente;
              
              GmailApp.sendEmail("celyamex@gmail.com", subject, body, {
                attachments: [pdfBlob]
              });
            } catch (errEmail) {
              console.error("Error enviando email: " + errEmail.toString());
              escribirLogDebug("Error enviando correo del PDF: " + errEmail.toString());
            }
          } else {
            console.error("n8n no retornó un PDF. Respuesta: " + response.getContentText());
            escribirLogDebug("n8n no retornó un PDF: " + response.getContentText().substring(0, 200));
          }
        } else {
          console.error("Error en respuesta de n8n. Código: " + responseCode + " - Detalle: " + response.getContentText());
          escribirLogDebug("Error en respuesta de n8n: " + responseCode);
        }
      }
    } catch (errN8n) {
      console.error("Error enviando webhook a n8n: " + errN8n.toString());
      escribirLogDebug("Error enviando webhook a n8n: " + errN8n.toString());
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({result: "success", idCliente: nuevoIdCliente})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Error en doPost: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      result: "error", 
      message: "Error en Script: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function enviarBlobATelegram(credenciales, blob, caption) {
  const url = `https://api.telegram.org/bot${credenciales.TELEGRAM_TOKEN}/sendPhoto`;
  UrlFetchApp.fetch(url, { 'method': 'post', 'payload': { 'chat_id': credenciales.CHAT_ID, 'photo': blob, 'caption': caption }, 'muteHttpExceptions': true });
}

// 🤖 ENVÍA MENSAJES DE TEXTO A TELEGRAM DE FORMA SEGURA Y VALIDANDO CREDENCIALES
function enviarMensajeTelegram(arg1, arg2, arg3) {
  // Soporta dos firmas:
  //   enviarMensajeTelegram(credenciales, mensaje)        — 2 args
  //   enviarMensajeTelegram(botToken, chatId, mensaje)    — 3 args
  let botToken, chatId, mensaje;
  
  if (arg3 !== undefined) {
    // Llamada con 3 args: (botToken, chatId, mensaje)
    botToken = arg1;
    chatId = arg2;
    mensaje = arg3;
  } else if (typeof arg1 === 'object' && arg1 !== null) {
    // Llamada con 2 args: (credenciales, mensaje)
    botToken = arg1.TELEGRAM_TOKEN;
    chatId = arg1.CHAT_ID;
    mensaje = arg2;
  } else {
    // Llamada con 2 args donde arg1 es botToken string (fallback)
    botToken = arg1;
    chatId = arg2;
    mensaje = '';
    Logger.log("enviarMensajeTelegram: firma no reconocida");
    return false;
  }

  if (!botToken || !chatId) {
    Logger.log("TELEGRAM_TOKEN o CHAT_ID no están configurados.");
    return false;
  }
  
  const url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  const payload = {
    chat_id: chatId,
    text: mensaje,
    parse_mode: "HTML"
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log("Error Telegram. Código: " + responseCode + ", Respuesta: " + response.getContentText());
      return false;
    }
    Logger.log("Telegram Response OK: " + response.getContentText());
    return true;
  } catch (err) {
    Logger.log("Error de conexión Telegram: " + err.toString());
    return false;
  }
}

// 🔠 ESCAPA CARACTERES ESPECIALES PARA HTML DE TELEGRAM (parse_mode=HTML)
function escapeMarkdownV1(text) {
  // Alias para compatibilidad — ahora usa escape HTML
  return escapeHTML(text);
}

function escapeHTML(text) {
  if (!text && text !== 0) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// 📱 EXTRAE EL NOMBRE DEL CELULAR DESDE EL PLAZO
function extraerNombreCelular(plazo, color, capacidad) {
  if (!plazo) return "Celular";
  const match = plazo.match(/\(([^)]+)\)$/);
  if (match) {
    let name = match[1];
    name = name.replace(/^PEDIDO ESPECIAL:\s*/i, "");
    if (color) {
      name = name.replace(new RegExp('\\s*-\\s*' + escapeRegex(color), 'gi'), '');
    }
    if (capacidad) {
      name = name.replace(new RegExp('\\s*-\\s*' + escapeRegex(capacidad), 'gi'), '');
    }
    name = name.replace(/[-\s,]+$/, "").trim();
    if (name) return name;
  }
  return "Celular";
}

// 🔍 ESCAPA CARACTERES PARA EXPRESIONES REGULARES
function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

// 🧼 SANITIZA EL NOMBRE DEL CELULAR PARA EVITAR CARACTERES CONTROLADORES DE MARKDOWN
function sanitizarNombreCelular(nombre) {
  if (!nombre) return "Celular";
  let clean = nombre.replace(/[_*`\[\]()]/g, "");
  clean = clean.replace(/\s+/g, " ").trim();
  return clean || "Celular";
}

// 🧪 FUNCIÓN DE PRUEBA PARA VALIDAR EL BOT DE TELEGRAM
function probarTelegram() {
  const credenciales = obtenerPropiedadesEcosistema();
  Logger.log("TELEGRAM_TOKEN: " + (credenciales.TELEGRAM_TOKEN ? "Configurado (comienza con " + credenciales.TELEGRAM_TOKEN.substring(0, 5) + "...)" : "No configurado"));
  Logger.log("CHAT_ID: " + (credenciales.CHAT_ID ? "Configurado (" + credenciales.CHAT_ID + ")" : "No configurado"));
  
  if (credenciales.TELEGRAM_TOKEN && credenciales.CHAT_ID) {
    const exito = enviarMensajeTelegram(credenciales, "🤖 *Prueba de Bot de Telegram* 🤖\n\nLas propiedades están activas y el bot funciona correctamente.");
    Logger.log("Resultado del envío de prueba: " + (exito ? "ÉXITO" : "FALLÓ"));
  } else {
    Logger.log("No se puede probar el bot porque faltan credenciales.");
  }
}

// 🏷️ VERIFICA SI EL MODELO ES DE UNA MARCA AUTORIZADA POR CELYA
function esMarcaAutorizada(modelo) {
  if (!modelo) return false;
  const modeloUpper = String(modelo).toUpperCase();
  return modeloUpper.includes("SAMSUNG") || 
         modeloUpper.includes("MOTOROLA") || 
         modeloUpper.includes("MOTO") || 
         modeloUpper.includes("XIAOMI") || 
         modeloUpper.includes("REDMI") || 
         modeloUpper.includes("POCO") || 
         modeloUpper.includes("OPPO") ||
         modeloUpper.includes("CELULAR COTIZADO") ||
         modeloUpper.includes("IPHONE") ||
         modeloUpper.includes("APPLE");
}

// 🔤 CONVIERTE UN NÚMERO ENTERO A SU REPRESENTACIÓN EN TEXTO EN ESPAÑOL
function numeroALetras(num) {
  const unidades = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
  const decenas = ["", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
  const especiales = {
    11: "ONCE", 12: "DOCE", 13: "TRECE", 14: "CATORCE", 15: "QUINCE",
    16: "DIECISEIS", 17: "DIECISIETE", 18: "DIECIOCHO", 19: "DIECINUEVE",
    21: "VEINTIUN", 22: "VEINTIDOS", 23: "VEINTITRES", 24: "VEINTICUATRO",
    25: "VEINTICINCO", 26: "VEINTISEIS", 27: "VEINTISIETE", 28: "VEINTIOCHO",
    29: "VEINTINUEVE"
  };
  const centenas = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];

  function convertirGrupo(n) {
    if (n === 0) return "";
    let output = "";
    const c = Math.floor(n / 100);
    const d = Math.floor((n % 100) / 10);
    const u = n % 10;

    if (c > 0) {
      if (c === 1 && d === 0 && u === 0) {
        output += "CIEN ";
      } else {
        output += centenas[c] + " ";
      }
    }

    const du = d * 10 + u;
    if (du > 0) {
      if (especiales[du]) {
        output += especiales[du] + " ";
      } else {
        if (d > 0) {
          if (d === 2 && u > 0) {
            output += "VEINTI" + unidades[u] + " ";
          } else {
            output += decenas[d] + (u > 0 ? " Y " + unidades[u] : "") + " ";
          }
        } else {
          output += unidades[u] + " ";
        }
      }
    }
    return output.trim();
  }

  num = Math.round(num);
  if (num === 0) return "CERO PESOS 00/100 M.N.";
  if (num === 1) return "UN PESO 00/100 M.N.";

  let result = "";
  const millones = Math.floor(num / 1000000);
  const miles = Math.floor((num % 1000000) / 1000);
  const resto = num % 1000;

  if (millones > 0) {
    if (millones === 1) {
      result += "UN MILLON ";
    } else {
      result += convertirGrupo(millones) + " MILLONES ";
    }
  }

  if (miles > 0) {
    if (miles === 1) {
      result += "MIL ";
    } else {
      result += convertirGrupo(miles) + " MIL ";
    }
  }

  if (resto > 0) {
    result += convertirGrupo(resto) + " ";
  }

  result = result.trim();
  
  if (millones > 0 && miles === 0 && resto === 0) {
    return result + " DE PESOS 00/100 M.N.";
  }
  return result + " PESOS 00/100 M.N.";
}

/**
 * Convierte un número a letras en pesos mexicanos.
 * @param {number} numero El número a convertir.
 * @return {string} El número en mayúsculas expresado en letras.
 * @customfunction
 */
function NUMERO_A_LETRAS(numero) {
  if (numero === null || numero === undefined || isNaN(numero) || numero === "") {
    return "";
  }
  return numeroALetras(numero);
}

// 🆔 GENERADOR OFICIAL DE ID CORRELATIVO C-Y000
function generarSiguienteIdCelYa(sheetClientes) {
  const totalFilas = sheetClientes.getLastRow();
  
  // Si la tabla está vacía o solo tiene la fila 1 de encabezados, el primero es el 001
  if (totalFilas <= 1) return "C-Y001";
  
  // Leemos toda la columna T (columna 20) a partir de la fila 2
  const valoresId = sheetClientes.getRange(2, 20, totalFilas - 1, 1).getValues();
  let ultimoNumeroId = 0;
  
  // Buscamos de abajo hacia arriba si ya existe algún ID real guardado
  for (let i = valoresId.length - 1; i >= 0; i--) {
    const celda = valoresId[i][0] ? valoresId[i][0].toString().trim() : "";
    if (celda.startsWith("C-Y")) {
      // Extraemos el número (ejemplo: "C-Y003" -> 3)
      const numeroExtraido = parseInt(celda.replace("C-Y", ""), 10);
      if (!isNaN(numeroExtraido)) {
        ultimoNumeroId = numeroExtraido;
        break; // Encontramos el último ID válido consecutivo
      }
    }
  }
  
  // Si no se encontró ningún ID con formato "C-Y", significa que está limpia; iniciamos en 1
  if (ultimoNumeroId === 0) return "C-Y001";
  
  // Si ya había, le sumamos 1 al último número real de la base de datos
  return "C-Y" + String(ultimoNumeroId + 1).padStart(3, '0');
}

function probarCatalogoLocal() {
  const credenciales = obtenerPropiedadesEcosistema();
  escribirLogDebug("🔑 SHEET_ID configurado: " + credenciales.SHEET_ID);
  if (!credenciales.SHEET_ID) {
    escribirLogDebug("❌ ERROR: Falta SHEET_ID");
    return;
  }
  const ss = SpreadsheetApp.openById(credenciales.SHEET_ID);
  const sheetInv = ss.getSheetByName('Catalogo');
  const sheetConfig = ss.getSheetByName('Configuracion') || ss.getSheetByName('Configuración');
  
  if (!sheetInv || !sheetConfig) {
    escribirLogDebug("❌ ERROR: No se encontró la pestaña Catalogo o Configuracion");
    return;
  }
  
  const ultFila = sheetInv.getLastRow();
  escribirLogDebug("📊 Última fila de Catálogo: " + ultFila);
  if (ultFila <= 1) {
    escribirLogDebug("⚠️ No hay datos en Catálogo.");
    return;
  }
  
  const isCatalogo = (sheetInv.getName() === 'Catalogo');
  const minColsNeeded = isCatalogo ? 12 : 7;
  const maxCols = sheetInv.getLastColumn();
  if (maxCols < minColsNeeded) {
    sheetInv.insertColumnsAfter(maxCols, minColsNeeded - maxCols);
  }
  const totalCols = sheetInv.getLastColumn();
  const datosInv = sheetInv.getRange(2, 1, ultFila - 1, totalCols).getValues();
  
  let linksRicos = [];
  let formulasLinks = [];
  if (isCatalogo) {
    linksRicos = sheetInv.getRange(2, 7, ultFila - 1, 2).getRichTextValues();
    formulasLinks = sheetInv.getRange(2, 7, ultFila - 1, 2).getFormulas();
  }
  escribirLogDebug("📊 Filas leídas: " + datosInv.length);
  
  let peticionesRealizadas = 0;
  const MAX_PETICIONES = 3;
  
  datosInv.forEach((fila, index) => {
    const numeroFilaReal = index + 2;
    
    let imei = "";
    if (typeof fila[0] === "number") {
      imei = fila[0].toFixed(0);
    } else if (fila[0] !== null && fila[0] !== undefined) {
      imei = fila[0].toString().trim();
    }
    imei = imei.toUpperCase();

    let modeloTitulo = fila[1] ? fila[1].toString().trim() : "";
    let precioContadoManual = 0;
    let precioContadoFinalFromSheet = 0;
    let estado = "";
    let linkMercadoLibre = "";
    let cachedFoto = "";
    let cachedMemoria = "";
    let cachedColor = "";

    if (isCatalogo) {
      precioContadoManual = parseFloat(fila[3]) || 0;
      precioContadoFinalFromSheet = (fila.length > 11 && fila[11] !== undefined && fila[11] !== null && String(fila[11]).trim() !== "") ? parseFloat(fila[11]) || 0 : 0;
      estado = fila[5] ? fila[5].toString().toUpperCase().trim() : "";
      
      const formula = formulasLinks[index] && formulasLinks[index][0] ? formulasLinks[index][0] : "";
      let formulaLink = "";
      if (formula.startsWith("=HYPERLINK")) {
        const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
        if (match) formulaLink = match[1];
      }
      const richLink = linksRicos[index] && linksRicos[index][0] ? linksRicos[index][0].getLinkUrl() : "";
      linkMercadoLibre = richLink || formulaLink || (fila[6] ? fila[6].toString().trim() : "");

      const formulaFoto = formulasLinks[index] && formulasLinks[index][1] ? formulasLinks[index][1] : "";
      let formulaFotoLink = "";
      if (formulaFoto.startsWith("=HYPERLINK")) {
        const match = formulaFoto.match(/HYPERLINK\("([^"]+)"/i) || formulaFoto.match(/HYPERLINK\('([^']+)'/i);
        if (match) formulaFotoLink = match[1];
      }
      const richFotoLink = linksRicos[index] && linksRicos[index][1] ? linksRicos[index][1].getLinkUrl() : "";
      cachedFoto = richFotoLink || formulaFotoLink || (fila[7] ? fila[7].toString().trim() : "");
      if (cachedFoto.toLowerCase().trim() === "foto") cachedFoto = "";

      cachedMemoria = fila[8] ? fila[8].toString().trim() : "";
      cachedColor = fila[9] ? fila[9].toString().trim() : "";
    } else {
      precioContadoManual = parseFloat(fila[6]) || 0;
      estado = fila[4] ? fila[4].toString().toUpperCase().trim() : "";
      cachedMemoria = fila[3] ? fila[3].toString().trim() : "";
      cachedColor = fila[2] ? fila[2].toString().trim() : "";
    }
    
    escribirLogDebug(`\n--- Fila ${numeroFilaReal} ---`);
    escribirLogDebug(`IMEI: ${imei} | Modelo: ${modeloTitulo} | Precio: ${precioContadoManual} | Estado: ${estado} | Link: ${linkMercadoLibre}`);
    
    const cumpleEstado = (estado === "DISPONIBLE" || estado === "MAS VENDIDO" || estado === "BAJO DEMANDA" || imei.startsWith("TOP") || imei.startsWith("999999"));
    escribirLogDebug("¿Cumple Estado?: " + cumpleEstado);
    
    if (cumpleEstado) {
      if ((precioContadoManual === 0 || cachedFoto === "") && linkMercadoLibre.includes("mercadolibre")) {
        if (peticionesRealizadas < MAX_PETICIONES) {
          escribirLogDebug("👉 Intentando consultar Mercado Libre...");
          const res = obtenerDatosDeMercadoLibre(linkMercadoLibre);
          if (res) {
            if (precioContadoManual === 0) {
              precioContadoManual = Math.round(res.price);
              escribirLogDebug("💉 Guardando en Sheets (Fila " + numeroFilaReal + "): Precio Costo = $" + precioContadoManual);
              sheetInv.getRange(numeroFilaReal, 4).setValue(precioContadoManual);
            }
            if (modeloTitulo === "") {
              modeloTitulo = res.title;
              sheetInv.getRange(numeroFilaReal, 2).setValue(modeloTitulo);
            }
            if (res.pictures && res.pictures.length > 0 && cachedFoto === "") {
              cachedFoto = res.pictures[0].secure_url;
              if (isCatalogo) {
                sheetInv.getRange(numeroFilaReal, 8).setFormula(`=HYPERLINK("${cachedFoto}", "foto")`);
              }
            }
            if (res.memoria && cachedMemoria === "") {
              cachedMemoria = res.memoria;
              if (isCatalogo) {
                sheetInv.getRange(numeroFilaReal, 9).setValue(cachedMemoria);
              }
            }
            if (res.color && cachedColor === "") {
              cachedColor = res.color;
              if (isCatalogo) {
                sheetInv.getRange(numeroFilaReal, 10).setValue(cachedColor);
              }
            }
            peticionesRealizadas++;
          }
        } else {
          escribirLogDebug("⏭️ Se omitió consulta por alcanzar el límite de peticiones");
        }
      } else {
        escribirLogDebug("⏭️ Se salta la consulta (Ya tiene precio y foto, o no tiene link ML)");
      }
    }
  });
  SpreadsheetApp.flush();
}

// 🎨 NORMALIZA EL COLOR A UN FORMATO SENCILLO Y LEGIBLE EN ESPAÑOL PARA EL CATÁLOGO
function normalizarColorEspanol(colorName) {
  if (!colorName) return "";
  const name = colorName.toLowerCase().trim();

  // Mapeos específicos de nombres comerciales a básicos
  const mapeosEspecificos = {
    "matte charcoal": "Negro",
    "phantom black": "Negro",
    "midnight": "Negro",
    "charcoal": "Negro",
    "graphite": "Negro",
    "grafito": "Negro",
    "awesome graphite": "Negro",
    "space gray": "Gris",
    "space grey": "Gris",
    "gris oscuro": "Gris",
    "gris espacial": "Gris",
    "awesome iceblue": "Azul",
    "awesome lemon": "Amarillo",
    "lime green": "Verde",
    "awesome lime": "Verde",
    "mint": "Verde",
    "menta": "Verde",
    "mint green": "Verde",
    "starlight": "Blanco",
    "silver": "Plata",
    "plata": "Plata"
  };

  if (mapeosEspecificos[name]) {
    return mapeosEspecificos[name];
  }

  // Grupos genéricos por búsqueda de palabras clave
  const groups = [
    { canonical: "Negro", keys: ["negro", "black", "space", "carbono", "carbon", "oscur"] },
    { canonical: "Blanco", keys: ["blanco", "white", "claro", "light"] },
    { canonical: "Plata", keys: ["plata", "silver"] },
    { canonical: "Gris", keys: ["gris", "grey", "gray"] },
    { canonical: "Crema", keys: ["cream", "crema", "marfil", "ivory", "beige"] },
    { canonical: "Azul", keys: ["azul", "blue", "celeste", "sky", "indigo", "marino"] },
    { canonical: "Verde", keys: ["verde", "green", "lima", "lime", "menta", "mint", "musgo"] },
    { canonical: "Rojo", keys: ["rojo", "red", "naranja", "orange", "coral", "atardecer"] },
    { canonical: "Rosa", keys: ["rosa", "pink", "gold", "oro", "dorado", "bronce", "bronze"] },
    { canonical: "Violeta", keys: ["violeta", "violet", "purpura", "púrpura", "purple", "morado", "lavanda", "lavender", "lila", "lilac"] }
  ];

  // Buscamos primero en los colores base prioritarios (evitando falsos positivos de oscur/claro)
  const gruposPrioritarios = ["Gris", "Azul", "Verde", "Rojo", "Rosa", "Violeta", "Crema", "Plata"];
  for (let i = 0; i < gruposPrioritarios.length; i++) {
    const label = gruposPrioritarios[i];
    const grp = groups.find(function(g) { return g.canonical === label; });
    if (grp && grp.keys.some(function(key) { return name.indexOf(key) !== -1; })) {
      return grp.canonical;
    }
  }

  // Buscamos en Negro y Blanco
  for (let i = 0; i < groups.length; i++) {
    const grp = groups[i];
    if (gruposPrioritarios.indexOf(grp.canonical) === -1) {
      if (grp.keys.some(function(key) { return name.indexOf(key) !== -1; })) {
        return grp.canonical;
      }
    }
  }

  // Capitalizar la primera letra si no coincide
  return colorName.charAt(0).toUpperCase() + colorName.slice(1).toLowerCase();
}

// 🌐 OBTIENE TÍTULO, IMAGEN Y MENOR PRECIO DESDE LA API DE MERCADO LIBRE
function obtenerDatosDeMercadoLibre(linkMercadoLibre, targetCondition) {
  try {
    const regex = /(MLM\-?\d+)/i;
    const coincidencia = linkMercadoLibre.match(regex);
    if (!coincidencia) {
      escribirLogDebug("❌ No se encontró formato MLM en el link: " + linkMercadoLibre);
      return null;
    }
    
    const id = coincidencia[0].replace("-", "").toUpperCase();
    
    // Si es un link de producto de catálogo (/p/)
    if (linkMercadoLibre.includes("/p/")) {
      escribirLogDebug("📦 Es link de catálogo (/p/). Consultando API de producto para: " + id);
      
      // 1. Obtener detalles del producto (nombre, imagen)
      const detailUrl = `https://api.mercadolibre.com/products/${id}`;
      const resDetail = fetchConToken(detailUrl);
      escribirLogDebug("📊 Código detalles producto: " + resDetail.getResponseCode());
      
      if (resDetail.getResponseCode() !== 200) {
        escribirLogDebug("❌ Error al obtener detalles del catálogo: " + resDetail.getContentText());
        return null;
      }
      
      const productDetails = JSON.parse(resDetail.getContentText());
      
      // 2. Obtener los precios de los ítems del producto
      const itemsUrl = `https://api.mercadolibre.com/products/${id}/items`;
      const resItems = fetchConToken(itemsUrl);
      escribirLogDebug("📊 Código items de producto: " + resItems.getResponseCode());
      
      if (resItems.getResponseCode() !== 200) {
        escribirLogDebug("❌ Error al obtener items del catálogo: " + resItems.getContentText());
        return null;
      }
      
      const productItems = JSON.parse(resItems.getContentText());
      const results = productItems.results || [];
      
      let minPrice = Infinity;
      const targetCond = targetCondition || "new";
      const itemsFiltrados = results.filter(function(item) {
        if (targetCond === "used") {
          return item.condition === 'used' || item.condition === 'refurbished';
        }
        return item.condition === targetCond;
      });
      const itemsAProcesar = itemsFiltrados.length > 0 ? itemsFiltrados : results;
      
      itemsAProcesar.forEach(item => {
        if (item.price && item.price < minPrice) {
          minPrice = item.price;
        }
      });
      
      if (minPrice === Infinity) {
        escribirLogDebug("⚠️ No se encontraron precios válidos en los ítems de catálogo.");
        return null;
      }
      
      const thumbnail = productDetails.pictures && productDetails.pictures.length > 0
        ? productDetails.pictures[0].url
        : "";
        
      let color = "";
      let memoria = "";
      let colors = [];
      let memories = [];
      
      if (productDetails.pickers && Array.isArray(productDetails.pickers)) {
        productDetails.pickers.forEach(picker => {
          const pIdUpper = picker.picker_id ? picker.picker_id.toUpperCase() : "";
          const isColor = pIdUpper.includes("COLOR") || pIdUpper === "FINISH" || pIdUpper === "ACABADO";
          const isMemory = pIdUpper.includes("MEMORY") || pIdUpper.includes("STORAGE") || pIdUpper.includes("CAPACITY") || pIdUpper.includes("CAPACIDAD");
          
          if (picker.products && Array.isArray(picker.products)) {
            picker.products.forEach(p => {
              const isOutOfStock = p.tags && (p.tags.indexOf("out-of-stock") !== -1 || p.tags.indexOf("disabled") !== -1);
              if (!isOutOfStock && p.picker_label) {
                const label = p.picker_label.trim();
                if (isColor && colors.indexOf(label) === -1) colors.push(label);
                if (isMemory && memories.indexOf(label) === -1) memories.push(label);
              }
            });
          }
        });
      }
      
      if (colors.length === 0 && productDetails.variations && productDetails.variations.length > 0) {
        productDetails.variations.forEach(v => {
          const isAvailable = v.available_quantity !== undefined ? v.available_quantity > 0 : true;
          if (isAvailable) {
            const combinations = v.attribute_combinations || [];
            const attrs = v.attributes || [];
            const allAttrs = combinations.concat(attrs);
            const attrColor = allAttrs.find(a => a.id === 'COLOR' || a.id === 'MAIN_COLOR');
            const attrStorage = allAttrs.find(a => a.id === 'INTERNAL_MEMORY' || a.id === 'MEMORY_CAPACITY');
            if (attrColor && attrColor.value_name && colors.indexOf(attrColor.value_name.trim()) === -1) {
              colors.push(attrColor.value_name.trim());
            }
            if (attrStorage && attrStorage.value_name && memories.indexOf(attrStorage.value_name.trim()) === -1) {
              memories.push(attrStorage.value_name.trim());
            }
          }
        });
      }
      
      if (colors.length === 0) {
        const rootColor = extraerAtributo(productDetails.attributes, "COLOR") || extraerAtributo(productDetails.attributes, "MAIN_COLOR");
        if (rootColor) colors.push(rootColor);
      }
      if (memories.length === 0) {
        const rootMemoria = extraerAtributo(productDetails.attributes, "INTERNAL_MEMORY") || extraerAtributo(productDetails.attributes, "MEMORY_CAPACITY");
        if (rootMemoria) memories.push(rootMemoria);
      }
      
      color = colors.filter(Boolean).map(normalizarColorEspanol).filter(function(v, i, a) { return v && a.indexOf(v) === i; }).join(", ");
      memoria = memories.filter(Boolean).join(", ");
        
      escribirLogDebug("✅ Sincronizado desde catálogo: " + productDetails.name + " - Precio: $" + minPrice + " - Memoria: " + memoria + " - Color: " + color);
      return {
        title: productDetails.name,
        price: minPrice,
        pictures: thumbnail ? [{ secure_url: thumbnail.replace("-I.jpg", "-O.jpg").replace("http://", "https://") }] : [],
        color: color,
        memoria: memoria
      };
    } else {
      // Es una publicación individual (/items/)
      escribirLogDebug("🛍️ Es link de publicación individual (/items/). Consultando API de item para: " + id);
      
      const itemUrl = `https://api.mercadolibre.com/items/${id}`;
      const resItem = fetchConToken(itemUrl);
      escribirLogDebug("📊 Código item de publicación: " + resItem.getResponseCode());
      
      if (resItem.getResponseCode() !== 200) {
        escribirLogDebug("❌ Error al obtener detalles del item: " + resItem.getContentText());
        return null;
      }
      
      const itemDetails = JSON.parse(resItem.getContentText());
      
      // Auto-resolution to catalog product if present
      if (itemDetails.catalog_product_id) {
        escribirLogDebug("📦 Publicación individual tiene catalog_product_id: " + itemDetails.catalog_product_id + ". Resolviendo a catálogo...");
        return obtenerDatosDeMercadoLibre("https://www.mercadolibre.com.mx/p/" + itemDetails.catalog_product_id);
      }
      
      const price = itemDetails.price || 0;
      
      if (price === 0) {
        escribirLogDebug("⚠️ No se encontró precio válido para el item.");
        return null;
      }
      
      const picturesList = itemDetails.pictures || [];
      const secureUrl = picturesList.length > 0 ? picturesList[0].secure_url : (itemDetails.thumbnail || "");
      
      let color = "";
      let memoria = "";
      
      if (itemDetails.variations && itemDetails.variations.length > 0) {
        let colors = [];
        let memories = [];
        itemDetails.variations.forEach(v => {
          const isAvailable = v.available_quantity !== undefined ? v.available_quantity > 0 : true;
          if (isAvailable) {
            const combinations = v.attribute_combinations || [];
            const attrs = v.attributes || [];
            const allAttrs = combinations.concat(attrs);
            const attrColor = allAttrs.find(a => a.id === 'COLOR' || a.id === 'MAIN_COLOR');
            const attrStorage = allAttrs.find(a => a.id === 'INTERNAL_MEMORY' || a.id === 'MEMORY_CAPACITY');
            if (attrColor && attrColor.value_name && colors.indexOf(attrColor.value_name.trim()) === -1) {
              colors.push(attrColor.value_name.trim());
            }
            if (attrStorage && attrStorage.value_name && memories.indexOf(attrStorage.value_name.trim()) === -1) {
              memories.push(attrStorage.value_name.trim());
            }
          }
        });
        
        if (colors.length === 0) {
          const rootColor = extraerAtributo(itemDetails.attributes, "COLOR") || extraerAtributo(itemDetails.attributes, "MAIN_COLOR");
          if (rootColor) colors.push(rootColor);
        }
        if (memories.length === 0) {
          const rootMemoria = extraerAtributo(itemDetails.attributes, "INTERNAL_MEMORY") || extraerAtributo(itemDetails.attributes, "MEMORY_CAPACITY");
          if (rootMemoria) memories.push(rootMemoria);
        }
        
        color = colors.filter(Boolean).map(normalizarColorEspanol).filter(function(v, i, a) { return v && a.indexOf(v) === i; }).join(", ");
        memoria = memories.filter(Boolean).join(", ");
      } else {
        const rawColor = extraerAtributo(itemDetails.attributes, "COLOR") || extraerAtributo(itemDetails.attributes, "MAIN_COLOR") || "";
        color = normalizarColorEspanol(rawColor);
        memoria = extraerAtributo(itemDetails.attributes, "INTERNAL_MEMORY") || extraerAtributo(itemDetails.attributes, "MEMORY_CAPACITY");
      }
      
      escribirLogDebug("✅ Sincronizado desde publicación: " + itemDetails.title + " - Precio: $" + price + " - Memoria: " + memoria + " - Color: " + color);
      return {
        title: itemDetails.title,
        price: price,
        pictures: secureUrl ? [{ secure_url: secureUrl.replace("http://", "https://") }] : [],
        color: color,
        memoria: memoria
      };
    }
  } catch (err) {
    escribirLogDebug("❌ Excepción en obtenerDatosDeMercadoLibre: " + err.toString());
    return null;
  }
}

function escribirLogDebug(mensaje) {
  try {
    console.log(mensaje);
  } catch(e) {}
  
  try {
    const creds = obtenerPropiedadesEcosistema();
    if (!creds.SHEET_ID) return;
    const ss = SpreadsheetApp.openById(creds.SHEET_ID);
    let sheetLog = ss.getSheetByName("DEBUG_LOGS");
    if (!sheetLog) {
      sheetLog = ss.insertSheet("DEBUG_LOGS");
      sheetLog.appendRow(["Fecha y Hora", "Mensaje"]);
    }
    const fecha = Utilities.formatDate(new Date(), "America/Mexico_City", "yyyy-MM-dd HH:mm:ss");
    sheetLog.appendRow([fecha, mensaje]);
  } catch(e) {
    // Ignorar si falla
  }
}


// 🌐 REALIZA UN FETCH PASANDO EL TOKEN DE OAUTH Y REALIZA UN RETRY CON AUTO-REFRESCO SI EXPIRÓ
function fetchConToken(url) {
  const credenciales = obtenerPropiedadesEcosistema();
  let token = credenciales.ML_ACCESS_TOKEN;
  
  const headers = {};
  if (token) {
    headers["Authorization"] = "Bearer " + token;
  }
  
  let response = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
  const code = response.getResponseCode();
  
  if (code === 401 || code === 403) {
    escribirLogDebug("🔄 Token expirado o inválido (Código " + code + "). Intentando renovar automáticamente...");
    const nuevoToken = refrescarTokenML();
    if (nuevoToken) {
      headers["Authorization"] = "Bearer " + nuevoToken;
      response = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
      escribirLogDebug("📊 Respuesta tras renovar token: " + response.getResponseCode());
    }
  }
  return response;
}

// 🔄 RENOVAR EL ACCESS TOKEN DE MERCADO LIBRE USANDO EL REFRESH TOKEN
function refrescarTokenML() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('ML_CLIENT_ID') || "370163762307604";
  const clientSecret = props.getProperty('ML_CLIENT_SECRET') || "3uu7IZZB5u33IFGmND1FXyMF0bsYLo0u";
  const refreshToken = props.getProperty('ML_REFRESH_TOKEN');
  
  if (!clientId || !clientSecret || !refreshToken) {
    escribirLogDebug("❌ Error al refrescar token: Faltan credenciales en Propiedades del Script (ML_CLIENT_ID, ML_CLIENT_SECRET o ML_REFRESH_TOKEN).");
    return null;
  }
  
  const url = "https://api.mercadolibre.com/oauth/token";
  const payload = {
    grant_type: "refresh_token",
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken
  };
  
  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(text);
      props.setProperty('ML_ACCESS_TOKEN', data.access_token);
      props.setProperty('ML_REFRESH_TOKEN', data.refresh_token);
      escribirLogDebug("✅ Token de Mercado Libre renovado con éxito.");
      return data.access_token;
    } else {
      escribirLogDebug("❌ Error al renovar token de Mercado Libre (Código " + response.getResponseCode() + "): " + text);
      return null;
    }
  } catch (e) {
    escribirLogDebug("❌ Excepción al renovar token: " + e.toString());
    return null;
  }
}

// ⏰ CREAR DISPARADOR DE TIEMPO PARA RENOVAR EL TOKEN AUTOMÁTICAMENTE CADA 5 HORAS
function crearTriggerRefrescoToken() {
  const triggers = ScriptApp.getProjectTriggers();
  let yaExiste = false;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "refrescarTokenML") {
      yaExiste = true;
      break;
    }
  }
  
  if (!yaExiste) {
    ScriptApp.newTrigger("refrescarTokenML")
             .timeBased()
             .everyHours(5)
             .create();
    escribirLogDebug("⏰ Disparador para refrescarTokenML creado con éxito (cada 5 horas).");
  }
}

// 🏷️ AUXILIAR PARA EXTRAER ATRIBUTOS ESPECÍFICOS DE LA API DE MERCADO LIBRE
function extraerAtributo(attributes, idAttr) {
  if (!attributes || !Array.isArray(attributes)) return "";
  const attr = attributes.find(a => a.id === idAttr);
  return attr ? attr.value_name : "";
}

// 🧹 REPARACIÓN Y LIMPIEZA DE PRECIOS CACHEADOS EN EL CATÁLOGO DE MERCADOLIBRE
function limpiarPreciosML() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("Error al obtener Spreadsheet activo en limpiarPreciosML: " + e.toString());
  }
  if (!ss) {
    const creds = obtenerPropiedadesEcosistema();
    if (creds && creds.SHEET_ID) {
      ss = SpreadsheetApp.openById(creds.SHEET_ID);
    }
  }
  if (!ss) {
    Logger.log("❌ No se pudo obtener el Spreadsheet en limpiarPreciosML.");
    return "No se pudo obtener el Spreadsheet.";
  }
  const sheetInv = ss.getSheetByName("Catalogo");
  if (!sheetInv) {
    Logger.log("❌ No se encontró la pestaña 'Catalogo'");
    return "No se encontró la pestaña 'Catalogo'";
  }
  
  const ultFila = sheetInv.getLastRow();
  if (ultFila < 2) return "Catálogo vacío";
  
  const range = sheetInv.getRange(2, 1, ultFila - 1, 12);
  const data = range.getValues();
  let contadorLimpiados = 0;
  
  for (let i = 0; i < data.length; i++) {
    const imei = data[i][0].toString().trim();
    const estado = data[i][5].toString().toUpperCase().trim();
    const link = data[i][6] ? data[i][6].toString().trim() : "";
    const numeroFilaReal = i + 2;
    
    // Si es un celular de Mercado Libre (bajo demanda, mas vendido o inicia con 999999)
    if (link.includes("mercadolibre") || imei.startsWith("999999") || estado === "MAS VENDIDO" || estado === "BAJO DEMANDA") {
      // Ponemos el Precio Contado en 0 o vacío para obligar al sistema a recargar el precio costo bruto real en la siguiente búsqueda
      sheetInv.getRange(numeroFilaReal, 4).setValue(0);
      contadorLimpiados++;
    }
  }
  
  SpreadsheetApp.flush();
  Logger.log("✅ Se limpiaron " + contadorLimpiados + " precios de Mercado Libre para forzar su recálculo dinámico.");
  return "Se limpiaron " + contadorLimpiados + " precios de Mercado Libre.";
}

// 🌐 OBTIENE LAS VARIANTES COMPLETAS DE MERCADO LIBRE PARA CACHEAR EN EL GOOGLE SHEET
function obtenerDatosVariantesML(linkMercadoLibre, targetCondition) {
  try {
    const regex = /(MLM\-?\d+)/i;
    const coincidencia = linkMercadoLibre.match(regex);
    if (!coincidencia) return null;
    
    const cleanId = coincidencia[0].replace("-", "").toUpperCase();
    const isProduct = linkMercadoLibre.includes("/p/");
    
    let response = null;
    if (isProduct) {
      // Intentar primero el endpoint del producto de catálogo principal
      try {
        response = fetchConToken("https://api.mercadolibre.com/products/" + cleanId);
      } catch (prodErr) {
        Logger.log("Error consultando producto principal en obtenerDatosVariantesML: " + prodErr.toString());
      }
      
      // Fallback al primer item hijo si el producto falla o no es exitoso
      if (!response || response.getResponseCode() !== 200) {
        try {
          const itemsUrl = "https://api.mercadolibre.com/products/" + cleanId + "/items";
          const resItems = fetchConToken(itemsUrl);
          if (resItems && resItems.getResponseCode() === 200) {
            const itemsData = JSON.parse(resItems.getContentText());
            const results = itemsData.results || [];
            if (results.length > 0) {
              const firstItemId = results[0].id;
              response = fetchConToken("https://api.mercadolibre.com/items/" + firstItemId);
            }
          }
        } catch (itemErr) {
          Logger.log("Error buscando items en obtenerDatosVariantesML: " + itemErr.toString());
        }
      }
    } else {
      response = fetchConToken("https://api.mercadolibre.com/items/" + cleanId);
    }
    
    let finalJson = null;
    if (response && response.getResponseCode() === 200) {
      finalJson = JSON.parse(response.getContentText());
    } else {
      // Fallback
      let fallbackResponse = null;
      if (isProduct) {
        fallbackResponse = fetchConToken("https://api.mercadolibre.com/items/" + cleanId);
      } else {
        fallbackResponse = fetchConToken("https://api.mercadolibre.com/products/" + cleanId);
      }
      if (fallbackResponse && fallbackResponse.getResponseCode() === 200) {
        finalJson = JSON.parse(fallbackResponse.getContentText());
      }
    }

    if (finalJson) {
      if (finalJson.catalog_product_id) {
        try {
          const prodRes = fetchConToken("https://api.mercadolibre.com/products/" + finalJson.catalog_product_id);
          if (prodRes && prodRes.getResponseCode() === 200) {
            const prodData = JSON.parse(prodRes.getContentText());
            // Inject price if missing
            if (!prodData.price || !prodData.buy_box_winner) {
              const itemsUrl = "https://api.mercadolibre.com/products/" + finalJson.catalog_product_id + "/items";
              const itemsRes = fetchConToken(itemsUrl);
              if (itemsRes && itemsRes.getResponseCode() === 200) {
                const itemsData = JSON.parse(itemsRes.getContentText());
                const results = itemsData.results || [];
                if (results.length > 0) {
                  const targetCond = targetCondition || "new";
                  const resultsFiltrados = results.filter(function(r) { 
                    if (targetCond === "used") {
                      return r.condition === 'used' || r.condition === 'refurbished';
                    }
                    return r.condition === targetCond; 
                  });
                  const resultsAProcesar = resultsFiltrados.length > 0 ? resultsFiltrados : results;
                  resultsAProcesar.sort(function(a, b) { return (a.price || 0) - (b.price || 0); });
                  const bestItem = resultsAProcesar[0];
                  prodData.price = bestItem.price;
                  prodData.buy_box_winner = {
                    item_id: bestItem.item_id || bestItem.id,
                    price: bestItem.price,
                    shipping: bestItem.shipping,
                    seller_id: bestItem.seller_id,
                    condition: bestItem.condition
                  };
                }
              }
            }
            finalJson = prodData;
          }
        } catch (e) {
          Logger.log("Error auto-resolving catalog product in obtenerDatosVariantesML: " + e.toString());
        }
      } else if (isProduct && (!finalJson.price || !finalJson.buy_box_winner)) {
        // Inject price if catalog product directly loaded lacks price info
        try {
          const itemsUrl = "https://api.mercadolibre.com/products/" + cleanId + "/items";
          const itemsRes = fetchConToken(itemsUrl);
          if (itemsRes && itemsRes.getResponseCode() === 200) {
            const itemsData = JSON.parse(itemsRes.getContentText());
            const results = itemsData.results || [];
            if (results.length > 0) {
              const targetCond = targetCondition || "new";
              const resultsFiltrados = results.filter(function(r) { 
                if (targetCond === "used") {
                  return r.condition === 'used' || r.condition === 'refurbished';
                }
                return r.condition === targetCond; 
              });
              const resultsAProcesar = resultsFiltrados.length > 0 ? resultsFiltrados : results;
              resultsAProcesar.sort(function(a, b) { return (a.price || 0) - (b.price || 0); });
              const bestItem = resultsAProcesar[0];
              finalJson.price = bestItem.price;
              finalJson.buy_box_winner = {
                item_id: bestItem.item_id || bestItem.id,
                price: bestItem.price,
                shipping: bestItem.shipping,
                seller_id: bestItem.seller_id,
                condition: bestItem.condition
              };
            }
          }
        } catch (e) {
          Logger.log("Error injecting price in obtenerDatosVariantesML: " + e.toString());
        }
      }
      return finalJson;
    }
  } catch (err) {
    Logger.log("Excepcion en obtenerDatosVariantesML: " + err.toString());
  }
  return null;
}

// 💵 OBTIENE UN MAPA DE PRECIOS COSTO Y PRECIOS CONTADO FINAL DE CADA MODELO DESDE LA PESTAÑA 'Catalogo'
function obtenerPreciosDeCatalogo(ss) {
  const sheetCat = ss.getSheetByName('Catalogo');
  const preciosMap = {};
  if (sheetCat) {
    const dataCat = sheetCat.getDataRange().getValues();
    for (let j = 1; j < dataCat.length; j++) {
      const modeloStr = String(dataCat[j][1] || '').trim().toUpperCase();
      const costo = parseFloat(dataCat[j][3]) || 0; // Columna D: ML Costo (Costo)
      const precioContadoFinal = parseFloat(dataCat[j][11]) || 0; // Columna L: Precio Contado Final
      if (modeloStr) {
        preciosMap[modeloStr] = {
          costo: costo,
          precioContadoFinal: precioContadoFinal
        };
      }
    }
  }
  return preciosMap;
}

/**
 * Inicializa y configura la estructura completa del Spreadsheet.
 * Crea todas las pestañas necesarias (Clientes, Cotizaciones, Inventario, Catalogo, Pagos, Configuración, DEBUG_LOGS)
 * y establece las cabeceras exactas para evitar corrupciones de datos.
 */
function inicializarEstructuraSpreadsheet() {
  let ss = null;
  let getActiveError = "";
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    getActiveError = e.toString();
  }
  
  let openByIdError = "";
  let sheetIdUsed = "";
  if (!ss) {
    try {
      const creds = obtenerPropiedadesEcosistema();
      sheetIdUsed = creds ? creds.SHEET_ID : null;
      if (sheetIdUsed) {
        ss = SpreadsheetApp.openById(sheetIdUsed.trim());
      }
    } catch (e) {
      openByIdError = e.toString();
    }
  }
  
  if (!ss) {
    let propKeys = [];
    try {
      const allProps = PropertiesService.getScriptProperties().getProperties();
      propKeys = Object.keys(allProps);
    } catch (e) {
      propKeys = ["Error leyendo propiedades: " + e.toString()];
    }
    const msg = "❌ Error: No se pudo obtener el Spreadsheet.\n" +
                "- getActiveSpreadsheet error: " + (getActiveError || "Ninguno (devolvió null)") + "\n" +
                "- openById error: " + (openByIdError || "Ninguno") + "\n" +
                "- SHEET_ID intentado: '" + (sheetIdUsed || "") + "'\n" +
                "- Propiedades encontradas en el Script: " + JSON.stringify(propKeys);
    Logger.log(msg);
    throw new Error(msg);
  }
  
  // 1. Pestaña 'Clientes'
  let sheetClientes = ss.getSheetByName('Clientes');
  if (!sheetClientes) {
    sheetClientes = ss.insertSheet('Clientes');
  }
  const headersClientes = [
    "Timestamp", "Cliente", "Teléfono / WhatsApp", "Fecha de Nacimiento", 
    "Modelo", "IMEI", "Monto Enganche", "Dirección", "Teléfono Aval", 
    "Link Firma", "Cuota", "Interés", "Horas Tolerancia", "Semanas Tolerancia", 
    "Plazo", "Total Financiado", "Total en Letra", "Fecha Vencimiento", 
    "Día Raya", "ID Cliente", "Tipo de Periodo"
  ];
  sheetClientes.getRange(1, 1, 1, headersClientes.length).setValues([headersClientes]);
  
  // 2. Pestaña 'Cotizaciones'
  let sheetCot = ss.getSheetByName('Cotizaciones');
  if (!sheetCot) {
    sheetCot = ss.insertSheet('Cotizaciones');
  }
  const headersCot = [
    "Fecha", "Cliente", "WhatsApp / Teléfono", "Enlace Mercado Libre", 
    "Plazo Solicitado", "Estado de Crédito", "Sueldo", "Préstamos Activos", "Color", 
    "Capacidad", "Link Prellenado", "Plan Elegido", "Frecuencia", "ID Cliente", "Estado de Cotización"
  ];
  sheetCot.getRange(1, 1, 1, headersCot.length).setValues([headersCot]);
  
  // 3. Pestaña 'Inventario'
  let sheetInv = ss.getSheetByName('Inventario');
  if (!sheetInv) {
    sheetInv = ss.insertSheet('Inventario');
  }
  const headersInv = [
    "IMEI", "Modelo", "Color", "Memoria", "Estado", "Notas", "Precio Final"
  ];
  sheetInv.getRange(1, 1, 1, headersInv.length).setValues([headersInv]);
  
  // 4. Pestaña 'Catalogo'
  let sheetCat = ss.getSheetByName('Catalogo');
  if (!sheetCat) {
    sheetCat = ss.insertSheet('Catalogo');
  }
  const headersCat = [
    "IMEI", "Modelo", "Color", "ML Costo", "Notas", "Estatus",
    "Enlace Mercado Libre", "Link Imagen", "Capacidad", "Color Original",
    "Variantes JSON", "Precio Contado Final", "Colores Disponibles"
  ];
  // Limpiar posibles columnas viejas sobrantes (ej. Filtro4, etc.) si ya existen
  if (sheetCat.getLastColumn() > headersCat.length) {
    const diff = sheetCat.getLastColumn() - headersCat.length;
    sheetCat.deleteColumns(headersCat.length + 1, diff);
  }
  sheetCat.getRange(1, 1, 1, headersCat.length).setValues([headersCat]);
  
  // 5. Pestaña 'Pagos'
  let sheetPagos = ss.getSheetByName('Pagos');
  if (!sheetPagos) {
    sheetPagos = ss.insertSheet('Pagos');
  }
  const headersPagos = [
    "IMEI", "Cliente", "Fecha Pago", "Monto Pago", "Método Pago", 
    "Semana Pago", "Folio Pago", "Notas Pago"
  ];
  sheetPagos.getRange(1, 1, 1, headersPagos.length).setValues([headersPagos]);
  
  // 6. Pestaña 'Configuración'
  let sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet('Configuración');
  }
  const headersConfig = [
    "Tasa de Interés Semanal por Atraso (%)", 
    "Horas de Tolerancia antes de Bloquear Celular (Horas)", 
    "Semanas Máximas de Atraso antes de Rescisión (Semanas)", 
    "Tasa Interés Semanal Rápido (%)", 
    "Tasa Interés Semanal Cómodo (%)", 
    "Margen de Ganancia para Precio de Contado (%)", 
    "Enganche Mínimo Semanal Rápido (%)", 
    "--- Reservado (No Usar) ---", 
    "Enganche Mínimo Semanal Cómodo (%)", 
    "Descuento por Liquidación Anticipada (%)",
    "Telegram Bot Token (API)", 
    "Telegram Chat ID para Notificaciones",
    "Tasa Interés Quincenal Rápido (%)",
    "Tasa Interés Quincenal Cómodo (%)",
    "Enganche Mínimo Quincenal Rápido (%)",
    "Enganche Mínimo Quincenal Cómodo (%)"
  ];
  sheetConfig.getRange(1, 1, 1, headersConfig.length).setValues([headersConfig]);
  
  // Rellenar valores por defecto en Fila 2 de Configuración si está vacía
  if (sheetConfig.getLastRow() < 2) {
    sheetConfig.getRange("A2:P2").setValues([[6, 24, 3, 50, 75, 15, 20, 33, 30, 15, "", "", 45, 65, 25, 35]]);
  } else {
    // Si ya existe la fila 2, asegurar que los campos tengan valores correctos o por defecto (migración a nuevos valores)
    const rangeTodos = sheetConfig.getRange("A2:P2");
    const valoresTodos = rangeTodos.getValues()[0];
    const actualizados = [
      valoresTodos[0] !== "" && valoresTodos[0] !== undefined && valoresTodos[0] !== null ? valoresTodos[0] : 6,
      valoresTodos[1] !== "" && valoresTodos[1] !== undefined && valoresTodos[1] !== null ? valoresTodos[1] : 24,
      valoresTodos[2] !== "" && valoresTodos[2] !== undefined && valoresTodos[2] !== null ? valoresTodos[2] : 3,
      (valoresTodos[3] === "" || valoresTodos[3] === undefined || valoresTodos[3] === null || valoresTodos[3] === 66) ? 50 : valoresTodos[3],
      (valoresTodos[4] === "" || valoresTodos[4] === undefined || valoresTodos[4] === null || valoresTodos[4] === 90) ? 75 : valoresTodos[4],
      valoresTodos[5] !== "" && valoresTodos[5] !== undefined && valoresTodos[5] !== null ? valoresTodos[5] : 15,
      (valoresTodos[6] === "" || valoresTodos[6] === undefined || valoresTodos[6] === null || valoresTodos[6] === 16.5) ? 20 : valoresTodos[6],
      valoresTodos[7] !== "" && valoresTodos[7] !== undefined && valoresTodos[7] !== null ? valoresTodos[7] : 33,
      (valoresTodos[8] === "" || valoresTodos[8] === undefined || valoresTodos[8] === null || valoresTodos[8] === 11) ? 30 : valoresTodos[8],
      valoresTodos[9] !== "" && valoresTodos[9] !== undefined && valoresTodos[9] !== null ? valoresTodos[9] : 15,
      valoresTodos[10] || "",
      valoresTodos[11] || "",
      (valoresTodos[12] === "" || valoresTodos[12] === undefined || valoresTodos[12] === null) ? 45 : valoresTodos[12],
      (valoresTodos[13] === "" || valoresTodos[13] === undefined || valoresTodos[13] === null) ? 65 : valoresTodos[13],
      (valoresTodos[14] === "" || valoresTodos[14] === undefined || valoresTodos[14] === null) ? 25 : valoresTodos[14],
      (valoresTodos[15] === "" || valoresTodos[15] === undefined || valoresTodos[15] === null) ? 35 : valoresTodos[15]
    ];
    rangeTodos.setValues([actualizados]);
  }
  
  // Desactivar triggers antiguos para que los maneje n8n
  try {
    desactivarTriggersAntiguos();
  } catch (trigErr) {
    escribirLogDebug("Error al limpiar triggers antiguos: " + trigErr.toString());
  }

  // 7. Pestaña 'DEBUG_LOGS'
  let sheetLog = ss.getSheetByName('DEBUG_LOGS');
  if (!sheetLog) {
    sheetLog = ss.insertSheet('DEBUG_LOGS');
  }
  const headersLog = ["Fecha y Hora", "Mensaje"];
  sheetLog.getRange(1, 1, 1, headersLog.length).setValues([headersLog]);

  // 8. Pestaña 'Promesas'
  let sheetPromesas = ss.getSheetByName('Promesas');
  if (!sheetPromesas) {
    sheetPromesas = ss.insertSheet('Promesas');
  }
  const headersPromesas = ["IMEI", "Cliente", "Fecha Promesa", "Monto Prometido", "Estatus", "Fecha Registro", "Notas"];
  sheetPromesas.getRange(1, 1, 1, headersPromesas.length).setValues([headersPromesas]);
  
  // 9. Pestaña 'Cierres'
  let sheetCierres = ss.getSheetByName('Cierres');
  if (!sheetCierres) {
    sheetCierres = ss.insertSheet('Cierres');
  }
  const headersCierres = ["Fecha Cierre", "Efectivo Cobrado", "Transferencia Cobrada", "Tarjeta Cobrada", "Total Cobrado", "Diferencia Efectivo", "Notas Cierre"];
  sheetCierres.getRange(1, 1, 1, headersCierres.length).setValues([headersCierres]);
  
  Logger.log("✅ ¡Estructura del Spreadsheet inicializada y configurada correctamente!");
}

// 🎨 EXTRAE LOS COLORES DISPONIBLES DESDE EL JSON DE VARIANTES DE MERCADO LIBRE
function extraerColoresDeDatosML(data) {
  if (!data) return "";
  
  // Robustez en caso de recibir string
  if (typeof data === "string") {
    try {
      data = JSON.parse(data);
    } catch (e) {
      Logger.log("Error parseando data en extraerColoresDeDatosML: " + e.toString());
      return "";
    }
  }

  const colores = [];
  
  // Caso 1: Tiene pickers (Producto de catálogo)
  if (data.pickers && Array.isArray(data.pickers)) {
    const colorPicker = data.pickers.find(p => {
      if (!p.picker_id) return false;
      const idUpper = p.picker_id.toUpperCase();
      return idUpper.includes("COLOR") || idUpper === "FINISH" || idUpper === "ACABADO";
    });
    if (colorPicker && colorPicker.products && Array.isArray(colorPicker.products)) {
      colorPicker.products.forEach(prod => {
        // Filtrar si la variante de color está agotada (tags: out-of-stock o disabled)
        const isOutOfStock = prod.tags && (prod.tags.indexOf("out-of-stock") !== -1 || prod.tags.indexOf("disabled") !== -1);
        if (!isOutOfStock && prod.picker_label) {
          const c = prod.picker_label.trim();
          if (c && colores.indexOf(c) === -1) {
            colores.push(c);
          }
        }
      });
    }
  }
  
  // Caso 2: Tiene variations (Item individual de Mercado Libre)
  if (colores.length === 0 && data.variations && Array.isArray(data.variations)) {
    data.variations.forEach(v => {
      // Filtrar si la cantidad disponible es menor o igual a 0
      const isAvailable = v.available_quantity !== undefined ? v.available_quantity > 0 : true;
      if (isAvailable) {
        if (v.attribute_combinations && Array.isArray(v.attribute_combinations)) {
          v.attribute_combinations.forEach(attr => {
            if (attr.id === "COLOR" || (attr.name && attr.name.toLowerCase().includes("color"))) {
              if (attr.value_name) {
                const c = attr.value_name.trim();
                if (c && colores.indexOf(c) === -1) {
                  colores.push(c);
                }
              }
            }
          });
        }
      }
    });
  }

  // Caso 3: Fallback al color principal del item
  if (colores.length === 0 && data.attributes && Array.isArray(data.attributes)) {
    const attrColor = data.attributes.find(a => a.id === "COLOR" || (a.name && a.name.toLowerCase().includes("color")));
    if (attrColor && attrColor.value_name) {
      const c = attrColor.value_name.trim();
      if (c) colores.push(c);
    }
  }
  
  // Caso 4: Recorrer results (si es un fallback de items de catálogo con múltiples ofertas de vendedores)
  if (colores.length === 0 && data.results && Array.isArray(data.results)) {
    data.results.forEach(function(item) {
      let colorEncontrado = "";
      
      // 1. Buscar en atributos
      if (item.attributes && Array.isArray(item.attributes)) {
        const attrColor = item.attributes.find(a => a.id === "COLOR" || (a.name && a.name.toLowerCase().includes("color")));
        if (attrColor && attrColor.value_name) {
          colorEncontrado = attrColor.value_name.trim();
        }
      }
      
      // 2. Si no se encontró en atributos, escanear el título (title) o permalink
      if (!colorEncontrado) {
        const textoBuscar = String((item.title || "") + " " + (item.permalink || "")).toLowerCase();
        
        // Mapeo de palabras clave comunes a nombres estándar
        const mapaColores = {
          "negro": "Negro", "black": "Negro", "dark": "Negro",
          "gris": "Gris", "gray": "Gris", "grey": "Gris", "charcoal": "Gris", "plata": "Gris", "silver": "Gris",
          "azul": "Azul", "blue": "Azul",
          "verde": "Verde", "green": "Verde", "mint": "Verde",
          "rojo": "Rojo", "red": "Rojo",
          "rosa": "Rosa", "pink": "Rosa",
          "lila": "Lila", "violeta": "Lila", "morado": "Lila", "purple": "Lila",
          "dorado": "Dorado", "gold": "Dorado",
          "blanco": "Blanco", "white": "Blanco"
        };
        
        for (const clave in mapaColores) {
          if (textoBuscar.indexOf(clave) !== -1) {
            colorEncontrado = mapaColores[clave];
            break;
          }
        }
      }
      
      if (colorEncontrado && colores.indexOf(colorEncontrado) === -1) {
        colores.push(colorEncontrado);
      }
    });
  }
  
  const coloresNormalizados = colores.map(normalizarColorEspanol).filter(function(v, i, a) { return v && a.indexOf(v) === i; });
  return coloresNormalizados.join(", ");
}

// 🧼 LIMPIA Y SANITIZA EL NOMBRE DEL MODELO PARA EL SPREADSHEET (Quita color, capacidad, etc.)
function limpiarNombreModelo(modelo) {
  if (!modelo) return "";
  let clean = modelo.trim();

  // 0. Quitar todo a partir de un dos puntos (:)
  if (clean.includes(":")) {
    clean = clean.split(":")[0].trim();
  }

  const isRefurb = /reacondicionado|refurbished|refurb/i.test(modelo);

  // 1. Remover fragmentos de reacondicionado y detalles de caja/condición
  clean = clean.replace(/-\s*(excelente|muy\s*bueno|bueno|aceptable|regular)\s*(\(reacondicionado\))?/gi, "");
  clean = clean.replace(/\b(reacondicionado|refurbished|refurb)\b/gi, "");
  clean = clean.replace(/\b(caja\s*(maltratada|dañada|abierta|original|sellada))\b/gi, "");
  clean = clean.replace(/\b(dual\s*sim|single\s*sim)\b/gi, "");
  clean = clean.replace(/\b(desbloqueado|liberado|unlocked)\b/gi, "");
  clean = clean.replace(/\b(nacional|internacional|global)\b/gi, "");

  // 2. Remover "Smartphone", "Celular", "Teléfono", "Telefono" del inicio del nombre
  clean = clean.replace(/^(smartphone|celular|teléfono|telefono)\b/gi, "").trim();

  // 3. Remover cualquier paréntesis () que NO contenga "reacondicionado"
  clean = clean.replace(/\((?!reacondicionado\b)[^)]*\)/gi, "").trim();

  // 4. Remover colores conocidos (incluyendo variaciones y acabados)
  const colores = [
    "black", "negro", "blue", "azul", "white", "blanco", "green", "verde", "silver", "plata",
    "gray", "gris", "gold", "oro", "purple", "morado", "purpura", "púrpura", "pink", "rosa", "red", "rojo", "light",
    "dark", "sky", "navy", "rose", "violet", "violeta", "lavender", "lavanda", "carbon", "carbono",
    "claro", "oscuro", "space", "estelar", "celeste", "brillante", "mate",
    "phantom", "cream", "crema", "beige", "coral", "titanium", "titanio",
    "graphite", "grafito", "burgundy", "borgoña", "midnight", "medianoche",
    "starlight", "ice", "hielo", "bronze", "bronce", "sand", "arena",
    "indigo", "teal", "mint", "menta", "lime", "lima", "ivory", "marfil",
    "cuero", "vegano", "platino", "platinum", "obsidiana", "obsidian",
    "naranja", "orange", "amarillo", "yellow", "marrón", "marron", "brown",
    "tierras", "raras", "aurora", "glowing", "rainbow", "sunset", "sunrise",
    "glacier", "glaciar", "plateado", "arandano", "charcoal", "matte"
  ];
  colores.forEach(color => {
    const regex = new RegExp("\\b" + color + "\\b", "gi");
    clean = clean.replace(regex, "");
  });

  // 5. Remover capacidades de almacenamiento y RAM por completo
  const regexGb = /(\d+)\s*(gb|tb|ram)\b/gi;
  clean = clean.replace(regexGb, "");

  // Remover palabras conectoras/redundantes
  clean = clean.replace(/\b(con|de|color|ram)\b/gi, "");

  // Limpiar múltiples espacios en blanco y caracteres colgantes
  clean = clean.replace(/\s+/g, " ").trim();
  clean = clean.replace(/^[,\-\s]+|[,\-\s]+$/g, "");

  if (isRefurb) {
    clean += " (Reacondicionado)";
  }
  return clean;
}

function extraerLinkDeArrays(values, formulas, richTexts, rowIdx, colIdx) {
  if (richTexts && richTexts[rowIdx] && richTexts[rowIdx][colIdx]) {
    const url = richTexts[rowIdx][colIdx].getLinkUrl();
    if (url) return url;
  }
  if (formulas && formulas[rowIdx] && formulas[rowIdx][colIdx]) {
    const formula = formulas[rowIdx][colIdx];
    if (formula && formula.startsWith("=")) {
      if (formula.toUpperCase().indexOf("HYPERLINK") !== -1) {
        const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
        if (match) return match[1];
      }
    }
  }
  const val = values && values[rowIdx] ? String(values[rowIdx][colIdx] || '') : "";
  if (val.indexOf("http://") === 0 || val.indexOf("https://") === 0) {
    return val;
  }
  return "";
}

// ☀️ TELEGRAM BOT NOTIFICATIONS & TRIGGERS ☀️


function crearTriggerResumenMatutino() {
  const triggers = ScriptApp.getProjectTriggers();
  let existe = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarResumenMatutinoTelegram') {
      existe = true;
      break;
    }
  }
  if (!existe) {
    ScriptApp.newTrigger('enviarResumenMatutinoTelegram')
             .timeBased()
             .everyDays(1)
             .atHour(8)
             .nearMinute(0)
             .create();
    escribirLogDebug("Trigger para enviarResumenMatutinoTelegram creado exitosamente.");
  }
}

function desactivarTriggersAntiguos() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fnName = triggers[i].getHandlerFunction();
    if (fnName === "enviarResumenMatutinoTelegram" || fnName === "refrescarTokenML") {
      try {
        ScriptApp.deleteTrigger(triggers[i]);
        escribirLogDebug("⏰ Trigger " + fnName + " eliminado con éxito (migrado a n8n).");
      } catch (e) {
        escribirLogDebug("⚠️ Error eliminando trigger " + fnName + ": " + e.toString());
      }
    }
  }
}

function normalizarDiaSemana(dia) {
  if (!dia) return "";
  return String(dia).toUpperCase().trim()
    .replace(/[ÁÁ]/g, "A")
    .replace(/[ÉÉ]/g, "E")
    .replace(/[ÍÍ]/g, "I")
    .replace(/[ÓÓ]/g, "O")
    .replace(/[ÚÚ]/g, "U")
    .replace(/[ÑÑ]/g, "N");
}

function enviarResumenMatutinoTelegram() {
  try {
    let ss = null;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
      Logger.log("getActiveSpreadsheet falló: " + e.toString());
    }
    if (!ss) {
      const creds = obtenerPropiedadesEcosistema();
      if (creds && creds.SHEET_ID) {
        ss = SpreadsheetApp.openById(creds.SHEET_ID);
      }
    }
    if (!ss) {
      Logger.log("No se pudo obtener el Spreadsheet en enviarResumenMatutinoTelegram.");
      return;
    }
    
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    if (!sheetConfig) return;
    
    const filaConfig = sheetConfig.getRange("A2:L2").getValues()[0];
    const botToken = String(filaConfig[10] || '').trim();
    const chatId = String(filaConfig[11] || '').trim();
    
    if (!botToken || !chatId) {
      Logger.log("Telegram Token o Chat ID vacíos en Configuración.");
      return;
    }
    
    const interesTasa = parseFloat(filaConfig[0]) || 6;
    const semanasToleradas = parseInt(filaConfig[2]) || 3;
    
    const sheetClientes = ss.getSheetByName('Clientes');
    const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];
    
    const sheetPagos = ss.getSheetByName('Pagos');
    const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
    
    // Agrupar pagos por IMEI
    const pagosPorImei = {};
    for (let j = 1; j < dataPagos.length; j++) {
      const imeiPago = String(dataPagos[j][0]).trim();
      const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
      if (imeiPago) {
        pagosPorImei[imeiPago] = (pagosPorImei[imeiPago] || 0) + montoPago;
      }
    }
    
    const diasSemana = ["DOMINGO", "LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"];
    const hoyDiaSemana = diasSemana[new Date().getDay()];
    
    let totalClientes = 0;
    let saldoTotalFinanciado = 0;
    let saldoTotalPendiente = 0;
    let totalAlCorriente = 0;
    let totalAtrasoLeve = 0;
    let totalAtrasoModerado = 0;
    let totalAtrasoCritico = 0;
    let listaRescision = [];
    let clientesRayaHoy = [];
    
    for (let i = 1; i < dataClientes.length; i++) {
      const startTimestamp = dataClientes[i][0];
      const clienteNombre = String(dataClientes[i][1] || '').trim();
      const imeiKey = String(dataClientes[i][5] || '').trim();
      const cuotaNum = parseFloat(String(dataClientes[i][10] || "0").replace(/[\$,]/g, "")) || 0;
      const totalFinanciadoVal = parseFloat(String(dataClientes[i][15] || "0").replace(/[\$,]/g, "")) || 0;
      const tipoPeriodo = String(dataClientes[i][20] || "SEMANAL").toUpperCase();
      
      if (clienteNombre && imeiKey) {
        totalClientes++;
        saldoTotalFinanciado += totalFinanciadoVal;
        
        const totalPagado = pagosPorImei[imeiKey] || 0;
        const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);
        
        // Calcular atraso
        let diasAtrasoVal = 0;
        if (startTimestamp instanceof Date && cuotaNum > 0 && saldoPendienteBase > 0) {
          const diasTranscurridos = Math.floor((new Date() - startTimestamp) / (1000 * 60 * 60 * 24));
          const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
          const periodosTranscurridos = Math.floor(diasTranscurridos / diasPorPeriodo);
          const totalEsperado = periodosTranscurridos * cuotaNum;
          if (totalPagado < totalEsperado) {
            const atrasoMonto = totalEsperado - totalPagado;
            diasAtrasoVal = Math.ceil((atrasoMonto / cuotaNum) * diasPorPeriodo);
          }
        }
        
        let penalidadMonto = 0;
        if (diasAtrasoVal > 0) {
          penalidadMonto = diasAtrasoVal * (cuotaNum * (interesTasa / 100));
        }
        const saldoPendienteReal = saldoPendienteBase + penalidadMonto;
        saldoTotalPendiente += saldoPendienteReal;
        
        const diaRayaCliente = normalizarDiaSemana(dataClientes[i][18]); // Columna S: Día Raya (index 18)
        let esDiaDePago = false;
        
        if (diaRayaCliente === hoyDiaSemana) {
          esDiaDePago = true;
        } else if (diaRayaCliente.includes("15") || diaRayaCliente.includes("30")) {
          // Cliente con esquema quincenal fijo (ej. "15 Y 30")
          const hoy = new Date();
          const diaMes = hoy.getDate();
          const ultimoDiaMes = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
          
          // Cobrar el día 15, el 30, el 31 (si el mes tiene 31 días) o el último día del mes (para Febrero)
          if (diaMes === 15 || diaMes === 30 || diaMes === 31 || (diaMes === ultimoDiaMes && ultimoDiaMes < 30)) {
            esDiaDePago = true;
          }
        }
        
        if (esDiaDePago && saldoPendienteReal > 0) {
          clientesRayaHoy.push(`• <b>${clienteNombre}</b> (Cuota: $${Math.round(cuotaNum)} MXN) - /cobro_${imeiKey}`);
        }
        
        if (diasAtrasoVal === 0) {
          totalAlCorriente++;
        } else if (diasAtrasoVal <= 7) {
          totalAtrasoLeve++;
        } else if (diasAtrasoVal <= 21) {
          totalAtrasoModerado++;
        } else {
          totalAtrasoCritico++;
          listaRescision.push(clienteNombre + " (" + imeiKey + ") - " + semanasAtrasoVal + " sem atraso");
        }
      }
    }
    
    // Buscar promesas de pago para hoy
    const sheetPromesas = ss.getSheetByName('Promesas');
    const dataProm = sheetPromesas ? sheetPromesas.getDataRange().getValues() : [];
    let promesasHoyTexto = "";
    
    const hoyStr1 = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const hoyStr2 = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for (let k = 1; k < dataProm.length; k++) {
      const imeiP = String(dataProm[k][0] || '').trim();
      const clienteP = String(dataProm[k][1] || '').trim();
      let fechaP = dataProm[k][2];
      const montoP = parseFloat(dataProm[k][3]) || 0;
      const estatusP = String(dataProm[k][4] || '').trim().toUpperCase();
      
      if (estatusP === "PENDIENTE") {
        let fStr = "";
        if (fechaP instanceof Date) {
          fStr = Utilities.formatDate(fechaP, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (fechaP) {
          fStr = String(fechaP).trim();
        }
        
        if (fStr === hoyStr1 || fStr === hoyStr2) {
          promesasHoyTexto += `📌 <b>${clienteP}</b> te prometió pagar hoy <b>$${montoP}</b> (IMEI: ${imeiP})\n`;
          promesasHoyTexto += `📝 <i>Mensaje para copiar y pegar:</i>\n`;
          promesasHoyTexto += `<code>Hola ${clienteP}, te recordamos tu promesa de pago registrada para hoy por $${montoP}. Quedamos al pendiente de tu comprobante. ¡Gracias!</code>\n\n`;
        }
      }
    }
    
    let msg = `☀️ <b>Resumen Matutino CelYa (8:00 AM)</b>\n\n`;
    msg += `📊 <b>Métricas de Cartera:</b>\n`;
    msg += `• Clientes Activos: <b>${totalClientes}</b>\n`;
    msg += `• Cartera Pendiente Total: <b>$${Math.round(saldoTotalPendiente).toLocaleString('es-MX')}</b>\n\n`;
    msg += `🚦 <b>Semáforo Financiero:</b>\n`;
    msg += `• Al Corriente: <b>${totalAlCorriente}</b>\n`;
    msg += `• Atraso Leve (1-7 días): <b>${totalAtrasoLeve}</b>\n`;
    msg += `• Atraso Moderado (8-21 días): <b>${totalAtrasoModerado}</b>\n`;
    msg += `• Atraso Crítico (22+ días): <b>${totalAtrasoCritico}</b>\n\n`;
    
    if (listaRescision.length > 0) {
      msg += `⚠️ <b>Clientes en Atraso Crítico / Bloqueo:</b>\n`;
      listaRescision.forEach(function(item) {
        msg += `• ${item}\n`;
      });
      msg += `\n`;
    }
    
    let rayaHoyTexto = "";
    if (clientesRayaHoy.length > 0) {
      rayaHoyTexto = `💸 <b>Días de Raya de Hoy (${hoyDiaSemana}):</b>\n` + clientesRayaHoy.join("\n") + `\n\n`;
    } else {
      rayaHoyTexto = `💸 No hay cobros programados para hoy por Día de Raya.\n\n`;
    }
    msg += rayaHoyTexto;
    
    if (promesasHoyTexto) {
      msg += `📅 <b>Promesas de Pago de Hoy:</b>\n${promesasHoyTexto}`;
    } else {
      msg += `📅 No hay promesas de pago programadas para hoy.\n`;
    }
    
    enviarMensajeTelegram(botToken, chatId, msg);
    
    // Enviar recordatorios individuales de bloqueo y rescisión con mensajes de copiar y pegar
    const alertas = obtenerAlertasClientes(ss);
    
    if (alertas.rescisiones.length > 0) {
      alertas.rescisiones.forEach(function(c) {
        let alertMsg = `⚠️ <b>AVISO DE RESCISIÓN DE CONTRATO</b> ⚠️\n\n` +
                       `👤 <b>Cliente:</b> ${c.cliente}\n` +
                       `📱 <b>Equipo:</b> ${c.modelo}\n` +
                       `⏳ <b>Atraso:</b> ${c.semanasAtraso} semanas\n` +
                       `💵 <b>Saldo Pendiente:</b> $${c.saldo.toLocaleString('es-MX')}\n\n` +
                       `📝 <b>Mensaje para copiar y pegar:</b>\n` +
                       `<code>${c.mensajeCopiar}</code>`;
        enviarMensajeTelegram(botToken, chatId, alertMsg);
      });
    }
    
    if (alertas.bloqueos.length > 0) {
      alertas.bloqueos.forEach(function(c) {
        let alertMsg = `🚨 <b>AVISO DE BLOQUEO DE ACCESO</b> 🚨\n\n` +
                       `👤 <b>Cliente:</b> ${c.cliente}\n` +
                       `📱 <b>Equipo:</b> ${c.modelo}\n` +
                       `⏳ <b>Atraso:</b> ${c.horasAtraso} horas\n` +
                       `💵 <b>Saldo Pendiente:</b> $${c.saldo.toLocaleString('es-MX')}\n\n` +
                       `📝 <b>Mensaje para copiar y pegar:</b>\n` +
                       `<code>${c.mensajeCopiar}</code>`;
        enviarMensajeTelegram(botToken, chatId, alertMsg);
      });
    }
  } catch (error) {
    Logger.log("Error en enviarResumenMatutinoTelegram: " + error.toString());
    throw new Error("enviarResumenMatutinoTelegram falló: " + error.message);
  }
}

// 🤖 PROCESAMIENTO DE WEBHOOK Y COMANDOS DE TELEGRAM 🤖

function procesarWebhookTelegram(update, credenciales) {
  let chatId = "";
  if (update.message) {
    chatId = String(update.message.chat.id);
  } else if (update.callback_query) {
    chatId = String(update.callback_query.message.chat.id);
  }

  try {
    const ss = SpreadsheetApp.openById(credenciales.SHEET_ID);
    
    // Obtener token y chatId de configuración en la hoja como fallback
    const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
    const filaConfig = sheetConfig ? sheetConfig.getRange("A2:L2").getValues()[0] : null;
    const sheetBotToken = filaConfig ? String(filaConfig[10] || '').trim() : "";
    
    const botToken = credenciales.TELEGRAM_TOKEN || sheetBotToken;
    if (!botToken) {
      Logger.log("Telegram Bot Token no configurado.");
      return;
    }
    let text = "";
    let isCallback = false;
    let callbackQueryId = "";

    if (update.message) {
      text = String(update.message.text || '').trim();
    } else if (update.callback_query) {
      isCallback = true;
      callbackQueryId = update.callback_query.id;
      text = String(update.callback_query.data || '').trim();
    }

    if (!chatId || !text) return;

    // Comandos generales
    if (text === "/start" || text === "/help") {
      let msg = `🤖 <b>CelYa Bot de Cobranza</b>\n\n`;
      msg += `Comandos disponibles:\n`;
      msg += `/clientes - Ver la lista de todos los clientes activos\n`;
      msg += `/morosos - Ver la lista de clientes con atraso\n`;
      msg += `/bloqueos - Ver clientes listos para bloqueo de equipo\n`;
      msg += `/rescisiones - Ver clientes listos para rescisión de contrato\n\n`;
      msg += `Toca en cualquier cliente de la lista para obtener su mensaje de cobro copiable al instante.`;
      enviarMensajeTelegram(botToken, chatId, msg);
      return;
    }

    if (text === "/clientes" || text === "ver_todos_clientes") {
      enviarListaClientesTelegram(ss, botToken, chatId, false);
      if (isCallback) {
        responderCallbackQuery(botToken, callbackQueryId);
      }
      return;
    }

    if (text === "/morosos" || text === "ver_clientes_morosos") {
      enviarListaClientesTelegram(ss, botToken, chatId, true);
      if (isCallback) {
        responderCallbackQuery(botToken, callbackQueryId);
      }
      return;
    }

    if (text === "/bloqueos" || text === "ver_bloqueos") {
      const alertas = obtenerAlertasClientes(ss);
      if (alertas.bloqueos.length === 0) {
        enviarMensajeTelegram(botToken, chatId, "🟢 No hay equipos con atraso suficiente para bloquear acceso.");
      } else {
        alertas.bloqueos.forEach(function(c) {
          let alertMsg = `🚨 <b>AVISO DE BLOQUEO DE ACCESO</b> 🚨\n\n` +
                         `👤 <b>Cliente:</b> ${c.cliente}\n` +
                         `📱 <b>Equipo:</b> ${c.modelo}\n` +
                         `⏳ <b>Atraso:</b> ${c.horasAtraso} horas\n` +
                         `💵 <b>Saldo Pendiente:</b> $${c.saldo.toLocaleString('es-MX')}\n\n` +
                         `📝 <b>Mensaje para copiar y pegar:</b>\n` +
                         `<code>${c.mensajeCopiar}</code>`;
          enviarMensajeTelegram(botToken, chatId, alertMsg);
        });
      }
      if (isCallback) {
        responderCallbackQuery(botToken, callbackQueryId);
      }
      return;
    }

    if (text === "/rescisiones" || text === "ver_rescisiones") {
      const alertas = obtenerAlertasClientes(ss);
      if (alertas.rescisiones.length === 0) {
        enviarMensajeTelegram(botToken, chatId, "🟢 No hay contratos con atraso suficiente para rescisión.");
      } else {
        alertas.rescisiones.forEach(function(c) {
          let alertMsg = `⚠️ <b>AVISO DE RESCISIÓN DE CONTRATO</b> ⚠️\n\n` +
                         `👤 <b>Cliente:</b> ${c.cliente}\n` +
                         `📱 <b>Equipo:</b> ${c.modelo}\n` +
                         `⏳ <b>Atraso:</b> ${c.semanasAtraso} semanas\n` +
                         `💵 <b>Saldo Pendiente:</b> $${c.saldo.toLocaleString('es-MX')}\n\n` +
                         `📝 <b>Mensaje para copiar y pegar:</b>\n` +
                         `<code>${c.mensajeCopiar}</code>`;
          enviarMensajeTelegram(botToken, chatId, alertMsg);
        });
      }
      if (isCallback) {
        responderCallbackQuery(botToken, callbackQueryId);
      }
      return;
    }

    // Si es un comando de cobro: /cobro_XXXXX o callback cobro_XXXXX o callback_query
    if (text.indexOf('/cobro_') === 0 || text.indexOf('cobro_') === 0) {
      const imei = text.startsWith('/cobro_') ? text.substring(7) : text.substring(6);
      const msg = generarMensajeCobroPorImei(ss, imei);
      
      const keyboard = {
        inline_keyboard: [
          [
            { text: "📋 Ver Clientes", callback_data: "ver_todos_clientes" },
            { text: "🚦 Ver Morosos", callback_data: "ver_clientes_morosos" }
          ]
        ]
      };
      
      enviarMensajeTelegramConTeclado(botToken, chatId, msg, keyboard);
      
      if (isCallback) {
        responderCallbackQuery(botToken, callbackQueryId);
      }
      return;
    }

    // Comando desconocido
    let unknownMsg = `❓ Comando no reconocido. Usa /clientes para listar los clientes.`;
    enviarMensajeTelegram(botToken, chatId, unknownMsg);
  } catch (error) {
    Logger.log("Error en procesarWebhookTelegram: " + error.toString());
    if (botToken && chatId) {
      enviarMensajeTelegram(botToken, chatId, "⚠️ Error procesando comando: " + error.toString());
    }
  }
}

function generarMensajeCobroPorImei(ss, imeiKey) {
  const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
  const filaConfig = sheetConfig ? sheetConfig.getRange("A2:L2").getValues()[0] : [6, 24, 3, 66, 90, 15, 16.5, 33, 11, 15, "", ""];
  const interesTasa = parseFloat(filaConfig[0]) || 6;

  const sheetClientes = ss.getSheetByName('Clientes');
  const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];
  
  const sheetPagos = ss.getSheetByName('Pagos');
  const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
  
  // Calcular total pagado
  let totalPagado = 0;
  for (let j = 1; j < dataPagos.length; j++) {
    if (String(dataPagos[j][0]).trim() === imeiKey) {
      totalPagado += parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
    }
  }
  
  // Buscar cliente
  let clienteRow = null;
  for (let i = 1; i < dataClientes.length; i++) {
    if (String(dataClientes[i][5] || '').trim() === imeiKey) {
      clienteRow = dataClientes[i];
      break;
    }
  }
  
  if (!clienteRow) {
    return "❌ Cliente no encontrado.";
  }
  
  const startTimestamp = clienteRow[0];
  const clienteNombre = String(clienteRow[1] || '').trim();
  const cuotaNum = parseFloat(String(clienteRow[10] || "0").replace(/[\$,]/g, "")) || 0;
  const totalFinanciadoVal = parseFloat(String(clienteRow[15] || "0").replace(/[\$,]/g, "")) || 0;
  const tipoPeriodo = String(clienteRow[20] || "SEMANAL").toUpperCase();
  
  const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);
  
  // Calcular atraso
  let diasAtrasoVal = 0;
  if (startTimestamp instanceof Date && cuotaNum > 0 && saldoPendienteBase > 0) {
    const diasTranscurridos = Math.floor((new Date() - startTimestamp) / (1000 * 60 * 60 * 24));
    const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
    const periodosTranscurridos = Math.floor(diasTranscurridos / diasPorPeriodo);
    const totalEsperado = periodosTranscurridos * cuotaNum;
    if (totalPagado < totalEsperado) {
      const atrasoMonto = totalEsperado - totalPagado;
      diasAtrasoVal = Math.ceil((atrasoMonto / cuotaNum) * diasPorPeriodo);
    }
  }
  
  let semanasAtrasoVal = Math.floor(diasAtrasoVal / 7);
  let penalidadMonto = 0;
  if (semanasAtrasoVal > 0) {
    penalidadMonto = semanasAtrasoVal * (cuotaNum * (interesTasa / 100));
  }
  const saldoPendienteReal = saldoPendienteBase + penalidadMonto;
  
  let mensajeCobro = "";
  if (diasAtrasoVal > 0) {
    mensajeCobro = `¡Hola, ${clienteNombre}! 👋\n` +
                   `Te recordamos que presentas un atraso en tu pago semanal.\n\n` +
                   `💵 *Monto de tu cuota:* $${Math.round(cuotaNum).toLocaleString('es-MX')} MXN\n` +
                   `📉 *Saldo pendiente total:* $${Math.round(saldoPendienteReal).toLocaleString('es-MX')} MXN\n` +
                   `⏳ *Días de atraso:* ${diasAtrasoVal} días\n\n` +
                   `🔗 *Consulta tu estado de cuenta:* https://portal.estrenacelya.com/?imei=${imeiKey}\n\n`;
  } else {
    mensajeCobro = `¡Hola, ${clienteNombre}! 👋\n` +
                   `Te recordamos realizar tu pago de esta semana.\n\n` +
                   `💵 *Monto de tu cuota:* $${Math.round(cuotaNum).toLocaleString('es-MX')} MXN\n` +
                   `📉 *Saldo pendiente total:* $${Math.round(saldoPendienteReal).toLocaleString('es-MX')} MXN\n\n` +
                   `🔗 *Consulta tu estado de cuenta:* https://portal.estrenacelya.com/?imei=${imeiKey}\n\n`;
  }
  
  // Agregar métodos de pago desde propiedades del script si están configurados
  let metodosPago = "";
  try {
    metodosPago = PropertiesService.getScriptProperties().getProperty('METODOS_PAGO') || "";
    // Convertir saltos de línea de texto literales (\n) a saltos de línea reales
    metodosPago = metodosPago.replace(/\\n/g, '\n').replace(/\\r/g, '\r');
  } catch(e) {
    Logger.log("Error leyendo METODOS_PAGO: " + e.toString());
  }
  
  if (metodosPago) {
    mensajeCobro += `💳 *Métodos de pago:*\n${metodosPago}\n\n`;
  }
  
  mensajeCobro += `Por favor, envíanos el comprobante una vez realizado el pago. ¡Muchas gracias!`;
                        
  let msg = `👤 <b>Cliente:</b> ${clienteNombre}\n`;
  msg += `📱 <b>IMEI:</b> <code>${imeiKey}</code>\n`;
  msg += `💵 <b>Saldo Pendiente Real:</b> $${Math.round(saldoPendienteReal).toLocaleString('es-MX')}\n`;
  msg += `⏳ <b>Atraso:</b> ${diasAtrasoVal} días (${semanasAtrasoVal} sem)\n\n`;
  msg += `📝 <b>Mensaje de Cobro para copiar:</b>\n`;
  msg += `<code>${mensajeCobro}</code>`;
  
  return msg;
}

function enviarListaClientesTelegram(ss, botToken, chatId, soloMorosos) {
  const sheetClientes = ss.getSheetByName('Clientes');
  const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];
  
  const sheetPagos = ss.getSheetByName('Pagos');
  const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
  
  // Agrupar pagos por IMEI
  const pagosPorImei = {};
  for (let j = 1; j < dataPagos.length; j++) {
    const imeiPago = String(dataPagos[j][0]).trim();
    const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
    if (imeiPago) {
      pagosPorImei[imeiPago] = (pagosPorImei[imeiPago] || 0) + montoPago;
    }
  }

  let msg = soloMorosos ? `🚦 <b>Clientes con Atraso (Morosos)</b>\n\n` : `📋 <b>Lista de Clientes Activos</b>\n\n`;
  let inlineButtons = [];

  for (let i = 1; i < dataClientes.length; i++) {
    const startTimestamp = dataClientes[i][0];
    const clienteNombre = String(dataClientes[i][1] || '').trim();
    const imeiKey = String(dataClientes[i][5] || '').trim();
    const cuotaNum = parseFloat(String(dataClientes[i][10] || "0").replace(/[\$,]/g, "")) || 0;
    const totalFinanciadoVal = parseFloat(String(dataClientes[i][15] || "0").replace(/[\$,]/g, "")) || 0;
    const tipoPeriodo = String(dataClientes[i][20] || "SEMANAL").toUpperCase();
    
    if (clienteNombre && imeiKey) {
      const totalPagado = pagosPorImei[imeiKey] || 0;
      const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);
      
      // Calcular atraso
      let diasAtrasoVal = 0;
      if (startTimestamp instanceof Date && cuotaNum > 0 && saldoPendienteBase > 0) {
        const diasTranscurridos = Math.floor((new Date() - startTimestamp) / (1000 * 60 * 60 * 24));
        const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
        const periodosTranscurridos = Math.floor(diasTranscurridos / diasPorPeriodo);
        const totalEsperado = periodosTranscurridos * cuotaNum;
        if (totalPagado < totalEsperado) {
          const atrasoMonto = totalEsperado - totalPagado;
          diasAtrasoVal = Math.ceil((atrasoMonto / cuotaNum) * diasPorPeriodo);
        }
      }
      
      if (soloMorosos && diasAtrasoVal === 0) {
        continue;
      }
      
      const statusEmoji = diasAtrasoVal === 0 ? "🟢" : (diasAtrasoVal <= 7 ? "🟡" : "🔴");
      const shortDesc = `${statusEmoji} ${clienteNombre} (${diasAtrasoVal}d)`;
      
      // Añadir botón en línea
      inlineButtons.push([
        { text: shortDesc, callback_data: `cobro_${imeiKey}` }
      ]);
      
      // Comando clickable alternativo
      msg += `• ${statusEmoji} ${clienteNombre} (${diasAtrasoVal}d): /cobro_${imeiKey}\n`;
    }
  }

  if (inlineButtons.length === 0) {
    msg += soloMorosos ? `✅ No hay clientes morosos registrados actualmente.` : `No hay clientes activos registrados.`;
    enviarMensajeTelegram(botToken, chatId, msg);
  } else {
    msg += `\n👉 <i>Toca en los botones de abajo o haz clic en los comandos /cobro_[imei] del texto para ver su plantilla.</i>`;
    const keyboard = {
      inline_keyboard: inlineButtons
    };
    enviarMensajeTelegramConTeclado(botToken, chatId, msg, keyboard);
  }
}

function enviarMensajeTelegramConTeclado(botToken, chatId, mensaje, teclado) {
  if (!botToken || !chatId) return;
  
  const url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  const payload = {
    chat_id: chatId,
    text: mensaje,
    parse_mode: "HTML",
    reply_markup: teclado
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200) {
      const errText = response.getContentText();
      Logger.log("Error al enviar Telegram con teclado. Código: " + code + ", Respuesta: " + errText);
      let desc = "Error desconocido";
      try {
        desc = JSON.parse(errText).description || desc;
      } catch (pe) {}
      throw new Error("Telegram API Error (" + code + "): " + desc);
    }
  } catch (e) {
    Logger.log("Error al enviar Telegram con teclado: " + e.toString());
    throw e;
  }
}

function responderCallbackQuery(botToken, callbackQueryId) {
  if (!botToken || !callbackQueryId) return;
  const url = "https://api.telegram.org/bot" + botToken + "/answerCallbackQuery";
  const payload = {
    callback_query_id: callbackQueryId
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("Error respondiendo callbackQuery: " + e.toString());
  }
}

// 🤖 CÁLCULO DE ALERTAS POR RETRASO EN HORAS / SEMANAS Y GENERACIÓN DE PLANTILLAS 🤖
function obtenerAlertasClientes(ss) {
  const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
  const filaConfig = sheetConfig ? sheetConfig.getRange("A2:L2").getValues()[0] : [6, 24, 3, 66, 90, 15, 16.5, 33, 11, 15, "", ""];
  
  const tasaInteresGlobal = parseFloat(filaConfig[0]) || 6;
  const horasToleranciaGlobal = parseInt(filaConfig[1]) || 24;
  const semanasToleranciaGlobal = parseInt(filaConfig[2]) || 3;
  
  const sheetClientes = ss.getSheetByName('Clientes');
  const dataClientes = sheetClientes ? sheetClientes.getDataRange().getValues() : [];
  
  const sheetPagos = ss.getSheetByName('Pagos');
  const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
  
  // Agrupar pagos por IMEI
  const pagosPorImei = {};
  for (let j = 1; j < dataPagos.length; j++) {
    const imeiPago = String(dataPagos[j][0]).trim();
    const montoPago = parseFloat(String(dataPagos[j][3]).replace(/[\$,]/g, "")) || 0;
    if (imeiPago) {
      pagosPorImei[imeiPago] = (pagosPorImei[imeiPago] || 0) + montoPago;
    }
  }
  
  let metodosPago = "";
  try {
    metodosPago = PropertiesService.getScriptProperties().getProperty('METODOS_PAGO') || "";
    metodosPago = metodosPago.replace(/\\n/g, '\n').replace(/\\r/g, '\r');
  } catch(e) {}
  
  let alertasBloqueo = [];
  let alertasRescision = [];
  
  for (let i = 1; i < dataClientes.length; i++) {
    const startTimestamp = dataClientes[i][0];
    const clienteNombre = String(dataClientes[i][1] || '').trim();
    const modelo = String(dataClientes[i][4] || 'Genérico').trim();
    const imeiKey = String(dataClientes[i][5] || '').trim();
    const cuotaNum = parseFloat(String(dataClientes[i][10] || "0").replace(/[\$,]/g, "")) || 0;
    const totalFinanciadoVal = parseFloat(String(dataClientes[i][15] || "0").replace(/[\$,]/g, "")) || 0;
    const tipoPeriodo = String(dataClientes[i][20] || "SEMANAL").toUpperCase();
    
    if (clienteNombre && imeiKey && totalFinanciadoVal > 0) {
      const totalPagado = pagosPorImei[imeiKey] || 0;
      const saldoPendienteBase = Math.max(0, totalFinanciadoVal - totalPagado);
      
      if (saldoPendienteBase <= 0) continue; // Ya liquidó, no hay deuda
      
      let fechaInicio = startTimestamp;
      if (fechaInicio && !(fechaInicio instanceof Date)) {
        fechaInicio = new Date(fechaInicio);
      }
      
      let horasTol = horasToleranciaGlobal;
      const colMVal = dataClientes[i][12];
      if (colMVal !== undefined && colMVal !== null && String(colMVal).trim() !== "") {
        const cleanM = String(colMVal).replace(/[^0-9]/g, "");
        if (cleanM) horasTol = parseInt(cleanM) || horasToleranciaGlobal;
      }
      
      let semTol = semanasToleranciaGlobal;
      const colNVal = dataClientes[i][13];
      if (colNVal !== undefined && colNVal !== null && String(colNVal).trim() !== "") {
        const cleanN = String(colNVal).replace(/[^0-9]/g, "");
        if (cleanN) semTol = parseInt(cleanN) || semanasToleranciaGlobal;
      }
      
      let tasaInteresCliente = tasaInteresGlobal;
      const colLVal = dataClientes[i][11];
      if (colLVal !== undefined && colLVal !== null && String(colLVal).trim() !== "") {
        const cleanL = String(colLVal).replace(/[^0-9\.]/g, "");
        if (cleanL) tasaInteresCliente = parseFloat(cleanL) || tasaInteresGlobal;
      }
      
      let horasAtraso = 0;
      let semanasAtraso = 0;
      let hasValidDate = fechaInicio instanceof Date && !isNaN(fechaInicio.getTime());
      
      if (hasValidDate && cuotaNum > 0) {
        const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
        const numPeriodosPagados = Math.floor(totalPagado / cuotaNum);
        const proximoVencimiento = new Date(fechaInicio.getTime() + (numPeriodosPagados + 1) * diasPorPeriodo * 24 * 60 * 60 * 1000);
        
        const ahora = new Date();
        if (ahora > proximoVencimiento) {
          const diffMs = ahora.getTime() - proximoVencimiento.getTime();
          horasAtraso = Math.floor(diffMs / (1000 * 60 * 60));
          semanasAtraso = Math.floor(horasAtraso / (24 * 7));
        }
      }
      
      // Calcular penalidad e intereses
      let penalidadMonto = 0;
      const diasAtraso = Math.max(0, Math.floor(horasAtraso / 24));
      if (diasAtraso > 0) {
        penalidadMonto = diasAtraso * (cuotaNum * (tasaInteresCliente / 100));
      }
      const saldoPendienteReal = saldoPendienteBase + penalidadMonto;
      
      let fVencimientoStr = "N/A";
      if (hasValidDate) {
        const diasPorPeriodo = (tipoPeriodo === "QUINCENAL") ? 15 : 7;
        const numPeriodosPagados = Math.floor(totalPagado / cuotaNum);
        const vencimientoDate = new Date(fechaInicio.getTime() + (numPeriodosPagados + 1) * diasPorPeriodo * 24 * 60 * 60 * 1000);
        fVencimientoStr = Utilities.formatDate(vencimientoDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      }
      
      const infoCliente = {
        cliente: clienteNombre,
        modelo: modelo,
        imei: imeiKey,
        cuota: Math.round(cuotaNum),
        saldo: Math.round(saldoPendienteReal),
        horasAtraso: horasAtraso,
        semanasAtraso: semanasAtraso,
        fechaVencimiento: fVencimientoStr
      };
      
      if (semanasAtraso >= semTol) {
        // Alerta de Rescisión
        let msgRescisionTemplate = `⚠️ AVISO DE RESCISIÓN DE CONTRATO ⚠️\n\n` +
                                   `Hola ${clienteNombre}, lamentamos informarte que, al haber acumulado ${semanasAtraso} semanas de atraso en tus pagos de tu equipo ${modelo} con IMEI ${imeiKey}, se ha procedido a la rescisión de tu contrato de financiamiento.\n\n` +
                                   `💵 Cuota: $${Math.round(cuotaNum).toLocaleString('es-MX')} MXN\n` +
                                   `📉 Saldo vencido acumulado: $${Math.round(saldoPendienteReal).toLocaleString('es-MX')} MXN\n\n` +
                                   `Te solicitamos realizar el pago total del saldo vencido de inmediato o comunicarte para la devolución del equipo y evitar el inicio de las acciones legales correspondientes.\n\n` +
                                   `✍️ Firma tu rescisión aquí: https://portal.estrenacelya.com/?imei=${imeiKey}\n\n`;
        if (metodosPago) {
          msgRescisionTemplate += `💳 Métodos de pago:\n${metodosPago}\n\n`;
        }
        msgRescisionTemplate += `Quedamos al pendiente.`;
        
        infoCliente.mensajeCopiar = msgRescisionTemplate;
        alertasRescision.push(infoCliente);
      } else if (horasAtraso >= horasTol) {
        // Alerta de Bloqueo
        let msgBlockTemplate = `🚨 *AVISO DE BLOQUEO DE EQUIPO* 🚨\n\n` +
                               `Hola *${clienteNombre}*, te informamos que tu equipo *${modelo}* con IMEI *${imeiKey}* presenta un atraso de *${horasAtraso}* horas en su pago.\n\n` +
                               `💵 *Cuota:* $${Math.round(cuotaNum).toLocaleString('es-MX')} MXN\n` +
                               `📉 *Total para regularizar:* $${Math.round(saldoPendienteReal).toLocaleString('es-MX')} MXN\n\n` +
                               `Te recordamos que tu fecha límite de pago fue el *${fVencimientoStr}*. Para restablecer el acceso a tu equipo, realiza tu pago y envía tu comprobante a la brevedad.\n\n` +
                               `🔗 *Consulta tu estado de cuenta:* https://portal.estrenacelya.com/?imei=${imeiKey}\n\n`;
        if (metodosPago) {
          msgBlockTemplate += `💳 *Métodos de pago:*\n${metodosPago}\n\n`;
        }
        msgBlockTemplate += `Por favor, envíanos el comprobante una vez realizado el pago. ¡Muchas gracias!`;
        
        infoCliente.mensajeCopiar = msgBlockTemplate;
        alertasBloqueo.push(infoCliente);
      }
    }
  }
  
  return {
    bloqueos: alertasBloqueo,
    rescisiones: alertasRescision
  };
}

// 🔧 Función utilitaria para inicializar la propiedad del webhook de n8n
function configurarWebhookN8N() {
  PropertiesService.getScriptProperties().setProperty("N8N_CONTRATO_WEBHOOK", "http://107.175.122.33:5678/webhook/3/webhook/contrato-pagare");
  console.log("Propiedad N8N_CONTRATO_WEBHOOK configurada en Properties Service.");
}

// 🔄 Sincroniza todos los precios del catálogo con Mercado Libre utilizando el token oficial (en paralelo)
function sincronizarPreciosCatalogoCompletoML() {
  const credenciales = obtenerPropiedadesEcosistema();
  const ss = SpreadsheetApp.openById(credenciales.SHEET_ID);
  const sheetCat = ss.getSheetByName('Catalogo');
  if (!sheetCat) {
    return "No se encontró la pestaña 'Catalogo'";
  }
  
  const ultFila = sheetCat.getLastRow();
  if (ultFila < 2) return "Catálogo vacío";
  
  const dataCat = sheetCat.getRange(2, 1, ultFila - 1, sheetCat.getLastColumn()).getValues();
  const formulasLinks = sheetCat.getRange(2, 7, ultFila - 1, 1).getFormulas();
  const linksRicos = sheetCat.getRange(2, 7, ultFila - 1, 1).getRichTextValues();
  
  const tokenML = refrescarTokenML(); // Asegurar token fresco
  
  let peticiones = [];
  let mapeoItems = [];
  
  for (let j = 0; j < dataCat.length; j++) {
    const filaReal = j + 2;
    const modelo = String(dataCat[j][1] || '').trim();
    
    const formula = formulasLinks[j] && formulasLinks[j][0] ? formulasLinks[j][0] : "";
    let formulaLink = "";
    if (formula.startsWith("=HYPERLINK")) {
      const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
      if (match) formulaLink = match[1];
    }
    const richLink = linksRicos[j] && linksRicos[j][0] ? linksRicos[j][0].getLinkUrl() : "";
    const link = richLink || formulaLink || String(dataCat[j][6] || '').trim();
    
    if (link && link.indexOf("mercadolibre") !== -1) {
      const regex = /(MLM\-?\d+)/i;
      const coincidencia = link.match(regex);
      if (coincidencia) {
        const id = coincidencia[0].replace("-", "").toUpperCase();
        let targetUrl = `https://api.mercadolibre.com/items/${id}`;
        let esProducto = false;
        
        if (link.includes("/p/")) {
          // Para catálogo (/p/), consultamos la lista de publicaciones activas para obtener precios y variantes con stock real
          targetUrl = `https://api.mercadolibre.com/products/${id}/items`;
          esProducto = true;
        }
        
        peticiones.push({
          url: targetUrl,
          method: "get",
          headers: {
            "Authorization": "Bearer " + tokenML
          },
          muteHttpExceptions: true
        });
        
        mapeoItems.push({
          filaReal: filaReal,
          modelo: modelo,
          esProducto: esProducto,
          id: id,
          colorOriginal: String(dataCat[j][9] || '').trim()
        });
      }
    }
  }
  
  if (peticiones.length === 0) {
    return { result: "success", actualizados: 0, mensaje: "No hay links de Mercado Libre que procesar" };
  }
  
  // ⚡ EJECUCIÓN CONCURRENTE EN LA RED DE GOOGLE (Tarda ~1-2 segundos en total)
  const respuestas = UrlFetchApp.fetchAll(peticiones);
  let contadorActualizados = 0;
  let errores = [];
  
  for (let k = 0; k < respuestas.length; k++) {
    const res = respuestas[k];
    const info = mapeoItems[k];
    
    if (res.getResponseCode() === 200) {
      try {
        const resJson = JSON.parse(res.getContentText());
        let precio = 0;
        
        if (info.esProducto) {
          // resJson contiene {"results": [...]} de la API de items de catálogo
          const results = resJson.results || [];
          if (results.length > 0) {
            results.sort(function(a, b) { return (a.price || 0) - (b.price || 0); });
            precio = parseFloat(results[0].price) || 0;
          }
        } else {
          precio = parseFloat(resJson.price) || 0;
        }
        
        if (precio > 0) {
          sheetCat.getRange(info.filaReal, 4).setValue(Math.round(precio));
          
          // Guardar variantes JSON en la columna 11 (K) y colores en la columna 13 (M) de forma automática
          try {
            let jsonString = "";
            if (info.esProducto && resJson.results && Array.isArray(resJson.results)) {
              // Quedarse solo con los primeros 10 items (las mejores ofertas)
              const depurado = {
                results: resJson.results.slice(0, 10)
              };
              jsonString = JSON.stringify(depurado);
            } else {
              jsonString = JSON.stringify(resJson);
            }
            
            // Si por alguna razón sigue siendo demasiado largo, recortar el array
            if (jsonString.length > 48000) {
              if (info.esProducto && resJson.results && Array.isArray(resJson.results)) {
                const depurado = {
                  results: resJson.results.slice(0, 3)
                };
                jsonString = JSON.stringify(depurado);
              }
            }
            
            sheetCat.getRange(info.filaReal, 11).setValue(jsonString);
            const colDisponibles = extraerColoresDeDatosML(resJson);
            sheetCat.getRange(info.filaReal, 13).setValue(colDisponibles || "");
          } catch(eVar) {
            Logger.log("Error guardando JSON o colores en sincronización horaria: " + eVar.toString());
          }
          
          contadorActualizados++;
        } else {
          errores.push("Fila " + info.filaReal + " (" + info.modelo + "): Precio no encontrado en resultados");
        }
      } catch (parseErr) {
        errores.push("Fila " + info.filaReal + " (" + info.modelo + "): Error de parsing de JSON");
      }
    } else {
      errores.push("Fila " + info.filaReal + " (" + info.modelo + "): HTTP " + res.getResponseCode() + " - " + res.getContentText().substring(0, 100));
    }
  }
  
  SpreadsheetApp.flush();
  
  const logMsg = "🔄 Sincronización Catálogo Completo: " + contadorActualizados + " precios de Mercado Libre actualizados con éxito.";
  escribirLogDebug(logMsg);
  if (errores.length > 0) {
    escribirLogDebug("⚠️ Errores en sincronización: " + errores.join(" | "));
  }
  
  return {
    result: "success",
    actualizados: contadorActualizados,
    errores: errores
  };
}

// ⏰ Crea el activador para que Google Sheets actualice el catálogo solo cada hora
function crearTriggerSincronizacionCatalogo() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'ejecutarSincronizacionHorariaML') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('ejecutarSincronizacionHorariaML')
           .timeBased()
           .everyHours(1)
           .create();
  
  console.log("⏰ Activador horario para sincronización de catálogo configurado con éxito.");
}

// Función que ejecuta el activador horario
function ejecutarSincronizacionHorariaML() {
  // 1. Sincronizar precios y variantes
  sincronizarPreciosCatalogoCompletoML();
  
  // 2. Ejecutar verificación de stock y notificar a Telegram si hay productos agotados
  try {
    const alertas = verificarStockGlobal();
    const productosSinStock = alertas.filter(function(item) { return !item.tieneStock; });
    
    if (productosSinStock.length > 0) {
      const creds = obtenerPropiedadesEcosistema();
      let ss = null;
      try {
        ss = SpreadsheetApp.getActiveSpreadsheet();
      } catch(e) {}
      if (!ss && creds.SHEET_ID) {
        ss = SpreadsheetApp.openById(creds.SHEET_ID.trim());
      }
      
      if (ss) {
        const sheetConfig = ss.getSheetByName('Configuración') || ss.getSheetByName('Configuracion');
        if (sheetConfig) {
          const filaConfig = sheetConfig.getRange("A2:L2").getValues()[0];
          const botToken = String(filaConfig[10] || '').trim();
          const chatId = String(filaConfig[11] || '').trim();
          
          if (botToken && chatId) {
            let msg = `⚠️ <b>Alerta de Quiebre de Stock CelYa</b> ⚠️\n\n`;
            msg += `Se detectaron <b>${productosSinStock.length}</b> celulares activos en tu catálogo sin stock en Mercado Libre:\n\n`;
            
            productosSinStock.forEach(function(item) {
              msg += `• <b>${item.modelo}</b> (Fila ${item.filaNum})\n`;
              msg += `  ❌ Motivo: <i>${item.mensaje}</i>\n`;
              msg += `  🔗 Link: <a href="${item.link}">Ver publicación</a>\n\n`;
            });
            
            msg += `💡 <i>Tip: Puedes pausar estos productos en la columna 'Estatus' de tu Google Sheet para que no aparezcan en el catálogo de los clientes.</i>`;
            enviarMensajeTelegram(botToken, chatId, msg);
          }
        }
      }
    }
  } catch (errStock) {
    Logger.log("Error al verificar stock o notificar Telegram en sincronización horaria: " + errStock.toString());
  }
}

// 📦 Verifica el stock en vivo de todos los celulares activos en paralelo y guarda el caché
function verificarStockGlobal() {
  const credenciales = obtenerPropiedadesEcosistema();
  const ss = SpreadsheetApp.openById(credenciales.SHEET_ID);
  const sheetCat = ss.getSheetByName('Catalogo');
  if (!sheetCat) return [];
  
  const ultFila = sheetCat.getLastRow();
  if (ultFila < 2) return [];
  
  const dataCat = sheetCat.getRange(2, 1, ultFila - 1, sheetCat.getLastColumn()).getValues();
  const formulasLinks = sheetCat.getRange(2, 7, ultFila - 1, 2).getFormulas();
  const linksRicos = sheetCat.getRange(2, 7, ultFila - 1, 2).getRichTextValues();
  
  const tokenML = refrescarTokenML();
  
  let peticiones = [];
  let mapeo = [];
  
  for (let j = 0; j < dataCat.length; j++) {
    const filaReal = j + 2;
    const modelo = String(dataCat[j][1] || '').trim();
    const estado = String(dataCat[j][5] || '').toUpperCase().trim();
    
    // Sólo evaluar stock de productos activos
    if (!["DISPONIBLE", "MAS VENDIDO", "BAJO DEMANDA"].includes(estado)) continue;
    
    const formula = formulasLinks[j] && formulasLinks[j][0] ? formulasLinks[j][0] : "";
    let formulaLink = "";
    if (formula.startsWith("=HYPERLINK")) {
      const match = formula.match(/HYPERLINK\("([^"]+)"/i) || formula.match(/HYPERLINK\('([^']+)'/i);
      if (match) formulaLink = match[1];
    }
    const richLink = linksRicos[j] && linksRicos[j][0] ? linksRicos[j][0].getLinkUrl() : "";
    const link = richLink || formulaLink || String(dataCat[j][6] || '').trim();
    
    // Extraer foto
    const formulaFoto = formulasLinks[j] && formulasLinks[j][1] ? formulasLinks[j][1] : "";
    let formulaFotoLink = "";
    if (formulaFoto.startsWith("=HYPERLINK")) {
      const matchFoto = formulaFoto.match(/HYPERLINK\("([^"]+)"/i) || formulaFoto.match(/HYPERLINK\('([^']+)'/i);
      if (matchFoto) formulaFotoLink = matchFoto[1];
    }
    const richFotoLink = linksRicos[j] && linksRicos[j][1] ? linksRicos[j][1].getLinkUrl() : "";
    let foto = richFotoLink || formulaFotoLink || String(dataCat[j][7] || '').trim();
    if (foto.toLowerCase() === "foto") foto = "";
    
    if (link && link.indexOf("mercadolibre") !== -1) {
      const regex = /(MLM\-?\d+)/i;
      const coincidencia = link.match(regex);
      if (coincidencia) {
        const id = coincidencia[0].replace("-", "").toUpperCase();
        let targetUrl = `https://api.mercadolibre.com/items/${id}`;
        let esProducto = false;
        
        if (link.includes("/p/")) {
          targetUrl = `https://api.mercadolibre.com/products/${id}`;
          esProducto = true;
        }
        
        peticiones.push({
          url: targetUrl,
          method: "get",
          headers: {
            "Authorization": "Bearer " + tokenML
          },
          muteHttpExceptions: true
        });
        
        mapeo.push({
          filaReal: filaReal,
          modelo: modelo,
          esProducto: esProducto,
          id: id,
          link: link,
          estado: estado,
          variantesJsonActual: dataCat[j][10] ? String(dataCat[j][10]).trim() : "",
          foto: foto
        });
      }
    }
  }
  
  if (peticiones.length === 0) return [];
  
  const respuestas = UrlFetchApp.fetchAll(peticiones);
  let resultadosStock = [];
  let actualizacionesCache = [];
  
  for (let k = 0; k < respuestas.length; k++) {
    const res = respuestas[k];
    const info = mapeo[k];
    let liveData = null;
    const responseCode = res.getResponseCode();
    
    if (responseCode === 200) {
      try {
        liveData = JSON.parse(res.getContentText());
      } catch (e) {}
    }
    
    let tieneStock = false;
    let razon = "";
    let mensaje = "";
    let dataLimpia = {};
    
    if (responseCode === 200 && liveData) {
      dataLimpia = limpiarJsonVariantesML(liveData);
      
      let fotoApi = "";
      if (info.esProducto) {
        fotoApi = liveData.pictures && liveData.pictures.length > 0 ? (liveData.pictures[0].url || liveData.pictures[0].secure_url || "") : (liveData.thumbnail || "");
        
        let activeCount = 0;
        if (liveData.pickers && liveData.pickers.length > 0) {
          const productMap = {};
          liveData.pickers.forEach(picker => {
            if (!picker.products) return;
            picker.products.forEach(p => {
              if (!p.product_id) return;
              const tags = p.tags || [];
              const outOfStock = tags.includes("out-of-stock") || tags.includes("disabled") || tags.includes("no-bids") || tags.includes("no-winner");
              if (!productMap[p.product_id]) {
                productMap[p.product_id] = { isOutOfStock: outOfStock };
              } else if (outOfStock) {
                productMap[p.product_id].isOutOfStock = true;
              }
            });
          });
          const allProductIds = Object.keys(productMap);
          const activeProducts = allProductIds.filter(pid => !productMap[pid].isOutOfStock);
          activeCount = activeProducts.length;
        } else {
          activeCount = (liveData.price || (liveData.buy_box_winner && liveData.buy_box_winner.price)) ? 1 : 0;
        }
        
        if (activeCount > 0) {
          tieneStock = true;
        } else {
          tieneStock = false;
          razon = "all_pickers_out";
          mensaje = "Variantes de catálogo agotadas";
        }
      } else {
        fotoApi = liveData.thumbnail || (liveData.pictures && liveData.pictures.length > 0 ? liveData.pictures[0].url : "");
        
        if (liveData.variations && liveData.variations.length > 0) {
          const activeVariations = liveData.variations.filter(v => parseInt(v.available_quantity, 10) > 0);
          if (activeVariations.length > 0) {
            tieneStock = true;
          } else {
            tieneStock = false;
            razon = "all_variations_out";
            mensaje = `${liveData.variations.length} variaciones, todas sin stock`;
          }
        } else {
          const status = (liveData.status || "").toLowerCase();
          const qty = parseInt(liveData.available_quantity, 10);
          if (status && status !== "active") {
            tieneStock = false;
            razon = "inactive";
            mensaje = `Estado: ${liveData.status || 'inactivo'}`;
          } else if (!isNaN(qty) && qty > 0) {
            tieneStock = true;
          } else {
            tieneStock = false;
            razon = "no_stock";
            mensaje = "Sin unidades disponibles";
          }
        }
      }
      
      // Sanitizar la foto para HTTPS y mejor tamaño
      if (fotoApi) {
        if (fotoApi.startsWith("http:")) {
          fotoApi = fotoApi.replace("http:", "https:");
        }
        fotoApi = fotoApi.replace("-I.jpg", "-O.jpg");
      }
      
      // Auto-completar la foto en Sheets si la celda estaba vacía
      if ((!info.foto || info.foto.trim() === "") && fotoApi) {
        sheetCat.getRange(info.filaReal, 8).setValue(`=HYPERLINK("${fotoApi}", "FOTO")`);
        info.foto = fotoApi;
      }
      
      const cacheStr = JSON.stringify(dataLimpia);
      if (cacheStr !== info.variantesJsonActual) {
        actualizacionesCache.push({
          fila: info.filaReal,
          variantes: cacheStr
        });
      }
    } else {
      tieneStock = false;
      razon = "inactive";
      if (responseCode === 404) {
        mensaje = "Publicación pausada o sin stock en Mercado Libre (404/No Winners)";
      } else {
        mensaje = "Error de conexión con Mercado Libre (HTTP " + responseCode + ")";
      }
    }
    
    resultadosStock.push({
      filaNum: info.filaReal,
      modelo: info.modelo,
      link: info.link,
      id: info.id,
      estado: info.estado,
      tieneStock: tieneStock,
      razon: razon,
      mensaje: mensaje,
      liveData: dataLimpia,
      foto: info.foto
    });
  }
  
  if (actualizacionesCache.length > 0) {
    actualizacionesCache.forEach(upd => {
      sheetCat.getRange(upd.fila, 11).setValue(upd.variantes);
    });
    SpreadsheetApp.flush();
  }
  
  return resultadosStock;
}

// 🧹 Recorta y limpia el JSON de Mercado Libre para no saturar las celdas de Sheets
function limpiarJsonVariantesML(data) {
  if (!data) return {};
  
  // 1. Si es respuesta de /products/{id} (producto de catálogo con pickers)
  if (data.pickers && Array.isArray(data.pickers)) {
    return {
      id: data.id,
      name: data.name,
      price: data.price,
      buy_box_winner: data.buy_box_winner ? {
        item_id: data.buy_box_winner.item_id || data.buy_box_winner.id,
        price: data.buy_box_winner.price,
        condition: data.buy_box_winner.condition
      } : null,
      pictures: (data.pictures || []).slice(0, 1).map(p => { return { url: p.url || p.secure_url }; }),
      pickers: data.pickers.map(p => {
        return {
          picker_id: p.picker_id,
          products: (p.products || []).map(prod => {
            return {
              product_id: prod.product_id,
              picker_label: prod.picker_label,
              product_name: prod.product_name || prod.product_title || "",
              tags: prod.tags || []
            };
          })
        };
      })
    };
  }
  
  // 2. Si es respuesta de /products/{id}/items (lista de ofertas para catálogo)
  if (data.results && Array.isArray(data.results)) {
    return {
      results: data.results.map(item => {
        return {
          id: item.id,
          price: item.price,
          available_quantity: item.available_quantity,
          attributes: (item.attributes || []).map(attr => {
            return {
              id: attr.id,
              name: attr.name,
              value_name: attr.value_name
            };
          })
        };
      })
    };
  }
  
  // 3. Si es respuesta de /items/{id} (ítem con variaciones)
  if (data.variations && Array.isArray(data.variations)) {
    return {
      id: data.id,
      status: data.status,
      price: data.price,
      available_quantity: data.available_quantity,
      variations: data.variations.map(v => {
        return {
          id: v.id,
          price: v.price,
          available_quantity: v.available_quantity,
          attribute_combinations: (v.attribute_combinations || []).map(comb => {
            return {
              id: comb.id,
              name: comb.name,
              value_name: comb.value_name
            };
          })
        };
      })
    };
  }
  
  // 4. Si es respuesta de /items/{id} (ítem simple)
  return {
    id: data.id,
    status: data.status,
    price: data.price,
    available_quantity: data.available_quantity,
    pictures: (data.pictures || []).slice(0, 1).map(p => { return { url: p.url }; })
  };
}

// ⚡ Limpia automáticamente el caché de una fila del catálogo si se edita el link o el modelo
function onEdit(e) {
  if (!e) return;
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    if (sheet.getName() === 'Catalogo') {
      const fila = range.getRow();
      const col = range.getColumn();
      
      // Si se edita el IMEI (1), Modelo (2) o Link de Mercado Libre (7)
      if (fila > 1 && (col === 1 || col === 2 || col === 7)) {
        // Limpiar celda de variantesJson (columna 11) y coloresDisponibles (columna 13)
        sheet.getRange(fila, 11).setValue("");
        sheet.getRange(fila, 13).setValue("");
      }
    }
  } catch (err) {
    console.error("Error en trigger onEdit:", err.toString());
  }
}