// ==========================================
// CONFIGURATION
// ==========================================

/*
function inicializarPropiedadesSeguras() {
  const props = PropertiesService.getScriptProperties();
  
  const secretos = {
    // Convertimos el array a texto (JSON) para que Apps Script permita guardarlo
    'EXCEPTIONS': JSON.stringify([
      { 
        name: "---------------", 
        spreadsheetId: "-----------------" 
      }
    ])
  };
  
  props.setProperties(secretos);
  console.log('✅ EXCEPTIONS guardados en el entorno seguro.');
}
*/

const PA_TEAM_MEMBERS = ["Krizza", "Mark", "Giovanni", "Roan", "Anthoney", "Czarina"];

/**
 * Función principal llamada desde Code.gs para evaluar e inyectar métricas de PA
 */
function getPAMetricsData(reportMode, teamName, employeeEmail, startDateStr, endDateStr, finalName) {
  const scriptProps = PropertiesService.getScriptProperties();
  
  // IMPORTANTE: Asegúrate de guardar estas propiedades en el script de Reports
  const PA_SPREADSHEET_ID = scriptProps.getProperty('MAIN_SPREADSHEET_ID');
  if (!PA_SPREADSHEET_ID) {
    Logger.log("PA_SPREADSHEET_ID no configurado en propiedades.");
    return null; 
  }
  
  const EXCEPTIONS = JSON.parse(scriptProps.getProperty('PA_EXCEPTIONS') || '[]');
  const allPAMembers = [...PA_TEAM_MEMBERS, ...EXCEPTIONS.map(e => e.name)];
  
  let selection = null;

  // 1. Verificar si se está pidiendo el reporte del Equipo PA
  if (reportMode === 'teams' && teamName) {
    let tName = teamName.toLowerCase();
    // Hacemos match si se llama PA o si incluye Prior Auth
    if (tName.includes('pa') || tName.includes('prior auth')) {
      selection = "All";
    }
  }
  // 2. Verificar si se está pidiendo el reporte de un Empleado que pertenece a PA
  else if (reportMode === 'employee') {
    // Unimos el email de Weave y el nombre del HR Sheet en una sola cadena a prueba de balas
    let searchString = ((employeeEmail || "") + " " + (finalName || "")).toLowerCase();
    
    for (let i = 0; i < allPAMembers.length; i++) {
      let memberName = allPAMembers[i].toLowerCase().trim();
      // Si la cadena combinada incluye "anna", el match es exitoso
      if (searchString.includes(memberName)) {
        selection = allPAMembers[i];
        break;
      }
    }
  }

  // Si no es un reporte de PA, retornamos nulo para no interferir con el resto del script
  if (!selection) return null;

  return processPAData(selection, startDateStr, endDateStr, PA_SPREADSHEET_ID, EXCEPTIONS);
}

/**
  Procesa la lectura del Excel y cruza las fechas
 **/
/**
  Procesa la lectura del Excel como texto puro sin zonas horarias
 **/
function processPAData(selection, startDateStr, endDateStr, PA_SPREADSHEET_ID, EXCEPTIONS) {
  // 1. ELIMINAMOS LAS CONVERSIONES. startDateStr y endDateStr ya vienen como texto "YYYY-MM-DD"
  
  let targetMembers = (selection === "All") ? [...PA_TEAM_MEMBERS, ...EXCEPTIONS.map(e => e.name)] : [selection];
  const mainSpreadsheet = SpreadsheetApp.openById(PA_SPREADSHEET_ID);
  let allRows = [];

  targetMembers.forEach(member => {
    let sheetToProcess = null;
    let isException = EXCEPTIONS.find(e => e.name === member);
    let currentSpreadsheet = mainSpreadsheet; 

    if (isException) {
      try {
        const extSs = SpreadsheetApp.openById(isException.spreadsheetId);
        currentSpreadsheet = extSs;
        const extSheets = extSs.getSheets();
        for (let i = 0; i < extSheets.length; i++) {
          if (extSheets[i].getName().toLowerCase().includes(member.toLowerCase())) {
            sheetToProcess = extSheets[i];
            break;
          }
        }
        if (!sheetToProcess) sheetToProcess = extSheets[0];
      } catch (e) {
        Logger.log("Could not open spreadsheet for exception: " + member);
      }
    } else {
      const sheets = mainSpreadsheet.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        if (sheets[i].getName().toLowerCase().includes(member.toLowerCase())) {
          sheetToProcess = sheets[i];
          break;
        }
      }
    }

    if (sheetToProcess) {
      // Tomamos la zona horaria nativa de ESTE Excel en particular (ej. Montevideo) para extraer el texto exacto que tú ves en la pantalla
      const sheetTz = currentSpreadsheet.getSpreadsheetTimeZone(); 
      const data = sheetToProcess.getDataRange().getValues();
      
      if (data.length > 1) {
        let headerRowIdx = 0;
        for(let r = 0; r < Math.min(5, data.length); r++) {
          const tempHeaders = data[r].map(h => String(h).trim().toLowerCase());
          if (tempHeaders.some(h => h.includes("start date") || h.includes("status"))) {
            headerRowIdx = r;
            break;
          }
        }

        const headers = data[headerRowIdx].map(h => String(h).trim().toLowerCase());
        const statusIdx = headers.findIndex(h => h.includes("status"));
        const dateIdx = headers.findIndex(h => h.includes("start date"));
        const medIdx = headers.findIndex(h => h.includes("medication") || h.includes("procedure"));
        const insIdx = headers.findIndex(h => h.includes("insurance") || h.includes("pharmacy plan"));

        for (let r = headerRowIdx + 1; r < data.length; r++) {
          const row = data[r];
          if (dateIdx === -1 || !row[dateIdx]) continue;
          
          // 2. CONVERTIMOS LA CELDA A TEXTO PURO
          let rowDateStr = "";
          if (row[dateIdx] instanceof Date) {
             // Si Google Sheets lo lee como fecha, extraemos el texto YYYY-MM-DD tal cual se ve en tu monitor
             rowDateStr = Utilities.formatDate(row[dateIdx], sheetTz, "yyyy-MM-dd");
          } else {
             // Si lo tipearon como texto libre
             let tempD = new Date(row[dateIdx]);
             if (!isNaN(tempD.getTime())) rowDateStr = Utilities.formatDate(tempD, sheetTz, "yyyy-MM-dd");
          }

          if (!rowDateStr) continue;

          // 3. COMPARACIÓN ALFABÉTICA (Sin cálculos matemáticos de tiempo)
          if (rowDateStr >= startDateStr && rowDateStr <= endDateStr) {
            allRows.push({
              member: member,
              status: statusIdx !== -1 ? row[statusIdx] : "Unknown",
              date: rowDateStr, // Ya es un texto fijo
              medication: medIdx !== -1 ? row[medIdx] : "Unknown",
              insurance: insIdx !== -1 ? row[insIdx] : "Unknown"
            });
          }
        }
      }
    }
  });

  let metrics = aggregatePAMetrics(allRows);
  metrics.selection = selection; 
  return metrics;
}

/**
 * Agrega y resume los datos extraídos
 */
function aggregatePAMetrics(dataRows) {
  const metrics = {
    statusCounts: {}, dailyCounts: {}, medicationCounts: {}, insuranceCounts: {}, memberCounts: {} 
  };

  dataRows.forEach(row => {
    let status = row.status ? String(row.status).trim().toUpperCase() : "BLANK";
    if (status === "COMPLETED") status = "COMPLETE";
    metrics.statusCounts[status] = (metrics.statusCounts[status] || 0) + 1;
    metrics.dailyCounts[row.date] = (metrics.dailyCounts[row.date] || 0) + 1;

    let med = row.medication ? String(row.medication).trim().toUpperCase() : "";
    if (med !== "" && med !== "BLANK") metrics.medicationCounts[med] = (metrics.medicationCounts[med] || 0) + 1;

    let ins = row.insurance ? String(row.insurance).trim().toUpperCase() : "";
    if (ins !== "" && ins !== "BLANK") metrics.insuranceCounts[ins] = (metrics.insuranceCounts[ins] || 0) + 1;

    metrics.memberCounts[row.member] = (metrics.memberCounts[row.member] || 0) + 1;
  });

  const sortedDaily = {};
  Object.keys(metrics.dailyCounts).sort().forEach(k => { sortedDaily[k] = metrics.dailyCounts[k]; });
  metrics.dailyCounts = sortedDaily;

  const getTopN = (obj, n) => {
    return Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, n).reduce((acc, curr) => ({...acc, [curr[0]]: curr[1]}), {});
  };

  metrics.medicationCounts = getTopN(metrics.medicationCounts, 10);
  metrics.insuranceCounts = getTopN(metrics.insuranceCounts, 5);

  return metrics;
}