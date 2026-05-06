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
  Procesa la lectura del Excel como texto puro sin zonas horarias
 **/
function processPAData(selection, startDateStr, endDateStr, PA_SPREADSHEET_ID, EXCEPTIONS) {
  let targetMembers = (selection === "All") ? [...PA_TEAM_MEMBERS, ...EXCEPTIONS.map(e => e.name)] : [selection];
  const mainSpreadsheet = SpreadsheetApp.openById(PA_SPREADSHEET_ID);
  let allRows = [];

  // Configuración de estados pendientes (en minúsculas para comparar fácil)
  const PENDING_STATUSES = ["in progress", "submitted", "appeal submitted", "appeal in progress", "other", "cart's open"];
  
  // Objeto para agrupar: { "Chart_Medicación": { statuses: [], latestEntry: {} } }
  let paGroups = {}; 
  let patientsThisPeriod = [];

  const today = new Date();
  const twoWeeksAgo = new Date();
  twoWeeksAgo.setDate(today.getDate() - 14);
  twoWeeksAgo.setHours(0,0,0,0);

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
        Logger.log("Error con excepción: " + member);
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
        const chartIdx = 3; // Columna D
        const nameIdx = 4;  // Columna E

        for (let r = headerRowIdx + 1; r < data.length; r++) {
          const row = data[r];
          if (dateIdx === -1 || !row[dateIdx]) continue;
          
          let rowDateObj = row[dateIdx] instanceof Date ? row[dateIdx] : new Date(row[dateIdx]);
          if (isNaN(rowDateObj.getTime())) continue;

          let rowDateStr = Utilities.formatDate(rowDateObj, sheetTz, "yyyy-MM-dd");
          let statusRaw = statusIdx !== -1 ? String(row[statusIdx]).trim() : "Unknown";
          let statusLower = statusRaw.toLowerCase();
          let chartNum = String(row[chartIdx] || "").trim();
          let patName = String(row[nameIdx] || "").trim();
          let medication = medIdx !== -1 ? String(row[medIdx]).trim().toLowerCase() : "unknown_med";

          // --- LÓGICA DE AGRUPAMIENTO PARA OVERDUE ---
          if (chartNum) {
            let groupKey = chartNum + "_" + medication;
            if (!paGroups[groupKey]) {
              paGroups[groupKey] = { statuses: [], entries: [] };
            }
            paGroups[groupKey].statuses.push(statusLower);
            paGroups[groupKey].entries.push({
              date: rowDateStr,
              dateObj: rowDateObj,
              status: statusRaw,
              patient: patName || "Unknown",
              chart: chartNum,
              medication: medication
            });
          }

          // Métricas normales (volumen diario, etc.)
          if (rowDateStr >= startDateStr && rowDateStr <= endDateStr) {
            if (patName || chartNum) patientsThisPeriod.push({ chart: chartNum, name: patName });

            allRows.push({
              member: member, status: statusRaw, date: rowDateStr,
              medication: medIdx !== -1 ? row[medIdx] : "Unknown",
              insurance: insIdx !== -1 ? row[insIdx] : "Unknown"
            });
          }
        }
      }
    }
  });

  // --- FILTRADO FINAL DE ATRASADOS (OVERDUE) ---
  let finalOverdue = [];
  for (let key in paGroups) {
    let group = paGroups[key];
    
    // Regla 1: ¿Todos los estados de este grupo (Paciente+Med) son pendientes?
    let allArePending = group.statuses.every(s => PENDING_STATUSES.includes(s));
    
    if (allArePending) {
      // Regla 2: Si todos son pendientes, buscamos la entrada más reciente
      group.entries.sort((a, b) => b.dateObj.getTime() - a.dateObj.getTime());
      let latest = group.entries[0];
      
      // Regla 3: ¿La fecha más reciente es de hace más de 2 semanas?
      if (latest.dateObj < twoWeeksAgo) {
        finalOverdue.push(latest);
      }
    }
    // Si hay aunque sea uno que NO es pendiente (ej. Approved), el hilo se ignora completo.
  }

  let metrics = aggregatePAMetrics(allRows);
  metrics.selection = selection; 
  metrics.patientsThisPeriod = patientsThisPeriod;
  
  finalOverdue.sort((a, b) => a.date > b.date ? 1 : -1);
  metrics.overduePAs = finalOverdue;
  
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