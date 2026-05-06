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
function getPAMetricsData(reportMode, teamName, employeeEmail, startDateStr, endDateStr) {
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
  else if (reportMode === 'employee' && employeeEmail) {
    let emailPrefix = employeeEmail.split('@')[0].toLowerCase();
    for (let i = 0; i < allPAMembers.length; i++) {
      let memberName = allPAMembers[i].toLowerCase();
      // Buscamos coincidencia (ej: "krizza" dentro de "krizzajohnson@...")
      if (emailPrefix.includes(memberName) || memberName.includes(emailPrefix)) {
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
function processPAData(selection, startDateStr, endDateStr, PA_SPREADSHEET_ID, EXCEPTIONS) {
  const startDate = new Date(startDateStr + "T00:00:00-05:00"); 
  const endDate = new Date(endDateStr + "T23:59:59-05:00");     
  
  let targetMembers = (selection === "All") ? [...PA_TEAM_MEMBERS, ...EXCEPTIONS.map(e => e.name)] : [selection];
  const mainSpreadsheet = SpreadsheetApp.openById(PA_SPREADSHEET_ID);
  let allRows = [];

  targetMembers.forEach(member => {
    let sheetToProcess = null;
    let isException = EXCEPTIONS.find(e => e.name === member);

    if (isException) {
      try {
        const extSs = SpreadsheetApp.openById(isException.spreadsheetId);
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
          
          const rowDate = new Date(row[dateIdx]);
          if (isNaN(rowDate.getTime())) continue;

          if (rowDate >= startDate && rowDate <= endDate) {
            allRows.push({
              member: member,
              status: statusIdx !== -1 ? row[statusIdx] : "Unknown",
              date: Utilities.formatDate(rowDate, "EST", "yyyy-MM-dd"),
              medication: medIdx !== -1 ? row[medIdx] : "Unknown",
              insurance: insIdx !== -1 ? row[insIdx] : "Unknown"
            });
          }
        }
      }
    }
  });

  let metrics = aggregatePAMetrics(allRows);
  metrics.selection = selection; // Guardamos el tipo de selección para el Frontend
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