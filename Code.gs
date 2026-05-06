// ==========================================
// GLOBAL CONFIGURATION
// ==========================================
const scriptProps = PropertiesService.getScriptProperties();

const CALLS_FOLDER_ID = scriptProps.getProperty('CALLS_FOLDER_ID'); 
const SMS_FOLDER_ID = scriptProps.getProperty('SMS_FOLDER_ID'); 
const GEMINI_API_KEY = scriptProps.getProperty('GEMINI_API_KEY');
const SPREADSHEET_ID = scriptProps.getProperty('SPREADSHEET_ID');
const CALL_LOG_SHEET_ID = scriptProps.getProperty('CALL_LOG_SHEET_ID');
const REPORTS_PARENT_FOLDER_ID = scriptProps.getProperty('REPORTS_PARENT_FOLDER_ID');

const AUTHORIZED_MANAGERS = [
  "genesiscastillo@allhealthmedgroup.com", 
  "fernandopochintesta@allhealthmedgroup.com",
  "linatascon@allhealthmedgroup.com"
];

function doGet() {
  // Get the email of the person trying to open the web app
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  // Check if their email is on our VIP list
  if (AUTHORIZED_MANAGERS.indexOf(userEmail) === -1) {
    // If they are NOT on the list, show a friendly blocked screen
    var htmlBlocked = "<div style='font-family: sans-serif; text-align: center; margin-top: 100px; color: #333;'>" +
                      "<h1 style='color: #e74c3c;'>⛔ Restricted Access</h1>" +
                      "<p>We apologize, <b>" + (userEmail || "User") + "</b>.</p>" +
                      "<p>You do not have the required permissions to access the Performance Dashboard.</p>" +
                      "<p style='color: #7f8c8d; font-size: 12px; margin-top: 50px;'>If you believe this is an error, please contact the administrator.</p>" +
                      "</div>";
                      
    return HtmlService.createHtmlOutput(htmlBlocked)
        .setTitle('Access Denied')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // If they ARE on the list, load the Dashboard normally
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('All Health Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // <-- Very important for embedding inside Google Sites
}

function getTeamsFromSheet() {
  try {
    if (SPREADSHEET_ID === "PEGAR_AQUI_EL_ID_DE_TU_GOOGLE_SHEET") return { status: "error", message: "Please configure SPREADSHEET_ID in Code.gs" };
    var SHEET_NAME = 'Data';
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { status: "error", message: "Sheet is empty or missing data." };

    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var teamIdx = headers.indexOf("team");
    var idIdx = headers.indexOf("id");
    var nameIdx = headers.indexOf("name");

    if (teamIdx === -1 || idIdx === -1 || nameIdx === -1) return { status: "error", message: "Spreadsheet must have columns: Team, ID, Name" };

    var teamsMap = {};
    for (var i = 1; i < data.length; i++) {
      var teamName = data[i][teamIdx] ? data[i][teamIdx].toString().trim() : "";
      var empId = data[i][idIdx] ? data[i][idIdx].toString().trim() : "";
      var empName = data[i][nameIdx] ? data[i][nameIdx].toString().trim() : "";
      
      if (teamName && empId) {
        if (!teamsMap[teamName]) teamsMap[teamName] = [];
        teamsMap[teamName].push({ id: empId.toLowerCase(), name: empName || empId });
      }
    }
    return { status: "success", data: teamsMap };
  } catch (e) { return { status: "error", message: "Error reading Sheet: " + e.message }; }
}

function processUserFilters(form) {
  var reportMode = form.reportMode || 'employee'; 
  var employee = form.employeeEmail ? form.employeeEmail.trim() : "";
  var locationFilter = form.locationFilter || 'All'; 
  var teamName = form.teamFilter || '';
  
  var teamMembers = [];
  if (reportMode === 'teams' && teamName) {
    var teamsDb = getTeamsFromSheet();
    if (teamsDb.status === 'success' && teamsDb.data[teamName]) teamMembers = teamsDb.data[teamName];
  }
  
  var startParts = form.startDate.split("-");
  var endParts = form.endDate.split("-");
  var startDate = new Date(startParts[0], startParts[1] - 1, startParts[2], 0, 0, 0);
  var endDate = new Date(endParts[0], endParts[1] - 1, endParts[2], 23, 59, 59);
  var channel = form.channel;
  
  var diffTime = Math.abs(endDate - startDate);
  var diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
  if (diffDays > 31) return { status: "error", message: "Date range cannot exceed 31 days." };

  var metrics = {
    calls: { total: 0, inbound: 0, outbound: 0, duration: 0, daily: {}, dailyPatients: {} },
    sms: { total: 0, daily: {}, dailyPatients: {} },
    patients: {}, dailyPatients: {}, locations: {},
    hours: { morning: 0, afternoon: 0, evening: 0 },
    matchedEmails: {}, agentVolume: {}, teamStats: {},
    callResults: { answered: 0, missed: 0, abandoned: 0, other: 0 },
    busiestDay: { calls: [0,0,0,0,0,0,0], sms: [0,0,0,0,0,0,0], dayCounts: [0,0,0,0,0,0,0] },
    busiestHour: { calls: Array(24).fill(0), sms: Array(24).fill(0), dayCount: 0 },
    activeSpans: {}
  };

  // Pre-calcular cuántos lunes, martes, etc., hay en el rango para sacar promedios exactos
  var curr = new Date(startDate);
  var totalDays = 0;
  while (curr <= endDate) {
      metrics.busiestDay.dayCounts[curr.getDay()]++;
      totalDays++;
      curr.setDate(curr.getDate() + 1);
  }
  metrics.busiestHour.dayCount = totalDays;

  if (reportMode === 'teams') {
    teamMembers.forEach(function(m) {
      metrics.teamStats[m.id] = { 
        name: m.name, calls: 0, callsIn: 0, callsOut: 0, duration: 0, sms: 0, 
        dailyPatients: {}, activeSpans: {} 
      };
    });
  }
  
  var interactionsTimeline = []; 
  var months = getMonthsInRange(startDate, endDate);

  try {
    if (channel === 'calls' || channel === 'both') processFolder(CALLS_FOLDER_ID, "Weave_Phone_Report", months, startDate, endDate, reportMode, employee, locationFilter, teamMembers, "calls", metrics, interactionsTimeline);
    if (channel === 'sms' || channel === 'both') processFolder(SMS_FOLDER_ID, "Weave_SMS_Report", months, startDate, endDate, reportMode, employee, locationFilter, teamMembers, "sms", metrics, interactionsTimeline);
  } catch (e) { return { status: "error", message: e.message }; }

  var insights = compileInsights(metrics);
  var officialEmployeeEmail = (reportMode === 'employee') ? "" : (reportMode === 'teams' ? teamName : "Global");
  var gaps = []; var attendance = null;

  if (reportMode === 'employee') {
      var matchedArr = Object.keys(metrics.matchedEmails);
      var officialEmployeeEmail = "";
      if (matchedArr.length === 0) {
        // Si no hay llamadas, no matamos el script. Usamos lo que escribió el usuario para buscar en PA.
        officialEmployeeEmail = employee; 
      } else if (matchedArr.length > 1) {
        var errorHtml = "Multiple employees found. Please use one of these exact emails:<br><br><ul style='text-align:left; margin-top:5px;'>";
        matchedArr.forEach(function(e) { errorHtml += "<li>" + e + "</li>"; });
        return { status: "error", message: errorHtml + "</ul>" };
      } else {
        officialEmployeeEmail = matchedArr[0];
      }
      
      // 1. Buscamos el perfil en el Spreadsheet usando el correo oficial
      var employeeProfile = getEmployeeProfile(officialEmployeeEmail);
      // 2. Definimos variables seguras (si no está en el Excel, usa el correo/Default)
      var finalName = employeeProfile ? employeeProfile.name : officialEmployeeEmail;
      var finalTeam = employeeProfile ? employeeProfile.team : "General";
      var finalSchedule = employeeProfile ? employeeProfile.schedule : null;
      var callLogData = getCallLogMetrics(finalName);

      // Le pasamos finalSchedule (Excel) y form (Web) a ambas funciones
      if (channel === 'both') gaps = calculateGaps(interactionsTimeline, finalSchedule, form);
      if (channel === 'both' && (finalSchedule || form.useSchedule === 'true')) attendance = evaluateAttendance(startDate, endDate, finalSchedule, form, metrics.calls.daily, metrics.sms.daily, channel);
      
      // Calcular "Pace" (Ritmo de trabajo) para Empleado
      var totalHours = 0;
      for (var d in metrics.activeSpans) {
          var span = (metrics.activeSpans[d].max - metrics.activeSpans[d].min) / 3600000;
          totalHours += Math.max(1, span); // Mínimo 1 hora para no inflar métricas
      }
      insights.pace = totalHours > 0 ? Math.round((metrics.calls.total + metrics.sms.total) / totalHours) : 0;
      
  } else if (reportMode === 'global') {
      var topAgents = [];
      for (var agent in metrics.agentVolume) topAgents.push({name: agent, count: metrics.agentVolume[agent]});
      topAgents.sort(function(a,b) { return b.count - a.count; });
      insights.topAgents = topAgents;
  } else if (reportMode === 'teams') {
      for (var k in metrics.teamStats) {
        var t = metrics.teamStats[k];
        t.aht = t.calls > 0 ? Math.round(t.duration / t.calls) : 0;
        t.avgDailyPatients = calculateAvgDailyPatients(t.dailyPatients); 
        
        // Calcular "Pace" individual
        var tHrs = 0;
        for(var d in t.activeSpans) {
            var sp = (t.activeSpans[d].max - t.activeSpans[d].min) / 3600000;
            tHrs += Math.max(1, sp);
        }
        t.pace = tHrs > 0 ? Math.round((t.calls + t.sms) / tHrs) : 0;
      }
  }

  var callsData = metrics.calls;
  var aht = callsData.total > 0 ? Math.round(callsData.duration / callsData.total) : 0;
  callsData.ahtMin = Math.floor(aht / 60); callsData.ahtSec = aht % 60;
  callsData.talkHours = (callsData.duration / 3600).toFixed(2);
  callsData.stats = calculateDailyStats(callsData.daily, callsData.total);
  callsData.avgDailyPatients = calculateAvgDailyPatients(metrics.calls.dailyPatients);

  var smsData = metrics.sms;
  smsData.stats = calculateDailyStats(smsData.daily, smsData.total);
  smsData.avgDailyPatients = calculateAvgDailyPatients(metrics.sms.dailyPatients);
  insights.avgDailyPatients = calculateAvgDailyPatients(metrics.dailyPatients);

  var paDashboardData = getPAMetricsData(reportMode, teamName, officialEmployeeEmail, form.startDate, form.endDate);

  return {
    status: "success",
    data: {
      reportMode: reportMode, 
      locationName: locationFilter, 
      employee: officialEmployeeEmail,
      employeeDisplayName: finalName,
      employeeTeam: finalTeam,
      employeeSchedule: finalSchedule,
      callLogs: callLogData, 
      teamMembers: teamMembers,
      teamStats: metrics.teamStats, startDate: form.startDate, endDate: form.endDate, channel: channel,
      calls: callsData, sms: smsData, callResults: metrics.callResults,
      busiestDay: metrics.busiestDay, busiestHour: metrics.busiestHour,
      insights: insights, topAgents: insights.topAgents || [], hoursData: metrics.hours, gaps: gaps, attendance: attendance,
      paDashboardData: paDashboardData
    }
  };
}

function processFolder(folderId, prefix, months, startDt, endDt, reportMode, email, locationFilter, teamMembers, type, metrics, timeline) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    
    if (fileName.indexOf(prefix) !== -1 && fileName.toLowerCase().indexOf(".csv") !== -1) {
      if (months.some(function(m) { return fileName.toLowerCase().indexOf(m.toLowerCase()) !== -1; })) {
        var textData = file.getBlob().getDataAsString("UTF-8");
        if (!textData || textData.trim() === "") continue;

        var csvData;
        try { csvData = Utilities.parseCsv(textData); } catch (e) { try { csvData = Utilities.parseCsv(textData.replace(/"/g, "'")); } catch (e2) { continue; } }
        if (!csvData || csvData.length <= 1) continue;

        var headers = csvData[0].map(function(h) { return h.trim(); });
        var dateIdx = headers.indexOf(type === 'calls' ? 'Date/Time' : 'Date/ Time');
        var userIdx = headers.indexOf(type === 'calls' ? 'Office User' : 'Sender Name');
        var locIdx = headers.indexOf('Location');
        var patIdx = headers.indexOf('Patient Name');
        var durIdx = headers.indexOf('Duration');
        var cTypeIdx = headers.indexOf('Call Type');
        var resIdx = headers.indexOf('Result'); // NUEVO
        var smsTypeIdx = headers.indexOf('SMS Type'); // NUEVO

        if (dateIdx === -1) continue;

        for (var i = 1; i < csvData.length; i++) {
          var row = csvData[i];
          if (!row[dateIdx]) continue;

          var rawDateStr = row[dateIdx].trim();
          var cleanDateStr = rawDateStr.replace(/\+0000/g, "").replace(/UTC/g, "").trim().replace(" ", "T"); 
          var rowDate = new Date(cleanDateStr + "Z");
          
          if (rowDate >= startDt && rowDate <= endDt) {
            var rowUser = (userIdx > -1 && row[userIdx]) ? row[userIdx].trim() : "Unknown Agent";
            var rowLoc = (locIdx > -1 && row[locIdx]) ? row[locIdx].trim() : "Unknown";
            var isMatch = false; var matchedTeamId = null;

            if (reportMode === 'employee') {
                if (!rowUser) continue;
                var cleanUser = rowUser.toLowerCase(); var searchInput = email.toLowerCase(); 
                if (searchInput.indexOf("@") !== -1) { isMatch = (cleanUser === searchInput); } 
                else {
                  var searchStr = searchInput.replace(/\s+/g, ""); var userPrefix = cleanUser.split("@")[0].replace(/\./g, ""); 
                  if (userPrefix.indexOf(searchStr) !== -1 || searchStr.indexOf(userPrefix) !== -1) isMatch = true;
                }
            } else if (reportMode === 'global') {
                if (locationFilter === 'All' || rowLoc.toLowerCase().indexOf(locationFilter.toLowerCase()) !== -1) isMatch = true;
            } else if (reportMode === 'teams') {
                if (!rowUser) continue;
                var cu = rowUser.toLowerCase();
                for (var j = 0; j < teamMembers.length; j++) {
                   if (cu === teamMembers[j].id || cu.split("@")[0].replace(/\./g, "") === teamMembers[j].id.replace(/\s+/g, "")) {
                      isMatch = true; matchedTeamId = teamMembers[j].id; break;
                   }
                }
            }

            if (isMatch) {
              var startMs = rowDate.getTime();
              var endMs = startMs; 
              var dayKey = (rowDate.getMonth() + 1) + "/" + rowDate.getDate() + "/" + rowDate.getFullYear();
              var dayOfWeek = rowDate.getDay();
              var hour = rowDate.getHours();

              // Track Active Span (Horas Activas Reales)
              if (reportMode === 'employee') {
                  metrics.matchedEmails[rowUser] = true; 
                  metrics.activeSpans[dayKey] = metrics.activeSpans[dayKey] || { min: startMs, max: startMs };
                  metrics.activeSpans[dayKey].min = Math.min(metrics.activeSpans[dayKey].min, startMs);
                  metrics.activeSpans[dayKey].max = Math.max(metrics.activeSpans[dayKey].max, startMs);
              } else if (reportMode === 'teams' && matchedTeamId) {
                  metrics.teamStats[matchedTeamId].activeSpans[dayKey] = metrics.teamStats[matchedTeamId].activeSpans[dayKey] || { min: startMs, max: startMs };
                  metrics.teamStats[matchedTeamId].activeSpans[dayKey].min = Math.min(metrics.teamStats[matchedTeamId].activeSpans[dayKey].min, startMs);
                  metrics.teamStats[matchedTeamId].activeSpans[dayKey].max = Math.max(metrics.teamStats[matchedTeamId].activeSpans[dayKey].max, startMs);
              }

              metrics.agentVolume[rowUser] = (metrics.agentVolume[rowUser] || 0) + 1;

              if (type === 'calls') {
                metrics.calls.total++;
                metrics.calls.daily[dayKey] = (metrics.calls.daily[dayKey] || 0) + 1;
                var duration = durIdx > -1 && row[durIdx] ? parseInt(row[durIdx].trim()) : 0;
                if (!isNaN(duration)) { metrics.calls.duration += duration; endMs += (duration * 1000); }
                
                var cType = cTypeIdx > -1 && row[cTypeIdx] ? row[cTypeIdx].trim().toLowerCase() : "";
                if (cType === 'inbound') {
                    metrics.calls.inbound++;
                    metrics.busiestDay.calls[dayOfWeek]++;
                    metrics.busiestHour.calls[hour]++;
                    
                    var resultText = (resIdx > -1 && row[resIdx]) ? row[resIdx].trim().toLowerCase() : "";
                    if (resultText.indexOf('answered') > -1) metrics.callResults.answered++;
                    else if (resultText.indexOf('missed') > -1) metrics.callResults.missed++;
                    else if (resultText.indexOf('abandoned') > -1) metrics.callResults.abandoned++;
                    else metrics.callResults.other++;
                }
                if (cType === 'outbound') metrics.calls.outbound++;
                
                if (reportMode === 'teams' && matchedTeamId) {
                  metrics.teamStats[matchedTeamId].calls++;
                  metrics.teamStats[matchedTeamId].duration += (isNaN(duration) ? 0 : duration);
                  if (cType === 'inbound') metrics.teamStats[matchedTeamId].callsIn++;
                  if (cType === 'outbound') metrics.teamStats[matchedTeamId].callsOut++;
                }
              } else {
                metrics.sms.total++;
                metrics.sms.daily[dayKey] = (metrics.sms.daily[dayKey] || 0) + 1;
                var sType = (smsTypeIdx > -1 && row[smsTypeIdx]) ? row[smsTypeIdx].trim().toLowerCase() : "";
                if (sType === 'inbound') {
                    metrics.busiestDay.sms[dayOfWeek]++;
                    metrics.busiestHour.sms[hour]++;
                }
                if (reportMode === 'teams' && matchedTeamId) metrics.teamStats[matchedTeamId].sms++;
              }

              if (reportMode === 'employee') timeline.push({ start: startMs, end: endMs });

              var patient = patIdx > -1 && row[patIdx] ? row[patIdx].trim() : "";
              if (patient) {
                metrics.patients[patient] = true;
                metrics.dailyPatients[dayKey] = metrics.dailyPatients[dayKey] || {};
                metrics.dailyPatients[dayKey][patient] = true;

                if (type === 'calls') {
                  metrics.calls.dailyPatients[dayKey] = metrics.calls.dailyPatients[dayKey] || {};
                  metrics.calls.dailyPatients[dayKey][patient] = true;
                } else {
                  metrics.sms.dailyPatients[dayKey] = metrics.sms.dailyPatients[dayKey] || {};
                  metrics.sms.dailyPatients[dayKey][patient] = true;
                }
                
                if (reportMode === 'teams' && matchedTeamId) {
                  metrics.teamStats[matchedTeamId].dailyPatients[dayKey] = metrics.teamStats[matchedTeamId].dailyPatients[dayKey] || {};
                  metrics.teamStats[matchedTeamId].dailyPatients[dayKey][patient] = true;
                }
              }

              var locName = rowLoc.replace("All Health Medical Group - ", "");
              metrics.locations[locName] = (metrics.locations[locName] || 0) + 1;

              var hr = rowDate.getHours();
              if (hr >= 6 && hr < 12) metrics.hours.morning++;
              else if (hr >= 12 && hr < 17) metrics.hours.afternoon++;
              else metrics.hours.evening++;
            }
          }
        }
      }
    }
  }
}

// Auxiliares (sin cambios)
function getMonthsInRange(s, e) { 
  var m=["January","February","March","April","May","June","July","August","September","October","November","December"]; 
  var r=[m[s.getMonth()]]; 
  if(s.getMonth()!==e.getMonth()) r.push(m[e.getMonth()]); 
  return r; 
}

function compileInsights(m) { var p="Morning (6 AM - 12 PM)", mh=m.hours.morning; if(m.hours.afternoon>mh){p="Afternoon (12 PM - 5 PM)";mh=m.hours.afternoon;} if(m.hours.evening>mh)p="Evening (5 PM onwards)"; return {uniquePatients:Object.keys(m.patients).length, peakTime:p}; }

function calculateGaps(tl, schedule, form) {
  if (tl.length === 0) return [];
  var d = {};
  tl.forEach(function(ix) {
    var dt = new Date(ix.start);
    var ds = (dt.getMonth() + 1) + "/" + dt.getDate() + "/" + dt.getFullYear();
    if (!d[ds]) d[ds] = { rawDate: dt, intervals: [] };
    d[ds].intervals.push(ix);
  });
  var dg = [];

  for (var ds in d) {
    var ixs = d[ds].intervals;
    ixs.sort(function(a, b) { return a.start - b.start; });
    var m = [ixs[0]];
    for (var i = 1; i < ixs.length; i++) {
      var l = m[m.length - 1], c = ixs[i];
      if (c.start <= l.end) l.end = Math.max(l.end, c.end);
      else m.push(c);
    }
    var sd = d[ds].rawDate;
    var dayOfWeek = sd.getDay();
    var isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
    var shiftS, shiftE, lh;

    if (schedule && schedule[dayOfWeek]) {
      shiftS = schedule[dayOfWeek].start;
      shiftE = schedule[dayOfWeek].end;
      lh = schedule[dayOfWeek].lunchStart;
    } else {
      shiftS = isWeekend ? form.weStart : form.wdStart;
      shiftE = isWeekend ? form.weEnd : form.wdEnd;
      lh = isWeekend ? form.weLunch : form.wdLunch;
    }

    if (shiftS === undefined || shiftS === null || shiftS === "") continue; 

    var b = parseScheduleBounds(shiftS, shiftE, sd, m[0].start, m[m.length - 1].end);
    var dgf = [];
    var minGapThreshold = isWeekend ? 120 : 30; // 2 horas finde, 30 mins semana

    // --- NUEVO MOTOR: BLOQUES ACTIVOS ESTRICTOS ---
    var activeBlocks = [];
    
    // Si hay almuerzo, partimos el turno en dos bloques
    if (lh !== undefined && lh !== null && lh !== "") {
      var lhf = parseFloat(lh); // Acepta decimales (ej. 18.66 para 6:40 PM)
      var lunchStartDt = new Date(sd);
      // Calculamos la hora exacta y el minuto exacto
      lunchStartDt.setHours(Math.floor(lhf), Math.round((lhf - Math.floor(lhf)) * 60), 0, 0);
      var lunchStartMs = lunchStartDt.getTime();
      var lunchEndMs = lunchStartMs + 3600000; // El almuerzo termina EXACTAMENTE 1 hora después (60 * 60 * 1000)

      activeBlocks.push({ start: b.start, end: lunchStartMs }); // Bloque 1: Ingreso a Inicio de Almuerzo
      activeBlocks.push({ start: lunchEndMs, end: b.end });     // Bloque 2: Fin de Almuerzo a Salida
    } else {
      // Si no hay almuerzo, es un solo bloque de corrido
      activeBlocks.push({ start: b.start, end: b.end });
    }

    // Evaluamos los gaps bloque por bloque
    for (var bIdx = 0; bIdx < activeBlocks.length; bIdx++) {
      var block = activeBlocks[bIdx];
      var cl = block.start; // Reloj actual

      for (var j = 0; j < m.length; j++) {
        var mg = m[j]; // Intervalo de actividad
        
        if (mg.end <= cl) continue; // Actividad en el pasado, ignorar
        if (mg.start >= block.end) break; // Actividad fuera de este bloque, pasar al siguiente

        // Si hay espacio entre nuestro reloj actual y la próxima actividad
        if (mg.start > cl) {
          var gapStart = cl;
          var gapEnd = Math.min(mg.start, block.end); // El gap topa contra el límite del bloque si es necesario
          var dm = Math.floor((gapEnd - gapStart) / 60000);
          
          if (dm >= minGapThreshold) {
            dgf.push(formatGapObj(gapStart, gapEnd, dm, ""));
          }
        }
        // Avanzar el reloj
        cl = Math.max(cl, Math.min(mg.end, block.end));
      }

      // Verificamos si quedó un gap al final del bloque (Ej: si se fue temprano antes del almuerzo o antes de salir)
      if (block.end > cl) {
        var dm = Math.floor((block.end - cl) / 60000);
        if (dm >= minGapThreshold) {
          dgf.push(formatGapObj(cl, block.end, dm, ""));
        }
      }
    }

    if (dgf.length > 0) dg.push({ rawDateMs: sd.getTime(), date: ds, items: dgf });
  }
  dg.sort(function(a, b) { return a.rawDateMs - b.rawDateMs; });
  return dg;
}

function parseScheduleBounds(s, e, doj, fi, li) {
  if (s === undefined || s === null || s === "" || e === undefined || e === null || e === "") return {start: fi, end: li};
  
  var sf = parseFloat(s);
  var ef = parseFloat(e);
  
  var sd = new Date(doj);
  sd.setHours(Math.floor(sf), Math.round((sf - Math.floor(sf)) * 60), 0, 0);
  
  var ed = new Date(doj);
  ed.setHours(Math.floor(ef), Math.round((ef - Math.floor(ef)) * 60), 0, 0);
  
  return {start: sd.getTime(), end: ed.getTime()};
}
function formatGapObj(s,e,m,t){return{startStr:formatTimeAMPM(new Date(s)),endStr:formatTimeAMPM(new Date(e)),durationStr:formatDurationStr(m)+(t||"")};}
function formatDurationStr(m){if(m<60)return m+" mins";var h=Math.floor(m/60),mi=m%60;return h+"h "+(mi>0?mi+"m":"");}
function calculateDailyStats(doj,t){var d=Object.keys(doj);if(d.length===0)return{avg:0,maxDay:"N/A",maxVal:0,minDay:"N/A",minVal:0};var a=[];for(var k in doj)a.push({date:k,count:doj[k]});a.sort(function(x,y){return y.count-x.count;});return{avg:Math.round(t/d.length),maxDay:a[0].date,maxVal:a[0].count,minDay:a[a.length-1].date,minVal:a[a.length-1].count};}
function formatTimeAMPM(d){var h=d.getHours(),m=d.getMinutes(),a=h>=12?'PM':'AM';h=h%12||12;m=m<10?'0'+m:m;return h+':'+m+' '+a;}

function evaluateAttendance(s,e,schedule,f,cd,sd,c){
  var ex=[f.wd_0==='true',f.wd_1==='true',f.wd_2==='true',f.wd_3==='true',f.wd_4==='true',f.wd_5==='true',f.wd_6==='true'],ad={};
  if(c==='calls'||c==='both')for(var d in cd)ad[d]=true;
  if(c==='sms'||c==='both')for(var d in sd)ad[d]=true;
  var ab=[],ow=[],cu=new Date(s),dn=["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],mn=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  while(cu<=e){
    var dw=cu.getDay();
    var ie = (schedule && schedule[dw]) ? true : ex[dw];
    var ds=(cu.getMonth()+1)+"/"+cu.getDate()+"/"+cu.getFullYear(); 
    var hw=ad[ds]===true, dd=dn[dw]+", "+mn[cu.getMonth()]+" "+cu.getDate();
    if(ie&&!hw)ab.push(dd);
    else if(!ie&&hw)ow.push(dd);
    cu.setDate(cu.getDate()+1);
  }return{
    absences:ab,offDayWork:ow
  };
}

function calculateAvgDailyPatients(dp){var d=Object.keys(dp);if(d.length===0)return 0;var tp=0;for(var k in dp)tp+=Object.keys(dp[k]).length;return Math.round(tp/d.length);}

function runGeminiAnalysis(payload) {
  var statsData = payload.stats; var startDateStr = payload.startDate; var endDateStr = payload.endDate;
  var reportMode = payload.reportMode || 'employee'; var locationName = payload.locationName || 'All';
  var teamMembers = payload.teamMembers || []; var email = statsData.employee; 
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "TU_CLAVE_DE_API_GEMINI_AQUI") return "<b>⚠️ API Key Missing</b>";

  try {
    var startParts = startDateStr.split("-"); var endParts = endDateStr.split("-");
    var startDt = new Date(startParts[0], startParts[1] - 1, startParts[2], 0, 0, 0); var endDt = new Date(endParts[0], endParts[1] - 1, endParts[2], 23, 59, 59);
    var months = getMonthsInRange(startDt, endDt); var conversations = {}; var smsCount = 0;
    var folder = DriveApp.getFolderById(SMS_FOLDER_ID); var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next(); var fileName = file.getName();
      if (fileName.indexOf("Weave_SMS_Report") !== -1 && fileName.toLowerCase().indexOf(".csv") !== -1) {
        if (months.some(function(m) { return fileName.toLowerCase().indexOf(m.toLowerCase()) !== -1; })) {
          var textData = file.getBlob().getDataAsString("UTF-8"); var csvData;
          try { csvData = Utilities.parseCsv(textData); } catch (e) { try { csvData = Utilities.parseCsv(textData.replace(/"/g, "'")); } catch (e2) { continue; } }
          if (!csvData || csvData.length <= 1) continue;
          var headers = csvData[0].map(function(h) { return h.trim(); });
          var dateIdx = headers.indexOf('Date/ Time'); var userIdx = headers.indexOf('Sender Name'); var patIdx = headers.indexOf('Patient Name'); var typeIdx = headers.indexOf('SMS Type'); var textIdx = headers.indexOf('Sms Text'); var locIdx = headers.indexOf('Location');

          for (var i = 1; i < csvData.length; i++) {
            var row = csvData[i]; if (!row[dateIdx] || !row[patIdx] || !row[textIdx]) continue;
            var rawDateStr = row[dateIdx].trim(); var cleanDateStr = rawDateStr.replace(/\+0000/g, "").replace(/UTC/g, "").trim().replace(" ", "T"); var rowDate = new Date(cleanDateStr + "Z");
            if (rowDate >= startDt && rowDate <= endDt) {
                var rowUser = row[userIdx] ? row[userIdx].trim() : ""; var rowLoc = (locIdx > -1 && row[locIdx]) ? row[locIdx].trim() : ""; var isFromPatient = (row[typeIdx] && row[typeIdx].trim().toLowerCase() === 'inbound'); var isValidForAI = false;
                if (reportMode === 'employee') { if (rowUser.toLowerCase() === email.toLowerCase() || isFromPatient) isValidForAI = true; } 
                else if (reportMode === 'global') { if (locationName === 'All' || rowLoc.toLowerCase().indexOf(locationName.toLowerCase()) !== -1) isValidForAI = true; } 
                else if (reportMode === 'teams') {
                    var cu = rowUser.toLowerCase(); for (var j = 0; j < teamMembers.length; j++) { if (cu === teamMembers[j].id || cu.split("@")[0].replace(/\./g, "") === teamMembers[j].id.replace(/\s+/g, "")) { isValidForAI = true; break; } }
                    if (isFromPatient) isValidForAI = true;
                }
                if (isValidForAI) {
                  var patName = row[patIdx].trim(); if (!conversations[patName]) conversations[patName] = [];
                  var timeStr = rowDate.getHours() + ":" + (rowDate.getMinutes() < 10 ? '0' : '') + rowDate.getMinutes();
                  var speaker = isFromPatient ? "Patient" : (rowUser ? "Agent (" + rowUser.split("@")[0] + ")" : "Agent");
                  conversations[patName].push("[" + timeStr + "] " + speaker + ": " + row[textIdx].trim()); smsCount++;
                }
            }
          }
        }
      }
    }

    if (smsCount === 0) return "<b>Notice:</b> No SMS conversations found.";
    var allPatients = Object.keys(conversations); 
    var shuffledPatients = shuffleArray(allPatients); 
    var selectedPatients = shuffledPatients.slice(0, 40); 
    var conversationLog = ""; selectedPatients.forEach(function(p) { conversationLog += "--- Patient: " + p + " ---\n" + conversations[p].join("\n") + "\n\n"; });
    var prompt = "";

    if (reportMode === 'employee') {
        prompt = "Act as a QA Manager for All Health Medical Group. Audit this specific employee.\n\nTRANSCRIPTS:\n" + conversationLog + "\nProvide EXACTLY this HTML structure. Do not use Markdown:\n1. <b>Overall Quality:</b> [Score]/10 - [Phrase]<br><br>\n2. <b>Professional Tone:</b> [Score]/10<br><i>Feedback:</i> [Phrase]<br><i>Example:</i> '[Quote]'<br><br>\n3. <b>Clarity & Grammar:</b> [Score]/10<br><i>Feedback:</i> [Phrase]<br><i>Example:</i> '[Quote]'<br><br>\n4. <b>First Contact Resolution:</b> [Score]/10<br><br>\n5. <b>Manager Feedback:</b><br><ul><li>[Point 1]</li></ul>\n6. <b>Recommendations:</b><br><ul><li>[Rec 1]</li></ul>\n";
    } else if (reportMode === 'global') {
        prompt = "Act as an Operations Director. Analyze this random sample of recent SMS traffic for: " + locationName + "\n\nTRANSCRIPTS:\n" + conversationLog + "\nProvide EXACTLY this HTML structure. Do not use Markdown:\n1. <b>Patient Sentiment & Tone:</b> [Positive/Neutral/Negative] - [Brief explanation].<br><br>\n2. <b>Top Contact Drivers:</b><br><ul><li>[Reason 1]</li></ul><br>\n3. <b>Operational Bottlenecks Detected:</b><br><ul><li>[Issue 1]</li></ul><br>\n4. <b>Front Desk Performance (General):</b> [Brief feedback].<br><br>\n5. <b>Actionable Recommendations:</b><br><ul><li>[Rec 1]</li></ul>\n";
    } else if (reportMode === 'teams') {
        prompt = "Act as a Team Performance Supervisor. Evaluate the communication style and effectiveness of this specific team.\n\nTRANSCRIPTS:\n" + conversationLog + "\nProvide EXACTLY this HTML structure. Do not use Markdown:\n1. <b>Team Synergy & Standardization:</b> [Analysis of whether all members follow similar protocols].<br><br>\n2. <b>Member Spotlights:</b><br><ul><li><b>Top Communicator:</b> [Name] - [Why]</li><li><b>Needs Coaching:</b> [Name] - [Why]</li></ul><br>\n3. <b>Common Mistakes / Areas to Improve:</b><br><ul><li>[Mistake 1]</li></ul><br>\n4. <b>Action Plan for Next Team Meeting:</b><br><ul><li>[Topic 1 to discuss]</li></ul>\n";
    }

    const apiKey = GEMINI_API_KEY.trim();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    var apiPayload = { "contents": [{ "parts": [{"text": prompt}] }], "generationConfig": { "temperature": 0.3 } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(apiPayload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    if (json.error) return "<b>API Error:</b> " + json.error.message;
    return json.candidates[0].content.parts[0].text;
  } catch (e) { return "<b>Critical AI Error:</b> " + e.message; }
}

function shuffleArray(array) { 
  for (var i = array.length - 1; i > 0; i--) { 
    var j = Math.floor(Math.random() * (i + 1)); 
    var temp = array[i]; array[i] = array[j]; array[j] = temp; 
  } 
  return array; 
}

// --- NUEVO MÓDULO: INTEGRACIÓN CON HR SPREADSHEET ---
function getEmployeeProfile(email) {
  if (!email) return null;
  
  try {
    var SHEET_NAME = 'Data';
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    
    // Empezamos en 1 asumiendo que la fila 0 son los títulos (Team, Email, Name, Mon, Tue...)
    for (var i = 1; i < data.length; i++) {
      var rowEmail = String(data[i][1]).trim().toLowerCase(); // Columna B (1)
      
      if (rowEmail === email.toLowerCase()) {
        return {
          team: data[i][0] ? String(data[i][0]).trim() : "Unassigned", // Columna A
          name: data[i][2] ? String(data[i][2]).trim() : email,        // Columna C
          schedule: {
            1: parseShiftString(data[i][3]), // D: Lunes
            2: parseShiftString(data[i][4]), // E: Martes
            3: parseShiftString(data[i][5]), // F: Miércoles
            4: parseShiftString(data[i][6]), // G: Jueves
            5: parseShiftString(data[i][7]), // H: Viernes
            6: parseShiftString(data[i][8]), // I: Sábado
            0: parseShiftString(data[i][9])  // J: Domingo (0 en JavaScript)
          }
        };
      }
    }
  } catch (e) {
    Logger.log("Error leyendo Spreadsheet: " + e.message);
  }
  
  return null; // Retorna null si no lo encuentra (falla elegantemente)
}

// Convierte "8 AM - 8 PM (1 PM)" a {start: 8, end: 20, lunchStart: 13, lunchEnd: 14}
function parseShiftString(shiftStr) {
  if (!shiftStr) return null;
  var str = String(shiftStr).trim().toUpperCase();
  if (str === "" || str === "OFF") return null;

  try {
    // Separa el shift principal del lunch que está entre paréntesis
    var parts = str.split("(");
    var shiftParts = parts[0].split("-");
    
    var startHour = convertTo24H(shiftParts[0]);
    var endHour = convertTo24H(shiftParts[1]);
    
    var lunchStart = null;
    var lunchEnd = null;
    
    if (parts.length > 1) {
      var lunchStr = parts[1].replace(")", "");
      lunchStart = convertTo24H(lunchStr);
      lunchEnd = lunchStart + 1; // Asumimos 1 hora de lunch
    }

    return {
      start: startHour,
      end: endHour,
      lunchStart: lunchStart,
      lunchEnd: lunchEnd
    };
  } catch(e) {
    return null; // Si el formato está mal escrito, lo ignora
  }
}

// Función auxiliar para pasar de "8 AM" a 8, o "1:30 PM" a 13.5
function convertTo24H(timeStr) {
  if (!timeStr) return 0;
  // Este Regex nuevo captura opcionalmente los minutos después de los dos puntos
  var match = timeStr.match(/(\d+)(?::(\d+))?\s*(AM|PM)/i);
  if (!match) return 0;
  
  var hour = parseInt(match[1], 10);
  var minutes = match[2] ? parseInt(match[2], 10) : 0;
  var ampm = match[3].toUpperCase();
  
  if (ampm === "PM" && hour < 12) hour += 12;
  if (ampm === "AM" && hour === 12) hour = 0; // Midnight
  
  // Retornamos el número en formato decimal (Ej: 1:30 PM = 13.5)
  return hour + (minutes / 60); 
}

// --- NUEVO MÓDULO: CALL LOG METRICS ---
function getCallLogMetrics(employeeName) {
  if (!employeeName) return null;
  
  try {
    // CAMBIA "Sheet1" por el nombre real de la pestaña en tu Excel de Call Logs
    var sheet = SpreadsheetApp.openById(CALL_LOG_SHEET_ID).getSheetByName("Phone/Weave Ev"); 
    var data = sheet.getDataRange().getValues();
    var searchName = String(employeeName).trim().toLowerCase();

    // Empezamos en 1 asumiendo que la fila 0 son los títulos
    for (var i = 1; i < data.length; i++) {
      var rowName = String(data[i][0]).trim().toLowerCase(); // Columna A
      
      // Búsqueda flexible (por si en un lado dice "Chanielle Walters" y en el otro solo "Chanielle")
      if (rowName === searchName || searchName.indexOf(rowName) !== -1 || rowName.indexOf(searchName) !== -1) {
        // Procesamos y redondeamos los datos antes de enviarlos
        var fVal = data[i][5]; // Rolling Avg
        var gVal = data[i][6]; // Last Week
        var hVal = data[i][7]; // Perf Change
        var lVal = data[i][11]; // Efficiency

        // Redondeo sin decimales (Average y Last Week)
        var rollingAvg = (fVal !== "" && !isNaN(fVal)) ? Math.round(Number(fVal)) : (fVal || "N/A");
        var lastWeek = (gVal !== "" && !isNaN(gVal)) ? Math.round(Number(gVal)) : (gVal || "N/A");
        
        // Efficiency: Sin decimales
        var efficiency = (lVal !== "" && !isNaN(lVal)) ? Math.round(Number(lVal)) : (lVal || "N/A");

        // Performance Change: 2 decimales y símbolo %
        var perfChange = "N/A";
        if (hVal !== "") {
          if (!isNaN(hVal)) {
            var numH = Number(hVal);
            // Si en Google Sheets está en formato %, llega como decimal (ej. 0.15). Lo multiplicamos.
            // (Si en tu Excel es un número puro como "15", simplemente bórra el "* 100")
            if (Math.abs(numH) < 2) numH = numH * 100; 
            perfChange = numH.toFixed(2) + "%";
          } else {
            perfChange = hVal; // Por si la celda ya es un texto como "15%"
          }
        }

        return {
          rollingAvg: rollingAvg,
          lastWeek: lastWeek,
          perfChange: perfChange,
          efficiency: efficiency
        };
      }
    }
  } catch (e) {
    Logger.log("Error leyendo Call Logs: " + e.message);
  }
  
  return null;
}

// --- NUEVO MÓDULO: GUARDADO DE PDF EN DRIVE (POR MES/AÑO) ---
// --- NUEVO MÓDULO: GUARDADO DE PDF EN DRIVE (POR MES/AÑO) ---
function createPdfInDrive(htmlContent, fileName, targetDateStr) {
  try {
    // 1. Convertir el HTML a un documento PDF
    var blob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
    blob.setName(fileName);
    
    // 2. Extraer Mes y Año de la fecha final (Ej: "2026-04-10" -> "April 2026")
    var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var dateParts = targetDateStr.split("-"); // [YYYY, MM, DD]
    var year = dateParts[0];
    var monthIndex = parseInt(dateParts[1], 10) - 1; // Convertir "04" a índice 3 (April)
    var folderMonthYear = monthNames[monthIndex] + " " + year; 
    
    // 3. Conectar a la carpeta padre
    var parentFolder = DriveApp.getFolderById(REPORTS_PARENT_FOLDER_ID);
    
    // 4. Buscar si la subcarpeta "April 2026" ya existe, si no, crearla
    var subFolders = parentFolder.getFoldersByName(folderMonthYear);
    var targetFolder;
    
    if (subFolders.hasNext()) {
      targetFolder = subFolders.next();
    } else {
      targetFolder = parentFolder.createFolder(folderMonthYear);
    }
    
    // 5. Guardar el archivo en la subcarpeta correcta
    var file = targetFolder.createFile(blob);
    
    return file.getUrl();
    
  } catch (e) {
    throw new Error("Failed to create PDF in Drive: " + e.message);
  }
}