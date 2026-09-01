/**
 * BACKEND - Logika aplikacji (Apps Script)
 */

// Konfiguracja: ID folderu na Dysku Google, gdzie mają być zapisywane raporty
const TARGET_FOLDER_ID = "1svYq84kpL3m6fenoE5aVHvyCZZO3dvzp";

// Funkcja wymuszająca uprawnienia do Google Drive podczas autoryzacji
function forceDrivePermission() {
  try { DriveApp.getFileById("12345_test_id"); } catch (e) { /* Ignoruj */ }
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SUPER JACHT-dev')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Pobiera początkowe dane do dashboardu (KPI + Wykresy)
 * Scalona wersja obsługująca zarówno spisHours, jak i extraJobs!
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let technicians = [];
  let techniciansDetails = [];
  let techStats = { total: 0, active: 0 };
  const dataSheet = ss.getSheetByName('Data sheet');
  if (dataSheet) {
    const lastRow = dataSheet.getLastRow();
    if (lastRow >= 9) {
      const techData = dataSheet.getRange(9, 1, lastRow - 8, 3).getValues(); 
      techData.forEach(r => {
        let team = String(r[0]).trim();
        let name = String(r[1]).trim();
        let status = String(r[2]).trim().toLowerCase();
        if (name !== "") {
          technicians.push(name);
          techniciansDetails.push({ name: name, team: team, status: status, hoursByDeck: {}, totalHours: 0 });
          techStats.total++;
          if (status === "aktywny" || status === "active") techStats.active++;
        }
      });
    }
  }

  let totalWorkedHours = 0;
  let workedHoursByTask = {};
  let spisHours = []; 

  let extraJobs = {
    totalHours: 0,
    byDeck: {},
    byArea: {}
  };

  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (sheetSpis && sheetSpis.getLastRow() >= 2) {
    const spisData = sheetSpis.getRange(2, 2, sheetSpis.getLastRow() - 1, 7).getDisplayValues(); 
    spisData.forEach(row => {
      let timeStart = String(row[0]).trim(); // B
      let timeEnd = String(row[1]).trim();   // C

      let deck = String(row[2]).trim().toUpperCase(); // D
      let area = String(row[3]).trim(); // E
      let taskName = String(row[4]).trim();  // F
      let taskId = String(row[5]).trim();    // G 
      let isCancelled = String(row[4]).includes("[ANULOWANE]") || taskId.includes("[ANULOWANE]");
      let techNameFromSpis = String(row[6]).trim(); // H

      if (timeStart && timeEnd && !isCancelled) {
        let duration = parseWorkTimeBackend(timeEnd) - parseWorkTimeBackend(timeStart);
        if (duration < 0) duration += 24; 

        duration = Math.round(duration * 100) / 100;
        if (duration > 0) {
          if (taskId.toLowerCase().includes("- extra job")) {
            extraJobs.totalHours += duration;
            if (!extraJobs.byDeck[deck]) extraJobs.byDeck[deck] = 0;
            extraJobs.byDeck[deck] += duration;
            if (!extraJobs.byArea[area]) extraJobs.byArea[area] = 0;
            extraJobs.byArea[area] += duration;
          }

          totalWorkedHours += duration;
          let key = deck + "_" + taskName;
          if(!workedHoursByTask[key]) workedHoursByTask[key] = 0;
          workedHoursByTask[key] += duration;

          spisHours.push({
            hours: duration,
            deck: deck,
            area: area,
            idPrac: taskId, 
            taskName: taskName
          });

          let techObj = techniciansDetails.find(t => t.name.toLowerCase() === techNameFromSpis.toLowerCase());
          if (techObj) {
            techObj.totalHours += duration;
            if (!techObj.hoursByDeck[deck]) techObj.hoursByDeck[deck] = 0;
            techObj.hoursByDeck[deck] += duration;
          }
        }
      }
    });
  }

  totalWorkedHours = Math.round(totalWorkedHours * 10) / 10;
  extraJobs.totalHours = Math.round(extraJobs.totalHours * 10) / 10;
  Object.keys(extraJobs.byDeck).forEach(k => {
    extraJobs.byDeck[k] = Math.round(extraJobs.byDeck[k] * 10) / 10;
  });
  Object.keys(extraJobs.byArea).forEach(k => {
    extraJobs.byArea[k] = Math.round(extraJobs.byArea[k] * 10) / 10;
  });
  techniciansDetails.forEach(tech => {
    tech.totalHours = Math.round(tech.totalHours * 10) / 10;
    Object.keys(tech.hoursByDeck).forEach(d => {
      tech.hoursByDeck[d] = Math.round(tech.hoursByDeck[d] * 10) / 10;
    });
  });

  let tasksTree = {}; 
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido) {
    const lastRow = sheetCandido.getLastRow();

    if (lastRow >= 3) {
      const data = sheetCandido.getRange(1, 1, lastRow, 16).getValues();
      let currentDeck = "";
      let currentArea = "";

      for (let i = 2; i < data.length; i++) {
        let deckVal = String(data[i][2] || "").trim().toUpperCase();
        if (deckVal === "" && String(data[i][1] || "").toLowerCase().includes("dk")) {
          deckVal = String(data[i][1] || "").trim().toUpperCase();
        }
        let areaVal = String(data[i][3] || "").trim();

        let taskName = String(data[i][8] || "").trim();
        if (taskName === "") {
            let taskH = String(data[i][7] || "").trim();
            let taskJ = String(data[i][9] || "").trim();
            if (taskH !== "" && !taskH.toLowerCase().includes("seavia")) taskName = taskH;
            else if (taskJ !== "") taskName = taskJ;
        }

        let plannedStart = data[i][9];
        let plannedEnd = data[i][12];
        let plannedDays = 1;
        if (plannedStart instanceof Date && plannedEnd instanceof Date) {
          let diffTime = Math.abs(plannedEnd - plannedStart);
          plannedDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
        }

        let rawProgress = data[i][15]; 
        let progressNum = 0;
        if (typeof rawProgress === 'number') {
          progressNum = (rawProgress > 0 && rawProgress <= 1) ? Math.round(rawProgress * 100) : Math.round(rawProgress);
        } else if (typeof rawProgress === 'string') {
          progressNum = parseInt(rawProgress.replace('%', '').trim()) || 0;
        }

        if (deckVal !== "") {
          currentDeck = deckVal; currentArea = ""; 
          if (!tasksTree[currentDeck]) tasksTree[currentDeck] = {};
        }
        if (areaVal !== "") currentArea = areaVal;

        if (currentDeck === "" || taskName === "") continue;
        let area = currentArea === "" ? "Brak strefy" : currentArea;

        if (!tasksTree[currentDeck]) tasksTree[currentDeck] = {}; 
        if (!tasksTree[currentDeck][area]) tasksTree[currentDeck][area] = [];

        let taskId = String(data[i][5] || "").trim(); 
        let key = currentDeck + "_" + taskName;
        let taskHours = workedHoursByTask[key] || 0;
        taskHours = Math.round(taskHours * 10) / 10;

        tasksTree[currentDeck][area].push({
          row: i + 1,
          taskId: taskId,
          taskName: taskName,
          progress: progressNum,
          plannedDays: plannedDays,
          workedHours: taskHours
        });
      }
    }
  }

  return {
    technicians: technicians,
    techniciansDetails: techniciansDetails,
    techStats: techStats,           
    totalWorkedHours: totalWorkedHours, 
    tasksTree: tasksTree,
    spisHours: spisHours, 
    extraJobs: extraJobs 
  };
}

function parseWorkTimeBackend(val) {
  if (!val) return 0;
  if (val instanceof Date) return val.getHours() + (val.getMinutes() / 60);
  let str = String(val).replace(',', '.').trim();
  if (str.includes(':')) {
    let parts = str.split(':');
    return parseInt(parts[0]||0) + (parseInt(parts[1]||0)/60);
  }
  return parseFloat(str) || 0;
}

function getTasksForTechnicians(techNamesArray, dateStr) {
  let tasks = [];
  if (!techNamesArray || techNamesArray.length === 0 || !dateStr) return tasks;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis) return tasks;

  const lastRow = sheetSpis.getLastRow();
  if (lastRow < 2) return tasks; 

  const data = sheetSpis.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, lastRow - 1, 1).getValues();

  const searchTechs = techNamesArray.map(t => String(t).trim().toLowerCase());

  for (let i = 0; i < data.length; i++) {
    let rowDate = parseDateToYMD(rawDates[i][0]);
    let rowStart = data[i][1];
    let rowEnd = data[i][2];
    let rowDeck = String(data[i][3]).trim().toUpperCase();
    let rowTask = data[i][5];
    let rowTech = String(data[i][7] || "").trim().toLowerCase();
    let rowTechOriginal = String(data[i][7] || "").trim();

    if (!rowStart || !rowEnd) continue;

    if (searchTechs.includes(rowTech) && rowDate === dateStr) {
      tasks.push({
        techName: rowTechOriginal,
        start: rowStart,
        end: rowEnd,
        deck: rowDeck,
        taskName: rowTask,
        dbRowIndex: i + 2 
      });
    }
  }

  tasks.sort((a, b) => {
    let nameCmp = a.techName.localeCompare(b.techName);
    if (nameCmp !== 0) return nameCmp;
    return a.start.localeCompare(b.start);
  });

  return tasks;
}

function cancelWorkReport(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
    if (sheetSpis) {
      if (rowIndex > 1 && rowIndex <= sheetSpis.getLastRow()) {
        let userEmail = Session.getActiveUser().getEmail() || "Nieznany użytkownik";
        let timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm");
        let cancelMsg = "[ANULOWANE] przez: " + userEmail + " (" + timestamp + ")";
        sheetSpis.getRange(rowIndex, 2, 1, 2).clearContent();
        sheetSpis.getRange(rowIndex, 6).setValue(cancelMsg);
        return { success: true, message: "Wpis anulowano pomyślnie (ślad pozostał w bazie)." };
      }
    }
    return { success: false, message: "Nie udało się zlokalizować wpisu w bazie." };
  } catch (error) {
    return { success: false, message: "Błąd podczas usuwania: " + error.toString() };
  }
}

function addTechnician(name) {
  if (!name || name.trim() === "") return { success: false, message: "Nazwisko nie może być puste!" };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const techName = name.trim();
    const dataSheet = ss.getSheetByName('Data sheet');
    if (!dataSheet) return { success: false, message: "Błąd: Brak zakładki 'Data sheet' w arkuszu." };

    const lastRow = dataSheet.getLastRow();
    let added = false;

    if (lastRow >= 9) {
      const values = dataSheet.getRange("B9:B").getValues();
      for (let i = 0; i < values.length; i++) {
        if (String(values[i][0]).trim() === "") {
          dataSheet.getRange(9 + i, 2).setValue(techName);
          added = true;
          break;
        }
      }
    }
    if (!added) dataSheet.getRange(Math.max(lastRow + 1, 9), 2).setValue(techName);
    return { success: true, message: "Pomyślnie dodano technika: " + techName, newTech: techName };
  } catch (error) {
    return { success: false, message: "Błąd podczas dodawania technika: " + error.toString() };
  }
}

function toggleTechnicianStatus(name, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data sheet');
    if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };

    const lastRow = dataSheet.getLastRow();
    if (lastRow >= 9) {
      const techData = dataSheet.getRange(9, 2, lastRow - 8, 1).getValues();
      for (let i = 0; i < techData.length; i++) {
        if (String(techData[i][0]).trim().toLowerCase() === name.trim().toLowerCase()) {
          dataSheet.getRange(9 + i, 3).setValue(newStatus);
          return { success: true, message: "Status technika zaktualizowany!" };
        }
      }
    }
    return { success: false, message: "Nie znaleziono technika w bazie." };
  } catch (error) {
    return { success: false, message: "Błąd zmiany statusu: " + error.toString() };
  }
}

function assignTechnicianToTeam(techName, teamName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName('Data sheet');
    if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };

    const lastRow = dataSheet.getLastRow();
    if (lastRow >= 9) {
      const techData = dataSheet.getRange(9, 2, lastRow - 8, 1).getValues(); 
      for (let i = 0; i < techData.length; i++) {
        if (String(techData[i][0]).trim().toLowerCase() === techName.trim().toLowerCase()) {
          dataSheet.getRange(9 + i, 1).setValue(teamName);
          return { success: true, message: "Drużyna zaktualizowana!" };
        }
      }
    }
    return { success: false, message: "Nie znaleziono technika w bazie." };
  } catch (error) {
    return { success: false, message: "Błąd zmiany drużyny: " + error.toString() };
  }
}

function parseDateToYMD(val) {
  if (!val) return null;
  if (val instanceof Date) {
    let d = val.getDate().toString().padStart(2, '0');
    let m = (val.getMonth() + 1).toString().padStart(2, '0');
    let y = val.getFullYear();
    return y + "-" + m + "-" + d;
  }
  if (typeof val === 'string') {
    let str = val.trim();
    let m1 = str.match(/(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/);
    if (m1) {
      let d = m1[1].padStart(2, '0');
      let m = m1[2].padStart(2, '0');
      let y = m1[3];
      if (y.length === 2) y = "20" + y;
      return y + "-" + m + "-" + d;
    }
    if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;
  }
  return null;
}

function getIsoWeekNumber(dateValue) {
  if (!dateValue) return null;
  const date = dateValue instanceof Date ? new Date(dateValue.getTime()) : new Date(dateValue);
  if (isNaN(date.getTime())) return null;

  const utcDate = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = utcDate.getUTCDay() || 7;
  utcDate.setUTCDate(utcDate.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(utcDate.getUTCFullYear(), 0, 1));
  const weekNumber = Math.ceil((((utcDate - yearStart) / 86400000) + 1) / 7);
  return weekNumber;
}

function submitWorkReport(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let idPrac = "";
    const sheetCandido = ss.getSheetByName('Candido');
    if (formData.isCustomTask) {
      if (sheetCandido) {
        idPrac = formData.customTaskId || "Dodatkowe";
        if (formData.isExtraJob && !idPrac.toUpperCase().includes("- EXTRA JOB")) {
          idPrac += " - Extra Job";
        }
        let formattedDate = formData.workDate;
        let dParts = formData.workDate.split('-');
        if (dParts.length === 3) {
           formattedDate = dParts[1] + "/" + dParts[2] + "/" + dParts[0];
        }
        let newRow = new Array(16).fill("");
        newRow[2] = formData.deck;
        newRow[3] = formData.area;
        newRow[5] = idPrac;
        newRow[7] = "Seavia";
        newRow[8] = formData.taskName;
        newRow[9] = formattedDate;
        newRow[12] = formattedDate;
        if (formData.progress) {
          newRow[15] = formData.progress + "%";
        }
        sheetCandido.appendRow(newRow);
      }
    } else {
      if (formData.taskRow) {
        if (sheetCandido) {
          idPrac = sheetCandido.getRange(parseInt(formData.taskRow), 6).getValue();
          if (!idPrac || idPrac === "") idPrac = "Wiersz " + formData.taskRow;
          if (formData.progress) {
            sheetCandido.getRange(parseInt(formData.taskRow), 16).setValue(formData.progress + "%");
          }
        }
      }
    }
    let sheetSpis = ss.getSheetByName('Spis wykonanych prac');
    if (!sheetSpis) sheetSpis = ss.insertSheet('Spis wykonanych prac');
    if (sheetSpis.getLastRow() === 0) {
      sheetSpis.appendRow(["Data", "Godzina rozpoczęcia pracy", "Godzina zakończenia pracy", "Pokład", "Typ prac", "Nazwa prac", "ID prac", "Technik"]);
      sheetSpis.getRange(1, 1, 1, 8).setFontWeight("bold");
    }
    let techsArray = formData.techniciansArray;
    if (!techsArray || !Array.isArray(techsArray)) {
      return { success: false, message: "Brak techników do przypisania." };
    }
    techsArray.forEach(tech => {
       sheetSpis.appendRow([
         formData.workDate, 
         formData.timeStart, 
         formData.timeEnd, 
         formData.deck ? String(formData.deck).trim().toUpperCase() : "",
         formData.area,
         formData.taskName, 
         idPrac, 
         tech
       ]);
    });
    let msgSuffix = techsArray.length > 1 ? " (Dla " + techsArray.length + " techników)" : '';
    return { success: true, message: "Pomyślnie zapisano godziny i postęp zadania!" + msgSuffix };
  } catch (error) {
    return { success: false, message: "Błąd podczas zapisu: " + error.toString() };
  }
}

function getCandidoWeeklyReport(dateStr, dateToStr, includeBreaks) {
  if (arguments.length === 2 && typeof dateToStr === 'boolean') {
    includeBreaks = dateToStr;
    dateToStr = null;
  }
  if (!dateStr) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, weekInfo: {} };

  const startDate = new Date(dateStr);
  const endDate = dateToStr ? new Date(dateToStr) : new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 7);
  const weekDates = [];
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    weekDates.push(parseDateToYMD(new Date(d)));
  }

  let weekNum = getIsoWeekNumber(startDate);
  if (!weekNum) {
    weekNum = dateStr;
  }

  let reportData = {
    settings: { includeBreaks: includeBreaks },
    weekInfo: {
      weekNumber: weekNum,
      start: weekDates[0],
      end: weekDates[weekDates.length - 1]
    },
    decks: {},
    allTasks: [],
    extraTasks: [],
    extraDeck: { tasks: {} }
  };

  let uniqueTasks = new Set();
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido && sheetCandido.getLastRow() >= 3) {
    const candidoTasks = sheetCandido.getRange(3, 9, sheetCandido.getLastRow() - 2, 1).getValues();
    candidoTasks.forEach(row => {
      let t = String(row[0]).trim();
      if (t && !t.includes("[ANULOWANE]") && !t.toLowerCase().includes("extra job")) {
        uniqueTasks.add(t);
      }
    });
  }

  const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < spisData.length; i++) {
    let rowDate = parseDateToYMD(rawDates[i][0]);
    let dayIndex = weekDates.indexOf(rowDate);
    if (dayIndex === -1) continue;

    let rowStart = spisData[i][1];
    let rowEnd = spisData[i][2];
    let rowDeck = String(spisData[i][3]).trim().toUpperCase();
    let rowTask = String(spisData[i][5]).trim();
    let taskId = String(spisData[i][6]).trim();
    let isCancelled = rowTask.includes("[ANULOWANE]") || taskId.includes("[ANULOWANE]");
    let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break");

    if (isBreak && !includeBreaks) continue;
    if (!rowStart || !rowEnd || isCancelled || !rowDeck || !rowTask) continue;

    if (rowTask.toLowerCase().includes("site manager") || rowTask.toLowerCase().includes("meetings") || rowTask.toLowerCase() === "supervision") {
       rowTask = "Supervision";
    }

    let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
    if (duration < 0) duration += 24; 
    duration = Math.round(duration * 100) / 100;

    if (duration > 0) {
      let isExtraJob = taskId.toLowerCase().includes("- extra job");
      if (isExtraJob) {
        let extraTaskKey = rowDeck + ", " + rowTask;
        if (!reportData.extraDeck.tasks[extraTaskKey]) {
          reportData.extraDeck.tasks[extraTaskKey] = [0, 0, 0, 0, 0, 0, 0];
        }
        reportData.extraDeck.tasks[extraTaskKey][dayIndex] += duration;
      } else {
        if (rowDeck === "TRAVEL") {
          rowTask = "Travel";
        }
        if (!reportData.decks[rowDeck]) {
           reportData.decks[rowDeck] = { tasks: {} };
        }
        if (!reportData.decks[rowDeck].tasks[rowTask]) {
           reportData.decks[rowDeck].tasks[rowTask] = [0, 0, 0, 0, 0, 0, 0];
        }
        reportData.decks[rowDeck].tasks[rowTask][dayIndex] += duration;
        if(rowTask === "Supervision") {
          uniqueTasks.add(rowTask);
        }
      }
    }
  }

  reportData.allTasks = Array.from(uniqueTasks).sort((a, b) => {
    return a.localeCompare(b);
  });

  let extraTasksSet = new Set();
  Object.keys(reportData.extraDeck.tasks).forEach(k => {
    let total = reportData.extraDeck.tasks[k].reduce((a, b) => a + b, 0);
    if (total > 0) {
      extraTasksSet.add(k);
    }
  });
  reportData.extraTasks = Array.from(extraTasksSet).sort((a, b) => a.localeCompare(b));

  Object.keys(reportData.decks).forEach(d => {
     Object.keys(reportData.decks[d].tasks).forEach(t => {
       reportData.decks[d].tasks[t] = reportData.decks[d].tasks[t].map(h => Math.round(h * 10) / 10);
     });
  });
  Object.keys(reportData.extraDeck.tasks).forEach(t => {
    reportData.extraDeck.tasks[t] = reportData.extraDeck.tasks[t].map(h => Math.round(h * 10) / 10);
  });
  return reportData;
}

function getTimeByWeekReport(dateStr, dateToStr, includeBreaks) {
  if (arguments.length === 2 && typeof dateToStr === 'boolean') {
    includeBreaks = dateToStr;
    dateToStr = null;
  }
  if (!dateStr) return { decks: {}, allTasks: [], allWeeks: [] };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, allTasks: [], allWeeks: [] };

  const projectTasks = new Set();
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido && sheetCandido.getLastRow() >= 3) {
    const candidoTasks = sheetCandido.getRange(3, 9, sheetCandido.getLastRow() - 2, 1).getValues();
    candidoTasks.forEach(row => {
      const taskName = String(row[0]).trim();
      if (taskName && !taskName.includes("[ANULOWANE]") && !taskName.toLowerCase().includes("extra job")) {
        projectTasks.add(taskName);
      }
    });
  }

  const startDate = new Date(dateStr);
  const endDate = dateToStr ? new Date(dateToStr) : new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + 7);
  const rangeDates = [];
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    rangeDates.push(parseDateToYMD(new Date(d)));
  }

  let reportData = {
    settings: { includeBreaks: includeBreaks },
    decks: {},
    allTasks: Array.from(projectTasks),
    allWeeks: new Set(),
    extraTasks: [],
    extraDeck: { tasks: {} }
  };

  const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < spisData.length; i++) {
    let rowDate = parseDateToYMD(rawDates[i][0]);
    if (rangeDates.indexOf(rowDate) === -1) continue;
    let rowStart = spisData[i][1];
    let rowEnd = spisData[i][2];
    let rowDeck = String(spisData[i][3]).trim().toUpperCase();
    let rowTask = String(spisData[i][5]).trim();
    let taskId = String(spisData[i][6]).trim();
    let isCancelled = rowTask.includes("[ANULOWANE]") || taskId.includes("[ANULOWANE]");
    let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break");

    if (isBreak && !includeBreaks) continue;
    if (!rowStart || !rowEnd || isCancelled || !rowDeck || !rowTask) continue;

    if (rowTask.toLowerCase().includes("site manager") || rowTask.toLowerCase().includes("meetings") || rowTask.toLowerCase() === "supervision") {
       rowTask = "Supervision";
    }

    let rowWeekNum = getIsoWeekNumber(new Date(rowDate + 'T00:00:00'));
    if (!rowWeekNum) {
      rowWeekNum = getIsoWeekNumber(startDate);
    }
    let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
    if (duration < 0) duration += 24;
    duration = Math.round(duration * 100) / 100;

    if (duration > 0) {
      let isExtraJob = taskId.toLowerCase().includes("- extra job");
      if (isExtraJob) {
        let extraTaskKey = rowDeck + ", " + rowTask;
        if (!reportData.extraDeck.tasks[extraTaskKey]) {
          reportData.extraDeck.tasks[extraTaskKey] = { total: 0, weeklyHours: {} };
        }
        if (!reportData.extraDeck.tasks[extraTaskKey].weeklyHours[rowWeekNum]) {
          reportData.extraDeck.tasks[extraTaskKey].weeklyHours[rowWeekNum] = 0;
        }
        reportData.extraDeck.tasks[extraTaskKey].weeklyHours[rowWeekNum] += duration;
        reportData.extraDeck.tasks[extraTaskKey].total += duration;
      } else {
        if (rowDeck === "TRAVEL") {
          rowTask = "Travel";
        }
        if (!reportData.decks[rowDeck]) {
           reportData.decks[rowDeck] = { tasks: {} };
        }
        if (!reportData.decks[rowDeck].tasks[rowTask]) {
           reportData.decks[rowDeck].tasks[rowTask] = { total: 0, weeklyHours: {} };
        }
        if (!reportData.decks[rowDeck].tasks[rowTask].weeklyHours[rowWeekNum]) {
          reportData.decks[rowDeck].tasks[rowTask].weeklyHours[rowWeekNum] = 0;
        }
        reportData.decks[rowDeck].tasks[rowTask].weeklyHours[rowWeekNum] += duration;
        reportData.decks[rowDeck].tasks[rowTask].total += duration;
        if(rowTask === "Supervision") {
          if (reportData.allTasks.indexOf(rowTask) === -1) {
            reportData.allTasks.push(rowTask);
          }
        }
      }
      reportData.allWeeks.add(rowWeekNum);
    }
  }

  reportData.allTasks = reportData.allTasks.sort((a, b) => a.localeCompare(b));
  reportData.allWeeks = Array.from(reportData.allWeeks).sort((a, b) => a - b);

  let extraTasksSet = new Set();
  Object.keys(reportData.extraDeck.tasks).forEach(k => {
    if (reportData.extraDeck.tasks[k].total > 0) {
      extraTasksSet.add(k);
    }
  });
  reportData.extraTasks = Array.from(extraTasksSet).sort((a, b) => a.localeCompare(b));

  Object.keys(reportData.decks).forEach(d => {
     Object.keys(reportData.decks[d].tasks).forEach(t => {
       let task = reportData.decks[d].tasks[t];
       task.total = Math.round(task.total * 10) / 10;
       Object.keys(task.weeklyHours).forEach(w => {
         task.weeklyHours[w] = Math.round(task.weeklyHours[w] * 10) / 10;
       });
     });
  });
  Object.keys(reportData.extraDeck.tasks).forEach(t => {
     let task = reportData.extraDeck.tasks[t];
     task.total = Math.round(task.total * 10) / 10;
     Object.keys(task.weeklyHours).forEach(w => {
       task.weeklyHours[w] = Math.round(task.weeklyHours[w] * 10) / 10;
     });
  });
  return reportData;
}

function getTechnicianWeeklyReports(dateStr, dateToStr, includeBreaks) {
  if (arguments.length === 2 && typeof dateToStr === 'boolean') {
    includeBreaks = dateToStr;
    dateToStr = null;
  }
  if (!dateStr) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return null;

  const fromDate = new Date(dateStr);
  const toDate = dateToStr ? new Date(dateToStr) : new Date(fromDate);
  const rangeDates = [];
  const currentDate = new Date(fromDate);
  const endDate = new Date(toDate);

  while (currentDate <= endDate) {
    rangeDates.push(parseDateToYMD(new Date(currentDate)));
    currentDate.setDate(currentDate.getDate() + 1);
  }

  let dateRangeLabel = rangeDates.length > 1 ? `${rangeDates[0]} / ${rangeDates[rangeDates.length - 1]}` : rangeDates[0];
  let weekNum = getIsoWeekNumber(fromDate);
  if (!weekNum) {
    weekNum = rangeDates[0] || dateStr;
  }
  let year = fromDate.getFullYear();

  let reportData = {
    weekInfo: {
      weekNumber: weekNum,
      year: year,
      dates: rangeDates,
      rangeLabel: dateRangeLabel
    },
    technicians: {}
  };

  const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < spisData.length; i++) {
    let rowDate = parseDateToYMD(rawDates[i][0]);
    let dayIndex = rangeDates.indexOf(rowDate);
    if (dayIndex === -1) continue;

    let rowStart = String(spisData[i][1]).trim();
    let rowEnd = String(spisData[i][2]).trim();
    let rowDeck = String(spisData[i][3]).trim().toUpperCase();
    let rowArea = String(spisData[i][4]).trim();
    let rowTask = String(spisData[i][5]).trim();
    let rowRoomId = String(spisData[i][6]).trim();
    let techName = String(spisData[i][7]).trim();
    let isCancelled = rowTask.includes("[ANULOWANE]") || rowTask.includes("ANULOWANE");
    let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break") || rowTask.toLowerCase().includes("brake");

    if (isBreak && !includeBreaks) continue;
    if (!rowStart || !rowEnd || isCancelled || !techName || !rowTask) continue;

    let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
    if (duration < 0) duration += 24;
    duration = Math.round(duration * 100) / 100;

    if (duration > 0) {
      if (!reportData.technicians[techName]) {
        reportData.technicians[techName] = {
          totalHours: 0,
          days: Array(rangeDates.length).fill(null).map(() => [])
        };
      }
      reportData.technicians[techName].days[dayIndex].push({
        start: rowStart,
        end: rowEnd,
        hours: duration,
        task: rowTask,
        deck: rowDeck,
        area: rowArea,
        roomId: rowRoomId
      });
      reportData.technicians[techName].totalHours += duration;
    }
  }

  Object.keys(reportData.technicians).forEach(techName => {
    reportData.technicians[techName].totalHours = Math.round(reportData.technicians[techName].totalHours * 10) / 10;
  });

  // URUCHOMIENIE W TLE: Synchronizacja z REV OCEAN - timesheet
  syncWeeklyTimesheet(reportData);

  return reportData;
}

function generateStyledExcelReport(reportData, reportType) {
  try {
    const isWeekly = reportType === 'weekly';
    const baseName = "Candido_" + (isWeekly ? reportData.weekInfo.weekNumber : 'Annual') + "_Report";
    const tempFileName = baseName + "_TEMP_" + Date.now();
    const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);

    const spreadsheet = SpreadsheetApp.create(tempFileName);
    const tempFile = DriveApp.getFileById(spreadsheet.getId());
    tempFile.moveTo(targetFolder);

    const sheet = spreadsheet.getSheets()[0];
    sheet.setName(isWeekly ? "Weekly Report" : "Annual Report");
    const headerColor = "#f1f5f9";
    const totalHeaderColor = "#e2e8f0";
    const seaviaBlue = "#0f3460";
    const lightGreen = "#dcfce7";
    const green = "#bbf7d0";
    const borderColor = "#cbd5e1";
    let currentRow = 1;

    sheet.getRange(currentRow, 1, 1, 2).setValues([["Project name:", "REV OCEAN 19577"]]).setFontWeight("bold");
    sheet.getRange(currentRow, 2).setFontColor(seaviaBlue).setHorizontalAlignment("right");
    currentRow++;

    if (isWeekly) {
      sheet.getRange(currentRow, 1, 1, 2).setValues([["Week no.:", reportData.weekInfo.weekNumber]]).setFontWeight("bold");
      sheet.getRange(currentRow, 2).setFontColor(seaviaBlue).setHorizontalAlignment("right");
      currentRow++;
      sheet.getRange(currentRow, 1, 1, 2).setValues([["Date range:", reportData.weekInfo.start + " / " + reportData.weekInfo.end]]).setFontWeight("bold");
      sheet.getRange(currentRow, 2).setHorizontalAlignment("right");
      currentRow++;
    } else {
      sheet.getRange(currentRow, 1, 1, 2).setValues([["Data generacji:", new Date().toLocaleDateString('pl-PL')]]).setFontWeight("bold");
      sheet.getRange(currentRow, 2).setHorizontalAlignment("right");
      currentRow++;
    }

    sheet.getRange(currentRow, 1, 1, 2).setValues([["Paid Breaks:", reportData.settings.includeBreaks ? "YES" : "NO"]]).setFontWeight("bold");
    sheet.getRange(currentRow, 2).setFontColor(reportData.settings.includeBreaks ? "#166534" : "#94a3b8").setHorizontalAlignment("right");
    currentRow += 2;

    Object.keys(reportData.decks).sort().forEach(deckName => {
      let deckData = reportData.decks[deckName];
      if (!deckData.tasks || Object.keys(deckData.tasks).length === 0) return;

      sheet.getRange(currentRow, 1).setValue(isWeekly ? "WEEKLY SUM " + deckName : "ANNUAL TIME REPORT - " + deckName).setFontWeight("bold").setFontColor(seaviaBlue);
      currentRow++;

      let headers = ["No.", "Tasks", "Total"];
      let columnCount;
      if (isWeekly) {
        headers.push('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
        columnCount = 10;
      } else {
        let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
        sortedWeeks.forEach(w => headers.push("W" + w));
        columnCount = 3 + sortedWeeks.length;
      }

      if (columnCount > 0) {
        let headerRange = sheet.getRange(currentRow, 1, 1, columnCount);
        headerRange.setValues([headers]).setBackground(headerColor).setFontWeight("bold");
        sheet.getRange(currentRow, 3).setBackground(totalHeaderColor);
      }
      currentRow++;

      const renderRow = (taskName, index) => {
        let rowData = [index, taskName];
        if (isWeekly) {
          let hours = (deckData.tasks && deckData.tasks[taskName]) ? deckData.tasks[taskName] : Array(7).fill(0);
          let total = hours.reduce((a, b) => a + b, 0);
          rowData.push(total > 0 ? total : 0, ...hours.map(h => h > 0 ? h : 0));
        } else {
          let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
          let task = (deckData.tasks && deckData.tasks[taskName]) ? deckData.tasks[taskName] : { total: 0, weeklyHours: {} };
          rowData.push(task.total > 0 ? task.total : 0);
          sortedWeeks.forEach(w => {
            let h = task.weeklyHours[w] || 0;
            rowData.push(h > 0 ? h : 0);
          });
        }

        if (rowData.length > 2) {
          let rowRange = sheet.getRange(currentRow, 1, 1, rowData.length);
          rowRange.setValues([rowData]).setNumberFormat("0.0");
          sheet.getRange(currentRow, 1, 1, 2).setNumberFormat("0");
          sheet.getRange(currentRow, 2).setHorizontalAlignment("left");

          let totalCell = sheet.getRange(currentRow, 3);
          if (rowData[2] > 0) totalCell.setBackground(green).setFontWeight("bold"); else totalCell.setBackground(totalHeaderColor);

          for (let i = 4; i <= rowData.length; i++) {
            let cell = sheet.getRange(currentRow, i);
            if (cell.getValue() > 0) cell.setBackground(lightGreen).setFontWeight("bold");
          }
        }
        currentRow++;
      };

      if (deckName !== "TRAVEL") {
        renderRow("Supervision", 0);
      }

      let taskIndex = 1;
      let filteredTasks = reportData.allTasks;
      if (!reportData.settings.includeBreaks) {
        filteredTasks = filteredTasks.filter(t => !t.toLowerCase().includes('break') && !t.toLowerCase().includes('przerwa'));
      }

      if (deckName === "TRAVEL") {
        filteredTasks = filteredTasks.filter(t => t.toLowerCase() === "travel");
      } else {
        filteredTasks = filteredTasks.filter(t => t.toLowerCase() !== "travel");
      }

      filteredTasks.forEach(taskName => {
        if (taskName !== "Supervision") {
          renderRow(taskName, taskIndex++);
        }
      });

      let renderedRowCount = taskIndex;
      if (deckName !== "TRAVEL") {
        renderedRowCount += 1;
      } else {
        renderedRowCount = filteredTasks.length;
      }
      if (renderedRowCount > 0) {
        let tableStartRow = currentRow - renderedRowCount - 1; 
        let numRows = renderedRowCount + 1; 
        if (tableStartRow > 0 && numRows > 0 && columnCount > 0) {
          let tableRange = sheet.getRange(tableStartRow, 1, numRows, columnCount);
          tableRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
        }
      }
      sheet.setColumnWidth(2, 250);
      currentRow += 2;
    });

    if (reportData.extraTasks && reportData.extraTasks.length > 0) {
      sheet.getRange(currentRow, 1).setValue("PRACE DODATKOWE (EXTRA JOBS)").setFontWeight("bold").setFontColor("#dc2626");
      currentRow++;

      let headers = ["No.", "Tasks", "Total"];
      let columnCount;
      if (isWeekly) {
        headers.push('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
        columnCount = 10;
      } else {
        let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
        sortedWeeks.forEach(w => headers.push("W" + w));
        columnCount = 3 + sortedWeeks.length;
      }

      if (columnCount > 0) {
        let headerRange = sheet.getRange(currentRow, 1, 1, columnCount);
        headerRange.setValues([headers]).setBackground(headerColor).setFontWeight("bold");
        sheet.getRange(currentRow, 3).setBackground(totalHeaderColor);
      }
      currentRow++;

      let extraTaskIndex = 1;
      reportData.extraTasks.forEach(taskName => {
        let rowData = [extraTaskIndex, taskName];
        if (isWeekly) {
          let hours = reportData.extraDeck.tasks[taskName] || Array(7).fill(0);
          let total = hours.reduce((a, b) => a + b, 0);
          rowData.push(total > 0 ? total : 0, ...hours.map(h => h > 0 ? h : 0));
        } else {
          let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
          let task = reportData.extraDeck.tasks[taskName] || { total: 0, weeklyHours: {} };
          rowData.push(task.total > 0 ? task.total : 0);
          sortedWeeks.forEach(w => {
            let h = task.weeklyHours[w] || 0;
            rowData.push(h > 0 ? h : 0);
          });
        }

        if (rowData.length > 2) {
          let rowRange = sheet.getRange(currentRow, 1, 1, rowData.length);
          rowRange.setValues([rowData]).setNumberFormat("0.0");
          sheet.getRange(currentRow, 1, 1, 2).setNumberFormat("0");
          sheet.getRange(currentRow, 2).setHorizontalAlignment("left");
          rowRange.setFontColor("#dc2626"); 

          let totalCell = sheet.getRange(currentRow, 3);
          if (rowData[2] > 0) totalCell.setBackground(green).setFontWeight("bold"); else totalCell.setBackground(totalHeaderColor);

          for (let i = 4; i <= rowData.length; i++) {
            let cell = sheet.getRange(currentRow, i);
            if (cell.getValue() > 0) cell.setBackground(lightGreen).setFontWeight("bold");
          }
        }
        currentRow++;
        extraTaskIndex++;
      });

      let renderedRowCount = reportData.extraTasks.length;
      if (renderedRowCount > 0) {
        let tableStartRow = currentRow - renderedRowCount - 1;
        let numRows = renderedRowCount + 1;
        if (tableStartRow > 0 && numRows > 0 && columnCount > 0) {
          let tableRange = sheet.getRange(tableStartRow, 1, numRows, columnCount);
          tableRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
        }
      }
      sheet.setColumnWidth(2, 250);
      currentRow += 2;
    }

    SpreadsheetApp.flush();

    const url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Błąd podczas konwersji arkusza do formatu Excel (.xlsx): " + response.getContentText());
    }

    const excelBlob = response.getBlob().setName(`${baseName}.xlsx`);
    const excelFile = targetFolder.createFile(excelBlob);
    tempFile.setTrashed(true);

    const downloadUrl = `https://drive.google.com/uc?export=download&id=${excelFile.getId()}`;
    return {
      fileUrl: downloadUrl,
      fileName: excelFile.getName()
    };
  } catch (e) {
    Logger.log("Krytyczny błąd w generateStyledExcelReport: " + e.message + " Stack: " + e.stack);
    return { error: "Wystąpił błąd po stronie serwera: " + e.message };
  }
}

function deleteTeam(teamName) {
 if (!teamName || teamName.trim() === "") {
   return { success: false, message: "Nazwa drużyny nie może być pusta!" };
 }
 try {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const dataSheet = ss.getSheetByName('Data sheet');
   if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };
   const lastRow = dataSheet.getLastRow();
   let updatedCount = 0;
   if (lastRow >= 9) {
     const range = dataSheet.getRange(9, 1, lastRow - 8, 1);
     const values = range.getValues();
     for (let i = 0; i < values.length; i++) {
       if (String(values[i][0]).trim().toLowerCase() === teamName.trim().toLowerCase()) {
         dataSheet.getRange(9 + i, 1).setValue(""); 
         updatedCount++;
       }
     }
   }
   return { success: true, message: `Pomyślnie usunięto drużynę ${teamName} (odpięto ${updatedCount} techników).`, updatedCount: updatedCount };
 } catch (error) {
   return { success: false, message: "Błąd podczas usuwania drużyny: " + error.toString() };
 }
}

function renameTeam(oldName, newName) {
 if (!oldName || !newName || newName.trim() === "") {
   return { success: false, message: "Nazwy drużyn nie mogą być puste!" };
 }
 try {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const dataSheet = ss.getSheetByName('Data sheet');
   if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };
   const lastRow = dataSheet.getLastRow();
   let updatedCount = 0;
   if (lastRow >= 9) {
     const range = dataSheet.getRange(9, 1, lastRow - 8, 1);
     const values = range.getValues();
     for (let i = 0; i < values.length; i++) {
       if (String(values[i][0]).trim().toLowerCase() === oldName.trim().toLowerCase()) {
         dataSheet.getRange(9 + i, 1).setValue(newName.trim());
         updatedCount++;
       }
     }
   }
   return { success: true, message: `Pomyślnie zmieniono nazwę drużyny z ${oldName} na ${newName} (zaktualizowano ${updatedCount} techników).`, updatedCount: updatedCount };
 } catch (error) {
   return { success: false, message: "Błąd podczas zmiany nazwy drużyny: " + error.toString() };
 }
}

function uploadCostFile(base64Data, fileName, costName, costDate, fileIndex = 0, totalFiles = 1) {
  try {
    const folderId = "1dtvg-bx_ePiMfsuf2kf0vGadlzNgHeeI";
    const folder = DriveApp.getFolderById(folderId);
    
    const parts = base64Data.split(',');
    const contentType = parts[0].match(/:(.*?);/)[1];
    const rawBase64 = parts[1];
    const decoded = Utilities.base64Decode(rawBase64);
    
    // 1. Wyciąganie oryginalnego rozszerzenia (np. .jpg, .pdf)
    let extension = "";
    const lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex !== -1) {
      extension = fileName.substring(lastDotIndex);
    } else {
      if (contentType.includes('pdf')) extension = '.pdf';
      else if (contentType.includes('jpeg') || contentType.includes('jpg')) extension = '.jpg';
      else if (contentType.includes('png')) extension = '.png';
    }
    
    // 2. Czyszczenie nazwy (usuwamy tylko niedozwolone znaki w nazwach plików, pozwalamy na spacje)
    const cleanCostName = costName.replace(/[\/\\?%*:|"<>]/g, '').trim();
    
    // 3. Dodanie numeracji, jeśli w jednej paczce jest więcej niż 1 plik 
    let suffix = "";
    if (totalFiles > 1) {
      suffix = ` - plik ${fileIndex + 1}`;
    }
    
    // 4. Budowa finalnej nowej nazwy
    const newFileName = `${costDate} ${cleanCostName}${suffix}${extension}`;
    
    const blob = Utilities.newBlob(decoded, contentType, newFileName);
    const file = folder.createFile(blob);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetKoszty = ss.getSheetByName('Koszty');
    if (!sheetKoszty) {
      sheetKoszty = ss.insertSheet('Koszty');
      sheetKoszty.appendRow(["Data dodania", "Data kosztu", "Nazwa kosztu", "Link do pliku", "Użytkownik"]);
      sheetKoszty.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
    const userEmail = Session.getActiveUser().getEmail() || "Nieznany użytkownik";
    // Zapis do arkusza Google z zachowaniem oryginalnie wpisanej nazwy
    sheetKoszty.appendRow([new Date(), costDate, costName, file.getUrl(), userEmail]);
    
    return { success: true, message: "Koszt został pomyślnie zapisany w bazie i folderze!", fileUrl: file.getUrl() };
  } catch (e) {
    Logger.log("Błąd w uploadCostFile: " + e.message);
    return { success: false, message: "Błąd serwera podczas zapisywania pliku: " + e.toString() };
  }
}

// =======================================================================
// NOWE: SYNCHRONIZACJA Z REV OCEAN - TIMESHEET (W tle)
// =======================================================================
function syncWeeklyTimesheet(reportData) {
  if (!reportData || !reportData.weekInfo || !reportData.technicians) return { success: false, message: "Brak danych." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTime = ss.getSheetByName('REV OCEAN - timesheet');
  if (!sheetTime) return { success: false, message: "Brak zakładki 'REV OCEAN - timesheet'." };

  // Wymuszamy pobranie pełnego zakresu
  const timeData = sheetTime.getDataRange().getValues();
  if (timeData.length < 3) return { success: false, message: "Pusty arkusz timesheet." };

  const timezone = ss.getSpreadsheetTimeZone();

  // 1. Zabezpieczone mapowanie dat (nagłówki) z Wiersza 3
  let dateToColMap = {};
  if (timeData[2]) {
    for (let c = 3; c < timeData[2].length; c++) {
      let rawDate = timeData[2][c];
      let dStr = "";
      
      if (rawDate instanceof Date) {
         dStr = Utilities.formatDate(rawDate, timezone, "yyyy-MM-dd");
      } else if (typeof rawDate === 'string') {
         let parsed = parseDateToYMD(rawDate);
         if (parsed) dStr = parsed;
      }

      if (dStr) {
        dateToColMap[dStr] = c + 1; // +1 dla metody getRange()
      }
    }
  }

  // 2. Bezpieczne mapowanie wszystkich istniejących techników z kolumny B!
  let techToRowMap = {};
  
  // Skanujemy CAŁY plik, patrząc TYLKO NA KOLUMNĘ B (timeData[r][1])
  for (let r = 3; r < timeData.length; r++) {
    if (!timeData[r]) continue;

    let tName = String(timeData[r][1] || "").trim(); // TYLKO KOLUMNA B
    
    if (tName) {
      let normalized = tName.replace(/\s+/g, '').toLowerCase();
      techToRowMap[normalized] = r + 1;
    }
  }

  let updatesCount = 0;
  let addedTechs = 0;

  // Najpierw: Czyścimy komórki z obecnego tygodnia dla WSZYSTKICH techników z reportData
  Object.keys(reportData.technicians).forEach(techName => {
    let normalizedTechKey = String(techName).replace(/\s+/g, '').toLowerCase();
    let rowIdx = techToRowMap[normalizedTechKey];
    
    if (rowIdx) {
      reportData.weekInfo.dates.forEach(dDate => {
        let colIdx = dateToColMap[dDate];
        if (colIdx) {
           let cell = sheetTime.getRange(rowIdx, colIdx);
           cell.setValue("");
           cell.setBackground(null);
        }
      });
    }
  });


  // 3. Główna pętla wpisująca godziny i ewentualnie nowe osoby na czysto
  Object.keys(reportData.technicians).forEach(techName => {
    
    let normalizedTechKey = String(techName).replace(/\s+/g, '').toLowerCase();
    let rowIdx = techToRowMap[normalizedTechKey];

    // Jeśli technika nie ma na liście w kolumnie B - dopisujemy!
    if (!rowIdx) {
       let targetSheetLastRow = sheetTime.getLastRow();
       
       // Szukamy pierwszego pustego miejsca od dołu patrząc WYŁĄCZNIE w kolumnę B.
       let insertRow = targetSheetLastRow + 1;
       const columnBValues = sheetTime.getRange("B1:B" + targetSheetLastRow).getValues();
       
       for(let i = columnBValues.length - 1; i >= 0; i--) {
         if (String(columnBValues[i][0]).trim() !== "") {
           insertRow = i + 2; 
           break;
         }
       }

       // Zapis do nowej linii - ZAWSZE w Kolumnie B
       sheetTime.getRange(insertRow, 1).setValue(insertRow - 3); 
       sheetTime.getRange(insertRow, 2).setValue(techName.toUpperCase()); 
       sheetTime.getRange(insertRow, 2).setFontWeight("bold");

       techToRowMap[normalizedTechKey] = insertRow; 
       rowIdx = insertRow;
       addedTechs++;
    }

    let techObj = reportData.technicians[techName];

    // Iteracja po 7 dniach analizowanego tygodnia
    techObj.days.forEach((tasksArray, dayIndex) => {
      if (tasksArray.length === 0) return;

      let dDate = reportData.weekInfo.dates[dayIndex];
      let colIdx = dateToColMap[dDate];

      if (!colIdx) return; 

      let dailyTotalHours = 0;
      let hasTravel = false;

      tasksArray.forEach(t => {
        dailyTotalHours += t.hours;
        if (t.task.toLowerCase().includes("travel")) {
          hasTravel = true;
        }
      });

      if (dailyTotalHours > 0) {
        let cell = sheetTime.getRange(rowIdx, colIdx);
        cell.setValue(dailyTotalHours).setHorizontalAlignment("center");
        
        // Oznaczanie Travel
        if (hasTravel) {
          cell.setBackground("#FFF2CC");
        }
        updatesCount++;
      }
    });
  });

  // Wymuszamy, żeby Google od razu to narysował w bazie
  SpreadsheetApp.flush(); 

  return { success: true, updates: updatesCount, newTechs: addedTechs };
}

// =======================================================================
// EKSPORT ZAKŁADKI REV OCEAN DO EXCELA
// =======================================================================
function exportRevOceanTimesheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('REV OCEAN - timesheet');
    if (!sheet) return { error: "Brak zakładki 'REV OCEAN - timesheet' w bazie." };

    const tempFileName = "Eksport_Timesheet_" + new Date().getTime();
    const tempSs = SpreadsheetApp.create(tempFileName);
    
    sheet.copyTo(tempSs).setName("REV OCEAN - timesheet");
    
    const defaultSheet = tempSs.getSheets()[0];
    if(tempSs.getSheets().length > 1) {
      tempSs.deleteSheet(defaultSheet);
    }

    const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    const tempFile = DriveApp.getFileById(tempSs.getId());
    tempFile.moveTo(targetFolder);

    SpreadsheetApp.flush();

    const url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Błąd podczas konwersji arkusza.");
    }

    const dateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const excelBlob = response.getBlob().setName(`Candido_Timesheet_${dateFormatted}.xlsx`);
    
    const finalFile = targetFolder.createFile(excelBlob);
    tempFile.setTrashed(true);

    return { success: true, fileUrl: `https://drive.google.com/uc?export=download&id=${finalFile.getId()}` };
  } catch (e) {
    return { error: e.message };
  }
}
