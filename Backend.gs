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
    .setTitle('Seavia - Panel Brygadzisty')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

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
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (sheetSpis && sheetSpis.getLastRow() >= 2) {
    const spisData = sheetSpis.getRange(2, 2, sheetSpis.getLastRow() - 1, 7).getDisplayValues(); 
    spisData.forEach(row => {
      let timeStart = String(row[0]).trim();
      let timeEnd = String(row[1]).trim();
      let deck = String(row[2]).trim();
      let taskName = String(row[4]).trim();
      let isCancelled = String(row[4]).includes("[ANULOWANE]");
      let techNameFromSpis = String(row[6]).trim();

      if (timeStart && timeEnd && !isCancelled) {
        let duration = parseWorkTimeBackend(timeEnd) - parseWorkTimeBackend(timeStart);
        if (duration < 0) duration += 24; 
        if (duration > 0) {
          totalWorkedHours += duration;
          let key = deck + "_" + taskName;
          if(!workedHoursByTask[key]) workedHoursByTask[key] = 0;
          workedHoursByTask[key] += duration;

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

  let tasksTree = {}; 
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido) {
    const lastRow = sheetCandido.getLastRow();
    if (lastRow >= 3) {
      const data = sheetCandido.getRange(1, 1, lastRow, 16).getValues();
      let currentDeck = "";
      let currentArea = "";
      for (let i = 2; i < data.length; i++) {
        let deckVal = String(data[i][2] || "").trim();
        if (deckVal === "" && String(data[i][1] || "").toLowerCase().includes("dk")) deckVal = String(data[i][1] || "").trim();
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

        let key = currentDeck + "_" + taskName;
        let taskHours = workedHoursByTask[key] || 0;

        tasksTree[currentDeck][area].push({
          row: i + 1,
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
    tasksTree: tasksTree
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
    let rowDeck = data[i][3];
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
        let cancelMsg = `[ANULOWANE] przez: ${userEmail} (${timestamp})`;
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
    return `${y}-${m}-${d}`;
  }
  if (typeof val === 'string') {
    let str = val.trim();
    let m1 = str.match(/(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})/);
    if (m1) {
      let d = m1[1].padStart(2, '0');
      let m = m1[2].padStart(2, '0');
      let y = m1[3];
      if (y.length === 2) y = "20" + y;
      return `${y}-${m}-${d}`;
    }
    if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;
  }
  return null;
}

function submitWorkReport(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let idPrac = "";
    let isDuplicate = false;
    const sheetCandido = ss.getSheetByName('Candido');
    if (formData.isCustomTask) {
      if (sheetCandido) {
        // AUTOMATYZACJA: Dodawanie dopisku " - Extra Job" do ręcznie wpisanego ID prac
        let customId = formData.customTaskId ? String(formData.customTaskId).trim() : "";
        if (customId) {
          if (!customId.endsWith(" - Extra Job")) {
            idPrac = customId + " - Extra Job";
          } else {
            idPrac = customId;
          }
        } else {
          idPrac = "Extra Job";
        }

        // Sprawdzamy, czy zadanie o tym ID już istnieje na wybranym pokładzie
        let duplicateRow = -1;
        const lastRow = sheetCandido.getLastRow();
        if (lastRow >= 3) {
          const candidoData = sheetCandido.getRange(1, 1, lastRow, 16).getValues();
          let currentDeck = "";
          for (let r = 2; r < candidoData.length; r++) {
            let deckVal = String(candidoData[r][2] || "").trim();
            if (deckVal === "" && String(candidoData[r][1] || "").toLowerCase().includes("dk")) {
              deckVal = String(candidoData[r][1] || "").trim();
            }
            if (deckVal !== "") {
              currentDeck = deckVal;
            }
            
            let rowIdPrac = String(candidoData[r][5] || "").trim();
            if (currentDeck.toLowerCase() === formData.deck.toLowerCase() && rowIdPrac.toLowerCase() === idPrac.toLowerCase()) {
              duplicateRow = r + 1; // 1-based index w arkuszu
              break;
            }
          }
        }

        if (duplicateRow !== -1) {
          // Zadanie już istnieje! Aktualizujemy postęp na istniejącym wierszu i nie dodajemy duplikatu
          isDuplicate = true;
          if (formData.progress) {
            sheetCandido.getRange(duplicateRow, 16).setValue(formData.progress + "%");
          }
        } else {
          // Zadanie nie istnieje, dodajemy je do Candido
          let formattedDate = formData.workDate;
          let dParts = formData.workDate.split('-');
          if (dParts.length === 3) {
             formattedDate = `${dParts[1]}/${dParts[2]}/${dParts[0]}`; 
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
    let snowyTechs = formData.techniciansArray;
    if (!snowyTechs || !Array.isArray(snowyTechs)) {
      return { success: false, message: "Brak techników do przypisania." };
    }
    snowyTechs.forEach(tech => {
       sheetSpis.appendRow([
         formData.workDate, 
         formData.timeStart, 
         formData.timeEnd, 
         formData.deck, 
         formData.area, 
         formData.taskName, 
         idPrac, 
         tech
       ]);
    });
    let msgSuffix = snowyTechs.length > 1 ? ` (Dla ${snowyTechs.length} techników)` : '';
    let duplicateNote = isDuplicate ? " (Zadanie z tym ID już istniało na tym pokładzie, godziny dopisano do istniejącego zadania)" : "";
    return { success: true, message: `Pomyślnie zapisano godziny i postęp zadania!${duplicateNote}${msgSuffix}` };
  } catch (error) {
    return { success: false, message: "Błąd podczas zapisu: " + error.toString() };
  }
}

// -------------------------------------------------------------
// MODUŁ RAPORTÓW CANDIDO (ROZLICZENIA DLA KLIENTA)
// -------------------------------------------------------------
function getCandidoWeeklyReport(dateStr, includeBreaks) {
  if (!dateStr) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, weekInfo: {} };

  let selectedDate = new Date(dateStr);
  let dayOfWeek = selectedDate.getDay(); 
  let diff = selectedDate.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  let monday = new Date(selectedDate.setDate(diff));
  let weekDates = [];
  for (let i = 0; i < 7; i++) {
    let d = new Date(monday);
    d.setDate(monday.getDate() + i);
    weekDates.push(parseDateToYMD(d));
  }
  let target = new Date(monday.valueOf());
  let dayNr = (monday.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  let firstThursday = target.valueOf();
  target.setMonth(0, 1);
  if (target.getDay() !== 4) {
      target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
  }
  let weekNum = 1 + Math.ceil((firstThursday - target) / 604800000);

  let reportData = {
    settings: { includeBreaks: includeBreaks },
    weekInfo: {
      weekNumber: "W" + weekNum,
      start: weekDates[0],
      end: weekDates[6]
    },
    decks: {},
    allTasks: []
  };

  let uniqueTasks = new Set();
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido && sheetCandido.getLastRow() >= 3) {
    const candidoTasks = sheetCandido.getRange(3, 9, sheetCandido.getLastRow() - 2, 1).getValues();
    candidoTasks.forEach(row => {
      let t = String(row[0]).trim();
      if (t) {
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
    let rowDeck = String(spisData[i][3]).trim();
    let rowTask = String(spisData[i][5]).trim();
    let isCancelled = String(spisData[i][5]).includes("[ANULOWANE]");
    let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break");
    if (isBreak && !includeBreaks) continue;
    if (!rowStart || !rowEnd || isCancelled || !rowDeck || !rowTask) continue;
    if (rowTask.toLowerCase().includes("site manager") || rowTask.toLowerCase().includes("meetings") || rowTask.toLowerCase() === "supervision") {
       rowTask = "Supervision";
    }
    let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
    if (duration < 0) duration += 24; 
    if (duration > 0) {
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
  reportData.allTasks = Array.from(uniqueTasks).sort((a, b) => a.localeCompare(b));
  Object.keys(reportData.decks).forEach(d => {
     Object.keys(reportData.decks[d].tasks).forEach(t => {
       reportData.decks[d].tasks[t] = reportData.decks[d].tasks[t].map(h => Math.round(h * 10) / 10);
     });
  });
  return reportData;
}

function getTimeByWeekReport(includeBreaks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, allTasks: [], allWeeks: [] };
  const projectTasks = new Set();
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido && sheetCandido.getLastRow() >= 3) {
    const candidoTasks = sheetCandido.getRange(3, 9, sheetCandido.getLastRow() - 2, 1).getValues();
    candidoTasks.forEach(row => {
      const taskName = String(row[0]).trim();
      if (taskName) {
        projectTasks.add(taskName);
      }
    });
  }
  let reportData = {
    settings: { includeBreaks: includeBreaks },
    decks: {},
    allTasks: projectTasks,
    allWeeks: new Set()
  };
  const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();
  const getWeekNumber = (d) => {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
  };
  for (let i = 0; i < spisData.length; i++) {
    let rowDate = rawDates[i][0];
    if (!(rowDate instanceof Date)) continue;
    let weekNum = getWeekNumber(rowDate);
    let rowStart = spisData[i][1];
    let rowEnd = spisData[i][2];
    let rowDeck = String(spisData[i][3]).trim();
    let rowTask = String(spisData[i][5]).trim();
    let isCancelled = rowTask.includes("[ANULOWANE]");
    let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break");
    if (isBreak && !includeBreaks) continue;
    if (!rowStart || !rowEnd || isCancelled || !rowDeck || !rowTask) continue;
    if (rowTask.toLowerCase().includes("site manager") || rowTask.toLowerCase().includes("meetings") || rowTask.toLowerCase() === "supervision") {
       rowTask = "Supervision";
    }
    let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
    if (duration < 0) duration += 24;
    if (duration > 0) {
      if (!reportData.decks[rowDeck]) {
         reportData.decks[rowDeck] = { tasks: {} };
      }
      if (!reportData.decks[rowDeck].tasks[rowTask]) {
         reportData.decks[rowDeck].tasks[rowTask] = { total: 0, weeklyHours: {} };
      }
      if (!reportData.decks[rowDeck].tasks[rowTask].weeklyHours[weekNum]) {
        reportData.decks[rowDeck].tasks[rowTask].weeklyHours[weekNum] = 0;
      }
      reportData.decks[rowDeck].tasks[rowTask].weeklyHours[weekNum] += duration;
      reportData.decks[rowDeck].tasks[rowTask].total += duration;
      if(rowTask === "Supervision") {
        reportData.allTasks.add(rowTask);
      }
      reportData.allWeeks.add(weekNum);
    }
  }
  reportData.allTasks = Array.from(reportData.allTasks).sort((a, b) => a.localeCompare(b));
  reportData.allWeeks = Array.from(reportData.allWeeks);
  Object.keys(reportData.decks).forEach(d => {
     Object.keys(reportData.decks[d].tasks).forEach(t => {
       let task = reportData.decks[d].tasks[t];
       task.total = Math.round(task.total * 10) / 10;
       Object.keys(task.weeklyHours).forEach(w => {
         task.weeklyHours[w] = Math.round(task.weeklyHours[w] * 10) / 10;
       });
     });
  });
  return reportData;
}

function getTechnicianWeeklyReports(dateStr, includeBreaks) {
  if (!dateStr) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return null;
  let selectedDate = new Date(dateStr);
  let dayOfWeek = selectedDate.getDay(); 
  let diff = selectedDate.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
  let monday = new Date(selectedDate.setDate(diff));
  let weekDates = []; 
  for (let i = 0; i < 7; i++) {
    let d = new Date(monday);
    d.setDate(monday.getDate() + i);
    weekDates.push(parseDateToYMD(d));
  }
  let target = new Date(monday.valueOf());
  let dayNr = (monday.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  let firstThursday = target.valueOf();
  target.setMonth(0, 1);
  if (target.getDay() !== 4) target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
  let weekNum = 1 + Math.ceil((firstThursday - target) / 604800000);
  let year = target.getFullYear();
  let reportData = {
    weekInfo: {
      weekNumber: weekNum,
      year: year,
      dates: weekDates
    },
    technicians: {}
  };
  const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < spisData.length; i++) {
    let rowDate = parseDateToYMD(rawDates[i][0]);
    let dayIndex = weekDates.indexOf(rowDate);
    if (dayIndex === -1) continue;
    let rowStart = String(spisData[i][1]).trim();
    let rowEnd = String(spisData[i][2]).trim();
    let rowDeck = String(spisData[i][3]).trim();
    let rowArea = String(spisData[i][4]).trim();
    let rowTask = String(spisData[i][5]).trim();
    let rowRoomId = String(spisData[i][6]).trim();
    let techName = String(spisData[i][7]).trim();
    let isCancelled = rowTask.includes("[ANULOWANE]");

    let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break") || rowTask.toLowerCase().includes("brake");
    if (isBreak && !includeBreaks) continue;
    if (!rowStart || !rowEnd || isCancelled || !techName || !rowTask) continue;
    let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
    if (duration < 0) duration += 24; 

    if (duration > 0) {
      if (!reportData.technicians[techName]) {
         reportData.technicians[techName] = { 
             totalHours: 0, 
             days: [[], [], [], [], [], [], []] 
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
  return reportData;
}

// --- ZMIENIONA FUNKCJA GENEROWANIA EXCELA (AUTOMATYCZNY ZAPIS JAKO .XLSX NA DYSKU) ---
function generateStyledExcelReport(reportData, reportType) {
  try {
    const isWeekly = reportType === 'weekly';
    const baseName = `Candido_${isWeekly ? reportData.weekInfo.weekNumber : 'Annual'}_Report`;
    const tempFileName = `${baseName}_TEMP_${Date.now()}`;
    const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);

    // 1. Tworzymy tymczasowy Arkusz Google bezpośrednio w folderze docelowym
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
      sheet.getRange(currentRow, 1, 1, 2).setValues([["Date range:", `${reportData.weekInfo.start} / ${reportData.weekInfo.end}`]]).setFontWeight("bold");
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

      sheet.getRange(currentRow, 1).setValue(isWeekly ? `WEEKLY SUM ${deckName}` : `ANNUAL TIME REPORT - ${deckName}`).setFontWeight("bold").setFontColor(seaviaBlue);
      currentRow++;

      let headers = ["No.", "Tasks", "Total"];
      let columnCount;
      if (isWeekly) {
        headers.push('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
        columnCount = 10;
      } else {
        let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
        sortedWeeks.forEach(w => headers.push(`W${w}`));
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
          sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]).setNumberFormat("0.0");
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

      // ZADANIE SUPERVISION JEST RENDEROWANE ZAWSZE NA SAMEJ GÓRZE (POZYCJA 0)
      renderRow("Supervision", 0);

      let taskIndex = 1;
      let filteredTasks = reportData.allTasks;
      if (!reportData.settings.includeBreaks) {
        filteredTasks = filteredTasks.filter(t => !t.toLowerCase().includes('break') && !t.toLowerCase().includes('przerwa'));
      }
      filteredTasks.forEach(taskName => {
        // Renderujemy unikalne zadania (nawet o wartości 0)
        if (taskName !== "Supervision") {
          renderRow(taskName, taskIndex++);
        }
      });

      // Wyliczamy faktyczną liczbę wygenerowanych wierszy danych na tym pokładzie
      let renderedRowCount = taskIndex; 
      if (renderedRowCount > 0) {
        let tableStartRow = currentRow - renderedRowCount - 1; // Wracamy do wiersza nagłówka
        let numRows = renderedRowCount + 1; // headers + wiersze danych
        if (tableStartRow > 0 && numRows > 0 && columnCount > 0) {
          let tableRange = sheet.getRange(tableStartRow, 1, numRows, columnCount);
          tableRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
        }
      }

      sheet.setColumnWidth(2, 250);
      currentRow += 2;
    });

    // Wymuszamy natychmiastowe zrzucenie stylów i danych do arkusza Google Sheets przed konwersją
    SpreadsheetApp.flush();

    // 2. Eksport pliku do formatu Excel (.xlsx) za pomocą wewnętrznego API Google
    const url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error("Błąd podczas konwersji arkusza do formatu Excel (.xlsx): " + response.getContentText());
    }
    const excelBlob = response.getBlob().setName(`${baseName}.xlsx`);

    // 3. Zapisujemy ostateczny plik Excel (.xlsx) w Twoim docelowym folderze na Google Drive
    const excelFile = targetFolder.createFile(excelBlob);

    // 4. Usuwamy tymczasowy Arkusz Google Sheets, aby nie zaśmiecać folderu
    tempFile.setTrashed(true);

    // 5. Zwracamy bezpośredni link do pobrania pliku Excel (.xlsx) z Dysku Google
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

/**
 * NOWE: Aktualizuje nazwę drużyny (teamu) dla wszystkich przypisanych techników w Data sheet.
 */
function updateTeamName(oldTeamName, newTeamName) {
  if (!oldTeamName || oldTeamName.trim() === "" || !newTeamName || newTeamName.trim() === "") {
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
        if (String(values[i][0]).trim().toLowerCase() === oldTeamName.trim().toLowerCase()) {
          dataSheet.getRange(9 + i, 1).setValue(newTeamName.trim());
          updatedCount++;
        }
      }
    }
    return { success: true, message: `Pomyślnie zaktualizowano nazwę drużyny dla ${updatedCount} techników.`, updatedCount: updatedCount };
  } catch (error) {
    return { success: false, message: "Błąd podczas zmiany nazwy drużyny: " + error.toString() };
  }
}

/**
 * NOWE: Usuwa drużynę (team) - ustawia pustą wartość dla wszystkich techników z danej drużyny.
 */
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
          dataSheet.getRange(9 + i, 1).setValue(""); // Czyścimy drużynę
          updatedCount++;
        }
      }
    }
    return { success: true, message: `Pomyślnie usunięto drużynę ${teamName} (odpięto ${updatedCount} techników).`, updatedCount: updatedCount };
  } catch (error) {
    return { success: false, message: "Błąd podczas usuwania drużyny: " + error.toString() };
  }
}
