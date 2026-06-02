/**
 * BACKEND - Logika aplikacji (Apps Script)
 */

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
        let team = String(r[0]).trim(); // Kolumna A (Drużyna)
        let name = String(r[1]).trim(); // Kolumna B (Nazwisko)
        let status = String(r[2]).trim().toLowerCase(); // Kolumna C (Status)
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
      let timeStart = String(row[0]).trim(); // B
      let timeEnd = String(row[1]).trim();   // C
      let deck = String(row[2]).trim();      // D
      let taskName = String(row[4]).trim();  // F
      let isCancelled = String(row[4]).includes("[ANULOWANE]");
      let techNameFromSpis = String(row[6]).trim(); // H

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
        let deckVal = String(data[i][2] || "").trim(); // C
        if (deckVal === "" && String(data[i][1] || "").toLowerCase().includes("dk")) deckVal = String(data[i][1] || "").trim();
        let areaVal = String(data[i][3] || "").trim(); // D - OBSZAR/AREA
        
        let taskName = String(data[i][8] || "").trim(); // I - NAZWA ZADANIA
        if (taskName === "") {
            let taskH = String(data[i][7] || "").trim();
            let taskJ = String(data[i][9] || "").trim();
            if (taskH !== "" && !taskH.toLowerCase().includes("seavia")) taskName = taskH;
            else if (taskJ !== "") taskName = taskJ;
        }
        
        let plannedStart = data[i][9]; // J
        let plannedEnd = data[i][12]; // M
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
    
    const sheetCandido = ss.getSheetByName('Candido');
    
    if (formData.isCustomTask) {
      if (sheetCandido) {
        idPrac = formData.customTaskId || "Dodatkowe";
        
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
         formData.deck, 
         formData.area, 
         formData.taskName, 
         idPrac, 
         tech
       ]);
    });

    let msgSuffix = techsArray.length > 1 ? ` (Dla ${techsArray.length} techników)` : '';
    return { success: true, message: `Pomyślnie zapisano godziny i postęp zadania!${msgSuffix}` };
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
      
      // Dodajemy zadania z godzinami, jeśli nie ma ich na głównej liście (np. Supervision)
      if (rowTask === "Supervision" || !Array.from(uniqueTasks).find(t => t.toLowerCase() === rowTask.toLowerCase())) {
         uniqueTasks.add(rowTask);
      }
    }
  }

  reportData.allTasks = Array.from(uniqueTasks).sort((a, b) => a.localeCompare(b));
  // Odrzucamy zadania, które nie są Supervision i nie pochodzą z Candido
  let candidoTasksLower = Array.from(uniqueTasks).map(t=>t.toLowerCase());
  reportData.allTasks = reportData.allTasks.filter(t => t === "Supervision" || candidoTasksLower.includes(t.toLowerCase()));

  Object.keys(reportData.decks).forEach(d => {
     Object.keys(reportData.decks[d].tasks).forEach(t => {
       reportData.decks[d].tasks[t] = reportData.decks[d].tasks[t].map(h => Math.round(h * 10) / 10);
     });
  });

  return reportData;
}


/**
 * POPRAWIONA FUNKCJA: Generuje raport czasu pracy w podziale na tygodnie dla całego projektu.
 */
function getTimeByWeekReport(includeBreaks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, allTasks: [], allWeeks: [] };

  // 1. Zbuduj główną listę zadań z zakładki Candido, kolumna I
  const projectTasks = new Set();
  const sheetCandido = ss.getSheetByName('Candido');
  if (sheetCandido && sheetCandido.getLastRow() >= 3) {
    // Pobieramy dane z kolumny I (indeks 9)
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
    allTasks: projectTasks, // Używamy seta z Candido jako bazę
    allWeeks: new Set()
  };

  const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
  const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();

  // Funkcja pomocnicza do obliczania numeru tygodnia ISO
  const getWeekNumber = (d) => {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
  };

  // 2. Przetwórz godziny z arkusza 'Spis wykonanych prac'
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
      
      // Dodaj zadania specjalne (np. Supervision) do głównej listy, jeśli ich tam nie ma
      if(rowTask === "Supervision") {
        reportData.allTasks.add(rowTask);
      }
      reportData.allWeeks.add(weekNum);
    }
  }

  // 3. Przekonwertuj Sety na posortowane tablice
  reportData.allTasks = Array.from(reportData.allTasks).sort((a, b) => a.localeCompare(b));
  reportData.allWeeks = Array.from(reportData.allWeeks);

  // Zaokrąglij godziny
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
