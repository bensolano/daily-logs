// ============================================================
// Activity Logger — Code.gs
// Container-bound Google Apps Script for Google Sheets
// ============================================================

// ── Constants ────────────────────────────────────────────────

const RECAP_MARKER = "▬ RÉCAPITULATIF"; // sentinel in col 1 of recap header row
const ARAMIS_DAY_H = 7; // reference working day in hours for Aramis ratio
const ARAMIS_DAY_HOURS_KEY = "aramis_day_hours"; // configurable ref hours in DocumentProperties
const TASKS_SHEET = "Tasks"; // name of the tasks reference sheet
const TODO_STORAGE_PREFIX = "daily_todos:"; // key prefix in UserProperties

// ── Lifecycle ────────────────────────────────────────────────

/**
 * Adds the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 Logger")
    .addItem("Ouvrir le journal", "showSidebar")
    .addSeparator()
    .addItem("Configurer heures Aramis", "promptAramisReferenceHours")
    .addToUi();
}

/**
 * Opens the activity logger sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle("Journal d'activité").setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Helpers ──────────────────────────────────────────────────

/**
 * Parses a duration value → total minutes.
 * Accepts either:
 *   - a string "hh:mm"  (from the sidebar input)
 *   - a Date object     (Google Sheets auto-converts "hh:mm" cells to time serials,
 *                        which getValues() returns as Date objects)
 * @param {string|Date} val
 * @returns {number}
 */
function durationToMinutes(val) {
  if (!val) return 0;
  // Google Sheets time serial → Date object
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }
  // Plain "hh:mm" string
  const str = String(val).trim();
  if (!str.includes(":")) return 0;
  const parts = str.split(":");
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  return h * 60 + m;
}

/**
 * Converts total minutes to an Aramis ratio (0–1) relative to a 7-hour day.
 * Rounded to 3 decimal places.
 * @param {number} minutes
 * @returns {number}
 */
function minutesToAramis(minutes) {
  return Math.round((minutes / 60 / getAramisDayHours()) * 1000) / 1000;
}

/**
 * Returns the configured Aramis reference day hours for this spreadsheet.
 * Falls back to default ARAMIS_DAY_H when unset or invalid.
 * @returns {number}
 */
function getAramisDayHours() {
  const raw = PropertiesService.getDocumentProperties().getProperty(ARAMIS_DAY_HOURS_KEY);
  if (!raw) return ARAMIS_DAY_H;
  const parsed = parseFloat(String(raw).replace(",", "."));
  return Number.isFinite(parsed) && parsed > 0 ? parsed : ARAMIS_DAY_H;
}

/**
 * Prompts user to change the Aramis reference day hours.
 * Stored in DocumentProperties and used for all future Aramis calculations.
 */
function promptAramisReferenceHours() {
  const ui = SpreadsheetApp.getUi();
  const current = getAramisDayHours();
  const response = ui.prompt(
    "Configurer heures Aramis",
    `Heures de référence actuelles : ${current}h\nEntrez une valeur entre 0 et 24 (ex: 7 ou 7.5).`,
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim().replace(",", ".");
  const value = parseFloat(input);
  if (!Number.isFinite(value) || value <= 0 || value > 24) {
    ui.alert("Valeur invalide. Entrez un nombre > 0 et ≤ 24 (ex: 7 ou 7.5).");
    return;
  }

  const normalized = Math.round(value * 100) / 100;
  PropertiesService.getDocumentProperties().setProperty(ARAMIS_DAY_HOURS_KEY, String(normalized));
  ui.alert(
    `Heures Aramis mises à jour: ${normalized}h. Pensez à lancer \"Recalculer totaux\" pour mettre à jour l'historique.`,
  );
}

/**
 * Formats total minutes as "hh:mm".
 * @param {number} minutes
 * @returns {string}
 */
function minutesToHHMM(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

// ── Tasks sheet ──────────────────────────────────────────────

/**
 * Finds or creates the "Tasks" reference sheet.
 * Structure: col A = task name, col B = active (TRUE/FALSE checkbox)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateTasksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASKS_SHEET);
    setupTasksSheet(sheet);
  }
  return sheet;
}

/**
 * Initialises a fresh Tasks sheet with styled headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupTasksSheet(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.setValues([["Tâche", "Active"]]);
  headerRange.setFontWeight("bold").setBackground("#1a3a5c").setFontColor("#ffffff").setHorizontalAlignment("center");
  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 80);
  sheet.setFrozenRows(1);
}

/**
 * Returns the list of active task names (for the sidebar dropdown).
 * Called from the sidebar via google.script.run.
 * @returns {string[]}
 */
function getTasks() {
  const sheet = getOrCreateTasksSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return data
    .filter((row) => row[1] === true || String(row[1]).toLowerCase() === "true")
    .map((row) => String(row[0]).trim())
    .filter((name) => name.length > 0);
}

/**
 * Adds a new task to the Tasks sheet (active by default).
 * If the task already exists (case-insensitive), silently returns its name.
 * Called from the sidebar via google.script.run.
 * @param {string} taskName
 * @returns {{ success: boolean, name: string }}
 */
function addTask(taskName) {
  const name = taskName.trim();
  if (!name) return { success: false, name: "" };

  const sheet = getOrCreateTasksSheet();
  const lastRow = sheet.getLastRow();

  // Check for duplicate (case-insensitive)
  if (lastRow >= 2) {
    const existing = sheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues()
      .flat()
      .map((v) => String(v).trim().toLowerCase());
    if (existing.includes(name.toLowerCase())) {
      return { success: true, name }; // already exists — treat as success
    }
  }

  // Append new active task with a real checkbox in col B
  sheet.appendRow([name, true]);
  sheet.getRange(sheet.getLastRow(), 2).insertCheckboxes();

  return { success: true, name };
}

// ── Data exposed to the sidebar ──────────────────────────────

/**
 * Returns today's date formatted in French for the sidebar greeting.
 * e.g. "Mercredi 11 mars 2026"
 * @returns {string}
 */
function getTodayFrench() {
  const now = new Date();
  const DAYS = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
  const MONTHS = [
    "janvier",
    "février",
    "mars",
    "avril",
    "mai",
    "juin",
    "juillet",
    "août",
    "septembre",
    "octobre",
    "novembre",
    "décembre",
  ];
  return `${DAYS[now.getDay()]} ${now.getDate()} ${MONTHS[now.getMonth()]} ${now.getFullYear()}`;
}

/**
 * Returns the local date key (yyyy-MM-dd) in script timezone.
 * @returns {string}
 */
function getTodayStorageKey() {
  const tz = Session.getScriptTimeZone() || "Europe/Paris";
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
}

/**
 * Returns today's TODO list for the sidebar.
 * Stored per user, per day.
 * @returns {{ items: { text: string, done: boolean, plannedMinutes: number }[] }}
 */
function getDailyTodos() {
  const key = TODO_STORAGE_PREFIX + getTodayStorageKey();
  const raw = PropertiesService.getUserProperties().getProperty(key);
  if (!raw) return { items: [] };

  try {
    const parsed = JSON.parse(raw);
    const items = Array.isArray(parsed.items)
      ? parsed.items
          .map((item) => ({
            text: String(item && item.text ? item.text : "").trim(),
            done: !!(item && item.done),
            plannedMinutes: Math.max(0, parseInt(item && item.plannedMinutes, 10) || 0),
          }))
          .filter((item) => item.text.length > 0)
      : [];
    return { items };
  } catch (e) {
    return { items: [] };
  }
}

/**
 * Saves today's TODO list from the sidebar.
 * @param {{ items: { text: string, done: boolean, plannedMinutes: number }[] }} payload
 * @returns {{ success: boolean }}
 */
function saveDailyTodos(payload) {
  const incoming = payload && Array.isArray(payload.items) ? payload.items : [];
  const items = incoming
    .map((item) => ({
      text: String(item && item.text ? item.text : "").trim(),
      done: !!(item && item.done),
      plannedMinutes: Math.max(0, parseInt(item && item.plannedMinutes, 10) || 0),
    }))
    .filter((item) => item.text.length > 0)
    .slice(0, 200);

  const key = TODO_STORAGE_PREFIX + getTodayStorageKey();
  PropertiesService.getUserProperties().setProperty(key, JSON.stringify({ items }));
  return { success: true };
}

// ── Sheet management ─────────────────────────────────────────

/**
 * Returns ISO week number (Monday = first day, ISO 8601).
 * @param {Date} date
 * @returns {number}
 */
function getISOWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7; // Sunday = 7
  d.setUTCDate(d.getUTCDate() + 4 - dayNum); // nearest Thursday
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

/**
 * Builds the tab name and title string for the current week.
 * Tab  : "34 3/02 7/02"
 * Title: "Semaine 34  Lun. 3/02/2026  Ven. 7/02/2026"
 * @returns {{ tabName: string, titleText: string }}
 */
function buildWeekStrings() {
  const now = new Date();
  const weekNum = getISOWeekNumber(now);

  // Monday of current ISO week
  const dayOfWeek = now.getDay() === 0 ? 6 : now.getDay() - 1; // 0 = Mon
  const monday = new Date(now);
  monday.setDate(now.getDate() - dayOfWeek);
  monday.setHours(0, 0, 0, 0);

  // Friday
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);

  const pad = (n) => String(n).padStart(2, "0");

  const monD = monday.getDate();
  const monM = pad(monday.getMonth() + 1);
  const friD = friday.getDate();
  const friM = pad(friday.getMonth() + 1);
  const year = friday.getFullYear();

  // Short tab name  — matches user example "34 3/02 7/02"
  const tabName = `${weekNum} ${monD}/${monM} ${friD}/${friM}`;

  // Long title cell — matches user example "Semaine 34 Lun. 3/02/2026 Ven. 7/02/2026"
  const titleText = `Semaine ${weekNum}   Lun. ${monD}/${monM}/${year}   Ven. ${friD}/${friM}/${year}`;

  return { tabName, titleText };
}

/**
 * Finds or creates the sheet for the current week, then returns it.
 * A new sheet gets a styled header row.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateWeekSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { tabName, titleText } = buildWeekStrings();

  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    setupWeekSheet(sheet, titleText);
  }
  return sheet;
}

/**
 * Initialises a fresh week sheet with a title cell and column headers.
 * Columns: Horodatage | Tâche | Durée | Durée totale tâche | Description | Aramis | Aramis total tâche | Durée totale jour
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} titleText
 */
function setupWeekSheet(sheet, titleText) {
  // ── Row 1 : week title ──────────────────────────────────────
  const titleRange = sheet.getRange(1, 1, 1, 8);
  titleRange.merge();
  titleRange.setValue(titleText);
  titleRange
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground("#1a3a5c")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 30);

  // ── Row 2 : column headers ──────────────────────────────────
  const headers = [
    "Horodatage",
    "Tâche",
    "Durée",
    "Durée totale tâche",
    "Description",
    "Aramis",
    "Aramis total tâche",
    "Durée totale jour",
  ];
  const headerRange = sheet.getRange(2, 1, 1, 8);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff").setHorizontalAlignment("center");
  sheet.setRowHeight(2, 24);

  // ── Column widths ───────────────────────────────────────────
  sheet.setColumnWidth(1, 185); // Horodatage
  sheet.setColumnWidth(2, 200); // Tâche
  sheet.setColumnWidth(3, 65); // Durée
  sheet.setColumnWidth(4, 150); // Durée totale tâche
  sheet.setColumnWidth(5, 300); // Description
  sheet.setColumnWidth(6, 80); // Aramis
  sheet.setColumnWidth(7, 150); // Aramis total tâche
  sheet.setColumnWidth(8, 150); // Durée totale jour

  sheet.setFrozenRows(2);
}

// ── Log entry ────────────────────────────────────────────────

/**
 * Appends a log entry to this week's sheet.
 *
 * Called from the sidebar via google.script.run.
 *
 * @param {{ title: string, description: string, duration: string }} payload
 * @returns {{ success: boolean, message: string }}
 */
function logEntry(payload) {
  try {
    const sheet = getOrCreateWeekSheet();
    const now = new Date();

    // French timestamp  "Mer. 11 mars 2026  14:32"
    const DAYS_SHORT = ["Dim.", "Lun.", "Mar.", "Mer.", "Jeu.", "Ven.", "Sam."];
    const MONTHS_LONG = [
      "janvier",
      "février",
      "mars",
      "avril",
      "mai",
      "juin",
      "juillet",
      "août",
      "septembre",
      "octobre",
      "novembre",
      "décembre",
    ];
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const timestamp = `${DAYS_SHORT[now.getDay()]} ${now.getDate()} ${MONTHS_LONG[now.getMonth()]} ${now.getFullYear()}  ${hh}:${mm}`;

    // Aramis ratio for this entry
    const minutes = durationToMinutes(payload.duration);
    const aramis = minutesToAramis(minutes);

    // ── Daily running totals ──────────────────────────────────
    // todayKey e.g. "12 mars 2026" — used to match existing rows from today
    const todayKey = now.getDate() + " " + MONTHS_LONG[now.getMonth()] + " " + now.getFullYear();
    const taskTitle = (payload.title || "").trim();
    let dailyTaskMinutes = minutes; // running total for THIS task today
    let totalDayMinutes = minutes; // running total for ALL tasks today
    let isFirstEntryOfDay = true; // auto-detected: no prior rows today
    const currentLastRow = sheet.getLastRow();
    if (currentLastRow >= 3) {
      // Read cols 1–3: Horodatage(0), Tâche(1), Durée(2)
      const existing = sheet.getRange(3, 1, currentLastRow - 2, 3).getValues();
      existing.forEach(function (r) {
        const rowTimestamp = String(r[0]).trim();
        if (!rowTimestamp) return; // blank separator row
        if (rowTimestamp.startsWith(RECAP_MARKER)) return; // recap header
        const rowDayKey = extractDayKey(rowTimestamp); // "12 mars 2026" or null
        if (!rowDayKey || rowDayKey !== todayKey) return; // different day
        // ── This row belongs to today ──
        const rowMins = durationToMinutes(r[2]); // col 3 — Durée
        totalDayMinutes += rowMins;
        isFirstEntryOfDay = false;
        if (String(r[1]).trim() === taskTitle) {
          // col 2 — Tâche
          dailyTaskMinutes += rowMins;
        }
      });
    }
    const dailyTaskHHMM = minutesToHHMM(dailyTaskMinutes);
    const dailyTaskAramis = minutesToAramis(dailyTaskMinutes);
    const totalDayHHMM = minutesToHHMM(totalDayMinutes);

    // ── Append the entry (with optional blank separator row before it) ────────
    // Col order: Horodatage | Tâche | Durée | Durée totale tâche | Description | Aramis | Aramis total tâche | Durée totale jour
    const entryData = [
      timestamp,
      payload.title || "",
      payload.duration || "",
      dailyTaskHHMM,
      payload.description || "",
      aramis,
      dailyTaskAramis,
      totalDayHHMM,
    ];

    const lastDataRow = sheet.getLastRow();
    if (isFirstEntryOfDay && lastDataRow > 2) {
      // insertRowAfter leaves the new row empty → getLastRow() stays at lastDataRow.
      // We must write the entry to lastDataRow+2 directly — appendRow would land on lastDataRow+1 (the blank row).
      sheet.insertRowAfter(lastDataRow);
      // insertRowAfter copies formatting from the row above — clear it so the separator stays visually blank.
      sheet
        .getRange(lastDataRow + 1, 1, 1, 8)
        .clearFormat()
        .setBackground(null);
      sheet.getRange(lastDataRow + 2, 1, 1, 8).setValues([entryData]);
    } else {
      sheet.appendRow(entryData);
    }

    // ── Row styling ──────────────────────────────────────────
    const row = sheet.getLastRow();

    // Col 1 — Horodatage: bold, dark blue font, light blue bg
    sheet.getRange(row, 1).setFontWeight("bold").setFontColor("#1a3a5c").setBackground("#eaf0fb");

    // Col 2 — Tâche: bold
    sheet.getRange(row, 2).setFontWeight("bold").setFontColor("#1a1a1a");

    // Col 3 — Durée: bold, dark green, centered
    sheet
      .getRange(row, 3)
      .setFontWeight("bold")
      .setFontColor("#2b6b2b")
      .setBackground("#e6f4ea")
      .setHorizontalAlignment("center");

    // Col 4 — Durée totale tâche: bold, muted green, centered
    sheet
      .getRange(row, 4)
      .setFontWeight("bold")
      .setFontColor("#1a5c2b")
      .setBackground("#c8e6c9")
      .setHorizontalAlignment("center");

    // Col 5 — Description: normal, lighter grey, wrap
    sheet.getRange(row, 5).setFontColor("#444444").setWrap(true);

    // Col 6 — Aramis: orange accent, number format, centered
    sheet
      .getRange(row, 6)
      .setNumberFormat("0.000")
      .setFontWeight("bold")
      .setFontColor("#7a4800")
      .setBackground("#fff8e6")
      .setHorizontalAlignment("center");

    // Col 7 — Aramis total tâche: deep orange, centered
    sheet
      .getRange(row, 7)
      .setNumberFormat("0.000")
      .setFontWeight("bold")
      .setFontColor("#7a4800")
      .setBackground("#ffe082")
      .setHorizontalAlignment("center");

    // Col 8 — Durée totale jour: bold, dark blue font, light blue bg, centered
    sheet
      .getRange(row, 8)
      .setFontWeight("bold")
      .setFontColor("#1a3a5c")
      .setBackground("#d0e4ff")
      .setHorizontalAlignment("center");

    return { success: true, message: "Entrée enregistrée ✓" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ── Weekly recap ─────────────────────────────────────────────

/**
 * Writes (or overwrites) a recap section at the bottom of the current week's sheet.
 * Groups all log entries by task title, sums duration and Aramis.
 * Safe to call repeatedly — always removes the previous recap first.
 * Triggered automatically every Saturday, or called manually from the menu.
 */
function writeWeeklyRecap() {
  const sheet = getOrCreateWeekSheet();
  const lastRow = sheet.getLastRow();

  // ── 1. Remove any existing recap ─────────────────────────
  if (lastRow >= 3) {
    const col1Values = sheet
      .getRange(3, 1, lastRow - 2, 1)
      .getValues()
      .flat();
    const recapStart = col1Values.findIndex((v) => String(v).startsWith(RECAP_MARKER));
    if (recapStart !== -1) {
      const firstRecapRow = recapStart + 3; // +2 header rows, +1 for 1-index
      sheet.deleteRows(firstRecapRow, lastRow - firstRecapRow + 1);
    }
  }

  // ── 2. Read all data rows ─────────────────────────────────
  const dataLastRow = sheet.getLastRow();
  if (dataLastRow < 3) {
    SpreadsheetApp.getUi().alert("Aucune donnée à récapituler pour cette semaine.");
    return;
  }

  // New layout: col1=Horodatage(0), col2=Tâche(1), col3=Durée(2)
  const data = sheet.getRange(3, 1, dataLastRow - 2, 3).getValues();

  // Accumulate minutes per task title
  const taskMap = {}; // { taskTitle: { minutes: number } }
  data.forEach((row) => {
    const title = String(row[1]).trim(); // col 2 — Tâche
    const duration = row[2]; // col 3 — Durée (may be a Date object)
    if (String(row[0]).startsWith(RECAP_MARKER)) return; // residual recap rows
    if (!title || duration === "" || duration === null || duration === undefined) return; // blank / separator rows
    const mins = durationToMinutes(duration);
    if (!taskMap[title]) taskMap[title] = { minutes: 0 };
    taskMap[title].minutes += mins;
  });

  const tasks = Object.keys(taskMap);
  if (tasks.length === 0) {
    SpreadsheetApp.getUi().alert("Aucune entrée valide trouvée pour cette semaine.");
    return;
  }

  // ── 3. Write the recap ────────────────────────────────────
  const writeRow = sheet.getLastRow() + 2; // leave one blank row as visual gap

  // Recap title row (merged across all 8 cols)
  const titleCell = sheet.getRange(writeRow, 1, 1, 8);
  titleCell.merge();
  titleCell.setValue(RECAP_MARKER + "  —  Récapitulatif de la semaine");
  titleCell
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground("#1a3a5c")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(writeRow, 28);

  // Recap column headers
  const recapHeaderRow = writeRow + 1;
  sheet
    .getRange(recapHeaderRow, 1)
    .setValue("Tâche")
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");
  sheet
    .getRange(recapHeaderRow, 2)
    .setValue("Durée totale")
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");
  sheet
    .getRange(recapHeaderRow, 4)
    .setValue("Aramis total")
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  // One row per task
  let totalMinutes = 0;
  let totalAramis = 0;

  tasks.forEach((title, idx) => {
    const mins = taskMap[title].minutes;
    const aramis = minutesToAramis(mins);
    const hhmm = minutesToHHMM(mins);
    totalMinutes += mins;
    totalAramis += aramis;

    const r = recapHeaderRow + 1 + idx;
    sheet.getRange(r, 1).setValue(title).setFontWeight("bold").setFontColor("#1a1a1a");
    sheet
      .getRange(r, 2)
      .setValue(hhmm)
      .setFontWeight("bold")
      .setFontColor("#2b6b2b")
      .setBackground("#e6f4ea")
      .setHorizontalAlignment("center");
    sheet
      .getRange(r, 4)
      .setValue(aramis)
      .setNumberFormat("0.000")
      .setFontWeight("bold")
      .setFontColor("#7a4800")
      .setBackground("#fff8e6")
      .setHorizontalAlignment("center");
  });

  // Total row
  const totalRow = recapHeaderRow + 1 + tasks.length;
  sheet.getRange(totalRow, 1).setValue("TOTAL").setFontWeight("bold").setFontColor("#1a3a5c").setBackground("#e8edf5");
  sheet
    .getRange(totalRow, 2)
    .setValue(minutesToHHMM(totalMinutes))
    .setFontWeight("bold")
    .setFontColor("#2b6b2b")
    .setBackground("#c8e6c9")
    .setHorizontalAlignment("center");
  sheet
    .getRange(totalRow, 4)
    .setValue(Math.round(totalAramis * 1000) / 1000)
    .setNumberFormat("0.000")
    .setFontWeight("bold")
    .setFontColor("#7a4800")
    .setBackground("#ffe082")
    .setHorizontalAlignment("center");

  Logger.log("Récapitulatif écrit — %s tâche(s), %s minutes.", tasks.length, totalMinutes);
}

// ── Recalculate totals ───────────────────────────────────────

/**
 * Extracts a day-key string from a French timestamp like "Mer. 11 mars 2026  14:32".
 * Returns "11 mars 2026" or null if it doesn't match.
 * @param {string} timestamp
 * @returns {string|null}
 */
function extractDayKey(timestamp) {
  const match = String(timestamp).match(/(\d{1,2}\s+\w+\s+\d{4})/);
  return match ? match[1] : null;
}

/**
 * Re-walks every data row of the current week's sheet in order, grouped by day,
 * and overwrites the four computed columns:
 *   col 4 — Durée totale tâche   (running total for this task today)
 *   col 6 — Aramis               (per-entry ratio, derived from col 3 Durée)
 *   col 7 — Aramis total tâche   (running Aramis total for this task today)
 *   col 8 — Durée totale jour    (running total for ALL tasks today)
 *
 * Safe to run after any manual edit to task names or durations in the sheet.
 * Called from the custom menu.
 */
function recalculateTotals() {
  const sheet = getOrCreateWeekSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert("Aucune donnée à recalculer sur cette feuille.");
    return;
  }

  // ── Load known task names from the Tasks sheet ────────────
  // Build a lowercase lookup set so we can detect & register missing tasks.
  const tasksSheet = getOrCreateTasksSheet();
  const tasksLastRow = tasksSheet.getLastRow();
  const knownTasks = new Set(); // lowercase names already in the Tasks sheet
  if (tasksLastRow >= 2) {
    tasksSheet
      .getRange(2, 1, tasksLastRow - 1, 1)
      .getValues()
      .flat()
      .forEach((v) => knownTasks.add(String(v).trim().toLowerCase()));
  }
  const addedTasks = []; // names we register during this run

  // Read all data rows — cols 1–8 (Horodatage … Durée totale jour)
  const numRows = lastRow - 2;
  const data = sheet.getRange(3, 1, numRows, 8).getValues();

  // Per-day accumulators (reset whenever the date portion of the timestamp changes)
  let currentDayKey = null;
  let taskMinutesMap = {}; // { taskName → minutes accumulated today }
  let totalDayMinutes = 0;
  let updated = 0;

  data.forEach(function (row, i) {
    const sheetRow = i + 3; // convert 0-based index → 1-based sheet row

    const timestamp = String(row[0]).trim();
    const taskTitle = String(row[1]).trim();
    const duration = row[2]; // "hh:mm" string or Date serial from Sheets

    // Skip blank separator rows
    if (!timestamp) return;
    // Skip recap header rows (sentinel stored in col 1)
    if (timestamp.startsWith(RECAP_MARKER)) return;
    // Skip rows without a parsable day or a task name or a duration
    const dayKey = extractDayKey(timestamp);
    if (!dayKey || !taskTitle || duration === "" || duration === null || duration === undefined) return;

    // ── Register task if it doesn’t exist in the Tasks sheet ────
    if (!knownTasks.has(taskTitle.toLowerCase())) {
      addTask(taskTitle);
      knownTasks.add(taskTitle.toLowerCase());
      addedTasks.push(taskTitle);
    }

    // ── Day boundary — reset per-day accumulators ─────────────
    if (dayKey !== currentDayKey) {
      currentDayKey = dayKey;
      taskMinutesMap = {};
      totalDayMinutes = 0;
    }

    const mins = durationToMinutes(duration);

    if (!taskMinutesMap[taskTitle]) taskMinutesMap[taskTitle] = 0;
    taskMinutesMap[taskTitle] += mins;
    totalDayMinutes += mins;

    // Recomputed values
    const aramis = minutesToAramis(mins);
    const dailyTaskHHMM = minutesToHHMM(taskMinutesMap[taskTitle]);
    const dailyTaskAramis = minutesToAramis(taskMinutesMap[taskTitle]);
    const totalDayHHMM = minutesToHHMM(totalDayMinutes);

    // Write the four columns back (styles already applied when row was first logged)
    sheet.getRange(sheetRow, 4).setValue(dailyTaskHHMM);
    sheet.getRange(sheetRow, 6).setValue(aramis).setNumberFormat("0.000");
    sheet.getRange(sheetRow, 7).setValue(dailyTaskAramis).setNumberFormat("0.000");
    sheet.getRange(sheetRow, 8).setValue(totalDayHHMM);

    updated++;
  });

  const taskMsg =
    addedTasks.length > 0 ? `\n\n${addedTasks.length} tâche(s) ajoutée(s) à la liste : ${addedTasks.join(", ")}.` : "";
  SpreadsheetApp.getUi().alert(
    updated > 0
      ? `✓ Recalcul terminé — ${updated} entrée(s) mise(s) à jour.${taskMsg}`
      : `Aucune entrée valide trouvée — rien n'a été modifié.${taskMsg}`,
  );
}
