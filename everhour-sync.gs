// === Constants ===
/**
 * WARNING: The Everhour API Key MUST be stored in Script Properties for security.
 * Go to: File > Project properties > Script properties, and add a property
 * with key 'EVERHOUR_API_KEY' and your Everhour token as the value.
 */
// === Constants ===
const EVERHOUR_API_KEY = 'YOUR_API_KEY';
// 💡 IMPORTANT: If you kept the hardcoded key, replace it now with the line above!


// === Add Menu on Open ===
/**
 * Creates a custom menu in the Google Sheet named 'Everhour Tools'
 * with actions for syncing and maintenance.
 */

// === Add Menu on Open ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Everhour Tools')
    .addItem('List All Projects', 'listAllEverhourProjects')
    .addItem('Sync All Budgets', 'syncAllBudgets')
    .addItem('Sync Dashboard', 'generateDashboardCharts')
    .addItem('Sync One Sheet Group', 'promptAndSyncGroupDetails')
    .addItem('Clean Empty Rows', 'cleanEmptyProjectNameRows')
    .addToUi();
}

// === Normalize Helper ===
/**
 * Normalizes a string for reliable lookup by trimming whitespace,
 * replacing multiple spaces with one, and converting to lowercase.
 * @param {string} str - The input string (e.g., a project name).
 * @returns {string} The normalized string.
 */
function normalize(str) {
  return String(str || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

// === Helper: Remove All-Blank Rows ===
/**
 * Iterates backward through a sheet and deletes any row where ALL cells are blank.
 * This helps keep budget sheets clean between syncs.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to clean.
 */
function removeBlankRows(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  let deleted = 0;
  let checked = 0;
  for (let i = lastRow; i > 1; i--) {
    const rowValues = sheet.getRange(i, 1, 1, lastCol).getValues()[0];
    checked++;
    const reallyBlank = rowValues.every(cell => cell === '' || cell === null || (typeof cell === 'string' && cell.trim() === ''));
    if (reallyBlank) {
      Logger.log(`🗑️ Removing blank row ${i} (content: ${JSON.stringify(rowValues)})`);
      sheet.deleteRow(i);
      deleted++;
    }
  }
  Logger.log(`removeBlankRows(): Checked ${checked} rows, deleted ${deleted} blank rows.`);
}

// === List All Everhour Projects ===
/**
 * Fetches all projects from the Everhour API and populates the 'CONFIG' sheet.
 * Preserves existing 'Sheet Name' mappings based on Project ID.
 */
function listAllEverhourProjects() {
  Logger.log("▶️ listAllEverhourProjects()");
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG')
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet('CONFIG');

  const url = 'https://api.everhour.com/projects';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Api-Key': EVERHOUR_API_KEY }
  });

  const projects = JSON.parse(response.getContentText());
  Logger.log(`✅ ${projects.length} projects fetched from Everhour.`);

  const existingData = configSheet.getRange(2, 1, Math.max(0, configSheet.getLastRow() - 1), 3).getValues();
  const sheetNameMap = new Map(existingData.map(row => [row[1], row[0]]));

  configSheet.clearContents();
  configSheet.getRange(1, 1, 1, 3).setValues([["Sheet Name", "Project ID", "Project Name"]]);

  let row = 2;
  projects.forEach(project => {
    const existingSheetName = sheetNameMap.get(project.id) || '';
    configSheet.getRange(row, 1, 1, 3).setValues([[existingSheetName, project.id, project.name]]);
    Logger.log(`🔁 Listed project: ${project.name} (${project.id})`);
    row++;
  });

  SpreadsheetApp.getUi().alert(`✅ ${projects.length} projects listed in CONFIG.`);
}

// === Prompt and Sync One Sheet Group ===
function promptAndSyncGroupDetails() {
  Logger.log("▶️ promptAndSyncGroupDetails()");
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter the sheet name to sync only that group:');
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const sheetName = response.getResponseText().trim();
  if (!sheetName) return;

  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  const rawData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 3).getValues();
  const projectIds = rawData.filter(r => r[0] === sheetName).map(r => r[1]);

  Logger.log(`📥 Projects to sync for sheet ${sheetName}: ${projectIds.join(', ')}`);

  if (projectIds.length === 0) {
    Logger.log(`❌ No projects found for sheet: ${sheetName}`);
    SpreadsheetApp.getUi().alert(`❌ No projects found in CONFIG for sheet: ${sheetName}`);
    return;
  }

  const errors = syncCustomProjectsToSheet(sheetName, projectIds);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Sync for ${sheetName} done. ${errors} errors.`, 'Group Sync');
}

// === Sync All Budgets ===
/**
 * Main synchronization routine. Reads project groupings from CONFIG
 * and triggers sync for each unique sheet. Finally, updates the Dashboard.
 */
function syncAllBudgets() {
  Logger.log("▶️ syncAllBudgets()");
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  if (!configSheet) {
    Logger.log("❌ CONFIG sheet not found.");
    return;
  }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("❌ CONFIG sheet is empty.");
    return;
  }

  const rawData = configSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const projectMap = {};
  rawData.forEach(([sheetName, projectId]) => {
    if (!sheetName || !projectId) return;
    if (!projectMap[sheetName]) projectMap[sheetName] = [];
    projectMap[sheetName].push(projectId);
  });

  Object.entries(projectMap).forEach(([sheetName, projectIds]) => {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      removeBlankRows(sheet); // Clean before every sync
      Logger.log(`🧹 Removed blank rows from sheet: ${sheetName}`);
    } else {
      Logger.log(`⚠️ Sheet ${sheetName} does not exist. Skipping blank row removal.`);
    }
    Logger.log(`🔄 Syncing: ${sheetName} with ${projectIds.length} projects.`);
    syncCustomProjectsToSheet(sheetName, projectIds);
  });

  Logger.log("📊 Generating dashboard charts...");
  generateDashboardCharts();
}

// === Sync Projects to Sheet (preserving Consommé Devis and Restant Devis) ===
/**
 * Fetches budget details for a list of Everhour projects and updates a single budget sheet.
 * CRITICALLY, it preserves manually entered data in columns F and G.
 * @param {string} sheetName - The name of the target sheet to sync.
 * @param {string[]} projectIds - An array of Everhour Project IDs to fetch.
 * @returns {number} The count of errors encountered during API fetching.
 */
function syncCustomProjectsToSheet(sheetName, projectIds) {
  Logger.log(`▶️ syncCustomProjectsToSheet(${sheetName}, ${projectIds.length} projects)`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  // === 1. Capture manually entered Consommé Devis and Restant Devis ===
  const existingData = sheet.getDataRange().getValues();
  // Build a map: normalized Project Name => [Consommé Devis, Restant Devis]
  const preservedMap = new Map(
    existingData
      .slice(1)
      .filter(row => !!row[1] && String(row[1]).trim() !== '')
      .map(row => [normalize(row[1]), [row[5], row[6]]])
  );

  // === 2. Wipe all rows except the header ===
  const maxRows = sheet.getLastRow();
  if (maxRows > 2) {
    // Only delete if more than one data row (never delete the last row)
    sheet.deleteRows(2, maxRows - 2);
  } else if (maxRows === 2) {
    // Just clear the contents of row 2 if it exists (cannot delete it)
    sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
  }

  // === 3. Insert headers ===
  sheet.getRange(1, 1, 1, 8).setValues([[
    'Devis', 'Project Name', 'Budget Total', 'Consommé Réel', 'Restant Réel',
    'Consommé Devis', 'Restant Devis', '% Consommé réel'
  ]]);

  let rowIdx = 2;
  let errors = 0;

  for (const projectId of projectIds) {
    Utilities.sleep(300);
    try {
      Logger.log(`🔎 Fetching projectId: ${projectId}`);
      const url = `https://api.everhour.com/projects/${projectId}`;
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'X-Api-Key': EVERHOUR_API_KEY },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() !== 200) {
        Logger.log(`❌ HTTP ${response.getResponseCode()} for ${projectId}: ${response.getContentText()}`);
        errors++;
        continue;
      }

      const data = JSON.parse(response.getContentText());
      const name = (data.name || '').trim();
      if (!name) {
        Logger.log(`⚠️ Project with ID ${projectId} has no name. Skipping row insert.`);
        continue;
      }
      Logger.log(`➡️ Prepared row: '${name}' (budget: ${data.budget?.budget ?? 0}, used: ${data.budget?.progress ?? 0})`);

      const budget = (data.budget?.budget ?? 0) / 100;
      const used = (data.budget?.progress ?? 0) / 100;
      const remaining = budget - used;
      const percentUsed = budget > 0 ? used / budget : '';

      // Retrieve preserved values (if any)
      const [consommeDevis, restantDevis] = preservedMap.get(normalize(name)) || ['', ''];

      sheet.getRange(rowIdx, 1, 1, 8).setValues([[
        '', name, budget, used, remaining, consommeDevis, restantDevis, percentUsed
      ]]);
      rowIdx++;
    } catch (e) {
      Logger.log(`❌ Exception for project ${projectId}: ${e}`);
      errors++;
    }
  }
  Logger.log(`✅ Finished syncCustomProjectsToSheet for '${sheetName}'. Errors: ${errors}`);
  return errors;
}

// === Dashboard with % Used Chart ===
function generateDashboardCharts() {
  Logger.log('▶️ generateDashboardCharts()');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) dashboard = ss.insertSheet('Dashboard');
  else {
    dashboard.getCharts().forEach(chart => dashboard.removeChart(chart));
    dashboard.clearContents();
  }

  const headers = [['Sheet', 'Project', '% Used']];
  let data = [];

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === 'CONFIG' || name === 'Dashboard') return;

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const project = row[1];
      const percent = parseFloat(row[7]);
      if (project && !isNaN(percent)) {
        data.push([name, project, percent]);
      }
    }
  });

  data.sort((a, b) => b[2] - a[2]);
  const allData = headers.concat(data);
  dashboard.getRange(1, 1, allData.length, 3).setValues(allData);

  if (data.length > 0) {
    const dataRange = dashboard.getRange(2, 2, data.length, 2);

    const chart = dashboard.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dataRange)
      .setPosition(1, 6, 0, 0)
      .setOption('title', '% Consommé Réel par Projet')
      .setOption('hAxis', {
        title: '% Utilisé',
        minValue: 0,
        maxValue: Math.max(1, Math.max(...data.map(d => d[2])) * 1.2),
        format: 'percent'
      })
      .setOption('vAxis', { title: 'Projet', textStyle: { fontSize: 10 } })
      .setOption('legend', { position: 'none' })
      .setOption('annotations', {
        alwaysOutside: true,
        textStyle: { fontSize: 10, bold: true, color: '#000' }
      })
      .build();

    dashboard.insertChart(chart);
  }
  Logger.log('✅ Dashboard generated.');
}

// === Strict Cleaner for orphaned rows (delete-only) ===
function cleanEmptyProjectNameRows() {
  Logger.log('▶️ cleanEmptyProjectNameRows()');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data[0][1] !== 'Project Name') {
    SpreadsheetApp.getUi().alert('⚠️ This doesn’t look like a budget sheet. Column B must be "Project Name".');
    return;
  }

  let removed = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const projectName = row[1];
    const budget = row[2];
    const consumed = row[3];
    const remaining = row[4];
    if ((projectName === '' || projectName === null) &&
        (budget !== 0 || consumed !== 0 || remaining !== 0)) {
      sheet.deleteRow(i + 1);
      removed++;
      Logger.log(`❌ Deleted unnamed row at sheet row ${i + 1}`);
    }
  }
  SpreadsheetApp.getUi().alert(`✅ Clean-up complete: ${removed} unnamed rows deleted.`);
}
