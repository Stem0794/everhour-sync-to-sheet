/**
 * Everhour Google Sheets Budget Sync Tool
 * Adds custom menu to allow users to import Everhour project budgets into Sheets.
 *
 * Requirements:
 * - Valid Everhour API key (replace `EVERHOUR_API_KEY`)
 * - Everhour projects must use `money`-based budgets
 * 
 * Author: [Your Name or Org]
 */

const EVERHOUR_API_KEY = 'YOUR_API_KEY_HERE'; // Replace with your real Everhour API Key

// Add menu on sheet open
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Everhour Tools')
    .addItem('Add Projects to Sheet', 'addProjectsToSheet')
    .addToUi();
}

// Menu interaction
function addProjectsToSheet() {
  const ui = SpreadsheetApp.getUi();

  const namePrompt = ui.prompt('📝 Sheet Name', 'Enter the sheet/tab name (will be created if needed):', ui.ButtonSet.OK_CANCEL);
  if (namePrompt.getSelectedButton() !== ui.Button.OK) return;

  const sheetName = namePrompt.getResponseText().trim();
  if (!sheetName) return;

  const idsPrompt = ui.prompt('🔗 Everhour Project IDs', 'Enter one or more Everhour project IDs (comma or newline separated):', ui.ButtonSet.OK_CANCEL);
  if (idsPrompt.getSelectedButton() !== ui.Button.OK) return;

  const rawInput = idsPrompt.getResponseText().trim();
  const projectIds = rawInput
    .split(/[,\n]/)
    .map(id => id.trim())
    .filter(id => id.startsWith('li:'));

  if (!projectIds.length) {
    ui.alert('❌ No valid project IDs entered. All IDs must start with "li:".');
    return;
  }

  syncMultipleProjectsToSheet(sheetName, projectIds);
}

// Sync budget info for multiple Everhour projects to one sheet
function syncMultipleProjectsToSheet(sheetName, projectIds) {
  const sheet = getOrCreateSheet(sheetName);
  sheet.clearContents();

  sheet.getRange(1, 1, 1, 5).setValues([
    ['Project Name', 'Budget Total', 'Budget Used', 'Remaining', '% Used']
  ]);

  let row = 2;

  projectIds.forEach(projectId => {
    const url = `https://api.everhour.com/projects/${encodeURIComponent(projectId)}`;

    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'X-Api-Key': EVERHOUR_API_KEY }
      });

      const data = JSON.parse(response.getContentText());
      const name = data.name || projectId;
      const budget = data.budget || {};

      let total = '', used = '', remaining = '', percent = '';

      if (budget.type === 'money') {
        const totalCents = budget.budget ?? null;
        const usedCents = budget.progress ?? null;

        if (typeof totalCents === 'number' && typeof usedCents === 'number') {
          total = totalCents / 100;
          used = usedCents / 100;
          remaining = total - used;
          percent = total > 0 ? used / total : '';
        }
      }

      sheet.getRange(row, 1, 1, 5).setValues([
        [name, total, used, remaining, percent]
      ]);
    } catch (e) {
      Logger.log(`❌ Error for project ${projectId}: ${e.message}`);
      sheet.getRange(row, 1, 1, 5).setValues([
        ['ERROR', '', '', '', '']
      ]);
    }

    row++;
  });

  applyConditionalFormattingToSheet(sheet);
}

// Get or create a sheet tab
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// Add color formatting for budget % used
function applyConditionalFormattingToSheet(sheet) {
  const range = sheet.getRange('E2:E');
  range.clearFormat();

  const rules = [];

  const overbudgetRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.9)
    .setBackground('#f4cccc')
    .setRanges([range])
    .build();

  const underbudgetRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.9)
    .setBackground('#d9ead3')
    .setRanges([range])
    .build();

  rules.push(overbudgetRule, underbudgetRule);
  sheet.setConditionalFormatRules(rules);

  sheet.getRange('B2:D').setNumberFormat('#,##0.00 €');
  sheet.getRange('E2:E').setNumberFormat('0.00%');
}
