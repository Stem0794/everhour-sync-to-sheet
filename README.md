# Everhour Budget Sync for Google Sheets

This repository contains Google Apps Script code for an automated **Everhour budget synchronization** tool. This tool fetches project and budget data from the Everhour API and organizes it into a Google Sheet, preserving manual inputs and generating a summary dashboard.

## 🚀 Getting Started

This guide assumes you are familiar with Google Apps Script and Google Sheets.

### 1\. Initial Setup

1.  **Create a New Google Sheet.**
2.  Open the **Script editor** (Extensions \> Apps Script).
3.  Copy and paste the entire script code into the `Code.gs` file, replacing any existing content.
4.  **Save** the project (File \> Save).

### 2\. Configure the Everhour API Key

You must replace the placeholder API key in the script with your actual Everhour API key.

1.  Obtain your **Everhour API Key** from your Everhour profile settings.

2.  In the Apps Script editor, modify the following constant at the top of the script:

    ```javascript
    // === Constants ===
    const EVERHOUR_API_KEY = 'YOUR_ACTUAL_EVERHOUR_API_KEY_HERE'; // ⚠️ REPLACE THIS!
    ```

### 3\. Run the Initial Sync

1.  **Open or Refresh** your Google Sheet. A new menu called **"Everhour Tools"** will appear after a moment.
2.  Go to **Everhour Tools** \> **List All Projects**.
3.  The first time you run this, you will be prompted to grant **authorization**. Review the permissions and allow access.
4.  This function creates a new sheet named **`CONFIG`** and populates it with all your Everhour project IDs and names.

-----

## 🛠️ Configuration and Mapping

The **`CONFIG`** sheet is the **master list** that dictates which Everhour projects go into which budget tracking sheets.

| Column | Name | Description | Action |
| :--- | :--- | :--- | :--- |
| **A** | **Sheet Name** | The name of the custom budget sheet where the project's data will be synced (e.g., `Client A`, `Internal Projects`). | **Manually fill this column** for every project you want to track. Projects with an empty cell here will be ignored by the sync functions. |
| **B** | Project ID | The unique ID from Everhour. | **Do not edit.** Used by the script to fetch data. |
| **C** | Project Name | The name of the project from Everhour. | **Do not edit.** For reference. |

-----

## ⚙️ Everhour Tools Menu Functions

The custom menu provides the primary way to interact with the synchronization script.

### 1\. List All Projects (`listAllEverhourProjects`)

  * **Purpose:** Fetches a fresh list of **all** projects from Everhour and updates the `CONFIG` sheet.
  * **When to Use:** Run this whenever you create new projects in Everhour and need to map them to budget sheets. It **preserves** any existing "Sheet Name" values based on the Project ID.

### 2\. Sync All Budgets (`syncAllBudgets`)

  * **Purpose:** The main synchronization function. It iterates through the `CONFIG` sheet, groups projects by their **Sheet Name**, performs the data sync for each group, and then generates the **Dashboard**.
  * **Process:**
    1.  **Cleans** any fully blank rows from the target budget sheets.
    2.  **Fetches** the `Budget Total` and current time **consumption** from Everhour.
    3.  **Preserves** the data in the manual input columns (`Consommé Devis` and `Restant Devis`).
    4.  **Updates/overwrites** the Everhour-derived columns.

### 3\. Sync One Sheet Group (`promptAndSyncGroupDetails`)

  * **Purpose:** Allows you to quickly sync data for a single sheet (e.g., during testing or after editing a project's budget in Everhour).
  * **Usage:** Prompts you to enter the **Sheet Name** (e.g., `Client A`) and only updates that specific sheet.

### 4\. Sync Dashboard (`generateDashboardCharts`)

  * **Purpose:** Creates or refreshes a sheet named **`Dashboard`** with a list of all projects and their **`% Consommé Réel`** (Real % Used). It also generates a **Bar Chart** for quick visual analysis, sorting projects by the highest percentage used.

### 5\. Clean Empty Rows (`cleanEmptyProjectNameRows`)

  * **Purpose:** A maintenance function for your individual budget sheets. It checks the **active sheet** (the one you currently have open) for rows where the `Project Name` is blank, but budget/consumption figures are non-zero (orphaned data).
  * **Usage:** Click on a budget sheet (e.g., `Client A`) and run this to clean up the data.

-----

## 📊 Budget Sheet Columns (Data Structure)

The budget sheets created by the script are designed to merge live Everhour data with your manual tracking.

| Column Header | Data Source | Description |
| :--- | :--- | :--- |
| **Devis** | Manual Input | Optional: For recording the total estimated budget in hours/cost. |
| **Project Name** | Everhour | The name of the project. |
| **Budget Total** | Everhour | The total budget set in Everhour (in hours/cost, depending on your Everhour setup). |
| **Consommé Réel** | Everhour | The total time/cost consumed for the project as reported by Everhour. |
| **Restant Réel** | Calculation | `Budget Total` minus `Consommé Réel`. |
| **Consommé Devis** | **Manual Input** | For manually tracking consumption against the initial `Devis`. **This column is preserved during sync.** |
| **Restant Devis** | **Manual Input** | For manually tracking remaining budget against the initial `Devis`. **This column is preserved during sync.** |
| **% Consommé réel** | Calculation | Percentage of `Consommé Réel` against `Budget Total`. |

-----

## 🔒 Error Handling and Preservation

The script includes several features to ensure data integrity:

  * **Data Preservation:** The **`syncCustomProjectsToSheet`** function explicitly captures and re-applies the values in the **`Consommé Devis`** and **`Restant Devis`** columns before overwriting the sheet, meaning you won't lose your manual tracking data.
  * **Blank Row Cleaning:** The **`removeBlankRows`** function is called before every full sync to prevent accidental orphaned rows from accumulating.
  * **Rate Limiting:** A `Utilities.sleep(300)` call is included between API calls in the sync function to prevent hitting Everhour's API rate limits.
