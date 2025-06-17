# 📊 Everhour Budget Sync for Google Sheets

Sync your Everhour project budgets into Google Sheets — with one click.

This Google Apps Script adds a custom menu to your spreadsheet, allowing you to pull budget info (total, used, remaining, % used) from one or multiple Everhour projects into clean, auto-formatted tabs.

---

## ✅ Features

* Add **one or many projects** to any sheet
* Auto-creates or updates named tabs
* Displays:

  * Project Name
  * Budget Total
  * Budget Used
  * Remaining
  * % Used (with conditional formatting)
* Works for projects using **money-based budgets**
* Custom UI menu: **Everhour Tools > Add Projects to Sheet**

---

## 🚀 Setup Instructions

### 1. Get your Everhour API key

→ [https://app.everhour.com/#/account/profile](https://app.everhour.com/#/account/profile)

### 2. Open your Google Sheet

Go to `Extensions > Apps Script`

### 3. Paste the script

Use the full script from `everhour-sync.gs`
Don’t forget to replace:

```javascript
const EVERHOUR_API_KEY = 'YOUR_API_KEY_HERE';
```

### 4. Save and reload the sheet

A new menu will appear: **Everhour Tools**

---

## 💎 How to Use

1. Click **Everhour Tools > Add Projects to Sheet**
2. Enter a sheet/tab name (will be created if it doesn’t exist)
3. Paste one or more Everhour project IDs, like:

```
li:abcd-1234...
li:efgh-5678...
```

You can find the project ID by clicking on a project in Everhour — it's in the URL (e.g. `https://app.everhour.com/#/projects/li:abcd-1234...`).

4. Click OK — the tab will fill automatically.

👉 Don’t forget to apply formatting:

* Set columns B, C, D to currency format (€)
* Set column E to percent format and reduce to 0 or 1 decimal

---

## 🧷 Embedding in Notion

To view your synced budgets inside Notion:

1. In Google Sheets, click **Share > Publish to the web**.
2. Select **Entire Document** or just your budget tab, then click **Publish**.
3. Copy the **public link** generated.
4. In Notion:

   * Type `/embed`
   * Choose **Embed** block
   * Paste your Google Sheets link
   * Resize the block as needed

> 🔒 Note: The sheet must be public (or at least visible to those with the link) for the embed to work.

---

## 💡 Notes

* Budget values are fetched in **euros (€)** and converted from cents.
* Projects must have a `type: "money"` budget set in Everhour.
* Conditional formatting:

  * Red if over 90% used
  * Green if under

---

## 🧑‍💻 Author

Made by me
Feel free to fork, remix, or suggest improvements.
