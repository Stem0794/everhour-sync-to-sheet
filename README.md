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

→ [https://everhour.dev/docs/#section/Authentication](https://everhour.dev/docs/#section/Authentication)

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

4. Click OK — the tab will fill automatically.

---

## 💡 Notes

* Budget values are fetched in **euros (€)** and converted from cents.
* Projects must have a `type: "money"` budget set in Everhour.
* Conditional formatting:

  * Red if over 90% used
  * Green if under

---

## 🧑‍💼 Author

Made by me
Feel free to fork, remix, or suggest improvements.
