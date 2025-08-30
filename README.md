# 🚨 Incident Number Tracker

A lightweight desktop tool built with **PySide6** and **openpyxl** to help teams log and track **incident tickets**.
The app provides a clean UI to add, preview, and manage incident records stored in an Excel file (`incident_numbers.xlsx`).

---

## ✨ Features

* 📅 **Date Picker** – Select or auto-fill the "Created On" date.
* 🔢 **Smart Ticket ID**

  * Auto-suggests next sequential ID in format: **`THyymmddNN`**
  * Example: `TH25083001`, `TH25083002` for incidents on Aug 30, 2025
  * User can override with their own (e.g., `INC2453903`).
* 📝 **Multi-line Description** – Write detailed notes; lines are auto-joined into a single line.
* 👀 **Preview Panel** – Shows real-time single-line version of your description.
* 📊 **Excel Integration** – All data is stored in `incident_numbers.xlsx` → **INCIDENTS** sheet.
* 🔄 **Quick Actions**

  * Add new incident
  * Clear description
  * Refresh table
  * Open Excel directly from the app
* ⚡ **Keyboard Shortcuts**

  * `Ctrl+Enter` → Add incident
  * `Ctrl+L` → Clear description
  * `Ctrl+T` → Set today’s date
  * `Ctrl+O` → Open Excel file
  * `F5` → Refresh list
* 📋 **Context Menu** – Right-click a row to copy it (tab-separated) to clipboard.
* 🎨 **Modern UI** – Styled with a **red banner** to clearly differentiate it from other trackers.

---

## 📂 File Structure

* `incident_numbers.xlsx` → Auto-created if missing, contains a sheet: `INCIDENTS`.
* `app.py` (your script) → Runs the GUI.

---

## 🚀 Getting Started

### 1. Clone or Download

```bash
git clone https://github.com/your-username/incident-number-tracker.git
cd incident-number-tracker
```

### 2. Install Requirements

```bash
pip install PySide6 openpyxl
```

### 3. Run the App

```bash
python app.py
```

---

## 📑 Excel Format

The `INCIDENTS` sheet has **3 columns**:

| Created On | Ticket ID  | Description                        |
| ---------- | ---------- | ---------------------------------- |
| 2025-08-30 | TH25083001 | Router crash, rebooted, monitoring |
| 2025-08-30 | INC1234567 | External ticket logged with vendor |

* Dates are stored in `yyyy-mm-dd` format.
* Ticket ID is either **auto-generated** (`THyymmddNN`) or manually entered.
* Description is a **single line**, with multi-line input automatically joined.

---

## 🖥️ Platform Support

* ✅ Windows (with proper taskbar icon grouping)
* ✅ macOS
* ✅ Linux

---

## ⚠️ Common Issues

* **Excel file won’t save** → Close it if already open in Excel.
* **PermissionError** → Move the app and Excel file to a writable folder (e.g., Desktop/Documents).

---

## 📜 License

MIT License – free to use, modify, and distribute.

