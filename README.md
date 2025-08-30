Great — I see your new version adds an **Activity sheet** with Start/Stop tracking, a **stacked main/activity view**, extra shortcuts, and refined UI logic.
Here’s the **updated README.md** to match this codebase:

---

# 🚨 Incident Number Tracker

A desktop tool built with **PySide6** and **openpyxl** to help teams log and manage **incident tickets** and their **activity timelines**.
Incidents are stored in an Excel file (`incident_numbers.xlsx`) with two sheets:

* **INCIDENTS** → log of tickets with date, ID, and description
* **Activity** → start/stop records with timestamps per ticket

---

## ✨ Features

### Incident Management

* 📅 **Date Picker** – Select or auto-fill the "Created On" date.
* 🔢 **Smart Ticket ID**

  * Auto-suggests next sequential ID in format: **`THyymmddNN`**
  * Example: `TH25083001`, `TH25083002` for Aug 30, 2025.
  * You can override with your own (e.g., `INC2453903`).
* 📝 **Multi-line Description** – Input multiple lines; stored as a single joined line.
* 👀 **Preview Panel** – Shows single-line preview of your description.
* 📊 **Excel Integration** – All tickets stored in `incident_numbers.xlsx` → **INCIDENTS** sheet.

### Activity Tracking

* ⏱️ **Start/Stop Buttons** – Log activity sessions per ticket with precise timestamps.
* 🗂 **Activity Sheet** – Records:

  * Ticket ID
  * Start Time
  * End Time
* 🚫 **No Overlaps** – Prevents starting a new activity for a ticket if one is already running.
* 📋 **Activity Table** – Displays history of all sessions for a ticket.

### Usability

* 🔄 **Quick Actions**

  * Add new incident
  * Clear description
  * Refresh lists
  * Open Excel directly from the app
* ⚡ **Keyboard Shortcuts**

  * `Ctrl+Enter` → Add incident
  * `Ctrl+L` → Clear description
  * `Ctrl+T` → Today’s date
  * `Ctrl+O` → Open Excel file
  * `F5` → Refresh incidents
  * `Ctrl+Shift+S` → Start activity (Activity view)
  * `Ctrl+E` → Stop activity (Activity view)
* 📋 **Context Menu** – Right-click a row → Copy row / View Activity.
* 🎨 **Modern UI**

  * Red banner to distinguish from other trackers
  * Alternating row colors
  * Clear tooltips & hints

---

## 📂 File Structure

* `incident_numbers.xlsx`

  * `INCIDENTS` → Ticket log
  * `Activity` → Start/Stop sessions
* `app.py` → Main script (GUI).

---

## 🚀 Getting Started

### 1. Clone or Download

```bash
git clone https://github.com/engkufizz/IncidentNumber_Tracker.git
cd IncidentNumber_Tracker
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

### INCIDENTS sheet

| Created On | Ticket ID  | Description                        |
| ---------- | ---------- | ---------------------------------- |
| 2025-08-30 | TH25083001 | Router crash, rebooted, monitoring |

### Activity sheet

| Ticket ID  | Start Time          | End Time            |
| ---------- | ------------------- | ------------------- |
| TH25083001 | 2025-08-30 09:15:00 | 2025-08-30 09:45:12 |

---

## 🖥️ Platform Support

* ✅ Windows (with proper taskbar icon grouping)
* ✅ macOS
* ✅ Linux

---

## ⚠️ Common Issues

* **Excel file won’t save** → Close it if already open in Excel.
* **PermissionError** → Move the app and Excel file to a writable folder (e.g., Desktop/Documents).
* **Open Activity not stopping** → Ensure you’re on the correct ticket row before stopping.

---

## 📜 License

MIT License – free to use, modify, and distribute.

