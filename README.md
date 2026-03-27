# Intuit Hiring Dashboard

An interactive dashboard for tracking open hiring requisitions and expected hires, built with Streamlit.

---

## Quick Start (teammates)

### Step 1 — Install Python (one time only)
Download and install Python 3.9 or later from **https://www.python.org/downloads/**

- **Mac:** Download the `.pkg` installer and run it
- **Windows:** Download the `.exe` installer — make sure to tick **"Add Python to PATH"** during install

Verify it worked by opening Terminal (Mac) or Command Prompt (Windows) and running:
```
python3 --version    # Mac
python --version     # Windows
```

### Step 2 — Get the code (one time only)
Clone this repo, or download it as a ZIP from GitHub:

**Option A — Clone with git:**
```bash
git clone https://github.com/mohankumar-int/hiring-dashboard.git
cd hiring-dashboard
```

**Option B — Download ZIP:**
1. Go to https://github.com/mohankumar-int/hiring-dashboard
2. Click the green **Code** button → **Download ZIP**
3. Unzip it and open the folder

### Step 3 — Run the dashboard

**Mac / Linux:**
```bash
bash setup_and_run.sh
```

**Windows:**
Double-click `setup_and_run.bat`
*(or run it from Command Prompt)*

The script will automatically install all dependencies and open the dashboard at **http://localhost:8501**

---

## Daily Use

Every day when you want to open the dashboard, just re-run the same script:

| Mac/Linux | `bash setup_and_run.sh` |
|-----------|------------------------|
| Windows   | Double-click `setup_and_run.bat` |

Then open **http://localhost:8501** in your browser.

---

## Loading fresh data

The dashboard reads data from two Excel files that you upload via the sidebar each session:

| File | What it contains |
|------|-----------------|
| `Open Hiring Requisitions Standard.xlsx` | All open reqs — source for Tab 1 |
| `Expected Hires.xlsx` | Pipeline offers — source for Tab 2 & Pipeline filter |

Your previous session's data loads automatically on restart. Upload new files whenever you receive a refreshed export.

---

## Getting updates

When the dashboard is updated, pull the latest code before running:

```bash
git pull
bash setup_and_run.sh
```

---

## Requirements
- Python 3.9+
- Internet connection (first run only, to install packages)
- The two Excel data files (shared separately)
