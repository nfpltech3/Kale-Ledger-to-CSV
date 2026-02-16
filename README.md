# Ledger to CSV Converter

Converts ledger reports to Logisys-compatible CSVs.

## Tech Stack
- Python 3.11
- Tkinter (GUI)
- Pandas (Data Processing)
- OpenPyXL (Excel)
- Pillow (Image handling)

---

## Installation

### Clone
```bash
git clone https://github.com/username/ledger-to-csv.git
cd ledger-to-csv
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. Create virtual environment
```bash
python -m venv venv
```

2. Activate (REQUIRED)

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. Run application
```bash
python Ledger_to_CSV.py
```

---

### Build Executable

1. Install PyInstaller (Inside venv):
```bash
pip install pyinstaller
```

2. Build using the included Spec file (Ensure you do not run main.py directly):
```bash
pyinstaller Ledger_to_CSV.spec
```

3. Locate Executable:
The application will be generated in the `dist/` folder.

---

## Usage

1. Launch application.
2. Select Job Register CSV.
3. Select Ledger Excel Report.
4. Click 'Process'.
5. Output CSV is saved in `Kale Output` directory.
