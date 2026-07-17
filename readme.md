# Madrasa Data Automation Pipeline

Welcome to the Madrasa Data Automation Pipeline! This project automates the extraction, enrichment, and visualization of Madrasa data in West Bengal. It takes raw image data, processes it via OCR, supplements it with geographical and transport data (trains, auto stands, etc.), and compiles it into a highly interactive, macro-enabled Excel dashboard.

---

## 🏗️ Project Architecture

This project is divided into three main operational phases, all orchestrated by Python scripts and automated batch files. 

### 1. Data Source (`extracted_data.csv`)
The canonical source of truth for all Madrasa data. This CSV stores the core information such as the Madrasa Name, Code, PIN, Raabta Centre, Total Students, etc. It acts as the bridge between the Python scripts and the final Excel workbook.

### 2. The Excel Dashboard (`sheets/forms_v4.xlsm`)
A highly stylized and fully automated Excel workbook containing:
- **Forms**: A frontend UI for manual data entry, featuring the Madrasa Entry Portal, as well as 3 new "Smart Forms" (Student Information, Teacher Information, Madaris Information) with dynamic SmartArt integrations and VBA Submit buttons.
- **Database**: The master table where all Madrasa data is stored, including dynamic calculations for `Total No of Students Involved` and `Madaris Involved` grouped by `Raabta Code`.
- **Dashboard**: A search interface to query the database by PIN Code or District.
- **Hidden DBs**: `Student_DB`, `Teacher_DB`, and `Madaris_DB` store the relational data submitted via the Smart Forms.

---

## 🚀 How to Run the Pipeline

The pipeline is split into three sequential batch files. You can run these by simply double-clicking them in the root directory.

### Phase 1: OCR & Data Injection (`run_phase1_ocr.bat`)
This phase is responsible for reading raw images and pushing the extracted text into our Excel database.
1. **OCR Extraction (`scripts/ocr_pipeline.py`)**: Reads images from the `images/` folder, extracts text using OCR, processes the data, and updates `extracted_data.csv` (using `(CODE, NAME OF MADRASAH, PIN)` as a unique composite key to prevent duplicates).
2. **Excel Injection (`scripts/inject_csv_to_excel.py`)**: Safely injects the newly extracted data into `forms_v4.xlsm`. 
   - *Note: It uses COM objects to attach to Excel. The Excel file must be CLOSED in the GUI before running this step to prevent write locks.*

### Phase 2: Train & Station Locator (`run_phase2_trains.bat`)
This phase adds inter-city transport logistics to the dataset.
1. **Find Trains (`scripts/find_trains.py`)**: Uses APIs/geodata to determine the *Nearest Railway Station* and *Train from Sealdah* for each Madrasa based on its PIN/Village.
2. **Inject Trains (`scripts/inject_trains.py`)**: Updates the `extracted_data.csv` and the Excel Database with this new transportation data.

### Phase 3: Local Transport Locator (`run_phase3_local.bat`)
This phase handles the "last mile" logistics.
1. **Find Local Transport (`scripts/find_local.py`)**: Locates the closest *Auto Stand* and *Main Bazaar* for each Madrasa.
2. **Inject Local Transport (`scripts/inject_local.py`)**: Pushes this final layer of data into the Excel Database.

---

## 🛠️ Advanced Tools & Scripts

If you need to make structural changes to the project, you can use the standalone Python scripts located in the `scripts/` directory:

- **`build_excel.py`**: The architect script. If the Excel file gets corrupted, deleted, or if you need to redesign the UI, running `python scripts/build_excel.py` will rebuild `forms_v4.xlsm` from scratch. It generates all sheets, applies color themes (Emerald, Blue, Purple, Amber), writes the VBA macros, and merges existing data from `extracted_data.csv`.
- **`compare_and_merge.py`**: A utility script used to cross-reference our data against external datasets to fill in missing gaps without overwriting existing manual entries.
- **`reparse_missed.py`**: A fallback script to re-run OCR extraction specifically on images that failed or were skipped during Phase 1.
- **`check_errors.py`**: Parses `error_tracker.csv` to help you diagnose why certain rows failed validation or were flagged as duplicates during injection.

---

## ⚠️ Important Notes

1. **"Security Risk / MotW" Banner in Excel**: 
   If Windows blocks macros in `forms_v4.xlsm` because it came from another computer:
   - Close Excel.
   - Right-click the `.xlsm` file -> `Properties`.
   - Check **Unblock** at the bottom of the General tab, click OK.
   - *Pro-tip: You can add the `sheets/` folder to Excel's "Trusted Locations" (Options > Trust Center > Trust Center Settings > Trusted Locations) to prevent this forever.*

2. **Rearranging Excel Columns**: 
   You can manually cut/paste columns in the Excel file, but doing so will break the hardcoded column mappings (like `Entry Date`, `PIN`) inside the VBA macros. If you move columns, you MUST press `Alt+F11` in Excel and update the column index numbers in `Module1`, or else the Search and Submit features will crash. If you run `build_excel.py` again, your manual column shifts will be overwritten by the Python script layout.