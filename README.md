# AAT.ECF Cross-Validation Project

This project automates the processing and validation of AAT and ECF data with three main modules.

## Project Structure

```
New_AAT.ECF_CrossValidation/
├── main.py                      # Main entry point
├── config.py                    # Centralized configuration
├── modules/
│   ├── __init__.py
│   ├── cross_validation.py      # Module 1: Cross-validation report
│   ├── copy_comments.py         # Module 2: Copy AAT comments
│   └── large_deal_summary.py    # Module 3: Large deal summary for Dave
├── Cross-validation.py          # Legacy script (deprecated)
└── README.md                    # This file
```

## Quick Start

### 1. Set the Date

Edit `main.py` and set the processing date:

```python
DATE_STR = '20251130'  # Format: YYYYMMDD
```

### 2. Choose Modules to Run

Comment/uncomment the module lines in `main.py`:

```python
run_cross_validation(DATE_STR)       # Module 1: Cross-Validation Report
run_copy_comments(DATE_STR)          # Module 2: Copy Comments from Previous Version
run_large_deal_summary(DATE_STR)     # Module 3: Large Deal Summary for Dave
```

### 3. Run the Project

```bash
python main.py
```

Or run modules independently:

```bash
# Run individual modules
python modules/cross_validation.py
python modules/copy_comments.py
python modules/large_deal_summary.py
```

> **Note**: When running modules independently, set the `DEFAULT_DATE` variable at the bottom of each module file.

## Modules

### Module 1: Cross-Validation Report

**Purpose**: Generate AAT vs ECF comparison report with highlighting and categorization.

**Input**:
- `S:/Audrey/Audrey/AAT.DCF/{date}/AAT vs ECF {date}.xlsx`
- `{AAT_OUTPUT_BASE_PATH}/{date}/Status_Final_{date}.xlsx`

**Output**:
- `S:/Audrey/Audrey/AAT.DCF/{date}/AAT vs ECF {date}.xlsx` (updated)

**Features**:
- Merges AAT and ECF data with Status_Final (deal-level MV)
- Calculates IRR and Duration differences
- Highlights significant discrepancies
- Categorizes deals based on thresholds
- Creates two highlight tabs: IRR Diffs and Duration Diffs

### Module 2: Copy AAT Comments

**Purpose**: Automatically copy AAT Comments from previous version to current version.

**Input**:
- Versioned files in: `C:\Users\guoau\Magnetar Capital LLC\WKG - AAT - General\Project Management\AAT vs DCF`
- File pattern: `AAT vs ECF {date}.v{version}.xlsx`

**Output**:
- Updates AAT Comments in latest version file

**Logic**:
- Finds latest version (e.g., v5)
- Copies comments from v(n-1) (e.g., v4)
- If current is v1, copies from last version of previous month
- Updates two tabs: "Highlight IRR Diffs" and "Highlight Duration Diffs"

### Module 3: Large Deal Summary for Dave

**Purpose**: Create a simplified summary report of large deals.

**Input**:
- `C:\Users\guoau\Magnetar Capital LLC\WKG - Desk Cash Flows-DCF Production - Documents\Project Management\AAT vs ECF discrepancies\AAT vs ECF {date}.xlsx`

**Output**:
- `S:\Audrey\Audrey\Large Deal Summary for Dave\Large Deals Summary for Dave.xlsx`

**Processing Steps**:
1. Copy source file and rename to "Large Deals Summary for Dave.xlsx"
2. Process Summary tab (skip pivot table at top, work on data table below)
3. Keep only specified columns:
   - Deal Name
   - Sr. Portfolio Manager
   - {date} AAT IRR
   - {date} ECF IRR
   - AAT&ECF IRR Diffs
   - Duration AAT
   - Duration ECF
   - Duration Diffs
   - Liq Cap
   - % LC (calculated)
   - Category
4. Remove all deals containing "CoreWeave" in Deal Name
5. Calculate % LC (percentage of total Liq Cap for each deal)
6. Highlight top 10 deals by Liq Cap in gold

## Configuration

Edit `config.py` to customize:

### Paths

```python
BASE_PATH = 'S:/Audrey/Audrey/AAT.DCF'
AAT_OUTPUT_BASE_PATH = '//evprodfsg01/QR_Workspace/AssetAllocation/CFValidation/Prod/AATOutput'
VERSIONED_FILES_FOLDER = r'C:\Users\guoau\Magnetar Capital LLC\WKG - AAT - General\Project Management\AAT vs DCF'
AAT_ECF_SUMMARY_REPORT = r'C:\Users\guoau\Magnetar Capital LLC\WKG - Desk Cash Flows-DCF Production - Documents\Project Management\AAT vs ECF discrepancies'
LARGE_DEAL_SUMMARY_FOLDER = r'S:\Audrey\Audrey\Large Deal Summary for Dave'
```

### Thresholds

```python
SIGNIFICANT_MV_THRESHOLD = 25_000_000  # $25 million for significance
IRR_DIFF_THRESHOLD = 0.05              # 5% IRR difference threshold
DURATION_DIFF_THRESHOLD = 0.5          # 0.5 duration difference threshold
```

## Requirements

### Python Version
- Python 3.8+

### Required Packages

```bash
pip install pandas openpyxl python-dateutil
```

Or install from requirements file (if available):

```bash
pip install -r requirements.txt
```

## Workflow

### Typical Monthly Process

1. **Generate Cross-Validation Report** (Module 1)
   ```python
   DATE_STR = '20251130'
   run_cross_validation(DATE_STR)
   ```

2. **Copy Comments from Previous Version** (Module 2)
   ```python
   run_copy_comments(DATE_STR)
   ```

3. **Create Large Deal Summary** (Module 3)
   ```python
   run_large_deal_summary(DATE_STR)
   ```

### Standalone Module Testing

Each module can be run independently by:
1. Opening the module file (e.g., `modules/large_deal_summary.py`)
2. Setting `DEFAULT_DATE` at the bottom
3. Running: `python modules/large_deal_summary.py`

## Troubleshooting

### No Data in Output
- Ensure input files exist for the specified date
- Check that Status_Final file has deal-level rows (Instrument column is NaN)

### File Corruption Warning
- Module 2 uses `keep_vba=True, keep_links=True` to preserve Excel formatting
- If warnings persist, check source file integrity

### Module Not Running
- Check that required input files exist
- Verify paths in `config.py`
- Ensure date format is correct (YYYYMMDD)

## Notes

- **Date Format**: Always use YYYYMMDD format (e.g., '20251130')
- **File Versioning**: Module 2 expects versioned files like `AAT vs ECF 20251130.v2.xlsx`
- **Pivot Tables**: Module 3 preserves pivot tables in the output for email copying
- **CoreWeave Deals**: Always removed in Module 3 before % LC calculation

## Legacy Files

- `Cross-validation.py`: Original script, kept for reference only
- Use the modular structure in `modules/` for all new development
