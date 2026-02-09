# AAT.ECF Cross-Validation Project

This project processes and validates AAT (Asset Allocation Tool) and ECF (Enterprise Cash Flow) data.

## Project Structure

```
New_AAT.ECF_CrossValidation/
├── main.py                    # Main entry point
├── config.py                  # Configuration file
├── utils.py                   # Utility functions
├── modules/
│   ├── __init__.py
│   ├── cross_validation.py   # Cross-validation report module
│   └── [future modules]
├── Cross-validation.py        # Legacy script (keep for reference)
└── README.md                  # This file
```

## Quick Start

### Run All Modules

```bash
python main.py
```

### Run Specific Modules

**Method 1: Using config.py**
Edit `config.py` and set the desired flags:
```python
ENABLE_CROSS_VALIDATION = True  # Enable
# ENABLE_MODULE2 = False         # Disable
```

**Method 2: Comment/Uncomment in main.py**
Open `main.py` and comment out the modules you don't want to run.

## Configuration

Edit `config.py` to change:
- **Date**: `DATE_STR = '20251130'`
- **Paths**: `BASE_PATH`, `AAT_OUTPUT_BASE_PATH`
- **Thresholds**: `SIGNIFICANT_MV_THRESHOLD`, `IRR_DIFF_THRESHOLD`, etc.

## Modules

### 1. Cross-Validation Report
Generates AAT vs ECF comparison reports with highlighting and categorization.

## Adding New Modules

1. Create `modules/new_module.py`
2. Add `ENABLE_NEW_MODULE = True` to `config.py`
3. Add module call to `main.py`

## Requirements

- Python 3.8+
- pandas
- openpyxl
