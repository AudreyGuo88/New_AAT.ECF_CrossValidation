"""
Project Configuration File

Global constants for paths and thresholds.
Note: Date is configured in main.py, not here.
"""

# ===== Path Configuration =====
BASE_PATH = 'S:/Audrey/Audrey/AAT.DCF'
AAT_OUTPUT_BASE_PATH = '//evprodfsg01/QR_Workspace/AssetAllocation/CFValidation/Prod/AATOutput'

# Source folder for versioned files (used by copy_comments module)
VERSIONED_FILES_FOLDER = r'C:\Users\guoau\Magnetar Capital LLC\WKG - AAT - General\Project Management\AAT vs DCF'

# AAT vs ECF discrepancies summary report folder
AAT_ECF_SUMMARY_REPORT = r'C:\Users\guoau\Magnetar Capital LLC\WKG - Desk Cash Flows-DCF Production - Documents\Project Management\AAT vs ECF discrepancies'

# Large Deal Summary output folder for Dave
LARGE_DEAL_SUMMARY_FOLDER = r'S:\Audrey\Audrey\Large Deal Summary for Dave'

# ===== Threshold Configuration =====
SIGNIFICANT_MV_THRESHOLD = 25_000_000  # $25 million threshold for significance
IRR_DIFF_THRESHOLD = 0.05  # 5% IRR difference threshold
DURATION_DIFF_THRESHOLD = 0.5  # 0.5 duration difference threshold
