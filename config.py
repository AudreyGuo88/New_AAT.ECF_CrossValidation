"""
Project Configuration File

Global constants for paths and thresholds.
Note: Date is configured in main.py, not here.
"""

# ===== Path Configuration =====
BASE_PATH = 'S:/Audrey/Audrey/AAT.DCF'
AAT_OUTPUT_BASE_PATH = '//evprodfsg01/QR_Workspace/AssetAllocation/CFValidation/Prod/AATOutput'

# ===== Threshold Configuration =====
SIGNIFICANT_MV_THRESHOLD = 25_000_000  # $25 million threshold for significance
IRR_DIFF_THRESHOLD = 0.05  # 5% IRR difference threshold
DURATION_DIFF_THRESHOLD = 0.5  # 0.5 duration difference threshold
