"""
AAT.ECF Cross-Validation Project - Main Entry Point

Usage:
    1. Set the date below (DATE_STR)
    2. Comment/uncomment module lines to control which modules to run
    3. Run: python main.py
"""

import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Import modules
from modules.cross_validation import run_cross_validation
from modules.historical_validation_comments import run_copy_comments
from modules.large_deal_summary import run_large_deal_summary



def main():
    """
    Main entry point.

    To run/skip modules: Comment/uncomment the lines below.
    """

    # ===== Configuration =====
    # Set the processing date here (format: YYYYMMDD)
    DATE_STR = '20251130'

    print("=" * 80)
    print("AAT.ECF Cross-Validation Project")
    print(f"Processing Date: {DATE_STR}")
    print("=" * 80)
    print()

    # ===== Run Modules =====
    # Comment out any module you don't want to run

    # run_cross_validation(DATE_STR)           # Module 1: Cross-Validation Report
    # run_copy_comments(DATE_STR)            # Module 2: Historical Validation Comments
    run_large_deal_summary(DATE_STR)       # Module 3: Large Deal Summary for Dave

    print()
    print("=" * 80)
    print("All modules completed!")
    print("=" * 80)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nFatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
