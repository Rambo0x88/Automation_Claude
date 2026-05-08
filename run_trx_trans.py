#!/usr/bin/env python3
"""
TRX Transaction History Extraction & Comparison

Thin orchestrator that calls utils/trx_transaction/trongrid_dam_comparison.py
with sensible defaults (date defaults to today).

Usage:
  python3 run_trx_trans.py <TRX_ADDRESS>                        # today's date
  python3 run_trx_trans.py <TRX_ADDRESS> <DDMMYYYY>             # specific date
  python3 run_trx_trans.py <TRX_ADDRESS> <FROM_DDMMYYYY> <TO>   # date range
  python3 run_trx_trans.py <PORTFOLIO_NAME>                      # today, DAM-only
  python3 run_trx_trans.py <PORTFOLIO_NAME> <DDMMYYYY>           # specific date, DAM-only

Examples:
  python3 run_trx_trans.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6
  python3 run_trx_trans.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6 16042026
  python3 run_trx_trans.py trx2_Mkx
  python3 run_trx_trans.py trx2_Mkx 01032026 31032026
"""

import os
import sys
import subprocess
from datetime import datetime

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_COMPARISON_SCRIPT = os.path.join(_SCRIPT_DIR, "utils", "trx_transaction", "trongrid_dam_comparison.py")


def _today_ddmmyyyy():
    """Return today's date in DDMMYYYY format."""
    return datetime.now().strftime("%d%m%Y")


def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python3 run_trx_trans.py <TRX_ADDRESS_OR_PORTFOLIO_NAME>")
        print("  python3 run_trx_trans.py <TRX_ADDRESS_OR_PORTFOLIO_NAME> <DDMMYYYY>")
        print("  python3 run_trx_trans.py <TRX_ADDRESS_OR_PORTFOLIO_NAME> <FROM_DDMMYYYY> <TO_DDMMYYYY>")
        print()
        print("If no date is provided, defaults to today.")
        print()
        print("Examples:")
        print("  python3 run_trx_trans.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6")
        print("  python3 run_trx_trans.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6 16042026")
        print("  python3 run_trx_trans.py trx2_Mkx")
        sys.exit(1)

    target = sys.argv[1]

    # Build args for the comparison script
    if len(sys.argv) >= 4:
        # Date range provided: target FROM TO
        date_from = sys.argv[2]
        date_to = sys.argv[3]
        cmd_args = [sys.executable, _COMPARISON_SCRIPT, target, date_from, date_to]
    elif len(sys.argv) >= 3:
        # Single date provided: target DATE
        date = sys.argv[2]
        cmd_args = [sys.executable, _COMPARISON_SCRIPT, target, date]
    else:
        # No date — default to today
        today = _today_ddmmyyyy()
        cmd_args = [sys.executable, _COMPARISON_SCRIPT, target, today]
        print(f"ℹ️  No date specified, using today: {today}")

    print(f"\n{'='*70}")
    print(f"TRX Transaction History Comparison")
    print(f"{'='*70}")
    print(f"Target: {target}")
    print(f"Command: {' '.join(cmd_args)}")
    print(f"{'='*70}\n")

    # Run the comparison script
    result = subprocess.run(cmd_args, cwd=_SCRIPT_DIR)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
