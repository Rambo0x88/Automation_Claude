#!/usr/bin/env python3
"""
Full Portfolio Check — Overview + Transaction History

Orchestrator that runs:
  1. Overview Extraction & Comparison (run_overview.py)
  2. Transaction History Comparison (run_trx_trans.py / run_eth_trans.py)

Detects address types and runs the applicable pipelines.
Produces 2 separate Excel outputs: one for Overview, one for Transaction History.

Usage:
  python3 run_all.py --trx T...                                  # Overview + TRX transactions (today)
  python3 run_all.py --trx T... --date 16042026                  # Overview + TRX transactions (specific date)
  python3 run_all.py --trx T... --date 01032026 --date-to 31032026  # Overview + TRX transactions (date range)
  python3 run_all.py --evm 0x...                                  # Overview + ETH transactions (placeholder)
  python3 run_all.py --trx T... --evm 0x...                       # Overview (both) + TRX trans + ETH trans
  python3 run_all.py -p "portfolio_name"                           # Overview + auto-detect transactions
  python3 run_all.py -p "portfolio_name" --date 16042026           # Overview + transactions for date

Examples:
  python3 run_all.py --trx TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6
  python3 run_all.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab
  python3 run_all.py --trx TUqEg3dzVEJNQSVW2HY98z5X8SBdhmao8D --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --date 16042026
  python3 run_all.py -p trx2_Mkx --date 16042026
"""

import argparse
import os
import sys
import subprocess
from datetime import datetime

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _today_ddmmyyyy():
    return datetime.now().strftime("%d%m%Y")


def _is_trx_address(s):
    return s.startswith("T") and len(s) == 34 and s.isalnum()


def _is_evm_address(s):
    return s.startswith("0x") and len(s) == 42


def _run_script(script_name, args, label):
    """Run a script as subprocess and return the exit code."""
    script_path = os.path.join(_SCRIPT_DIR, script_name)
    cmd = [sys.executable, script_path] + args
    print(f"\n{'='*70}")
    print(f"▶ {label}")
    print(f"  Command: {' '.join(cmd)}")
    print(f"{'='*70}\n")
    result = subprocess.run(cmd, cwd=_SCRIPT_DIR)
    if result.returncode != 0:
        print(f"\n⚠️  {label} exited with code {result.returncode}")
    else:
        print(f"\n✅ {label} completed successfully")
    return result.returncode


def main():
    parser = argparse.ArgumentParser(
        description="Full Portfolio Check — Overview + Transaction History",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 run_all.py --trx TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6
  python3 run_all.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab
  python3 run_all.py --trx T... --evm 0x... --date 16042026
  python3 run_all.py -p trx2_Mkx --date 16042026
        """
    )
    parser.add_argument('--trx', type=str, action='append', default=[],
                        help='TRX address (T..., 34 chars). Can specify multiple.')
    parser.add_argument('--evm', type=str, action='append', default=[],
                        help='EVM address (0x..., 42 chars). Can specify multiple.')
    parser.add_argument('-p', '--portfolio', type=str, default=None,
                        help='Portfolio name (searches DAM dropdown)')
    parser.add_argument('--date', type=str, default=None,
                        help='Transaction date filter (DDMMYYYY). Defaults to today.')
    parser.add_argument('--date-to', type=str, default=None,
                        help='End date for range (DDMMYYYY). If omitted, same as --date.')
    parser.add_argument('-q', '--quiet', action='store_true',
                        help='Quiet mode for overview extraction')

    args = parser.parse_args()

    has_trx = len(args.trx) > 0
    has_evm = len(args.evm) > 0
    has_portfolio = args.portfolio is not None
    date = args.date or _today_ddmmyyyy()
    date_to = args.date_to

    if not has_trx and not has_evm and not has_portfolio:
        parser.print_help()
        print("\n❌ Please provide at least one of: --trx, --evm, or -p")
        sys.exit(1)

    print("=" * 70)
    print("FULL PORTFOLIO CHECK — Overview + Transaction History")
    print("=" * 70)
    if has_trx:
        print(f"  TRX addresses: {args.trx}")
    if has_evm:
        print(f"  EVM addresses: {args.evm}")
    if has_portfolio:
        print(f"  Portfolio: {args.portfolio}")
    print(f"  Transaction date: {date}" + (f" to {date_to}" if date_to else ""))
    print("=" * 70)

    results = {}

    # ── PART 1: Overview Extraction ──────────────────────────────────────
    overview_args = []
    if has_trx:
        for addr in args.trx:
            overview_args.extend(["--trx", addr])
    if has_evm:
        for addr in args.evm:
            overview_args.extend(["--evm", addr])
    if has_portfolio:
        overview_args.extend(["-p", args.portfolio])
    if args.quiet:
        overview_args.append("-q")

    results["overview"] = _run_script("run_overview.py", overview_args, "PART 1: Overview Extraction & Comparison")

    # ── PART 2: TRX Transaction History ──────────────────────────────────
    if has_trx:
        for trx_addr in args.trx:
            trans_args = [trx_addr, date]
            if date_to:
                trans_args.append(date_to)
            results[f"trx_trans_{trx_addr[-8:]}"] = _run_script(
                "run_trx_trans.py", trans_args,
                f"PART 2: TRX Transaction History ({trx_addr[-8:]})"
            )
    elif has_portfolio and not has_evm:
        # Portfolio name mode — pass to TRX transaction script (DAM-only, no TronGrid)
        trans_args = [args.portfolio, date]
        if date_to:
            trans_args.append(date_to)
        results["trx_trans_portfolio"] = _run_script(
            "run_trx_trans.py", trans_args,
            f"PART 2: TRX Transaction History ({args.portfolio})"
        )

    # ── PART 3: ETH Transaction History ──────────────────────────────────
    if has_evm:
        for evm_addr in args.evm:
            results[f"eth_trans_{evm_addr[-8:]}"] = _run_script(
                "run_eth_trans.py", [evm_addr],
                f"PART 3: ETH Transaction History ({evm_addr[-8:]})"
            )

    # ── Summary ──────────────────────────────────────────────────────────
    print(f"\n{'='*70}")
    print("SUMMARY")
    print(f"{'='*70}")
    for step, code in results.items():
        status = "✅" if code == 0 else "⚠️"
        print(f"  {status} {step}: exit code {code}")
    print(f"{'='*70}")

    # Exit with non-zero if any step failed
    if any(code != 0 for code in results.values()):
        sys.exit(1)


if __name__ == "__main__":
    main()
