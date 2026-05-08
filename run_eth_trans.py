#!/usr/bin/env python3
"""
ETH Transaction History Extraction & Comparison (Placeholder)

This script will handle ETH/EVM transaction history extraction and comparison
against blockchain explorer data (Etherscan, BSCScan, BaseScan, etc.).

Status: NOT YET IMPLEMENTED — waiting for ETH transaction comparison code.

Usage (future):
  python3 run_eth_trans.py <EVM_ADDRESS>
  python3 run_eth_trans.py <EVM_ADDRESS> <DDMMYYYY>
  python3 run_eth_trans.py <EVM_ADDRESS> <FROM_DDMMYYYY> <TO_DDMMYYYY>
"""

import sys


def main():
    print("=" * 70)
    print("ETH Transaction History Comparison")
    print("=" * 70)
    print()
    print("⏳ NOT YET IMPLEMENTED")
    print()
    print("This script will be implemented when the ETH transaction")
    print("comparison code is available. It will:")
    print("  1. Fetch transaction history from Etherscan/BSCScan/BaseScan API")
    print("  2. Extract transaction rows from DAM Transactions tab")
    print("  3. Build side-by-side comparison Excel")
    print()
    print("For now, use the blockchain explorer API flow:")
    print("  python3 -m tests.test_blockchain_explorer_api --evm <ADDRESS> --chain ethereum")
    print()
    print("=" * 70)
    sys.exit(0)


if __name__ == "__main__":
    main()
