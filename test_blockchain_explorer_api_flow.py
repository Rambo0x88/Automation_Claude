"""
Blockchain Explorer API Flow - Alternative to Manual CSV Export

This is a SEPARATE flow that uses Etherscan/BSCScan/BaseScan APIs directly
instead of manual CSV downloads. Run this INSTEAD of the main run_overview.py
if you want API-based blockchain explorer data.

Usage:
    python3 test_blockchain_explorer_api_flow.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain ethereum
    python3 test_blockchain_explorer_api_flow.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain bsc
    python3 test_blockchain_explorer_api_flow.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain base
"""

import os
import sys
import json
import argparse
from datetime import datetime
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))

from utils.blockchain_explorer_api import (
    fetch_all_explorer_data,
    export_explorer_data_to_excel,
    EXPLORER_APIS,
)


def load_api_keys() -> dict:
    """Load API keys from config file."""
    config_file = "test_data/blockchain_explorer_api_keys.json"
    
    if not os.path.exists(config_file):
        print(f"\n❌ Config file not found: {config_file}")
        print(f"\nCreate it with your API keys:")
        print(f"{{")
        print(f'  "ethereum": "YOUR_ETHERSCAN_API_KEY",')
        print(f'  "bsc": "YOUR_BSCSCAN_API_KEY",')
        print(f'  "base": "YOUR_BASESCAN_API_KEY"')
        print(f"}}\n")
        return {}
    
    try:
        with open(config_file, 'r') as f:
            return json.load(f)
    except Exception as e:
        print(f"❌ Error loading API keys: {e}")
        return {}


def run_blockchain_explorer_api_flow(
    evm_addresses: list = None,
    chain: str = "ethereum",
    api_key: str = None,
    output_folder: str = "test-results/blockchain-explorer-api"
):
    """
    Run blockchain explorer API flow for given addresses.
    
    This is a SEPARATE flow from the main run_overview.py
    Use this when you want API-based blockchain explorer data instead of manual CSV export.
    
    Args:
        evm_addresses: List of EVM addresses to fetch data for
        chain: "ethereum", "bsc", or "base"
        api_key: Explorer API key (if not provided, loads from config)
        output_folder: Where to save Excel files
    """
    
    print("\n" + "="*80)
    print("BLOCKCHAIN EXPLORER API FLOW (Alternative to Manual CSV Export)")
    print("="*80)
    
    if not evm_addresses:
        print("❌ No EVM addresses provided")
        return
    
    # Create output directory
    os.makedirs(output_folder, exist_ok=True)
    
    # Load API keys if not provided
    if not api_key:
        api_keys = load_api_keys()
        api_key = api_keys.get(chain)
        
        if not api_key:
            print(f"❌ No API key found for {chain}")
            print(f"   Please set up {output_folder}/../blockchain_explorer_api_keys.json")
            return
    
    explorer = EXPLORER_APIS.get(chain)
    if not explorer:
        print(f"❌ Unknown chain: {chain}")
        print(f"   Supported: ethereum, bsc, base")
        return
    
    print(f"\n📊 Using {explorer['name']} API")
    print(f"   Chain: {chain}")
    print(f"   Addresses: {len(evm_addresses)}")
    print(f"   Output: {output_folder}")
    
    # Fetch data for each address
    all_results = []
    for idx, address in enumerate(evm_addresses, 1):
        print(f"\n{'='*60}")
        print(f"[{idx}/{len(evm_addresses)}] Address: {address}")
        print(f"{'='*60}")
        
        # Fetch all data
        explorer_data = fetch_all_explorer_data(address, chain, api_key)
        
        if not explorer_data["success"]:
            print(f"❌ Failed to fetch data for {address}")
            continue
        
        # Export to Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"{explorer['name']}_{address[:10]}_{timestamp}.xlsx"
        excel_path = os.path.join(output_folder, excel_filename)
        
        if export_explorer_data_to_excel(explorer_data, excel_path):
            result = {
                "address": address,
                "chain": chain,
                "excel_file": excel_filename,
                "excel_path": excel_path,
                "transactions": len(explorer_data["transactions"]),
                "internal_transactions": len(explorer_data["internal_transactions"]),
                "token_transfers": len(explorer_data["token_transfers"]),
                "nft_transfers": len(explorer_data["nft_transfers"]),
            }
            all_results.append(result)
            
            print(f"✅ Data fetched and exported")
            print(f"   Transactions: {result['transactions']}")
            print(f"   Internal Transactions: {result['internal_transactions']}")
            print(f"   Token Transfers: {result['token_transfers']}")
            print(f"   NFT Transfers: {result['nft_transfers']}")
            print(f"   File: {excel_filename}")
        else:
            print(f"❌ Failed to export data for {address}")
    
    # Print summary
    print(f"\n{'='*80}")
    print("SUMMARY")
    print(f"{'='*80}")
    print(f"Processed: {len(all_results)}/{len(evm_addresses)} addresses")
    
    for result in all_results:
        print(f"\n✅ {result['address']}")
        print(f"   Transactions: {result['transactions']}")
        print(f"   Internal Transactions: {result['internal_transactions']}")
        print(f"   Token Transfers: {result['token_transfers']}")
        print(f"   NFT Transfers: {result['nft_transfers']}")
        print(f"   File: {result['excel_file']}")
    
    print(f"\n✅ All files saved to: {output_folder}")
    print(f"{'='*80}\n")
    
    return all_results


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Blockchain Explorer API Flow - Fetch transaction data via API (Alternative to Manual CSV Export)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 test_blockchain_explorer_api_flow.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain ethereum
  python3 test_blockchain_explorer_api_flow.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain bsc
  python3 test_blockchain_explorer_api_flow.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab 0x1234567890abcdef --chain ethereum
        """
    )
    parser.add_argument(
        "--evm",
        nargs="+",
        required=True,
        help="EVM address(es) to fetch data for"
    )
    parser.add_argument(
        "--chain",
        default="ethereum",
        choices=["ethereum", "bsc", "base"],
        help="Blockchain to query (default: ethereum)"
    )
    parser.add_argument(
        "--api-key",
        help="Explorer API key (if not provided, loads from config)"
    )
    parser.add_argument(
        "--output",
        default="test-results/blockchain-explorer-api",
        help="Output folder for Excel files"
    )
    
    args = parser.parse_args()
    
    run_blockchain_explorer_api_flow(
        evm_addresses=args.evm,
        chain=args.chain,
        api_key=args.api_key,
        output_folder=args.output,
    )


if __name__ == "__main__":
    main()
