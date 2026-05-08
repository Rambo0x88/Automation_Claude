"""
Blockchain Explorer API Test - STEP 5 (API-based alternative)

This is a NEW flow that uses Etherscan/BSCScan/BaseScan APIs directly
instead of manual CSV downloads. The old CSV method is still available.

Usage:
    python3 -m tests.test_blockchain_explorer_api --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain ethereum --api-key YOUR_KEY
"""

import os
import sys
import json
import argparse
from datetime import datetime
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils.blockchain_explorer_api import (
    fetch_all_explorer_data,
    export_explorer_data_to_excel,
    EXPLORER_APIS,
)


def load_api_keys() -> dict:
    """Load API keys from config file."""
    config_file = "test_data/blockchain_explorer_api_keys.json"
    
    if not os.path.exists(config_file):
        print(f"⚠️  Config file not found: {config_file}")
        print(f"   Create it with:")
        print(f"   {{")
        print(f'     "ethereum": "YOUR_ETHERSCAN_API_KEY",')
        print(f'     "bsc": "YOUR_BSCSCAN_API_KEY",')
        print(f'     "base": "YOUR_BASESCAN_API_KEY"')
        print(f"   }}")
        return {}
    
    try:
        with open(config_file, 'r') as f:
            return json.load(f)
    except Exception as e:
        print(f"❌ Error loading API keys: {e}")
        return {}


def run_blockchain_explorer_api_test(
    evm_addresses: list = None,
    chain: str = "ethereum",
    api_key: str = None,
    output_folder: str = "test-results/blockchain-explorer-api"
):
    """
    Run blockchain explorer API test for given addresses.
    
    Args:
        evm_addresses: List of EVM addresses to fetch data for
        chain: "ethereum", "bsc", or "base"
        api_key: Explorer API key (if not provided, loads from config)
        output_folder: Where to save Excel files
    """
    
    print("\n" + "="*80)
    print("STEP 5: Blockchain Explorer API Test (NEW - API-based flow)")
    print("="*80)
    
    if not evm_addresses:
        print("⏭️  Skipped (no EVM addresses provided)")
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
        return
    
    print(f"\n📊 Using {explorer['name']} API")
    print(f"   Chain: {chain}")
    print(f"   Addresses: {len(evm_addresses)}")
    
    # Fetch data for each address
    all_results = []
    for address in evm_addresses:
        print(f"\n{'='*60}")
        print(f"Address: {address}")
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
            all_results.append({
                "address": address,
                "chain": chain,
                "excel_file": excel_filename,
                "excel_path": excel_path,
                "transactions": len(explorer_data["transactions"]),
                "internal_transactions": len(explorer_data["internal_transactions"]),
                "token_transfers": len(explorer_data["token_transfers"]),
                "nft_transfers": len(explorer_data["nft_transfers"]),
            })
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
        description="Blockchain Explorer API Test - Fetch transaction data via API"
    )
    parser.add_argument(
        "--evm",
        nargs="+",
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
    
    if not args.evm:
        parser.print_help()
        return
    
    run_blockchain_explorer_api_test(
        evm_addresses=args.evm,
        chain=args.chain,
        api_key=args.api_key,
        output_folder=args.output,
    )


if __name__ == "__main__":
    main()
