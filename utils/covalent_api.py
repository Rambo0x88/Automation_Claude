"""
Covalent API Utility for Address Discovery
Fetches blockchain activity data to find addresses with ETH-only, BSC-only, or both
"""
import requests
from typing import Dict, List, Optional
from config.config import Config


class CovalentAPI:
    """Covalent API wrapper for blockchain address discovery"""

    # Covalent API Chain IDs
    ETH_CHAIN_ID = "1"  # Ethereum Mainnet
    BSC_CHAIN_ID = "56"  # Binance Smart Chain

    def __init__(self, api_key: str = None):
        """
        Initialize Covalent API client
        Args:
            api_key: Covalent API key (defaults to Config.COVALENT_API_KEY)
        """
        self.api_key = api_key or getattr(Config, 'COVALENT_API_KEY', '')
        self.base_url = "https://api.covalenthq.com/v1"

    def get_address_activity(self, address: str, chain_id: str = ETH_CHAIN_ID) -> Optional[Dict]:
        """
        Get activity data for an address on a specific chain

        Args:
            address: Wallet address (e.g., '0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb')
            chain_id: Chain ID (1 for ETH, 56 for BSC)

        Returns:
            Dict with activity data or None if request fails
        """
        if not self.api_key:
            print("⚠️ Warning: COVALENT_API_KEY not configured")
            return None

        url = f"{self.base_url}/address/{address}/activity/?key={self.api_key}"

        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()

            if data.get('error'):
                print(f"⚠️ API Error: {data.get('error_message')}")
                return None

            return data.get('data', {})
        except requests.exceptions.RequestException as e:
            print(f"⚠️ Request failed: {e}")
            return None

    def has_eth_activity(self, address: str) -> bool:
        """
        Check if address has Ethereum activity

        Args:
            address: Wallet address

        Returns:
            True if address has ETH activity, False otherwise
        """
        activity = self.get_address_activity(address, self.ETH_CHAIN_ID)
        if not activity:
            return False

        # Check if there are any transactions or token holdings
        items = activity.get('items', [])
        return len(items) > 0

    def has_bsc_activity(self, address: str) -> bool:
        """
        Check if address has BSC activity

        Args:
            address: Wallet address

        Returns:
            True if address has BSC activity, False otherwise
        """
        activity = self.get_address_activity(address, self.BSC_CHAIN_ID)
        if not activity:
            return False

        # Check if there are any transactions or token holdings
        items = activity.get('items', [])
        return len(items) > 0

    def get_address_type(self, address: str) -> str:
        """
        Determine if address has ETH-only, BSC-only, or both

        Args:
            address: Wallet address

        Returns:
            'eth_only', 'bsc_only', 'both', or 'none'
        """
        has_eth = self.has_eth_activity(address)
        has_bsc = self.has_bsc_activity(address)

        if has_eth and has_bsc:
            return 'both'
        elif has_eth:
            return 'eth_only'
        elif has_bsc:
            return 'bsc_only'
        else:
            return 'none'

    def find_addresses_by_type(self, addresses: List[str]) -> Dict[str, List[str]]:
        """
        Categorize a list of addresses by their blockchain activity

        Args:
            addresses: List of wallet addresses

        Returns:
            Dict with keys 'eth_only', 'bsc_only', 'both', 'none'
            Example: {
                'eth_only': ['0x123...'],
                'bsc_only': ['0x456...'],
                'both': ['0x789...'],
                'none': []
            }
        """
        categorized = {
            'eth_only': [],
            'bsc_only': [],
            'both': [],
            'none': []
        }

        for address in addresses:
            print(f"🔍 Checking address: {address}")
            addr_type = self.get_address_type(address)
            categorized[addr_type].append(address)
            print(f"   Type: {addr_type}")

        return categorized

    def get_sample_addresses(self) -> Dict[str, str]:
        """
        Get sample addresses for testing (you can replace with real addresses)

        Returns:
            Dict with sample addresses by type
        """
        # These are example addresses - replace with actual addresses that have activity
        return {
            'eth_only': '0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb',
            'bsc_only': '0xAb5801a7D398351b8bE11C439e05C5B3259aeC9B',
            'both': '0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb'  # Replace with actual multi-chain address
        }


def find_test_addresses(api_key: str = None) -> Dict[str, str]:
    """
    Helper function to find suitable test addresses

    Args:
        api_key: Covalent API key

    Returns:
        Dict with addresses categorized by type
    """
    client = CovalentAPI(api_key)

    # List of candidate addresses to check (you can add more)
    candidate_addresses = [
        '0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb',
        '0xAb5801a7D398351b8bE11C439e05C5B3259aeC9B',
        '0x8626f6940E2eb28930eFb4CeF49B2d1F2C9C1199',  # Example address
        # Add more addresses here
    ]

    categorized = client.find_addresses_by_type(candidate_addresses)

    print("\n" + "="*80)
    print("📊 ADDRESS CATEGORIZATION RESULTS")
    print("="*80)
    print(f"ETH Only: {len(categorized['eth_only'])} addresses")
    for addr in categorized['eth_only']:
        print(f"  - {addr}")

    print(f"\nBSC Only: {len(categorized['bsc_only'])} addresses")
    for addr in categorized['bsc_only']:
        print(f"  - {addr}")

    print(f"\nBoth Chains: {len(categorized['both'])} addresses")
    for addr in categorized['both']:
        print(f"  - {addr}")

    print("="*80 + "\n")

    return {
        'eth_only': categorized['eth_only'][0] if categorized['eth_only'] else None,
        'bsc_only': categorized['bsc_only'][0] if categorized['bsc_only'] else None,
        'both': categorized['both'][0] if categorized['both'] else None
    }


if __name__ == "__main__":
    # Test the API
    import os
    api_key = os.getenv('COVALENT_API_KEY', '')

    if not api_key:
        print("⚠️ Please set COVALENT_API_KEY environment variable")
    else:
        print("🔍 Finding test addresses using Covalent API...")
        addresses = find_test_addresses(api_key)

        print("\n✅ Found addresses:")
        print(f"ETH Only: {addresses.get('eth_only')}")
        print(f"BSC Only: {addresses.get('bsc_only')}")
        print(f"Both Chains: {addresses.get('both')}")
