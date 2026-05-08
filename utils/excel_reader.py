"""
Excel Reader Utilities for Portfolio Address Management
Reads portfolio addresses and test data from Excel files
"""
import pandas as pd
from typing import List, Dict, Optional
import os


class ExcelReader:
    """Utility class for reading portfolio test data from Excel files"""

    def __init__(self, excel_path: str):
        """
        Initialize Excel reader

        Args:
            excel_path: Path to Excel file containing portfolio addresses
        """
        self.excel_path = excel_path

        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        self.df = None
        self.portfolios = []

    def load_portfolios(self, sheet_name: str = "Portfolios") -> List[Dict]:
        """
        Load portfolio data from Excel

        Expected Excel structure:
        | Portfolio Name | Chain | Address | Tags | Expected Balance | Notes |
        |----------------|-------|---------|------|------------------|-------|
        | Portfolio 1    | ETH   | 0x123...| tag1 | 1000             | ...   |
        | Portfolio 1    | BSC   | 0x456...| tag2 | 500              | ...   |
        | Portfolio 2    | ETH   | 0x789...| tag3 | 2000             | ...   |

        Returns:
            List of portfolio dictionaries:
            [
                {
                    "name": "Portfolio 1",
                    "addresses": [
                        {"chain": "ETH", "address": "0x123...", "tags": ["tag1"], "expected_balance": 1000},
                        {"chain": "BSC", "address": "0x456...", "tags": ["tag2"], "expected_balance": 500}
                    ]
                },
                {
                    "name": "Portfolio 2",
                    "addresses": [...]
                }
            ]
        """
        print(f"📖 Loading portfolios from Excel: {self.excel_path}")

        try:
            # Try reading with default sheet name first, fallback to first sheet
            try:
                self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            except:
                # If sheet not found, use first sheet
                self.df = pd.read_excel(self.excel_path, sheet_name=0)
                sheet_name = "Sheet1"

            print(f"✅ Loaded {len(self.df)} rows from sheet '{sheet_name}'")

            # Group by portfolio name
            portfolios_dict = {}

            for index, row in self.df.iterrows():
                portfolio_name = str(row.get('Portfolio Name', '')).strip()
                if not portfolio_name:
                    continue

                # Initialize portfolio if not exists
                if portfolio_name not in portfolios_dict:
                    portfolios_dict[portfolio_name] = {
                        "name": portfolio_name,
                        "test_email": str(row.get('Test Email', '')).strip(),
                        "password": str(row.get('Password', '')).strip(),
                        "addresses": []
                    }

                # Parse address data
                # Chain is optional - if not provided, DAM will auto-detect
                address_data = {
                    "chain": str(row.get('Chain', 'ETH')).strip() if 'Chain' in row and pd.notna(row.get('Chain')) else 'ETH',
                    "address": str(row.get('Address', '')).strip(),
                    "tags": self._parse_tags(row.get('Tags', '')) if 'Tags' in row else [],
                    "expected_balance": row.get('Expected Balance', None) if 'Expected Balance' in row else None,
                    "notes": str(row.get('Notes', '')).strip() if 'Notes' in row else ''
                }

                # Validate address
                if address_data['address'] and address_data['address'].startswith('0x'):
                    portfolios_dict[portfolio_name]["addresses"].append(address_data)

            # Convert to list
            self.portfolios = list(portfolios_dict.values())

            # Print summary
            total_addresses = sum(len(p['addresses']) for p in self.portfolios)
            print(f"📊 Loaded {len(self.portfolios)} portfolios with {total_addresses} total addresses")

            for portfolio in self.portfolios:
                print(f"  - {portfolio['name']}: {len(portfolio['addresses'])} addresses")

            return self.portfolios

        except Exception as e:
            print(f"❌ Error loading portfolios from Excel: {e}")
            raise

    def _parse_tags(self, tags_str) -> List[str]:
        """
        Parse tags string into list

        Args:
            tags_str: Tags as string (comma/semicolon separated)

        Returns:
            List of tag strings
        """
        if pd.isna(tags_str):
            return []

        tags_str = str(tags_str).strip()
        if not tags_str:
            return []

        # Split by comma or semicolon
        tags = [tag.strip() for tag in tags_str.replace(';', ',').split(',')]
        return [tag for tag in tags if tag]

    def get_all_addresses(self) -> List[str]:
        """
        Get flat list of all addresses across all portfolios

        Returns:
            List of address strings
        """
        all_addresses = []
        for portfolio in self.portfolios:
            for addr_data in portfolio['addresses']:
                all_addresses.append(addr_data['address'])

        return all_addresses

    def get_addresses_by_portfolio(self, portfolio_name: str) -> List[Dict]:
        """
        Get addresses for specific portfolio

        Args:
            portfolio_name: Name of portfolio

        Returns:
            List of address dictionaries
        """
        for portfolio in self.portfolios:
            if portfolio['name'] == portfolio_name:
                return portfolio['addresses']

        return []

    def get_addresses_by_chain(self, chain: str) -> List[str]:
        """
        Get all addresses for specific blockchain

        Args:
            chain: Chain name (e.g., "ETH", "BSC")

        Returns:
            List of addresses for that chain
        """
        addresses = []
        for portfolio in self.portfolios:
            for addr_data in portfolio['addresses']:
                if addr_data['chain'].upper() == chain.upper():
                    addresses.append(addr_data['address'])

        return addresses


def load_test_addresses(excel_path: str, sheet_name: str = "Portfolios") -> List[Dict]:
    """
    Convenience function to load portfolio addresses from Excel

    Args:
        excel_path: Path to Excel file
        sheet_name: Name of sheet to read (default: "Portfolios")

    Returns:
        List of portfolio dictionaries

    Example:
        >>> portfolios = load_test_addresses("test_data/portfolios.xlsx")
        >>> print(portfolios[0]['name'])
        'My ETH Portfolio'
        >>> print(portfolios[0]['addresses'][0]['address'])
        '0x123...'
    """
    reader = ExcelReader(excel_path)
    return reader.load_portfolios(sheet_name)
