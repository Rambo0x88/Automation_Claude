"""
Helper script to create Excel file with your test addresses
Edit the addresses list below with your real addresses, then run this script
"""
import pandas as pd
from datetime import datetime


def create_addresses_excel(output_path: str = "test_data/my_portfolio_addresses.xlsx"):
    """
    Create Excel file with your portfolio addresses

    INSTRUCTIONS:
    1. Edit the 'addresses' list below with your real wallet addresses
    2. Run: python3 utils/create_test_addresses_excel.py
    3. The Excel file will be created at: test_data/my_portfolio_addresses.xlsx
    """

    # =========================================================================
    # EDIT THIS SECTION WITH YOUR REAL ADDRESSES
    # =========================================================================

    addresses = [
        # Portfolio 1
        {
            "Portfolio Name": "Portfolio 1",
            "Chain": "ETH",
            "Address": "0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb",
            "Tags": "Main, Trading",
            "Expected Balance": 10000,
            "Notes": "Main trading wallet"
        },
        {
            "Portfolio Name": "Portfolio 1",
            "Chain": "BSC",
            "Address": "0x8894E0a0c962CB723c1976a4421c95949bE2D4E3",
            "Tags": "Trading",
            "Expected Balance": 5000,
            "Notes": "BSC trading"
        },

        # Portfolio 2
        {
            "Portfolio Name": "Portfolio 2",
            "Chain": "ETH",
            "Address": "0xD551234Ae421e3BcBa6A39BA2e21e3cE6e2E1234",
            "Tags": "Personal",
            "Expected Balance": 50000,
            "Notes": "Primary holding"
        },

        # Add more addresses here...
        # Just copy the format above and change the values
    ]

    # =========================================================================
    # DON'T EDIT BELOW THIS LINE
    # =========================================================================

    # Create DataFrame
    df = pd.DataFrame(addresses)

    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Portfolios', index=False)

        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Portfolios']

        # Auto-adjust column widths
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(col)
            )
            worksheet.column_dimensions[chr(65 + idx)].width = max_length + 2

    print(f"✅ Excel file created: {output_path}")
    print(f"📊 Created {len(df)} address entries across {df['Portfolio Name'].nunique()} portfolios")
    print("\nPortfolio Summary:")
    print(df.groupby('Portfolio Name')['Address'].count())
    print("\nChain Distribution:")
    print(df.groupby('Chain')['Address'].count())
    print(f"\n💡 Tip: Edit this script to add your real addresses, then run it again!")


if __name__ == "__main__":
    create_addresses_excel()
