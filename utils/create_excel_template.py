"""
Script to create Excel template for portfolio address testing
"""
import pandas as pd


def create_portfolio_template(output_path: str = "test_data/portfolio_addresses_template.xlsx"):
    """
    Create Excel template with sample data

    Template structure:
    | Portfolio Name | Chain | Address | Tags | Expected Balance | Notes |
    """
    # Sample data
    data = {
        "Portfolio Name": [
            "Portfolio 1",
            "Portfolio 1",
            "Portfolio 2",
            "Portfolio 2",
            "Portfolio 3",
        ],
        "Chain": [
            "ETH",
            "BSC",
            "ETH",
            "Polygon",
            "ETH"
        ],
        "Address": [
            "0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb",
            "0x8894E0a0c962CB723c1976a4421c95949bE2D4E3",
            "0xD551234Ae421e3BcBa6A39BA2e21e3cE6e2E1234",
            "0xE12345f1234567890abcdef1234567890abcde12",
            "0xF234567890abcdef1234567890abcdef12345678"
        ],
        "Tags": [
            "Exchange, Hot Wallet",
            "Trading",
            "Personal, Long-term",
            "DeFi, Yield Farming",
            "NFTs, Collectibles"
        ],
        "Expected Balance": [
            10000.50,
            5000.25,
            50000.00,
            15000.75,
            25000.00
        ],
        "Notes": [
            "Main trading wallet",
            "BSC trading account",
            "Primary holding wallet",
            "Polygon DeFi portfolio",
            "NFT collection wallet"
        ]
    }

    # Create DataFrame
    df = pd.DataFrame(data)

    # Write to Excel with formatting
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

    print(f"✅ Excel template created: {output_path}")
    print(f"📊 Template contains {len(df)} sample addresses across {df['Portfolio Name'].nunique()} portfolios")
    print("\nPortfolio Summary:")
    print(df.groupby('Portfolio Name')['Address'].count())


if __name__ == "__main__":
    create_portfolio_template()
