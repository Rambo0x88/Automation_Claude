"""
Helper script to create Excel file with portfolio addresses and associated test emails
For daily testing with specific test accounts
"""
import pandas as pd
from datetime import datetime


def create_addresses_with_emails_excel(output_path: str = "test_data/daily_test_addresses.xlsx"):
    """
    Create Excel file with portfolio addresses and their associated test emails

    INSTRUCTIONS:
    1. Edit the 'test_accounts' list below with your real addresses and emails
    2. Email format: zgtestdataDDMM001@ (e.g., zgtestdata120101@merqbcqa.33mail.com)
    3. Run: python3 utils/create_addresses_with_emails.py
    4. Excel file created at: test_data/daily_test_addresses.xlsx
    """

    # =========================================================================
    # EDIT THIS SECTION WITH YOUR REAL TEST ACCOUNTS AND ADDRESSES
    # =========================================================================

    test_accounts = [
        # Test Account 1: zgtestdata120101@...
        {
            "Test Email": "zgtestdata120101@merqbcqa.33mail.com",
            "Password": "YourPassword123!",
            "Portfolio Name": "Main Portfolio",
            "Chain": "ETH",
            "Address": "0x742d35Cc6634C0532925a3b844Bc9e7595f0bEb",
            "Tags": "Main, Trading",
            "Expected Balance": 10000,
            "Notes": "Main trading wallet"
        },
        {
            "Test Email": "zgtestdata120101@merqbcqa.33mail.com",
            "Password": "YourPassword123!",
            "Portfolio Name": "Main Portfolio",
            "Chain": "BSC",
            "Address": "0x8894E0a0c962CB723c1976a4421c95949bE2D4E3",
            "Tags": "Trading",
            "Expected Balance": 5000,
            "Notes": "BSC trading account"
        },

        # Test Account 2: zgtestdata120102@...
        {
            "Test Email": "zgtestdata120102@merqbcqa.33mail.com",
            "Password": "YourPassword123!",
            "Portfolio Name": "DeFi Portfolio",
            "Chain": "ETH",
            "Address": "0xD551234Ae421e3BcBa6A39BA2e21e3cE6e2E1234",
            "Tags": "DeFi, Staking",
            "Expected Balance": 50000,
            "Notes": "Primary DeFi holdings"
        },
        {
            "Test Email": "zgtestdata120102@merqbcqa.33mail.com",
            "Password": "YourPassword123!",
            "Portfolio Name": "DeFi Portfolio",
            "Chain": "Polygon",
            "Address": "0xE12345f1234567890abcdef1234567890abcde12",
            "Tags": "Yield Farming",
            "Expected Balance": 15000,
            "Notes": "Polygon yield farming"
        },

        # Test Account 3: zgtestdata120103@...
        {
            "Test Email": "zgtestdata120103@merqbcqa.33mail.com",
            "Password": "YourPassword123!",
            "Portfolio Name": "NFT Portfolio",
            "Chain": "ETH",
            "Address": "0xF234567890abcdef1234567890abcdef12345678",
            "Tags": "NFTs, Collectibles",
            "Expected Balance": 25000,
            "Notes": "NFT collection wallet"
        },

        # Add more test accounts and addresses here...
        # Just copy the format above and change the values
    ]

    # =========================================================================
    # DON'T EDIT BELOW THIS LINE
    # =========================================================================

    # Create DataFrame
    df = pd.DataFrame(test_accounts)

    # Reorder columns for better readability
    column_order = [
        "Test Email",
        "Password",
        "Portfolio Name",
        "Chain",
        "Address",
        "Tags",
        "Expected Balance",
        "Notes"
    ]
    df = df[column_order]

    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Test Accounts', index=False)

        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Test Accounts']

        # Auto-adjust column widths
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(col)
            )
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)

    print(f"✅ Excel file created: {output_path}")
    print(f"\n📊 Test Account Summary:")
    print(f"   Total addresses: {len(df)}")
    print(f"   Unique test emails: {df['Test Email'].nunique()}")
    print(f"   Portfolios: {df['Portfolio Name'].nunique()}")

    print("\n📧 Test Accounts:")
    for email in df['Test Email'].unique():
        email_data = df[df['Test Email'] == email]
        portfolios = email_data['Portfolio Name'].unique()
        address_count = len(email_data)
        print(f"   {email}")
        print(f"      - {len(portfolios)} portfolio(s): {', '.join(portfolios)}")
        print(f"      - {address_count} address(es)")

    print("\n🔗 Chain Distribution:")
    print(df.groupby('Chain')['Address'].count().to_string())

    print(f"\n💡 Tip: Edit this script to add your real test accounts and addresses!")
    print(f"📝 Email format: zgtestdataDDMM001@merqbcqa.33mail.com")


if __name__ == "__main__":
    create_addresses_with_emails_excel()
