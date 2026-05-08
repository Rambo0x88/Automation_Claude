#!/usr/bin/env python3
"""
Script to manually create portfolios in an existing DAM account.
This is used to set up an account with portfolios before running TC1.
"""
import os
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import pandas as pd

# Add parent directory to path to import page objects
sys.path.insert(0, str(Path(__file__).parent.parent))

from pages.sign_in_page import SignInPage
from pages.portfolio_page import PortfolioPage
from utils.excel_reader import ExcelReader

def create_portfolios_for_account(email: str, password: str, excel_file: str):
    """
    Create portfolios in an existing DAM account.

    Args:
        email: Account email
        password: Account password
        excel_file: Path to Excel file with portfolio data
    """
    print("=" * 80)
    print("Manual Portfolio Creation Script")
    print("=" * 80)
    print(f"\n📧 Account: {email}")
    print(f"📁 Excel File: {excel_file}\n")

    # Load portfolios from Excel
    print("📖 Loading portfolios from Excel...")
    reader = ExcelReader(excel_file)
    portfolios = reader.load_portfolios()
    print(f"✅ Loaded {len(portfolios)} portfolios with {sum(len(p['addresses']) for p in portfolios)} total addresses\n")

    with sync_playwright() as p:
        # Launch browser
        print("🌐 Launching browser...")
        browser = p.chromium.launch(headless=False, slow_mo=500)
        context = browser.new_context(viewport={'width': 1280, 'height': 720})
        page = context.new_page()

        try:
            # Sign in
            print(f"🔐 Signing in as {email}...")
            sign_in_page = SignInPage(page)
            sign_in_page.navigate()
            sign_in_page.sign_in(email, password)
            print("✅ Signed in successfully\n")

            # Navigate to portfolio page
            print("📂 Navigating to portfolio page...")
            portfolio_page = PortfolioPage(page)
            portfolio_page.navigate()
            print("✅ On portfolio page\n")

            # Create each portfolio
            print(f"💼 Creating {len(portfolios)} portfolios...\n")
            created_count = 0

            for idx, portfolio_data in enumerate(portfolios, 1):
                portfolio_name = portfolio_data['name']
                addresses = portfolio_data['addresses']

                print(f"  [{idx}/{len(portfolios)}] Creating portfolio: {portfolio_name}")
                print(f"  Addresses: {len(addresses)}")

                try:
                    # Check if we're on Overview page (need to click dropdown first)
                    # or if we're already on create portfolio dialog
                    is_on_overview = page.locator("text=Overview").is_visible(timeout=3000)

                    if is_on_overview:
                        print(f"  📊 Overview page detected, opening dropdown menu...")
                        # Click dropdown button
                        dropdown_button = page.locator("button:has-text('Portfolio'), button[aria-haspopup='menu']").first
                        dropdown_button.click()
                        print(f"  ⏳ Waiting 3 seconds for dropdown menu...")
                        page.wait_for_timeout(3000)

                        # Click "Create portfolio" from dropdown
                        print(f"  ➕ Clicking 'Create portfolio' from dropdown...")
                        portfolio_page.click_create_portfolio()
                        print(f"  ⏳ Waiting 3 seconds after clicking Create portfolio...")
                        page.wait_for_timeout(3000)
                    else:
                        print(f"  📝 Already on create form, proceeding directly...")
                        # Already on create form, just proceed
                        portfolio_page.click_create_portfolio()

                    # Enter portfolio name
                    portfolio_page.enter_portfolio_name(portfolio_name)

                    # Add each address
                    for addr_idx, address_data in enumerate(addresses, 1):
                        address = address_data['address']
                        chain = address_data.get('chain', 'ETH')
                        tags = address_data.get('tags', [])

                        print(f"    Adding address {addr_idx}/{len(addresses)}: {address[:10]}...")

                        # Add address
                        portfolio_page.add_address(
                            address=address,
                            chain=chain,
                            tags=tags
                        )

                    # Save portfolio
                    print(f"  💾 Saving portfolio '{portfolio_name}'...")
                    portfolio_page.save_portfolio()
                    print(f"  ⏳ Waiting 3 seconds after clicking Save...")
                    page.wait_for_timeout(3000)

                    created_count += 1
                    print(f"  ✅ Portfolio '{portfolio_name}' created successfully ({created_count}/{len(portfolios)})")

                    # Take full-page screenshot
                    from datetime import datetime
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    clean_name = portfolio_name.replace(' ', '_').replace('/', '_').replace("'", '')
                    screenshot_path = f"test-results/screenshots/portfolio_{clean_name}_{timestamp}.png"
                    print(f"  📸 Taking full-page screenshot...")
                    page.screenshot(path=screenshot_path, full_page=True)
                    print(f"  ✅ Screenshot saved: {screenshot_path}")

                    # Wait for portfolio to be fully saved
                    print(f"  ⏳ Waiting 3 seconds for portfolio to be saved...")
                    page.wait_for_timeout(3000)
                    print(f"  ✅ Ready for next portfolio\n")

                except Exception as e:
                    print(f"  ❌ Error creating portfolio '{portfolio_name}': {e}\n")
                    continue

            print("=" * 80)
            print(f"✅ Portfolio creation completed!")
            print(f"   Created: {created_count}/{len(portfolios)} portfolios")
            print("=" * 80)

            # Keep browser open for verification
            print("\n⏸️  Browser will stay open for 10 seconds for verification...")
            page.wait_for_timeout(10000)

        except Exception as e:
            print(f"\n❌ Error: {e}")
            import traceback
            traceback.print_exc()
        finally:
            browser.close()

def main():
    """Main entry point."""
    # Load environment variables
    load_dotenv()

    # Get credentials from .env
    email = os.getenv('TEST_EMAIL')
    password = os.getenv('TEST_PASSWORD')

    if not email or not password:
        print("❌ Error: TEST_EMAIL and TEST_PASSWORD must be set in .env file")
        sys.exit(1)

    # Get Excel file path
    excel_file = os.getenv('PORTFOLIO_EXCEL_PATH', 'test_data/DAM addresses.xlsx')

    if not os.path.exists(excel_file):
        print(f"❌ Error: Excel file not found: {excel_file}")
        sys.exit(1)

    # Create portfolios
    create_portfolios_for_account(email, password, excel_file)

if __name__ == '__main__':
    main()
