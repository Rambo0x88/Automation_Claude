"""
Utility script to create a portfolio with all addresses from an Excel file

Usage:
    python utils/create_portfolio_from_excel.py --excel "test_data/DAM addresses.xlsx" --portfolio-name "All Addresses"
    python utils/create_portfolio_from_excel.py --excel "test_data/DAM addresses.xlsx" --portfolio-name "My Portfolio" --email "user@example.com" --password "password"
"""
import argparse
import sys
import os
import pandas as pd
from playwright.sync_api import sync_playwright
from pages.sign_in_page import SignInPage
from pages.portfolio_page import PortfolioPage
from utils.excel_reader import ExcelReader
from config.config import Config


def get_all_addresses_from_excel(excel_path: str) -> list:
    """
    Extract all unique addresses from Excel file
    
    Supports multiple Excel formats:
    - Standard format with "Portfolio Name", "Address" columns
    - Simple format with just "Address" column
    - Any format that ExcelReader can parse
    
    Args:
        excel_path: Path to Excel file
        
    Returns:
        List of unique address strings
    """
    print(f"📖 Reading addresses from Excel: {excel_path}")
    
    try:
        # Try using ExcelReader first (for standard format)
        try:
            reader = ExcelReader(excel_path)
            portfolios = reader.load_portfolios()
            
            # Collect all unique addresses from all portfolios
            all_addresses = set()
            for portfolio in portfolios:
                for addr_data in portfolio['addresses']:
                    address = addr_data.get('address', '').strip()
                    if address and address.startswith('0x'):
                        all_addresses.add(address)
            
            if len(all_addresses) > 0:
                addresses_list = sorted(list(all_addresses))
                print(f"✅ Found {len(addresses_list)} unique addresses in Excel file")
                return addresses_list
        except Exception as e:
            print(f"⚠️  ExcelReader format failed, trying direct Excel read: {e}")
        
        # Fallback: Try reading Excel directly (for simple format with just addresses)
        try:
            df = pd.read_excel(excel_path, sheet_name=0)
            print(f"📊 Loaded {len(df)} rows from Excel")
            
            # Look for address column (case-insensitive)
            address_column = None
            for col in df.columns:
                if 'address' in str(col).lower():
                    address_column = col
                    break
            
            if address_column is None:
                # Try first column if no address column found
                address_column = df.columns[0]
                print(f"⚠️  No 'Address' column found, using first column: {address_column}")
            
            # Extract addresses
            all_addresses = set()
            for idx, row in df.iterrows():
                address = str(row[address_column]).strip()
                # Validate address format
                if address and address.startswith('0x') and len(address) >= 42:
                    all_addresses.add(address[:42])  # Ensure 42 chars (0x + 40 hex)
            
            addresses_list = sorted(list(all_addresses))
            print(f"✅ Found {len(addresses_list)} unique addresses in Excel file")
            return addresses_list
            
        except Exception as e2:
            print(f"❌ Error reading Excel file: {e2}")
            raise Exception(f"Could not read addresses from Excel. Please ensure the file has an 'Address' column or uses the standard portfolio format. Error: {e2}")
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        raise


def create_portfolio_with_addresses(
    excel_path: str,
    portfolio_name: str = "Portfolio from Excel",
    email: str = None,
    password: str = None,
    headless: bool = False
):
    """
    Create a portfolio with all addresses from Excel file
    
    Args:
        excel_path: Path to Excel file containing addresses
        portfolio_name: Name for the new portfolio
        email: Email for sign-in (uses Config if not provided)
        password: Password for sign-in (uses Config if not provided)
        headless: Run browser in headless mode
    """
    # Get credentials - try multiple sources
    if not email:
        # Try loading from tc1_account.json first
        tc1_account_path = os.path.join(Config.PROJECT_ROOT, "test_data", "tc1_account.json")
        if os.path.exists(tc1_account_path):
            try:
                import json
                with open(tc1_account_path, "r") as f:
                    account_data = json.load(f)
                    email = account_data.get('email', '')
                    print(f"📁 Loaded email from {tc1_account_path}")
            except Exception as e:
                print(f"⚠️  Could not load from {tc1_account_path}: {e}")
        
        # Fallback to config
        if not email:
            email = Config.TEST_EMAIL
    
    if not password:
        # Try loading from tc1_account.json first
        tc1_account_path = os.path.join(Config.PROJECT_ROOT, "test_data", "tc1_account.json")
        if os.path.exists(tc1_account_path):
            try:
                import json
                with open(tc1_account_path, "r") as f:
                    account_data = json.load(f)
                    password = account_data.get('password', '')
                    print(f"📁 Loaded password from {tc1_account_path}")
            except Exception as e:
                print(f"⚠️  Could not load from {tc1_account_path}: {e}")
        
        # Fallback to config
        if not password:
            password = Config.TEST_PASSWORD
    
    if not email or not password:
        print("❌ Error: Email and password required.")
        print("   Options:")
        print("   1. Pass --email and --password as arguments")
        print("   2. Set TEST_EMAIL and TEST_PASSWORD environment variables")
        print("   3. Place credentials in test_data/tc1_account.json")
        sys.exit(1)
    
    # Read addresses from Excel
    addresses = get_all_addresses_from_excel(excel_path)
    
    if len(addresses) == 0:
        print("❌ No addresses found in Excel file. Please check the file format.")
        sys.exit(1)
    
    print(f"\n📋 Will create portfolio '{portfolio_name}' with {len(addresses)} addresses")
    print(f"   First address: {addresses[0][:10]}...{addresses[0][-8:]}")
    print(f"   Last address: {addresses[-1][:10]}...{addresses[-1][-8:]}")
    
    # Launch browser
    with sync_playwright() as p:
        print(f"\n🌐 Launching browser (headless={headless})...")
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            record_video_dir='videos' if not headless else None
        )
        page = context.new_page()
        
        try:
            # Step 1: Sign in
            print(f"\n🔐 Step 1: Signing in as {email}...")
            sign_in_page = SignInPage(page, Config.BASE_URL)
            sign_in_page.navigate()
            page.wait_for_timeout(2000)
            
            # Check for CAPTCHA and handle it
            try:
                from utils.captcha_solver import wait_for_captcha_and_solve
                print("🔍 Checking for CAPTCHA...")
                captcha_solved = wait_for_captcha_and_solve(page, timeout=120000)
                if not captcha_solved:
                    print("⚠️  CAPTCHA was not solved, but continuing...")
            except Exception as e:
                print(f"⚠️  CAPTCHA check failed: {e}, continuing...")
            
            # Enter credentials and sign in
            sign_in_page.sign_in(email, password)
            page.wait_for_timeout(5000)  # Wait longer for sign-in to complete
            
            # Verify sign-in success - check URL and wait for redirect
            max_wait = 10
            sign_in_success = False
            for i in range(max_wait):
                current_url = page.url
                if '/portfolio' in current_url or '/sign-in' not in current_url:
                    # Check if we're actually logged in by looking for portfolio elements
                    try:
                        # Look for portfolio page indicators
                        portfolio_indicators = [
                            'button:has-text("Create portfolio")',
                            'button:has-text("Portfolio")',
                            '[class*="portfolio"]'
                        ]
                        for indicator in portfolio_indicators:
                            if page.locator(indicator).count() > 0:
                                sign_in_success = True
                                break
                    except:
                        pass
                    
                    if sign_in_success or '/portfolio' in current_url:
                        sign_in_success = True
                        break
                
                page.wait_for_timeout(1000)
            
            if not sign_in_success:
                print("❌ Sign-in failed. Please check credentials.")
                print(f"   Current URL: {page.url}")
                page.screenshot(path="test-results/screenshots/signin_failed.png")
                print("📸 Screenshot saved: test-results/screenshots/signin_failed.png")
                sys.exit(1)
            
            print("✅ Signed in successfully")
            
            # Step 2: Navigate to portfolio page
            print(f"\n📂 Step 2: Navigating to portfolio page...")
            portfolio_page = PortfolioPage(page)
            portfolio_page.navigate()
            page.wait_for_timeout(2000)
            
            assert portfolio_page.is_on_portfolio_page(), "Failed to reach portfolio page"
            print("✅ On portfolio page")
            
            # Step 3: Click Create Portfolio
            print(f"\n➕ Step 3: Clicking Create Portfolio button...")
            portfolio_page.click_create_portfolio()
            page.wait_for_timeout(2000)
            print("✅ Create Portfolio dialog opened")
            
            # Step 4: Enter portfolio name
            print(f"\n✏️  Step 4: Entering portfolio name: '{portfolio_name}'...")
            portfolio_page.enter_portfolio_name(portfolio_name)
            page.wait_for_timeout(1000)
            print("✅ Portfolio name entered")
            
            # Step 5: Add all addresses
            print(f"\n📝 Step 5: Adding {len(addresses)} addresses...")
            failed_addresses = []
            
            for idx, address in enumerate(addresses, 1):
                try:
                    print(f"   [{idx}/{len(addresses)}] Adding: {address[:10]}...{address[-8:]}")
                    portfolio_page.add_address(address)
                    
                    # Add a small delay between addresses to avoid overwhelming the UI
                    if idx < len(addresses):
                        page.wait_for_timeout(500)
                    
                except Exception as e:
                    print(f"   ⚠️  Failed to add address {address[:10]}...{address[-8:]}: {e}")
                    failed_addresses.append(address)
                    continue
            
            if failed_addresses:
                print(f"\n⚠️  Warning: {len(failed_addresses)} addresses failed to add:")
                for addr in failed_addresses[:5]:  # Show first 5
                    print(f"   - {addr}")
                if len(failed_addresses) > 5:
                    print(f"   ... and {len(failed_addresses) - 5} more")
            
            successful_count = len(addresses) - len(failed_addresses)
            print(f"\n✅ Successfully added {successful_count}/{len(addresses)} addresses")
            
            # Step 6: Save portfolio
            print(f"\n💾 Step 6: Saving portfolio...")
            portfolio_page.click_save_portfolio()
            page.wait_for_timeout(5000)  # Wait for portfolio to be created
            
            # Verify portfolio was created
            print(f"\n✅ Portfolio '{portfolio_name}' created successfully!")
            print(f"   Total addresses: {successful_count}")
            if failed_addresses:
                print(f"   Failed addresses: {len(failed_addresses)}")
            
            # Keep browser open for a moment to verify
            if not headless:
                print(f"\n⏸️  Keeping browser open for 10 seconds for verification...")
                page.wait_for_timeout(10000)
            else:
                page.wait_for_timeout(2000)
            
            print("\n✅ Portfolio creation completed!")
            
        except Exception as e:
            print(f"\n❌ Error during portfolio creation: {e}")
            import traceback
            traceback.print_exc()
            
            # Take screenshot for debugging
            try:
                screenshot_path = "test-results/screenshots/create_portfolio_error.png"
                page.screenshot(path=screenshot_path)
                print(f"📸 Error screenshot saved: {screenshot_path}")
            except:
                pass
            
            sys.exit(1)
            
        finally:
            if not headless:
                print("\n⏸️  Closing browser in 3 seconds...")
                page.wait_for_timeout(3000)
            browser.close()


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description='Create a portfolio with all addresses from an Excel file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create portfolio with default name
  python utils/create_portfolio_from_excel.py --excel "test_data/DAM addresses.xlsx"
  
  # Create portfolio with custom name
  python utils/create_portfolio_from_excel.py --excel "test_data/DAM addresses.xlsx" --portfolio-name "All My Addresses"
  
  # Create portfolio with custom credentials
  python utils/create_portfolio_from_excel.py --excel "test_data/DAM addresses.xlsx" --email "user@example.com" --password "password"
  
  # Run in headless mode
  python utils/create_portfolio_from_excel.py --excel "test_data/DAM addresses.xlsx" --headless
        """
    )
    
    parser.add_argument(
        '--excel',
        type=str,
        required=True,
        help='Path to Excel file containing addresses'
    )
    
    parser.add_argument(
        '--portfolio-name',
        type=str,
        default='Portfolio from Excel',
        help='Name for the new portfolio (default: "Portfolio from Excel")'
    )
    
    parser.add_argument(
        '--email',
        type=str,
        default=None,
        help='Email for sign-in (uses Config.TEST_EMAIL if not provided)'
    )
    
    parser.add_argument(
        '--password',
        type=str,
        default=None,
        help='Password for sign-in (uses Config.TEST_PASSWORD if not provided)'
    )
    
    parser.add_argument(
        '--headless',
        action='store_true',
        help='Run browser in headless mode'
    )
    
    args = parser.parse_args()
    
    # Validate Excel file exists
    if not os.path.exists(args.excel):
        print(f"❌ Error: Excel file not found: {args.excel}")
        sys.exit(1)
    
    # Create portfolio
    create_portfolio_with_addresses(
        excel_path=args.excel,
        portfolio_name=args.portfolio_name,
        email=args.email,
        password=args.password,
        headless=args.headless
    )


if __name__ == '__main__':
    main()

