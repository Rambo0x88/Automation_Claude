#!/usr/bin/env python3
"""
Create Portfolio with Exchange Accounts

Flow:
1. Sign in to DAM with existing account
2. Create new portfolio with given name
3. Add wallet addresses (optional)
4. Connect exchange accounts (Binance)
5. Save portfolio

Usage:
  python3 run_create_portfolio.py "my portfolio name"
  python3 run_create_portfolio.py "my portfolio" --address 0x4e14fc11...
  python3 run_create_portfolio.py "my portfolio" --address 0x4e14... --address TUqEg3...
  python3 run_create_portfolio.py "my portfolio" --exchange binance
  python3 run_create_portfolio.py "my portfolio" --exchange binance --exchange-key "moon api key"
  python3 run_create_portfolio.py "my portfolio" --address 0x4e14... --exchange binance
"""

import os
import sys
import json
import argparse
from playwright.sync_api import sync_playwright

# Add project root to path
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

from config.config import Config
from utils.exchange_connector import add_exchanges_to_portfolio, connect_exchange_from_config


def load_credentials():
    """Load credentials from tc1_account.json or .env"""
    tc1_path = os.path.join(Config.PROJECT_ROOT, "test_data", "tc1_account.json")
    if os.path.exists(tc1_path):
        with open(tc1_path) as f:
            acc = json.load(f)
        return acc["email"], acc["password"]
    return Config.TEST_EMAIL, Config.TEST_PASSWORD


def run(portfolio_name, addresses=None, exchange=None, exchange_key=None, headless=False):
    """
    Create portfolio with optional addresses and exchange connections.
    """
    email, password = load_credentials()
    addresses = addresses or []

    print(f"\n{'='*60}")
    print(f"Create Portfolio: {portfolio_name}")
    print(f"Account: {email}")
    if addresses:
        for i, addr in enumerate(addresses, 1):
            print(f"Address {i}: {addr}")
    if exchange:
        print(f"Exchange: {exchange}")
        if exchange_key:
            print(f"Key: {exchange_key}")
        else:
            print(f"Keys: all configured")
    print(f"{'='*60}\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            slow_mo=300,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        try:
            from playwright_stealth import Stealth
            Stealth(navigator_platform_override="MacIntel").apply_stealth_sync(context)
        except ImportError:
            pass

        page = context.new_page()
        page.set_default_timeout(30000)

        # Step 1: Sign in
        print("🔐 Signing in...")
        page.goto(Config.SIGN_IN_URL)
        page.wait_for_timeout(2000)
        page.fill('input[data-testid="input-email"]', email)
        page.fill('input[data-testid="input-password"]', password)
        page.click('button[data-testid="sign-in-btn"]')
        page.wait_for_timeout(8000)

        if "sign-in" in page.url:
            print(f"❌ Sign-in failed. Still on {page.url}")
            print(f"   Email: {email}")
            browser.close()
            sys.exit(1)

        print("✅ Signed in")

        # Step 2: Navigate to portfolio and open Create Portfolio
        print("\n📂 Opening Create Portfolio...")

        # Go to portfolio page
        if "portfolioId=" in page.url:
            page.goto(f"{Config.BASE_URL}/portfolio")
            page.wait_for_timeout(3000)

        # Open dropdown
        for sel in [
            'button:has-text("Portfolio")',
            'button[aria-haspopup]',
        ]:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=3000):
                    btn.click()
                    page.wait_for_timeout(2000)
                    break
            except:
                pass

        # Scroll to bottom of dropdown to find Create portfolio
        page.mouse.move(490, 450)
        for _ in range(15):
            page.mouse.wheel(0, 300)
            page.wait_for_timeout(100)
        page.wait_for_timeout(500)

        # Click Create portfolio
        create_clicked = False
        for sel_fn in [
            lambda: page.get_by_role("menuitem", name="Create portfolio", exact=True),
            lambda: page.get_by_text("Create portfolio", exact=True),
            lambda: page.get_by_text("Create portfolio"),
        ]:
            try:
                btn = sel_fn()
                if btn.count() > 0 and btn.first.is_visible():
                    btn.first.click()
                    create_clicked = True
                    break
            except:
                pass

        if not create_clicked:
            print("❌ Could not find Create portfolio button")
            browser.close()
            sys.exit(1)

        page.wait_for_timeout(2000)
        print("✅ Create Portfolio page opened")

        # Step 3: Enter portfolio name
        print(f"\n📝 Portfolio name: {portfolio_name}")
        name_input = page.locator('input[placeholder*="portfolio name" i], input[placeholder*="Enter portfolio" i]').first
        name_input.wait_for(state="visible", timeout=5000)
        name_input.fill(portfolio_name)
        page.wait_for_timeout(500)

        # Step 4: Add addresses (if any)
        if addresses:
            print(f"\n📍 Adding {len(addresses)} address(es)...")
            # Make sure we're on Address tab
            for sel in ['text="Address(0)"', 'text="Address (0)"', '[role="tab"]:has-text("Address")']:
                try:
                    tab = page.locator(sel).first
                    if tab.is_visible(timeout=2000):
                        tab.click()
                        page.wait_for_timeout(1000)
                        break
                except:
                    pass

            for addr in addresses:
                addr_input = page.locator('input[data-testid*="input-wallet"], input[placeholder*="wallet address" i], input[placeholder*="Paste" i]').last
                addr_input.wait_for(state="visible", timeout=5000)
                addr_input.fill(addr)
                page.wait_for_timeout(500)
                addr_input.press("Enter")
                page.wait_for_timeout(2000)
                print(f"   ✅ Added: {addr[:12]}...{addr[-8:]}")

        # Step 5: Connect exchange (if requested)
        if exchange:
            print(f"\n🔗 Connecting {exchange} exchange...")
            if exchange_key:
                connect_exchange_from_config(page, exchange=exchange, key_name=exchange_key)
            else:
                add_exchanges_to_portfolio(page, exchanges=[exchange])

        # Step 6: Save portfolio
        print("\n💾 Saving portfolio...")
        save_btn = page.locator('button:has-text("Save")').first
        try:
            save_btn.wait_for(state="visible", timeout=10000)
            page.wait_for_timeout(2000)
            save_btn.click(timeout=15000)
            page.wait_for_timeout(5000)
            print("✅ Portfolio saved!")
        except Exception as e:
            print(f"⚠️  Save button issue: {e}")
            try:
                page.locator('button[type="submit"]').first.click(timeout=10000)
                page.wait_for_timeout(5000)
                print("✅ Portfolio saved (submit fallback)")
            except:
                print("❌ Could not save portfolio")

        print(f"\n{'='*60}")
        print(f"✅ Done — Portfolio '{portfolio_name}' created")
        print(f"   URL: {page.url}")
        print(f"{'='*60}\n")

        browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Create DAM portfolio with exchange accounts",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 run_create_portfolio.py "my portfolio"
  python3 run_create_portfolio.py "my portfolio" --address 0x4e14fc11...
  python3 run_create_portfolio.py "my portfolio" --exchange binance
  python3 run_create_portfolio.py "my portfolio" --exchange binance --exchange-key "moon api key"
  python3 run_create_portfolio.py "my portfolio" --address 0x4e14... --exchange binance
        """
    )
    parser.add_argument('name', help='Portfolio name')
    parser.add_argument('--address', action='append', default=None,
                        help='Wallet address to add (repeat for multiple)')
    parser.add_argument('--exchange', type=str, default=None,
                        help='Exchange type to connect (binance or bit)')
    parser.add_argument('--exchange-key', type=str, default=None,
                        help='Specific exchange key name from config')
    parser.add_argument('--headless', action='store_true', default=False,
                        help='Run in headless mode')

    args = parser.parse_args()

    run(
        portfolio_name=args.name,
        addresses=args.address,
        exchange=args.exchange,
        exchange_key=args.exchange_key,
        headless=args.headless,
    )
