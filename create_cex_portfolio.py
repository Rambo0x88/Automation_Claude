#!/usr/bin/env python3
"""
Create a CEX-only portfolio and connect exchange accounts.

Usage:
    cd core/projects/DAM/automationv2
    source venv/bin/activate

    # Both binance + bit
    TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py -e binance bit

    # With custom portfolio name
    TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py --name "My CEX Portfolio" -e binance bit

    # With specific exchange key
    TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py -e bit --key "david"
"""
import sys
import os
import argparse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from playwright.sync_api import sync_playwright
from config.config import Config
from utils.exchange_connector import (
    click_exchange_tab,
    connect_exchange_from_config,
)
from utils.helpers import make_screenshot_folder

# Will be set in main() after parsing email
SCREENSHOT_DIR = None


def ss(page, name):
    """Save screenshot to the run's screenshot folder"""
    path = os.path.join(SCREENSHOT_DIR, f"{name}.png")
    page.screenshot(path=path)
    print(f"   📸 {path}")


def main():
    parser = argparse.ArgumentParser(description="Create CEX-only portfolio with exchange connection")
    parser.add_argument("--name", "-n", default=None, help="Portfolio name (default: auto-generated)")
    parser.add_argument("--exchange", "-e", nargs="+", default=["binance"], help="Exchange type(s): binance, bit, or both")
    parser.add_argument("--key", "-k", default=None, help="Specific exchange key display_name (default: first available)")
    args = parser.parse_args()

    test_email = os.getenv("TEST_EMAIL", Config.TEST_EMAIL)
    test_password = os.getenv("TEST_PASSWORD", Config.TEST_PASSWORD)

    if not test_email:
        print("❌ TEST_EMAIL not set.")
        sys.exit(1)

    global SCREENSHOT_DIR
    SCREENSHOT_DIR = make_screenshot_folder("tc1b-2", test_email)

    portfolio_name = args.name or f"CEX_{datetime.now().strftime('%m%d_%H%M')}"

    print("=" * 80)
    print("CREATE CEX-ONLY PORTFOLIO + EXCHANGE CONNECTION")
    print("=" * 80)
    print(f"   Email:     {test_email}")
    print(f"   Portfolio: {portfolio_name}")
    print(f"   Exchange:  {', '.join(args.exchange)}")
    print(f"   Screenshots: {SCREENSHOT_DIR}")
    print("=" * 80)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=250)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()

        # ── STEP 1: Sign in ──────────────────────────────────────────
        print("\n📋 STEP 1: Sign In")
        print("-" * 40)
        page.goto(Config.SIGN_IN_URL)
        page.wait_for_timeout(2000)

        email_input = page.locator('input[type="email"], input[name="email"], input[placeholder*="email" i]').first
        email_input.fill(test_email)
        page.wait_for_timeout(500)

        pwd_input = page.locator('input[type="password"]').first
        pwd_input.fill(test_password)
        page.wait_for_timeout(500)

        sign_in_btn = page.locator('button[type="submit"], button:has-text("Sign In"), button:has-text("Log In")').first
        sign_in_btn.click()
        page.wait_for_timeout(5000)

        if "sign-in" in page.url.lower() or "login" in page.url.lower():
            print("❌ Sign in failed — check credentials")
            ss(page, "01_signin_failed")
            browser.close()
            sys.exit(1)
        print("✅ Signed in successfully")
        ss(page, "01_signed_in")

        # ── STEP 2: Navigate to Create Portfolio page ─────────────────
        print("\n📋 STEP 2: Navigate to Create Portfolio")
        print("-" * 40)

        page.goto(Config.PORTFOLIO_URL)
        page.wait_for_timeout(3000)
        print(f"✅ On portfolio page ({page.url})")

        # Check if this is a new account (no portfolios) — shows "Create portfolio" button in center
        on_create_page = False
        new_account = page.locator('text="Start Managing Your Digital Assets"').first
        try:
            if new_account.is_visible(timeout=3000):
                print("   ✅ New account detected (no portfolios)")
                create_btn = page.locator('button:has-text("Create portfolio")').first
                create_btn.click()
                page.wait_for_timeout(5000)
                on_create_page = True
                print(f"   ✅ Clicked Create portfolio → {page.url}")
        except:
            pass

        # If we're now on the create page, don't do anything else
        if not on_create_page:
            # Check if create form is already visible
            try:
                for sel in ['form fieldset input', 'input[placeholder*="portfolio" i]']:
                    if page.locator(sel).first.is_visible(timeout=2000):
                        on_create_page = True
                        print("   ✅ Already on Create Portfolio page")
                        break
            except:
                pass

        # Existing account with portfolios — use dropdown
        if not on_create_page:
            # Use dropdown to navigate to Create Portfolio
            for sel in ['button:has-text("Portfolio")', '[class*="portfolio"] button']:
                try:
                    btn = page.locator(sel).first
                    if btn.is_visible(timeout=3000):
                        btn.click()
                        page.wait_for_timeout(2000)
                        print(f"   ✅ Dropdown opened ({sel})")
                        break
                except:
                    continue

            # Scroll dropdown to find Create portfolio
            page.mouse.move(490, 450)
            for _ in range(15):
                page.mouse.wheel(0, 200)
                page.wait_for_timeout(200)

            # Click Create portfolio
            for sel in ['text="Create portfolio"', '[role="menuitem"]:has-text("Create portfolio")', 'button:has-text("Create portfolio")']:
                try:
                    btn = page.locator(sel).first
                    if btn.is_visible(timeout=3000):
                        btn.click()
                        print(f"   ✅ Clicked Create portfolio (dropdown)")
                        break
                except:
                    continue

            # Wait for Create Portfolio page to load
            page.wait_for_timeout(5000)

        # Scroll to top
        page.evaluate("window.scrollTo(0, 0)")
        page.wait_for_timeout(1000)
        ss(page, "02_create_portfolio_page")

        # ── STEP 3: Enter portfolio name ──────────────────────────────
        print("\n📋 STEP 3: Enter Portfolio Name")
        print("-" * 40)

        name_input = None
        for selector in [
            'form fieldset input',
            'input[placeholder*="portfolio" i]',
            'input[placeholder*="Enter portfolio" i]',
            'xpath=/html/body/div/div/main/div/div/div/div/div/div/form/fieldset/div[1]/div[1]/input',
            'input[type="text"]',
        ]:
            try:
                candidate = page.locator(selector).first
                if candidate.count() > 0 and candidate.is_visible(timeout=2000):
                    name_input = candidate
                    print(f"   Found input: '{selector}'")
                    break
            except:
                continue

        if name_input:
            name_input.click()
            name_input.fill(portfolio_name)
            page.wait_for_timeout(500)
            print(f"✅ Portfolio name entered: {portfolio_name}")
        else:
            print("❌ Could not find portfolio name input")
            all_inputs = page.locator('input').all()
            print(f"   DEBUG: {len(all_inputs)} inputs on page:")
            for inp in all_inputs:
                try:
                    print(f"   - visible={inp.is_visible()} placeholder='{inp.get_attribute('placeholder')}' name='{inp.get_attribute('name')}'")
                except:
                    pass
            print(f"   DEBUG: URL: {page.url}")
            ss(page, "03_no_name_input")
            browser.close()
            sys.exit(1)

        ss(page, "03_name_entered")

        # ── STEP 4: Switch to Exchange tab and connect ────────────────
        print("\n📋 STEP 4: Connect Exchange")
        print("-" * 40)

        try:
            click_exchange_tab(page)
            page.wait_for_timeout(1000)
            ss(page, "04_exchange_tab")

            for exchange in args.exchange:
                print(f"   🔄 Connecting {exchange}...")
                if args.key:
                    connect_exchange_from_config(page, exchange=exchange, key_name=args.key)
                else:
                    connect_exchange_from_config(page, exchange=exchange)
                page.wait_for_timeout(2000)
                ss(page, f"04_{exchange}_connected")
                print(f"   ✅ {exchange} connected")

            print("✅ All exchanges connected")
        except Exception as e:
            print(f"⚠️  Exchange connection error: {e}")
            ss(page, "04_exchange_error")

        # ── STEP 4b: Tick all exchange account checkboxes ─────────────
        print("\n📋 STEP 4b: Select All Exchange Accounts")
        print("-" * 40)
        try:
            page.wait_for_timeout(1000)
            # The checkboxes are inside table rows: label > input[type="checkbox"]
            # They may already be checked after connecting — verify and tick unchecked ones
            checkboxes = page.locator('tbody input[type="checkbox"]').all()
            ticked = 0
            already_checked = 0
            for cb in checkboxes:
                try:
                    if cb.is_visible():
                        if cb.is_checked():
                            already_checked += 1
                        else:
                            # Click the parent label to toggle
                            cb.locator('xpath=..').click()
                            page.wait_for_timeout(300)
                            ticked += 1
                except:
                    continue
            print(f"   ✅ {already_checked} already checked, {ticked} newly ticked")
            if ticked == 0 and already_checked == 0:
                # Fallback: click all labels with checkbox styling
                labels = page.locator('tbody label:has(input[type="checkbox"])').all()
                for lbl in labels:
                    try:
                        if lbl.is_visible():
                            lbl.click()
                            page.wait_for_timeout(300)
                            ticked += 1
                    except:
                        continue
                print(f"   ✅ Fallback: ticked {ticked} label(s)")
            ss(page, "04b_checkboxes_ticked")
        except Exception as e:
            print(f"   ⚠️  Checkbox error: {e}")
            ss(page, "04b_checkbox_error")

        # ── STEP 5: Save portfolio ────────────────────────────────────
        print("\n📋 STEP 5: Save Portfolio")
        print("-" * 40)

        save_button = page.locator('button:has-text("Save")').first
        try:
            save_button.wait_for(state="visible", timeout=10000)
            page.wait_for_timeout(2000)
            save_button.click(timeout=15000)
            page.wait_for_timeout(5000)
            print("✅ Portfolio saved!")
        except Exception as e:
            print(f"⚠️  Save issue: {e}")
            try:
                page.locator('button[type="submit"]').first.click(timeout=10000)
                page.wait_for_timeout(5000)
                print("✅ Portfolio saved (submit fallback)!")
            except:
                print("❌ Could not save portfolio")

        ss(page, "05_portfolio_saved")

        print("\n" + "=" * 80)
        print(f"✅ DONE — Portfolio '{portfolio_name}' created with {', '.join(args.exchange)} exchange(s)")
        print("=" * 80)

        page.wait_for_timeout(5000)
        browser.close()


if __name__ == "__main__":
    main()
