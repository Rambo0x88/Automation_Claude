"""
DAM Authentication — Shared sign-in, credentials loading, and browser setup.

Reusable across all scripts that need to sign into DAM:
  - run_overview.py
  - utils/trx_transaction/trongrid_dam_comparison.py
  - utils/trx_transaction/dam_transaction_extractor.py
  - utils/trx_transaction/dam_transaction_extractor_v2.py

Usage:
    from utils.dam_auth import load_credentials, sign_in_to_dam, launch_browser

    email, password = load_credentials()
    browser, context, page = launch_browser()
    sign_in_to_dam(page, email, password)
    # ... do work ...
    browser.close()
"""

import os
import json
from config.config import Config


def load_credentials(tc1_path=None):
    """
    Load DAM credentials from tc1_account.json or fall back to .env/Config.

    Args:
        tc1_path: Absolute path to tc1_account.json (defaults to PROJECT_ROOT/test_data/tc1_account.json)

    Returns:
        tuple: (email, password)
    """
    if tc1_path is None:
        tc1_path = os.path.join(Config.PROJECT_ROOT, "test_data", "tc1_account.json")
    if os.path.exists(tc1_path):
        with open(tc1_path) as f:
            acc = json.load(f)
        email = acc["email"]
        password = acc["password"]
        print(f"   📂 Loaded credentials from {tc1_path}: {email}")
        return email, password

    email = Config.TEST_EMAIL
    password = Config.TEST_PASSWORD
    if email:
        print(f"   📂 Using .env credentials: {email}")
    else:
        print(f"   ⚠️  No credentials found in {tc1_path} or .env")
    return email, password


def launch_browser(headless=None, slow_mo=400, use_stealth=True):
    """
    Launch a Playwright Chromium browser with Chrome stealth settings.

    Args:
        headless: True/False/None (None = use HEADLESS env var, default True)
        slow_mo: Milliseconds delay between actions
        use_stealth: Apply playwright-stealth to bypass bot detection

    Returns:
        tuple: (browser, context, page)
    """
    from playwright.sync_api import sync_playwright

    if headless is None:
        headless = os.environ.get('HEADLESS', 'true').lower() == 'true'

    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=headless,
        slow_mo=slow_mo,
        channel="chrome",
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
        ],
    )
    context = browser.new_context(
        viewport={"width": 1920, "height": 1080},
        user_agent=(
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )

    if use_stealth:
        try:
            from playwright_stealth import Stealth
            Stealth(navigator_platform_override="MacIntel").apply_stealth_sync(context)
        except ImportError:
            print("   ⚠️  playwright-stealth not installed, skipping stealth")

    page = context.new_page()
    return browser, context, page


def sign_in_to_dam(page, email, password, base_url=None, max_attempts=3):
    """
    Sign in to DAM with retry mechanism.

    Args:
        page: Playwright page object
        email: DAM account email
        password: DAM account password
        base_url: DAM base URL (defaults to Config.BASE_URL)
        max_attempts: Maximum sign-in attempts

    Returns:
        bool: True if sign-in succeeded, False otherwise
    """
    if base_url is None:
        base_url = Config.BASE_URL

    sign_in_url = f"{base_url}/sign-in"

    for attempt in range(1, max_attempts + 1):
        if attempt > 1:
            print(f"\n   🔄 Sign-in retry attempt {attempt}/{max_attempts}...")

        page.goto(sign_in_url)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(2000)

        page.fill('input[data-testid="input-email"]', email)
        page.fill('input[data-testid="input-password"]', password)
        page.click('button[data-testid="sign-in-btn"]')
        print(f"   🔐 Signing in as {email}...")

        # Wait for redirect
        page.wait_for_timeout(8000)

        # Close any popup
        try:
            for selector in ['button:has-text("×")', '[aria-label="close"]']:
                if page.locator(selector).is_visible(timeout=1000):
                    page.locator(selector).first.click()
                    page.wait_for_timeout(1000)
                    break
        except Exception:
            pass

        # Check if sign-in succeeded
        current_url = page.url
        if '/sign-in' not in current_url:
            print(f"   ✅ Signed in successfully")
            return True
        else:
            print(f"   ⚠️  Still on sign-in page after attempt {attempt}")
            if attempt < max_attempts:
                page.wait_for_timeout(5000)

    print(f"   ❌ Sign-in failed after {max_attempts} attempts")
    return False
