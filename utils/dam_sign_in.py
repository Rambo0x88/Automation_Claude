"""
Shared DAM sign-in function.

Usage:
    from utils.dam_sign_in import dam_sign_in

    with sync_playwright() as p:
        browser, page = dam_sign_in(p)
        # ... do work with page ...
        browser.close()

Or as a context manager:
    from utils.dam_sign_in import dam_session

    with dam_session() as page:
        # page is signed in and ready
        page.goto(...)
"""

import os
import json
import re
from playwright.sync_api import sync_playwright


def _load_credentials():
    """Load DAM credentials from tc1_account.json or Config fallback."""
    from config.config import Config

    tc1_path = os.path.join(Config.PROJECT_ROOT, "test_data", "tc1_account.json")
    if os.path.exists(tc1_path):
        with open(tc1_path) as f:
            acc = json.load(f)
        return acc["email"], acc["password"], Config.BASE_URL, Config.SIGN_IN_URL
    else:
        return Config.TEST_EMAIL, Config.TEST_PASSWORD, Config.BASE_URL, Config.SIGN_IN_URL


def dam_sign_in(playwright_instance, headless=True):
    """
    Launch browser and sign in to DAM.

    Args:
        playwright_instance: The Playwright instance from sync_playwright()
        headless: Run browser in headless mode (default True)

    Returns:
        (browser, context, page, base_url) on success
        Raises RuntimeError on sign-in failure with error details.
    """
    email, password, base_url, sign_in_url = _load_credentials()

    browser = playwright_instance.chromium.launch(
        headless=headless,
        slow_mo=300,
        channel="chrome",
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
        ],
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
        pass  # playwright_stealth not installed, continue without it

    page = context.new_page()

    # Sign in
    print(f"   🔐 Signing in as {email}...")
    page.goto(sign_in_url)
    page.wait_for_timeout(2000)

    page.fill('input[data-testid="input-email"]', email)
    page.fill('input[data-testid="input-password"]', password)
    page.click('button[data-testid="sign-in-btn"]')
    page.wait_for_timeout(8000)

    # Verify sign-in succeeded
    current_url = page.url
    if "sign-in" in current_url or "sign-up" in current_url:
        # Still on sign-in page — extract error message
        error_msg = _extract_error(page)

        print(f"   ❌ Sign-in FAILED — still on {current_url}")
        if error_msg:
            print(f"   ❌ Error: {error_msg}")
        else:
            print(f"   ❌ No error message found. Page may require CAPTCHA or credentials may be wrong.")
        print(f"   ❌ Credentials used:")
        print(f"      Email:    {email}")
        print(f"      Password: {password[:3]}{'*' * (len(password) - 3)}")

        browser.close()
        raise RuntimeError(f"DAM sign-in failed: {error_msg or 'unknown error'}")

    # Close popup if any
    try:
        for selector in ['button:has-text("×")', '[aria-label="close"]']:
            if page.locator(selector).is_visible(timeout=1000):
                page.locator(selector).first.click()
                page.wait_for_timeout(500)
                break
    except:
        pass

    # Handle redirect to specific portfolio
    if "portfolioId=" in page.url:
        page.goto(f"{base_url}/portfolio")
        page.wait_for_timeout(3000)

    print("   ✅ Signed in")
    return browser, context, page, base_url


def find_portfolio_in_dropdown(page, portfolio_name):
    """
    Open the portfolio dropdown and find a portfolio by name.

    Args:
        page: Playwright page (must be signed in and on portfolio page)
        portfolio_name: Portfolio name to search for

    Returns:
        (found, portfolio_element, display_name) tuple.
        found=True if portfolio was located.
    """
    # Open dropdown
    dropdown_opened = False
    for sel in [
        'button:has-text("Portfolio")',
        'div:has-text("Portfolio") >> button',
        '[class*="portfolio"] button',
        'button[aria-haspopup]',
    ]:
        try:
            btn = page.locator(sel).first
            if btn.count() > 0 and btn.is_visible(timeout=3000):
                btn.click()
                page.wait_for_timeout(3000)
                dropdown_opened = True
                break
        except:
            pass

    if not dropdown_opened:
        page.mouse.click(395, 141)
        page.wait_for_timeout(3000)

    # Search for portfolio
    portfolio_found = False
    portfolio_element = None
    found_name = None

    # Method 1: CSS class selectors
    for selector in [
        'div.text-mono-900.break-all',
        'div[class*="text-mono-900"][class*="break-all"]',
    ]:
        try:
            elements = page.locator(selector).all()
            for elem in elements:
                try:
                    if elem.is_visible():
                        text = elem.text_content().strip()
                        if text.lower() == portfolio_name.lower():
                            parent = elem.locator('xpath=ancestor::button').first
                            if parent.count() > 0:
                                portfolio_element = parent
                            else:
                                portfolio_element = elem
                            portfolio_found = True
                            found_name = text
                            break
                except:
                    pass
            if portfolio_found:
                break
        except:
            pass

    # Method 2: getByText
    if not portfolio_found:
        try:
            matches = page.get_by_text(portfolio_name)
            if matches.count() > 0:
                for i in range(matches.count()):
                    elem = matches.nth(i)
                    if elem.is_visible():
                        text = elem.text_content().strip()
                        first_line = text.split('\n')[0].strip()
                        if first_line.lower() == portfolio_name.lower():
                            portfolio_element = elem
                            portfolio_found = True
                            found_name = first_line
                            break
        except:
            pass

    # Method 3: Scroll and retry
    if not portfolio_found:
        try:
            for _ in range(30):
                page.keyboard.press('ArrowDown')
                page.wait_for_timeout(200)
                # Re-check after scroll
                try:
                    matches = page.get_by_text(portfolio_name)
                    if matches.count() > 0:
                        for i in range(matches.count()):
                            elem = matches.nth(i)
                            if elem.is_visible():
                                text = elem.text_content().strip()
                                first_line = text.split('\n')[0].strip()
                                if first_line.lower() == portfolio_name.lower():
                                    portfolio_element = elem
                                    portfolio_found = True
                                    found_name = first_line
                                    break
                except:
                    pass
                if portfolio_found:
                    break
        except:
            pass

    return portfolio_found, portfolio_element, found_name


def extract_addresses_from_portfolio(page):
    """
    Extract all addresses from the Combined Net Worth section of the current portfolio.

    Args:
        page: Playwright page (must be on a portfolio page)

    Returns:
        (tron_addresses, evm_addresses, exchanges) — lists of strings
    """
    tron = set()
    evm = set()
    exchanges = set()

    def classify_and_add(addr):
        addr = addr.strip()
        if addr.startswith('T') and len(addr) == 34:
            tron.add(addr)
        elif addr.startswith('0x') and len(addr) == 42:
            evm.add(addr)

    # Method 1: data-tooltip-id attributes
    try:
        elements = page.locator('[data-tooltip-id*="address-display-tooltip"]').all()
        for elem in elements:
            try:
                elem.hover()
                page.wait_for_timeout(1500)

                tooltip = page.locator('[role="tooltip"]').first
                if tooltip.count() > 0 and tooltip.is_visible():
                    text = tooltip.text_content().strip()
                    for m in re.findall(r'([Tt][A-Za-z0-9]{33})', text):
                        classify_and_add(m)
                    for m in re.findall(r'(0x[A-Fa-f0-9]{40})', text):
                        classify_and_add(m)

                tid = elem.get_attribute('data-tooltip-id') or ""
                for m in re.findall(r'address-display-tooltip-([Tt][A-Za-z0-9]{33})', tid):
                    classify_and_add(m)
                for m in re.findall(r'address-display-tooltip-(0x[A-Fa-f0-9]{40})', tid):
                    classify_and_add(m)
            except:
                pass
    except:
        pass

    # Method 2: Page source
    if not tron and not evm:
        try:
            content = page.content()
            for m in re.findall(r'address-display-tooltip-([Tt][A-Za-z0-9]{33})', content):
                classify_and_add(m)
            for m in re.findall(r'address-display-tooltip-(0x[A-Fa-f0-9]{40})', content):
                classify_and_add(m)
        except:
            pass

    # Method 3: Full page scan
    if not tron and not evm:
        try:
            content = page.content()
            for m in re.findall(r'([Tt][A-Za-z0-9]{33})', content):
                classify_and_add(m)
            for m in re.findall(r'(0x[A-Fa-f0-9]{40})', content):
                classify_and_add(m)
        except:
            pass

    return sorted(tron), sorted(evm), sorted(exchanges)


def _extract_error(page):
    """Extract error message from the current page."""
    for sel in [
        '[data-testid*="error"]', '[class*="error"]', '[class*="toast"]',
        '[class*="alert"]', '[class*="notification"]',
        '[role="alert"]', '[role="status"]',
        'div:has-text("internal server error")',
        'div:has-text("try again")', 'div:has-text("Invalid")',
    ]:
        try:
            el = page.locator(sel).first
            if el.is_visible(timeout=500):
                txt = el.text_content().strip()
                if txt and len(txt) < 200:
                    if any(kw in txt.lower() for kw in ['error', 'invalid', 'failed', 'try again', 'incorrect']):
                        return txt
        except:
            pass

    # Fallback: scan body text
    try:
        body = page.locator("body").text_content().strip()
        for line in body.split('\n'):
            line = line.strip()
            if line and 5 < len(line) < 200:
                if any(kw in line.lower() for kw in ['error', 'invalid', 'incorrect', 'failed', 'try again']):
                    return line
    except:
        pass

    return ""


class dam_session:
    """
    Context manager for a signed-in DAM browser session.

    Usage:
        with dam_session() as (page, base_url):
            page.goto(f"{base_url}/portfolio")
            # ... do work ...
    """

    def __init__(self, headless=True):
        self.headless = headless
        self._pw = None
        self._browser = None

    def __enter__(self):
        self._pw = sync_playwright().start()
        self._browser, self._context, self._page, self._base_url = dam_sign_in(
            self._pw, headless=self.headless
        )
        return self._page, self._base_url

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._browser:
            self._browser.close()
        if self._pw:
            self._pw.stop()
        return False
