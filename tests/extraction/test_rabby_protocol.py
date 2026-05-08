"""
Rabby Protocol & Hyperliquid API test with DAM screen capture.

Tests:
  test_dam_screenshot  — Sign in to DAM, navigate to portfolio, take screenshot
  test_rabby_export    — Fetch Rabby Protocol + Hyperliquid, export to Excel (no browser)

Run all:
  pytest tests/test_rabby_protocol.py -v -s

Run Rabby export only (no browser):
  pytest tests/test_rabby_protocol.py::test_rabby_export -v -s

Run DAM screenshot only:
  pytest tests/test_rabby_protocol.py::test_dam_screenshot -v -s --headed

Test data loaded from test_data/tc_dune_wallet.json:
  - wallet_address: EVM address to query
  - portfolio_name: DAM portfolio name to navigate to (leave blank to screenshot portfolio list)
"""

import json
import os
import re
import sys
from datetime import datetime

import pytest
from playwright.sync_api import Page

# Ensure project root is on path when running directly
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config.config import Config
from pages.sign_in_page import SignInPage
from pages.portfolio_page import PortfolioPage
from utils.rabby_api import export_combined_excel

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

_TEST_DATA_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "test_data", "tc_dune_wallet.json",
)
_OUTPUT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "test-results", "excel-exports",
)
_SCREENSHOT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "test-results", "screenshots",
)
_PORTFOLIO_NAME = "test_portfolio"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _load_test_data() -> dict:
    """Load test data from test_data/tc_dune_wallet.json."""
    try:
        with open(_TEST_DATA_PATH) as f:
            return json.load(f)
    except Exception as e:
        print(f"⚠️  Could not load tc_dune_wallet.json: {e}")
        return {}


def _load_addresses() -> list:
    """Return EVM address list from test data."""
    data = _load_test_data()
    addr = data.get("wallet_address", "")
    if addr:
        return [addr]
    return ["0xeb2Eb5C68156250C368914761bB8F1208d56AcD0"]


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def authenticated_user(page: Page, config):
    """Sign in to DAM with the pre-created test account."""
    sign_in_page = SignInPage(page)
    sign_in_page.navigate()
    sign_in_page.sign_in(Config.TEST_EMAIL, Config.TEST_PASSWORD)
    assert sign_in_page.wait_for_successful_sign_in(), "DAM sign-in failed"
    return {"email": Config.TEST_EMAIL}


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_dam_screenshot(page: Page, authenticated_user):
    """
    Sign in to DAM, navigate to the configured portfolio, and take a
    full-page screenshot saved to test-results/screenshots/.
    """
    data = _load_test_data()
    portfolio_name = data.get("portfolio_name", "").strip()
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    print(f"\n{'='*70}", flush=True)
    print(f"📸 DAM Portfolio Screen Capture", flush=True)
    print(f"{'='*70}", flush=True)

    portfolio_page = PortfolioPage(page)
    portfolio_page.navigate()
    page.wait_for_load_state("load")
    page.wait_for_timeout(3000)

    if portfolio_name:
        try:
            print(f"   🔍 Navigating to portfolio: {portfolio_name}", flush=True)
            page.locator(f"text={portfolio_name}").first.click()
            page.wait_for_load_state("load")
            page.wait_for_timeout(8000)  # Wait for protocol positions to load
            print(f"   ✅ Portfolio loaded", flush=True)
        except Exception as e:
            print(f"   ⚠️  Could not navigate to '{portfolio_name}': {e}", flush=True)
            print(f"   ℹ️  Screenshotting portfolio list instead", flush=True)
    else:
        print(f"   ℹ️  No portfolio_name set — screenshotting portfolio list", flush=True)

    os.makedirs(_SCREENSHOT_DIR, exist_ok=True)
    safe_name = re.sub(r'[^A-Za-z0-9_\-]', '_', portfolio_name or "portfolio_list")
    screenshot_path = os.path.join(_SCREENSHOT_DIR, f"DAM_Protocol_{safe_name}_{timestamp}.png")
    page.screenshot(path=screenshot_path, full_page=True)

    assert os.path.exists(screenshot_path), f"Screenshot not created: {screenshot_path}"
    print(f"\n✅ Screenshot saved: {screenshot_path}", flush=True)


def test_rabby_export():
    """
    Fetch Rabby Protocol (Aave, Morpho, Compound) + Hyperliquid data and
    export to a single Excel file. No browser required.
    """
    addresses = _load_addresses()
    data = _load_test_data()
    portfolio_name = data.get("portfolio_name", "").strip() or _PORTFOLIO_NAME

    print(f"\n{'='*70}", flush=True)
    print(f"TEST: Rabby Export (Protocol + Hyperliquid)", flush=True)
    print(f"Addresses: {addresses}", flush=True)
    print(f"{'='*70}", flush=True)

    filepath = export_combined_excel(
        addresses=addresses,
        portfolio_name=portfolio_name,
        output_dir=_OUTPUT_DIR,
    )

    assert os.path.exists(filepath), f"Excel file not created: {filepath}"
    print(f"\n✅ Excel saved: {filepath}", flush=True)
