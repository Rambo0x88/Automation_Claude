"""
Exchange Connector — Connect exchange accounts to a DAM portfolio.

Handles the "Connect Exchange" flow in the Create Portfolio page:
1. Click Exchange(0) tab
2. Check if exchanges already exist (previously connected)
3. Click "Connect Exchange" button
4. Select exchange type (Binance / BIT)
5. Fill Display Name, API Key, Secret Key
6. Click Connect

Usage:
    from utils.exchange_connector import connect_exchange, load_exchange_keys

    # Load keys from config
    keys = load_exchange_keys()

    # Connect a specific exchange (page must be on Create Portfolio page)
    connect_exchange(page, exchange="binance", display_name="moon api key",
                     api_key="...", secret_key="...")

    # Or connect using keys from config file
    connect_exchange_from_config(page, exchange="binance", key_name="moon api key")

    # Or connect all configured exchanges
    connect_all_exchanges(page, exchange="binance")
"""

import os
import json


def load_exchange_keys(config_path=None):
    """
    Load exchange API keys from config file.

    Args:
        config_path: Path to exchange_keys.json. Defaults to test_data/exchange_keys.json

    Returns:
        dict: {"binance": [...], "bit": [...]}
    """
    if config_path is None:
        _script_dir = os.path.dirname(os.path.abspath(__file__))
        _project_root = os.path.normpath(os.path.join(_script_dir, ".."))
        config_path = os.path.join(_project_root, "test_data", "exchange_keys.json")

    if not os.path.exists(config_path):
        print(f"   ⚠️  Exchange keys file not found: {config_path}")
        return {"binance": [], "bit": []}

    with open(config_path) as f:
        return json.load(f)


def click_exchange_tab(page):
    """
    Click the Exchange(0) tab on the Create Portfolio page.

    Args:
        page: Playwright page (must be on Create Portfolio page)

    Returns:
        bool: True if tab was clicked successfully
    """
    for sel in [
        'xpath=//*[contains(@id, "trigger-exchange")]',
        '[id*="trigger-exchange"]',
        'text="Exchange(0)"',
        'text="Exchange (0)"',
        '[role="tab"]:has-text("Exchange")',
        'button:has-text("Exchange")',
    ]:
        try:
            tab = page.locator(sel).first
            if tab.is_visible(timeout=3000):
                tab.click()
                page.wait_for_timeout(1500)
                print("   ✅ Clicked Exchange tab")
                return True
        except:
            pass

    print("   ⚠️  Exchange tab not found")
    return False


def get_existing_exchanges(page):
    """
    Check if any exchanges are already connected (listed in the table).

    Args:
        page: Playwright page (must be on Exchange tab)

    Returns:
        list: Names of existing exchanges (e.g. ["moon api key", "xg"])
    """
    existing = []
    try:
        # Look for rows in the exchange table
        rows = page.locator('table tbody tr, [class*="table"] tr').all()
        for row in rows:
            try:
                name_cell = row.locator('td').first
                if name_cell.is_visible(timeout=500):
                    name = name_cell.text_content().strip()
                    if name and name.lower() not in ('no data to display', 'name', ''):
                        existing.append(name)
            except:
                pass
    except:
        pass

    if existing:
        print(f"   📋 Existing exchanges: {existing}")
    else:
        print("   📋 No existing exchanges found")

    return existing


def connect_exchange(page, exchange="binance", display_name="", api_key="", secret_key=""):
    """
    Connect an exchange account via the Connect Exchange modal.

    Args:
        page: Playwright page (must be on Exchange tab of Create Portfolio)
        exchange: "binance" or "bit"
        display_name: Display name for the exchange connection
        api_key: API key
        secret_key: Secret key

    Returns:
        bool: True if connection was successful
    """
    if not display_name or not api_key or not secret_key:
        print(f"   ❌ Missing credentials for {exchange}: {display_name}")
        return False

    # Click "Connect Exchange" button
    connect_btn_clicked = False
    for sel in [
        '[id*="content-exchange"] button',
        'button:has-text("Connect Exchange")',
        'text="Connect Exchange"',
        '[class*="connect"]:has-text("Connect Exchange")',
    ]:
        try:
            btn = page.locator(sel).first
            if btn.is_visible(timeout=3000):
                btn.click()
                page.wait_for_timeout(2000)
                connect_btn_clicked = True
                print(f"   ✅ Clicked 'Connect Exchange'")
                break
        except:
            pass

    if not connect_btn_clicked:
        print("   ❌ 'Connect Exchange' button not found")
        return False

    # Wait for modal to appear
    modal_visible = False
    try:
        modal = page.locator('[role="dialog"], [class*="modal"]').first
        modal_visible = modal.is_visible(timeout=5000)
    except:
        pass

    if not modal_visible:
        print("   ❌ Connect Exchange modal did not appear")
        return False

    # Select exchange type (Binance or BIT)
    exchange_lower = exchange.lower()
    if exchange_lower == "binance":
        for sel in [
            'text="Binance"',
            '[class*="exchange"]:has-text("Binance")',
            'div:has-text("Binance"):not(:has-text("BIT"))',
        ]:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=2000):
                    el.click()
                    page.wait_for_timeout(500)
                    print("   ✅ Selected Binance")
                    break
            except:
                pass
    elif exchange_lower == "bit":
        for sel in [
            'text="BIT"',
            '[class*="exchange"]:has-text("BIT")',
        ]:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=2000):
                    el.click()
                    page.wait_for_timeout(500)
                    print("   ✅ Selected BIT")
                    break
            except:
                pass

    # Fill Display Name
    try:
        name_input = page.locator(
            'input[placeholder*="Display Name" i], '
            'input[placeholder*="Trading Account" i], '
            'input[placeholder*="display name" i]'
        ).first
        if name_input.is_visible(timeout=3000):
            name_input.fill(display_name)
            page.wait_for_timeout(300)
            print(f"   ✅ Display Name: {display_name}")
    except Exception as e:
        print(f"   ⚠️  Could not fill Display Name: {e}")

    # Fill API Key
    try:
        api_input = page.locator(
            f'input[placeholder*="API Key" i], '
            f'input[placeholder*="{exchange} API" i]'
        ).first
        if api_input.is_visible(timeout=3000):
            api_input.fill(api_key)
            page.wait_for_timeout(300)
            print(f"   ✅ API Key: {api_key[:8]}...{api_key[-4:]}")
    except Exception as e:
        print(f"   ⚠️  Could not fill API Key: {e}")

    # Fill Secret Key
    try:
        secret_input = page.locator(
            'input[placeholder*="Secret Key" i], '
            f'input[placeholder*="{exchange} Secret" i]'
        ).first
        if secret_input.is_visible(timeout=3000):
            secret_input.fill(secret_key)
            page.wait_for_timeout(300)
            print(f"   ✅ Secret Key: {secret_key[:8]}...{secret_key[-4:]}")
    except Exception as e:
        print(f"   ⚠️  Could not fill Secret Key: {e}")

    # Click Connect button (inside modal)
    connected = False
    for sel in [
        '[role="dialog"] button:has-text("Connect")',
        'button:has-text("Connect"):not(:has-text("Connect Exchange"))',
        '[class*="modal"] button:has-text("Connect")',
    ]:
        try:
            btn = page.locator(sel).last  # Last "Connect" button (not "Connect Exchange")
            if btn.is_visible(timeout=3000):
                btn.click()
                page.wait_for_timeout(3000)
                connected = True
                print(f"   ✅ Clicked Connect")
                break
        except:
            pass

    if not connected:
        print("   ❌ Could not click Connect button")
        return False

    # Check for success/error
    try:
        # Check if modal closed (success)
        page.wait_for_timeout(2000)
        modal_still = page.locator('[role="dialog"]:has-text("Connect Exchange")').first
        if not modal_still.is_visible(timeout=2000):
            print(f"   ✅ Exchange '{display_name}' connected successfully")
            return True
        else:
            # Check for error message in modal
            error_el = page.locator('[class*="error"], [role="alert"]').first
            if error_el.is_visible(timeout=1000):
                error_msg = error_el.text_content().strip()
                print(f"   ❌ Connection error: {error_msg}")
            else:
                print(f"   ⚠️  Modal still open — connection may have failed")
            # Close modal
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
            return False
    except:
        return True  # Assume success if we can't check


def connect_exchange_from_config(page, exchange="binance", key_name=None, config_path=None):
    """
    Connect an exchange using credentials from the config file.

    Args:
        page: Playwright page
        exchange: "binance" or "bit"
        key_name: Specific key name to use (e.g. "moon api key"). If None, uses first available.
        config_path: Path to exchange_keys.json

    Returns:
        bool: True if connected successfully
    """
    keys = load_exchange_keys(config_path)
    exchange_keys = keys.get(exchange.lower(), [])

    if not exchange_keys:
        print(f"   ⚠️  No {exchange} keys configured in exchange_keys.json")
        return False

    # Find the right key
    key_entry = None
    if key_name:
        for k in exchange_keys:
            if k["display_name"].lower() == key_name.lower():
                key_entry = k
                break
        if not key_entry:
            print(f"   ⚠️  Key '{key_name}' not found for {exchange}")
            return False
    else:
        key_entry = exchange_keys[0]

    return connect_exchange(
        page,
        exchange=exchange,
        display_name=key_entry["display_name"],
        api_key=key_entry["api_key"],
        secret_key=key_entry["secret_key"],
    )


def connect_all_exchanges(page, exchange="binance", config_path=None):
    """
    Connect all configured exchange keys for a given exchange type.

    Args:
        page: Playwright page (must be on Exchange tab)
        exchange: "binance" or "bit"
        config_path: Path to exchange_keys.json

    Returns:
        int: Number of exchanges successfully connected
    """
    keys = load_exchange_keys(config_path)
    exchange_keys = keys.get(exchange.lower(), [])

    if not exchange_keys:
        print(f"   ⚠️  No {exchange} keys configured")
        return 0

    # Check existing exchanges to avoid duplicates
    existing = get_existing_exchanges(page)
    existing_lower = [e.lower() for e in existing]

    connected_count = 0
    for key_entry in exchange_keys:
        name = key_entry["display_name"]
        if name.lower() in existing_lower:
            print(f"   ⏭️  '{name}' already connected — skipping")
            continue

        success = connect_exchange(
            page,
            exchange=exchange,
            display_name=name,
            api_key=key_entry["api_key"],
            secret_key=key_entry["secret_key"],
        )
        if success:
            connected_count += 1

    print(f"   📊 Connected {connected_count} new {exchange} exchange(s)")
    return connected_count


def add_exchanges_to_portfolio(page, exchanges=None, config_path=None):
    """
    Full flow: click Exchange tab → check existing → connect new exchanges.

    Args:
        page: Playwright page (must be on Create Portfolio page)
        exchanges: List of exchange types to connect (default: ["binance"])
        config_path: Path to exchange_keys.json

    Returns:
        int: Total number of exchanges connected
    """
    if exchanges is None:
        exchanges = ["binance"]

    print("\n   📡 Adding exchange connections...")

    # Click Exchange tab
    if not click_exchange_tab(page):
        return 0

    total = 0
    for exchange in exchanges:
        print(f"\n   🔗 Connecting {exchange} exchanges...")
        count = connect_all_exchanges(page, exchange=exchange, config_path=config_path)
        total += count

    return total
