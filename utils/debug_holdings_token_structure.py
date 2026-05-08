#!/usr/bin/env python3
"""
Debug script to inspect Holdings - Token table structure
"""

from playwright.sync_api import sync_playwright
import time

def debug_holdings_token_structure():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # Navigate to DAM
        print("🌐 Navigating to DAM...")
        page.goto("https://app.dam.io/account/login")
        page.wait_for_timeout(2000)

        # Login
        print("🔐 Logging in...")
        page.fill('input[type="email"]', 'testaccount2@wealthgen.xyz')
        page.fill('input[type="password"]', 'Orion888!!!!')
        page.click('button[type="submit"]')
        page.wait_for_timeout(5000)

        # Select portfolio
        print("📁 Selecting portfolio 'zg's address - 1'...")
        try:
            page.click("button:has-text('zg\\'s address - 1')")
            page.wait_for_timeout(3000)
        except:
            print("⚠️  Could not find portfolio selector")

        # Navigate to Holdings
        print("\n📋 Navigating to Holdings tab...")
        try:
            page.click('button[role="tab"]:has-text("Holdings")')
            page.wait_for_timeout(3000)
        except:
            print("⚠️  Could not find Holdings tab")

        # Get first token row with button
        print("\n🔍 Analyzing first Token_Main row structure...")
        try:
            first_token_row = page.locator("tbody tr").filter(has=page.locator("button.cursor-pointer.rounded")).first
            cells = first_token_row.locator("td, [role='cell']").all()

            print(f"\n  Total cells: {len(cells)}")
            for idx, cell in enumerate(cells):
                text = cell.inner_text().strip()
                # Show first 100 chars
                print(f"  Cell {idx}: {text[:100]}")

        except Exception as e:
            print(f"  ❌ Error: {e}")

        # Get first chain detail row (without button)
        print("\n🔍 Analyzing first Chain_Detail row structure...")
        try:
            all_rows = page.locator("tbody tr").all()
            for idx, row in enumerate(all_rows):
                has_button = row.locator("td:first-child button.cursor-pointer.rounded").count() > 0
                if not has_button:
                    cells = row.locator("td, [role='cell']").all()
                    print(f"\n  Row {idx + 1} (Chain Detail):")
                    print(f"  Total cells: {len(cells)}")
                    for cell_idx, cell in enumerate(cells):
                        text = cell.inner_text().strip()
                        print(f"  Cell {cell_idx}: {text[:100]}")
                    break

        except Exception as e:
            print(f"  ❌ Error: {e}")

        # Take screenshot
        print("\n📸 Taking screenshot...")
        page.screenshot(path="debug_holdings_token_structure.png")
        print("  ✅ Screenshot saved as debug_holdings_token_structure.png")

        print("\n⏸️  Browser will stay open for 30 seconds for manual inspection...")
        time.sleep(30)

        browser.close()

if __name__ == "__main__":
    debug_holdings_token_structure()
