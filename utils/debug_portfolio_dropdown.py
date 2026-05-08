#!/usr/bin/env python3
"""
Debug script to inspect portfolio dropdown structure
"""

from playwright.sync_api import sync_playwright
import time

def debug_portfolio_dropdown():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # Navigate to DAM
        print("🌐 Navigating to DAM...")
        page.goto("https://dam-sit.mqbc21.com/sign-in")
        page.wait_for_timeout(3000)

        # Login
        print("🔐 Logging in...")
        page.fill('input[data-testid="input-email"]', 'moontest1311@gmail.com')
        page.fill('input[data-testid="input-password"]', 'Orion888!!!!')
        page.click('button[data-testid="sign-in-btn"]')
        page.wait_for_timeout(5000)

        # Click portfolio dropdown - look for button that contains "Portfolio" text
        print("\n📁 Looking for portfolio selector button...")
        portfolio_clicked = False
        try:
            # Look for button with "Portfolio" text in it
            portfolio_button = page.locator("button").filter(has_text="Portfolio").first
            if portfolio_button.is_visible(timeout=3000):
                print(f"   ✅ Found portfolio button: {portfolio_button.inner_text()[:50]}")
                portfolio_button.click()
                page.wait_for_timeout(2000)
                portfolio_clicked = True
                print("   ✅ Clicked portfolio dropdown")

                # Take screenshot
                page.screenshot(path="debug_portfolio_menu_opened.png")
                print("   📸 Screenshot: debug_portfolio_menu_opened.png")

                # Look for the menu with portfolio list
                menu = page.locator("[role='menu']").last  # Use last menu as there might be multiple
                if menu.is_visible(timeout=1000):
                    menu_html = menu.inner_html()
                    print(f"\n   Menu HTML (first 1000 chars):\n{menu_html[:1000]}")

                    # Look for all menu items
                    menu_items = menu.locator("[role='menuitem']").all()
                    print(f"\n   Found {len(menu_items)} menu items:")
                    for idx, item in enumerate(menu_items):
                        text = item.inner_text().strip()
                        print(f"      {idx}: {text}")

        except Exception as e:
            print(f"   ❌ Error: {e}")

        # Analyze dropdown structure
        print("\n🔍 Analyzing portfolio dropdown items...")
        try:
            # Get all buttons in the dropdown
            all_buttons = page.locator("button").all()
            print(f"\n   Found {len(all_buttons)} total buttons on page")

            # Look for buttons that might be portfolio items
            print("\n   Looking for portfolio item buttons...")
            for idx, btn in enumerate(all_buttons):
                try:
                    if btn.is_visible(timeout=100):
                        btn_text = btn.inner_text().strip()
                        if "address" in btn_text.lower() or "zg" in btn_text.lower():
                            class_attr = btn.get_attribute("class")
                            print(f"\n   Button {idx}:")
                            print(f"      Text: {btn_text}")
                            print(f"      Class: {class_attr}")

                            # Get the HTML
                            html = btn.inner_html()
                            print(f"      HTML (first 200 chars): {html[:200]}")
                except:
                    continue

            # Look for all divs containing portfolio text
            print("\n\n🎯 Searching for all portfolio-like divs...")
            all_divs = page.locator("div.text-mono-900.typography-body").all()
            print(f"   Found {len(all_divs)} divs with portfolio styling")
            for idx, div in enumerate(all_divs):
                try:
                    if div.is_visible(timeout=100):
                        text = div.inner_text().strip()
                        if text and len(text) > 0:
                            print(f"   Div {idx}: {text}")
                            # Check if this is "zg's address - 1"
                            if "zg" in text.lower() or "address" in text.lower():
                                parent = div.locator("xpath=..").first  # Get parent element
                                parent_tag = parent.evaluate("el => el.tagName")
                                parent_class = parent.get_attribute("class")
                                print(f"      👆 Parent tag: {parent_tag}")
                                print(f"      👆 Parent class: {parent_class[:100]}")
                except:
                    continue

            # Try to find the specific portfolio
            print("\n\n🎯 Searching for 'zg's address - 1'...")

            # Method 1: text-is
            try:
                option1 = page.locator("button").filter(has_text="zg's address - 1").first
                if option1.is_visible(timeout=1000):
                    print("   ✅ Found with filter(has_text)")
                    print(f"      Class: {option1.get_attribute('class')}")
            except:
                print("   ❌ Not found with filter(has_text)")

            # Method 2: has div with text
            try:
                option2 = page.locator("button:has(div:text-is('zg\\'s address - 1'))").first
                if option2.is_visible(timeout=1000):
                    print("   ✅ Found with :has(div:text-is)")
                    print(f"      Class: {option2.get_attribute('class')}")
            except:
                print("   ❌ Not found with :has(div:text-is)")

            # Method 3: Direct text locator
            try:
                option3 = page.get_by_text("zg's address - 1", exact=True).first
                if option3.is_visible(timeout=1000):
                    print("   ✅ Found with get_by_text")
                    tag = option3.evaluate("el => el.tagName")
                    print(f"      Tag: {tag}")
                    print(f"      Class: {option3.get_attribute('class')}")
            except:
                print("   ❌ Not found with get_by_text")

            # Method 4: Click on the div that contains the text with exact match
            try:
                # Look for div with exact text match
                target_portfolio = "zg's address - 1"
                print(f"\n   🎯 Looking for exact match: '{target_portfolio}'")

                # Get all divs and check for exact match
                all_portfolio_divs = page.locator("div.text-mono-900.typography-body.font-normal.break-all").all()
                print(f"   Found {len(all_portfolio_divs)} portfolio name divs")

                for idx, div in enumerate(all_portfolio_divs):
                    try:
                        if div.is_visible(timeout=100):
                            text = div.inner_text().strip()
                            if text == target_portfolio:
                                print(f"   ✅ Found exact match at index {idx}: '{text}'")
                                # Get the parent menuitem and click it
                                menu_item = div.locator("xpath=ancestor::div[@role='menuitem']").first
                                if menu_item.is_visible(timeout=1000):
                                    print("      Clicking menu item...")
                                    menu_item.click()
                                    page.wait_for_timeout(2000)
                                    print("      ✅ Successfully clicked!")
                                    break
                    except:
                        continue
            except Exception as e:
                print(f"   ❌ Error with exact match method: {e}")

        except Exception as e:
            print(f"   ❌ Error analyzing dropdown: {e}")

        # Take screenshot
        print("\n📸 Taking screenshot...")
        page.screenshot(path="debug_portfolio_dropdown.png")
        print("   ✅ Screenshot saved as debug_portfolio_dropdown.png")

        print("\n⏸️  Browser will stay open for 30 seconds for manual inspection...")
        time.sleep(30)

        browser.close()

if __name__ == "__main__":
    debug_portfolio_dropdown()
