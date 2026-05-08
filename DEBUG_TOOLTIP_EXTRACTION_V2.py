"""
Debug script to investigate why tooltips are not being extracted.
Tests the extraction logic with detailed logging.
"""

from playwright.sync_api import sync_playwright
import time

def debug_tooltip_extraction():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # Show browser
        page = browser.new_page()
        
        # Navigate to DAM
        print("🌐 Navigating to DAM...")
        page.goto("https://dam.defiassets.com/", wait_until="networkidle")
        page.wait_for_timeout(2000)
        
        # Login (if needed)
        print("🔐 Checking if login needed...")
        try:
            page.fill('input[type="email"]', "moontest1311@gmail.com")
            page.fill('input[type="password"]', "Test@123456")
            page.click('button:has-text("Sign In")')
            page.wait_for_timeout(3000)
        except:
            print("   Already logged in or login not needed")
        
        # Navigate to portfolio
        print("📊 Navigating to portfolio...")
        page.goto("https://dam.defiassets.com/portfolio/moontest1311_CEX", wait_until="networkidle")
        page.wait_for_timeout(2000)
        
        # Click on Overview - Wallet tab
        print("📋 Clicking Overview - Wallet tab...")
        try:
            page.click('text=Overview - Wallet')
            page.wait_for_timeout(2000)
        except:
            print("   Tab already selected")
        
        # Get first few rows
        print("\n" + "="*80)
        print("DEBUGGING TOOLTIP EXTRACTION")
        print("="*80)
        
        table_rows = page.locator('table tbody tr').all()
        print(f"Total rows found: {len(table_rows)}")
        
        # Test first 5 rows
        for row_idx in range(min(5, len(table_rows))):
            row = table_rows[row_idx]
            token_name = row.locator('td:nth-child(2)').text_content().strip()
            
            print(f"\n{'='*80}")
            print(f"ROW {row_idx}: {token_name}")
            print(f"{'='*80}")
            
            # Check if price tooltip element exists
            price_elem = row.locator('[data-tooltip-id*="price-tooltip"]').first
            print(f"Price element found: {price_elem.count() > 0}")
            
            if price_elem.count() > 0:
                price_tooltip_id = price_elem.get_attribute('data-tooltip-id')
                print(f"Tooltip ID: {price_tooltip_id}")
                
                # Check visibility BEFORE hover
                is_visible_before = price_elem.is_visible()
                print(f"Visible BEFORE hover: {is_visible_before}")
                
                # Hover
                print("Hovering...")
                price_elem.hover()
                page.wait_for_timeout(500)
                
                # Check visibility AFTER hover
                is_visible_after = price_elem.is_visible()
                print(f"Visible AFTER hover: {is_visible_after}")
                
                # Check if tooltip element exists in DOM
                tooltip_exists = page.evaluate(f"""
                    (id) => {{
                        const tooltip = document.getElementById(id);
                        return tooltip !== null;
                    }}
                """, price_tooltip_id)
                print(f"Tooltip element exists in DOM: {tooltip_exists}")
                
                # Check offsetParent
                offset_parent_null = page.evaluate(f"""
                    (id) => {{
                        const tooltip = document.getElementById(id);
                        if (tooltip) {{
                            return tooltip.offsetParent === null;
                        }}
                        return null;
                    }}
                """, price_tooltip_id)
                print(f"offsetParent === null: {offset_parent_null}")
                
                # Check display style
                display_style = page.evaluate(f"""
                    (id) => {{
                        const tooltip = document.getElementById(id);
                        if (tooltip) {{
                            return window.getComputedStyle(tooltip).display;
                        }}
                        return null;
                    }}
                """, price_tooltip_id)
                print(f"Display style: {display_style}")
                
                # Check visibility style
                visibility_style = page.evaluate(f"""
                    (id) => {{
                        const tooltip = document.getElementById(id);
                        if (tooltip) {{
                            return window.getComputedStyle(tooltip).visibility;
                        }}
                        return null;
                    }}
                """, price_tooltip_id)
                print(f"Visibility style: {visibility_style}")
                
                # Try to extract text
                tooltip_text = page.evaluate(f"""
                    (id) => {{
                        const tooltip = document.getElementById(id);
                        if (tooltip) {{
                            return tooltip.textContent.trim();
                        }}
                        return null;
                    }}
                """, price_tooltip_id)
                print(f"Tooltip text: {tooltip_text}")
                
                # Try alternative extraction (check all children)
                tooltip_html = page.evaluate(f"""
                    (id) => {{
                        const tooltip = document.getElementById(id);
                        if (tooltip) {{
                            return tooltip.innerHTML;
                        }}
                        return null;
                    }}
                """, price_tooltip_id)
                print(f"Tooltip HTML (first 200 chars): {str(tooltip_html)[:200] if tooltip_html else 'None'}")
        
        browser.close()

if __name__ == "__main__":
    debug_tooltip_extraction()
