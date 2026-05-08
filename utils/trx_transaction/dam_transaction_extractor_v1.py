#!/usr/bin/env python3
"""
DAM Transaction Extraction Script
- Logs into DAM at https://dam-sit.mqbc21.com/sign-in
- Opens portfolio for address TWd4WrZ9wn84f5x1hZhL4DHvk738ns5jwb
- Filters transactions for 2026-01-21
- Captures all transaction data and generates Excel file
"""

import time
import json
import re
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

TARGET_ADDRESS = "TWd4WrZ9wn84f5x1hZhL4DHvk738ns5jwb"
TARGET_DATE = "2026-01-21"
EMAIL = "roninx688@gmail.com"
PASSWORD = "787193@PyBt7871"
BASE_URL = "https://dam-sit.mqbc21.com"


def run():
    transactions = []
    screenshot_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test-results", "screenshots")
    import os
    os.makedirs(screenshot_dir, exist_ok=True)

    with sync_playwright() as p:
        HEADLESS = os.environ.get('HEADLESS', 'false').lower() == 'true'
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=300)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()
        page.set_default_timeout(30000)

        # ── STEP 1: Sign in ──────────────────────────────────────────────────
        print("\n[1] Navigating to sign-in page...")
        page.goto(f"{BASE_URL}/sign-in")
        page.wait_for_load_state("domcontentloaded")

        page.fill('input[data-testid="input-email"]', EMAIL)
        page.fill('input[data-testid="input-password"]', PASSWORD)
        page.click('button[data-testid="sign-in-btn"]')
        print("   Submitted sign-in form, waiting for redirect...")

        try:
            page.wait_for_url("**/portfolio**", timeout=20000)
            print("   ✅ Signed in successfully")
        except PlaywrightTimeoutError:
            page.screenshot(path=f"{screenshot_dir}/dam_signin_error.png")
            print("   ❌ Sign-in redirect timed out. Screenshot saved.")
            browser.close()
            return []

        page.wait_for_timeout(2000)
        # Close any popup/modal
        for sel in ["button[aria-label*='close' i]", "button[aria-label*='dismiss' i]",
                    "button:has-text('×')", "button:has-text('Close')"]:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=1000):
                    btn.click()
                    break
            except Exception:
                pass

        # ── STEP 2: Find portfolio for target address ────────────────────────
        print(f"\n[2] Looking for portfolio with address {TARGET_ADDRESS}...")

        # Try clicking the portfolio dropdown/selector
        portfolio_found = False
        portfolio_name = None

        # Look for portfolio dropdown button
        dropdown_selectors = [
            "button[data-testid*='portfolio']",
            "[class*='portfolio-selector']",
            "[class*='portfolio-dropdown']",
            "button:has-text('Portfolio')",
        ]
        for sel in dropdown_selectors:
            try:
                btn = page.locator(sel).first
                if btn.is_visible(timeout=2000):
                    btn.click()
                    page.wait_for_timeout(1000)
                    print(f"   Opened dropdown via: {sel}")
                    break
            except Exception:
                pass

        # Try to find portfolio items and look for our address
        try:
            # Scroll through portfolio list to find the one with our address
            portfolio_items = page.locator("[class*='portfolio-item'], [class*='portfolio-card'], li[class*='portfolio']")
            count = portfolio_items.count()
            print(f"   Found {count} portfolio items in dropdown")

            for i in range(count):
                item = portfolio_items.nth(i)
                text = item.text_content() or ""
                if TARGET_ADDRESS[:8] in text or TARGET_ADDRESS in text:
                    portfolio_name = text.strip().split('\n')[0]
                    item.click()
                    portfolio_found = True
                    print(f"   ✅ Found portfolio: {portfolio_name}")
                    break
        except Exception as e:
            print(f"   Portfolio search error: {e}")

        if not portfolio_found:
            # Try searching via search input in portfolio dropdown
            try:
                search_input = page.locator("input[placeholder*='search' i], input[placeholder*='portfolio' i]").first
                if search_input.is_visible(timeout=2000):
                    search_input.fill(TARGET_ADDRESS[:10])
                    page.wait_for_timeout(1500)
                    result = page.locator(f"[class*='portfolio-item']:has-text('{TARGET_ADDRESS[:10]}'), li:has-text('{TARGET_ADDRESS[:10]}')").first
                    if result.is_visible(timeout=3000):
                        portfolio_name = result.text_content().strip().split('\n')[0]
                        result.click()
                        portfolio_found = True
                        print(f"   ✅ Found portfolio via search: {portfolio_name}")
            except Exception as e:
                print(f"   Search attempt failed: {e}")

        page.wait_for_timeout(2000)
        page.screenshot(path=f"{screenshot_dir}/dam_portfolio_loaded.png")
        print(f"   Current URL: {page.url}")

        # ── STEP 3: Navigate to Transactions / Activity section ──────────────
        print(f"\n[3] Looking for Transactions/Activity section...")
        page.wait_for_timeout(1000)

        # Try to find and click a "Transactions" or "Activity" tab/link
        tx_nav_selectors = [
            "a:has-text('Transaction')", "button:has-text('Transaction')",
            "a:has-text('Activity')", "button:has-text('Activity')",
            "[data-testid*='transaction']", "[data-testid*='activity']",
            "a[href*='transaction']", "a[href*='activity']",
            "[class*='transaction-tab']", "[class*='activity-tab']",
            "li:has-text('Transaction')", "li:has-text('Activity')",
        ]

        nav_clicked = False
        for sel in tx_nav_selectors:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=1500):
                    el.click()
                    page.wait_for_timeout(2000)
                    nav_clicked = True
                    print(f"   ✅ Clicked nav: {sel}")
                    break
            except Exception:
                pass

        if not nav_clicked:
            # Try URL-based navigation
            current_url = page.url
            for suffix in ["/transactions", "/activity", "?tab=transactions", "?tab=activity"]:
                try:
                    if "portfolioId=" in current_url or "/portfolio/" in current_url:
                        nav_url = current_url.split("?")[0] + suffix
                    else:
                        nav_url = current_url.rstrip("/") + suffix
                    page.goto(nav_url)
                    page.wait_for_load_state("networkidle")
                    if page.locator("table, [class*='transaction']").first.is_visible(timeout=3000):
                        nav_clicked = True
                        print(f"   ✅ Navigated to: {nav_url}")
                        break
                except Exception:
                    pass

        page.wait_for_timeout(2000)
        page.screenshot(path=f"{screenshot_dir}/dam_transactions_page.png")
        print(f"   URL after navigation: {page.url}")

        # ── STEP 4: Apply date filter for 2026-01-21 ────────────────────────
        print(f"\n[4] Applying date filter for {TARGET_DATE}...")
        page.wait_for_timeout(1000)

        # Try to find date filter inputs
        date_filter_applied = False

        # Look for date range picker or date inputs
        date_selectors = [
            "input[type='date']",
            "input[placeholder*='date' i]",
            "input[placeholder*='from' i]",
            "[data-testid*='date']",
            "[class*='date-picker']",
            "[class*='datepicker']",
            "button:has-text('Date')",
            "button:has-text('Filter')",
            "[class*='date-filter']",
        ]

        for sel in date_selectors:
            try:
                el = page.locator(sel).first
                if el.is_visible(timeout=1500):
                    print(f"   Found date element: {sel}")
                    el.click()
                    page.wait_for_timeout(1000)
                    break
            except Exception:
                pass

        # Try filling date inputs directly
        try:
            date_inputs = page.locator("input[type='date']")
            if date_inputs.count() >= 1:
                date_inputs.first.fill(TARGET_DATE)
                page.keyboard.press("Enter")
                if date_inputs.count() >= 2:
                    date_inputs.nth(1).fill(TARGET_DATE)
                    page.keyboard.press("Enter")
                page.wait_for_timeout(2000)
                date_filter_applied = True
                print(f"   ✅ Date inputs filled with {TARGET_DATE}")
        except Exception as e:
            print(f"   Date input attempt: {e}")

        page.wait_for_timeout(2000)
        page.screenshot(path=f"{screenshot_dir}/dam_date_filtered.png")

        # ── STEP 5: Extract transaction data from table ──────────────────────
        print(f"\n[5] Extracting transaction data...")
        page.wait_for_timeout(2000)

        # Try multiple table/list selectors
        page_content = page.content()

        # Look for transaction rows
        tx_row_selectors = [
            "table tbody tr",
            "[class*='transaction-row']",
            "[class*='tx-row']",
            "[class*='transaction-item']",
            "[class*='activity-row']",
            "[data-testid*='transaction-row']",
        ]

        rows_found = False
        for sel in tx_row_selectors:
            try:
                rows = page.locator(sel)
                count = rows.count()
                if count > 0:
                    print(f"   Found {count} rows with selector: {sel}")
                    rows_found = True

                    for i in range(count):
                        row = rows.nth(i)
                        try:
                            row_text = row.text_content() or ""
                            # Try to get all cells
                            cells = row.locator("td, [class*='cell'], [class*='col']")
                            cell_texts = []
                            for j in range(cells.count()):
                                cell_texts.append(cells.nth(j).text_content().strip())

                            tx_data = {
                                "row_index": i + 1,
                                "full_text": row_text.strip(),
                                "cells": cell_texts,
                            }
                            transactions.append(tx_data)
                            print(f"   Row {i+1}: {row_text[:100].strip()}")
                        except Exception as e:
                            print(f"   Error reading row {i}: {e}")
                    break
            except Exception as e:
                print(f"   Selector {sel} error: {e}")

        if not rows_found:
            print("   No transaction rows found with standard selectors.")
            # Capture full page text for analysis
            all_text = page.locator("body").text_content()
            print(f"   Page text preview: {all_text[:500]}")
            page.screenshot(path=f"{screenshot_dir}/dam_no_transactions.png", full_page=True)

        # ── STEP 6: Try to click individual transactions for detail ──────────
        print(f"\n[6] Attempting to get detailed transaction info...")

        detailed_transactions = []

        if rows_found and len(transactions) > 0:
            for i, tx in enumerate(transactions):
                try:
                    rows = page.locator(tx_row_selectors[0] if tx_row_selectors else "table tbody tr")
                    row = rows.nth(i)
                    row.click()
                    page.wait_for_timeout(1500)

                    # Try to extract detail from modal/panel
                    detail_selectors = [
                        "[class*='transaction-detail']",
                        "[class*='modal']",
                        "[class*='drawer']",
                        "[class*='panel']",
                        "[role='dialog']",
                    ]

                    detail_text = ""
                    for d_sel in detail_selectors:
                        try:
                            detail = page.locator(d_sel).first
                            if detail.is_visible(timeout=2000):
                                detail_text = detail.text_content().strip()
                                break
                        except Exception:
                            pass

                    detailed_transactions.append({
                        **tx,
                        "detail_text": detail_text,
                    })

                    # Close modal/detail
                    for close_sel in ["button[aria-label*='close' i]", "button:has-text('×')", "button:has-text('Close')", "Escape"]:
                        try:
                            if close_sel == "Escape":
                                page.keyboard.press("Escape")
                            else:
                                page.locator(close_sel).first.click()
                            page.wait_for_timeout(500)
                            break
                        except Exception:
                            pass

                except Exception as e:
                    print(f"   Error getting detail for row {i}: {e}")
                    detailed_transactions.append(tx)

        final_transactions = detailed_transactions if detailed_transactions else transactions

        # ── STEP 7: Screenshot full page ────────────────────────────────────
        page.screenshot(path=f"{screenshot_dir}/dam_final_state.png", full_page=True)
        print(f"\n   Screenshots saved to {screenshot_dir}")

        # Capture complete page HTML for analysis
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test-results", "dam_transactions_page.html")
        with open(html_path, "w") as f:
            f.write(page.content())
        print(f"   HTML saved to: {html_path}")

        browser.close()

    return final_transactions


def save_transactions_to_json(transactions):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test-results", "dam_transactions_raw.json")
    import os
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(transactions, f, indent=2)
    print(f"   Raw transaction data saved to: {path}")


if __name__ == "__main__":
    print("=" * 60)
    print(f"DAM Transaction Extraction")
    print(f"Address: {TARGET_ADDRESS}")
    print(f"Date: {TARGET_DATE}")
    print("=" * 60)

    transactions = run()
    save_transactions_to_json(transactions)
    print(f"\nTotal transactions captured: {len(transactions)}")
    for i, tx in enumerate(transactions, 1):
        print(f"  {i}: {str(tx)[:120]}")
