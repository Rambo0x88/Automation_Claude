"""
Portfolio Page Object Model (Playwright)
"""
from playwright.sync_api import Page, expect
from pages.base_page import BasePage
from config.config import Config


class PortfolioPage(BasePage):
    """Portfolio page object"""

    def __init__(self, page: Page, base_url: str = None):
        super().__init__(page, base_url)
        self.url = Config.PORTFOLIO_URL

        # Locators (Playwright style)
        self.create_portfolio_button = page.locator("button:has-text('Create portfolio')")
        self.portfolio_name_input = page.locator("input[placeholder*='portfolio name' i], input[name='name'], input[id*='portfolioName' i]")
        self.portfolio_description_input = page.locator("textarea[placeholder*='description' i], textarea[name='description']")
        # More specific selector for Save button - use type="submit" to distinguish from "Create portfolio" button
        self.save_portfolio_button = page.locator("button[type='submit']:has-text('Save'), button[type='submit']:has-text('Create'), button[type='submit']:has-text('Submit')")
        self.cancel_button = page.locator("button:has-text('Cancel')")

        # Portfolio list
        self.portfolio_cards = page.locator("[class*='portfolio-card'], [class*='portfolio']")
        self.portfolio_titles = page.locator("h2[class*='portfolio'], [class*='portfolio-name'], [class*='portfolio-title']")

        # Empty state
        self.empty_state_message = page.locator("text=/haven't created a portfolio yet|Start Managing Your Digital Assets/i")
        self.empty_state_button = page.locator("button:has-text('Create portfolio')")

        # Success messages
        self.success_message = page.locator("text=/Success|created successfully/i")

        # User profile
        self.user_profile_icon = page.locator("button[aria-label*='user' i], [class*='profile'], [class*='user-icon']")

    def navigate(self):
        """Navigate to portfolio page"""
        self.navigate_to(self.url)

    def is_on_portfolio_page(self) -> bool:
        """Check if currently on portfolio page"""
        return '/portfolio' in self.get_current_url()

    def is_empty_state_displayed(self) -> bool:
        """Check if empty state (no portfolios) is displayed"""
        try:
            expect(self.empty_state_message).to_be_visible(timeout=10000)
            return True
        except:
            return False

    def click_create_portfolio(self):
        """Click Create Portfolio button"""
        if self.is_empty_state_displayed():
            self.empty_state_button.click()
        else:
            # Try to find and click the Create Portfolio button
            # If not visible, scroll down in the dropdown to find it
            try:
                if not self.create_portfolio_button.is_visible(timeout=2000):
                    print("      🔍 Create portfolio button not visible, scrolling down...")
                    # Scroll down in the dropdown by pressing Page Down or using mouse wheel
                    self.page.keyboard.press("PageDown")
                    self.page.wait_for_timeout(500)

                # Now click the button
                self.create_portfolio_button.wait_for(state='visible', timeout=5000)
                self.create_portfolio_button.scroll_into_view_if_needed()
                self.create_portfolio_button.click()
            except Exception as e:
                print(f"      ⚠️ Error finding Create portfolio button: {e}")
                # Fallback: just try to click
                self.create_portfolio_button.click()

    def enter_portfolio_name(self, name: str):
        """Enter portfolio name"""
        self.portfolio_name_input.fill(name)

    def enter_portfolio_description(self, description: str):
        """Enter portfolio description"""
        if self.portfolio_description_input.is_visible(timeout=3000):
            self.portfolio_description_input.fill(description)

    def add_address(self, address: str, chain: str = None, tags: list = None):
        """
        Add a crypto address to portfolio
        Based on HTML structure: input[data-testid="input-wallet.0.address"]

        Flow (as observed by user):
        1. Input portfolio name
        2. Input address and press Enter
        3. Address gets added to list, "Add Tag(s)" button appears
        4. Save button becomes enabled

        Args:
            address: Wallet address (ETH/BSC format)
            chain: Blockchain network (ETH, BSC, Polygon, etc.) - currently not used
            tags: Optional list of tag names to add (not implemented yet)
        """
        print(f"      🔍 Looking for address input field...")

        # Find address input fields - they have data-testid like "input-wallet.0.address"
        # Get the last visible one (the empty input for new address)
        address_inputs = self.page.locator("input[data-testid*='input-wallet']")

        # Wait for at least one input to be visible
        print(f"      ⏳ Waiting for address input to be visible...")
        address_inputs.first.wait_for(state='visible', timeout=10000)

        # Get count of inputs
        input_count = address_inputs.count()
        print(f"      ✅ Found {input_count} address input field(s)")

        # Get the last input (the empty one)
        last_input = address_inputs.last

        # Ensure it's visible and enabled
        last_input.wait_for(state='visible', timeout=5000)

        # Click to focus
        print(f"      🖱️  Clicking address input field...")
        last_input.click()
        self.page.wait_for_timeout(500)

        # Clear any existing value first
        last_input.fill("")
        self.page.wait_for_timeout(300)

        # Fill the address
        print(f"      ⌨️  Typing address: {address[:10]}...{address[-8:]}")
        last_input.fill(address)
        self.page.wait_for_timeout(800)

        # Press Enter to add address to the list (this makes Save button appear)
        print(f"      ⏎  Pressing Enter to add address...")
        last_input.press("Enter")

        # Wait for address to be added and Save button to become enabled
        # Need to wait longer for address validation to complete
        self.page.wait_for_timeout(3000)
        print(f"      ✅ Address input completed")

        # Optional: Add tags if provided (for future implementation)
        if tags and len(tags) > 0:
            try:
                print(f"      🏷️  Adding tags: {tags}")
                # Click "Add Tag(s)" button that appears after adding address
                add_tag_button = self.page.locator("button:has-text('Add Tag(s)')").last
                add_tag_button.wait_for(state='visible', timeout=3000)
                add_tag_button.click()
                self.page.wait_for_timeout(500)

                # Tag modal should appear - add tags one by one
                for tag in tags:
                    tag_input = self.page.locator("input[placeholder*='tag' i]").first
                    tag_input.fill(tag)
                    tag_input.press("Enter")
                    self.page.wait_for_timeout(300)

                # Click Save button in tag modal
                save_tag_button = self.page.locator("button:has-text('Save')").first
                save_tag_button.click()
                self.page.wait_for_timeout(500)
                print(f"      ✅ Tags added successfully")
            except Exception as e:
                print(f"      ⚠️ Could not add tags: {e}")
                # Continue anyway - tags are optional

    def click_save_portfolio(self):
        """Click Save/Create button"""
        print("💾 Clicking Save Portfolio button...")

        # Wait for Save button to become enabled (validation complete)
        print("   ⏳ Waiting for Save button to be enabled...")
        try:
            # Wait for button to NOT have disabled attribute
            enabled_button = self.page.locator("button[type='submit']:has-text('Save'):not([disabled]), button[type='submit']:has-text('Create'):not([disabled]), button[type='submit']:has-text('Submit'):not([disabled])")
            enabled_button.wait_for(state='visible', timeout=15000)
            print("   ✅ Save button is now enabled")
            enabled_button.click()
        except Exception as e:
            print(f"   ⚠️ Timeout waiting for Save button to be enabled: {e}")
            print("   Attempting to click anyway...")
            self.save_portfolio_button.click()

        print("✅ Save Portfolio button clicked successfully")

    def save_portfolio(self):
        """
        Save portfolio and wait for it to be created
        This is a wrapper that handles the complete save flow
        """
        self.click_save_portfolio()
        # Wait for portfolio to be saved and dialog to close
        self.page.wait_for_timeout(3000)

    def click_cancel(self):
        """Click Cancel button"""
        self.cancel_button.click()

    def create_portfolio(self, name: str, description: str = ""):
        """Complete portfolio creation flow"""
        self.click_create_portfolio()
        self.enter_portfolio_name(name)
        if description:
            self.enter_portfolio_description(description)
        self.click_save_portfolio()

    def wait_for_portfolio_created(self, portfolio_name: str, timeout: int = 15000) -> bool:
        """
        Wait for portfolio to be created and displayed
        Returns True if portfolio appears, False otherwise
        """
        try:
            # Wait a moment for UI update
            self.page.wait_for_timeout(2000)
            return self.is_portfolio_displayed(portfolio_name)
        except:
            return False

    def is_portfolio_displayed(self, portfolio_name: str) -> bool:
        """Check if portfolio with given name is displayed"""
        try:
            portfolio_locator = self.page.locator(f"text={portfolio_name}")
            return portfolio_locator.is_visible(timeout=5000)
        except:
            return False

    def get_all_portfolio_names(self) -> list:
        """Get list of all portfolio names"""
        try:
            # First, wait for page to load
            self.page.wait_for_timeout(2000)

            # Try multiple selectors to find portfolios
            print("🔍 Searching for portfolios on page...")

            # Method 1: Look for portfolio cards/containers
            portfolio_containers = self.page.locator("[class*='portfolio'], [data-testid*='portfolio']").all()
            print(f"   Found {len(portfolio_containers)} elements with 'portfolio' in class/testid")

            # Method 2: Look for h1, h2, h3 headings that might be portfolio names
            headings = self.page.locator("h1, h2, h3, h4").all()
            print(f"   Found {len(headings)} headings on page")

            portfolio_names = []

            # Check if there are any portfolios using our original selector
            portfolios = self.portfolio_titles.all()
            if len(portfolios) > 0:
                portfolio_names = [p.text_content().strip() for p in portfolios if p.text_content()]
                print(f"   ✅ Method 1 (original selector): Found {len(portfolio_names)} portfolios")
                return portfolio_names

            # If original selector fails, try to find all portfolio items in the dropdown
            # Look for portfolio items - they typically have specific classes or roles
            portfolio_item_selectors = [
                "[role='menuitem']",  # Common dropdown menu item
                "[class*='menu-item']",
                "[class*='dropdown-item']",
                "[class*='portfolio-item']",
                "li[role='option']",
                "div[role='option']"
            ]

            # Write debug output to file since xdist suppresses print
            debug_file = open("/tmp/portfolio_debug.log", "w")

            for selector in portfolio_item_selectors:
                try:
                    items = self.page.locator(selector).all()
                    if len(items) > 0:
                        msg = f"   🔍 DEBUG: Found {len(items)} items with selector: {selector}\n"
                        print(msg)
                        debug_file.write(msg)
                        debug_file.flush()
                        # Extract text from each item
                        for idx, item in enumerate(items):
                            # Use inner_text so we can strip the second line ("1 addresses")
                            text = item.inner_text() if hasattr(item, "inner_text") else item.text_content()
                            if text:
                                original_text = text
                                text = text.strip()
                                # Portfolio items have THREE lines:
                                # Line 0: First letter of portfolio name (e.g., 'z', 'l')
                                # Line 1: Actual portfolio name (e.g., "zg's address - 27")
                                # Line 2: Address count (e.g., "1 addresses")
                                # We want line 1 (index 1)
                                if '\n' in text:
                                    lines = text.split('\n')
                                    if len(lines) >= 2:
                                        # Take the second line (actual portfolio name)
                                        text = lines[1].strip()
                                    else:
                                        # Fallback: if only one line, use it
                                        text = lines[0].strip()

                                msg = f"   🔍 DEBUG [{idx}]: Raw text: '{original_text}' → After split: '{text}'\n"
                                print(msg, end='')
                                debug_file.write(msg)
                                debug_file.flush()

                                # Filter out non-portfolio items (common UI elements)
                                if text and text not in ['Portfolio', 'Create Portfolio', 'Settings', 'Logout', 'Create portfolio']:
                                    # Remove trailing ": 1 addresses" text that sometimes lingers
                                    import re
                                    before_regex = text
                                    text = re.sub(r':?\s*\d+\s+address(?:es)?$', '', text).strip()
                                    if before_regex != text:
                                        print(f"   🔍 DEBUG: Regex changed '{before_regex}' → '{text}'")

                                    # Fix cases where portfolio names have duplicate first character due to HTML structure
                                    # Example: <span>l</span><span>large data</span> results in "llarge data"
                                    # ONLY apply this fix if the first TWO characters are identical
                                    if len(text) > 2:
                                        before_dedup = text
                                        # Case 1: "llarge data" -> "large data" (first char duplicates second char exactly)
                                        if text[0] == text[1] and text[0].isalpha():
                                            text = text[1:]
                                            print(f"   🔍 DEBUG: Dedup changed '{before_dedup}' → '{text}' (duplicate first char)")
                                        # Case 2 removed - it was too aggressive

                                    # Also handle pure digit strings (like "290")
                                    if text and not text.isdigit() and text not in portfolio_names:
                                        print(f"   ✅ DEBUG: Adding portfolio: '{text}'")
                                        portfolio_names.append(text)
                                    elif text in portfolio_names:
                                        print(f"   ⚠️  DEBUG: Skipping duplicate: '{text}'")

                        if len(portfolio_names) > 0:
                            print(f"   ✅ Method 2 (dropdown items with {selector}): Found {len(portfolio_names)} portfolios")
                            return portfolio_names
                except Exception as e:
                    continue

            # If still nothing, look for elements containing "zg's address" or other common patterns
            all_text_elements = self.page.locator("text=/zg's address|large data|portfolio/i").all()
            if len(all_text_elements) > 0:
                for elem in all_text_elements:
                    text = elem.text_content()
                    if text:
                        text = text.strip()
                        if text and text not in ['Portfolio', 'Create Portfolio', 'Settings', 'Logout']:
                            portfolio_names.append(text)

                if len(portfolio_names) > 0:
                    print(f"   ✅ Method 3 (text search): Found {len(portfolio_names)} portfolios")
                    return portfolio_names

            # If still nothing, check if we're on empty state
            if self.is_empty_state_displayed():
                print(f"   ℹ️  Empty state detected - no portfolios exist yet")
                return []

            # Last resort - print page content for debugging
            print(f"   ⚠️  Could not find portfolios. Current URL: {self.page.url}")
            print(f"   📄 Page title: {self.page.title()}")

            return []

        except Exception as e:
            print(f"   ❌ Error getting portfolio names: {e}")
            return []

    def is_success_message_displayed(self) -> bool:
        """Check if success message is displayed"""
        try:
            expect(self.success_message).to_be_visible(timeout=10000)
            return True
        except:
            return False

    def click_user_profile(self):
        """Click user profile icon"""
        self.user_profile_icon.click()

    def take_portfolio_screenshot(self, portfolio_name: str, page_type: str = "overview", view_button: str = None, filename: str = None, url_domain: str = "dam") -> str:
        """
        Take screenshots of the DAM portfolio page in sections with new naming convention:
        Format: {domain}_{portfolioname}_{page}_{view}_{timestamp}_p1.png

        Examples:
        - dam_Portfolio_1_Overview_Token_20251114_115630_p1.png
        - dam_Portfolio_1_Holding_Token_20251114_115630_p1.png

        Args:
            portfolio_name: Name of the portfolio being viewed (e.g., "Portfolio 1")
            page_type: "overview" or "holding" (default: "overview")
            view_button: "token", "chain", or "address" (optional)
            filename: Optional custom filename (without extension)
            url_domain: Domain source - "dam" (default: "dam")

        Returns:
            Path to first screenshot (top section)
        """
        from datetime import datetime
        import os

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        clean_name = portfolio_name.replace(' ', '_').replace('/', '_').replace("'", '')

        # Create dam folder if it doesn't exist
        dam_folder = "test-results/screenshots/dam"
        os.makedirs(dam_folder, exist_ok=True)

        # Build filename with new format: {domain}_{portfolioname}_{page}_{view}_{timestamp}_p1.png
        page_name = page_type.capitalize()  # "Overview" or "Holding"
        view_name = view_button.capitalize() if view_button else "Token"  # "Token", "Chain", or "Address"

        base_filename = f"{url_domain}_{clean_name}_{page_name}_{view_name}_{timestamp}"

        # Ensure at top of page
        self.page.evaluate("window.scrollTo(0, 0)")
        self.page.wait_for_timeout(1000)

        # Screenshot timeout to handle tall pages
        screenshot_timeout = 60000  # 60 seconds

        # Take screenshot of top section (displayed viewport)
        top_screenshot = f"{dam_folder}/{base_filename}_p1.png"
        self.page.screenshot(path=top_screenshot, timeout=screenshot_timeout)
        print(f"📸 Section p1 (Top): {top_screenshot}")

        # Get page dimensions
        viewport_height = self.page.viewport_size['height']
        total_height = self.page.evaluate("document.body.scrollHeight")

        print(f"📏 Viewport: {viewport_height}px, Total page: {total_height}px")

        # If page is taller than viewport, scroll down and capture sections (NO scroll back up)
        section_num = 1
        if total_height > viewport_height:
            section_num = 2
            current_scroll = viewport_height

            while current_scroll < total_height:
                # Scroll down by viewport height
                self.page.evaluate(f"window.scrollTo(0, {current_scroll})")
                self.page.wait_for_timeout(800)  # Wait for content to load

                # Take screenshot of this section
                section_screenshot = f"{dam_folder}/{base_filename}_p{section_num}.png"
                self.page.screenshot(path=section_screenshot, timeout=screenshot_timeout)
                print(f"📸 Section p{section_num}: {section_screenshot}")

                section_num += 1
                current_scroll += viewport_height

        print(f"✅ Captured {section_num} section(s) for portfolio '{portfolio_name}' (scrolled top to bottom only)")
        return top_screenshot

    def expect_on_portfolio_page(self):
        """Assert on portfolio page"""
        import re
        expect(self.page).to_have_url(re.compile(r".*portfolio.*"))

    def expect_portfolio_displayed(self, portfolio_name: str):
        """Assert portfolio is displayed"""
        expect(self.page.locator(f"text={portfolio_name}")).to_be_visible()

    # Holdings Tab Methods
    def click_holdings_tab(self):
        """Click Holdings tab"""
        holdings_tab = self.page.locator("button:has-text('Holdings'), a:has-text('Holdings'), [role='tab']:has-text('Holdings')")
        holdings_tab.click()
        self.page.wait_for_timeout(1000)  # Wait for tab content to load

    def click_overview_tab(self):
        """Click Overview tab"""
        overview_tab = self.page.locator("button:has-text('Overview'), a:has-text('Overview'), [role='tab']:has-text('Overview')")
        overview_tab.click()
        self.page.wait_for_timeout(1000)

    def click_transactions_tab(self):
        """Click Transactions tab"""
        transactions_tab = self.page.locator("button:has-text('Transactions'), a:has-text('Transactions'), [role='tab']:has-text('Transactions')")
        transactions_tab.click()
        self.page.wait_for_timeout(1000)

    def get_holdings_data(self) -> list:
        """
        Get all holdings from the Holdings tab
        Returns: List of dicts with holding information
        Example: [{'token': 'ETH', 'amount': '0.5', 'value': '$1,500.00'}, ...]
        """
        holdings = []

        # Wait for holdings table to load
        try:
            # Look for table rows in Holdings tab
            table_rows = self.page.locator("table tbody tr, [class*='holding-row'], [data-testid*='holding']")

            if table_rows.count() > 0:
                for i in range(table_rows.count()):
                    row = table_rows.nth(i)

                    # Extract token name (could be in different columns)
                    token_name = row.locator("td:nth-child(1), [class*='token-name'], [data-testid*='token']").first.text_content()

                    # Extract amount
                    amount = row.locator("td:nth-child(2), [class*='amount'], [data-testid*='amount']").first.text_content()

                    # Extract value (optional)
                    value = ""
                    try:
                        value = row.locator("td:nth-child(3), [class*='value'], [data-testid*='value']").first.text_content()
                    except:
                        pass

                    holdings.append({
                        'token': token_name.strip() if token_name else "",
                        'amount': amount.strip() if amount else "",
                        'value': value.strip() if value else ""
                    })

            return holdings
        except Exception as e:
            print(f"⚠️ Could not extract holdings data: {e}")
            return []

    def verify_token_exists(self, token_name: str) -> bool:
        """
        Verify a specific token exists in holdings
        Args:
            token_name: Token symbol (e.g., 'ETH', 'BTC', 'USDT')
        Returns: True if token found, False otherwise
        """
        try:
            token_locator = self.page.locator(f"text={token_name}").first
            return token_locator.is_visible(timeout=5000)
        except:
            return False

    def get_token_amount(self, token_name: str) -> str:
        """
        Get the amount of a specific token
        Args:
            token_name: Token symbol
        Returns: Token amount as string, or empty string if not found
        """
        holdings = self.get_holdings_data()
        for holding in holdings:
            if holding['token'] == token_name:
                return holding['amount']
        return ""

    def get_addresses_from_net_worth_section(self) -> list:
        """
        Get all addresses from 'Net Worth per Address' section
        Returns: List of dicts with address information
        Example: [{'chain': 'ETH', 'address': '0x123...', 'value': '$1,000'}, ...]
        """
        addresses = []

        try:
            # Look for address elements in Net Worth per Address section
            address_rows = self.page.locator("[class*='address-row'], [data-testid*='address']")

            if address_rows.count() > 0:
                for i in range(address_rows.count()):
                    row = address_rows.nth(i)

                    # Extract chain (ETH, BSC, etc.)
                    chain = row.locator("[class*='chain'], [data-testid*='chain']").first.text_content()

                    # Extract address
                    address = row.locator("[class*='address'], [data-testid*='address-value']").first.text_content()

                    # Extract value (optional)
                    value = ""
                    try:
                        value = row.locator("[class*='value'], [data-testid*='net-worth']").first.text_content()
                    except:
                        pass

                    addresses.append({
                        'chain': chain.strip() if chain else "",
                        'address': address.strip() if address else "",
                        'value': value.strip() if value else ""
                    })

            return addresses
        except Exception as e:
            print(f"⚠️ Could not extract addresses: {e}")
            return []

    def print_holdings_summary(self):
        """Print formatted summary of all holdings"""
        print("\n" + "="*60)
        print("📊 HOLDINGS SUMMARY")
        print("="*60)

        holdings = self.get_holdings_data()
        if holdings:
            for holding in holdings:
                print(f"Token: {holding['token']:<10} | Amount: {holding['amount']:<15} | Value: {holding['value']}")
        else:
            print("No holdings found")

        print("="*60 + "\n")

    def print_addresses_summary(self):
        """Print formatted summary of all addresses"""
        print("\n" + "="*60)
        print("📍 ADDRESSES SUMMARY")
        print("="*60)

        addresses = self.get_addresses_from_net_worth_section()
        if addresses:
            for addr in addresses:
                print(f"Chain: {addr['chain']:<5} | Address: {addr['address']:<45} | Value: {addr['value']}")
        else:
            print("No addresses found")

        print("="*60 + "\n")

    # Transactions Tab Methods
    def get_transaction_filter_period(self) -> str:
        """
        Get currently selected transaction filter period
        Returns: Filter period string (e.g., 'Last 3 months', 'Last 6 months', 'Last year')
        """
        try:
            # Look for selected filter button/dropdown
            filter_selector = self.page.locator("[class*='filter-selected'], [aria-selected='true'], select option[selected]")
            if filter_selector.is_visible(timeout=3000):
                return filter_selector.text_content().strip()
            return ""
        except:
            return ""

    def select_transaction_filter(self, period: str):
        """
        Select transaction filter period
        Args:
            period: Filter period (e.g., 'Last 3 months', 'Last 6 months', 'Last year', 'All time')
        """
        try:
            # Look for filter dropdown or button
            filter_button = self.page.locator(f"button:has-text('{period}'), option:has-text('{period}')")
            filter_button.click()
            self.page.wait_for_timeout(1000)  # Wait for transactions to reload
        except Exception as e:
            print(f"⚠️ Could not select filter '{period}': {e}")

    def get_transactions_data(self) -> list:
        """
        Get all transactions from the Transactions tab
        Returns: List of dicts with transaction information
        Example: [
            {
                'chain': 'ETH',  # ETH or BNB
                'type': 'Receive',  # Send/Receive
                'address': '0x123...',
                'amount': '1.5 ETH',
                'value': '$3,000.00',
                'time': '2025-10-27 14:30:00',
                'hash': '0xabc...'
            }, ...
        ]

        Note: Transactions are sorted by time (most recent first)
        If an address has both ETH and BNB transactions, both chains will be displayed
        """
        transactions = []

        try:
            # Look for transaction rows in Transactions tab
            tx_rows = self.page.locator("table tbody tr, [class*='transaction-row'], [data-testid*='transaction']")

            if tx_rows.count() > 0:
                for i in range(tx_rows.count()):
                    row = tx_rows.nth(i)

                    # Extract chain (ETH or BNB/BSC)
                    chain = ""
                    try:
                        chain = row.locator("[class*='chain'], [data-testid*='chain'], td:nth-child(1)").first.text_content()
                    except:
                        pass

                    # Extract transaction type (Send/Receive)
                    tx_type = ""
                    try:
                        tx_type = row.locator("[class*='type'], [data-testid*='type'], td:nth-child(2)").first.text_content()
                    except:
                        pass

                    # Extract address
                    address = ""
                    try:
                        address = row.locator("[class*='address'], [data-testid*='address'], td:nth-child(3)").first.text_content()
                    except:
                        pass

                    # Extract amount (with token symbol)
                    amount = ""
                    try:
                        amount = row.locator("[class*='amount'], [data-testid*='amount'], td:nth-child(4)").first.text_content()
                    except:
                        pass

                    # Extract value (USD equivalent)
                    value = ""
                    try:
                        value = row.locator("[class*='value'], [data-testid*='value'], td:nth-child(5)").first.text_content()
                    except:
                        pass

                    # Extract timestamp
                    time = ""
                    try:
                        time = row.locator("[class*='time'], [data-testid*='time'], td:nth-child(6)").first.text_content()
                    except:
                        pass

                    # Extract transaction hash
                    tx_hash = ""
                    try:
                        tx_hash = row.locator("[class*='hash'], [data-testid*='hash'], a[href*='etherscan'], a[href*='bscscan']").first.text_content()
                        # Alternative: get from href attribute
                        if not tx_hash:
                            hash_link = row.locator("a[href*='etherscan'], a[href*='bscscan']").first
                            href = hash_link.get_attribute('href')
                            if href:
                                # Extract hash from URL
                                tx_hash = href.split('/')[-1]
                    except:
                        pass

                    transactions.append({
                        'chain': chain.strip() if chain else "",
                        'type': tx_type.strip() if tx_type else "",
                        'address': address.strip() if address else "",
                        'amount': amount.strip() if amount else "",
                        'value': value.strip() if value else "",
                        'time': time.strip() if time else "",
                        'hash': tx_hash.strip() if tx_hash else ""
                    })

            return transactions
        except Exception as e:
            print(f"⚠️ Could not extract transactions data: {e}")
            return []

    def verify_transaction_exists(self, tx_hash: str) -> bool:
        """
        Verify a specific transaction exists in the list
        Args:
            tx_hash: Transaction hash to search for
        Returns: True if transaction found, False otherwise
        """
        try:
            tx_locator = self.page.locator(f"text={tx_hash}").first
            return tx_locator.is_visible(timeout=5000)
        except:
            return False

    def get_transactions_by_chain(self, chain: str) -> list:
        """
        Get all transactions for a specific blockchain
        Args:
            chain: Chain name ('ETH' or 'BNB'/'BSC')
        Returns: List of transactions for that chain
        """
        all_transactions = self.get_transactions_data()
        return [tx for tx in all_transactions if tx['chain'].upper() == chain.upper()]

    def get_transactions_by_address(self, address: str) -> list:
        """
        Get all transactions for a specific address
        Args:
            address: Wallet address (can be partial)
        Returns: List of transactions involving that address
        """
        all_transactions = self.get_transactions_data()
        return [tx for tx in all_transactions if address.lower() in tx['address'].lower()]

    def print_transactions_summary(self, limit: int = 10):
        """
        Print formatted summary of transactions
        Args:
            limit: Maximum number of transactions to display (default: 10)
        """
        print("\n" + "="*100)
        print("📋 TRANSACTIONS SUMMARY (sorted by time, most recent first)")
        print("="*100)

        transactions = self.get_transactions_data()
        if transactions:
            display_count = min(len(transactions), limit)
            for i, tx in enumerate(transactions[:display_count]):
                print(f"{i+1}. [{tx['chain']}] {tx['type']:<8} | "
                      f"Amount: {tx['amount']:<15} | "
                      f"Value: {tx['value']:<12} | "
                      f"Time: {tx['time']}")
                if tx['hash']:
                    print(f"   Hash: {tx['hash']}")
                if tx['address']:
                    print(f"   Address: {tx['address']}")
                print()

            if len(transactions) > limit:
                print(f"... and {len(transactions) - limit} more transactions")
        else:
            print("No transactions found")

        print("="*100 + "\n")

    def print_transactions_by_chain(self, chain: str):
        """
        Print transactions for a specific blockchain
        Args:
            chain: Chain name ('ETH' or 'BNB')
        """
        print(f"\n{'='*80}")
        print(f"📋 {chain.upper()} TRANSACTIONS")
        print(f"{'='*80}")

        chain_txs = self.get_transactions_by_chain(chain)
        if chain_txs:
            for i, tx in enumerate(chain_txs):
                print(f"{i+1}. {tx['type']:<8} | "
                      f"Amount: {tx['amount']:<15} | "
                      f"Time: {tx['time']}")
        else:
            print(f"No {chain.upper()} transactions found")

        print(f"{'='*80}\n")

    def click_download_button_overview_token(self):
        """
        Click the download icon button in Overview - Token page
        Then click the Export button in the modal that appears

        This button downloads the token breakdown table data as CSV/Excel
        """
        print("📥 Clicking download button in Overview - Token page...")

        try:
            # Wait a moment for page to be ready
            self.page.wait_for_timeout(1000)

            # Look for the download button specifically in the Token Breakdown section
            # The button should be near "Show tokens <1% value" toggle
            download_button_selectors = [
                # Button with aria-label="Export token breakdown"
                "button[aria-label='Export token breakdown']",
                # Button with download icon near "Token Breakdown" heading
                "button[aria-haspopup='dialog']:has(svg.size-5.text-mono-500)",
                "button[class*='rounded-sm'][class*='typography-body']:has(svg)",
                # More generic selectors
                "div:has-text('Token Breakdown') ~ div button:has(svg)",
            ]

            download_clicked = False

            for selector in download_button_selectors:
                try:
                    download_button = self.page.locator(selector).first
                    if download_button.is_visible(timeout=2000):
                        print(f"   ✅ Found download button: {selector}")
                        download_button.click()
                        self.page.wait_for_timeout(1500)  # Wait for modal to appear
                        print(f"   ✅ Download button clicked successfully")
                        download_clicked = True
                        break
                except Exception as e:
                    continue

            if not download_clicked:
                print(f"   ⚠️  Download button not found with specific selectors - trying to find by SVG near Token Breakdown...")
                # Try to find the download button by looking for SVG icon buttons in the Token Breakdown area
                try:
                    # Get all buttons with SVG that have aria-haspopup="dialog"
                    dialog_buttons = self.page.locator("button[aria-haspopup='dialog']:has(svg)").all()
                    if len(dialog_buttons) > 0:
                        print(f"   🔍 Found {len(dialog_buttons)} dialog buttons with SVG")
                        # Try the first one (should be the download button for Token Breakdown)
                        dialog_buttons[0].click()
                        self.page.wait_for_timeout(1500)
                        print(f"   ✅ Clicked dialog button (Token Breakdown download)")
                        download_clicked = True
                except Exception as e:
                    print(f"   ⚠️  Alternative approach failed: {e}")

            if not download_clicked:
                print(f"   ❌ Could not find or click download button")
                return False

            # Now click the Export button in the modal
            print(f"   🔍 Looking for Export button in modal...")
            try:
                # Wait for modal to be visible
                self.page.wait_for_timeout(1000)

                # Try multiple selectors for the Export button
                export_button_selectors = [
                    "button:has-text('Export')",
                    "button[type='submit']:has-text('Export')",
                    "[role='dialog'] button:has-text('Export')",
                    ".modal button:has-text('Export')",
                ]

                export_clicked = False
                for selector in export_button_selectors:
                    try:
                        export_button = self.page.locator(selector).first
                        if export_button.is_visible(timeout=3000):
                            print(f"   ✅ Found Export button: {selector}")

                            # Check if button is enabled
                            if export_button.is_enabled():
                                export_button.click()
                                self.page.wait_for_timeout(2000)  # Wait for download to start
                                print(f"   ✅ Export button clicked - download started")
                                export_clicked = True
                                break
                            else:
                                print(f"   ⚠️  Export button is disabled - closing modal and skipping export")
                                # Close the modal by pressing Escape
                                self.page.keyboard.press("Escape")
                                self.page.wait_for_timeout(500)
                                return False
                    except:
                        continue

                if not export_clicked:
                    print(f"   ⚠️  Export button not found in modal")
                    return False

            except Exception as e:
                print(f"   ⚠️  Error clicking Export button: {e}")
                return False

            # Check for quota error message and close modal if present
            try:
                quota_error = self.page.locator("text=/hit your daily export quota/i").first
                if quota_error.is_visible(timeout=2000):
                    print(f"   ⚠️  Daily export quota reached - closing modal")
                    # Click outside the modal to close it (click on backdrop)
                    self.page.keyboard.press("Escape")
                    self.page.wait_for_timeout(500)
            except:
                pass

            return True

        except Exception as e:
            print(f"   ❌ Error clicking download button: {e}")
            import traceback
            traceback.print_exc()
            return False

    def click_download_button_overview_chain(self):
        """
        Click the download icon button in Overview - Chain page
        Then click the Export button in the modal that appears

        This button downloads the chain breakdown table data as CSV/Excel
        """
        print("📥 Clicking download button in Overview - Chain page...")

        try:
            # Wait a moment for page to be ready
            self.page.wait_for_timeout(1000)

            # Look for the download button specifically in the Chain Breakdown section
            download_button_selectors = [
                # Button with aria-label="Export chain breakdown"
                "button[aria-label='Export chain breakdown']",
                # Button with download icon near "Chain Breakdown" or chain table
                "button[aria-haspopup='dialog']:has(svg.size-5.text-mono-500)",
                "button[class*='rounded-sm'][class*='typography-body']:has(svg)",
                # More generic selectors for chain section
                "div:has-text('Chain Breakdown') ~ div button:has(svg)",
            ]

            download_clicked = False

            for selector in download_button_selectors:
                try:
                    download_button = self.page.locator(selector).first
                    if download_button.is_visible(timeout=2000):
                        print(f"   ✅ Found download button: {selector}")
                        download_button.click()
                        self.page.wait_for_timeout(1500)  # Wait for modal to appear
                        print(f"   ✅ Download button clicked successfully")
                        download_clicked = True
                        break
                except Exception as e:
                    continue

            if not download_clicked:
                print(f"   ⚠️  Download button not found with specific selectors - trying to find by SVG near Chain section...")
                # Try to find the download button by looking for SVG icon buttons
                try:
                    # Get all buttons with SVG that have aria-haspopup="dialog"
                    dialog_buttons = self.page.locator("button[aria-haspopup='dialog']:has(svg)").all()
                    if len(dialog_buttons) > 0:
                        print(f"   🔍 Found {len(dialog_buttons)} dialog buttons with SVG")
                        # Try the first one (should be the download button)
                        dialog_buttons[0].click()
                        self.page.wait_for_timeout(1500)
                        print(f"   ✅ Clicked dialog button (Chain Breakdown download)")
                        download_clicked = True
                except Exception as e:
                    print(f"   ⚠️  Alternative approach failed: {e}")

            if not download_clicked:
                print(f"   ❌ Could not find or click download button")
                return False

            # Now click the Export button in the modal
            print(f"   🔍 Looking for Export button in modal...")
            try:
                # Wait for modal to be visible
                self.page.wait_for_timeout(1000)

                # Try multiple selectors for the Export button
                export_button_selectors = [
                    "button:has-text('Export')",
                    "button[type='submit']:has-text('Export')",
                    "[role='dialog'] button:has-text('Export')",
                    ".modal button:has-text('Export')",
                ]

                export_clicked = False
                for selector in export_button_selectors:
                    try:
                        export_button = self.page.locator(selector).first
                        if export_button.is_visible(timeout=3000):
                            print(f"   ✅ Found Export button: {selector}")

                            # Check if button is enabled
                            if export_button.is_enabled():
                                export_button.click()
                                self.page.wait_for_timeout(2000)  # Wait for download to start
                                print(f"   ✅ Export button clicked - download started")
                                export_clicked = True
                                break
                            else:
                                print(f"   ⚠️  Export button is disabled - closing modal and skipping export")
                                # Close the modal by pressing Escape
                                self.page.keyboard.press("Escape")
                                self.page.wait_for_timeout(500)
                                return False
                    except:
                        continue

                if not export_clicked:
                    print(f"   ⚠️  Export button not found in modal")
                    return False

            except Exception as e:
                print(f"   ⚠️  Error clicking Export button: {e}")
                return False

            # Check for quota error message and close modal if present
            try:
                quota_error = self.page.locator("text=/hit your daily export quota/i").first
                if quota_error.is_visible(timeout=2000):
                    print(f"   ⚠️  Daily export quota reached - closing modal")
                    # Click outside the modal to close it (click on backdrop)
                    self.page.keyboard.press("Escape")
                    self.page.wait_for_timeout(500)
            except:
                pass

            return True

        except Exception as e:
            print(f"   ❌ Error clicking download button: {e}")
            import traceback
            traceback.print_exc()
            return False

    def click_download_button_overview_address(self):
        """
        Click the download icon button in Overview - Address page
        Then click the Export button in the modal that appears

        This button downloads the address breakdown table data as CSV/Excel
        """
        print("📥 Clicking download button in Overview - Address page...")

        try:
            # Wait a moment for page to be ready
            self.page.wait_for_timeout(1000)

            # Look for the download button specifically in the Address Breakdown section
            download_button_selectors = [
                # Button with aria-label="Export address breakdown"
                "button[aria-label='Export address breakdown']",
                # Button with download icon near "Address Breakdown" or address table
                "button[aria-haspopup='dialog']:has(svg.size-5.text-mono-500)",
                "button[class*='rounded-sm'][class*='typography-body']:has(svg)",
                # More generic selectors for address section
                "div:has-text('Address Breakdown') ~ div button:has(svg)",
            ]

            download_clicked = False

            for selector in download_button_selectors:
                try:
                    download_button = self.page.locator(selector).first
                    if download_button.is_visible(timeout=2000):
                        print(f"   ✅ Found download button: {selector}")
                        download_button.click()
                        self.page.wait_for_timeout(1500)  # Wait for modal to appear
                        print(f"   ✅ Download button clicked successfully")
                        download_clicked = True
                        break
                except Exception as e:
                    continue

            if not download_clicked:
                print(f"   ⚠️  Download button not found with specific selectors - trying to find by SVG near Address section...")
                # Try to find the download button by looking for SVG icon buttons
                try:
                    # Get all buttons with SVG that have aria-haspopup="dialog"
                    dialog_buttons = self.page.locator("button[aria-haspopup='dialog']:has(svg)").all()
                    if len(dialog_buttons) > 0:
                        print(f"   🔍 Found {len(dialog_buttons)} dialog buttons with SVG")
                        # Try the first one (should be the download button)
                        dialog_buttons[0].click()
                        self.page.wait_for_timeout(1500)
                        print(f"   ✅ Clicked dialog button (Address Breakdown download)")
                        download_clicked = True
                except Exception as e:
                    print(f"   ⚠️  Alternative approach failed: {e}")

            if not download_clicked:
                print(f"   ❌ Could not find or click download button")
                return False

            # Now click the Export button in the modal
            print(f"   🔍 Looking for Export button in modal...")
            try:
                # Wait for modal to be visible
                self.page.wait_for_timeout(1000)

                # Try multiple selectors for the Export button
                export_button_selectors = [
                    "button:has-text('Export')",
                    "button[type='submit']:has-text('Export')",
                    "[role='dialog'] button:has-text('Export')",
                    ".modal button:has-text('Export')",
                ]

                export_clicked = False
                for selector in export_button_selectors:
                    try:
                        export_button = self.page.locator(selector).first
                        if export_button.is_visible(timeout=3000):
                            print(f"   ✅ Found Export button: {selector}")

                            # Check if button is enabled
                            if export_button.is_enabled():
                                export_button.click()
                                self.page.wait_for_timeout(2000)  # Wait for download to start
                                print(f"   ✅ Export button clicked - download started")
                                export_clicked = True
                                break
                            else:
                                print(f"   ⚠️  Export button is disabled - closing modal and skipping export")
                                # Close the modal by pressing Escape
                                self.page.keyboard.press("Escape")
                                self.page.wait_for_timeout(500)
                                return False
                    except:
                        continue

                if not export_clicked:
                    print(f"   ⚠️  Export button not found in modal")
                    return False

            except Exception as e:
                print(f"   ⚠️  Error clicking Export button: {e}")
                return False

            # Check for quota error message and close modal if present
            try:
                quota_error = self.page.locator("text=/hit your daily export quota/i").first
                if quota_error.is_visible(timeout=2000):
                    print(f"   ⚠️  Daily export quota reached - closing modal")
                    # Click outside the modal to close it (click on backdrop)
                    self.page.keyboard.press("Escape")
                    self.page.wait_for_timeout(500)
            except:
                pass

            return True

        except Exception as e:
            print(f"   ❌ Error clicking download button: {e}")
            import traceback
            traceback.print_exc()
            return False
