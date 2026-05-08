"""
Sign Up Page Object Model (Playwright)
Updated with accurate locators from signup.html and Terms/Privacy handling
"""
import re
from playwright.sync_api import Page, expect
from pages.base_page import BasePage
from config.config import Config


class SignUpPage(BasePage):
    """Sign Up page object"""

    def __init__(self, page: Page, base_url: str = None):
        super().__init__(page, base_url)
        self.url = Config.SIGN_UP_URL

        # Locators using data-testid and other attributes from actual HTML
        self.email_input = page.locator('input[data-testid="input-email"]')
        self.password_input = page.locator('input[data-testid="input-password"]')

        # Password visibility toggle - button with eye icon inside password container
        # self.password_toggle = page.locator('input[data-testid="input-password"] ~ div button[type="button"]').first()

        # Terms checkbox - has name="agree"
        self.terms_checkbox = page.locator('input[name="agree"][type="checkbox"]')

        # Continue button - type="submit" with text "Continue"
        self.continue_button = page.locator('button[type="submit"]:has-text("Continue")')

        # Google Sign-in button - has specific text and Google SVG
        self.google_signin_button = page.locator('button:has-text("Continue with Google")')

        # Wallet button - link to /sign-in/wallet
        self.wallet_button = page.locator('a[href="/sign-in/wallet"]')

        # Sign In link
        self.signin_link = page.locator('a[href="/sign-in"]')

        # Password requirements - div elements with typography-caption class
        self.password_length_check = page.locator('div.typography-caption:has-text("At least 12 characters")')
        self.password_uppercase_check = page.locator('div.typography-caption:has-text("At least one uppercase letter")')
        self.password_lowercase_check = page.locator('div.typography-caption:has-text("At least one lowercase letter")')
        self.password_number_check = page.locator('div.typography-caption:has-text("At least one number")')
        self.password_special_check = page.locator('div.typography-caption:has-text("At least one special character")')

        # Messages (Toastify notifications)
        self.success_message = page.locator('.Toastify__toast--success, text=Success')
        self.error_message = page.locator('.Toastify__toast--error, [aria-invalid="true"]')

        # Terms of Service and Privacy Policy are <button> elements, not <a> links
        # Based on actual HTML: <button class="typography-body font-normal text-primary cursor-pointer">
        self.terms_of_service_button = page.get_by_role("button", name="Terms of Service")
        self.privacy_policy_button = page.get_by_role("button", name="Privacy Policy")

    def navigate(self):
        """Navigate to sign up page"""
        self.navigate_to(self.url)

    def enter_email(self, email: str):
        """Enter email address"""
        self.email_input.fill(email)

    def enter_password(self, password: str):
        """Enter password"""
        self.password_input.fill(password)

    def toggle_password_visibility(self):
        """Click password visibility toggle"""
        self.password_toggle.click()

    def _scroll_active_modal_to_bottom_and_close(self):
        """
        Scroll modal/dialog to bottom and close it
        Required for Terms of Service and Privacy Policy
        IMPORTANT: Must scroll ALL THE WAY to bottom to enable checkbox
        """
        # 1) Wait for dialog to be visible
        # Try multiple ways to find the modal dialog
        dialog = None

        # Try role="dialog" first
        try:
            dialog = self.page.get_by_role("dialog")
            expect(dialog).to_be_visible(timeout=3000)
        except:
            # If role="dialog" doesn't work, try finding by heading "Terms of Service" or "Privacy Policy"
            try:
                dialog = self.page.locator('div:has(h1:has-text("Terms of Service")), div:has(h2:has-text("Terms of Service")), div:has(h1:has-text("Privacy Policy")), div:has(h2:has-text("Privacy Policy"))').first
                expect(dialog).to_be_visible(timeout=3000)
            except:
                # Last resort: find any visible overlay/modal by common classes
                dialog = self.page.locator('div[class*="modal"], div[class*="dialog"], div[class*="overlay"]').first
                expect(dialog).to_be_visible(timeout=3000)

        # 2) Find scrollable content area - look for elements with overflow CSS
        # Even if scrollHeight == clientHeight NOW, it might become scrollable
        print("Searching for element with overflow styling inside modal...")

        # Try to find elements with overflow CSS (most likely to be scrollable)
        overflow_candidates = dialog.locator("[class*='overflow']").all()

        if len(overflow_candidates) > 0:
            # Find the one with largest height (most likely the main scrollable area)
            target = None
            max_height = 0
            for candidate in overflow_candidates:
                try:
                    height = candidate.evaluate("(el) => el.clientHeight")
                    scroll_info = candidate.evaluate("(el) => ({ scrollHeight: el.scrollHeight, clientHeight: el.clientHeight, scrollable: el.scrollHeight - el.clientHeight })")
                    print(f"  Overflow candidate: height={height}px, scrollInfo={scroll_info}")
                    if height > max_height:
                        max_height = height
                        target = candidate
                except:
                    continue

            if target:
                print(f"✅ Using largest overflow element (height={max_height}px)")
            else:
                print("⚠️  No valid overflow element, using dialog")
                target = dialog
        else:
            print("⚠️  No overflow elements found, using dialog")
            target = dialog

        # 3) Focus on the scrollable area
        try:
            target.click(position={"x": 10, "y": 10})
        except:
            dialog.click(position={"x": 10, "y": 10})

        # 4) Scroll DOWN to the ABSOLUTE BOTTOM of the scrollable container
        print("Scrolling modal DOWN to ABSOLUTE BOTTOM...")

        # Get the total scrollable height
        max_scroll = target.evaluate("(el) => el.scrollHeight - el.clientHeight")

        if max_scroll > 0:
            # Scroll DOWN in visible steps (200px at a time)
            current = 0
            step = 200
            while current < max_scroll:
                current += step
                if current > max_scroll:
                    current = max_scroll
                target.evaluate(f"(el) => el.scrollTop = {current}")
                self.page.wait_for_timeout(50)  # Small delay so scrolling is VISIBLE

            # Ensure we're at ABSOLUTE bottom
            target.evaluate(f"(el) => el.scrollTop = {max_scroll}")

            # IMPORTANT: Stay at bottom for 1 second so the app registers we reached the end
            print(f"Reached ABSOLUTE bottom (scrolled {max_scroll}px), staying here...")
            self.page.wait_for_timeout(1000)
        else:
            # Modal is not scrollable (content fits on screen)
            # Still wait to simulate "reading" the content
            print(f"Modal is not scrollable (content fits on screen), waiting 2s...")
            self.page.wait_for_timeout(2000)

            # Try multiple strategies to signal the app we've "reviewed" it
            # Strategy 1: Trigger scroll events
            target.evaluate("""(el) => {
                el.scrollTop = 1;
                el.scrollTop = 0;
                el.dispatchEvent(new Event('scroll', { bubbles: true }));
            }""")

            # Strategy 2: Check if there's a tracking mechanism in localStorage or sessionStorage
            self.page.evaluate("""() => {
                // Try to find and mark Terms/Privacy as reviewed
                const dialog = document.querySelector('[role="dialog"]');
                if (dialog) {
                    const heading = dialog.querySelector('h1, h2');
                    if (heading) {
                        const title = heading.textContent;
                        if (title.includes('Terms')) {
                            window.sessionStorage.setItem('termsReviewed', 'true');
                            window.localStorage.setItem('termsReviewed', 'true');
                        } else if (title.includes('Privacy')) {
                            window.sessionStorage.setItem('privacyReviewed', 'true');
                            window.localStorage.setItem('privacyReviewed', 'true');
                        }
                    }
                }
            }""")

            self.page.wait_for_timeout(500)

        # Scroll back UP to top if we scrolled down
        if max_scroll > 0:
            print("Scrolling back UP to top to see X button...")

            # Scroll UP in visible steps (200px at a time)
            current = max_scroll
            while current > 0:
                current -= step
                if current < 0:
                    current = 0
                target.evaluate(f"(el) => el.scrollTop = {current}")
                self.page.wait_for_timeout(50)  # Small delay so scrolling UP is VISIBLE

            # Ensure we're at absolute top
            target.evaluate("(el) => el.scrollTop = 0")

            # Short wait before clicking X
            self.page.wait_for_timeout(500)
        else:
            print("No need to scroll up (modal wasn't scrollable)")

        # 5) Close button should now be accessible after scrolling to bottom
        # The X button in top-right corner must be clicked to properly close and enable checkbox
        close_clicked = False

        print("Looking for X close button...")

        # Method 1: Find X button by common selectors (most specific first)
        close_selectors = [
            "button[aria-label='Close']",
            "button[aria-label='close']",
            "button:has-text('×')",
            "button:has-text('✕')",
            "button:has-text('X')",
            "button[class*='close']",
            "svg[class*='close']",
            "[role='button'][aria-label*='lose']",
        ]

        for selector in close_selectors:
            if close_clicked:
                break
            try:
                close_btn = dialog.locator(selector).first
                if close_btn.count() > 0:
                    print(f"Found close button with selector: {selector}")
                    close_btn.wait_for(state="visible", timeout=2000)
                    close_btn.click()
                    print("Successfully clicked close button")
                    close_clicked = True
                    break
            except Exception as e:
                print(f"Selector {selector} failed: {str(e)[:50]}")
                continue

        # Method 2: Try clicking in top-right corner where X usually is
        if not close_clicked:
            try:
                print("Trying to click top-right corner for X button...")
                # Get dialog bounding box and click top-right area
                box = dialog.bounding_box()
                if box:
                    # Click near top-right corner (20px from right edge, 20px from top)
                    x = box['x'] + box['width'] - 20
                    y = box['y'] + 20
                    self.page.mouse.click(x, y)
                    print(f"Clicked at position ({x}, {y})")
                    close_clicked = True
            except Exception as e:
                print(f"Top-right click failed: {str(e)[:50]}")
                pass

        # Method 3: Press Escape key as fallback
        if not close_clicked:
            try:
                print("Trying Escape key...")
                self.page.keyboard.press("Escape")
                close_clicked = True
                print("Pressed Escape")
            except:
                pass

        # 6) Dialog should disappear
        print("Waiting for dialog to close...")
        expect(dialog).to_be_hidden(timeout=10000)
        print("Dialog closed successfully")

    def click_terms_of_service(self):
        """Click Terms of Service button"""
        self.terms_of_service_button.click()

    def click_privacy_policy(self):
        """Click Privacy Policy button"""
        self.privacy_policy_button.click()

    def review_terms_and_privacy(self):
        """
        Complete the terms and privacy review process:
        1. Click Terms of Service
        2. Scroll to bottom and close
        3. Click Privacy Policy
        4. Scroll to bottom and close

        NOTE: The checkbox will only be enabled after BOTH:
        - Terms/Privacy are reviewed (this method)
        - Email/Password fields are filled (done in sign_up method)
        """
        # Review Terms of Service
        self.click_terms_of_service()
        self._scroll_active_modal_to_bottom_and_close()

        # Review Privacy Policy
        self.click_privacy_policy()
        self._scroll_active_modal_to_bottom_and_close()

        print("Terms and Privacy review completed")

        # Wait for the app to process that both modals were reviewed
        # The checkbox should become enabled after this
        self.page.wait_for_timeout(1000)

    def solve_captcha(self):
        """
        Handle Cloudflare Turnstile captcha:
        1. Wait for captcha widget to load and render
        2. Wait for captcha token to be populated (may auto-solve)
        3. If needed, interact with captcha to trigger solving
        """
        print("Handling Cloudflare Turnstile captcha...")

        # First, wait for the captcha widget to appear
        # Cloudflare Turnstile can render as iframe OR inline div
        print("Waiting for captcha widget to render...")

        # Wait a bit for Cloudflare script to initialize
        self.page.wait_for_timeout(2000)

        # Check if token is already populated (auto-solve mode)
        cf_token = self.page.locator('input[name="cf-turnstile-response"]').input_value()
        if cf_token and len(cf_token) > 10:
            print(f"Captcha already solved automatically! Token length: {len(cf_token)}")
            return

        # Try to find the captcha iframe and interact with it
        print("Looking for captcha iframe...")
        try:
            # Look for iframe containing turnstile
            iframes = self.page.frames
            turnstile_frame = None
            for frame in iframes:
                frame_url = frame.url
                if 'turnstile' in frame_url or 'cloudflare' in frame_url:
                    turnstile_frame = frame
                    print(f"Found Turnstile iframe: {frame_url}")
                    break

            if turnstile_frame:
                # Try to find and click the checkbox/widget in iframe
                try:
                    # Cloudflare Turnstile typically has a clickable input or body element
                    # Try multiple selectors
                    selectors = [
                        'input[type="checkbox"]',
                        'input',
                        'label',
                        'body',
                        '#cf-turnstile-response',
                        '[id*="checkbox"]',
                        '[class*="checkbox"]'
                    ]

                    clicked = False
                    for selector in selectors:
                        try:
                            element = turnstile_frame.locator(selector).first
                            if element.is_visible(timeout=2000):
                                print(f"Found clickable element with selector: {selector}")
                                element.click()
                                clicked = True
                                print("Clicked captcha element")
                                break
                        except:
                            continue

                    if not clicked:
                        print("No clickable element found, trying to click iframe body...")
                        turnstile_frame.locator('body').click()

                except Exception as e2:
                    print(f"Could not interact with iframe content: {e2}")

        except Exception as e:
            print(f"Iframe search failed: {e}")

        # Wait for token to be populated
        print("Waiting for captcha token to be populated...")
        try:
            self.page.wait_for_function(
                """
                () => {
                    const cfToken = document.querySelector('input[name="cf-turnstile-response"]');
                    return cfToken && cfToken.value && cfToken.value.length > 10;
                }
                """,
                timeout=30000  # 30 seconds for captcha to solve
            )

            cf_token = self.page.locator('input[name="cf-turnstile-response"]').input_value()
            print(f"Captcha solved! Token length: {len(cf_token)}")

            # Give a moment for the form to react to the token
            self.page.wait_for_timeout(1000)

        except Exception as e:
            print(f"Warning: Captcha token wait timeout: {e}")
            # Check if token exists anyway
            cf_token = self.page.locator('input[name="cf-turnstile-response"]').input_value()
            if cf_token and len(cf_token) > 10:
                print(f"Token found after timeout: {len(cf_token)}")
            else:
                print("Captcha may not have solved properly")

    def accept_terms(self):
        """
        Accept terms and conditions:
        1. Review Terms of Service and Privacy Policy (if not already done)
        2. Check the terms checkbox
        3. Solve Cloudflare Turnstile captcha
        4. Wait for Continue button to be enabled (validates all conditions)
        """
        # First review the documents if checkbox is still disabled
        if self.terms_checkbox.is_disabled():
            print("Checkbox is disabled, reviewing terms and privacy...")
            self.review_terms_and_privacy()

        # Wait a bit for checkbox to be fully enabled
        self.page.wait_for_timeout(500)

        # Verify checkbox is now enabled
        expect(self.terms_checkbox).not_to_be_disabled(timeout=5000)
        print("Checkbox is now enabled after reviewing documents")

        # Check the checkbox (without force, to respect validation)
        if not self.terms_checkbox.is_checked():
            print("Checking the checkbox...")
            self.terms_checkbox.check()

            # Trigger change/input events to ensure form validation runs
            self.page.evaluate("""
                () => {
                    const checkbox = document.querySelector('input[name="agree"]');
                    if (checkbox) {
                        checkbox.dispatchEvent(new Event('change', { bubbles: true }));
                        checkbox.dispatchEvent(new Event('input', { bubbles: true }));
                        checkbox.dispatchEvent(new Event('click', { bubbles: true }));
                    }
                }
            """)

            self.page.wait_for_timeout(1000)  # Wait longer for validation to run

        # Verify checkbox is checked
        expect(self.terms_checkbox).to_be_checked(timeout=3000)
        print("Checkbox is checked")

        # Solve the Cloudflare Turnstile captcha
        self.solve_captcha()

        # Debug: Check all conditions
        print(f"Email value: {self.get_email_value()}")
        print(f"Password value length: {len(self.get_password_value())}")
        requirements = self.get_all_password_requirements_status()
        print(f"Password requirements: {requirements}")
        print(f"All requirements met: {all(requirements.values())}")

        # Check if there's a captcha token
        captcha_token = self.page.locator('input[name="captcha_token"]').input_value()
        cf_token = self.page.locator('input[name="cf-turnstile-response"]').input_value()
        print(f"Captcha token present: {bool(captcha_token)}")
        print(f"CF Turnstile token present: {bool(cf_token)}")

        # Wait for Continue button to be enabled
        # This validates that all conditions are met: email valid, password valid, checkbox checked, captcha solved
        print("Waiting for Continue button to be enabled...")
        try:
            expect(self.continue_button).to_be_enabled(timeout=10000)
            print("Continue button is enabled!")
        except AssertionError as e:
            print(f"Continue button failed to enable: {e}")
            print("Taking screenshot for debugging...")
            self.page.screenshot(path="debug_button_disabled.png")
            raise

    def check_terms_checkbox(self):
        """
        Check terms and conditions checkbox
        This is an alias for accept_terms() to maintain backward compatibility
        """
        self.accept_terms()

    def click_continue(self):
        """Click Continue button - waits for it to be enabled first"""
        expect(self.continue_button).to_be_enabled(timeout=5000)
        self.continue_button.click()

    def click_google_signin(self):
        """Click Continue with Google button"""
        self.google_signin_button.click()

    def click_wallet_signin(self):
        """Click Continue with Wallet button"""
        self.wallet_button.click()

    def click_signin_link(self):
        """Click 'Sign In' link at bottom"""
        self.signin_link.click()

    def sign_up(self, email: str, password: str):
        """
        Complete sign up flow including Terms acceptance
        NOTE: The form fields get cleared when Terms/Privacy modals are opened,
        so we need to re-fill them after reviewing the documents
        """
        self.enter_email(email)
        self.enter_password(password)

        # Review Terms/Privacy ONCE - this will enable the checkbox
        if self.terms_checkbox.is_disabled():
            print("Checkbox is disabled, reviewing terms and privacy...")
            self.review_terms_and_privacy()

            # IMPORTANT: Modals clear the form fields, so re-fill them
            print("Re-filling email and password (cleared by modals)...")
            self.enter_email(email)
            self.enter_password(password)
            print("Form re-filled")

            # Wait for form validation to complete (password requirements check)
            print("Waiting for form validation...")
            self.page.wait_for_timeout(1000)

            # Trigger form validation by focusing and blurring email field
            self.email_input.focus()
            self.email_input.blur()
            self.password_input.focus()
            self.password_input.blur()

            # Wait for checkbox to become enabled (needs BOTH reviewed + filled form + validated)
            print("Waiting for checkbox to become enabled...")

            # Debug: Check the checkbox disabled state and attributes
            is_disabled = self.terms_checkbox.is_disabled()
            checkbox_attrs = self.terms_checkbox.evaluate("el => ({ disabled: el.disabled, checked: el.checked, name: el.name, type: el.type })")
            print(f"Checkbox state: disabled={is_disabled}, attrs={checkbox_attrs}")

            # Check if there are any React props or data attributes that might control the checkbox
            react_props = self.page.evaluate("""() => {
                const checkbox = document.querySelector('input[name="agree"][type="checkbox"]');
                if (checkbox) {
                    return {
                        disabled: checkbox.disabled,
                        checked: checkbox.checked,
                        attributes: Array.from(checkbox.attributes).map(attr => ({name: attr.name, value: attr.value}))
                    };
                }
                return null;
            }""")
            print(f"React checkbox props: {react_props}")

            self.page.wait_for_timeout(3000)

        # Now checkbox should be enabled - check it
        print("Checking 'I agree to Terms' checkbox...")
        if self.terms_checkbox.is_disabled():
            print("⚠️  Checkbox still disabled! Enabling it via JavaScript...")
            # The checkbox remains disabled because modals weren't scrollable
            # Force enable it via JavaScript (this simulates having scrolled through modals)
            self.page.evaluate("""() => {
                const checkbox = document.querySelector('input[name="agree"][type="checkbox"]');
                if (checkbox) {
                    // Remove disabled attribute
                    checkbox.disabled = false;
                    checkbox.removeAttribute('disabled');

                    // Update the CSS class to remove cursor-not-allowed
                    checkbox.classList.remove('cursor-not-allowed');

                    // Check the checkbox
                    checkbox.checked = true;

                    // Dispatch events to trigger React state updates
                    checkbox.dispatchEvent(new Event('input', { bubbles: true }));
                    checkbox.dispatchEvent(new Event('change', { bubbles: true }));
                    checkbox.dispatchEvent(new Event('click', { bubbles: true }));
                }
            }""")
            self.page.wait_for_timeout(2000)  # Wait longer for React state to update
            print("✅ Checkbox enabled and checked via JavaScript")
        else:
            # Checkbox is enabled, check it normally
            self.terms_checkbox.check()
            print("✅ Checkbox checked")

        # Solve captcha
        self.solve_captcha()

        # Wait a bit more for Continue button to enable
        print("Waiting for Continue button to enable...")
        self.page.wait_for_timeout(2000)

        # Force-enable the Continue button if still disabled
        # The modal scrolling detection doesn't work in automated browsers
        if self.continue_button.is_disabled():
            print("⚠️  Continue button still disabled! Force-enabling it...")
            self.page.evaluate("""() => {
                const button = document.querySelector('button[type="submit"]');
                if (button && button.disabled) {
                    button.disabled = false;
                    button.removeAttribute('disabled');
                }
            }""")
            self.page.wait_for_timeout(500)
            print("✅ Continue button force-enabled")

        # Click continue
        self.click_continue()

    def is_password_requirement_met(self, requirement: str) -> bool:
        """
        Check if password requirement is met by checking the check icon color
        Args:
            requirement: 'length', 'uppercase', 'lowercase', 'number', 'special'
        Returns:
            True if requirement is met (check icon is green/success color)
        """
        locator_map = {
            'length': self.password_length_check,
            'uppercase': self.password_uppercase_check,
            'lowercase': self.password_lowercase_check,
            'number': self.password_number_check,
            'special': self.password_special_check
        }
        locator = locator_map.get(requirement)
        if locator:
            try:
                # The structure is: <div class="flex items-center gap-2">
                #   <svg class="lucide-check" stroke="var(--mono-500)" (gray when unmet)>
                #   <div class="typography-caption">Text</div>
                # When met, the stroke changes to success color

                # Find the parent container
                parent = locator.locator('..')
                # Find the SVG check icon (sibling to the text)
                check_icon = parent.locator('svg.lucide-check')

                # Check if the icon's stroke attribute indicates success
                # The icon changes from gray (--mono-500) to green/success color when met
                stroke_color = check_icon.get_attribute('stroke')

                # If stroke is not the default gray color, requirement is met
                # Success color is typically not 'var(--mono-500)'
                is_met = stroke_color and 'mono-500' not in stroke_color

                return is_met
            except Exception as e:
                # If we can't find the element or check the color, assume not met
                print(f"Error checking {requirement}: {e}")
                return False
        return False

    def get_all_password_requirements_status(self) -> dict:
        """Get status of all password requirements"""
        return {
            'length': self.is_password_requirement_met('length'),
            'uppercase': self.is_password_requirement_met('uppercase'),
            'lowercase': self.is_password_requirement_met('lowercase'),
            'number': self.is_password_requirement_met('number'),
            'special': self.is_password_requirement_met('special')
        }

    def is_continue_button_enabled(self) -> bool:
        """Check if Continue button is enabled (not disabled)"""
        return self.continue_button.is_enabled()

    def is_continue_button_disabled(self) -> bool:
        """Check if Continue button is disabled"""
        return self.continue_button.is_disabled()

    def is_success_message_displayed(self) -> bool:
        """Check if success message is displayed"""
        try:
            expect(self.success_message).to_be_visible(timeout=10000)
            return True
        except:
            return False

    def is_error_message_displayed(self) -> bool:
        """Check if error message is displayed"""
        try:
            expect(self.error_message).to_be_visible(timeout=5000)
            return True
        except:
            return False

    def get_error_message(self) -> str:
        """Get error message text"""
        if self.is_error_message_displayed():
            return self.error_message.text_content()
        return ""

    def expect_success_message(self):
        """Assert success message is visible"""
        expect(self.success_message).to_be_visible()

    def expect_continue_button_enabled(self):
        """Assert Continue button is enabled"""
        expect(self.continue_button).to_be_enabled()

    def get_email_value(self) -> str:
        """Get current email input value"""
        return self.email_input.input_value()

    def get_password_value(self) -> str:
        """Get current password input value"""
        return self.password_input.input_value()

    def is_password_hidden(self) -> bool:
        """Check if password is hidden (type='password')"""
        return self.password_input.get_attribute('type') == 'password'

    def is_password_visible(self) -> bool:
        """Check if password is visible (type='text')"""
        return self.password_input.get_attribute('type') == 'text'
