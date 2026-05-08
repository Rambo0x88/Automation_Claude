"""
Sign In Page Object Model (Playwright)
Updated with accurate locators from signin.html
"""
from playwright.sync_api import Page, expect
from pages.base_page import BasePage
from config.config import Config


class SignInPage(BasePage):
    """Sign In page object"""

    def __init__(self, page: Page, base_url: str = None):
        super().__init__(page, base_url)
        self.url = Config.SIGN_IN_URL

        # Locators using data-testid and other attributes from actual HTML
        self.email_input = page.locator('input[data-testid="input-email"]')
        self.password_input = page.locator('input[data-testid="input-password"]')

        # Password visibility toggle - button with eye icon inside password container
        self.password_toggle = page.locator('input[data-testid="input-password"] ~ div button[type="button"]').first

        # Sign In button - has data-testid
        self.sign_in_button = page.locator('button[data-testid="sign-in-btn"]')

        # Forgot password link
        self.forgot_password_link = page.locator('a[href="/forgot-password"]')

        # Google Sign-in button - has data-testid
        self.google_signin_button = page.locator('button[data-testid="sign-in-google-btn"]')

        # Wallet Sign-in link
        self.wallet_signin_button = page.locator('a[href="/sign-in/wallet"]')

        # Sign Up link
        self.sign_up_link = page.locator('a[href="/sign-up"]')

        # Messages (Toastify notifications)
        self.error_message = page.locator('.Toastify__toast--error, [aria-invalid="true"]')
        self.success_message = page.locator('.Toastify__toast--success')

    def navigate(self):
        """Navigate to sign in page"""
        self.navigate_to(self.url)

    def enter_email(self, email: str):
        """Enter email address"""
        self.email_input.fill(email)

    def enter_password(self, password: str):
        """Enter password"""
        self.password_input.fill(password)

    def toggle_password_visibility(self):
        """Click password visibility toggle"""
        if self.password_toggle.is_visible(timeout=3000):
            self.password_toggle.click()

    def click_sign_in(self):
        """Click Sign In button"""
        self.sign_in_button.click()

    def click_forgot_password(self):
        """Click Forgot Password link"""
        self.forgot_password_link.click()

    def click_google_signin(self):
        """Click Sign In with Google button"""
        self.google_signin_button.click()

    def click_wallet_signin(self):
        """Click Sign In with Wallet link"""
        self.wallet_signin_button.click()

    def click_sign_up_link(self):
        """Click Sign Up link"""
        self.sign_up_link.click()

    def sign_in(self, email: str, password: str):
        """Complete sign in flow"""
        self.enter_email(email)
        self.enter_password(password)
        self.click_sign_in()

    def is_sign_in_button_enabled(self) -> bool:
        """Check if Sign In button is enabled"""
        return self.sign_in_button.is_enabled()

    def is_sign_in_button_disabled(self) -> bool:
        """Check if Sign In button is disabled"""
        return self.sign_in_button.is_disabled()

    def is_error_message_displayed(self) -> bool:
        """Check if error message is displayed"""
        try:
            expect(self.error_message).to_be_visible(timeout=5000)
            return True
        except:
            return False

    def is_success_message_displayed(self) -> bool:
        """Check if success message is displayed"""
        try:
            expect(self.success_message).to_be_visible(timeout=10000)
            return True
        except:
            return False

    def get_error_message(self) -> str:
        """Get error message text"""
        if self.is_error_message_displayed():
            return self.error_message.text_content()
        return ""

    def wait_for_successful_sign_in(self, timeout: int = 15000) -> bool:
        """
        Wait for successful sign in by checking URL change to portfolio page
        Returns True if successful, False otherwise
        """
        try:
            self.page.wait_for_url("**/portfolio**", timeout=timeout)
            return True
        except:
            return False

    def expect_sign_in_button_enabled(self):
        """Assert Sign In button is enabled"""
        expect(self.sign_in_button).to_be_enabled()

    def expect_redirected_to_portfolio(self):
        """Assert redirected to portfolio page"""
        import re
        expect(self.page).to_have_url(re.compile(r".*portfolio.*"))

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