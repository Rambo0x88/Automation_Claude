"""
Account Settings Page Object Model (Playwright)
"""
from playwright.sync_api import Page, expect
from pages.base_page import BasePage
from config.config import Config


class AccountSettingsPage(BasePage):
    """Account Settings page object"""

    def __init__(self, page: Page, base_url: str = None):
        super().__init__(page, base_url)
        self.url = f"{Config.BASE_URL}/settings"

        # Locators (Playwright style)
        self.user_profile_icon = page.locator("button[aria-label*='user' i], [class*='profile'], [class*='user-icon']")
        self.account_settings_link = page.locator("a:has-text('Account Settings'), a:has-text('Settings'), button:has-text('Settings')")

        # Change Password section - Updated with modal selectors
        self.change_password_button = page.locator("button:has-text('Password'), button:has-text('Change Password')")
        self.password_modal = page.locator("[role='dialog'], [class*='modal']")
        self.current_password_input = page.locator("input[placeholder*='Current password' i], input[name='currentPassword'], input[id='currentPassword'], input[type='password']").first
        self.new_password_input = page.locator("input[placeholder*='New password' i], input[name='newPassword'], input[id='newPassword']").nth(1)
        self.confirm_password_input = page.locator("input[placeholder*='Confirm' i], input[name='confirmPassword'], input[id='confirmPassword']").last
        self.save_password_button = page.locator("button:has-text('Save'), button:has-text('Update'), button:has-text('Confirm')")
        self.cancel_button = page.locator("button:has-text('Cancel')")

        # Password requirements
        self.password_length_check = page.locator("text=At least 12 characters")
        self.password_uppercase_check = page.locator("text=At least one uppercase letter")
        self.password_lowercase_check = page.locator("text=At least one lowercase letter")
        self.password_number_check = page.locator("text=At least one number")
        self.password_special_check = page.locator("text=At least one special character")

        # Messages
        self.success_message = page.locator("text=/Success|successfully|Password changed/i")
        self.error_message = page.locator("[class*='error'], [class*='alert']")

    def navigate(self):
        """Navigate to account settings page"""
        self.navigate_to(self.url)

    def open_account_settings_from_profile(self):
        """Open account settings from user profile dropdown"""
        self.user_profile_icon.click()
        self.account_settings_link.wait_for(state='visible')
        self.account_settings_link.click()

    def click_change_password_button(self):
        """Click Change Password button to open modal"""
        if self.change_password_button.is_visible(timeout=5000):
            self.change_password_button.click()
            # Wait for modal to appear
            self.password_modal.wait_for(state='visible', timeout=5000)

    def enter_current_password(self, password: str):
        """Enter current password"""
        self.current_password_input.fill(password)

    def enter_new_password(self, password: str):
        """Enter new password"""
        self.new_password_input.fill(password)

    def enter_confirm_password(self, password: str):
        """Enter confirm password"""
        self.confirm_password_input.fill(password)

    def click_save_password(self):
        """Click Save/Update button"""
        self.save_password_button.click()

    def click_cancel(self):
        """Click Cancel button"""
        self.cancel_button.click()

    def change_password(self, current_password: str, new_password: str):
        """
        Complete password change flow
        Args:
            current_password: Current password
            new_password: New password
        """
        self.click_change_password_button()
        self.page.wait_for_timeout(1000)  # Wait for modal animation
        self.enter_current_password(current_password)
        self.page.wait_for_timeout(500)
        self.enter_new_password(new_password)
        self.page.wait_for_timeout(500)
        self.enter_confirm_password(new_password)
        self.page.wait_for_timeout(500)
        self.click_save_password()

    def is_success_message_displayed(self) -> bool:
        """Check if success message is displayed"""
        try:
            expect(self.success_message).to_be_visible(timeout=10000)
            return True
        except:
            return False

    def get_success_message(self) -> str:
        """Get success message text"""
        if self.is_success_message_displayed():
            return self.success_message.text_content()
        return ""

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

    def is_save_button_enabled(self) -> bool:
        """Check if Save button is enabled"""
        return self.save_password_button.is_enabled()

    def wait_for_password_changed(self, timeout: int = 10000) -> bool:
        """
        Wait for password change success
        Returns True if success message appears, False otherwise
        """
        return self.is_success_message_displayed()

    def expect_success_message(self):
        """Assert success message is visible"""
        expect(self.success_message).to_be_visible()

    def expect_save_button_enabled(self):
        """Assert Save button is enabled"""
        expect(self.save_password_button).to_be_enabled()
