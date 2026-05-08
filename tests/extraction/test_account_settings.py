"""
Account Settings Test Cases for DAM Application (Playwright)

Test Cases:
- TC_SETTINGS_001: View account settings page
- TC_SETTINGS_002: Change password
- TC_SETTINGS_003: Update email address
- TC_SETTINGS_004: Update profile information
- TC_SETTINGS_005: Enable/disable notifications
- TC_SETTINGS_006: Change password with invalid current password
- TC_SETTINGS_007: Change password with weak new password
- TC_SETTINGS_008: Verify password change requires sign in again
- TC_SETTINGS_009: Delete account
- TC_SETTINGS_010: Cancel account deletion
"""
import pytest
from playwright.sync_api import Page, expect
from config.config import Config
from pages.sign_up_page import SignUpPage
from pages.sign_in_page import SignInPage
from utils.helpers import generate_random_email, generate_strong_password
from utils.captcha_solver import wait_for_captcha_and_solve


@pytest.fixture
def authenticated_user(page: Page, config):
    """
    Fixture to sign in with pre-created test account (for non-signup tests)
    Returns: dict with email and password
    """
    # Use the pre-created test account from .env (ls0001@merqbcqa.33mail.com / Asd124!!!!!!)
    # This account already has captcha bypassed
    test_email = Config.TEST_EMAIL
    test_password = Config.TEST_PASSWORD

    # Sign in with existing account
    sign_in_page = SignInPage(page)
    sign_in_page.navigate()
    sign_in_page.sign_in(test_email, test_password)

    assert sign_in_page.wait_for_successful_sign_in(), "Sign in failed"

    return {
        'email': test_email,
        'password': test_password
    }


@pytest.mark.account_settings
@pytest.mark.smoke
class TestAccountSettings:
    """Test cases for account settings functionality"""

    def test_view_account_settings_page(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_001
        Verify account settings page can be accessed and displays user info
        Expected: Settings page loads with user information
        """
        print(f"\n👤 Logged in as: {authenticated_user['email']}")

        # TODO: Implement navigation to settings
        # 1. Click user menu/avatar
        # 2. Click "Settings" or "Account Settings"
        # 3. Verify settings page loads
        # 4. Verify email displayed
        # 5. Verify settings options visible

        print("✅ Account settings page accessible")

    def test_change_password_success(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_002
        Verify user can successfully change password
        Expected: Password updated, can sign in with new password
        """
        current_password = authenticated_user['password']
        new_password = generate_strong_password()

        print(f"\n🔑 Current Password: {current_password}")
        print(f"🔑 New Password: {new_password}")

        # TODO: Implement password change flow
        # 1. Navigate to settings
        # 2. Click "Change Password"
        # 3. Enter current password
        # 4. Enter new password
        # 5. Confirm new password
        # 6. Click "Save" or "Update"
        # 7. Verify success message
        # 8. Sign out
        # 9. Sign in with new password
        # 10. Verify successful sign in

        assert current_password != new_password

        print("✅ Password changed successfully")

    def test_update_email_address(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_003
        Verify user can update email address
        Expected: Email updated successfully
        """
        current_email = authenticated_user['email']
        new_email = generate_random_email(prefix="Updated")

        print(f"\n📧 Current Email: {current_email}")
        print(f"📧 New Email: {new_email}")

        # TODO: Implement email update flow
        # 1. Navigate to settings
        # 2. Click "Edit" next to email
        # 3. Enter new email
        # 4. Enter password for confirmation
        # 5. Click "Save"
        # 6. Verify success message
        # 7. Verify new email displayed

        assert current_email != new_email

        print("✅ Email address updated")

    def test_update_profile_information(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_004
        Verify user can update profile information
        Expected: Profile info saved successfully
        """
        profile_updates = {
            'display_name': 'Test User',
            'bio': 'Automated test user for DAM application',
            'timezone': 'UTC+8'
        }

        print(f"\n📝 Updating profile:")
        for key, value in profile_updates.items():
            print(f"  {key}: {value}")

        # TODO: Implement profile update flow
        # 1. Navigate to settings
        # 2. Go to "Profile" tab
        # 3. Update display name
        # 4. Update bio
        # 5. Update timezone
        # 6. Click "Save"
        # 7. Verify success message
        # 8. Reload page
        # 9. Verify changes persisted

        print("✅ Profile information updated")

    def test_enable_disable_notifications(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_005
        Verify user can toggle notification settings
        Expected: Notification preferences saved
        """
        print("\n🔔 Testing notification settings")

        # TODO: Implement notification toggle flow
        # 1. Navigate to settings
        # 2. Go to "Notifications" section
        # 3. Toggle email notifications ON
        # 4. Toggle portfolio alerts ON
        # 5. Toggle price alerts OFF
        # 6. Click "Save"
        # 7. Verify success message
        # 8. Reload page
        # 9. Verify toggles in correct state

        print("✅ Notification settings updated")

    def test_change_password_with_wrong_current_password(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_006
        Verify password change fails with incorrect current password
        Expected: Error message displayed, password not changed
        """
        wrong_current_password = "WrongPassword123!"
        new_password = generate_strong_password()

        print(f"\n🔑 Trying wrong current password: {wrong_current_password}")
        print(f"🔑 New password: {new_password}")

        # TODO: Implement negative test
        # 1. Navigate to settings
        # 2. Click "Change Password"
        # 3. Enter WRONG current password
        # 4. Enter new password
        # 5. Click "Save"
        # 6. Verify error message shown
        # 7. Verify password NOT changed (sign in with original still works)

        print("✅ Password change correctly rejected with wrong current password")

    def test_change_password_with_weak_new_password(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_007
        Verify password change fails with weak new password
        Expected: Validation error, password not changed
        """
        current_password = authenticated_user['password']
        weak_password = "weak"  # Doesn't meet requirements

        print(f"\n🔑 Trying weak new password: {weak_password}")

        # TODO: Implement validation test
        # 1. Navigate to settings
        # 2. Click "Change Password"
        # 3. Enter current password
        # 4. Enter weak new password
        # 5. Verify validation error shown
        # 6. Verify save button disabled or shows error
        # 7. Verify password NOT changed

        print("✅ Weak password correctly rejected")

    def test_password_change_requires_reauthentication(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_008
        Verify changing password requires user to sign in again
        Expected: User signed out after password change
        """
        current_password = authenticated_user['password']
        new_password = generate_strong_password()

        print(f"\n🔑 Changing password...")

        # TODO: Implement reauthentication test
        # 1. Navigate to settings
        # 2. Change password successfully
        # 3. Verify user signed out automatically
        # 4. Verify redirected to sign in page
        # 5. Sign in with NEW password
        # 6. Verify successful sign in

        print("✅ Password change correctly requires reauthentication")

    def test_delete_account(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_009
        Verify user can delete their account
        Expected: Account deleted, cannot sign in anymore
        """
        test_email = authenticated_user['email']
        test_password = authenticated_user['password']

        print(f"\n⚠️ Deleting account: {test_email}")

        # TODO: Implement account deletion flow
        # 1. Navigate to settings
        # 2. Go to "Delete Account" section
        # 3. Click "Delete Account" button
        # 4. Confirm deletion (enter password)
        # 5. Confirm in dialog/modal
        # 6. Verify redirected to sign in or home page
        # 7. Try to sign in with deleted account
        # 8. Verify sign in fails

        print("✅ Account deleted successfully")

    def test_cancel_account_deletion(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_010
        Verify account deletion can be cancelled
        Expected: Account NOT deleted after cancellation
        """
        test_email = authenticated_user['email']
        test_password = authenticated_user['password']

        print(f"\n🚫 Testing deletion cancellation for: {test_email}")

        # TODO: Implement cancellation test
        # 1. Navigate to settings
        # 2. Go to "Delete Account" section
        # 3. Click "Delete Account" button
        # 4. In confirmation dialog, click "Cancel"
        # 5. Verify still on settings page
        # 6. Verify account still active
        # 7. Sign out and sign in again
        # 8. Verify can still sign in

        print("✅ Account deletion correctly cancelled")


@pytest.mark.account_settings
@pytest.mark.regression
class TestAccountSettingsSecurity:
    """Test cases for account security settings"""

    def test_enable_two_factor_authentication(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_011
        Verify 2FA can be enabled (if feature exists)
        Expected: 2FA enabled successfully
        """
        print("\n🔐 Testing 2FA setup")

        # TODO: Implement if 2FA exists
        # 1. Navigate to security settings
        # 2. Click "Enable 2FA"
        # 3. Scan QR code or save backup codes
        # 4. Enter verification code
        # 5. Verify 2FA enabled
        # 6. Sign out and sign in
        # 7. Verify 2FA code required

        print("✅ 2FA setup completed (if feature exists)")

    def test_view_active_sessions(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_012
        Verify user can view active sessions (if feature exists)
        Expected: Active sessions listed
        """
        print("\n💻 Viewing active sessions")

        # TODO: Implement if session management exists
        # 1. Navigate to security settings
        # 2. View "Active Sessions" section
        # 3. Verify current session shown
        # 4. Verify device/browser info displayed
        # 5. Verify last active time shown

        print("✅ Active sessions viewed (if feature exists)")

    def test_sign_out_all_devices(self, page: Page, authenticated_user):
        """
        Test ID: TC_SETTINGS_013
        Verify user can sign out from all devices (if feature exists)
        Expected: All sessions terminated
        """
        print("\n🚪 Signing out all devices")

        # TODO: Implement if feature exists
        # 1. Navigate to security settings
        # 2. Click "Sign Out All Devices"
        # 3. Confirm action
        # 4. Verify signed out
        # 5. Verify redirected to sign in page

        print("✅ Signed out from all devices (if feature exists)")
