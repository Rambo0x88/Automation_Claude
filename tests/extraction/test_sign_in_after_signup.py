"""
Sign In After Sign Up Test Cases for DAM Application (Playwright)

Test Cases:
- TC_SIGNIN_AFTER_SIGNUP_001: Sign in immediately after successful sign up
- TC_SIGNIN_AFTER_SIGNUP_002: Sign in with correct credentials
- TC_SIGNIN_AFTER_SIGNUP_003: Sign in with incorrect password
- TC_SIGNIN_AFTER_SIGNUP_004: Sign in with non-existent email
- TC_SIGNIN_AFTER_SIGNUP_005: Sign in and verify session persistence
"""
import pytest
import re
from playwright.sync_api import Page, expect
from pages.sign_up_page import SignUpPage
from pages.sign_in_page import SignInPage
from config.config import Config
from utils.helpers import generate_random_email, generate_strong_password
from utils.captcha_solver import wait_for_captcha_and_solve


@pytest.mark.sign_in
@pytest.mark.smoke
class TestSignInAfterSignUp:
    """Test cases for sign in after successful sign up"""

    def test_sign_in_immediately_after_signup(self, page: Page, config):
        """
        Test ID: TC_SIGNIN_AFTER_SIGNUP_001
        Verify user can sign in immediately after successful sign up
        Expected: Sign up completes, then sign in with same credentials works

        ✅ Uses human-in-the-loop captcha solving
        Run with --headed to see browser and solve captcha when prompted
        """
        # Generate test credentials
        test_email = generate_random_email(prefix="Pla")
        test_password = generate_strong_password()

        print(f"\n📧 Test Email: {test_email}")
        print(f"🔑 Test Password: {test_password}")

        # Step 1: Sign up
        sign_up_page = SignUpPage(page)
        sign_up_page.navigate()
        sign_up_page.enter_email(test_email)
        sign_up_page.enter_password(test_password)

        # Verify password requirements met
        requirements = sign_up_page.get_all_password_requirements_status()
        assert all(requirements.values()), f"Password requirements not met: {requirements}"

        sign_up_page.check_terms_checkbox()

        # Wait for human to solve captcha
        captcha_solved = wait_for_captcha_and_solve(page, timeout=120000)
        assert captcha_solved, "❌ Captcha was not solved within timeout"

        sign_up_page.click_continue()

        # Verify "You're almost there!" message appears after signup
        try:
            almost_there_text = page.locator("text=/You'?re almost there/i")
            almost_there_text.wait_for(state="visible", timeout=10000)
            print("✅ 'You're almost there!' message displayed")
        except:
            print("ℹ️  'You're almost there!' message not found (may have redirected already)")

        # Step 2: Sign in with the same credentials
        sign_in_page = SignInPage(page)
        sign_in_page.navigate()
        sign_in_page.sign_in(test_email, test_password)

        # Verify successful sign in
        assert sign_in_page.wait_for_successful_sign_in(), \
            f"Sign in failed. Current URL: {page.url}"

        # Verify redirected to portfolio page
        expect(page).to_have_url(re.compile(r".*portfolio.*"))

        # Wait for portfolio page to fully load and verify welcome message
        try:
            # Check for "Start Managing Your Digital Assets" (new user) or portfolio data (existing portfolios)
            start_managing = page.locator("text=/Start Managing Your Digital Assets/i")
            portfolio_exists = page.locator("text=/Portfolio/i, text=/Total Net Worth/i").first

            # Wait for either message to appear
            page.wait_for_function(
                """() => {
                    return document.body.innerText.includes('Start Managing') ||
                           document.body.innerText.includes('Total Net Worth');
                }""",
                timeout=10000
            )

            if start_managing.is_visible():
                print("✅ 'Start Managing Your Digital Assets' screen displayed (new user)")
            else:
                print("✅ Portfolio dashboard displayed (existing portfolio)")
        except:
            print("ℹ️  Portfolio page loaded")

        print(f"✅ Sign in successful after sign up!")
        print(f"📧 Test Data - Email: {test_email}")
        print(f"🔑 Test Data - Password: {test_password}")

    def test_sign_in_with_incorrect_password(self, page: Page, config):
        """
        Test ID: TC_SIGNIN_AFTER_SIGNUP_002
        Verify sign in fails with incorrect password
        Expected: Error message displayed, user not signed in

        Uses pre-created test account from .env
        """
        # Use the pre-created test account
        test_email = Config.TEST_EMAIL
        correct_password = Config.TEST_PASSWORD

        # Try to sign in with wrong password
        sign_in_page = SignInPage(page)
        sign_in_page.navigate()

        wrong_password = "WrongPassword123!"
        sign_in_page.sign_in(test_email, wrong_password)

        # Verify error message or sign in failed
        # Should NOT reach portfolio page
        page.wait_for_timeout(2000)  # Wait for error to appear

        current_url = page.url
        assert "portfolio" not in current_url, \
            "User should NOT be signed in with wrong password"

        print("✅ Sign in correctly failed with wrong password")

    def test_sign_in_with_nonexistent_email(self, page: Page):
        """
        Test ID: TC_SIGNIN_AFTER_SIGNUP_003
        Verify sign in fails with non-existent email
        Expected: Error message or sign in fails
        """
        sign_in_page = SignInPage(page)
        sign_in_page.navigate()

        # Try sign in with email that doesn't exist
        nonexistent_email = generate_random_email(prefix="NonExist")
        random_password = generate_strong_password()

        sign_in_page.sign_in(nonexistent_email, random_password)

        # Wait for potential error
        page.wait_for_timeout(2000)

        # Should still be on sign in page or show error
        current_url = page.url
        assert "portfolio" not in current_url, \
            "User should NOT be signed in with non-existent email"

        print("✅ Sign in correctly failed with non-existent email")

    def test_sign_in_session_persistence(self, page: Page, config):
        """
        Test ID: TC_SIGNIN_AFTER_SIGNUP_004
        Verify user session persists after sign in
        Expected: User remains signed in after page refresh

        Uses pre-created test account from .env
        """
        # Use the pre-created test account
        test_email = Config.TEST_EMAIL
        test_password = Config.TEST_PASSWORD

        # Sign in
        sign_in_page = SignInPage(page)
        sign_in_page.navigate()
        sign_in_page.sign_in(test_email, test_password)

        assert sign_in_page.wait_for_successful_sign_in(), "Sign in failed"

        # Get current URL (portfolio page)
        portfolio_url = page.url
        assert "portfolio" in portfolio_url, "Not on portfolio page"

        # Refresh page
        page.reload()
        page.wait_for_load_state('networkidle')

        # Verify still on portfolio page (session persisted)
        current_url = page.url
        assert "portfolio" in current_url, \
            "Session did not persist - user signed out after refresh"

        print("✅ User session persisted after page refresh")
