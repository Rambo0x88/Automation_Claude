"""
Pytest configuration and fixtures for Playwright tests
"""
import pytest
import pytest_html
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page
from config.config import Config
from datetime import datetime
import os
import base64


def pytest_addoption(parser):
    """Add custom command line options"""
    # Note: pytest-playwright provides --browser, --headed, --slowmo, --tracing, --video
    # We add our custom options that don't conflict
    pass


@pytest.fixture(scope="session")
def browser_type_launch_args(pytestconfig, browser_name):
    """Browser launch arguments - override pytest-playwright defaults"""
    launch_args = {
        "headless": Config.HEADLESS,
        "slow_mo": Config.SLOW_MO,
    }

    # Use system Chrome instead of Playwright's Chromium to avoid crashes on Apple Silicon
    # This also ensures modal rendering matches the user's manual Chrome experience
    # Only apply channel for chromium, not firefox
    if browser_name == "chromium":
        launch_args["channel"] = "chrome"  # Use installed Chrome browser
        # Add Chrome arguments for maximum stealth and anti-detection
        launch_args["args"] = [
            "--incognito",  # Enable incognito mode for better Cloudflare bypass
            "--start-fullscreen",  # Start browser in fullscreen mode
            "--disable-blink-features=AutomationControlled",  # Hide automation (critical!)
            "--disable-dev-shm-usage",  # Overcome limited resource problems
            "--no-sandbox",  # Disable sandbox for better compatibility
            "--disable-infobars",  # Disable infobars
            "--disable-notifications",  # Disable notifications
            # Additional anti-detection flags
            "--disable-features=IsolateOrigins,site-per-process",  # Disable site isolation
            "--disable-site-isolation-trials",  # Disable site isolation trials
            "--disable-web-security",  # Allow cross-origin requests (use with caution)
            "--disable-features=VizDisplayCompositor",  # Disable compositor
            "--disable-breakpad",  # Disable crash reporting
            "--disable-client-side-phishing-detection",  # Disable phishing detection
            "--disable-sync",  # Disable sync
            "--disable-background-timer-throttling",  # Disable throttling
            "--disable-backgrounding-occluded-windows",  # Disable backgrounding
            "--disable-renderer-backgrounding",  # Keep renderer active
            "--disable-component-extensions-with-background-pages",  # Disable background extensions
            "--disable-default-apps",  # Disable default apps
            "--disable-extensions",  # Disable all extensions
            "--disable-hang-monitor",  # Disable hang monitor
            "--disable-popup-blocking",  # Disable popup blocking
            "--disable-prompt-on-repost",  # Disable repost prompt
            "--metrics-recording-only",  # Only record metrics
            "--no-first-run",  # Skip first run wizards
            "--safebrowsing-disable-auto-update",  # Disable safe browsing updates
            "--enable-automation=false",  # Explicitly disable automation flag
            "--password-store=basic",  # Use basic password store
            "--use-mock-keychain",  # Use mock keychain
            "--force-color-profile=srgb",  # Force color profile
            "--disable-gpu",  # Disable GPU hardware acceleration
            "--lang=en-US,en",  # Set language
        ]

    return launch_args


@pytest.fixture(scope="session")
def browser_context_args(pytestconfig):
    """Browser context arguments - override pytest-playwright defaults"""
    # Use standard viewport size for proper UI rendering
    # Retina display uses 2x scaling, so use logical pixels (half of physical)
    context_args = {
        "viewport": {"width": 1440, "height": 900},  # Standard MacBook viewport size
        "device_scale_factor": 2,  # Retina display scaling factor
        "ignore_https_errors": True,
        # Enable incognito mode to ensure clean state for each test
        "storage_state": None,  # No saved cookies/localStorage
        # Add realistic user agent to bypass Cloudflare detection
        "user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        # Add locale for more realistic browser fingerprint
        "locale": "en-US",
        "timezone_id": "America/New_York",
        # Add screen size to match viewport for more realistic fingerprint
        "screen": {"width": 1440, "height": 900},
        # Enable JavaScript explicitly (helps with "Enable JavaScript" errors)
        "java_script_enabled": True,
        # Accept downloads for more realistic behavior
        "accept_downloads": True,
        # Add extra HTTP headers to appear more like a real browser
        "extra_http_headers": {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Cache-Control": "max-age=0"
        },
        # Add realistic geolocation permissions
        "permissions": ["geolocation"],
        "geolocation": {"latitude": 40.7128, "longitude": -74.0060},  # New York
        # Enable color scheme
        "color_scheme": "light",
        # Disable automation indicators
        "bypass_csp": True,
    }

    # Configure video recording based on Config
    if Config.VIDEO_ENABLED:
        context_args["record_video_dir"] = Config.VIDEO_DIR
        # Use fullscreen dimensions for video - match Retina display
        context_args["record_video_size"] = {"width": 2560, "height": 1664}

    return context_args


# Removed custom page fixture - using pytest-playwright's built-in page fixture
# which properly handles browser lifecycle management


def _extract_test_id_from_docstring(item):
    """Extract Test ID from docstring if available"""
    if item.function.__doc__:
        for line in item.function.__doc__.split('\n'):
            if 'Test ID:' in line:
                test_id = line.split('Test ID:')[1].strip()
                return test_id
    return None


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    """
    Hook to capture screenshots on test failure and add test data to HTML report
    Embeds screenshots and test data in HTML report
    """
    outcome = yield
    rep = outcome.get_result()

    # Add test data to HTML report for all tests (pass or fail)
    if rep.when == "call":
        extra = getattr(rep, 'extra', [])

        # Extract test data from stdout/logs
        if hasattr(rep, 'capstdout'):
            stdout = rep.capstdout
            test_data_html = "<div style='margin-top: 10px; padding: 10px; background-color: #f5f5f5; border-left: 4px solid #4CAF50;'>"
            test_data_html += "<strong>📊 Test Data Used:</strong><br/>"

            # Extract email (supports both "Test Email" and "New User Email")
            if "Email:" in stdout:
                import re
                email_match = re.search(r'(?:📧 )?(?:Test |New User )?Email:\s*(.+?)(?:\n|$)', stdout)
                if email_match:
                    test_data_html += f"<strong>Email:</strong> {email_match.group(1)}<br/>"

            # Extract password (supports both "Test Password" and "New User Password")
            if "Password:" in stdout:
                import re
                password_match = re.search(r'(?:🔑 )?(?:Test |New User )?Password:\s*(.+?)(?:\n|$)', stdout)
                if password_match:
                    test_data_html += f"<strong>Password:</strong> {password_match.group(1)}<br/>"

            # Extract addresses if present
            if "Address:" in stdout or "ETH Address:" in stdout or "BSC Address:" in stdout:
                import re
                address_matches = re.findall(r'(?:ETH|BSC|) ?Address:\s*(.+?)(?:\n|$)', stdout)
                if address_matches:
                    test_data_html += "<strong>Addresses:</strong><ul>"
                    for addr in address_matches:
                        test_data_html += f"<li>{addr}</li>"
                    test_data_html += "</ul>"

            test_data_html += "</div>"

            # Only add if we found some test data
            if "Email:" in test_data_html or "Address:" in test_data_html:
                extra.append(pytest_html.extras.html(test_data_html))
                rep.extra = extra

    # Store test result in context for tracing decision
    if rep.when == "call" and rep.failed:
        if hasattr(item, 'funcargs') and 'context' in item.funcargs:
            item.funcargs['context']._test_failed = True

        # Take screenshot on failure and embed in HTML report
        if hasattr(item, 'funcargs') and 'page' in item.funcargs:
            page = item.funcargs['page']

            # Check if page is still open before attempting screenshot
            try:
                if page and not page.is_closed():
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    test_id = _extract_test_id_from_docstring(item) or item.name
                    screenshot_name = f"{test_id}_{timestamp}.png"

                    # Save to test-results/screenshots/
                    screenshot_dir = "test-results/screenshots"
                    os.makedirs(screenshot_dir, exist_ok=True)
                    screenshot_path = os.path.join(screenshot_dir, screenshot_name)

                    # Take screenshot
                    screenshot_bytes = page.screenshot(path=screenshot_path)
                    print(f"\n📸 Screenshot saved: {screenshot_path}")

                    # Embed screenshot in HTML report
                    if screenshot_bytes:
                        extra = getattr(rep, 'extra', [])
                        # Encode screenshot as base64 for embedding in HTML
                        screenshot_base64 = base64.b64encode(screenshot_bytes).decode('utf-8')
                        extra.append(pytest_html.extras.image(screenshot_base64, name=screenshot_name))
                        rep.extra = extra
                else:
                    print(f"\n⏭️  Skipping screenshot - page is closed")
            except Exception as e:
                print(f"\n⚠️  Could not take screenshot: {e}")


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_teardown(item, nextitem):
    """
    Hook to rename video files with test ID and timestamp after test completes
    """
    # Store video path before page is closed
    video_path = None
    if hasattr(item, 'funcargs') and 'page' in item.funcargs:
        try:
            page = item.funcargs['page']
            if hasattr(page, 'video') and page.video:
                video_path = page.video.path()
        except:
            pass

    # Let the test teardown complete (this closes the page and finalizes video)
    yield

    # Now rename the video after it's been finalized
    if video_path:
        try:
            import time
            # Wait a bit for video to be fully written
            time.sleep(0.5)

            if os.path.exists(video_path):
                # Extract test ID from docstring or use test name
                test_id = _extract_test_id_from_docstring(item) or item.name.replace('[chromium]', '')
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

                # Create new filename: testid_date_time.webm
                video_dir = os.path.dirname(video_path)
                new_video_name = f"{test_id}_{timestamp}.webm"
                new_video_path = os.path.join(video_dir, new_video_name)

                # Rename the video file
                os.rename(video_path, new_video_path)
                print(f"\n🎥 Video saved: {new_video_path}")
        except Exception as e:
            # Silently ignore video renaming errors
            pass


@pytest.fixture(scope="session")
def test_data():
    """
    Provide test data for tests
    """
    from utils.helpers import generate_test_data
    return generate_test_data()


@pytest.fixture(scope="session")
def config():
    """Provide configuration to tests"""
    return {
        'base_url': Config.BASE_URL,
        'sign_up_url': Config.SIGN_UP_URL,
        'sign_in_url': Config.SIGN_IN_URL,
        'portfolio_url': Config.PORTFOLIO_URL,
        'email_domain': Config.EMAIL_DOMAIN
    }
