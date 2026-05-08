"""
Base Page Object with common methods for all pages (Playwright)
"""
from playwright.sync_api import Page, expect
from config.config import Config


class BasePage:
    """Base page class with common methods"""

    def __init__(self, page: Page, base_url: str = None):
        self.page = page
        self.base_url = base_url or Config.BASE_URL
        # Set default timeout
        page.set_default_timeout(Config.DEFAULT_TIMEOUT)

    def navigate_to(self, url: str):
        """Navigate to a URL"""
        self.page.goto(url, timeout=Config.NAVIGATION_TIMEOUT)

    def click(self, selector: str):
        """Click element (Playwright auto-waits)"""
        self.page.locator(selector).click()

    def fill(self, selector: str, text: str):
        """Fill text into element (Playwright auto-waits)"""
        self.page.locator(selector).fill(text)

    def type(self, selector: str, text: str, delay: int = 0):
        """Type text into element with optional delay"""
        if delay > 0:
            self.page.locator(selector).type(text, delay=delay)
        else:
            self.page.locator(selector).fill(text)

    def get_text(self, selector: str) -> str:
        """Get element text"""
        return self.page.locator(selector).text_content()

    def is_visible(self, selector: str, timeout: int = 5000) -> bool:
        """Check if element is visible"""
        try:
            self.page.locator(selector).wait_for(state='visible', timeout=timeout)
            return True
        except:
            return False

    def is_hidden(self, selector: str, timeout: int = 5000) -> bool:
        """Check if element is hidden"""
        try:
            self.page.locator(selector).wait_for(state='hidden', timeout=timeout)
            return True
        except:
            return False

    def wait_for_url(self, url: str, timeout: int = None):
        """Wait for URL to match"""
        timeout = timeout or Config.DEFAULT_TIMEOUT
        self.page.wait_for_url(url, timeout=timeout)

    def wait_for_url_contains(self, url_part: str, timeout: int = None):
        """Wait for URL to contain substring"""
        timeout = timeout or Config.DEFAULT_TIMEOUT
        self.page.wait_for_url(f"**/*{url_part}*", timeout=timeout)

    def get_current_url(self) -> str:
        """Get current URL"""
        return self.page.url

    def check_checkbox(self, selector: str):
        """Check checkbox if not already checked"""
        checkbox = self.page.locator(selector)
        if not checkbox.is_checked():
            checkbox.check()

    def uncheck_checkbox(self, selector: str):
        """Uncheck checkbox if checked"""
        checkbox = self.page.locator(selector)
        if checkbox.is_checked():
            checkbox.uncheck()

    def is_checked(self, selector: str) -> bool:
        """Check if checkbox/radio is checked"""
        return self.page.locator(selector).is_checked()

    def is_enabled(self, selector: str) -> bool:
        """Check if element is enabled"""
        return self.page.locator(selector).is_enabled()

    def is_disabled(self, selector: str) -> bool:
        """Check if element is disabled"""
        return self.page.locator(selector).is_disabled()

    def get_attribute(self, selector: str, attribute: str) -> str:
        """Get element attribute"""
        return self.page.locator(selector).get_attribute(attribute)

    def execute_script(self, script: str, *args):
        """Execute JavaScript"""
        return self.page.evaluate(script, *args)

    def scroll_to_element(self, selector: str):
        """Scroll to element"""
        self.page.locator(selector).scroll_into_view_if_needed()

    def take_screenshot(self, filename: str, full_page: bool = False):
        """Take screenshot"""
        filepath = f"{Config.SCREENSHOT_DIR}/{filename}"
        self.page.screenshot(path=filepath, full_page=full_page)
        return filepath

    def get_page_title(self) -> str:
        """Get page title"""
        return self.page.title()

    def press_key(self, selector: str, key: str):
        """Press key in element"""
        self.page.locator(selector).press(key)

    def hover(self, selector: str):
        """Hover over element"""
        self.page.locator(selector).hover()

    def wait_for_selector(self, selector: str, state: str = 'visible', timeout: int = None):
        """
        Wait for selector
        Args:
            selector: CSS/XPath selector
            state: 'visible', 'hidden', 'attached', 'detached'
            timeout: Timeout in milliseconds
        """
        timeout = timeout or Config.DEFAULT_TIMEOUT
        self.page.locator(selector).wait_for(state=state, timeout=timeout)
