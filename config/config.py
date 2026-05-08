"""
Configuration management for DAM Playwright tests v2
"""
import os

# Try to load .env file if python-dotenv is installed
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


class Config:
    """Test configuration settings"""

    # Project root directory (automationv2/)
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Application URLs
    BASE_URL = os.getenv('BASE_URL', 'https://dam-sit.mqbc21.com')
    SIGN_UP_URL = f"{BASE_URL}/sign-up"
    SIGN_IN_URL = f"{BASE_URL}/sign-in"
    PORTFOLIO_URL = f"{BASE_URL}/portfolio"

    # Test credentials
    TEST_EMAIL = os.getenv('TEST_EMAIL', '')
    TEST_PASSWORD = os.getenv('TEST_PASSWORD', 'TestPassword123!')
    EMAIL_DOMAIN = '@merquri.io'

    # Browser settings
    BROWSER = os.getenv('BROWSER', 'chromium').lower()
    HEADLESS = os.getenv('HEADLESS', 'false').lower() == 'true'
    SLOW_MO = int(os.getenv('SLOW_MO', '250'))

    # Timeouts (milliseconds)
    DEFAULT_TIMEOUT = int(os.getenv('DEFAULT_TIMEOUT', '30000'))
    NAVIGATION_TIMEOUT = int(os.getenv('NAVIGATION_TIMEOUT', '30000'))

    # Video/Screenshot settings
    VIDEO_ENABLED = os.getenv('VIDEO_ENABLED', 'on-failure')
    SCREENSHOT_ON_FAILURE = True
    VIDEO_DIR = 'videos'
    SCREENSHOT_DIR = 'test-results/screenshots'

    # API settings
    COVALENT_API_KEY = os.getenv('COVALENT_API_KEY', '')
    ENABLE_COINGECKO = os.getenv('ENABLE_COINGECKO', 'false').lower() == 'true'

    # Extraction settings
    SKIP_PRICE_FETCHING = os.getenv('SKIP_PRICE_FETCHING', 'false').lower() == 'true'
    MAX_TOKENS_TO_PROCESS = int(os.getenv('MAX_TOKENS_TO_PROCESS', '0'))
    PARALLEL_API_CALLS = int(os.getenv('PARALLEL_API_CALLS', '3'))
    API_TIMEOUT_SECONDS = int(os.getenv('API_TIMEOUT_SECONDS', '5'))


# Create directories if they don't exist
os.makedirs(Config.VIDEO_DIR, exist_ok=True)
os.makedirs(Config.SCREENSHOT_DIR, exist_ok=True)
os.makedirs('test-results/excel-exports', exist_ok=True)
os.makedirs('test-results/reports', exist_ok=True)
