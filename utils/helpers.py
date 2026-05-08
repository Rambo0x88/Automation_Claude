"""
Helper utility functions for DAM tests
"""
import random
import string
from datetime import datetime
from typing import List
from faker import Faker

fake = Faker()


def generate_random_email(prefix: str = "lily.su", domain: str = "@merquri.io") -> str:
    """
    Generate random email address using plus-addressing.
    Format: lily.su+test{number}@merquri.io

    Args:
        prefix: Email prefix before the + (default: "lily.su")
        domain: Email domain (default: @merquri.io)
    Returns:
        Random email address like lily.su+test1714_0930@merquri.io
    """
    timestamp = datetime.now().strftime('%m%d_%H%M')
    random_id = ''.join(random.choices(string.digits, k=4))
    return f"{prefix}+test{random_id}_{timestamp}{domain}"


def generate_strong_password(length: int = 16) -> str:
    """
    Generate strong password meeting requirements:
    - At least 12 characters
    - At least one uppercase letter
    - At least one lowercase letter
    - At least one number
    - At least one special character
    """
    uppercase = random.choices(string.ascii_uppercase, k=2)
    lowercase = random.choices(string.ascii_lowercase, k=5)
    digits = random.choices(string.digits, k=3)
    special = random.choices('!@#$%^&*', k=2)

    # Combine and shuffle
    all_chars = uppercase + lowercase + digits + special
    remaining = length - len(all_chars)
    if remaining > 0:
        all_chars += random.choices(string.ascii_letters + string.digits, k=remaining)

    random.shuffle(all_chars)
    return ''.join(all_chars)


def generate_portfolio_name() -> str:
    """Generate random portfolio name"""
    return f"Portfolio_{fake.word().title()}_{datetime.now().strftime('%m%d%H%M')}"


def generate_eth_address() -> str:
    """
    Generate random Ethereum address
    Format: 0x followed by 40 hexadecimal characters
    """
    hex_chars = '0123456789abcdef'
    address = '0x' + ''.join(random.choices(hex_chars, k=40))
    return address


def generate_bsc_address() -> str:
    """
    Generate random BSC (Binance Smart Chain) address
    Format: Same as Ethereum (0x + 40 hex chars)
    """
    return generate_eth_address()  # BSC uses same format as ETH


def generate_crypto_addresses(count: int = 1, chain: str = 'ETH') -> List[str]:
    """
    Generate multiple crypto addresses

    Args:
        count: Number of addresses to generate (1-3)
        chain: 'ETH' or 'BSC'

    Returns:
        List of crypto addresses
    """
    if count < 1 or count > 3:
        raise ValueError("Count must be between 1 and 3")

    generator = generate_bsc_address if chain == 'BSC' else generate_eth_address
    return [generator() for _ in range(count)]


def generate_test_data() -> dict:
    """Generate complete test data set"""
    return {
        'email': generate_random_email(prefix="Pla"),
        'password': generate_strong_password(),
        'portfolio_name': generate_portfolio_name()
    }


def make_screenshot_folder(tc_id: str, test_email: str, base_dir: str = "test-results/screenshots") -> str:
    """
    Create a screenshot subfolder for a test run.

    Format: {base_dir}/{tc_id}_{email_username}_{MMDD_HHMM}/

    Args:
        tc_id: Test case identifier (e.g. "tc1b-2", "tc2c-1", "DAMSS")
        test_email: Test email address (e.g. "lily.su@merquri.io")
        base_dir: Base screenshots directory (default: "test-results/screenshots")

    Returns:
        str: Full path to the created folder

    Examples:
        make_screenshot_folder("tc1b-2", "lily.su@merquri.io")
        → "test-results/screenshots/tc1b-2_lily.su_0428_1101"

        make_screenshot_folder("DAMSS", "moontest1311@gmail.com")
        → "test-results/screenshots/DAMSS_moontest1311_0428_1101"
    """
    import os
    email_user = test_email.split('@')[0] if '@' in test_email else test_email
    timestamp = datetime.now().strftime('%m%d_%H%M')
    folder_name = f"{tc_id}_{email_user}_{timestamp}"
    folder_path = os.path.join(base_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path
