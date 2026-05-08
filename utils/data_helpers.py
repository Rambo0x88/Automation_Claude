"""Data cleaning and transformation helpers."""


def clean_currency_symbols(text):
    """Remove $ and % symbols from text data."""
    if isinstance(text, str):
        return text.replace('$', '').replace('%', '').strip()
    return text


def is_valid_evm_address(address_str):
    """Check if address is a valid EVM address (42 chars, starts with 0x)."""
    if not address_str:
        return False
    address_str = str(address_str).strip()
    return len(address_str) == 42 and address_str.lower().startswith("0x")
