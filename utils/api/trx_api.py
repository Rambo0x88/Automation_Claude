"""
TRX / TRON API — TronGrid and TronScan API functions.

Usage:
    from utils.api.trx_api import (
        fetch_token_details,
        fetch_trx_account_balance,
        fetch_trx_transactions,
        fetch_trx_price,
        fetch_token_list,
    )
"""

import time
import requests


# Default timeout for API calls
DEFAULT_TIMEOUT = 5


def fetch_token_details(contract_addr, has_all_info=False, timeout=None, max_retries=10):
    """
    Fetch TRC20 token details from TronScan API with aggressive retry logic.

    Args:
        contract_addr: TRC20 contract address
        has_all_info: If True, use shorter timeout (token already has basic info)
        timeout: API timeout in seconds (default: DEFAULT_TIMEOUT)
        max_retries: Maximum retry attempts

    Returns:
        tuple: (contract_addr, token_data_dict, success_bool)
    """
    if timeout is None:
        timeout = DEFAULT_TIMEOUT

    url = f"https://apilist.tronscanapi.com/api/token_trc20?contract={contract_addr}&showAll=1&start=&limit="
    effective_timeout = timeout if not has_all_info else max(1, timeout - 1)

    last_exception = None
    for attempt in range(max_retries):
        try:
            response = requests.get(url, timeout=effective_timeout)
            response.raise_for_status()
            data = response.json()

            if "trc20_tokens" in data and len(data["trc20_tokens"]) > 0:
                if attempt > 0:
                    print(f"      ✓ Token {contract_addr[:8]}... succeeded on attempt {attempt + 1}")
                return (contract_addr, data["trc20_tokens"][0], True)
            else:
                return (contract_addr, None, False)
        except Exception as e:
            last_exception = e
            if attempt < max_retries - 1:
                if attempt == 0:
                    print(f"      ⚠️  Token {contract_addr[:8]}... failed (attempt {attempt + 1}), retrying...")
                time.sleep(2)
            else:
                print(f"      ❌ Token {contract_addr[:8]}... failed after {max_retries} attempts: {str(last_exception)[:50]}")
                return (contract_addr, None, False)

    return (contract_addr, None, False)


def fetch_trx_account_balance(address, timeout=30):
    """
    Fetch TRX account balance from TronGrid API.

    Args:
        address: TRX wallet address (T..., 34 chars)
        timeout: API timeout in seconds

    Returns:
        tuple: (address, data_dict, success_bool)
    """
    url = f"https://api.trongrid.io/v1/accounts/{address}"
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        data = response.json()
        return (address, data, True)
    except Exception as e:
        print(f"   ❌ TronGrid account balance failed for {address[:10]}...: {str(e)[:50]}")
        return (address, None, False)


def fetch_trx_transactions(address, timeout=30):
    """
    Fetch recent transactions from TronGrid API.

    Args:
        address: TRX wallet address
        timeout: API timeout in seconds

    Returns:
        tuple: (address, data_dict, success_bool)
    """
    url = f"https://api.trongrid.io/v1/accounts/{address}/transactions"
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        data = response.json()
        return (address, data, True)
    except Exception as e:
        print(f"   ❌ TronGrid transactions failed for {address[:10]}...: {str(e)[:50]}")
        return (address, None, False)


def fetch_trx_price(timeout=30):
    """
    Fetch current TRX price from TronScan API.

    Returns:
        tuple: (price_usd, gain_24h, raw_data) or (None, None, None) on failure
    """
    url = "https://apilist.tronscanapi.com/api/token?id=0&showAll=1"
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        data = response.json()

        tokens = data.get("data", [])
        if tokens and len(tokens) > 0:
            market_info = tokens[0].get("market_info", {})
            price = market_info.get("priceInUsd")
            gain = market_info.get("gain")
            return (price, gain, data)
        return (None, None, data)
    except Exception as e:
        print(f"   ❌ TronScan TRX price failed: {str(e)[:50]}")
        return (None, None, None)


def fetch_token_list(timeout=30):
    """
    Fetch TRC20 token list from TronScan API (top 500 tokens).

    Returns:
        tuple: (tokens_list, raw_data) or ([], None) on failure
    """
    url = "https://apilist.tronscanapi.com/api/tokens/overview?start=0&limit=500&verifier=all&order=desc&filter=top&showAll=1&field="
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        data = response.json()
        tokens = data.get("tokens", [])
        return (tokens, data)
    except Exception as e:
        print(f"   ❌ TronScan token list failed: {str(e)[:50]}")
        return ([], None)
