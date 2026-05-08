# API Reference - DAM Automation v2

## Overview

This document covers all external APIs used in DAM automation:
- TRX Balance API (TronGrid account balance)
- TronGrid Transaction History API (transaction list + TRC20 transfers)
- TronScan Token Detail API (TRC20 token metadata and prices)
- SimDune API (EVM on-chain balances)
- Rabby Protocol API (DeFi positions)
- CoinGecko API (prices + 24H price change)

---

## 1. TRX / TRON APIs

All TRON-related APIs used across `run_overview.py`, `tests/extraction/test_trx_balance.py`, and `utils/trx_transaction/` scripts.

### 1.1 TronGrid Account Balance

**Purpose**: Fetch TRX native balance + all TRC20 token balances for an address.

**Endpoint**:
```
GET https://api.trongrid.io/v1/accounts/{address}
```

**Input**: TRX wallet address (starts with `T`, 34 characters)

**TRX Balance calculation** (all values in SUN; 1 TRX = 1,000,000 SUN):
```
TRX_raw = balance
         + SUM(frozen[].frozen_balance)
         + SUM(frozenV2[].amount)
         + SUM(unfrozenV2[].unfreeze_amount)
         + account_resource.delegated_frozenV2_balance_for_energy
         + delegated_frozenV2_balance_for_bandwidth (root level)
```

**TRC20 tokens**: from `account_data.trc20[]` — each entry is `{contractAddress: balance_raw_string}`.

**Implementation**:
- `run_overview.py` (inline)
- `utils/trx_transaction/trx_balance_and_dam_extraction.py` (inline)
- `tests/extraction/test_trx_balance.py`

**Excel output tabs** (in `API_TRXBalance_*.xlsx`):
- `API - TRX Balance` — raw JSON response
- `TRX Balance, Price` — parsed rows for TRX + all TRC20 tokens

**TRX Balance, Price columns**:

| Col | Field | Source |
|-----|-------|--------|
| A | Address | Input address |
| B | Decimal Places | Token List / Token Detail API |
| C | Token | Token symbol (`abbr`) |
| D | Contract Address | From TRC20 map |
| E | Balance | `Balance_Raw / 10^Decimals` |
| F | Balance (Raw) | Raw integer from API |
| G | Price | TronScan TRX Price API or Token Detail API |
| H | Price (24h) | `market_info.gain * 100` |
| I | Symbol Show | `symbolShow` from Token List |
| J | Calculated Value | `Balance × Price` |

**Rate Limiting**: No official limit. Recommended 1-2 req/s.

---

### 1.2 TronGrid Transaction History

**Purpose**: Fetch all transactions for a TRX address within a date range. Used by `utils/trx_transaction/trongrid_dam_comparison.py`.

**Endpoint**:
```
GET https://api.trongrid.io/v1/accounts/{address}/transactions
    ?min_timestamp={ms}&max_timestamp={ms}&limit=200&order_by=block_timestamp,asc
```

**Pagination**: Uses `fingerprint` from `meta` field. Keep fetching until no fingerprint or empty batch.

**Fields per transaction**:

| Field | Description |
|-------|-------------|
| `txID` | Transaction hash (64 hex chars) |
| `block_timestamp` | Unix timestamp in milliseconds |
| `ret[0].fee` | Fee in SUN |
| `ret[0].contractRet` | Status (`SUCCESS`, `REVERT`, etc.) |
| `raw_data.contract[0].type` | Contract type (see classification table below) |
| `raw_data.contract[0].parameter.value` | Transaction parameters (from, to, amount, etc.) |

**Inflow/Outflow classification** (in `parse_tx()` function):

| Contract Type | Direction | Net Transfer |
|---------------|-----------|-------------|
| `TransferContract` | Send/Receive | `±amount TRX` |
| `TriggerSmartContract` + TRC20 match | Send/Receive | `±amount TOKEN` |
| `TriggerSmartContract` (no TRC20) | Smart Contract Call | `Fee: -fee TRX` |
| `FreezeBalanceV2Contract` | Outflow | `Staked: amount TRX` |
| `UnfreezeBalanceV2Contract` | Inflow | `+amount TRX (unfrozen)` |
| `DelegateResourceContract` | Outflow | `Delegated amount TRX` |
| `UnDelegateResourceContract` | Inflow | `Undelegated amount TRX` |
| `VoteWitnessContract` | Neutral | `Vote cast` |
| `WithdrawExpireUnfreezeContract` | Inflow | `+TRX withdrawn from stake` |
| `WithdrawBalanceContract` | Inflow | `+TRX reward claimed` |

**Excel output** (Step 3 in comparison pipeline):
- `TronGrid Transactions` — Parsed: #, Trx Hash, Date/Time, Transaction Type, From, To, Amount, Resources Fee, Token Transfer, Net Transfer
- `TRC20 Transfers` — All TRC20 token transfers

---

### 1.3 TronGrid TRC20 Transfers

**Purpose**: Fetch all TRC20 token transfers for an address within a date range. Used alongside 1.2 to build a TRC20 lookup map.

**Endpoint**:
```
GET https://api.trongrid.io/v1/accounts/{address}/transactions/trc20
    ?min_timestamp={ms}&max_timestamp={ms}&limit=200&order_by=block_timestamp,asc
```

**Pagination**: Same fingerprint-based pagination as 1.2.

**Fields per transfer**:

| Field | Description |
|-------|-------------|
| `transaction_id` | Transaction hash (used to join with 1.2) |
| `block_timestamp` | Unix timestamp in ms |
| `from` | Sender address |
| `to` | Receiver address |
| `value` | Raw token amount (divide by `10^decimals`) |
| `token_info.symbol` | Token symbol |
| `token_info.name` | Token name |
| `token_info.decimals` | Token decimals |
| `token_info.address` | Token contract address |

---

### 1.4 TronScan Token List

**Purpose**: Pre-load token metadata (symbol, decimals, contract address) to avoid per-token API calls.

**Endpoint**:
```
GET https://apilist.tronscanapi.com/api/tokens/overview
    ?start=0&limit=500&verifier=all&order=desc&filter=top&showAll=1
```

**Fields used**:

| Response Field | Used As |
|---|---|
| `abbr` | Token symbol |
| `decimal` | Decimal places |
| `contractAddress` | Contract address (lookup key) |
| `canShow` | Whether token should be shown |

**Excel tabs**:
- `API - All Token Info` — raw JSON response
- `Token List` — parsed: `abbr`, `decimal`, `contractAddress`, `canShow`

---

### 1.5 TronScan TRX Price

**Purpose**: Fetch current TRX price and 24h change.

**Endpoint**:
```
GET https://apilist.tronscanapi.com/api/token?id=0&showAll=1
```

**Fields used**:

| Response Path | Used As |
|---|---|
| `data[0].market_info.priceInUsd` | TRX Price |
| `data[0].market_info.gain * 100` | TRX Price 24h % |

---

### 1.6 TronScan Token Detail (per TRC20 contract)

**Purpose**: Fetch symbol, decimals, and price for TRC20 tokens not already in the Token List.

**Endpoint**:
```
GET https://apilist.tronscanapi.com/api/token_trc20?contract={contractAddress}&showAll=1
```

**Only called when** the token is missing from the pre-loaded Token List.

**Retry logic**: Up to 10 attempts with 2s delay between failures.

**Fields used**:

| Response Path | Used As |
|---|---|
| `trc20_tokens[0].symbol` | Token symbol (if missing) |
| `trc20_tokens[0].decimals` | Decimal places (if missing) |
| `trc20_tokens[0].symbolShow` | Symbol show (if missing) |
| `trc20_tokens[0].market_info.priceInUsd` | Price |
| `trc20_tokens[0].market_info.gain * 100` | Price 24h % |

**Excel tab**: `API - TRC20 Token` — one row per contract: `contractAddress`, full JSON response

---

### 1.7 TRX API Raw JSON Output

File: `API_TRXBalance_<name>_<ts>_Raw.json`

```json
{
  "token_list": { "...TronScan token list response..." },
  "trx_price":  { "...TronScan TRX price response..." },
  "addresses": [
    {
      "address": "T...",
      "balance": { "...TronGrid account balance response..." },
      "transactions": { "...TronGrid transactions response..." },
      "token_details": [
        { "contract_address": "TR...", "response": { "trc20_tokens": ["..."] } }
      ]
    }
  ]
}
```

---

### 1.8 TRX Validation in DAM Excel

The `TRX Balance, Price` sheet is copied into the DAM Excel and used for:

| Validation | Source Column | DAM Column |
|---|---|---|
| Price Validation | Col G — Price | Compared against DAM Price Tooltip |
| Price 24h Validation | Col H — Price (24h) | Compared against DAM Price (24h) |
| Amount Validation | Col F — Balance (Raw) + Col B — Decimals | Compared against DAM Amount Tooltip |
| Combined Net Worth Calculated Value | Col J — Calculated Value | Summed per address |

---

## 2. SimDune API

### Purpose
Fetch EVM on-chain balances and token holdings for Ethereum addresses.

### Endpoint
```
GET https://api.simdune.com/v1/user/balances/{address}
```

### Input
- **Address**: EVM wallet address (0x + 40 hex characters)
- **Format**: `0x` + 40 hexadecimal characters
- **Length**: 42 characters total

### Example
```python
from utils.api.simdune_api import fetch_simdune_balance

balance = fetch_simdune_balance("0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab")
print(balance)  # Returns token balances
```

### Output
```json
{
  "address": "0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab",
  "tokens": [
    {
      "symbol": "ETH",
      "balance": "10.5",
      "decimals": 18,
      "contract": "0x0000000000000000000000000000000000000000"
    },
    {
      "symbol": "USDC",
      "balance": "5000.0",
      "decimals": 6,
      "contract": "0xa0b86991c6218b36c1d19d4a2e9eb0ce3606eb48"
    }
  ]
}
```

### Implementation
- **File**: `utils/api/simdune_api.py`
- **Test**: `tests/api/test_simdune_api.py`

### Rate Limiting
- 100 requests per minute
- Recommended: 1 request per second

---

## 3. Rabby Protocol API

### 3.1 Complex Protocol List (DeFi Protocols)

**Purpose**: Fetch DeFi protocol positions (Aave, Compound, Morpho, etc.)

**Endpoint**
```
GET https://api.rabby.io/v1/user/complex_protocol_list?id={address}
```

**Input**
- **Address**: EVM wallet address (0x + 40 hex characters)

**Example**
```python
from utils.api.rabby_api import fetch_protocol_list

protocols = fetch_protocol_list("0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab")
print(protocols)  # Returns DeFi positions
```

**Output**
```json
{
  "address": "0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab",
  "protocols": [
    {
      "name": "Aave V3",
      "positions": [
        {
          "asset": "ETH",
          "balance": "10.5",
          "value": "25000"
        }
      ]
    }
  ]
}
```

### 3.2 Complex App List (Off-Chain Apps)

**Purpose**: Fetch off-chain app positions (Hyperliquid, etc.)

**Endpoint**
```
GET https://api.rabby.io/v1/user/complex_app_list?id={address}
```

**Input**
- **Address**: EVM wallet address (0x + 40 hex characters)

**Example**
```python
from utils.api.rabby_api import fetch_app_list

apps = fetch_app_list("0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab")
print(apps)  # Returns off-chain app positions
```

**Output**
```json
{
  "address": "0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab",
  "apps": [
    {
      "name": "Hyperliquid",
      "positions": [
        {
          "asset": "PURR",
          "balance": "1000",
          "value": "5000"
        }
      ]
    }
  ]
}
```

### Implementation
- **File**: `utils/api/rabby_api.py`
- **Test**: `tests/api/test_rabby_api.py`

### Rate Limiting
- No official rate limit
- Recommended: 1-2 requests per second

---

## 4. CoinGecko API

### Purpose
Fetch cryptocurrency prices and market data. Used by `run_overview.py` and `utils/trx_transaction/` scripts.

### 4.1 Simple Price (Batch)

**Endpoint**:
```
GET https://api.coingecko.com/api/v3/simple/price?vs_currencies=usd&ids={comma_separated_ids}
```

**Headers**: `x-cg-demo-api-key: CG-F3KENg4b1mcvyeg6eo6LGDQU`

**Batch size**: Up to 250 IDs per request.

**Output**:
```json
{
  "ethereum": { "usd": 2500 },
  "bitcoin": { "usd": 45000 }
}
```

### 4.2 Coin Detail — 24H Price Change

**Purpose**: Fetch 24H price change percentage for individual coins.

**Endpoint**:
```
GET https://api.coingecko.com/api/v3/coins/{coin_id}
```

**Headers**: `x-cg-demo-api-key: CG-F3KENg4b1mcvyeg6eo6LGDQU`

**Field used**: `market_data.price_change_percentage_24h`

**Rate limiting**: 0.5s delay between individual coin requests to avoid hitting limits.

### 4.3 Coingecko Coin ID Lookup

**Purpose**: Map token addresses to CoinGecko coin IDs for price fetching.

**Source file**: `Coingecko Coin ID List.xlsx` (sheet: `Coin ID List`, ~26k rows)

**Columns**: A=ID, B=Symbol, C=Name, D=Platform Name, E=Platform Address

**Lookup priority**:
1. **(address + platform)**: `(token_address, platform_name)` → coin_id
2. **(address only)**: `token_address` → coin_id (fallback when platform name differs)
3. **(native symbol)**: `symbol` → coin_id (for native tokens like ETH, BNB where address = "native")

**Chain name mapping** (SimDune chain → CoinGecko platform):

| SimDune Chain | CoinGecko Platform |
|---|---|
| ethereum | ethereum |
| bsc / bnb | binance-smart-chain |
| polygon | polygon-pos |
| arbitrum | arbitrum-one |
| optimism | optimistic-ethereum |
| base | base |
| avalanche | avalanche |
| fantom | fantom |
| linea | linea |
| scroll | scroll |
| zksync | zksync |
| gnosis | xdai |

**Spam detection**: If a token address is valid EVM (42 chars, 0x prefix) but not found in CoinGecko, it's marked as `"Spam Token"`.

### Rate Limiting
- 10-50 calls/minute (free tier with demo API key)
- Recommended: 1 request per second
- Batch `simple/price` endpoint preferred over individual `coins/{id}` calls

### Note
- Disabled by default due to aggressive bot detection (403 Forbidden)
- To enable: Set `ENABLE_COINGECKO=true` in `.env`
- Use Firefox browser for better detection evasion

---

## 5. DAM Application API

### Purpose
Internal DAM application APIs for portfolio management.

### Endpoints

#### Get Portfolios
```
GET /api/portfolios
```

#### Create Portfolio
```
POST /api/portfolios
Body: {
  "name": "Portfolio Name",
  "addresses": ["0x...", "T..."]
}
```

#### Get Portfolio Details
```
GET /api/portfolios/{portfolioId}
```

#### Update Portfolio
```
PUT /api/portfolios/{portfolioId}
Body: {
  "name": "New Name",
  "addresses": ["0x...", "T..."]
}
```

### Implementation
- **File**: `utils/portfolio/portfolio_manager.py`
- **Test**: `tests/ui/test_portfolio.py`

---

## API Error Handling

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| 404 Not Found | Invalid address format | Check address format (0x... or T...) |
| 429 Too Many Requests | Rate limit exceeded | Reduce request frequency |
| 403 Forbidden | Bot detection (CoinGecko) | Use Firefox browser or disable |
| 500 Server Error | API server issue | Retry after 5 seconds |
| Connection Timeout | Network issue | Check internet connection |

### Retry Logic
```python
from utils.api.base_api import retry_request

# Automatically retries up to 3 times with exponential backoff
response = retry_request(
    url="https://api.example.com/endpoint",
    max_retries=3,
    timeout=5
)
```

---

## Performance Optimization

### Parallel API Calls
```python
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=3) as executor:
    futures = [
        executor.submit(fetch_simdune_balance, addr)
        for addr in addresses
    ]
    results = [f.result() for f in futures]
```

### Caching
```python
from utils.api.cache import cache_result

@cache_result(ttl=3600)  # Cache for 1 hour
def fetch_prices(coin_ids):
    return fetch_coingecko_prices(coin_ids)
```

### Batch Processing
```python
from utils.api.batch import batch_requests

# Process 100 addresses in batches of 10
results = batch_requests(
    addresses=addresses,
    batch_size=10,
    request_func=fetch_simdune_balance
)
```

---

## Testing APIs

### Unit Tests
```bash
pytest tests/api/ -v
```

### Integration Tests
```bash
pytest tests/extraction/ -v
```

### Mock API Responses
```python
from unittest.mock import patch

@patch('utils.api.rabby_api.fetch_protocol_list')
def test_with_mock(mock_fetch):
    mock_fetch.return_value = {"protocols": []}
    result = fetch_protocol_list("0x...")
    assert result == {"protocols": []}
```

---

## Troubleshooting

### API Not Responding
1. Check internet connection
2. Verify API endpoint is correct
3. Check API status page
4. Try with different address format

### Rate Limiting
1. Reduce request frequency
2. Implement exponential backoff
3. Use batch processing
4. Consider API key upgrade

### Invalid Response
1. Check address format
2. Verify API response structure
3. Check for API changes
4. Review error logs

---

## References

- [TronGrid API Documentation](https://developers.tron.network/reference/account-info-by-address)
- [TronScan API Documentation](https://docs.tronscan.org/)
- [SimDune API Documentation](https://simdune.com/docs)
- [Rabby API Documentation](https://api.rabby.io/)
- [CoinGecko API Documentation](https://www.coingecko.com/en/api)
