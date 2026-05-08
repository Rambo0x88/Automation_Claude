# API Data Source Analysis - Where Each Column Comes From

This document traces every output column back to its API source. Covers both the **SimDune/CoinGecko flow** (EVM portfolio extraction) and the **TronGrid/TronScan flow** (TRX transaction comparison).

---

# Part A: SimDune + CoinGecko Flow (EVM Portfolios)

Used by: `run_overview.py`, `utils/trx_transaction/trx_balance_and_dam_extraction.py`

## Column Data Sources

### Column 1: Chain
**Source**: Sim Dune API
**Field**: `item.get('chain', item.get('chain_name', ''))`
**Example**: "ethereum", "base", "binance-smart-chain"

### Column 2: Symbol
**Source**: Sim Dune API
**Field**: `item.get('symbol', item.get('token_symbol', ''))`
**Example**: "USDC", "ETH", "AERO"

### Column 3: Amount (Raw)
**Source**: Sim Dune API
**Field**: `item.get('amount', item.get('balance', item.get('value_raw', '')))`
**Example**: "1000000000"

### Column 4: Amount
**Source**: Calculated from Sim Dune API
**Formula**: `Amount_Raw / (10 ^ Decimals)`
**Example**: "1000000000" / 10^6 = "1000"

### Column 5: Decimals
**Source**: Sim Dune API
**Field**: `item.get('decimals', 18)`
**Example**: 6, 18, 8

### Column 6: Token Address
**Source**: Sim Dune API
**Field**: `item.get('address', '')`
**Example**: "0xa0b86991c6218b36c1d19d4a2e9eb0ce3606eb48"

### Column 7: ID (Coingecko ID)
**Source**: Coingecko (via lookup)
**Lookup Method**: 
- Uses Token Address from Sim Dune
- Looks up in Coingecko coin list
- Returns Coingecko ID or "Spam Token"
**Example**: "usd-coin", "ethereum", "aero"

### Column 8: Price
**Source**: Coingecko API
**API Endpoint**: `https://api.coingecko.com/api/v3/simple/price`
**Lookup**: Uses Coingecko ID from Column 7
**Example**: 1.00, 2403.07

### Column 9: 24H Price Change
**Source**: Coingecko API
**API Endpoint**: `https://api.coingecko.com/api/v3/coins/{id}`
**Lookup**: Uses Coingecko ID from Column 7
**Example**: 0.5, -2.3

---

## Data Flow Diagram

```
Sim Dune API Response
    ↓
    ├─ chain → Column 1: Chain
    ├─ symbol → Column 2: Symbol
    ├─ amount → Column 3: Amount (Raw)
    ├─ decimals → Column 5: Decimals
    ├─ address → Column 6: Token Address
    │
    └─ address + chain → Coingecko Lookup
        ↓
        ├─ Found → Column 7: ID (Coingecko ID)
        │   ↓
        │   ├─ Coingecko API (simple/price) → Column 8: Price
        │   └─ Coingecko API (coins/{id}) → Column 9: 24H Price Change
        │
        └─ Not Found → Column 7: "Spam Token"
            ↓
            └─ Debank API (fallback) → Column 8 & 9: Price & 24H Change
```

---

## Key Question: What if Sim Dune Doesn't Have Token Name?

**Answer**: The token name is NOT extracted from Sim Dune API at all!

### Current Extraction (Lines 500-510):
```python
chain = item.get('chain', item.get('chain_name', ''))
symbol = item.get('symbol', item.get('token_symbol', ''))
amount_raw = item.get('amount', item.get('balance', item.get('value_raw', '')))
decimals = item.get('decimals', 18)
token_address = item.get('address', '')
```

**Notice**: NO extraction of `name` field!

### Coingecko Lookup (Lines 525-545):
```python
# Case 1: Valid EVM address
if is_valid_evm_address(token_address):
    lookup_key = (token_address_lower, platform_name)
    coin_id = coingecko_map.get(lookup_key, "")
    # ... returns Coingecko ID, NOT token name
```

**Notice**: Coingecko lookup returns ONLY the Coingecko ID, not the token name!

---

## So Where Could Token Name Come From?

### Option 1: Sim Dune API (Currently NOT extracted)
```json
{
  "name": "USD Coin",  ← Available but NOT extracted
  "symbol": "USDC"
}
```

### Option 2: Coingecko API (Currently NOT extracted)
```json
{
  "id": "usd-coin",
  "name": "USD Coin",  ← Available but NOT extracted
  "symbol": "USDC"
}
```

### Option 3: Debank API (Currently NOT extracted)
```json
{
  "name": "USD Coin",  ← Available but NOT extracted
  "symbol": "USDC"
}
```

---

## Current Behavior

**If Sim Dune has token name**: NOT extracted (ignored)
**If Coingecko has token name**: NOT extracted (ignored)
**If Debank has token name**: NOT extracted (ignored)

**Result**: Token name is NEVER included in the output!

---

## To Add Token Name, You Need to:

1. **Extract from Sim Dune API** (easiest):
   ```python
   name = item.get('name', '')
   ```

2. **OR Extract from Coingecko API** (requires additional API call):
   ```python
   # After getting coin_id, fetch from Coingecko
   coingecko_response = requests.get(f"https://api.coingecko.com/api/v3/coins/{coin_id}")
   name = coingecko_response.json().get('name', '')
   ```

3. **OR Extract from Debank API** (requires additional API call):
   ```python
   # For spam tokens, Debank API is already called
   # Just extract name from response
   name = debank_response.get('name', '')
   ```

---

## Recommendation

**Use Sim Dune API** (Option 1) because:
- ✅ Already have the data
- ✅ No additional API calls needed
- ✅ Fastest implementation
- ✅ Preserves all characters and whitespace
- ✅ No rate limiting concerns

**Implementation**:
```python
name = item.get('name', '')  # Extract from Sim Dune
```

Then add to Excel output as Column 3 (after Symbol).


---

# Part B: TronGrid + TronScan Flow (TRX Transaction Comparison)

Used by: `utils/trx_transaction/trongrid_dam_comparison.py`

## Step 3 Excel — TronGrid Transactions Sheet

Each row is one parsed transaction. Data flows from TronGrid API → `parse_tx()` function → Excel.

| Column | Source | Extraction |
|--------|--------|------------|
| # | Sequential | Row counter |
| Trx Hash | TronGrid `txID` | Direct field |
| Date/Time (UTC) | TronGrid `block_timestamp` | `datetime.fromtimestamp(ms/1000, tz=utc)` |
| Transaction Type | TronGrid `raw_data.contract[0].type` | Mapped via `parse_tx()` (see classification table in 03_API_REFERENCE.md §1.2) |
| From | TronGrid `raw_data.contract[0].parameter.value.owner_address` | Hex → Base58 decoded |
| To | TronGrid `raw_data.contract[0].parameter.value.to_address` | Hex → Base58 decoded |
| Amount | TronGrid `parameter.value.amount` | Divided by 1,000,000 for TRX; by `10^decimals` for TRC20 |
| Resources Fee | TronGrid `ret[0].fee` | Formatted as `"{fee/1e6:.6f} TRX ({fee:,} SUN)"` |
| Token Transfer | TRC20 lookup map (joined by `txID`) | Symbol, name, decimals, from, to, contract |
| Net Transfer | Calculated | Direction prefix (`+`/`-`) + amount + symbol |

### TRC20 Lookup Map

Built from TronGrid TRC20 endpoint (§1.3 in API reference):
```
trc20_map = { transaction_id: transfer_object }
```
When `parse_tx()` encounters a `TriggerSmartContract`, it checks `trc20_map[txID]` to determine if it's a token transfer.

---

## Step 7 Excel — DAM Transactions Sheet

Each row is one transaction extracted from DAM UI via Playwright.

| Column | Source | Extraction |
|--------|--------|------------|
| # | Sequential | Row counter |
| Trx Hash | DAM UI table cell | Regex: `[0-9a-fA-F]{6}...` or full 64-char hash |
| Date/Time (DAM) | DAM UI table cell | Regex: `DD/MM/YYYY HH:MM:SS` |
| Transaction Type | DAM UI table cell | Text before datetime in merged cell |
| From | DAM UI table cell | TRON address pattern `T[a-zA-Z0-9]{33}` |
| To | DAM UI table cell | TRON address pattern |
| Amount | DAM UI table cell | Direct text |
| Token Transfer | DAM UI table cell | Direct text |
| Net Transfer | DAM UI table cell | Direct text |
| Raw Cell Data | DAM UI | All `<td>` texts joined by ` | ` |

### DAM Pagination

DAM shows ~25 transactions per page. The script:
1. Reads all `tr[class*='hover']` rows on current page
2. Deduplicates by truncated hash (avoids re-counting when pages overlap)
3. Navigates to next page via pagination input (`input#pagination-input`) or next button
4. Stops when no new rows are found or last page reached

---

## Step 8 Excel — Comparison Sheet

Matches TronGrid and DAM transactions by `trx_hash`.

| Column | Source |
|--------|--------|
| DAM Hash (truncated) | DAM Step 7 data |
| DAM Date/Time | DAM Step 7 data |
| DAM Type / From / To / Amount / Token | DAM Step 7 data |
| TronGrid Hash (full) | TronGrid Step 3 data |
| TronGrid Date (UTC) / Type / From / To / Amount / Token | TronGrid Step 3 data |
| MATCH? | `✅ MATCH` if hash found in both, `⚠️ Not found` otherwise |
| Notes | Timezone explanation for boundary mismatches |

### Hash Matching Logic

DAM often shows truncated hashes like `7fb8aa...44ba1a`. The matching uses:
```python
def norm_hash(h):
    if "..." in h:
        parts = h.split("...")
        return (parts[0].lower(), parts[1].lower())
    return (h[:6].lower(), h[-6:].lower())
```
Both full and truncated forms are indexed for lookup.

### Timezone Note

TronGrid uses **UTC**. DAM uses **UTC+7** (ICT). Transactions near midnight may appear on different calendar dates between the two sources. This is expected and documented in the Summary sheet.

---

# Part C: TRX Balance Flow

Used by: `tests/extraction/test_trx_balance.py`, `utils/trx_transaction/trx_balance_and_dam_extraction.py`

## TRX Balance, Price Sheet

| Column | Source API | Field Path |
|--------|-----------|------------|
| Address | Input | User-provided TRX address |
| Balance | Calculated | `balance_raw / 1,000,000` |
| Balance (Raw) | TronGrid Account | `data[0].balance + frozen + frozenV2 + unfrozenV2 + delegated` |
| Decimal Places | TronScan Token Detail | `trc20_tokens[0].decimals` |
| Contract Token | TronScan Token Detail | `trc20_tokens[0].symbol` |
| Contract Address | TronGrid Account | Key from `trc20[]` dict |
| Contract Balance | Calculated | `contract_balance_raw / 10^decimal_places` |
| Contract Balance (Raw) | TronGrid Account | Value from `trc20[]` dict |
| Price | TronScan Token Detail | `trc20_tokens[0].market_info.priceInUsd` |
| Price (24h) | TronScan Token Detail | `trc20_tokens[0].market_info.gain * 100` |
| Symbol Show | TronScan Token Detail | `trc20_tokens[0].symbolShow` |
