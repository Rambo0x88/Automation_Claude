# Field Extraction Sources & Calculation Guide

This document explains where each field is extracted from in the DAM UI and how validation/calculations are performed.

---

## Overview - Header & Token Holdings Header Sheet

### Token Count (Column C)

**Source**: Combined Net Worth section in DAM UI

**Extraction Method**:
- Scans all `div` and `span` elements on the page
- Looks for text pattern: `"ChainName (TokenCount)"`
- Examples: `"Tron (35)"`, `"Ethereum (17)"`, `"Base (8)"`

**Regex Pattern**:
```python
match = re.match(r'^([A-Za-z\s]+)\s*\((\d+)\)$', text)
# Group 1: Chain name (e.g., "Tron")
# Group 2: Token count (e.g., "35")
```

**Validation (Column D - Token Count Validation)**:

Compares UI token count against actual token count calculated from data:

1. **Count from Overview - Wallet sheet**
   - Counts unique tokens in that chain's wallet rows (Column B)
   - Example: Ethereum has 12 unique tokens

2. **Count from Overview - De-Fi sheet**
   - Counts unique tokens in that chain's DeFi rows (Column D)
   - Example: Ethereum has 8 unique tokens

3. **Union the two sets** (remove duplicates)
   - If same token appears in both wallet and DeFi, count once
   - `actual_count = len(wallet_tokens | defi_tokens)`
   - Example: 12 + 8 - 2 overlaps = 18 unique tokens

4. **Compare with UI token count**
   - If `actual_count == UI_token_count` → **"Passed"**
   - If `actual_count != UI_token_count` → **"Failed"**
   - If row is not "Token Holdings - Chain" → **"Not Applicable"**

**Chain Name Mapping**:
The script maps short chain names to full names:
- `"bnb"`, `"bsc"` → `"Binance Smart Chain"`
- `"eth"` → `"Ethereum"`
- `"trx"` → `"Tron"`
- `"matic"` → `"Polygon"`
- `"arb"` → `"Arbitrum"`
- etc.

**De-Fi Chain Code Mapping**:
Maps full chain names to De-Fi sheet chain codes:
- `"Binance Smart Chain"` → `"bsc"`
- `"Ethereum"` → `"eth"`
- `"Tron"` → `"tron"`
- `"Polygon"` → `"matic"`
- etc.

---

### Net Worth (Column E)

**Source**: Combined Net Worth section in DAM UI

**Extraction Method**:
- Extracts from parent element text after chain name pattern
- Looks for dollar value with `$` prefix
- Fallback: looks for numbers with commas (likely currency values)
- Preserves `<` prefix for small values (e.g., `"<$0.01"`)

**Regex Patterns**:
```python
# Primary: Look for $ prefix
value_match = re.search(r'\$([\d,]+\.?\d*)', cleaned_text)

# Fallback: Look for numbers with commas
value_match = re.search(r'([\d,]{4,}\.?\d*)', cleaned_text)
```

**Validation (Column F - Net Worth Validation)**:

Compares UI net worth against calculated net worth from wallet and DeFi data:

1. **Sum from Overview - Wallet sheet**
   - Sums all net worth values for that chain (Column E)
   - Example: Ethereum wallet total = $15,000

2. **Sum from Overview - De-Fi sheet**
   - Sums all net worth values for that chain (Column E)
   - Example: Ethereum DeFi total = $8,500

3. **Calculate total**
   - `actual_net_worth = wallet_sum + defi_sum`
   - Example: $15,000 + $8,500 = $23,500

4. **Compare with UI net worth**
   - If `actual_net_worth ≈ UI_net_worth` → **"Passed"**
   - If `actual_net_worth ≠ UI_net_worth` → **"Failed"**
   - Allows small tolerance for rounding differences

---

### Percentage (Column G)

**Source**: Combined Net Worth section in DAM UI

**Extraction Method**:
- Extracts from parent element text
- Looks for percentage pattern with optional `<` symbol
- Examples: `"0.01%"`, `"<0.01%"`, `"< 0.01%"`, `"50.25%"`

**Regex Pattern**:
```python
# Match percentage with optional < symbol and optional space
pct_match = re.search(r'(<\s*[\d.]+|[\d.]+)%', parent_text)
percentage = pct_match.group(1).replace(' ', '')  # Remove space if present
```

**Validation (Column H - Percentage Validation)**:

Compares UI percentage against calculated percentage from net worth:

1. **Calculate percentage**
   - `calculated_pct = (chain_net_worth / total_net_worth) * 100`
   - Example: ($23,500 / $100,000) * 100 = 23.5%

2. **Compare with UI percentage**
   - If `calculated_pct ≈ UI_pct` → **"Passed"**
   - If `calculated_pct ≠ UI_pct` → **"Failed"**
   - Allows small tolerance for rounding differences

---

## Overview - Wallet Sheet

### Chain (Column A)

**Source**: Overview - Wallet section in DAM UI

**Extraction Method**:
- Extracts from wallet breakdown table
- Each row represents a token holding on a specific chain
- Chain name is extracted from the row's chain indicator

---

### Token (Column B)

**Source**: Overview - Wallet section in DAM UI

**Extraction Method**:
- Extracts token symbol from wallet table
- Examples: `"ETH"`, `"USDC"`, `"DAI"`, `"WBTC"`

---

### Price (Column C)

**Source**: DAM UI token price display

**Extraction Method**:
- Hovers over price cell to trigger tooltip
- Reads tooltip content for full price value
- Fallback: reads price from cell text if tooltip unavailable

**Tooltip Extraction**:
- Looks for `data-tooltip-id` attribute on price element
- Hovers to make tooltip visible
- Reads tooltip div content by ID

**Validation (Column D - Price Validation)**:
- Compares UI price against CoinGecko API price
- If `|UI_price - API_price| < tolerance` → **"Passed"**
- If difference exceeds tolerance → **"Failed"**

---

### Price (24h) (Column H)

**Source**: DAM UI price change indicator

**Extraction Method**:
- Looks for percentage change element with color class
- Red/error class = negative change (prefixed with `-`)
- Green/success class = positive change
- Extracts percentage value

**Regex Pattern**:
```python
pct_match = re.search(r'([\d.]+)%?', pct_text)
pct_value = pct_match.group(1)
# Check if negative (red/error color)
if 'error' in pct_class.lower() or '↓' in pct_text:
    price_24h = f"-{pct_value}"
else:
    price_24h = pct_value
```

---

### Amount (Column K)

**Source**: DAM UI token amount display

**Extraction Method**:
- Extracts from amount cell in wallet table
- Removes token symbol suffix (e.g., "1.5 ETH" → "1.5")
- Hovers to get tooltip for full precision value

**Regex Pattern**:
```python
# Remove token symbol at end
amount = re.sub(r'[A-Z]+$', '', amount_cell_text).strip()
```

**Tooltip Extraction**:
- Looks for `data-tooltip-id` attribute on amount element
- Hovers to make tooltip visible
- Reads tooltip div content by ID
- Cleans up multiple decimals if present

---

### Amount Tooltip (Column L)

**Source**: DAM UI amount tooltip

**Extraction Method**:
- Hovers over amount cell to trigger tooltip
- Reads full precision amount from tooltip
- Cleans up formatting (removes commas, extra decimals)

**Fallback Methods** (if tooltip not visible):
1. Try direct ID lookup: `document.getElementById(tooltip_id)`
2. Try data-tooltip-id attribute lookup
3. Find any visible tooltip on page with numeric content
4. Find element with numeric/decimal content matching pattern

---

### Value (Column U)

**Source**: DAM UI calculated value (Amount × Price)

**Extraction Method**:
- Extracts from value cell in wallet table
- Removes currency symbols (`$`, commas)
- Cleans up formatting

**Calculation Validation**:
- `calculated_value = amount × price`
- Compares with UI value
- If `|calculated_value - UI_value| < tolerance` → **"Passed"**

---

## Overview - De-Fi Sheet

### Chain (Column B)

**Source**: De-Fi protocol section in DAM UI

**Extraction Method**:
- Extracts chain code from De-Fi protocol row
- Examples: `"eth"`, `"bsc"`, `"matic"`, `"arb"`

---

### Position Type (Column A)

**Source**: De-Fi protocol section in DAM UI

**Extraction Method**:
- Extracts position type from protocol row
- Examples: `"Supply"`, `"Borrow"`, `"Stake"`, `"LP"`

---

### Pool/Position (Column D)

**Source**: De-Fi protocol section in DAM UI

**Extraction Method**:
- Extracts pool or position pair from protocol row
- Examples: `"USDC"`, `"ETH/USDC"`, `"WETH"`, `"DAI"`

---

### Amount (Column E)

**Source**: De-Fi protocol section in DAM UI

**Extraction Method**:
- Extracts amount from protocol row
- Hovers to get tooltip for full precision value
- Handles "Borrow" positions (prefixed with `-`)

**Tooltip Extraction**:
- Looks for tooltip trigger in amount cell
- Hovers to make tooltip visible
- Reads tooltip content
- Extracts numeric value using regex: `r'(-?[\d,]+\.?\d*)'`

**Borrow Detection**:
- If tooltip contains `"Borrow:"` → prefixes amount with `-`
- Example: `"Borrow: 37376.113827 WETH"` → `"-37376.113827"`

---

### Value (Column F)

**Source**: De-Fi protocol section in DAM UI

**Extraction Method**:
- Extracts USD value from protocol row
- Removes currency symbols and commas
- Cleans up formatting

---

## Tooltip Extraction Strategy

### Why Some Tooltips Don't Show

Tooltips may not be visible because:
1. Trigger element is hidden (`display: none`)
2. Trigger element is outside viewport
3. Trigger element has `visibility: hidden`
4. Trigger element has `opacity: 0`

### Solution: Multi-Level Fallback

The script uses 4 fallback methods to extract tooltips:

**Method 1: Direct ID Lookup**
```javascript
const tooltip = document.getElementById(tooltip_id);
if (tooltip && tooltip.textContent.trim().length > 0) {
    return tooltip.textContent.trim();
}
```

**Method 2: Data-Tooltip-ID Attribute**
```javascript
const elem = document.querySelector('[data-tooltip-id="' + id + '"]');
if (elem && elem.textContent.trim().length > 0) {
    return elem.textContent.trim();
}
```

**Method 3: Find Any Visible Tooltip**
```javascript
const tooltips = document.querySelectorAll('[role="tooltip"], [class*="tooltip"], [id*="tooltip"]');
for (let tooltip of tooltips) {
    if (tooltip.style.display !== 'none' && tooltip.textContent.trim().length > 0) {
        return tooltip.textContent.trim();
    }
}
```

**Method 4: Find Numeric Content**
```javascript
// Find elements with numeric/decimal content
const allElements = document.querySelectorAll('*');
for (let elem of allElements) {
    const text = elem.textContent.trim();
    if (text.match(/^[0-9\.\-\+\s%<>]+$/) && text.length < 100) {
        return text;
    }
}
```

### Retry Logic

- Per-tooltip retry: 3 attempts with increasing wait times (400ms, 600ms, 800ms)
- Per-row retry: 2 attempts with 1 second wait between retries
- If any tooltip in a row is empty, entire row extraction retries

---

## Validation Tolerance

### Rounding Tolerance

Different fields have different tolerance levels:

| Field | Tolerance | Reason |
|-------|-----------|--------|
| Price | ±0.01% | API prices may vary slightly |
| Amount | ±0.00000001 | Blockchain precision |
| Value | ±$0.01 | Rounding differences |
| Percentage | ±0.01% | Rounding differences |
| Net Worth | ±$1 | Sum of rounded values |

### Validation Result Mapping

| Condition | Result |
|-----------|--------|
| Difference within tolerance | **"Passed"** |
| Difference exceeds tolerance | **"Failed"** |
| Field not applicable to row | **"Not Applicable"** |
| Data missing or error | **"Error"** |

---

## Common Extraction Issues

### Issue: Tooltip shows "Tooltip N/A, cant compare"

**Cause**: Tooltip trigger element is not visible or tooltip div not found

**Solution**: 
- Script uses 4 fallback methods to extract tooltips
- If all methods fail, marks as "Tooltip N/A, cant compare"
- This is expected for some tokens with rendering issues

### Issue: Token Count Validation shows "Failed"

**Cause**: Mismatch between UI token count and actual token count

**Possible Reasons**:
1. Token appears in both wallet and DeFi (counted once in union)
2. Token was added/removed after UI was rendered
3. Token is hidden or filtered in UI
4. Extraction missed some tokens

**Debug**: Check wallet and DeFi sheets for token counts

### Issue: Net Worth Validation shows "Failed"

**Cause**: Mismatch between UI net worth and calculated net worth

**Possible Reasons**:
1. Price data is stale or incorrect
2. Amount extraction missed decimal places
3. Rounding differences accumulated
4. Token value changed during extraction

**Debug**: Check individual token values and prices

---

## Field Extraction Summary Table

| Sheet | Column | Source | Extraction Method | Validation |
|-------|--------|--------|-------------------|-----------|
| Header & Holdings | Token Count | UI Combined Net Worth | Regex pattern match | Count union of wallet + DeFi |
| Header & Holdings | Net Worth | UI Combined Net Worth | Dollar value extraction | Sum of wallet + DeFi values |
| Header & Holdings | Percentage | UI Combined Net Worth | Percentage extraction | Calculated from net worth |
| Wallet | Chain | UI Wallet Table | Table cell extraction | N/A |
| Wallet | Token | UI Wallet Table | Table cell extraction | N/A |
| Wallet | Price | UI Tooltip | Hover + tooltip read | Compare with CoinGecko |
| Wallet | Amount | UI Tooltip | Hover + tooltip read | Calculated value check |
| Wallet | Value | UI Table Cell | Cell text extraction | Amount × Price |
---

## Overview - Combined Net Worth Sheet

### Value (Column B)

**Source**: DAM UI Combined Net Worth section

**Extraction Method**:
- Extracts USD value per address/exchange row from the Combined Net Worth table

---

### Value Validation (Column C)

**Validation Logic**: Compares DAM UI value (col B) against Calculated Value (col D)

- If `|DAM Value - Calculated Value| / Calculated Value ≤ 1%` → **"Passed"**
- If difference exceeds 1% → **"Failed"**
- If Calculated Value = 0 and DAM Value = 0 → **"Passed"**
- If Calculated Value = 0 and DAM Value ≠ 0 → **"Failed"**

---

### Calculated Value (Column D)

**Source**: Computed from API data sheets, varies by address type:

#### Tron Address (starts with `T`, 34 chars)

Sums from **TRX Balance, Price** sheet, filtered by matching address (col A):

```
Calculated Value = SUM( Balance_Raw / (10 ^ Decimal_Places) × Price )
```

- Col B = Decimal Places
- Col F = Balance (Raw)
- Col G = Price

#### EVM Address (starts with `0x`, 42 chars)

Sums two sources, filtered by matching address (col A):

1. **Sim + Coingecko + Debank API** sheet — col K (Calculated Price)
2. **Rabby Api Data** sheet — col M (Calculated Value)

```
Calculated Value = SUM(SimDune col K) + SUM(Rabby col M)
```

#### Exchange (name string, e.g. "Binance", "moontest")

Reads the **Total row** from the matching exchange sheet:

- Finds row where col P = "TOTAL" (case-insensitive)
- Takes value from col Q of that row

```
Calculated Value = Exchange sheet Q column value at "TOTAL" row
```

---

## Overview - Token Allocation Sheet

### Column Structure

| Col | Field | Source |
|-----|-------|--------|
| A | Token | UI Token Allocation table |
| B | Percentage | UI Token Allocation table |
| C | Percentage Validation | Calculated |
| D | Net Worth | UI Token Allocation table |
| E | Net Worth Validation | (reserved) |

### Percentage Validation (Column C)

**Formula**: `TRUNC(D / SUM(all D) * 100, 2) == B` → **"Passed"**, else **"Failed"**

- Total net worth = SUM of all rows in col D
- Calculated percentage = `row_D / total_D * 100`, truncated to 2 decimal places
- Special case: if calculated % < 0.01 → B must equal `"<0.01"` or `"< 0.01"`

---

## Overview - Chain Allocation Sheet

### Column Structure

| Col | Field | Source |
|-----|-------|--------|
| A | Chain | UI Chain Allocation table |
| B | Percentage | UI Chain Allocation table |
| C | Percentage Validation | Calculated |
| D | Net Worth | UI Chain Allocation table |
| E | Net Worth Validation | (reserved) |

### Percentage Validation (Column C)

Same formula as Token Allocation:

**Formula**: `TRUNC(D / SUM(all D) * 100, 2) == B` → **"Passed"**, else **"Failed"**

---

## Overview - Platform Allocation Sheet

### Column Structure

| Col | Field | Source |
|-----|-------|--------|
| A | Platform | UI Platform Allocation table |
| B | Percentage | UI Platform Allocation table |
| C | Percentage Validation | Calculated |
| D | Net Worth | UI Platform Allocation table |
| E | Net Worth Validation | (reserved) |

### Percentage Validation (Column C)

Same formula as Token Allocation:

**Formula**: `TRUNC(D / SUM(all D) * 100, 2) == B` → **"Passed"**, else **"Failed"**

