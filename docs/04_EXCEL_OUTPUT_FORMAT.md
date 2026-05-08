# Excel Output Format - DAM Automation v2

## Overview

All extraction tests generate Excel files with standardized formats. This document describes the structure and content of each sheet.

---

## Full Extraction Excel (`DAM_Full_{portfolio_name}_{timestamp}.xlsx`)

### Sheet 1: Overview - Wallet Breakdown
**Purpose**: Token holdings from DAM portfolio

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Token | Text | Token symbol (ETH, USDC, etc.) |
| Percentage | Number | % of total portfolio |
| Percentage Validation | Formula | Validates percentage calculation |
| Net Worth | Currency | USD value of holdings |
| Net Worth Validation | Formula | Validates net worth calculation |

**Example**:
```
Token | Percentage | Percentage Validation | Net Worth | Net Worth Validation
ETH   | 45.5%      | ✓                     | $25,000   | ✓
USDC  | 30.2%      | ✓                     | $16,600   | ✓
```

### Sheet 2: Overview - Token Allocation
**Purpose**: Token allocation breakdown

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Token | Text | Token symbol |
| Percentage | Number | % of total |
| Percentage Validation | Formula | Validation |
| Net Worth | Currency | USD value |
| Net Worth Validation | Formula | Validation |

### Sheet 3: Overview - Chain Allocation
**Purpose**: Blockchain chain allocation

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Chain | Text | Blockchain name (Ethereum, Polygon, etc.) |
| Percentage | Number | % of total |
| Percentage Validation | Formula | Validation |
| Net Worth | Currency | USD value |
| Net Worth Validation | Formula | Validation |

### Sheet 4: Overview - Platform Allocation
**Purpose**: DeFi platform allocation

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Platform | Text | Platform name (Aave, Compound, etc.) |
| Percentage | Number | % of total |
| Percentage Validation | Formula | Validation |
| Net Worth | Currency | USD value |
| Net Worth Validation | Formula | Validation |

### Sheet 5: Overview - Combined Net Worth
**Purpose**: Combined net worth summary

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Section | Text | Section name |
| Category | Text | Category name |
| Token Count | Number | Number of tokens |
| Token Count Validation | Formula | Validation |
| Net Worth | Currency | USD value |
| Net Worth Validation | Formula | Validation |
| Percentage | Number | % of total |
| Percentage Validation | Formula | Validation |

### Sheet 6: Overview - Header & Token Holdings Header
**Purpose**: Header information and token holdings summary

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Section | Text | Section name |
| Category | Text | Category name |
| Token Count | Number | Number of tokens |
| Token Count Validation | Formula | Validation |
| Net Worth | Currency | USD value |
| Net Worth Validation | Formula | Validation |
| Percentage | Number | % of total |
| Percentage Validation | Formula | Validation |

### Sheet 7: DeFi Tab
**Purpose**: DeFi protocol positions

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Protocol | Text | Protocol name (Aave V3, Compound, etc.) |
| Asset | Text | Asset symbol |
| Balance | Number | Token balance |
| Value | Currency | USD value |
| APY | Number | Annual percentage yield |

### Sheet 8: TRX Balance, Price
**Purpose**: TRX balance and price data (if TRX address provided)

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Address | Text | TRX wallet address |
| Balance | Number | TRX balance |
| Price | Currency | TRX price in USD |
| Total Value | Currency | Balance × Price |

### Sheet 9: SimDune (raw)
**Purpose**: Raw SimDune API response

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Address | Text | EVM wallet address |
| Token | Text | Token symbol |
| Balance | Number | Token balance |
| Decimals | Number | Token decimals |
| Contract | Text | Token contract address |

### Sheet 10: Sim Dune - Address Amount
**Purpose**: Extracted SimDune data

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Token | Text | Token symbol |
| Balance | Number | Token balance |
| Value | Currency | USD value |
| Chain | Text | Blockchain name |

### Sheet 11: Rabby Raw
**Purpose**: Raw Rabby API response

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Address | Text | EVM wallet address |
| Protocol | Text | Protocol name |
| Asset | Text | Asset symbol |
| Balance | Number | Asset balance |
| Value | Currency | USD value |

### Sheet 12: Rabby Api Data
**Purpose**: Extracted Rabby protocol data

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Protocol | Text | Protocol name |
| Asset | Text | Asset symbol |
| Balance | Number | Asset balance |
| Value | Currency | USD value |
| Chain | Text | Blockchain name |

---

## DAM Data Extraction Excel (`DAM_Overview_{portfolio_name}_{timestamp}.xlsx`)

### Single Sheet: DAM Overview
**Purpose**: All DAM portfolio data in one sheet

**Sections**:
1. **Header Information**
   - Portfolio name
   - Total net worth
   - Last updated

2. **Token Holdings**
   - Token symbol
   - Balance
   - Value
   - Percentage

3. **Chain Breakdown**
   - Chain name
   - Value
   - Percentage

4. **DeFi Protocols**
   - Protocol name
   - Asset
   - Balance
   - Value

---

## Rabby Protocol Excel (`Protocol_{portfolio_name}_{timestamp}.xlsx`)

### Sheet 1: Rabby Raw
**Purpose**: Raw API response

### Sheet 2: Rabby Api Data
**Purpose**: Extracted protocol data

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Protocol | Text | Protocol name |
| Asset | Text | Asset symbol |
| Balance | Number | Asset balance |
| Value | Currency | USD value |
| Chain | Text | Blockchain name |

---

## TRX Balance Excel (`TRX_Balance_{address}_{timestamp}.xlsx`)

### Sheet 1: TRX Balance, Price
**Purpose**: TRX balance and price

**Columns**:
| Column | Type | Description |
|--------|------|-------------|
| Address | Text | TRX wallet address |
| Balance | Number | TRX balance |
| Price | Currency | TRX price in USD |
| Total Value | Currency | Balance × Price |

### Sheet 2: Token List
**Purpose**: All tokens in TRX ecosystem

### Sheet 3: All Token Info
**Purpose**: Detailed token information

---

## Formatting Standards

### Colors
- **Header Row**: Light blue background
- **Validation Pass**: Green text
- **Validation Fail**: Red text
- **Total Row**: Light gray background

### Number Formats
- **Currency**: `$#,##0.00`
- **Percentage**: `0.00%`
- **Decimal**: `0.00000000` (8 decimals for crypto)

### Column Widths
- **Text Columns**: Auto-fit
- **Number Columns**: 15 characters
- **Currency Columns**: 18 characters

### Formulas
- **Percentage Validation**: `=IF(SUM(B:B)=100%, "✓", "✗")`
- **Net Worth Validation**: `=IF(C2=B2*D2, "✓", "✗")`

---

## Data Validation

### Validation Rules
1. **Percentages**: Must sum to 100%
2. **Net Worth**: Must equal Balance × Price
3. **Addresses**: Must be valid format (0x... or T...)
4. **Decimals**: Must match token decimals

### Error Indicators
- Red cell = Validation failed
- Yellow cell = Warning (unusual value)
- Green cell = Validation passed

---

## Export Options

### CSV Export
```bash
# Convert Excel to CSV
python3 -c "
import pandas as pd
df = pd.read_excel('DAM_Full_portfolio_20250101_120000.xlsx')
df.to_csv('DAM_Full_portfolio_20250101_120000.csv', index=False)
"
```

### JSON Export
```bash
# Convert Excel to JSON
python3 -c "
import pandas as pd
df = pd.read_excel('DAM_Full_portfolio_20250101_120000.xlsx')
df.to_json('DAM_Full_portfolio_20250101_120000.json', orient='records')
"
```

---

## File Locations

All Excel files are saved to:
```
test-results/excel-exports/
├── DAM_Full_portfolio_20250101_120000.xlsx
├── DAM_Overview_portfolio_20250101_120000.xlsx
├── Protocol_portfolio_20250101_120000.xlsx
└── TRX_Balance_address_20250101_120000.xlsx

test-results/API Result/
├── API_TRXBalance_{address}_{timestamp}.xlsx
└── SimDune_{portfolio}_{timestamp}.xlsx

# TRX Transaction Comparison (generated in automationv2 root):
Step3_TronGrid_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
Step7_DAM_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
Step8_Comparison_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
```

---

## TRX Transaction Comparison Excel Files

### Step3: TronGrid Transactions (`Step3_TronGrid_*.xlsx`)

**Sheet 1: TronGrid Transactions**

| Column | Type | Description |
|--------|------|-------------|
| # | Number | Row counter |
| Trx Hash | Text | 64-char transaction hash |
| Date/Time (UTC) | Text | UTC timestamp |
| Transaction Type | Text | Contract type (TRX Transfer, TRC20, Freeze, etc.) |
| From | Text | Sender address (Base58) |
| To | Text | Receiver address (Base58) |
| Amount | Text | Transaction amount with unit |
| Resources Consumed & Fee | Text | Fee in TRX + SUN |
| Token Transfer | Text | TRC20 token details (if applicable) |
| Net Transfer | Text | Direction prefix (+/-) + amount + symbol |

**Sheet 2: TRC20 Transfers** — All TRC20 token transfers with symbol, name, decimals, raw/adjusted values

### Step7: DAM Transactions (`Step7_DAM_*.xlsx`)

Same columns as Step3 but extracted from DAM UI, plus:

| Column | Type | Description |
|--------|------|-------------|
| Raw Cell Data | Text | All `<td>` texts joined by ` | ` |

### Step8: Comparison (`Step8_Comparison_*.xlsx`)

**Sheet 1: Summary & Conclusion** — Match statistics, timezone notes, final verdict

**Sheet 2: Comparison Detail** — Side-by-side TronGrid vs DAM fields matched by trx_hash

| Column | Description |
|--------|-------------|
| DAM Hash / Date / Type / From / To / Amount / Token | DAM extracted data |
| TronGrid Hash / Date / Type / From / To / Amount / Token | TronGrid API data |
| MATCH? | ✅ MATCH or ⚠️ Not found |
| Notes | Timezone boundary explanation |

**Sheet 3: TronGrid Data** — Full TronGrid parsed data
**Sheet 4: DAM Data** — Full DAM extracted data

---

## Troubleshooting

### Issue: Excel file is corrupted
**Solution**: Re-run extraction

### Issue: Formulas not calculating
**Solution**: 
1. Open Excel
2. Press Ctrl+Shift+F9 to recalculate all formulas
3. Save file

### Issue: Data is cut off
**Solution**:
1. Select all columns
2. Auto-fit column width
3. Save file

### Issue: Validation shows errors
**Solution**:
1. Check data accuracy
2. Verify calculations
3. Review source data
