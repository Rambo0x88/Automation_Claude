# Commands & Workflows - DAM Automation v2

## Setup (One-Time)
```bash
cd core/projects/DAM/automationv2
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
cp .env.example .env
```

---

# WORKFLOWS & COMMANDS

## Workflow 1: DAM UI Automation Tests

**Purpose:** Test DAM UI flows — authentication, portfolio management, and new user setup.

**Time:** ~2-15 minutes depending on scope

### Run All UI Tests
```bash
cd core/projects/DAM/automationv2
source venv/bin/activate
pytest tests/extraction/ -v --headed
```

### 1a. Authentication (Sign-up, Sign-in, Account Settings)
```bash
pytest tests/extraction/test_sign_up.py -v
pytest tests/extraction/test_sign_in_after_signup.py -v
pytest tests/extraction/test_account_settings.py -v
```

### 1b. Portfolio Management
```bash
pytest tests/extraction/test_portfolio.py -v --browser=chromium
pytest tests/extraction/test_portfolio.py -v --browser=chromium --portfolio "portfolio_name"
pytest tests/extraction/test_portfolio.py -v --headed --portfolio "my_portfolio"
```

### 1b-2. Connect Exchange to Portfolio

**Purpose:** Create a CEX-only portfolio and connect exchange accounts (Binance/BIT).

**Script:** `create_cex_portfolio.py` (standalone) or `utils/exchange_connector.py` (library)

**Config:** Exchange API keys stored in `test_data/exchange_keys.json`

**Usage:**
```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# Create portfolio with both Binance + BIT exchanges
TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py -e binance bit

# Custom portfolio name
TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py --name "My CEX Portfolio" -e binance bit

# Specific exchange key
TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py -e bit --key "david"

# Binance only (default)
TEST_EMAIL=lily.su@merquri.io TEST_PASSWORD='Orion888888!' python3 create_cex_portfolio.py
```

**Flow:**
1. Sign in with given credentials
2. Navigate to Create Portfolio page (handles new account + existing account)
3. Enter portfolio name
4. Click "Exchange" tab
5. Connect exchange(s) via modal (Binance / BIT)
6. Tick all exchange account checkboxes
7. Click Save

**Screenshots:** Saved to `test-results/screenshots/`

**Usage (from Python):**
```python
from utils.exchange_connector import add_exchanges_to_portfolio, connect_exchange_from_config

# Connect all configured Binance keys (skips duplicates)
add_exchanges_to_portfolio(page, exchanges=["binance"])

# Connect a specific key
connect_exchange_from_config(page, exchange="binance", key_name="moon api key")
```

**Exchange Keys Config** (`test_data/exchange_keys.json`):
```json
{
  "binance": [
    {"display_name": "moon api key", "api_key": "...", "secret_key": "..."},
    {"display_name": "xg", "api_key": "...", "secret_key": "..."},
    {"display_name": "david", "api_key": "...", "secret_key": "..."}
  ],
  "bit": []
}
```

**Functions:**

| Function | Description |
|----------|-------------|
| `load_exchange_keys()` | Load API keys from `test_data/exchange_keys.json` |
| `click_exchange_tab(page)` | Click "Exchange(0)" tab |
| `get_existing_exchanges(page)` | List already-connected exchanges |
| `connect_exchange(page, exchange, display_name, api_key, secret_key)` | Connect one exchange via modal |
| `connect_exchange_from_config(page, exchange, key_name)` | Connect using config file key |
| `connect_all_exchanges(page, exchange)` | Connect all keys, skip duplicates |
| `add_exchanges_to_portfolio(page, exchanges)` | Full flow: tab → check → connect all |

### 1c. New User Setup (Sign-up + Portfolio Creation)
```bash
pytest tests/extraction/test_portfolio_new_user.py -v --browser=chromium
```

**This will:**
1. Create new user account
2. Save credentials to test_data/tc1_account.json
3. Create portfolios from Excel
4. Generate report

### Output
```
test-results/reports/report_{YYYYMMDD_HHMM}.html
test-results/screenshots/
```
- New user credentials saved to `test_data/tc1_account.json` (1c only)


---

## Workflow 2: Data Extraction & Comparison

**Purpose:** Extract data from DAM UI, external APIs, and compare results.

**Time:** ~3-30 minutes depending on scope

---

### 2A. DAM UI Extraction

Scrape data directly from DAM portfolio pages via Playwright (no external API calls).

| # | DAM Page | Script | Status |
|---|----------|--------|--------|
| 2A-1 | Overview | `test_dam_data_extraction.py` | ✅ |
| 2A-2 | Report (EOD Balance, Movement Analysis, Allocation) | — | ⏳ |
| 2A-3 | Transactions | `dam_transaction_extractor.py` | ✅ |

#### 2A-1. Overview UI Extraction

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

pytest tests/extraction/test_dam_data_extraction.py -v -s --headed
pytest tests/extraction/test_dam_data_extraction.py -v -s --headed --portfolio "multi_eb2eb5c6_v3"
pytest tests/extraction/test_dam_data_extraction.py -v -s --headed --portfolioId "83081753-e3af-440a-9081-740120c3840d"
```

**Output:**
```
test-results/excel-exports/DAM_Overview_{portfolio_name}_{timestamp}.xlsx
test-results/screenshots/DAM_Overview_{portfolio_name}_{timestamp}.png
```

#### 2A-2. Report Page UI Extraction ⏳

**Status:** ⏳ Not yet implemented — no script or test exists for this page yet.

#### 2A-3. Transaction UI Extraction

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

python3 utils/trx_transaction/dam_transaction_extractor.py                                # Default portfolio ID
python3 utils/trx_transaction/dam_transaction_extractor.py trx2_Mkx                       # By portfolio name
python3 utils/trx_transaction/dam_transaction_extractor.py -p "My Portfolio"               # By portfolio name
python3 utils/trx_transaction/dam_transaction_extractor.py --id 8724c50c-...               # By portfolio ID
python3 utils/trx_transaction/dam_transaction_extractor.py --date 2026-03-15               # Specific date
python3 utils/trx_transaction/dam_transaction_extractor.py -p "trx2_Mkx" --date 2026-01-21 --xlsx
```

**Output:**
```
test-results/dam_transactions_detailed.json
test-results/DAM_Transactions_{date}_{timestamp}.xlsx  (with --xlsx flag)
test-results/screenshots/
```


---

### 2B. External Data Extraction (API / Blockchain Explorer)

Fetch data from external APIs — no DAM UI involved.

| # | Source | Script | Status |
|---|--------|--------|--------|
| 2B-1 | EVM — SimDune + Coingecko + Debank | Part of `run_overview.py` | ✅ |
| 2B-2 | TRX — TRX Balance API | Part of `run_overview.py` | ✅ |
| 2B-3 | TRX — TronGrid Transaction API | `trongrid_fetcher.py` (standalone) or part of `trongrid_dam_comparison.py` | ✅ |
| 2B-4 | EVM — BscScan / Etherscan Transaction API | — | ⏳ |
| 2B-5 | Rabby Protocol + Hyperliquid DeFi | `test_rabby_protocol.py::test_rabby_export` | ✅ |

#### 2B-5. Rabby Protocol Export (API only, no browser)

```bash
pytest tests/extraction/test_rabby_protocol.py::test_rabby_export -v -s
```

**Output:**
```
test-results/excel-exports/Protocol_{portfolio_name}_{timestamp}.xlsx
```

> Note: 2B-1, 2B-2 are not standalone scripts — they run as part of the comparison workflows in 2C.
> 2B-3 (TronGrid) can be run standalone — see below.

#### 2B-3. TronGrid Transaction Extraction (Standalone)

Fetch and parse TRX transactions from TronGrid API only — no DAM UI involved.

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# Single date
python3 -m utils.trx_transaction.trongrid_fetcher TFEC8v19pjw7bWU8umKvzqvPtFo83cqMkx 16042026

# Date range
python3 -m utils.trx_transaction.trongrid_fetcher TU6wEHdYzUrBQB8bSagTeCbFE29oXFLy7C 02032026 31032026
```

**Output:**
```
test-results/Step3_TronGrid_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
```

#### 2B-3b. DAM Transaction Extraction (Standalone, via Playwright)

Extract transactions from DAM UI only — no TronGrid API involved.

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# By portfolio name + date
python3 -m utils.trx_transaction.dam_extractor trx2_Mkx 16042026

# By TRX address + date range
python3 -m utils.trx_transaction.dam_extractor TFEC8v19pjw7bWU8umKvzqvPtFo83cqMkx 01032026 31032026
```

**Output:**
```
test-results/Step7_DAM_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
test-results/screenshots/{label}/dam_*.png
```


---

### 2C. DAM vs External Data Comparison

Compare DAM UI data with external API data side-by-side.

| # | Comparison | Script | Status |
|---|-----------|--------|--------|
| 2C-1 | Overview + EVM (SimDune/Rabby) | `run_overview.py` | ✅ |
| 2C-2 | Overview + TRX (TRX Balance API) | `run_overview.py` | ✅ |
| 2C-3 | Transaction + TronGrid | `trongrid_dam_comparison.py` | ✅ |
| 2C-4 | Transaction + BscScan/Etherscan | — | ⏳ |
| 2C-5 | Full Portfolio — TRX (Overview + TronGrid Transaction) | `run_all.py --trx` | ✅ |
| 2C-6 | Full Portfolio — EVM (Overview + BscScan/Etherscan Transaction) | `run_all.py --evm` | ⏳ (transaction part) |
| 2C-7 | Static Comparison (offline, no API) | `static_comparison_generator.py` | ✅ |

#### 2C-1 & 2C-2. Overview + API Comparison

Scrapes DAM Overview page + fetches SimDune/Rabby/TRX APIs → Excel with validation columns.

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

python3 run_overview.py                                        # Config defaults
python3 run_overview.py -p "zg's address - 1"    # By portfolio name
python3 run_overview.py -p A_b    # By portfolio name
python3 run_overview.py --evm 0x4e14...                        # EVM only
python3 run_overview.py --trx TUqEg3...                        # TRX only
python3 run_overview.py --trx TUqEg3... --evm 0x4e14...       # Both
```

**API Behavior:**

| Mode | APIs Run | APIs Skipped |
|------|----------|--------------|
| **EVM only** (`--evm`) | SimDune, Rabby Protocol, DAM Portfolio | TRX Balance |
| **TRX only** (`--trx`) | TRX Balance, DAM Portfolio | SimDune, Rabby Protocol |
| **Both** (`--trx` + `--evm`) | All APIs + DAM Portfolio | None |
| **CEX only** (no addresses) | DAM Portfolio only | All external APIs |

**Auto-Portfolio Creation:** If the portfolio doesn't exist in DAM, it will be automatically created.

**Output:**
```
test-results/excel-exports/DAM_Full_{portfolio_name}_{timestamp}.xlsx
test-results/screenshots/DAMSS_{portfolio}_{email}_{MMDD_HHMM}/
```

#### 2C-3. Transaction + TronGrid Comparison

Full pipeline: TronGrid API → DAM Transactions UI → single comparison Excel.

The pipeline is split into modular scripts that can run independently or together:

| Module | Purpose | Can run standalone? |
|--------|---------|---------------------|
| `shared.py` | Config, helpers, date parsing, Excel styling | No (library) |
| `trongrid_fetcher.py` | Step 3 — TronGrid API fetch + parse | ✅ Yes |
| `dam_extractor.py` | Steps 4-7 — DAM UI Playwright extraction | ✅ Yes |
| `trongrid_dam_comparison.py` | Orchestrator — runs both + builds comparison Excel | ✅ Yes |

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# TRX address + single date (DDMMYYYY format)
python3 utils/trx_transaction/trongrid_dam_comparison.py TFEC8v19pjw7bWU8umKvzqvPtFo83cqMkx 16042026

# TRX address + date range
python3 utils/trx_transaction/trongrid_dam_comparison.py TU6wEHdYzUrBQB8bSagTeCbFE29oXFLy7C 02032026 31032026

# Portfolio name + single date
python3 utils/trx_transaction/trongrid_dam_comparison.py trx2_Mkx 16042026

# Portfolio name + date range
python3 utils/trx_transaction/trongrid_dam_comparison.py trx2_Mkx 01032026 31032026
```

**Timezone handling:** The script auto-detects your system timezone. If you're in Singapore (UTC+8), `16042026` means April 16 00:00:00 SGT to 23:59:59 SGT. In Vietnam (UTC+7), the same date uses ICT boundaries. The UTC range used for TronGrid API calls is logged at startup.

**Pipeline Steps:**

| Step | Description |
|------|-------------|
| Steps 4-7 | Login to DAM, navigate to portfolio, apply date filter, extract rows |
| Steps 1-2 | Fetch transactions + TRC20 transfers from TronGrid API |
| Step 3 | Parse transactions, classify inflow/outflow |
| Step 8 | Match by trx_hash, compare fields → single Excel |

**When running the full comparison, only one Excel is produced** (the comparison file). It contains all data:

**Output:**
```
test-results/Comparison_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
utils/trx_transaction/test-results/screenshots/{label}/dam_*.png
utils/trx_transaction/test-results/{label}_raw_data.json
```

**Comparison Excel sheets:**
- `Summary & Conclusion` — match counts, timezone info, verdict
- `Comparison Detail` — side-by-side DAM vs TronGrid per transaction
- `TronGrid Data` — full TronGrid parsed transactions
- `DAM Data` — full DAM scraped transactions

**To run Step 3 or Step 7 individually** (produces its own standalone Excel):
```bash
# Step 3 only — TronGrid data
python3 -m utils.trx_transaction.trongrid_fetcher <TRX_ADDRESS> <DDMMYYYY>

# Step 7 only — DAM UI data
python3 -m utils.trx_transaction.dam_extractor <PORTFOLIO_NAME_OR_ADDRESS> <DDMMYYYY>
```

**Inflow/Outflow Classification (Step 3):**

| Contract Type | Inflow (+) | Outflow (-) |
|---------------|-----------|------------|
| TransferContract (TRX) | `+amount TRX` (receive) | `-amount TRX` (send) |
| TRC20 Transfer (tokens) | `+amount TOKEN` (receive) | `-amount TOKEN` (send) |
| FreezeBalanceV2 | — | `Staked: amount TRX` |
| UnfreezeBalanceV2 | `+amount TRX (unfrozen)` | — |
| DelegateResource | — | `Delegated amount TRX` |
| UnDelegateResource | `Undelegated amount TRX` | — |
| WithdrawExpireUnfreeze | `+TRX withdrawn from stake` | — |
| WithdrawBalance (reward) | `+TRX reward claimed` | — |
| VoteWitnessContract | `Vote cast` | — |

#### 2C-4. Transaction + BscScan/Etherscan Comparison ⏳

**Status:** ⏳ Not yet implemented

#### 2C-5. Full Portfolio — TRX (Overview + TronGrid Transaction + Comparison)

Runs Overview extraction + TRX API comparison + TronGrid transaction comparison in one go.

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# TRX address — overview + transactions (today)
python3 run_all.py --trx TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6

# TRX address + specific date
python3 run_all.py --trx TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6 --date 16042026

# TRX address + date range
python3 run_all.py --trx TUqEg3... --date 01032026 --date-to 31032026

# By portfolio name
python3 run_all.py -p trx2_Mkx --date 16042026
```

**What it runs:**
1. `run_overview.py` — DAM Overview extraction + TRX Balance API comparison
2. `trongrid_dam_comparison.py` — TronGrid vs DAM transaction comparison

**Output:** 2 separate Excel files (Overview + Transaction) + screenshots

#### 2C-6. Full Portfolio — EVM (Overview + BscScan/Etherscan Transaction + Comparison)

Runs Overview extraction + EVM API comparison + blockchain explorer transaction comparison.

```bash
# EVM address — overview + transactions
python3 run_all.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab

# Both EVM + TRX
python3 run_all.py --trx TUqEg3... --evm 0x4e14... --date 16042026
```

**What it runs:**
1. `run_overview.py` — DAM Overview extraction + SimDune/Rabby API comparison
2. EVM transaction comparison — ⏳ BscScan/Etherscan part not yet implemented

**Output:** Overview Excel + screenshots (transaction comparison pending)

#### 2C-7. Static Comparison Excel (No API Calls)

Generate comparison Excel using hardcoded data. No API or Playwright calls.

```bash
python3 utils/trx_transaction/static_comparison_generator.py
```

**Output:**
```
Step3_TronGrid_Transactions_2026-01-21.xlsx
Step7_DAM_Transactions_2026-01-21.xlsx
Step8_Comparison_TronGrid_vs_DAM_2026-01-21_{timestamp}.xlsx
```


---

## Workflow 3: Daily Testing Routine

**Purpose:** Run daily verification tests.

**Time:** ~15-30 minutes

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# 1. Run UI tests
pytest tests/extraction/ -v

# 2. Run extraction tests
pytest tests/extraction/ -v

# 3. View results
open test-results/reports/report_*.html
```

---

## Workflow 4: Debugging Failed Tests

**Purpose:** Debug and troubleshoot test failures.

**Time:** Variable

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate

# Visible browser + slow motion
pytest tests/extraction/test_portfolio.py -v --headed --slowmo=1000

# Tracing for time-travel debugging
pytest tests/extraction/test_portfolio.py -v --tracing=retain-on-failure
playwright show-trace test-results/trace_*.zip

# Playwright Inspector
PWDEBUG=1 pytest tests/extraction/test_portfolio.py -v
```

---

## Workflow 5: Run All Tests

**Purpose:** Full regression testing.

**Time:** ~30-60 minutes

```bash
cd core/projects/DAM/automationv2
source venv/bin/activate
pytest tests/ -v
```

---

# PYTEST OPTIONS

## Browser Control
```bash
pytest tests/ -v -s --headed
pytest tests/ -v --browser=chromium
pytest tests/ -v --browser=firefox
pytest tests/ -v --browser=webkit
```

## Execution Control
```bash
pytest tests/ -v --headed --slowmo=500
pytest tests/ -v -x                        # Stop on first failure
pytest tests/ -v -m smoke                  # Run by marker
pytest tests/ --collect-only               # List tests without running
pytest tests/ -v --durations=10            # Show slowest tests
```

## Debugging & Tracing
```bash
pytest tests/ -v --tracing=retain-on-failure
playwright show-trace test-results/trace_*.zip
PWDEBUG=1 pytest tests/extraction/test_dam_data_extraction.py -v
```

---

# QUICK REFERENCE TABLE

| Workflow | Command | Time | Output |
|----------|---------|------|--------|
| **1. UI Automation** | | | |
| Auth Tests | `pytest tests/extraction/test_sign_*.py -v` | 2-5m | Report |
| Portfolio Tests | `pytest tests/extraction/test_portfolio.py -v` | 2-5m | Report |
| Connect Exchange | `TEST_EMAIL=... python3 create_cex_portfolio.py -e binance bit` | 1-2m | Portfolio + Exchange |
| New User Setup | `pytest tests/extraction/test_portfolio_new_user.py -v` | 10-15m | New Account |
| All UI Tests | `pytest tests/extraction/ -v --headed` | 2-15m | Test Report |
| **2A. DAM UI Extraction** | | | |
| Overview UI Extraction | `pytest tests/extraction/test_dam_data_extraction.py -v` | 3-5m | Excel |
| Report Page UI Extraction | ⏳ Not yet implemented | — | — |
| Transaction UI Extraction | `python3 utils/trx_transaction/dam_transaction_extractor.py -p "name" --xlsx` | 3-10m | Excel + JSON |
| **2B. External Data Extraction** | | | |
| EVM (SimDune) | Part of `run_overview.py` | — | — |
| TRX Balance | Part of `run_overview.py` | — | — |
| TronGrid Transactions | `python3 -m utils.trx_transaction.trongrid_fetcher <ADDR> <DATE>` | 1-3m | Excel |
| BscScan/Etherscan | ⏳ Not yet implemented | — | — |
| Rabby Protocol (API only) | `pytest ...::test_rabby_export -v -s` | 3m | Excel |
| **2C. Comparison** | | | |
| Overview + EVM/TRX | `python3 run_overview.py -p P_Bd03` | 5-10m | Excel + Screenshots |
| Transaction + TronGrid | `python3 utils/trx_transaction/trongrid_dam_comparison.py <ADDR> <DATE>` | 5-15m | 1 Comparison Excel |
| DAM Transactions only | `python3 -m utils.trx_transaction.dam_extractor <NAME> <DATE>` | 3-10m | Excel |
| Transaction + BscScan | ⏳ Not yet implemented | — | — |
| Full Portfolio — TRX | `python3 run_all.py --trx T... --date DDMMYYYY` | 10-25m | Overview + Transaction Excel |
| Full Portfolio — EVM | `python3 run_all.py --evm 0x...` | 5-10m | Overview Excel (transaction ⏳) |
| Static Comparison | `python3 utils/trx_transaction/static_comparison_generator.py` | ~5s | 3 Excel files |
| **3-5. Operations** | | | |
| Daily Routine | `pytest tests/ -v` | 15-30m | Full Report |
| Debug | `pytest tests/extraction/ -v --headed --slowmo=1000` | Variable | Debug Info |
| All Tests | `pytest tests/ -v` | 30-60m | Full Report |

---

# TIPS & BEST PRACTICES

1. **Always activate virtual environment first:** `source venv/bin/activate`
2. **Use `--headed` to see browser:** `pytest tests/extraction/ -v --headed`
3. **Use `--slowmo` for debugging:** `pytest tests/extraction/ -v --headed --slowmo=500`
4. **Use tracing for complex debugging:** `pytest tests/extraction/ -v --tracing=retain-on-failure`
5. **Check test-results folder:** `ls -la test-results/`
6. **Run specific test:** `pytest tests/extraction/test_portfolio.py::test_create_portfolio -v --headed`
7. **Run by marker:** `pytest tests/ -v -m smoke`

---

# REPORT OUTPUT

```
test-results/
├── reports/
│   └── report_{YYYYMMDD_HHMM}.html
├── excel-exports/
│   └── DAM_Full_{portfolio_name}_{timestamp}.xlsx
├── API Result/
│   ├── API_TRXBalance_{address}_{timestamp}.xlsx
│   └── SimDune_{portfolio}_{timestamp}.xlsx
├── screenshots/
│   └── DAMSS_{portfolio}_{email}_{MMDD_HHMM}/
├── Comparison_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx
├── Step3_TronGrid_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx  (standalone only)
├── Step7_DAM_{addr}_{date_from}_to_{date_to}_{timestamp}.xlsx       (standalone only)
└── trace_*.zip
```

**Note:** When running the full comparison (`trongrid_dam_comparison.py`), only the `Comparison_*.xlsx` file is produced — it contains all TronGrid + DAM data plus the comparison. The `Step3_*` and `Step7_*` files are only produced when running `trongrid_fetcher.py` or `dam_extractor.py` individually.