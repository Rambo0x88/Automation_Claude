# DAM Automation v2

Automated testing and data extraction for the DAM (Digital Asset Management) platform.

## What It Does

- **Overview Extraction**: Extracts portfolio data from DAM Overview page, compares against on-chain APIs (TronGrid, SimDune, Rabby, CoinGecko)
- **TRX Transaction Comparison**: Fetches TronGrid transaction history, extracts DAM transaction rows, builds side-by-side comparison
- **ETH Transaction Comparison**: Placeholder — coming soon
- **UI Testing**: Sign-up, sign-in, portfolio creation, account settings

## Quick Start

```bash
cd core/projects/DAM/automationv2
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
cp .env.example .env
```

## Main Entry Points

```bash
# Overview extraction + API comparison
python3 run_overview.py --trx T... --evm 0x...

# TRX transaction history comparison (defaults to today)
python3 run_trx_trans.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6

# Both overview + transactions combined
python3 run_all.py --trx T... --date 16042026

# Run pytest tests
pytest tests/ -v
```

## Project Structure

```
automationv2/
├── run_overview.py              # Overview extraction & API comparison (14k lines)
├── run_trx_trans.py             # TRX transaction history comparison
├── run_eth_trans.py             # ETH transaction history (placeholder)
├── run_all.py                   # Master: overview + transactions combined
├── config/
│   └── config.py                # Configuration management
├── utils/
│   ├── trx_transaction/         # TRX transaction scripts (from automationv2TRX)
│   ├── portfolio/               # Portfolio management
│   ├── dam_data_extractor.py    # DAM UI data extraction
│   ├── chain_data_extractor.py  # Chain data extraction
│   ├── rabby_api.py             # Rabby Protocol API
│   ├── helpers.py               # General helpers
│   └── captcha_solver.py        # Captcha handling
├── tests/
│   ├── extraction/              # Data extraction tests
│   ├── ui/                      # UI tests (not yet populated)
│   └── api/                     # API tests (not yet populated)
├── test_data/                   # Test credentials and addresses
├── docs/                        # Documentation (13 files)
└── pages/                       # Page objects for Playwright
```

## Documentation

See `docs/` folder:
- [00_GETTING_STARTED.md](docs/00_GETTING_STARTED.md) — Setup
- [01_COMMANDS_AND_WORKFLOWS.md](docs/01_COMMANDS_AND_WORKFLOWS.md) — All workflows and commands
- [03_API_REFERENCE.md](docs/03_API_REFERENCE.md) — API documentation
- [04_EXCEL_OUTPUT_FORMAT.md](docs/04_EXCEL_OUTPUT_FORMAT.md) — Excel output structure
- [05_TROUBLESHOOTING.md](docs/05_TROUBLESHOOTING.md) — Common issues
