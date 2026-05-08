# Getting Started - DAM Automation v2

## Prerequisites
- Python 3.8 or higher
- No browser installation needed (Playwright manages browsers)

## Installation

### 1. Navigate to the project directory
```bash
cd core/projects/DAM/automationv2
```

### 2. Create and activate virtual environment
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Install Playwright browsers
```bash
playwright install chromium
```

### 5. Set up environment variables
```bash
cp .env.example .env
# Edit .env with your test credentials
```

## Project Structure

```
automationv2/
├── config/
│   ├── config.py               # Base configuration
│   └── environments/           # Environment-specific configs
├── utils/
│   ├── api/                    # API integrations (Rabby, SimDune, TRX)
│   ├── extraction/             # Data extraction utilities
│   ├── portfolio/              # Portfolio management
│   ├── trx_transaction/        # TRX transaction history & TronGrid comparison (from automationv2TRX)
│   ├── test_data/              # Test data generation
│   ├── security/               # Captcha handling
│   └── debug/                  # Debug utilities
├── tests/
│   ├── ui/                     # UI/Integration tests
│   ├── extraction/             # Data extraction tests
│   ├── api/                    # API tests
│   └── fixtures/               # Shared test fixtures
├── docs/                       # Documentation
├── run_overview.py             # Overview extraction & API comparison
├── run_trx_trans.py            # TRX transaction history comparison
├── run_eth_trans.py            # ETH transaction history (placeholder)
├── run_all.py                  # Master: overview + transactions combined
├── requirements.txt            # Python dependencies
├── pytest.ini                  # Pytest configuration
└── conftest.py                 # Root-level fixtures
```

## Quick Start

### Run Full Portfolio Check (Overview + Transactions)
```bash
cd core/projects/DAM/automationv2
source venv/bin/activate
python3 run_all.py --trx TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6
```

### Run Overview Extraction Only
```bash
python3 run_overview.py --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab
```

### Run TRX Transaction History Only
```bash
python3 run_trx_trans.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6 16042026
```

### Run UI Tests
```bash
pytest tests/ui/ -v --headed
```

### Run Extraction Tests
```bash
pytest tests/extraction/ -v
```

### Run All Tests
```bash
pytest tests/ -v
```

### Run TronGrid vs DAM Transaction Comparison
```bash
python3 utils/trx_transaction/trongrid_dam_comparison.py TQbqqt5kEfgXoQP31HUFumM5bYpieFcNQ6 16042026
```

## Environment Variables

Create `.env` file with:
```env
# Application URL
BASE_URL=https://dam-sit.mqbc21.com

# Test Credentials
TEST_EMAIL=your_email@example.com
TEST_PASSWORD=your_password

# Browser Settings
BROWSER=chromium
HEADLESS=false
SLOW_MO=250

# Timeouts (milliseconds)
DEFAULT_TIMEOUT=30000
NAVIGATION_TIMEOUT=30000
```

## Next Steps

- Read [01_COMMAND_REFERENCE.md](01_COMMAND_REFERENCE.md) for all available commands
- Read [02_TEST_WORKFLOWS.md](02_TEST_WORKFLOWS.md) for different testing scenarios
- Read [05_TROUBLESHOOTING.md](05_TROUBLESHOOTING.md) for common issues
