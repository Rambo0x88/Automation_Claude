# test_full_extraction.py Separation Plan

## Current State
- **File Size**: 13,191 lines
- **Main Functions**: 30+ functions
- **Structure**: Monolithic - all code in one file

## Proposed Separation Strategy

### Module 1: `utils/data_helpers.py` (Utility Functions)
**Purpose**: Data cleaning and transformation helpers
**Functions**:
- `clean_currency_symbols()` - Remove $ and % symbols
- `is_valid_evm_address()` - Validate EVM address format

**Lines**: ~50 lines
**Dependencies**: None

---

### Module 2: `utils/api_handlers.py` (API Integration)
**Purpose**: External API calls and data fetching
**Functions**:
- `fetch_token_details()` - TronScan API
- `fetch_sim_dune_balance()` - Sim Dune API
- `fetch_coingecko_prices_batch()` - CoinGecko prices
- `fetch_coingecko_price_change_batch()` - CoinGecko 24h changes
- `fetch_rabby_protocol()` - Rabby protocol API
- `fetch_rabby_app()` - Rabby app API
- `load_coingecko_coin_list()` - Load coin ID mappings

**Lines**: ~800 lines
**Dependencies**: requests, pandas, openpyxl

---

### Module 3: `utils/address_detection.py` (Address Detection)
**Purpose**: Detect and classify blockchain addresses
**Functions**:
- `detect_evm_addresses()` - Find EVM addresses (0x...)
- `detect_tron_addresses()` - Find Tron addresses (T...)
- `classify_address()` - Classify address type
- `lookup_address_in_excel()` - Look up address in DAM addresses.xlsx
- `lookup_portfolio_in_excel()` - Look up portfolio in DAM addresses.xlsx
- `lookup_portfolio_in_dam()` - Look up portfolio in DAM UI

**Lines**: ~600 lines
**Dependencies**: openpyxl, playwright

---

### Module 4: `utils/excel_exporters.py` (Excel Export Functions)
**Purpose**: Export API data to Excel files
**Functions**:
- `export_sim_dune_to_excel()` - Single address export
- `export_sim_dune_to_excel_combined()` - Multiple addresses export
- `export_rabby_to_excel()` - Single address export
- `export_rabby_to_excel_combined()` - Multiple addresses export
- `export_rabby_app_to_excel_combined()` - Hyperliquid export

**Lines**: ~1,200 lines
**Dependencies**: openpyxl, pandas

---

### Module 5: `utils/validation_helpers.py` (Validation & Calculations)
**Purpose**: Add validation columns and calculate percentages
**Functions**:
- `add_validation_columns_to_overview_token()` - Token validation
- `add_validation_columns_to_token_allocation()` - Token allocation validation
- `add_validation_columns_to_chain_allocation()` - Chain allocation validation
- `add_validation_columns_to_platform_allocation()` - Platform allocation validation
- `add_validation_columns_to_header_holdings()` - Header validation
- `add_validation_columns_to_combined_net_worth()` - Combined net worth validation
- `add_validation_to_defi_tab()` - DeFi validation
- `calculate_allocation_percentage_validation()` - Percentage calculations
- `extract_svg_networth_map()` - SVG parsing for net worth

**Lines**: ~1,000 lines
**Dependencies**: openpyxl, pandas, re, xml.etree

---

### Module 6: `tests/test_trx_balance.py` (TRX Balance Test)
**Purpose**: Part 1 - TRX Balance API Testing
**Functions**:
- `run_trx_balance_api_test()` - Main TRX test function

**Lines**: ~550 lines
**Dependencies**: All utils modules, playwright

---

### Module 7: `tests/test_dam_extraction.py` (DAM Portfolio Extraction)
**Purpose**: Part 2 - DAM Portfolio Full Extraction
**Functions**:
- `run_dam_portfolio_extraction()` - Main DAM extraction function

**Lines**: ~9,300 lines
**Dependencies**: All utils modules, playwright

---

### Module 8: `test_full_extraction.py` (Main Entry Point)
**Purpose**: Main script and CLI interface
**Functions**:
- `main()` - Main execution
- `print_usage()` - Usage instructions

**Lines**: ~100 lines
**Dependencies**: All test modules

---

## File Structure After Separation

```
automationv2/
├── test_full_extraction.py          (Main entry point - 100 lines)
├── utils/
│   ├── __init__.py
│   ├── data_helpers.py              (50 lines)
│   ├── api_handlers.py              (800 lines)
│   ├── address_detection.py         (600 lines)
│   ├── excel_exporters.py           (1,200 lines)
│   └── validation_helpers.py        (1,000 lines)
└── tests/
    ├── __init__.py
    ├── test_trx_balance.py          (550 lines)
    └── test_dam_extraction.py       (9,300 lines)
```

## Benefits

✅ **Modularity**: Each file has single responsibility
✅ **Maintainability**: Easier to find and fix bugs
✅ **Testability**: Can test individual modules
✅ **Reusability**: Import functions from other scripts
✅ **Readability**: Smaller files are easier to understand
✅ **Scalability**: Easy to add new features

## Implementation Steps

1. Create `utils/` directory structure
2. Extract utility functions to respective modules
3. Extract test functions to `tests/` directory
4. Update imports in main file
5. Test that all functionality still works
6. Delete separate docs (01_COMMAND_REFERENCE.md, 02_TEST_WORKFLOWS.md)
7. Keep combined 01_COMMANDS_AND_WORKFLOWS.md
8. Create test case summary document

## Estimated Impact

- **Total lines reduced in main file**: 13,191 → ~100 (99% reduction)
- **Code organization**: Monolithic → Modular
- **Import complexity**: Low (all in same project)
- **Backward compatibility**: 100% (same CLI interface)
