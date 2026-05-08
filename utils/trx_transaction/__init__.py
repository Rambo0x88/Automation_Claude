"""
TRX Transaction Utilities
=========================

Modular pipeline for TRX (TRON) transaction extraction and comparison.

Modules:
--------
- shared.py                  : Config, helpers, date parsing, Excel styling (shared state)
- trongrid_fetcher.py        : Step 3 — TronGrid API fetch + parse + standalone Excel
- dam_extractor.py           : Steps 4-7 — DAM UI Playwright extraction + standalone Excel
- trongrid_dam_comparison.py : Orchestrator — runs both + builds single comparison Excel (Step 8)

Standalone usage (from automationv2/ root):
-------------------------------------------
    # Step 3 only — TronGrid data
    python3 -m utils.trx_transaction.trongrid_fetcher <TRX_ADDRESS> <DDMMYYYY>

    # Step 7 only — DAM UI data
    python3 -m utils.trx_transaction.dam_extractor <PORTFOLIO_NAME> <DDMMYYYY>

    # Full comparison (Step 8) — single Excel with all data
    python3 utils/trx_transaction/trongrid_dam_comparison.py <ADDRESS_OR_NAME> <DDMMYYYY>

Legacy scripts (still available):
---------------------------------
- dam_transaction_extractor.py        : Standalone DAM extraction (own sign-in)
- dam_transaction_extractor_v2.py     : Pure extraction functions (caller provides page)
- dam_transaction_extractor_v1.py     : Legacy v1
- dam_transaction_extractor_jan21.py  : Legacy Jan 21 extraction
- static_comparison_generator.py      : Static comparison (no API calls)
"""
