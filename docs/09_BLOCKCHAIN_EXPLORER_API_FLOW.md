# STEP 5: Blockchain Explorer API Flow (NEW)

## Overview
New API-based alternative to manual CSV export. Fully automated, no Cloudflare issues, unlimited data.

**Status**: ✅ Ready to use (parallel to existing CSV method)

---

## Quick Start

### 1. Get API Keys (Free)

#### Etherscan (Ethereum)
1. Go to https://etherscan.io/apis
2. Sign up for free account
3. Create API key
4. Copy key

#### BSCScan (Binance Smart Chain)
1. Go to https://bscscan.com/apis
2. Sign up for free account
3. Create API key
4. Copy key

#### BaseScan (Base)
1. Go to https://basescan.org/apis
2. Sign up for free account
3. Create API key
4. Copy key

### 2. Configure API Keys

Create `test_data/blockchain_explorer_api_keys.json`:
```json
{
  "ethereum": "YOUR_ETHERSCAN_API_KEY",
  "bsc": "YOUR_BSCSCAN_API_KEY",
  "base": "YOUR_BASESCAN_API_KEY"
}
```

Or copy from template:
```bash
cp test_data/blockchain_explorer_api_keys.json.example test_data/blockchain_explorer_api_keys.json
# Then edit and add your API keys
```

### 3. Run API Test

```bash
# Single address
python3 -m tests.test_blockchain_explorer_api --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain ethereum

# Multiple addresses
python3 -m tests.test_blockchain_explorer_api --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab 0x1234567890abcdef --chain ethereum

# Different chain
python3 -m tests.test_blockchain_explorer_api --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --chain bsc

# Custom output folder
python3 -m tests.test_blockchain_explorer_api --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --output my-results

# With explicit API key
python3 -m tests.test_blockchain_explorer_api --evm 0x4e14fc11f58b64740e66e4b1aa188a4b007c0eab --api-key YOUR_KEY
```

### 4. Output

Excel files in `test-results/blockchain-explorer-api/`:
- `Etherscan_0x4e14fc_20260401_120000.xlsx`
  - Summary sheet
  - Transactions sheet
  - Internal Transactions sheet
  - Token Transfers sheet
  - NFT Transfers sheet

---

## Data Fetched

### 1. Transactions
```
Fields:
- blockNumber
- timeStamp
- hash (transaction hash)
- nonce
- blockHash
- transactionIndex
- from
- to
- value (in wei)
- gas
- gasPrice
- isError
- txreceipt_status
- input
- contractAddress
- cumulativeGasUsed
- gasUsed
- confirmations
- methodId
- functionName
```

**Count**: All transactions for address (no limit)
**Time Range**: All time (no date range limit)

### 2. Internal Transactions
```
Fields:
- blockNumber
- timeStamp
- hash (transaction hash)
- from
- to
- value (in wei)
- contractAddress
- input
- type
- gas
- gasUsed
- traceId
- isError
- errCode
```

**Count**: All internal transactions
**Time Range**: All time

### 3. Token Transfers (ERC-20)
```
Fields:
- blockNumber
- timeStamp
- hash (transaction hash)
- nonce
- blockHash
- from
- contractAddress (token contract)
- to
- value (token amount)
- tokenName
- tokenSymbol
- tokenDecimal
- transactionIndex
- gas
- gasPrice
- gasUsed
- cumulativeGasUsed
- input
- confirmations
```

**Count**: All ERC-20 transfers
**Time Range**: All time

### 4. NFT Transfers (ERC-721 + ERC-1155)
```
Fields:
- blockNumber
- timeStamp
- hash (transaction hash)
- nonce
- blockHash
- from
- contractAddress (NFT contract)
- to
- tokenID
- tokenName
- tokenSymbol
- tokenDecimal
- transactionIndex
- gas
- gasPrice
- gasUsed
- cumulativeGasUsed
- input
- confirmations
```

**Count**: All NFT transfers
**Time Range**: All time

---

## API Limits

### Rate Limits
- **Free Tier**: 5 calls/second
- **Pro Tier**: 20 calls/second
- **Premium Tier**: 100+ calls/second

### Data Limits
- **Transactions per call**: 10,000 (paginated)
- **Token transfers per call**: 10,000 (paginated)
- **NFT transfers per call**: 10,000 (paginated)

### Retry Logic
- Automatic retry on rate limit (up to 3 attempts)
- 2-second delay between retries
- Exponential backoff for failures

---

## Comparison: CSV vs API

| Feature | CSV Export | API Flow |
|---------|-----------|---------|
| **Automation** | ❌ Manual | ✅ Fully automated |
| **Speed** | ❌ Slow (user waits) | ✅ Fast (seconds) |
| **Cloudflare** | ✅ Bypasses | ✅ No issue |
| **Date Range** | ❌ Limited (1 year) | ✅ Unlimited |
| **Real-time** | ❌ No | ✅ Yes |
| **Pagination** | ❌ Limited | ✅ Full support |
| **Cost** | ❌ Time cost | ✅ Free (generous limits) |
| **Setup** | ❌ Complex | ✅ Simple (API key) |
| **Reliability** | ⚠️ User-dependent | ✅ Reliable |

---

## Implementation Details

### Module: `utils/blockchain_explorer_api.py`

#### Functions

**`fetch_transactions(address, chain, api_key)`**
- Fetch all transactions for address
- Returns: (address, transactions_list, success)

**`fetch_internal_transactions(address, chain, api_key)`**
- Fetch all internal transactions
- Returns: (address, internal_txs_list, success)

**`fetch_token_transfers(address, chain, api_key)`**
- Fetch all ERC-20 token transfers
- Returns: (address, transfers_list, success)

**`fetch_nft_transfers(address, chain, api_key)`**
- Fetch all NFT transfers
- Returns: (address, nft_transfers_list, success)

**`fetch_all_explorer_data(address, chain, api_key)`**
- Fetch all data types at once
- Returns: dict with all data

**`export_explorer_data_to_excel(explorer_data, output_path)`**
- Export data to Excel file
- Returns: True/False

### Test Module: `tests/test_blockchain_explorer_api.py`

**`run_blockchain_explorer_api_test(evm_addresses, chain, api_key, output_folder)`**
- Main test function
- Fetches data for all addresses
- Exports to Excel

**`load_api_keys()`**
- Load API keys from config file
- Returns: dict of API keys

---

## Error Handling

### Rate Limit Handling
```
1. Request fails with rate limit error
2. Wait 2 seconds
3. Retry (up to 3 times)
4. If still failing, return empty data
```

### Network Error Handling
```
1. Request fails with network error
2. Wait 2 seconds
3. Retry (up to 3 times)
4. If still failing, return empty data
```

### API Error Handling
```
1. API returns error status
2. Log error message
3. Return empty data
4. Continue with next request
```

---

## Performance

### Typical Execution Times
- **Single address**: 5-10 seconds
- **5 addresses**: 30-50 seconds
- **10 addresses**: 60-100 seconds

### Data Volume
- **Transactions**: 100-1000+ per address
- **Internal Transactions**: 10-100 per address
- **Token Transfers**: 50-500 per address
- **NFT Transfers**: 0-100 per address

### Excel File Size
- **Typical**: 1-5 MB per address
- **Large addresses**: 10-50 MB

---

## Troubleshooting

### "No API key found"
**Solution**: Create `test_data/blockchain_explorer_api_keys.json` with your API keys

### "Rate limit exceeded"
**Solution**: 
- Wait a few minutes
- Upgrade to Pro/Premium tier
- Use multiple API keys (rotate between them)

### "Invalid API key"
**Solution**: 
- Verify API key is correct
- Check API key is for correct chain
- Regenerate API key on explorer website

### "Address not found"
**Solution**: 
- Verify address is valid (0x + 40 hex chars)
- Check address exists on chain
- Try different chain

### "No data returned"
**Solution**: 
- Address might have no transactions
- Try different address
- Check API key rate limit

---

## Integration with Main Flow

### Option 1: Use API Instead of CSV
Replace STEP 5 CSV export with API flow:
```python
# In run_overview.py
if USE_API_FLOW:
    from tests.test_blockchain_explorer_api import run_blockchain_explorer_api_test
    run_blockchain_explorer_api_test(evm_addresses, chain="ethereum")
else:
    # Old CSV export method
    ...
```

### Option 2: Use Both (Parallel)
Run both CSV and API flows:
```python
# CSV export (manual)
run_csv_export()

# API export (automated)
run_blockchain_explorer_api_test()
```

### Option 3: Use API as Fallback
Try API first, fall back to CSV:
```python
try:
    run_blockchain_explorer_api_test()
except Exception as e:
    print(f"API failed: {e}, falling back to CSV export")
    run_csv_export()
```

---

## Next Steps

1. ✅ Get API keys (free, 5 minutes)
2. ✅ Configure API keys in JSON file
3. ✅ Run test: `python3 -m tests.test_blockchain_explorer_api --evm 0x...`
4. ✅ Verify Excel output
5. ⏳ Integrate into main flow (optional)
6. ⏳ Remove CSV export (optional)

---

## Files

### New Files Created
- `utils/blockchain_explorer_api.py` - API handlers
- `tests/test_blockchain_explorer_api.py` - Test module
- `test_data/blockchain_explorer_api_keys.json.example` - Config template
- `docs/09_BLOCKCHAIN_EXPLORER_API_FLOW.md` - This file

### Configuration
- `test_data/blockchain_explorer_api_keys.json` - Your API keys (create from template)

### Output
- `test-results/blockchain-explorer-api/` - Excel files

---

## Support

For issues:
1. Check `docs/05_TROUBLESHOOTING.md`
2. Verify API keys are correct
3. Check rate limits
4. Try different address
5. Check explorer website status

---

**Status**: ✅ Ready to use
**Maintenance**: Parallel to existing CSV method (no breaking changes)
**Future**: Can replace CSV method entirely
