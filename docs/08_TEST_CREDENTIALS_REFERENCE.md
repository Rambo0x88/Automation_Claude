# Test Credentials Reference

## Where Test Credentials Are Stored

### 1. **DAM Application Accounts**

#### TC0 Account (Portfolio Creation Test)
**File**: `test_data/tc0_account.json`
```json
{
  "email": "moontest2407@gmail.com",
  "password": "Orion222!!!!",
  "note": "Account for TC0 - should have NO existing portfolios"
}
```

#### TC1 Account (Existing User Test)
**File**: `test_data/tc1_account.json`
```json
{
  "email": "moontest1311@gmail.com",
  "password": "Orion888!!!!",
  "note": "Account created by TC2 for TC1 to reuse"
}
```

### 2. **Blockchain Explorer Accounts** ✅ MOVED TO test_data

**File**: `test_data/blockchain_explorers.json`
```json
{
  "explorers": [
    {
      "name": "Etherscan",
      "base_url": "https://etherscan.io",
      "username": "guatfern",
      "password": "Orion888!!!!"
    },
    {
      "name": "BSCScan",
      "base_url": "https://bscscan.com",
      "username": "guatfern",
      "password": "Orion888!!!!"
    },
    {
      "name": "BaseScan",
      "base_url": "https://basescan.org",
      "username": "moontest1803",
      "password": "Orion888!!!!"
    }
  ]
}
```

**Code Reference**: `run_overview.py` (loads from test_data file)

### 3. **Wallet Addresses**

**File**: `test_data/tc_dune_wallet.json`
- Contains wallet addresses for extraction tests
- Used by Rabby Protocol and SimDune API tests

**File**: `test_data/DAM addresses.xlsx`
- Excel file with portfolio names and wallet addresses
- Used for portfolio lookup and creation

**File**: `test_data/daily_test_addresses.xlsx`
- Daily testing wallet addresses

## Summary Table

| Account | Type | Email/Username | Password | File Location |
|---------|------|---|---|---|
| TC0 | DAM App | moontest2407@gmail.com | Orion222!!!! | `test_data/tc0_account.json` |
| TC1 | DAM App | moontest1311@gmail.com | Orion888!!!! | `test_data/tc1_account.json` |
| BaseScan | Blockchain Explorer | moontest1803 | Orion888!!!! | `test_data/blockchain_explorers.json` ✅ |
| Etherscan | Blockchain Explorer | guatfern | Orion888!!!! | `test_data/blockchain_explorers.json` ✅ |
| BSCScan | Blockchain Explorer | guatfern | Orion888!!!! | `test_data/blockchain_explorers.json` ✅ |

## How Credentials Are Used

### DAM Application
1. **TC0**: Used for portfolio creation tests (fresh account, no portfolios)
2. **TC1**: Used for existing user tests (reuses account created by TC2)

### Blockchain Explorers
- **BaseScan (moontest1803)**: Used for scanning Base blockchain data
- **Etherscan (guatfern)**: Used for scanning Ethereum blockchain data
- **BSCScan (guatfern)**: Used for scanning BSC blockchain data

### Wallet Addresses
- Stored in Excel files for portfolio creation
- Used for API calls (SimDune, Rabby, TRX Balance)
- Used for DAM portfolio extraction

## Configuration

### Environment Variables (.env)
```env
# DAM Application
TEST_EMAIL=moontest1311@gmail.com
TEST_PASSWORD=Orion888!!!!
```

### Test Data Files
```
test_data/
├── tc0_account.json                 # TC0 DAM account
├── tc1_account.json                 # TC1 DAM account
├── blockchain_explorers.json        # Blockchain explorer accounts ✅
├── tc_dune_wallet.json              # Wallet addresses
├── DAM addresses.xlsx               # Portfolio addresses
└── daily_test_addresses.xlsx        # Daily test addresses
```

### Config File (config/config.py)
```python
TEST_EMAIL = os.getenv('TEST_EMAIL', '')
TEST_PASSWORD = os.getenv('TEST_PASSWORD', 'TestPassword123!')
```

### Code Loading (run_overview.py)
```python
# Load blockchain explorer credentials from test_data
_explorer_configs_file = "test_data/blockchain_explorers.json"
with open(_explorer_configs_file, 'r') as _f:
    _explorer_data = json.load(_f)
    _explorer_configs = _explorer_data.get('explorers', [])
```

## Security Notes

⚠️ **Important**: These are test credentials for QA environment only
- Do NOT use in production
- Do NOT commit to public repositories
- Do NOT share with unauthorized users
- Rotate credentials periodically

## Updating Credentials

### To Update DAM Account Credentials:
1. Edit `test_data/tc0_account.json` or `test_data/tc1_account.json`
2. Update email and password
3. Re-run tests

### To Update Blockchain Explorer Credentials: ✅ NOW IN test_data
1. Edit `test_data/blockchain_explorers.json`
2. Update username and password for the explorer
3. Re-run tests

### To Update Environment Variables:
1. Edit `.env` file
2. Update `TEST_EMAIL` and `TEST_PASSWORD`
3. Restart tests

## Troubleshooting

### "Sign in failed" Error
- Check credentials in `.env` or `tc1_account.json`
- Verify account is not locked
- Try signing in manually in browser

### "Blockchain Explorer authentication failed"
- Check `moontest1803` credentials in `run_overview.py`
- Verify BaseScan account is active
- Check if IP is blocked

### "Portfolio not found"
- Check wallet addresses in `test_data/DAM addresses.xlsx`
- Verify portfolio exists in DAM UI
- Check portfolio name spelling

## Related Files

- `test_data/tc0_account.json` - TC0 account
- `test_data/tc1_account.json` - TC1 account
- `test_data/blockchain_explorers.json` - Blockchain explorer accounts ✅ NEW
- `test_data/tc_dune_wallet.json` - Wallet addresses
- `test_data/DAM addresses.xlsx` - Portfolio addresses
- `run_overview.py` - Loads credentials from test_data files
- `.env` - Environment variables
- `config/config.py` - Configuration management
