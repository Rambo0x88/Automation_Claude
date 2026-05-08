# Troubleshooting - DAM Automation v2

## Common Issues & Solutions

### Installation Issues

#### Issue: `ModuleNotFoundError: No module named 'playwright'`
**Cause**: Dependencies not installed

**Solution**:
```bash
cd core/projects/DAM/automationv2
source venv/bin/activate
pip install -r requirements.txt
```

#### Issue: `playwright: command not found`
**Cause**: Playwright browsers not installed

**Solution**:
```bash
playwright install chromium
```

#### Issue: `error: externally-managed-environment`
**Cause**: Python environment is system-managed (macOS with Homebrew)

**Solution**:
```bash
# Use virtual environment
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

### Runtime Issues

#### Issue: `TimeoutException: Timeout 30000ms exceeded`
**Cause**: Element not found or page not loaded in time

**Solution**:
1. Increase timeout in `.env`:
   ```env
   DEFAULT_TIMEOUT=60000
   NAVIGATION_TIMEOUT=60000
   ```

2. Run with visible browser to see what's happening:
   ```bash
   pytest tests/ui/test_portfolio.py -v --headed
   ```

3. Add slow motion to see actions:
   ```bash
   pytest tests/ui/test_portfolio.py -v --headed --slowmo=1000
   ```

#### Issue: `TargetClosedError: Target page, context or browser has been closed`
**Cause**: Browser crashed or closed unexpectedly

**Solution**:
1. Run in headed mode to see errors:
   ```bash
   pytest tests/ui/test_portfolio.py -v --headed
   ```

2. Reduce slow motion and timeouts
3. Check system resources (memory, CPU)
4. Try with Firefox browser:
   ```bash
   pytest tests/ui/test_portfolio.py -v --browser=firefox
   ```

#### Issue: `ElementNotFound: No element matches the selector`
**Cause**: UI element locator is incorrect or element doesn't exist

**Solution**:
1. Run with Playwright Inspector:
   ```bash
   PWDEBUG=1 pytest tests/ui/test_portfolio.py -v
   ```

2. Check if element exists on page:
   ```bash
   pytest tests/ui/test_portfolio.py -v --headed --slowmo=1000
   ```

3. Update selector in page object or test

#### Issue: `AssertionError: expected True but got False`
**Cause**: Assertion failed during test

**Solution**:
1. Check test output for details
2. Run with tracing:
   ```bash
   pytest tests/ui/test_portfolio.py -v --tracing=retain-on-failure
   playwright show-trace test-results/trace_*.zip
   ```

3. Review test logic and expected values

---

### API Issues

#### Issue: `ConnectionError: Failed to establish connection`
**Cause**: Network issue or API server down

**Solution**:
1. Check internet connection:
   ```bash
   ping google.com
   ```

2. Check API status:
   - TRX: https://api.trongrid.io/
   - SimDune: https://simdune.com/
   - Rabby: https://api.rabby.io/

3. Retry with exponential backoff (automatic in code)

#### Issue: `HTTPError: 429 Too Many Requests`
**Cause**: Rate limit exceeded

**Solution**:
1. Reduce request frequency
2. Implement delays between requests
3. Use batch processing
4. Check API rate limits in `docs/03_API_REFERENCE.md`

#### Issue: `HTTPError: 403 Forbidden`
**Cause**: Bot detection (usually CoinGecko)

**Solution**:
1. Disable CoinGecko in `.env`:
   ```env
   ENABLE_COINGECKO=false
   ```

2. Or use Firefox browser:
   ```bash
   pytest tests/extraction/ -v --browser=firefox
   ```

#### Issue: `Invalid address format`
**Cause**: Address doesn't match expected format

**Solution**:
1. Check address format:
   - EVM: `0x` + 40 hex characters (42 total)
   - TRX: `T` + 33 alphanumeric characters (34 total)

2. Verify address is correct:
   ```bash
   python3 -c "addr='0x...'; print(f'Valid: {len(addr)==42 and addr.startswith(\"0x\")}')"
   ```

---

### Excel Export Issues

#### Issue: `PermissionError: [Errno 13] Permission denied`
**Cause**: Excel file is open or read-only

**Solution**:
1. Close Excel file
2. Check file permissions:
   ```bash
   ls -la test-results/excel-exports/
   ```

3. Remove read-only flag:
   ```bash
   chmod 644 test-results/excel-exports/*.xlsx
   ```

#### Issue: `FileNotFoundError: [Errno 2] No such file or directory`
**Cause**: Output directory doesn't exist

**Solution**:
```bash
mkdir -p test-results/excel-exports
mkdir -p test-results/screenshots
```

#### Issue: Excel file is corrupted
**Cause**: Writing error during export

**Solution**:
1. Check disk space:
   ```bash
   df -h
   ```

2. Check file size:
   ```bash
   ls -lh test-results/excel-exports/
   ```

3. Re-run extraction

---

### Portfolio Issues

#### Issue: `Portfolio not found in DAM`
**Cause**: Portfolio doesn't exist or wrong name

**Solution**:
1. Check portfolio name:
   ```bash
   python3 run_overview.py -p "portfolio_name"
   ```

2. List available portfolios in DAM UI
3. Create portfolio if needed:
   ```bash
   python3 run_overview.py --evm 0x...
   ```

#### Issue: `Address already exists in portfolio`
**Cause**: Trying to add duplicate address

**Solution**:
1. Use different address
2. Or use existing portfolio with that address

#### Issue: `Portfolio creation failed`
**Cause**: UI interaction issue or validation error

**Solution**:
1. Run with visible browser:
   ```bash
   pytest tests/ui/test_portfolio.py -v --headed
   ```

2. Check address format
3. Check portfolio name is unique
4. Try manual creation in DAM UI

---

### Authentication Issues

#### Issue: `Sign in failed`
**Cause**: Invalid credentials or account locked

**Solution**:
1. Verify credentials in `.env`:
   ```bash
   cat .env | grep TEST_EMAIL
   cat .env | grep TEST_PASSWORD
   ```

2. Try signing in manually in browser
3. Check if account is locked
4. Reset password if needed

#### Issue: `Captcha challenge`
**Cause**: Cloudflare Turnstile captcha blocking automation

**Solution**:
1. Use Firefox browser (better fingerprint):
   ```bash
   pytest tests/ui/test_sign_up.py -v --browser=firefox
   ```

2. Or disable captcha in test environment
3. Or use MCP captcha solver (if configured)

---

### Performance Issues

#### Issue: Tests running slowly
**Cause**: Network latency, slow API, or system resources

**Solution**:
1. Check system resources:
   ```bash
   top -l 1 | head -20
   ```

2. Reduce slow motion:
   ```env
   SLOW_MO=0
   ```

3. Use parallel execution:
   ```bash
   pytest tests/ -v -n auto
   ```

4. Check network speed:
   ```bash
   ping api.rabby.io
   ```

#### Issue: Memory usage too high
**Cause**: Large data extraction or memory leak

**Solution**:
1. Reduce batch size
2. Process addresses one at a time
3. Clear cache between runs
4. Monitor memory:
   ```bash
   watch -n 1 'ps aux | grep python'
   ```

---

### Debugging Techniques

#### 1. Visible Browser Execution
```bash
pytest tests/ui/test_portfolio.py -v --headed
```
See exactly what the browser is doing.

#### 2. Slow Motion
```bash
pytest tests/ui/test_portfolio.py -v --headed --slowmo=1000
```
Slow down actions to 1 second each.

#### 3. Playwright Inspector
```bash
PWDEBUG=1 pytest tests/ui/test_portfolio.py -v
```
Interactive debugging with step-through.

#### 4. Tracing
```bash
pytest tests/ui/test_portfolio.py -v --tracing=retain-on-failure
playwright show-trace test-results/trace_*.zip
```
Time-travel debugging with full context.

#### 5. Screenshots
```bash
pytest tests/ui/test_portfolio.py -v --screenshot=on
```
Capture screenshots at each step.

#### 6. Videos
```bash
pytest tests/ui/test_portfolio.py -v --video=on
```
Record video of test execution.

#### 7. Logging
```python
import logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)
logger.debug("Debug message")
```

#### 8. Print Debugging
```python
print(f"Current URL: {page.url}")
print(f"Page content: {page.content()}")
```

---

### Getting Help

#### 1. Check Logs
```bash
# View test output
pytest tests/ui/test_portfolio.py -v -s

# View trace
playwright show-trace test-results/trace_*.zip
```

#### 2. Check Screenshots
```bash
ls -la test-results/screenshots/
open test-results/screenshots/
```

#### 3. Check Videos
```bash
ls -la videos/
open videos/
```

#### 4. Review Test Code
```bash
cat tests/ui/test_portfolio.py
```

#### 5. Check Configuration
```bash
cat .env
cat config/config.py
```

---

### Quick Checklist

- [ ] Virtual environment activated: `source venv/bin/activate`
- [ ] Dependencies installed: `pip install -r requirements.txt`
- [ ] Playwright browsers installed: `playwright install chromium`
- [ ] `.env` file configured with credentials
- [ ] Internet connection working
- [ ] Disk space available
- [ ] No other tests running
- [ ] Browser not open (for headless tests)
- [ ] Correct working directory: `core/projects/DAM/automationv2/`

---

### Still Having Issues?

1. **Check the logs**: `pytest tests/ -v -s`
2. **Run with visible browser**: `pytest tests/ -v --headed`
3. **Use Playwright Inspector**: `PWDEBUG=1 pytest tests/ -v`
4. **Check trace**: `playwright show-trace test-results/trace_*.zip`
5. **Review test code**: `cat tests/ui/test_portfolio.py`
6. **Check configuration**: `cat .env`
7. **Verify API status**: Check API documentation
8. **Check system resources**: `top`, `df -h`

---

### TRX Transaction Comparison Issues

#### Issue: `No transactions found with current filter`
**Cause**: Date filter didn't apply correctly or no transactions exist for that date

**Solution**:
1. Verify the date format is DDMMYYYY (e.g. `16042026` = April 16, 2026)
2. Check if the portfolio has transactions on that date in DAM UI manually
3. Try a wider date range: `python3 run_trx_trans.py <ADDR> 01042026 30042026`

#### Issue: `DAM sign-in FAILED after 2 attempts`
**Cause**: Credentials incorrect or DAM is down

**Solution**:
1. Check `test_data/tc1_account.json` credentials
2. Try signing in manually at https://dam-sit.mqbc21.com/sign-in
3. Check screenshot in `test-results/screenshots/dam_signin_error.png`

#### Issue: `Portfolio not found` in transaction comparison
**Cause**: Portfolio name doesn't match any dropdown item

**Solution**:
1. Use exact portfolio name as shown in DAM dropdown
2. Or use TRX address directly — it will create the portfolio: `python3 run_trx_trans.py TQbqqt5k... 16042026`

#### Issue: Timezone mismatch between TronGrid and DAM
**Cause**: TronGrid uses UTC, DAM uses UTC+7 (ICT)

**This is expected behavior**, not a bug. Transactions near midnight will appear on different calendar dates. The Step8 Comparison Excel documents this in the Summary sheet.

#### Issue: `Something went wrong` error in DAM after applying date filter
**Cause**: DAM backend error for the selected date range

**Solution**:
1. The script auto-retries with a Refresh click
2. If still failing, try a different date range
3. Check if DAM is having issues by browsing manually
