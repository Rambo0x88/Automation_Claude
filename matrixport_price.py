"""
Matrixport RFQ Price Fetcher — Direct HTTP (no ccxt)

1. Fetches all available trading pairs from /trader/v2/api/symbols-info
2. Filters for /USDT pairs
3. Fetches RFQ buy price for a hardcoded list of symbols

Usage:
    python matrixport_price.py              # Fetch prices for hardcoded PAIRS
    python matrixport_price.py --list-pairs # List all available USDT pairs
"""

import hashlib
import hmac
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

# ─── CREDENTIALS (hardcoded) ───
API_KEY = "ak-5fc6e4f8-6509-45dd-ada5-858a539b9d58"
SECRET = "f4Mnb68zUVFSJWm8tXO62m5cjjbbApFjQQkZ1tudu4qg4PFyYFlTrElNPK7vnksz"

BASE_URL = "https://mapi.matrixport.com"

# ─────────────────────────────────────────────────────────────────────────────
# HARDCODED PAIRS — Edit this list to add/remove pairs you want to check.
# Run with --list-pairs to see all available USDT pairs on Matrixport.
# ─────────────────────────────────────────────────────────────────────────────
PAIRS = [
    "1INCHUSDT",
    "AAVEUSDT",
    "APEUSDT",
    "ARBUSDT",
    "AVAXUSDT",
    "AXSUSDT",
    "BATUSDT",
    "BCHUSDT",
    "BNBUSDT",
    "BONKUSDT",
    "BTCUSDT",
    "CAKEUSDT",
    "CHZUSDT",
    "COMPUSDT",
    "CRVUSDT",
    "DCRUSDT",
    "DOGEUSDT",
    "DOTUSDT",
    "EIGENUSDT",
    "ENSUSDT",
    "ETCUSDT",
    "ETHUSDT",
    "FETUSDT",
    "FLOKIUSDT",
    "FTTUSDT",
    "GALAUSDT",
    "GMXUSDT",
    "GRTUSDT",
    "LINKUSDT",
    "LTCUSDT",
    "NEARUSDT",
    "OPUSDT",
    "ORDIUSDT",
    "PEOPLEUSDT",
    "PEPEUSDT",
    "POLUSDT",
    "PYTHUSDT",
    "RENDERUSDT",
    "SANDUSDT",
    "SHIBUSDT",
    "SOLUSDT",
    "SUSHIUSDT",
    "TONUSDT",
    "TRXUSDT",
    "UNIUSDT",
    "USDCUSDT",
    "WLDUSDT",
    "XAGMUSDT",
    "XAUMUSDT",
    "XECUSDT",
    "XLMUSDT",
    "XRPUSDT",
    "ZECUSDT",
]

# Minimum quantities per symbol for RFQ (adjust if you get "incorrect amount" errors)
MIN_QTY = {
    "1INCHUSDT": "30",
    "AAVEUSDT":  "0.2",
    "APEUSDT":   "2",
    "ARBUSDT":   "10",
    "AVAXUSDT":  "0.5",
    "AXSUSDT":   "4.5",
    "BATUSDT":   "4",
    "BCHUSDT":   "0.011",
    "BNBUSDT":   "0.1",
    "BONKUSDT":  "870000",
    "BTCUSDT":   "0.0001",
    "CAKEUSDT":  "4",
    "CHZUSDT":   "100",
    "COMPUSDT":  "0.2",
    "CRVUSDT":   "23",
    "DCRUSDT":   "0.5",
    "DOGEUSDT":  "200",
    "DOTUSDT":   "4.1",
    "EIGENUSDT": "31",
    "ENSUSDT":   "2",
    "ETCUSDT":   "0.63",
    "ETHUSDT":   "0.01",
    "FETUSDT":   "10",
    "FLOKIUSDT": "108000",
    "FTTUSDT":   "4",
    "GALAUSDT":  "1700",
    "GMXUSDT":   "0.4",
    "GRTUSDT":   "60",
    "LINKUSDT":  "10",
    "LTCUSDT":   "0.095",
    "NEARUSDT":  "5",
    "OPUSDT":    "10",
    "ORDIUSDT":  "0.5",
    "PEOPLEUSDT":"790",
    "PEPEUSDT":  "2000000",
    "POLUSDT":   "40",
    "PYTHUSDT":  "130",
    "RENDERUSDT":"3",
    "SANDUSDT":  "66",
    "SHIBUSDT":  "850000",
    "SOLUSDT":   "0.1",
    "SUSHIUSDT": "26",
    "TONUSDT":   "5",
    "TRXUSDT":   "300",
    "UNIUSDT":   "2",
    "USDCUSDT":  "20",
    "WLDUSDT":   "10",
    "XAGMUSDT":  "0.2",
    "XAUMUSDT":  "0.005",
    "XECUSDT":   "395365",
    "XLMUSDT":   "200",
    "XRPUSDT":   "20",
    "ZECUSDT":   "0.2",
}
DEFAULT_QTY = "1"


def sign_v2(method, path, query_string):
    """Generate Matrixport V2 auth headers."""
    ts = str(int(time.time() * 1000))
    prehash = ts + method + path + "&" + query_string
    signature = hmac.new(SECRET.encode(), prehash.encode(), hashlib.sha256).hexdigest()
    return {
        "X-MatrixPort-Access-Key": API_KEY,
        "X-Signature": signature,
        "X-Timestamp": ts,
        "X-Auth-Version": "v2",
    }


def fetch_symbols():
    """Fetch all available trading symbols from Matrixport."""
    path = "/trader/v2/api/symbols-info"
    headers = sign_v2("GET", path, "")
    resp = requests.get(BASE_URL + path, headers=headers, timeout=10)
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"symbols-info error: code={data.get('code')} msg={data.get('message')}")
    return data.get("data", [])


def fetch_rfq_price(symbol, side, qty):
    """Fetch RFQ price for a symbol. side: '1'=buy."""
    path = "/trader/v2/api/rfq-price"
    params = sorted([("qty", qty), ("side", side), ("symbol", symbol)])
    query_string = "&".join(f"{k}={v}" for k, v in params)

    headers = sign_v2("GET", path, query_string)
    resp = requests.get(BASE_URL + path + "?" + query_string, headers=headers, timeout=10)
    resp.raise_for_status()
    return resp.json()


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Matrixport RFQ Price Fetcher")
    parser.add_argument("--list-pairs", action="store_true", help="List all available USDT pairs and exit")
    args = parser.parse_args()

    print("=" * 80)
    print("  Matrixport RFQ Price Fetcher (Python, direct HTTP)")
    print("=" * 80)
    print(f"  API Key: {API_KEY[:12]}...")
    print()

    # Fetch all available symbols
    print("  Loading markets (symbols-info)...")
    symbols_info = fetch_symbols()
    usdt_pairs = [s for s in symbols_info if s.get("pair", "").endswith("USDT")]
    all_usdt_symbols = sorted([s["pair"] for s in usdt_pairs])

    print(f"  Found {len(usdt_pairs)} USDT pairs out of {len(symbols_info)} total")
    print()

    if args.list_pairs:
        print("  Available USDT pairs on Matrixport:")
        print("  " + "-" * 50)
        for sym_info in sorted(usdt_pairs, key=lambda x: x["pair"]):
            sym = sym_info["pair"]
            min_qty = sym_info.get("base_min_amount", "?")
            max_qty = sym_info.get("base_max_amount", "?")
            print(f"    {sym:<16} min_qty: {str(min_qty):<12} max_qty: {max_qty}")
        print()
        print("  Copy the symbols you want into the PAIRS list at the top of this script.")
        return

    # Load previous prices for % change comparison
    out_file = "matrixport_prices.json"
    prev_prices = {}
    if os.path.exists(out_file):
        with open(out_file) as f:
            prev_data = json.load(f)
        for row in prev_data:
            if "price" in row:
                prev_prices[row["symbol"]] = float(row["price"])

    # Fetch prices for hardcoded pairs
    print(f"  Fetching prices for {len(PAIRS)} pairs...")
    print()
    has_prev = bool(prev_prices)
    if has_prev:
        print(f"  {'Symbol':<14} {'Buy Price':>18} {'Change %':>10} {'Qty':>10} {'Status'}")
    else:
        print(f"  {'Symbol':<14} {'Buy Price':>18} {'Qty':>10} {'Status'}")
    print("  " + "-" * 70)

    results = []

    def get_price(symbol):
        qty = MIN_QTY.get(symbol, DEFAULT_QTY)
        try:
            resp = fetch_rfq_price(symbol, "1", qty)
            if resp.get("code") == 0:
                return {
                    "symbol":   symbol,
                    "price":    resp["data"]["price"],
                    "qty":      qty,
                    "cash":     resp["data"].get("cash", ""),
                    "price_id": resp["data"].get("price_id", ""),
                    "raw":      resp,
                }
            else:
                return {
                    "symbol": symbol,
                    "error":  f"code={resp.get('code')} msg={resp.get('message')}",
                    "qty":    qty,
                    "raw":    resp,
                }
        except Exception as e:
            return {"symbol": symbol, "error": str(e), "qty": qty}

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(get_price, sym): sym for sym in PAIRS}
        for future in as_completed(futures):
            results.append(future.result())

    # Sort results by original PAIRS order
    results.sort(key=lambda r: PAIRS.index(r["symbol"]))

    for r in results:
        if "error" in r:
            print(f"  {r['symbol']:<14} {'ERROR':>18} {'':>10} {r['qty']:>10} [{r['error']}]")
        else:
            cur = float(r["price"])
            if has_prev and r["symbol"] in prev_prices:
                prev = prev_prices[r["symbol"]]
                pct = (cur - prev) / prev * 100
                arrow = "▲" if pct >= 0 else "▼"
                pct_str = f"{arrow}{abs(pct):.4f}%"
            else:
                pct_str = "N/A"
            if has_prev:
                print(f"  {r['symbol']:<14} {r['price']:>18} {pct_str:>10} {r['qty']:>10} OK")
            else:
                print(f"  {r['symbol']:<14} {r['price']:>18} {r['qty']:>10} OK")

    # Export to JSON
    print()
    out_file = "matrixport_prices.json"
    with open(out_file, "w") as f:
        json.dump(results, f, indent=2)
    print(f"  Raw responses -> {out_file}")

    # Warn about pairs not available on the exchange
    available_set = set(all_usdt_symbols)
    missing = [p for p in PAIRS if p not in available_set]
    if missing:
        print()
        print(f"  [WARN] These pairs in PAIRS list are NOT available on Matrixport:")
        for m in missing:
            print(f"    - {m}")

    print()
    print("Done.")


if __name__ == "__main__":
    main()
