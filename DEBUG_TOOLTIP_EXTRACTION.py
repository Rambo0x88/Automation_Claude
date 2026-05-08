"""
Debug version of tooltip extraction with detailed logging.
This script adds comprehensive logging to identify root causes of missing tooltips.

Usage:
    python DEBUG_TOOLTIP_EXTRACTION.py
    
Output:
    - Detailed logging for each token
    - Element properties (visible, bounding box, etc.)
    - Tooltip properties (visible, text content, etc.)
    - JavaScript extraction results
    - Summary of failing tokens
"""

import json
from datetime import datetime
from pathlib import Path


def extract_tooltip_with_debug(page, elem, tooltip_id, token_name, row_idx, tooltip_type):
    """
    Extract tooltip with detailed debug logging.
    
    Args:
        page: Playwright page object
        elem: Element to hover over
        tooltip_id: ID of tooltip element
        token_name: Name of token (for logging)
        row_idx: Row index (for logging)
        tooltip_type: Type of tooltip (price, share, amount, price_24h)
    
    Returns:
        dict: {
            'value': extracted_value or None,
            'debug': {
                'element_visible': bool,
                'element_bounding_box': dict,
                'tooltip_visible': bool,
                'tooltip_text': str,
                'extraction_result': str,
                'error': str or None
            }
        }
    """
    
    debug_info = {
        'value': None,
        'debug': {
            'element_visible': False,
            'element_bounding_box': None,
            'tooltip_visible': False,
            'tooltip_text': None,
            'extraction_result': None,
            'error': None
        }
    }
    
    try:
        # ===== STEP 1: Check element visibility =====
        print(f"      [DEBUG] Checking element visibility...")
        
        try:
            is_visible = elem.is_visible()
            debug_info['debug']['element_visible'] = is_visible
            print(f"         Element visible: {is_visible}")
            
            if not is_visible:
                print(f"         ⚠️  Element NOT visible, attempting scroll into view...")
                elem.scroll_into_view_if_needed()
                page.wait_for_timeout(300)
                is_visible = elem.is_visible()
                debug_info['debug']['element_visible'] = is_visible
                print(f"         After scroll: {is_visible}")
        except Exception as e:
            debug_info['debug']['error'] = f"Visibility check failed: {str(e)}"
            print(f"         ❌ Visibility check error: {e}")
        
        # ===== STEP 2: Check element bounding box =====
        print(f"      [DEBUG] Checking element bounding box...")
        
        try:
            bbox = elem.bounding_box()
            debug_info['debug']['element_bounding_box'] = bbox
            if bbox:
                print(f"         Bounding box: x={bbox['x']}, y={bbox['y']}, w={bbox['width']}, h={bbox['height']}")
            else:
                print(f"         ⚠️  Bounding box is None (element might be hidden)")
        except Exception as e:
            debug_info['debug']['error'] = f"Bounding box check failed: {str(e)}"
            print(f"         ❌ Bounding box error: {e}")
        
        # ===== STEP 3: Hover and wait =====
        print(f"      [DEBUG] Hovering over element...")
        
        try:
            elem.hover()
            page.wait_for_timeout(500)
            print(f"         Hover successful, waited 500ms")
        except Exception as e:
            debug_info['debug']['error'] = f"Hover failed: {str(e)}"
            print(f"         ❌ Hover error: {e}")
            return debug_info
        
        # ===== STEP 4: Check tooltip visibility =====
        print(f"      [DEBUG] Checking tooltip visibility...")
        
        try:
            tooltip_elem = page.locator(f"#{tooltip_id}").first
            tooltip_visible = tooltip_elem.is_visible()
            debug_info['debug']['tooltip_visible'] = tooltip_visible
            print(f"         Tooltip visible: {tooltip_visible}")
            
            if not tooltip_visible:
                print(f"         ⚠️  Tooltip NOT visible after hover")
        except Exception as e:
            debug_info['debug']['error'] = f"Tooltip visibility check failed: {str(e)}"
            print(f"         ❌ Tooltip visibility error: {e}")
        
        # ===== STEP 5: Check tooltip text content =====
        print(f"      [DEBUG] Checking tooltip text content...")
        
        try:
            tooltip_text = page.evaluate(f"""
                (id) => {{
                    const tooltip = document.getElementById(id);
                    if (tooltip) {{
                        return {{
                            'text': tooltip.textContent.trim(),
                            'visible': tooltip.offsetParent !== null,
                            'display': window.getComputedStyle(tooltip).display,
                            'visibility': window.getComputedStyle(tooltip).visibility,
                            'opacity': window.getComputedStyle(tooltip).opacity,
                            'innerHTML': tooltip.innerHTML.substring(0, 100)
                        }};
                    }}
                    return null;
                }}
            """, tooltip_id)
            
            if tooltip_text:
                debug_info['debug']['tooltip_text'] = tooltip_text['text']
                print(f"         Tooltip text: {tooltip_text['text']}")
                print(f"         Tooltip offsetParent !== null: {tooltip_text['visible']}")
                print(f"         Tooltip display: {tooltip_text['display']}")
                print(f"         Tooltip visibility: {tooltip_text['visibility']}")
                print(f"         Tooltip opacity: {tooltip_text['opacity']}")
            else:
                print(f"         ⚠️  Tooltip element not found in DOM")
        except Exception as e:
            debug_info['debug']['error'] = f"Tooltip text check failed: {str(e)}"
            print(f"         ❌ Tooltip text error: {e}")
        
        # ===== STEP 6: Extract tooltip value =====
        print(f"      [DEBUG] Extracting tooltip value...")
        
        try:
            tooltip_value = page.evaluate(f"""
                (id) => {{
                    const tooltip = document.getElementById(id);
                    if (tooltip && tooltip.offsetParent !== null) {{
                        return tooltip.textContent.trim();
                    }}
                    return null;
                }}
            """, tooltip_id)
            
            debug_info['value'] = tooltip_value
            debug_info['debug']['extraction_result'] = tooltip_value
            
            if tooltip_value:
                print(f"         ✅ Extracted: {tooltip_value}")
            else:
                print(f"         ⚠️  Extraction returned None")
        except Exception as e:
            debug_info['debug']['error'] = f"Extraction failed: {str(e)}"
            print(f"         ❌ Extraction error: {e}")
        
    except Exception as e:
        debug_info['debug']['error'] = f"Unexpected error: {str(e)}"
        print(f"      ❌ Unexpected error: {e}")
    
    return debug_info


def run_debug_extraction(page, screenshot_folder):
    """
    Run tooltip extraction with detailed debug logging.
    
    Args:
        page: Playwright page object
        screenshot_folder: Path to save screenshots
    
    Returns:
        dict: Debug results for all tokens
    """
    
    print("\n" + "=" * 100)
    print("🔍 TOOLTIP EXTRACTION DEBUG - DETAILED LOGGING")
    print("=" * 100)
    
    # Get all table rows
    table_rows = page.locator('table tbody tr').all()
    print(f"\n📊 Total rows: {len(table_rows)}")
    print(f"📊 Expected tooltips: {len(table_rows)} rows × 4 types = {len(table_rows) * 4} tooltips")
    
    all_results = {
        'timestamp': datetime.now().isoformat(),
        'total_rows': len(table_rows),
        'tokens': []
    }
    
    failing_tokens = []
    
    for row_idx, row in enumerate(table_rows):
        try:
            # Get token name
            try:
                token_name_elem = row.locator('td').first
                token_name = token_name_elem.text_content().strip() if token_name_elem else f"token_{row_idx}"
                token_name = token_name.replace('/', '_').replace(' ', '_')[:20]
            except:
                token_name = f"token_{row_idx}"
            
            # Get chain (if available)
            try:
                chain_elem = row.locator('td').nth(0)
                chain = chain_elem.text_content().strip() if chain_elem else "unknown"
            except:
                chain = "unknown"
            
            # Get price (if available)
            try:
                price_elem = row.locator('td').nth(2)
                price = price_elem.text_content().strip() if price_elem else "unknown"
            except:
                price = "unknown"
            
            token_result = {
                'row_index': row_idx,
                'token_name': token_name,
                'chain': chain,
                'price': price,
                'tooltips': {}
            }
            
            print(f"\n{'=' * 100}")
            print(f"📍 Row {row_idx:03d} | Token: {token_name} | Chain: {chain} | Price: {price}")
            print(f"{'=' * 100}")
            
            # ===== PRICE TOOLTIP =====
            print(f"\n🔹 PRICE TOOLTIP")
            price_elem = row.locator('[data-tooltip-id*="price-tooltip"]').first
            if price_elem.count() > 0:
                try:
                    price_tooltip_id = price_elem.get_attribute('data-tooltip-id')
                    print(f"   Tooltip ID: {price_tooltip_id}")
                    
                    result = extract_tooltip_with_debug(page, price_elem, price_tooltip_id, token_name, row_idx, 'price')
                    token_result['tooltips']['price'] = result
                    
                    if result['value']:
                        print(f"   ✅ SUCCESS: {result['value']}")
                    else:
                        print(f"   ❌ FAILED: No value extracted")
                        failing_tokens.append({
                            'row': row_idx,
                            'token': token_name,
                            'chain': chain,
                            'price': price,
                            'type': 'price',
                            'debug': result['debug']
                        })
                except Exception as e:
                    print(f"   ❌ ERROR: {e}")
                    token_result['tooltips']['price'] = {'value': None, 'error': str(e)}
            else:
                print(f"   ⚠️  Element not found (selector didn't match)")
                token_result['tooltips']['price'] = {'value': None, 'error': 'Element not found'}
            
            # ===== SHARE TOOLTIP =====
            print(f"\n🔹 SHARE TOOLTIP")
            share_elem = row.locator('[data-tooltip-id*="share-tooltip"]').first
            if share_elem.count() > 0:
                try:
                    share_tooltip_id = share_elem.get_attribute('data-tooltip-id')
                    print(f"   Tooltip ID: {share_tooltip_id}")
                    
                    result = extract_tooltip_with_debug(page, share_elem, share_tooltip_id, token_name, row_idx, 'share')
                    token_result['tooltips']['share'] = result
                    
                    if result['value']:
                        print(f"   ✅ SUCCESS: {result['value']}")
                    else:
                        print(f"   ❌ FAILED: No value extracted")
                        failing_tokens.append({
                            'row': row_idx,
                            'token': token_name,
                            'chain': chain,
                            'price': price,
                            'type': 'share',
                            'debug': result['debug']
                        })
                except Exception as e:
                    print(f"   ❌ ERROR: {e}")
                    token_result['tooltips']['share'] = {'value': None, 'error': str(e)}
            else:
                print(f"   ⚠️  Element not found (selector didn't match)")
                token_result['tooltips']['share'] = {'value': None, 'error': 'Element not found'}
            
            # ===== AMOUNT TOOLTIP =====
            print(f"\n🔹 AMOUNT TOOLTIP")
            amount_elem = row.locator('[data-tooltip-id*="amount-tooltip"]').first
            if amount_elem.count() > 0:
                try:
                    amount_tooltip_id = amount_elem.get_attribute('data-tooltip-id')
                    print(f"   Tooltip ID: {amount_tooltip_id}")
                    
                    result = extract_tooltip_with_debug(page, amount_elem, amount_tooltip_id, token_name, row_idx, 'amount')
                    token_result['tooltips']['amount'] = result
                    
                    if result['value']:
                        print(f"   ✅ SUCCESS: {result['value']}")
                    else:
                        print(f"   ❌ FAILED: No value extracted")
                        failing_tokens.append({
                            'row': row_idx,
                            'token': token_name,
                            'chain': chain,
                            'price': price,
                            'type': 'amount',
                            'debug': result['debug']
                        })
                except Exception as e:
                    print(f"   ❌ ERROR: {e}")
                    token_result['tooltips']['amount'] = {'value': None, 'error': str(e)}
            else:
                print(f"   ⚠️  Element not found (selector didn't match)")
                token_result['tooltips']['amount'] = {'value': None, 'error': 'Element not found'}
            
            # ===== PRICE 24H TOOLTIP =====
            print(f"\n🔹 PRICE 24H TOOLTIP")
            price_24h_elem = row.locator('[data-tooltip-id*="price-24h-tooltip"]').first
            if price_24h_elem.count() > 0:
                try:
                    price_24h_tooltip_id = price_24h_elem.get_attribute('data-tooltip-id')
                    print(f"   Tooltip ID: {price_24h_tooltip_id}")
                    
                    result = extract_tooltip_with_debug(page, price_24h_elem, price_24h_tooltip_id, token_name, row_idx, 'price_24h')
                    token_result['tooltips']['price_24h'] = result
                    
                    if result['value']:
                        print(f"   ✅ SUCCESS: {result['value']}")
                    else:
                        print(f"   ❌ FAILED: No value extracted")
                        failing_tokens.append({
                            'row': row_idx,
                            'token': token_name,
                            'chain': chain,
                            'price': price,
                            'type': 'price_24h',
                            'debug': result['debug']
                        })
                except Exception as e:
                    print(f"   ❌ ERROR: {e}")
                    token_result['tooltips']['price_24h'] = {'value': None, 'error': str(e)}
            else:
                print(f"   ⚠️  Element not found (selector didn't match)")
                token_result['tooltips']['price_24h'] = {'value': None, 'error': 'Element not found'}
            
            all_results['tokens'].append(token_result)
            
        except Exception as e:
            print(f"❌ Row {row_idx}: Critical error - {e}")
            continue
    
    # ===== SUMMARY =====
    print(f"\n\n{'=' * 100}")
    print("📊 SUMMARY")
    print(f"{'=' * 100}")
    print(f"Total rows processed: {len(all_results['tokens'])}")
    print(f"Total failing tooltips: {len(failing_tokens)}")
    
    if failing_tokens:
        print(f"\n❌ FAILING TOKENS:")
        print(f"{'=' * 100}")
        
        # Group by token
        failing_by_token = {}
        for fail in failing_tokens:
            token_key = f"{fail['token']} ({fail['chain']})"
            if token_key not in failing_by_token:
                failing_by_token[token_key] = []
            failing_by_token[token_key].append(fail)
        
        for token_key, fails in failing_by_token.items():
            print(f"\n🔴 {token_key}")
            for fail in fails:
                print(f"   Row {fail['row']:03d} | Type: {fail['type']} | Price: {fail['price']}")
                if fail['debug']['error']:
                    print(f"      Error: {fail['debug']['error']}")
                if not fail['debug']['element_visible']:
                    print(f"      ⚠️  Element not visible")
                if not fail['debug']['tooltip_visible']:
                    print(f"      ⚠️  Tooltip not visible after hover")
                if fail['debug']['extraction_result'] is None:
                    print(f"      ⚠️  Extraction returned None")
    
    # Save results to JSON
    results_file = f"{screenshot_folder}/debug_results.json"
    with open(results_file, 'w') as f:
        json.dump(all_results, f, indent=2)
    print(f"\n✅ Debug results saved to: {results_file}")
    
    all_results['failing_tokens'] = failing_tokens
    return all_results


# Usage in run_overview.py:
# 
# # After getting table_rows, run debug extraction:
# debug_results = run_debug_extraction(page, screenshot_folder)
# 
# # Then use the results to identify root causes
# print(f"\nFailing tokens: {len(debug_results['failing_tokens'])}")
# for fail in debug_results['failing_tokens']:
#     print(f"  - {fail['token']} ({fail['chain']}): {fail['type']} tooltip")
