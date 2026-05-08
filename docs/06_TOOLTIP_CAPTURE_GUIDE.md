# Tooltip Capture Guide - Best Practices

## Current Issues

1. **Timing**: 300ms wait may be too short for tooltip to render
2. **Index Mismatch**: Loop index doesn't match actual row index
3. **No Retry**: Fails silently if tooltip doesn't appear
4. **No Validation**: Doesn't verify tooltip content is loaded

## Better Approaches

### Approach 1: Increase Wait Time & Add Retry Logic

```python
def extract_tooltip_with_retry(page, elem, tooltip_id, max_retries=3):
    """Extract tooltip with retry logic"""
    for attempt in range(max_retries):
        try:
            # Hover over element
            elem.hover()
            
            # Wait longer for tooltip to appear
            page.wait_for_timeout(500 + (attempt * 200))  # 500ms, 700ms, 900ms
            
            # Find tooltip div
            tooltip_div = page.locator(f'#{tooltip_id}').first
            
            # Wait for tooltip to be visible
            tooltip_div.wait_for(state="visible", timeout=2000)
            
            # Get text content
            tooltip_text = tooltip_div.text_content().strip()
            
            # Validate content is not empty
            if tooltip_text and len(tooltip_text) > 0:
                return tooltip_text
            
        except Exception as e:
            if attempt < max_retries - 1:
                page.wait_for_timeout(200)  # Wait before retry
                continue
            else:
                return None
    
    return None
```

### Approach 2: Use Attribute-Based Extraction (More Reliable)

Instead of hovering and reading DOM, read tooltip data from element attributes:

```python
def extract_tooltip_from_attributes(page, elem):
    """Extract tooltip from element attributes (more reliable)"""
    try:
        # Check for aria-label attribute
        aria_label = elem.get_attribute('aria-label')
        if aria_label:
            return aria_label
        
        # Check for title attribute
        title = elem.get_attribute('title')
        if title:
            return title
        
        # Check for data-tooltip attribute
        data_tooltip = elem.get_attribute('data-tooltip')
        if data_tooltip:
            return data_tooltip
        
        # Check for data-content attribute
        data_content = elem.get_attribute('data-content')
        if data_content:
            return data_content
        
    except Exception as e:
        print(f"Error extracting from attributes: {e}")
    
    return None
```

### Approach 3: Use JavaScript to Extract Tooltip

```python
def extract_tooltip_via_javascript(page, elem, tooltip_id):
    """Extract tooltip using JavaScript (fastest & most reliable)"""
    try:
        # Use JavaScript to get tooltip content
        tooltip_content = page.evaluate(f"""
            (tooltipId) => {{
                const tooltip = document.getElementById(tooltipId);
                if (tooltip && tooltip.offsetParent !== null) {{
                    return tooltip.textContent.trim();
                }}
                return null;
            }}
        """, tooltip_id)
        
        return tooltip_content
    except Exception as e:
        print(f"Error extracting via JavaScript: {e}")
    
    return None
```

### Approach 4: Parallel Extraction (Faster)

```python
from concurrent.futures import ThreadPoolExecutor

def extract_all_tooltips_parallel(page, tooltip_triggers, tooltip_type="price"):
    """Extract all tooltips in parallel"""
    tooltips_map = {}
    
    def extract_single_tooltip(idx, elem):
        try:
            tooltip_id = elem.get_attribute('data-tooltip-id')
            if not tooltip_id:
                return idx, None
            
            # Hover
            elem.hover()
            page.wait_for_timeout(500)
            
            # Extract via JavaScript (faster)
            tooltip_text = page.evaluate(f"""
                (id) => {{
                    const el = document.getElementById(id);
                    return el ? el.textContent.trim() : null;
                }}
            """, tooltip_id)
            
            return idx, tooltip_text
        except:
            return idx, None
    
    # Use ThreadPoolExecutor for parallel extraction
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [
            executor.submit(extract_single_tooltip, idx, elem)
            for idx, elem in enumerate(tooltip_triggers)
        ]
        
        for future in futures:
            idx, tooltip_text = future.result()
            if tooltip_text:
                tooltips_map[idx] = tooltip_text
    
    return tooltips_map
```

### Approach 5: Use Playwright's Built-in Tooltip Handling

```python
def extract_tooltip_builtin(page, elem):
    """Use Playwright's built-in tooltip handling"""
    try:
        # Get the element's bounding box
        box = elem.bounding_box()
        
        # Move mouse to element (triggers tooltip)
        page.mouse.move(box['x'] + box['width']/2, box['y'] + box['height']/2)
        
        # Wait for tooltip to appear
        page.wait_for_timeout(800)
        
        # Look for tooltip in common selectors
        tooltip_selectors = [
            '[role="tooltip"]',
            '.tooltip',
            '.popover',
            '[class*="tooltip"]',
            '[class*="popover"]'
        ]
        
        for selector in tooltip_selectors:
            tooltip = page.locator(selector).first
            if tooltip.is_visible(timeout=1000):
                return tooltip.text_content().strip()
        
    except Exception as e:
        print(f"Error: {e}")
    
    return None
```

## Recommended Solution

**Use Approach 3 (JavaScript) + Approach 1 (Retry Logic)**

```python
def extract_tooltip_best_practice(page, elem, tooltip_id, max_retries=3):
    """Best practice: JavaScript extraction with retry"""
    for attempt in range(max_retries):
        try:
            # Hover to trigger tooltip
            elem.hover()
            page.wait_for_timeout(400 + (attempt * 100))
            
            # Extract via JavaScript (fastest & most reliable)
            tooltip_text = page.evaluate(f"""
                (id) => {{
                    const tooltip = document.getElementById(id);
                    if (tooltip && tooltip.offsetParent !== null) {{
                        return tooltip.textContent.trim();
                    }}
                    return null;
                }}
            """, tooltip_id)
            
            # Validate
            if tooltip_text and len(tooltip_text) > 0:
                return tooltip_text
            
        except Exception as e:
            if attempt < max_retries - 1:
                page.wait_for_timeout(200)
                continue
        
        return None
```

## Implementation Steps

1. **Replace current tooltip extraction** with JavaScript-based approach
2. **Add retry logic** for failed extractions
3. **Add validation** to ensure tooltip content is not empty
4. **Increase wait times** from 300ms to 500-800ms
5. **Test with multiple tokens** to verify accuracy

## Expected Improvements

- ✅ More accurate tooltip capture
- ✅ Fewer failed cases
- ✅ Faster extraction (JavaScript is faster than DOM traversal)
- ✅ Better error handling
- ✅ Automatic retry on failure

## Testing

```python
# Test the new extraction
tooltip_text = extract_tooltip_best_practice(page, elem, tooltip_id)
print(f"Extracted: {tooltip_text}")

# Verify it's not empty
assert tooltip_text is not None
assert len(tooltip_text) > 0
```

## Performance Comparison

| Method | Speed | Reliability | Accuracy |
|--------|-------|-------------|----------|
| Current (DOM) | Slow | Low | Medium |
| JavaScript | Fast | High | High |
| Attributes | Very Fast | Medium | Medium |
| Parallel | Very Fast | High | High |

**Recommendation**: Use JavaScript + Retry (Approach 3 + 1)
