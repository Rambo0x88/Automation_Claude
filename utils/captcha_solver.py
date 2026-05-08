"""
Captcha Solver using MCP (Model Context Protocol)
Handles Cloudflare Turnstile captcha with human-in-the-loop
"""

from playwright.sync_api import Page
import time


def solve_captcha_manual(page: Page, timeout: int = 180000) -> bool:
    """
    Pause automation and wait for HUMAN to manually solve the captcha

    Args:
        page: Playwright Page object
        timeout: Maximum time to wait in milliseconds (default: 3 minutes)

    Returns:
        bool: True if captcha was solved, False if timeout
    """
    print("\n" + "="*70)
    print("⚠️  CLOUDFLARE VERIFICATION DETECTED")
    print("="*70)
    print("\n🔴 PLEASE WAIT - Cloudflare is verifying your browser...")
    print("   ✓ This is automatic in most cases")
    print("   ✓ If you see a checkbox, click it to complete verification")
    print("   ✓ The test will continue automatically once verified\n")
    print(f"⏱️  Maximum wait time: {timeout/1000} seconds ({timeout/60000} minutes)")
    print(f"   (Usually completes within 10-30 seconds)")
    print("="*70 + "\n")

    try:
        # First, try to wait for Turnstile token (most common case)
        try:
            page.wait_for_function(
                """
                () => {
                    const cfToken = document.querySelector('input[name="cf-turnstile-response"]');
                    return cfToken && cfToken.value && cfToken.value.length > 10;
                }
                """,
                timeout=timeout
            )
            
            # Get the token to confirm it worked
            cf_token = page.input_value('input[name="cf-turnstile-response"]')
            print("\n" + "="*70)
            print(f"✅ CAPTCHA SOLVED SUCCESSFULLY (Turnstile token detected)!")
            print(f"   Token received: {cf_token[:20]}...")
            print(f"   Token length: {len(cf_token)} characters")
            print("="*70 + "\n")
            page.wait_for_timeout(1000)
            return True
        except:
            # If Turnstile token doesn't appear, check if challenge disappeared
            pass
        
        # Alternative: Wait for challenge elements to disappear and content to appear
        # This handles cases where Cloudflare does automatic browser check
        try:
            start_time = time.time()
            check_interval = 2000  # Check every 2 seconds
            max_iterations = timeout // check_interval

            for i in range(max_iterations):
                elapsed = int(time.time() - start_time)
                print(f"   ⏱️  Waiting for verification... ({elapsed}s / {timeout//1000}s)")

                # Check if challenge elements are still present
                challenge_elements = [
                    'text=Verify you are human',
                    'text=Verifying you are human',
                    'text=Checking your browser',
                    '#challenge-running',
                    '.cf-browser-verification',
                ]

                challenge_still_present = False
                for element in challenge_elements:
                    try:
                        if page.locator(element).count() > 0:
                            challenge_still_present = True
                            break
                    except:
                        continue

                # If challenge is gone, check for actual content
                if not challenge_still_present:
                    # Look for page content indicators
                    content_indicators = [
                        'tbody tr',  # Table rows (common on CoinGecko)
                        'main',  # Main content area
                        '[class*="table"]',  # Table elements
                    ]

                    for indicator in content_indicators:
                        try:
                            if page.locator(indicator).count() > 0:
                                print("\n" + "="*70)
                                print(f"✅ CAPTCHA/CHALLENGE COMPLETED (content detected)!")
                                print(f"   Challenge elements disappeared and page content is visible")
                                print(f"   Total time: {elapsed}s")
                                print("="*70 + "\n")
                                page.wait_for_timeout(1000)
                                return True
                        except:
                            continue

                # Wait before next check
                page.wait_for_timeout(check_interval)
                
            # If we get here, challenge wasn't resolved
            raise Exception("Challenge not solved - no token or content detected")
            
        except Exception as e:
            # If both methods fail, raise the exception
            raise e

    except Exception as e:
        print("\n" + "="*70)
        print("❌ CAPTCHA TIMEOUT!")
        print(f"   Error: {str(e)}")
        print("   The captcha was not solved within the timeout period")
        print("="*70 + "\n")
        return False


def wait_for_captcha_and_solve(page: Page, timeout: int = 180000) -> bool:
    """
    Check if captcha exists and wait for manual solution

    Args:
        page: Playwright Page object
        timeout: Maximum time to wait in milliseconds (default: 3 minutes)

    Returns:
        bool: True if no captcha or captcha solved, False if timeout
    """
    # Wait a moment for page to load (reduced from 3000ms to 1500ms)
    page.wait_for_timeout(1500)

    # Check for various Cloudflare challenge indicators
    captcha_indicators = [
        'iframe[src*="turnstile"]',  # Turnstile iframe
        'iframe[src*="challenges.cloudflare.com"]',  # Cloudflare challenge iframe
        'text=Verify you are human',  # Text indicator
        'text=Verifying you are human',  # Alternative text (most common)
        'text=Checking your browser',  # Browser check text
        'text=Just a moment',  # Cloudflare loading text
        'text=Please wait',  # Wait message
        '#challenge-running',  # Challenge running element
        '.cf-browser-verification',  # Cloudflare verification class
        '[class*="challenge"]',  # Any element with challenge in class
        'input[name="cf-turnstile-response"]',  # Turnstile response input
        '[id*="challenge"]',  # Any element with challenge in ID
        '[class*="cloudflare"]',  # Cloudflare class
    ]
    
    try:
        captcha_detected = False
        
        # Check each indicator
        for indicator in captcha_indicators:
            try:
                if indicator.startswith('text='):
                    # Text-based selector
                    text_content = indicator.replace('text=', '')
                    if page.locator(f'text={text_content}').count() > 0:
                        captcha_detected = True
                        print(f"🔒 Captcha detected: '{text_content}' text found")
                        break
                else:
                    # CSS selector
                    if page.locator(indicator).count() > 0:
                        captcha_detected = True
                        print(f"🔒 Captcha detected: {indicator}")
                        break
            except:
                continue
        
        if captcha_detected:
            print("🔒 Captcha/challenge detected on page - waiting for manual solution...")
            return solve_captcha_manual(page, timeout)
        else:
            print("✅ No captcha detected")
            return True
    except Exception as e:
        print(f"⚠️  Error checking for captcha: {e}")
        # If we can't check properly, assume no captcha to avoid blocking
        return True
