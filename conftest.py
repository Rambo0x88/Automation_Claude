"""
Root conftest.py - pytest configuration hooks
This must be in root directory to run before pytest-html initializes
"""
import os
from datetime import datetime
from pathlib import Path


def pytest_addoption(parser):
    """Add custom command line options"""
    parser.addoption(
        "--portfolio",
        action="store",
        default=None,
        help="Specify portfolio name to extract data from (e.g., --portfolio \"zg's address - 46\")"
    )
    parser.addoption(
        "--portfolioId",
        action="store",
        default=None,
        help="Navigate directly to portfolio by ID (e.g., --portfolioId \"83081753-e3af-440a-9081-740120c3840d\")"
    )


def pytest_configure(config):
    """
    Create test-results directory before pytest-html tries to write.
    Auto-generates a timestamped HTML report in test-results/reports/
    named after the test file(s) being run.
    """
    test_results_dir = Path("test-results")
    screenshots_dir = test_results_dir / "screenshots"
    reports_dir = test_results_dir / "reports"

    # Create directories
    test_results_dir.mkdir(exist_ok=True)
    screenshots_dir.mkdir(exist_ok=True)
    reports_dir.mkdir(exist_ok=True)

    # Only auto-set htmlpath if not already provided via --html flag
    htmlpath = config.getoption('htmlpath', default=None)
    if not htmlpath:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        report_name = f"report_{timestamp}.html"
        config.option.htmlpath = str((reports_dir / report_name).resolve())

    print(f"\n✅ Created directories: {test_results_dir}, {screenshots_dir}, {reports_dir}")
