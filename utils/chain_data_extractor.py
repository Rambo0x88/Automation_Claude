#!/usr/bin/env python3
"""
Chain Data Extractor - Extract Chain Breakdown and Chain Activity tables from DAM
"""

import pandas as pd
from datetime import datetime
import os


class ChainDataExtractor:
    """Extract Chain Breakdown, Chain Activity, and Address Breakdown data from DAM portfolio page"""

    def __init__(self, page):
        self.page = page

    def extract_chain_breakdown_table(self):
        """
        Extract Chain Breakdown table data (main chains only: Base, Ethereum, Binance Smart Chain, etc.)
        Columns: Name, Share, Token, Protocol, Value
        """
        print(f"📊 Extracting Chain Breakdown table...")

        chain_breakdown_data = []

        try:
            # Wait for page to load
            self.page.wait_for_timeout(2000)

            # Wait for table to be attached
            self.page.wait_for_selector("tbody tr", state="attached", timeout=5000)
            self.page.wait_for_timeout(1000)

            # Attempt to locate the Chain Breakdown table by heading text
            chain_breakdown_tbody = None
            chain_heading = self.page.locator("text=/^\\s*Chain Breakdown\\s*$/i")
            if chain_heading.count() > 0:
                print(f"   ✅ Found 'Chain Breakdown' heading")
                chain_table = chain_heading.nth(0).locator("xpath=following::table[1]")
                if chain_table.count() > 0:
                    chain_breakdown_tbody = chain_table.first.locator("tbody").first
                    print(f"   ✅ Located table associated with 'Chain Breakdown' heading")

            # Fallback: inspect all tbody elements if heading-based lookup fails
            if chain_breakdown_tbody is None:
                all_tbodies = self.page.locator("tbody").all()
                print(f"   ⚠️  Heading-based lookup failed. Inspecting {len(all_tbodies)} tbody elements")

                # Debug: Print first few rows from each tbody to identify the correct one
                for tbody_idx, tbody in enumerate(all_tbodies[:3]):  # Check first 3 tbodies
                    rows = tbody.locator("tr").all()
                    if len(rows) > 0:
                        first_row_text = rows[0].inner_text()[:100] if len(rows) > 0 else ""
                        print(f"   tbody[{tbody_idx}]: {len(rows)} rows, first row: {first_row_text}...")

                for tbody in all_tbodies:
                    rows = tbody.locator("tr").all()
                    if len(rows) > 0:
                        first_row_text = rows[0].inner_text().lower()
                        # Heuristic: chain rows contain chain names and percentages but no token prices
                        has_chain_keyword = any(chain in first_row_text for chain in ["base", "ethereum", "binance", "polygon", "arbitrum", "optimism"])
                        contains_price_symbol = "$" in first_row_text and "%" in first_row_text and "share" not in first_row_text
                        if has_chain_keyword and not contains_price_symbol:
                            chain_breakdown_tbody = tbody
                            print(f"   ✅ Found Chain Breakdown tbody via heuristic matching")
                            break

            if chain_breakdown_tbody is None:
                print(f"   ⚠️  Could not identify Chain Breakdown tbody")
                return chain_breakdown_data

            table_rows = chain_breakdown_tbody.locator("tr").all()

            # Filter out rows without td elements
            table_rows = [row for row in table_rows if row.locator("td, [role='cell']").count() > 0]

            print(f"   Found {len(table_rows)} total rows in Chain Breakdown tbody")

            # Define headers for main rows
            headers = ["Name", "Share", "Token", "Protocol", "Value"]

            # For Chain Breakdown, we'll extract all rows and let the data structure determine validity
            # Rows should have 5 or 6 cells (with or without expand button)
            for idx, row in enumerate(table_rows, 1):
                try:
                    # Extract data from each row
                    cells = row.locator("td, [role='cell']").all()

                    if len(cells) == 0:
                        continue

                    # Extract text from each cell
                    cell_values = []
                    for i, cell in enumerate(cells):
                        try:
                            text = cell.inner_text().strip()
                            # Split by newline and take first non-empty line
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            text = lines[0] if lines else text
                            cell_values.append(text)
                        except:
                            text = cell.text_content().strip()
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            text = lines[0] if lines else text
                            cell_values.append(text)

                    # Only process rows with at least 5 values (Name, Share, Token, Protocol, Value)
                    # Skip rows that don't have the expected structure
                    if len(cell_values) < 5:
                        continue

                    # Create row data
                    row_data = {"Row_Number": idx}

                    # Map cell values to headers, starting from the correct offset
                    # If first cell is empty or button, data starts from cell 1
                    # Otherwise data starts from cell 0
                    offset = 1 if (not cell_values[0] or len(cell_values[0]) < 2) else 0

                    for i in range(len(headers)):
                        if offset + i < len(cell_values):
                            column_name = headers[i]
                            row_data[column_name] = cell_values[offset + i]

                    # Only add if we have the Name field
                    if row_data.get("Name"):
                        chain_breakdown_data.append(row_data)

                except Exception as e:
                    print(f"   ⚠️  Error extracting row {idx}: {e}")
                    continue

            print(f"✅ Extracted {len(chain_breakdown_data)} rows from Chain Breakdown table (main rows only)")

        except Exception as e:
            print(f"❌ Error extracting Chain Breakdown data: {e}")
            import traceback
            traceback.print_exc()

        return chain_breakdown_data

    def extract_chain_activity_table(self):
        """
        Extract Chain Activity table data (from second tbody)
        Columns: Chain, Trxns, Gas Spent, Current Value
        """
        print(f"📊 Extracting Chain Activity table...")

        chain_activity_data = []

        try:
            # Wait for page to load
            self.page.wait_for_timeout(2000)

            # Wait for table to be attached
            self.page.wait_for_selector("tbody tr", state="attached", timeout=5000)
            self.page.wait_for_timeout(1000)

            # Find the specific table whose headers match the Chain Activity columns,
            # preferably by locating the "Chain Activity" heading.
            def normalize_header(text: str) -> str:
                """Normalize header text for easier comparison"""
                return text.lower().replace(" ", "").replace("_", "")

            # Chain Activity table may have 3 or 4 columns depending on data availability
            # Required: Chain, Trxns, Gas Spent
            # Optional: Current Value
            expected_headers_3col = {normalize_header(text) for text in ["Chain", "Trxns", "Gas Spent"]}
            expected_headers_4col = {normalize_header(text) for text in ["Chain", "Trxns", "Gas Spent", "Current Value"]}
            chain_activity_tbody = None
            target_table_index = -1

            activity_heading = self.page.locator("text=/^\\s*Chain Activity\\s*$/i")
            if activity_heading.count() > 0:
                print(f"   ✅ Found 'Chain Activity' heading")
                activity_table = activity_heading.nth(0).locator("xpath=following::table[1]")
                if activity_table.count() > 0:
                    chain_activity_tbody = activity_table.first.locator("tbody").first
                    print(f"   ✅ Located table associated with 'Chain Activity' heading")

            # If heading-based lookup failed, scan all tables for matching headers
            tables = [] if chain_activity_tbody else self.page.locator("table").all()
            if chain_activity_tbody is None:
                print(f"   ⚠️  Heading-based lookup failed. Scanning {len(tables)} table elements on Chain tab")

                for table_idx, table in enumerate(tables):
                    try:
                        header_cells = table.locator("thead th, thead td, tr th").all()
                        header_texts = []
                        for cell in header_cells:
                            header_text = cell.text_content().strip()
                            if header_text:
                                header_texts.append(header_text)

                        normalized_headers = {normalize_header(text) for text in header_texts}

                        # Match either 3-column or 4-column format
                        if expected_headers_3col.issubset(normalized_headers) or expected_headers_4col.issubset(normalized_headers):
                            chain_activity_tbody = table.locator("tbody").first
                            target_table_index = table_idx
                            print(f"   ✅ Identified Chain Activity table at index {table_idx} with headers {header_texts}")
                            break
                    except Exception as e:
                        print(f"   ⚠️  Error inspecting table {table_idx}: {e}")

            # Fallback to legacy approach (second tbody) if detection failed
            if chain_activity_tbody is None:
                all_tbodies = self.page.locator("tbody").all()
                print(f"   ⚠️  Could not match headers directly. Falling back to tbody index heuristic (total tbodies: {len(all_tbodies)})")
                if len(all_tbodies) >= 2:
                    chain_activity_tbody = all_tbodies[1]
                    target_table_index = 1
                else:
                    print(f"   ❌ Chain Activity tbody not found (need at least 2 tbody elements)")
                    return chain_activity_data

            print(f"   Using tbody from table index {target_table_index} for Chain Activity extraction")
            table_rows = chain_activity_tbody.locator("tr").all()

            # Filter out rows without td elements
            table_rows = [row for row in table_rows if row.locator("td, [role='cell']").count() > 0]

            print(f"   Found {len(table_rows)} rows in Chain Activity table")

            # Define headers - support both 3-column and 4-column formats
            # 3-column format: Chain, Trxns, Gas_Spent
            # 4-column format: Chain, Trxns, Gas_Spent, Current_Value
            headers_3col = ["Chain", "Trxns", "Gas_Spent"]
            headers_4col = ["Chain", "Trxns", "Gas_Spent", "Current_Value"]

            for idx, row in enumerate(table_rows, 1):
                try:
                    # Extract data from each row
                    cells = row.locator("td, [role='cell']").all()

                    if len(cells) == 0:
                        continue

                    # Extract text from each cell
                    cell_values = []
                    for cell in cells:
                        try:
                            text = cell.inner_text().strip()
                            # Split by newline and take first non-empty line
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            text = lines[0] if lines else text
                            cell_values.append(text)
                        except:
                            text = cell.text_content().strip()
                            lines = [line.strip() for line in text.split('\n') if line.split()]
                            text = lines[0] if lines else text
                            cell_values.append(text)

                    # Process rows with 3 or 4 columns
                    if len(cell_values) >= 3:
                        # Create row data
                        row_data = {"Row_Number": idx}

                        # Determine which header format to use based on cell count
                        headers = headers_4col if len(cell_values) >= 4 else headers_3col

                        # Map cell values to headers
                        for i in range(min(len(cell_values), len(headers))):
                            column_name = headers[i]
                            row_data[column_name] = cell_values[i]

                        chain_activity_data.append(row_data)

                except Exception as e:
                    print(f"   ⚠️  Error extracting row {idx}: {e}")
                    continue

            print(f"✅ Extracted {len(chain_activity_data)} rows from Chain Activity table")

        except Exception as e:
            print(f"❌ Error extracting Chain Activity data: {e}")
            import traceback
            traceback.print_exc()

        return chain_activity_data

    def extract_address_breakdown_table(self):
        """
        Extract Address Breakdown table data
        Columns: Address, Chain, Share, Token, Protocol, Value
        """
        print(f"📊 Extracting Address Breakdown table...")

        address_breakdown_data = []

        try:
            # Wait for page to load
            self.page.wait_for_timeout(2000)

            # Wait for table to be attached
            self.page.wait_for_selector("tbody tr", state="attached", timeout=5000)
            self.page.wait_for_timeout(1000)

            # Find all tbody elements
            all_tbodies = self.page.locator("tbody").all()
            print(f"   Found {len(all_tbodies)} tbody elements")

            # The Address Breakdown table should be the first tbody
            # It contains addresses like "0x56D0...99CF81"
            address_breakdown_tbody = None
            for tbody_idx, tbody in enumerate(all_tbodies):
                rows = tbody.locator("tr").all()
                if len(rows) > 0:
                    # Check if first row contains an address pattern (0x...)
                    first_row_text = rows[0].inner_text().lower()
                    if "0x" in first_row_text and "100.00%" in first_row_text:  # Address with full share
                        address_breakdown_tbody = tbody
                        print(f"   ✅ Found Address Breakdown tbody at index {tbody_idx}")
                        break

            if address_breakdown_tbody is None:
                print(f"   ⚠️  Could not identify Address Breakdown tbody")
                return address_breakdown_data

            table_rows = address_breakdown_tbody.locator("tr").all()

            # Filter out rows without td elements
            table_rows = [row for row in table_rows if row.locator("td, [role='cell']").count() > 0]

            print(f"   Found {len(table_rows)} total rows in Address Breakdown tbody")

            # Define headers
            headers = ["Address", "Chain", "Share", "Token", "Protocol", "Value"]

            for idx, row in enumerate(table_rows, 1):
                try:
                    # Extract data from each row
                    cells = row.locator("td, [role='cell']").all()

                    if len(cells) == 0:
                        continue

                    # Extract text from each cell
                    cell_values = []
                    for cell in cells:
                        try:
                            text = cell.inner_text().strip()
                            # Split by newline and take first non-empty line
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            text = lines[0] if lines else text
                            cell_values.append(text)
                        except:
                            text = cell.text_content().strip()
                            lines = [line.strip() for line in text.split('\n') if line.strip()]
                            text = lines[0] if lines else text
                            cell_values.append(text)

                    # Only process rows with at least 6 values (Address, Chain, Share, Token, Protocol, Value)
                    if len(cell_values) < 6:
                        continue

                    # Create row data
                    row_data = {"Row_Number": idx}

                    # Map cell values to headers, starting from the correct offset
                    # If first cell is empty or button, data starts from cell 1
                    # Otherwise data starts from cell 0
                    offset = 1 if (not cell_values[0] or len(cell_values[0]) < 2) else 0

                    for i in range(len(headers)):
                        if offset + i < len(cell_values):
                            column_name = headers[i]
                            row_data[column_name] = cell_values[offset + i]

                    # Only add if we have the Address field
                    if row_data.get("Address"):
                        address_breakdown_data.append(row_data)

                except Exception as e:
                    print(f"   ⚠️  Error extracting row {idx}: {e}")
                    continue

            print(f"✅ Extracted {len(address_breakdown_data)} rows from Address Breakdown table")

        except Exception as e:
            print(f"❌ Error extracting Address Breakdown data: {e}")
            import traceback
            traceback.print_exc()

        return address_breakdown_data


def export_chain_data_to_excel(chain_breakdown_data, chain_activity_data, username, output_dir="test-results/chain-data"):
    """
    Export Chain Breakdown and Chain Activity data to Excel file

    Args:
        chain_breakdown_data: List of chain breakdown row dictionaries
        chain_activity_data: List of chain activity row dictionaries
        username: Username for filename
        output_dir: Directory to save the Excel file

    Returns:
        Path to the created Excel file
    """

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Generate filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"chain_data_{username}_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, filename)

    print(f"\n📁 Exporting chain data to Excel: {output_path}")

    try:
        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

            # Export Chain Breakdown data
            if chain_breakdown_data:
                df_breakdown = pd.DataFrame(chain_breakdown_data)

                # Reorganize columns
                breakdown_column_order = ['Row_Number', 'Name', 'Share', 'Token', 'Protocol', 'Value']
                existing_cols = [col for col in breakdown_column_order if col in df_breakdown.columns]
                df_breakdown = df_breakdown[existing_cols]

                df_breakdown.to_excel(writer, sheet_name='Chain Breakdown', index=False)
                print(f"   ✅ Exported {len(chain_breakdown_data)} rows to 'Chain Breakdown' sheet")

            # Export Chain Activity data
            if chain_activity_data:
                df_activity = pd.DataFrame(chain_activity_data)

                # Reorganize columns
                activity_column_order = ['Row_Number', 'Chain', 'Trxns', 'Gas_Spent', 'Current_Value']
                existing_cols = [col for col in activity_column_order if col in df_activity.columns]
                df_activity = df_activity[existing_cols]

                df_activity.to_excel(writer, sheet_name='Chain Activity', index=False)
                print(f"   ✅ Exported {len(chain_activity_data)} rows to 'Chain Activity' sheet")

            # Auto-adjust column widths
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter

                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass

                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Excel file saved: {output_path}")
        return output_path

    except Exception as e:
        print(f"❌ Error exporting to Excel: {e}")
        raise
