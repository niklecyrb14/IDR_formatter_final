import pandas as pd
import os
from datetime import datetime, timedelta
import re
from colorama import init, Fore, Style

# Initialize colorama for Windows compatibility
init()

# Color shortcuts
WHITE = Fore.WHITE  # Default text color
GREEN = Fore.LIGHTGREEN_EX
RESET = Style.RESET_ALL

def oprint(*args, **kwargs):
    """Print in default color"""
    # Just use regular print - no color change needed
    import builtins
    builtins.print(*args, **kwargs)

def is_comed_format(input_file):
    """Check if this is a COMED format file (with INTERVAL USAGE DATA header and KW_INTERVAL columns)"""
    file_ext = os.path.splitext(input_file)[1].lower()
    
    try:
        if file_ext == '.csv':
            with open(input_file, 'r') as f:
                first_lines = f.read(500)
            return 'INTERVAL USAGE DATA' in first_lines and 'KW_INTERVAL' in first_lines
        elif file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_file, header=None, nrows=20)
            content = ' '.join(str(v) for v in df.values.flatten())
            return 'INTERVAL USAGE DATA' in content or 'KW_INTERVAL' in content
        return False
    except:
        return False

def read_comed_format(input_file):
    """
    Read COMED format files and convert to standard datetime/usage format.
    
    This format has:
    - Header section with "INTERVAL USAGE DATA" and metadata
    - Data rows with RECORDING_DT (date), METER_NBR, and KW_INTERVAL_1 through KW_INTERVAL_48 (30-min intervals)
    - Multiple meters that need to be summed together
    - KW_INTERVAL_1 = 12:00-12:30 AM, KW_INTERVAL_48 = 11:30 PM-12:00 AM
    - Values are in kW (not kWh), so must be converted based on interval length:
      - 15-min data: divide by 4 to get kWh
      - 30-min data: divide by 2 to get kWh
      - Hourly data: no conversion needed
    """
    oprint("  Detected COMED format")
    
    file_ext = os.path.splitext(input_file)[1].lower()
    
    # Find the header row
    if file_ext == '.csv':
        with open(input_file, 'r') as f:
            lines = f.readlines()
        
        header_row = None
        for i, line in enumerate(lines):
            if 'Service Agreement/Choice ID' in line and 'METER_NBR' in line and 'RECORDING_DT' in line:
                header_row = i
                break
        
        if header_row is None:
            raise ValueError("Could not find COMED data header row")
        
        df = pd.read_csv(input_file, skiprows=header_row)
    else:
        # Excel file - find header row
        df_raw = pd.read_excel(input_file, header=None)
        header_row = None
        for i, row in df_raw.iterrows():
            row_str = ' '.join(str(v) for v in row.values)
            if 'Service Agreement' in row_str and 'METER_NBR' in row_str and 'RECORDING_DT' in row_str:
                header_row = i
                break
        
        if header_row is None:
            raise ValueError("Could not find COMED data header row")
        
        df = pd.read_excel(input_file, header=header_row)
    
    # Filter to just valid data rows (numeric METER_NBR)
    df = df[pd.to_numeric(df['METER_NBR'], errors='coerce').notna()].copy()
    
    # Filter for CHANNEL_NBR = 1 only (if column exists)
    if 'CHANNEL_NBR' in df.columns:
        total_rows = len(df)
        df = df[df['CHANNEL_NBR'] == 1].copy()
        filtered_rows = len(df)
        if filtered_rows < total_rows:
            oprint(f"  Filtered for CHANNEL_NBR=1 ({filtered_rows} of {total_rows} rows)")
    
    # Get unique meters
    meters = df['METER_NBR'].unique()
    oprint(f"  Found {len(meters)} meter(s) to combine")
    
    # Get interval columns (KW_INTERVAL_1 through KW_INTERVAL_48/96)
    interval_cols = [c for c in df.columns if c.startswith('KW_INTERVAL')]
    num_intervals = len([c for c in interval_cols if int(re.search(r'(\d+)', c).group(1)) <= 96])
    
    # Determine interval type based on number of columns
    # 96 intervals = 15-min, 48 intervals = 30-min, 24 intervals = hourly
    if num_intervals >= 96:
        interval_minutes = 15
    elif num_intervals >= 48:
        interval_minutes = 30
    else:
        interval_minutes = 60
    
    max_intervals = 1440 // interval_minutes  # Max intervals per day
    
    # Determine kW to kWh conversion factor
    # COMED data is in kW, need to convert to kWh based on interval length
    if interval_minutes == 15:
        kw_to_kwh_factor = 0.25  # 15 min = 1/4 hour
        oprint(f"  Found {len(interval_cols)} interval columns ({interval_minutes}-min data, converting kW to kWh ÷4)")
    elif interval_minutes == 30:
        kw_to_kwh_factor = 0.5   # 30 min = 1/2 hour
        oprint(f"  Found {len(interval_cols)} interval columns ({interval_minutes}-min data, converting kW to kWh ÷2)")
    else:
        kw_to_kwh_factor = 1.0   # Hourly = no conversion needed
        oprint(f"  Found {len(interval_cols)} interval columns ({interval_minutes}-min data)")
    
    # Convert interval values to numeric
    for col in interval_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Group by date and sum all meters together
    df_summed = df.groupby('RECORDING_DT')[interval_cols].sum().reset_index()
    
    oprint(f"  Combined {len(df)} rows into {len(df_summed)} dates")
    
    # Convert to long format (datetime, usage)
    records = []
    
    for _, row in df_summed.iterrows():
        date_str = row['RECORDING_DT']
        try:
            base_date = pd.to_datetime(date_str)
        except:
            continue
        
        # Process each interval
        for col in interval_cols:
            # Extract interval number
            match = re.search(r'KW_INTERVAL_(\d+)', col)
            if not match:
                continue
            
            interval_num = int(match.group(1))
            
            # Skip intervals beyond max (DST extras handled separately)
            if interval_num > max_intervals:
                continue
            
            # Calculate end time for this interval
            # Interval 1 ends at first interval_minutes, etc.
            total_minutes = interval_num * interval_minutes
            hour = total_minutes // 60
            minute = total_minutes % 60
            
            if hour == 24:
                interval_dt = base_date + timedelta(days=1)
            else:
                interval_dt = base_date + timedelta(hours=hour, minutes=minute)
            
            value = row[col]
            if pd.notna(value):
                # Convert kW to kWh using the interval-based factor
                kwh_value = float(value) * kw_to_kwh_factor
                records.append({'datetime': interval_dt, 'usage': kwh_value})
    
    result_df = pd.DataFrame(records)
    result_df = result_df.sort_values('datetime').reset_index(drop=True)
    
    # Sum any duplicate timestamps (shouldn't happen but just in case)
    result_df = result_df.groupby('datetime')['usage'].sum().reset_index()
    result_df = result_df.sort_values('datetime').reset_index(drop=True)
    
    oprint(f"  Converted to {len(result_df)} interval records")
    
    return result_df

def is_duq_format(input_file):
    """Check if this is a DUQ (Duquesne Light) format file with Customer Identity header and Detailed Interval Usage"""
    file_ext = os.path.splitext(input_file)[1].lower()

    try:
        if file_ext == '.csv':
            with open(input_file, 'r') as f:
                first_lines = f.read(3000)
            return 'Customer Identity' in first_lines and 'Detailed Interval Usage' in first_lines
        elif file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_file, header=None, nrows=40)
            content = ' '.join(str(v) for v in df.values.flatten())
            return 'Customer Identity' in content and 'Detailed Interval Usage' in content
        return False
    except:
        return False

def read_duq_format(input_file):
    """
    Read DUQ (Duquesne Light) format files and convert to standard datetime/usage format.

    This format has:
    - Header section with metadata (Customer Identity, EDC=Duquesne Light Co., etc.)
    - Summarized Monthly Billed Usage section (skipped)
    - Detailed Interval Usage section with hourly data:
      - Reading Date column + columns 1-24 (hours) interleaved with QTY columns + Quality column
      - Column 1 = hour ending 1:00 (00:00), Column 24 = hour ending 24:00 (23:00)
      - Data is already hourly kWh
    """
    oprint("  Detected DUQ (Duquesne Light) format")

    file_ext = os.path.splitext(input_file)[1].lower()

    # Find the "Detailed Interval Usage" header row
    if file_ext == '.csv':
        with open(input_file, 'r') as f:
            lines = f.readlines()

        detail_row = None
        for i, line in enumerate(lines):
            if 'Detailed Interval Usage' in line:
                detail_row = i
                break

        if detail_row is None:
            raise ValueError("Could not find 'Detailed Interval Usage' section in DUQ file")

        # The header row (Reading Date, 1, 1 QTY, 2, ...) is the next row
        header_row = detail_row + 1
        df = pd.read_csv(input_file, skiprows=header_row, header=0)
    else:
        df_raw = pd.read_excel(input_file, header=None)
        detail_row = None
        for i, row in df_raw.iterrows():
            if any('Detailed Interval Usage' in str(v) for v in row.values):
                detail_row = i
                break

        if detail_row is None:
            raise ValueError("Could not find 'Detailed Interval Usage' section in DUQ file")

        header_row = detail_row + 1
        df = pd.read_excel(input_file, header=header_row)

    # Identify the usage columns (numeric column names 1-24, skip QTY and Quality columns)
    usage_cols = []
    for col in df.columns:
        col_str = str(col).strip()
        if col_str.isdigit() and 1 <= int(col_str) <= 24:
            usage_cols.append(col)

    oprint(f"  Found {len(usage_cols)} hourly interval columns")

    # Build datetime/usage pairs
    all_rows = []
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]).strip()
        if not date_str or date_str == 'nan':
            continue

        try:
            date = pd.to_datetime(date_str)
        except:
            continue

        for col in usage_cols:
            hour_num = int(str(col).strip())  # 1-24
            hour_label = hour_num - 1  # Column 1 → 00:00, Column 24 → 23:00

            dt = date + timedelta(hours=hour_label)

            val = row[col]
            if pd.notna(val):
                try:
                    usage = float(val)
                    all_rows.append({'datetime': dt, 'usage': usage})
                except (ValueError, TypeError):
                    continue

    result_df = pd.DataFrame(all_rows)

    if result_df.empty:
        raise ValueError("No valid interval data found in DUQ file")

    result_df['datetime'] = pd.to_datetime(result_df['datetime'])
    result_df = result_df.sort_values('datetime').reset_index(drop=True)
    oprint(f"  Read {len(result_df)} hourly records")

    # Fill partial days (VEE cutoff) using data from 7 days prior (same weekday)
    # Skip the first and last days of the dataset as those are natural boundaries, not VEE issues
    existing_dts = set(result_df['datetime'])
    result_df['date'] = result_df['datetime'].dt.date
    hours_per_day = result_df.groupby('date').size()

    first_date = hours_per_day.index.min()
    last_date = hours_per_day.index.max()
    partial_days = hours_per_day[(hours_per_day < 24) & (hours_per_day.index != first_date) & (hours_per_day.index != last_date)]

    if len(partial_days) > 0:
        oprint(f"  Found {len(partial_days)} partial day(s) (VEE cutoff) - filling from same weekday prior week")
        fill_rows = []

        for day, hour_count in partial_days.items():
            day_dt = pd.Timestamp(day)
            missing_hours = []
            for h in range(24):
                dt = day_dt + timedelta(hours=h)
                if dt not in existing_dts:
                    missing_hours.append(h)

            if not missing_hours:
                continue

            # Look for donor day: 7 days prior, then 14, 21, etc.
            donor_day = None
            for weeks_back in range(1, 8):
                candidate = day_dt - timedelta(days=7 * weeks_back)
                candidate_date = candidate.date()
                if candidate_date in hours_per_day.index and hours_per_day[candidate_date] == 24:
                    donor_day = candidate
                    break

            if donor_day is None:
                oprint(f"    {day}: missing hours {missing_hours} - no donor day found, skipping")
                continue

            # Pull values from the donor day for missing hours
            donor_data = result_df[result_df['date'] == donor_day.date()].set_index(
                result_df[result_df['date'] == donor_day.date()]['datetime'].dt.hour
            )

            filled_count = 0
            for h in missing_hours:
                if h in donor_data.index:
                    donor_usage = donor_data.loc[h, 'usage']
                    fill_dt = day_dt + timedelta(hours=h)
                    fill_rows.append({'datetime': fill_dt, 'usage': float(donor_usage)})
                    filled_count += 1

            oprint(f"    {day}: filled {filled_count} missing hour(s) from {donor_day.date()}")

        if fill_rows:
            fill_df = pd.DataFrame(fill_rows)
            fill_df['datetime'] = pd.to_datetime(fill_df['datetime'])
            result_df = pd.concat([result_df, fill_df], ignore_index=True)
            result_df = result_df.sort_values('datetime').reset_index(drop=True)
            oprint(f"  Total after filling: {len(result_df)} hourly records")

    result_df = result_df.drop(columns=['date'], errors='ignore')
    return result_df

def is_first_energy_format(input_file):
    """Check if this is a First Energy format file (with Customer Identifier and Detailed Interval Usage)"""
    file_ext = os.path.splitext(input_file)[1].lower()
    
    try:
        if file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_file, header=None, nrows=50)
            col0_values = [str(v) for v in df.iloc[:, 0].values]
        elif file_ext == '.csv':
            # CSV may have inconsistent columns, so read with on_bad_lines='skip'
            # Try different encodings
            for encoding in ['utf-8', 'latin-1', None]:
                try:
                    df = pd.read_csv(input_file, header=None, nrows=50, on_bad_lines='skip', encoding=encoding)
                    break
                except:
                    continue
            col0_values = [str(v) for v in df.iloc[:, 0].values]
        else:
            return False
        
        # Check for First Energy markers in first column
        has_customer_id = any('Customer Identifier' in v for v in col0_values)
        
        return has_customer_id
    except Exception as e:
        return False

def read_first_energy_format(input_file):
    """
    Read First Energy format files and return a dict of {customer_id: dataframe} for each customer with interval data.
    
    This format has:
    - Multiple customers in one file, each starting with "Customer Identifier"
    - Some customers have "No Interval Data Found" (skip these)
    - Others have "Detailed Interval Usage" followed by "Reading Date" header row
    - Interval columns: 0015, 0030, 0045, 0100... (15-min intervals) with QTY columns between
    - DST columns are ignored
    """
    oprint("  Detected First Energy format")
    
    file_ext = os.path.splitext(input_file)[1].lower()
    
    if file_ext in ['.xlsx', '.xls']:
        df = pd.read_excel(input_file, header=None)
    else:
        # CSV has variable column counts - need to read with enough columns
        # First, find the max number of columns
        max_cols = 0
        with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                num_cols = len(line.split(','))
                if num_cols > max_cols:
                    max_cols = num_cols
        
        # Now read with that many columns
        col_names = list(range(max_cols))
        for encoding in ['utf-8', 'latin-1', None]:
            try:
                df = pd.read_csv(input_file, header=None, names=col_names, encoding=encoding)
                break
            except:
                continue
    
    # Find all customer sections
    customer_sections = []
    current_customer = None
    pending_interval = False

    for i, row in df.iterrows():
        cell_val = str(row[0]) if pd.notna(row[0]) else ''

        if 'Customer Identifier' in cell_val:
            # Extract customer ID (remove \t prefix if present)
            customer_id = str(row[1]).replace('\\t', '').replace('\t', '').strip()
            current_customer = {'id': customer_id, 'start_row': i, 'meter_blocks': []}
            customer_sections.append(current_customer)
            pending_interval = False

        elif 'No Interval Data Found' in cell_val and current_customer:
            pending_interval = False

        elif 'Detailed Interval Usage' in cell_val and current_customer:
            pending_interval = True

        elif 'Reading Date' in cell_val and current_customer and pending_interval:
            current_customer['meter_blocks'].append({'interval_start': i})
            pending_interval = False
    
    # Set end rows for each customer
    for i, customer in enumerate(customer_sections):
        if i + 1 < len(customer_sections):
            customer['end_row'] = customer_sections[i + 1]['start_row']
        else:
            customer['end_row'] = len(df)
    
    oprint(f"  Found {len(customer_sections)} customer(s)")
    
    # Process each customer with interval data
    customer_data = {}
    
    for customer in customer_sections:
        if not customer['meter_blocks']:
            oprint(f"    Customer {customer['id']}: No interval data - skipping")
            continue

        is_submeter = len(customer['meter_blocks']) > 1
        if is_submeter:
            oprint(f"    Customer {customer['id']}: Found {len(customer['meter_blocks'])} meters (submeter) - summing datasets")
        else:
            oprint(f"    Customer {customer['id']}: Processing interval data...")

        # Parse interval data from each meter block
        all_records = []

        for block_idx, block in enumerate(customer['meter_blocks']):
            # Get the header row and data rows
            header_row_idx = block['interval_start']
            header_row = df.iloc[header_row_idx]

            # Find data rows (from header+1 to end of customer section, stopping at empty rows)
            data_start = header_row_idx + 1
            data_end = customer['end_row']

            # Get interval data for this meter block
            records = []

            for row_idx in range(data_start, data_end):
                row = df.iloc[row_idx]

                # Stop if we hit an empty row or new section marker
                if pd.isna(row[0]) or 'Customer' in str(row[0]) or 'Detailed Interval Usage' in str(row[0]):
                    break

                # Parse the date
                try:
                    base_date = pd.to_datetime(row[0])
                except:
                    continue

                # Process each interval column (skip QTY columns and DST columns)
                for col_idx, col_name in enumerate(header_row):
                    col_str = str(col_name).strip()

                    # Skip non-interval columns
                    if col_str in ['Reading Date', 'nan', ''] or 'QTY' in col_str or 'DST' in col_str:
                        continue

                    # Skip "R " (Received) columns - we only want "D " (Delivered) or bare time values
                    if col_str.startswith('R '):
                        continue

                    # Strip "D " prefix if present (Delivered energy columns)
                    time_str_raw = col_str
                    if col_str.startswith('D '):
                        time_str_raw = col_str[2:].strip()

                    # Check if this looks like a time column (0015, 0030, etc. or 15, 30, 100, etc.)
                    try:
                        # Handle both "0015" string and 15.0 float formats
                        if isinstance(col_name, (int, float)) and not pd.isna(col_name):
                            time_val = int(col_name)
                        elif time_str_raw.isdigit():
                            time_val = int(time_str_raw)
                        else:
                            continue

                        # Convert to hour and minute
                        time_str = str(time_val).zfill(4)
                        hour = int(time_str[:-2]) if len(time_str) > 2 else 0
                        minute = int(time_str[-2:])

                        # Handle 2359 as the last interval of the day (actually ends at midnight)
                        # First Energy uses 2359 to represent the 23:45-00:00 interval
                        if time_val == 2359:
                            hour = 24
                            minute = 0
                        # Skip if not a valid interval (15-min, 30-min, or hourly)
                        elif minute not in [0, 15, 30, 45]:
                            continue

                        # Get the value
                        value = row[col_idx]
                        if pd.isna(value):
                            continue

                        # Create datetime
                        if hour == 24:
                            interval_dt = base_date + timedelta(days=1)
                        else:
                            interval_dt = base_date + timedelta(hours=hour, minutes=minute)

                        records.append({'datetime': interval_dt, 'usage': float(value)})

                    except (ValueError, TypeError):
                        continue

            if is_submeter:
                oprint(f"      Meter {block_idx + 1}: {len(records)} records")

            all_records.extend(records)

        if all_records:
            result_df = pd.DataFrame(all_records)
            result_df = result_df.sort_values('datetime').reset_index(drop=True)

            # Sum by datetime (handles both submeter summing and duplicate removal)
            result_df = result_df.groupby('datetime')['usage'].sum().reset_index()
            result_df = result_df.sort_values('datetime').reset_index(drop=True)

            customer_data[customer['id']] = result_df
            if is_submeter:
                oprint(f"      Summed to {len(result_df)} interval records")
            else:
                oprint(f"      Converted to {len(result_df)} interval records")
    
    return customer_data

def is_esg_multi_meter_format(input_file):
    """Check if this is an ESG format file with multiple meters that need to be summed."""
    file_ext = os.path.splitext(input_file)[1].lower()

    try:
        if file_ext in ['.xlsx', '.xls']:
            xlsx = pd.ExcelFile(input_file)
            if 'IDR Quantity' not in xlsx.sheet_names:
                return False
            df = pd.read_excel(xlsx, sheet_name='IDR Quantity', header=5, usecols=[3])
            if 'Meter Number' not in df.columns:
                return False
            unique_meters = df['Meter Number'].dropna().unique()
            return len(unique_meters) > 1
        elif file_ext == '.csv':
            with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
                for i, line in enumerate(f):
                    if i > 100:
                        break
                    if 'Report Period Date' in line and 'Interval Ending' in line and 'Meter Number' in line:
                        # Found header with Meter Number - read just that column to check for multiple meters
                        import csv
                        df = pd.read_csv(input_file, skiprows=i, usecols=['Meter Number'])
                        unique_meters = df['Meter Number'].dropna().unique()
                        return len(unique_meters) > 1
            return False
        return False
    except:
        return False

def read_esg_multi_meter_format(input_file):
    """
    Read ESG format files with multiple meters and sum their usage values.

    Same structure as standard ESG but with a Meter Number column containing
    multiple meters. Values for the same date/interval across meters are summed.
    """
    oprint("  Detected ESG multi-meter format")

    file_ext = os.path.splitext(input_file)[1].lower()

    # Read the file
    if file_ext in ['.xlsx', '.xls']:
        df = pd.read_excel(input_file, sheet_name='IDR Quantity', header=5)
    else:
        import csv
        header_row = None
        with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                row_str = ','.join(row)
                if 'Report Period Date' in row_str and 'Interval Ending' in row_str:
                    header_row = i
                    break
        if header_row is not None:
            df = pd.read_csv(input_file, skiprows=header_row, on_bad_lines='skip')
        else:
            df = pd.read_csv(input_file, on_bad_lines='skip')

    # Filter for KH Measurement Unit if the column exists
    if 'Measurement Unit' in df.columns:
        kh_count = (df['Measurement Unit'] == 'KH').sum()
        total_count = len(df)
        if kh_count > 0 and kh_count < total_count:
            oprint(f"  Filtering for 'KH' Measurement Unit ({kh_count} of {total_count} rows)")
            df = df[df['Measurement Unit'] == 'KH']

    # Report meters found
    if 'Meter Number' in df.columns:
        meters = df['Meter Number'].dropna().unique()
        oprint(f"  Found {len(meters)} meter(s) to combine: {', '.join(str(m) for m in meters)}")

    # Get interval columns (excluding DS columns)
    interval_cols = [col for col in df.columns if str(col).startswith('Interval Ending') and 'DS' not in str(col)]
    ds_cols = [col for col in df.columns if str(col).startswith('Interval Ending') and 'DS' in str(col)]

    oprint(f"  Found {len(interval_cols)} regular interval columns")
    oprint(f"  Found {len(ds_cols)} DST fall-back columns (ignoring to keep days uniform)")

    # Sum across meters for each date/interval by grouping by Report Period Date
    # This handles both multi-meter summing and duplicate date combining
    numeric_cols = interval_cols + ds_cols
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df_summed = df.groupby('Report Period Date')[numeric_cols].sum().reset_index()
    oprint(f"  Summed {len(df)} rows across meters to {len(df_summed)} unique dates")

    # Convert to long format (datetime, usage)
    records = []

    for _, row in df_summed.iterrows():
        date_val = row['Report Period Date']
        if pd.isna(date_val):
            continue

        try:
            date_int = int(float(date_val))
            date_str = str(date_int)
            if len(date_str) != 8:
                continue
            base_date = pd.to_datetime(date_str, format='%Y%m%d')
        except (ValueError, TypeError):
            continue

        # Process regular interval columns only (ignore DS columns)
        for col in interval_cols:
            time_match = re.search(r'(\d{4})$', col)
            if time_match:
                time_str = time_match.group(1)
                hour = int(time_str[:2])
                minute = int(time_str[2:])

                if hour == 24:
                    interval_dt = base_date + timedelta(days=1)
                else:
                    interval_dt = base_date + timedelta(hours=hour, minutes=minute)

                value = row[col]
                if pd.notna(value) and value != 0:
                    try:
                        records.append({'datetime': interval_dt, 'usage': float(value)})
                    except (ValueError, TypeError):
                        continue

    result_df = pd.DataFrame(records)

    # Group by datetime and sum (in case of any remaining overlaps)
    result_df = result_df.groupby('datetime')['usage'].sum().reset_index()
    result_df = result_df.sort_values('datetime').reset_index(drop=True)

    # Detect if hourly data - shift timestamps to start-of-hour
    if len(result_df) >= 2:
        time_diff = (result_df['datetime'].iloc[1] - result_df['datetime'].iloc[0]).total_seconds() / 60
        if time_diff == 60:
            oprint(f"  Adjusting hourly ESG timestamps to start-of-hour format")
            result_df['datetime'] = result_df['datetime'] - timedelta(hours=1)

    oprint(f"  Converted to {len(result_df)} interval records")

    return result_df

def is_esg_format(input_file):
    """Check if this is an ESG format file (Excel with IDR Quantity sheet, or CSV with same columns)"""
    file_ext = os.path.splitext(input_file)[1].lower()
    
    try:
        if file_ext in ['.xlsx', '.xls']:
            xlsx = pd.ExcelFile(input_file)
            return 'IDR Quantity' in xlsx.sheet_names
        elif file_ext == '.csv':
            # ESG CSVs can have multi-section format with headers scattered throughout
            # Read line by line to find the interval data section
            with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
                for i, line in enumerate(f):
                    if i > 100:  # Don't scan more than 100 lines
                        break
                    if 'Report Period Date' in line and 'Interval Ending' in line:
                        return True
            return False
        return False
    except:
        return False

def is_bge_format(input_file):
    """Check if this is a BGE format file (with RdgDate/ReadDate and Kwh columns)"""
    file_ext = os.path.splitext(input_file)[1].lower()
    
    try:
        # Read just the header row based on file type
        if file_ext == '.csv':
            df = pd.read_csv(input_file, nrows=1, index_col=False)
        elif file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_file, nrows=1)
        else:
            return False
        
        cols = [c.lower() for c in df.columns]
        
        # Check for BGE 15-min format (RdgDate) or BGE Hourly format (ReadDate)
        has_date_col = 'rdgdate' in cols or 'readdate' in cols
        has_kwh = 'kwh' in cols
        has_endtime = 'endtime' in cols
        has_starttime = 'starttime' in cols
        
        return has_date_col and has_kwh and (has_endtime or has_starttime)
    except:
        return False

def read_bge_format(input_file):
    """
    Read BGE format files (CSV or Excel) and convert to standard datetime/usage format.
    
    Supports multiple BGE formats:
    - BGE 15-min (EI): RdgDate, EndTime (15, 30, 45, 100, 115...), Kwh
    - BGE Hourly (HI): ReadDate, StartTime/EndTime (0000/0059, 0100/0159...), Kwh
      - StartTime 0000 with EndTime 0059 means hour 0 (00:00)
      - StartTime 0100 with EndTime 0159 means hour 1 (01:00)
    
    DST handling: For fall-back duplicates (e.g., two 1 AM hours on Nov 2), 
    keep only the first occurrence instead of summing.
    """
    file_ext = os.path.splitext(input_file)[1].lower()
    
    if file_ext == '.csv':
        df = pd.read_csv(input_file, index_col=False)
    else:
        df = pd.read_excel(input_file)
    
    # Determine which date column exists (case-insensitive search)
    cols_lower = {c.lower(): c for c in df.columns}
    
    if 'rdgdate' in cols_lower:
        date_col = cols_lower['rdgdate']
        format_type = 'EI'
        oprint("  Detected BGE 15-min format")
    elif 'readdate' in cols_lower:
        date_col = cols_lower['readdate']
        format_type = 'HI'
        oprint("  Detected BGE Hourly format")
    else:
        raise ValueError("Could not find date column (RdgDate or ReadDate)")
    
    kwh_col = cols_lower['kwh']
    
    # Check for StartTime column (newer HI format)
    has_starttime = 'starttime' in cols_lower
    if has_starttime:
        starttime_col = cols_lower['starttime']
        oprint(f"  Using columns: {date_col}, {starttime_col}, {kwh_col}")
    else:
        endtime_col = cols_lower['endtime']
        oprint(f"  Using columns: {date_col}, {endtime_col}, {kwh_col}")
    
    records = []
    for _, row in df.iterrows():
        date_val = row[date_col]
        kwh_val = row[kwh_col]
        
        if pd.isna(date_val) or pd.isna(kwh_val):
            continue
        
        # Parse date
        try:
            base_date = pd.to_datetime(date_val)
        except:
            continue
        
        if format_type == 'HI':
            if has_starttime:
                # New HI format: StartTime like 0000, 0100, 0200...
                # StartTime 0000 = hour 0, 0100 = hour 1, etc.
                starttime_val = row[starttime_col]
                if pd.isna(starttime_val):
                    continue
                starttime_int = int(starttime_val)
                # Parse HHMM format
                hour = starttime_int // 100
                minute = 0  # Use :00 for hourly data labels
            else:
                # Old HI format: EndTime like 59, 159, 259...
                endtime_val = row[cols_lower['endtime']]
                if pd.isna(endtime_val):
                    continue
                endtime_int = int(endtime_val)
                # EndTime 59 = hour 0, 159 = hour 1, etc.
                # Parse by looking at digits before the 59
                endtime_str = str(endtime_int)
                if endtime_str.endswith('59'):
                    hour = int(endtime_str[:-2]) if len(endtime_str) > 2 else 0
                else:
                    hour = endtime_int // 100
                minute = 0
        else:
            # EI format: EndTime like 15, 30, 45, 100, 115...
            # These are actual interval end times
            endtime_val = row[cols_lower['endtime']]
            if pd.isna(endtime_val):
                continue
            endtime_int = int(endtime_val)
            if endtime_int < 100:
                hour = 0
                minute = endtime_int
            else:
                endtime_str = str(endtime_int).zfill(4)
                hour = int(endtime_str[:-2]) if len(endtime_str) > 2 else 0
                minute = int(endtime_str[-2:])
        
        # Create the datetime
        if hour == 24:
            interval_dt = base_date + timedelta(days=1)
        else:
            interval_dt = base_date + timedelta(hours=hour, minutes=minute)
        
        records.append({'datetime': interval_dt, 'usage': float(kwh_val)})
    
    result_df = pd.DataFrame(records)
    result_df = result_df.sort_values('datetime').reset_index(drop=True)
    
    # Remove duplicates - keep FIRST occurrence (for DST fall-back, don't sum the duplicate hours)
    result_df = result_df.drop_duplicates(subset=['datetime'], keep='first')
    result_df = result_df.sort_values('datetime').reset_index(drop=True)
    
    oprint(f"  Converted to {len(result_df)} interval records")
    
    return result_df

def read_esg_format(input_file):
    """
    Read ESG format files (Excel or CSV) and convert to standard datetime/usage format.
    This format has:
    - One row per day with columns for each 15-min interval
    - Report Period Date column in YYYYMMDD format
    - Interval Ending columns like "Interval Ending 0015", "Interval Ending 0030", etc.
    - Special "DS" columns for November DST fall-back extra hour (ignored to keep days uniform)
    - Sometimes duplicate rows for same date that need to be combined
    - CSV files may have multiple sections with different headers (Document Info, Transaction Info, etc.)
    """
    oprint("  Detected ESG format")
    
    file_ext = os.path.splitext(input_file)[1].lower()
    
    # Read the file based on type
    if file_ext in ['.xlsx', '.xls']:
        df = pd.read_excel(input_file, sheet_name='IDR Quantity', header=5)
        
        # Filter for KH Measurement Unit if the column exists (Excel files)
        if 'Measurement Unit' in df.columns:
            kh_count = (df['Measurement Unit'] == 'KH').sum()
            total_count = len(df)
            if kh_count > 0 and kh_count < total_count:
                oprint(f"  Filtering for 'KH' Measurement Unit ({kh_count} of {total_count} rows)")
                df = df[df['Measurement Unit'] == 'KH']
            elif kh_count == total_count:
                oprint(f"  All {total_count} rows have 'KH' Measurement Unit")
    else:
        # CSV - find the header row with "Report Period Date" and "Interval Ending"
        # Multi-section CSVs may have the interval data section much further down
        # Use csv module to correctly count rows regardless of column count variations
        import csv
        header_row = None
        with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                row_str = ','.join(row)
                if 'Report Period Date' in row_str and 'Interval Ending' in row_str:
                    header_row = i
                    break
        
        if header_row is not None:
            # Use skiprows to skip all rows before the header, then pandas uses row 0 as header
            df = pd.read_csv(input_file, skiprows=header_row, on_bad_lines='skip')
        else:
            df = pd.read_csv(input_file, on_bad_lines='skip')
    
    # Filter for KH Measurement Unit if the column exists
    # KH = kWh data (the proper interval data to use)
    # Other values like K1, KW, etc. should be excluded
    if 'Measurement Unit' in df.columns:
        kh_count = (df['Measurement Unit'] == 'KH').sum()
        total_count = len(df)
        if kh_count > 0 and kh_count < total_count:
            oprint(f"  Filtering for 'KH' Measurement Unit ({kh_count} of {total_count} rows)")
            df = df[df['Measurement Unit'] == 'KH']
        elif kh_count == total_count:
            oprint(f"  All {total_count} rows have 'KH' Measurement Unit")
    
    # Get interval columns (excluding DS columns for now)
    interval_cols = [col for col in df.columns if str(col).startswith('Interval Ending') and 'DS' not in str(col)]
    ds_cols = [col for col in df.columns if str(col).startswith('Interval Ending') and 'DS' in str(col)]
    
    oprint(f"  Found {len(interval_cols)} regular interval columns")
    oprint(f"  Found {len(ds_cols)} DST fall-back columns (ignoring to keep days uniform)")
    
    # Check for duplicate dates and combine them
    date_counts = df['Report Period Date'].value_counts()
    dups = date_counts[date_counts > 1]
    if len(dups) > 0:
        oprint(f"  Found {len(dups)} dates with multiple rows - combining...")
        
        # Re-read the original file to get all data for combining
        if file_ext in ['.xlsx', '.xls']:
            df_orig = pd.read_excel(input_file, sheet_name='IDR Quantity', header=5)
        else:
            # CSV - use same header row detection for multi-section files
            import csv
            header_row = None
            with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
                reader = csv.reader(f)
                for i, row in enumerate(reader):
                    row_str = ','.join(row)
                    if 'Report Period Date' in row_str and 'Interval Ending' in row_str:
                        header_row = i
                        break
            if header_row is not None:
                df_orig = pd.read_csv(input_file, skiprows=header_row, on_bad_lines='skip')
            else:
                df_orig = pd.read_csv(input_file, on_bad_lines='skip')
        
        # For each date, combine the interval values (take the first non-null value)
        combined_rows = []
        for date in df_orig['Report Period Date'].unique():
            date_rows = df_orig[df_orig['Report Period Date'] == date]
            combined_row = {'Report Period Date': date}
            
            # For each interval column, take the first non-null value
            for col in interval_cols + ds_cols:
                values = date_rows[col].dropna()
                if len(values) > 0:
                    combined_row[col] = values.iloc[0]
                else:
                    combined_row[col] = None
            
            combined_rows.append(combined_row)
        
        df = pd.DataFrame(combined_rows)
        oprint(f"  Combined to {len(df)} unique dates")
    
    # Convert to long format (datetime, usage)
    records = []
    
    for _, row in df.iterrows():
        # Skip rows with missing or invalid dates
        date_val = row['Report Period Date']
        if pd.isna(date_val):
            continue
        
        # Try to convert to integer - skip if it fails (non-numeric like "Transaction Count")
        try:
            date_int = int(float(date_val))
            date_str = str(date_int)
            # Basic validation: should be 8 digits (YYYYMMDD)
            if len(date_str) != 8:
                continue
            base_date = pd.to_datetime(date_str, format='%Y%m%d')
        except (ValueError, TypeError):
            continue
        
        # Process regular interval columns only (ignore DS columns for November fall-back)
        # This keeps all days uniform at 24 hours
        for col in interval_cols:
            # Extract time from column name like "Interval Ending 0015" -> "0015"
            time_match = re.search(r'(\d{4})$', col)
            if time_match:
                time_str = time_match.group(1)
                hour = int(time_str[:2])
                minute = int(time_str[2:])
                
                # Handle "2400" as next day 00:00
                if hour == 24:
                    interval_dt = base_date + timedelta(days=1)
                else:
                    interval_dt = base_date + timedelta(hours=hour, minutes=minute)
                
                value = row[col]
                if pd.notna(value):
                    try:
                        records.append({'datetime': interval_dt, 'usage': float(value)})
                    except (ValueError, TypeError):
                        continue
    
    result_df = pd.DataFrame(records)
    
    # Sort and remove exact duplicates, sum values for same datetime
    result_df = result_df.groupby('datetime')['usage'].sum().reset_index()
    result_df = result_df.sort_values('datetime').reset_index(drop=True)
    
    # Detect if this is hourly data (intervals are 60 minutes apart)
    # If so, shift timestamps back by 1 hour since "Interval Ending 0100" means the 00:00 hour
    if len(result_df) >= 2:
        time_diff = (result_df['datetime'].iloc[1] - result_df['datetime'].iloc[0]).total_seconds() / 60
        if time_diff == 60:
            oprint(f"  Adjusting hourly ESG timestamps to start-of-hour format")
            result_df['datetime'] = result_df['datetime'] - timedelta(hours=1)
    
    oprint(f"  Converted to {len(result_df)} interval records")
    
    return result_df

def fill_dst_gap_intervals(df, interval_minutes):
    """
    Detect and fill DST gaps in raw interval data (before resampling).
    DST spring forward creates a gap where intervals are missing.
    We insert the missing intervals with the average of the values before and after the gap.
    
    For 30-min data: gap jumps from 2:00 to 3:30, missing 2:30 and 3:00
    For 15-min data: gap jumps from 2:00 to 3:15, missing 2:15, 2:30, 2:45, 3:00
    """
    df = df.sort_values('datetime').reset_index(drop=True)
    
    expected_diff = timedelta(minutes=interval_minutes)
    rows_to_insert = []
    
    for i in range(len(df) - 1):
        current_time = df.loc[i, 'datetime']
        next_time = df.loc[i + 1, 'datetime']
        actual_diff = next_time - current_time
        
        # Check if there's a gap larger than expected (DST gap)
        if actual_diff > expected_diff * 1.5:
            # Calculate how many intervals are missing
            missing_intervals = int(actual_diff / expected_diff) - 1
            
            if missing_intervals > 0:
                # Verify it's likely a DST gap (March, around 1-3 AM)
                # Hour 0-3 covers cases where timestamps are labeled by start or end time
                if current_time.month == 3 and 0 <= current_time.hour <= 3:
                    # Get the average of the value before and after the gap
                    value_before = df.loc[i, 'usage']
                    value_after = df.loc[i + 1, 'usage']
                    avg_value = (value_before + value_after) / 2
                    
                    oprint(f"    DST gap detected: {current_time} -> {next_time}")
                    oprint(f"    Inserting {missing_intervals} interval(s) with value {avg_value:.3f}")
                    
                    # Insert each missing interval with the averaged value
                    for j in range(1, missing_intervals + 1):
                        missing_time = current_time + (expected_diff * j)
                        rows_to_insert.append({
                            'datetime': missing_time,
                            'usage': round(avg_value, 3)
                        })
    
    if rows_to_insert:
        new_rows_df = pd.DataFrame(rows_to_insert)
        df = pd.concat([df, new_rows_df], ignore_index=True)
        df = df.sort_values('datetime').reset_index(drop=True)
        oprint(f"    ✓ Filled {len(rows_to_insert)} missing interval(s)")
    
    return df

def format_single_dataset(df, detected_interval):
    """
    Format a single interval dataset into the output format.
    Returns a DataFrame ready for output.
    """
    # For sub-hourly data, trim any data after the last 0:00 (midnight) timestamp
    # For hourly data, keep all data as-is (no trimming)
    last_row_time = df['datetime'].iloc[-1]
    
    if detected_interval == 60:
        # Hourly data - no trimming, keep all data
        oprint(f"  Hourly data - keeping all records (no trimming)")
    elif last_row_time.hour == 0 and last_row_time.minute == 0:
        oprint(f"  Data already ends at 0:00, no trimming needed")
    else:
        midnight_mask = (df['datetime'].dt.hour == 0) & (df['datetime'].dt.minute == 0)
        if midnight_mask.any():
            last_midnight_idx = df[midnight_mask].index[-1]
            last_midnight_time = df.loc[last_midnight_idx, 'datetime']
            oprint(f"  Last 0:00 found at index {last_midnight_idx}: {last_midnight_time}")
            
            if last_midnight_idx < len(df) - 1:
                trimmed_count = len(df) - last_midnight_idx - 1
                df = df.iloc[:last_midnight_idx + 1].copy()
                oprint(f"  Trimmed {trimmed_count} extra interval(s) after last 0:00")
                oprint(f"  Data now ends at: {df['datetime'].iloc[-1]}")
                oprint(f"  New total raw records: {len(df)}")
    
    # Fill DST gaps in raw interval data BEFORE resampling
    df = fill_dst_gap_intervals(df, detected_interval)
    
    # Resample to hourly if needed (sum the values)
    df.set_index('datetime', inplace=True)
    
    if detected_interval == 60:
        # Data is already hourly - just aggregate any duplicates, no time shifting needed
        hourly_df = df.groupby(df.index).sum()
        hourly_df.reset_index(inplace=True)
    else:
        # Sub-hourly data needs resampling
        # closed='right' means interval (0:00, 1:00] goes to 1:00 bucket
        # label='left' means the bucket is labeled with its start time (00:00-23:00)
        hourly_df = df.resample('h', closed='right', label='left').sum()
        hourly_df.reset_index(inplace=True)
    
    # Sort newest to oldest for splitting into years
    combined_df = hourly_df.sort_values('datetime', ascending=False)
    combined_df.reset_index(drop=True, inplace=True)
    
    # Round usage to 3 decimal places
    combined_df['usage'] = combined_df['usage'].round(3)
    
    oprint(f"  Total hourly records: {len(combined_df)}")
    
    # Create output DataFrame
    output_data = {}
    
    # Column A & B: Full dataset (oldest to newest)
    full_dataset_sorted = combined_df.sort_values('datetime', ascending=True)
    output_data['Intv End Date/Time'] = full_dataset_sorted['datetime'].dt.strftime('%m/%d/%Y %H:%M')
    output_data[' Usage'] = full_dataset_sorted['usage']
    
    # Blank column C
    output_data[''] = ''
    
    # Split into 8760-hour segments (years)
    total_hours = len(combined_df)
    num_years = (total_hours // 8760) + (1 if total_hours % 8760 > 0 else 0)
    
    oprint(f"  Segmenting into {num_years} year(s)...")
    
    # Sort newest to oldest for slicing into years
    newest_first = combined_df.sort_values('datetime', ascending=False).reset_index(drop=True)
    
    for year_num in range(num_years):
        start_idx = year_num * 8760
        end_idx = min((year_num + 1) * 8760, total_hours)
        
        year_data = newest_first.iloc[start_idx:end_idx].copy()
        year_data = year_data.sort_values('datetime', ascending=True)
        
        year_col_name_date = f'YEAR {year_num + 1} - Intv End Date/Time'
        year_col_name_usage = f'YEAR {year_num + 1} -  Usage'
        
        year_dates = year_data['datetime'].dt.strftime('%m/%d/%Y %H:%M').tolist()
        year_usage = year_data['usage'].tolist()
        
        while len(year_dates) < total_hours:
            year_dates.append('')
            year_usage.append('')
        
        output_data[year_col_name_date] = year_dates
        output_data[year_col_name_usage] = year_usage
        
        if year_num < num_years - 1:
            output_data[f'  {year_num}'] = ''
    
    # Create DataFrame
    output_df = pd.DataFrame(output_data)
    
    # Add header row
    header_row = {}
    header_row['Intv End Date/Time'] = 'OUTPUT'
    header_row[' Usage'] = ''
    header_row[''] = ''
    
    for year_num in range(num_years):
        year_col_name_date = f'YEAR {year_num + 1} - Intv End Date/Time'
        year_col_name_usage = f'YEAR {year_num + 1} -  Usage'
        
        header_row[year_col_name_date] = f'YEAR {year_num + 1}'
        header_row[year_col_name_usage] = ''
        
        if year_num < num_years - 1:
            header_row[f'  {year_num}'] = ''
    
    output_df = pd.concat([pd.DataFrame([header_row]), output_df], ignore_index=True)
    
    return output_df

def format_interval_data(input_file):
    """Format interval data from CSV or Excel file"""
    
    filename = os.path.basename(input_file)
    file_ext = os.path.splitext(input_file)[1].lower()
    
    oprint(f"\nProcessing: {filename}")
    
    try:
        # Check if this is First Energy format (must check first as it has special multi-customer handling)
        if is_first_energy_format(input_file):
            customer_data = read_first_energy_format(input_file)
            
            if not customer_data:
                oprint("  ✗ No interval data found in any customer section")
                return None
            
            # Generate output filename (Excel with multiple sheets)
            base_name = os.path.splitext(filename)[0]
            output_dir = os.path.dirname(input_file)
            output_file = os.path.join(output_dir, f"{base_name}_formatted.xlsx")
            
            # Process each customer and write to separate sheets
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for customer_id, df in customer_data.items():
                    oprint(f"\n  Processing customer {customer_id}...")
                    
                    # Sort and remove duplicates
                    df = df.sort_values('datetime').reset_index(drop=True)
                    original_len = len(df)
                    df = df.drop_duplicates(subset=['datetime'], keep='first').reset_index(drop=True)
                    if len(df) < original_len:
                        print(f"  Removed {original_len - len(df)} duplicate timestamp(s)")
                    
                    # Detect interval
                    detected_interval = 15  # Default for First Energy
                    if len(df) > 1:
                        time_diff = (df['datetime'].iloc[1] - df['datetime'].iloc[0]).total_seconds() / 60
                        detected_interval = int(time_diff)
                    
                    oprint(f"  Detected interval: {detected_interval} minutes")
                    oprint(f"  Total raw records: {len(df)}")
                    oprint(f"  Data range: {df['datetime'].iloc[0]} to {df['datetime'].iloc[-1]}")
                    
                    # Format the dataset
                    output_df = format_single_dataset(df, detected_interval)
                    
                    # Write to sheet (sheet name limited to 31 chars)
                    sheet_name = str(customer_id)[:31]
                    output_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    oprint(f"  ✓ Added sheet: {sheet_name}")
            
            oprint(f"\n  ✓ Saved to: {output_file}")
            return output_file
        
        # Check if this is DUQ (Duquesne Light) format
        if is_duq_format(input_file):
            df = read_duq_format(input_file)
        # Check if this is COMED format (with INTERVAL USAGE DATA and KW_INTERVAL columns)
        elif is_comed_format(input_file):
            df = read_comed_format(input_file)
        # Check if this is ESG multi-meter format (multiple meters to sum)
        elif is_esg_multi_meter_format(input_file):
            df = read_esg_multi_meter_format(input_file)
        # Check if this is ESG format (Excel with IDR Quantity sheet, or CSV with same columns)
        elif is_esg_format(input_file):
            df = read_esg_format(input_file)
        # Check if this is BGE format (with RdgDate/ReadDate and Kwh)
        elif is_bge_format(input_file):
            df = read_bge_format(input_file)
        # Standard PSEG format - Excel
        elif file_ext in ['.xlsx', '.xls']:
            oprint("  Detected PSEG format (Excel)")
            df = pd.read_excel(input_file, skiprows=3, usecols=[0, 1], names=['datetime', 'usage'])
            df['datetime'] = pd.to_datetime(df['datetime'])
        # Standard PSEG format - CSV
        elif file_ext == '.csv':
            oprint("  Detected PSEG format (CSV)")
            try:
                df = pd.read_csv(input_file, sep=',', skiprows=3, usecols=[0, 1], names=['datetime', 'usage'])
            except:
                df = pd.read_csv(input_file, sep='\t', skiprows=3, usecols=[0, 1], names=['datetime', 'usage'])
            df['datetime'] = pd.to_datetime(df['datetime'])
        else:
            oprint(f"  ✗ Unsupported file type: {file_ext}")
            oprint("    Supported types: .csv, .xlsx, .xls")
            return None
        
        # Sort by datetime
        df = df.sort_values('datetime').reset_index(drop=True)
        
        # Remove any duplicate timestamps (keep first occurrence)
        original_len = len(df)
        df = df.drop_duplicates(subset=['datetime'], keep='first').reset_index(drop=True)
        if len(df) < original_len:
            oprint(f"  Removed {original_len - len(df)} duplicate timestamp(s)")
        
        # Detect interval
        detected_interval = None
        if len(df) > 1:
            time_diff = (df['datetime'].iloc[1] - df['datetime'].iloc[0]).total_seconds() / 60
            detected_interval = int(time_diff)
        
        oprint(f"  Detected interval: {detected_interval} minutes")
        oprint(f"  Total raw records: {len(df)}")
        oprint(f"  Data range: {df['datetime'].iloc[0]} to {df['datetime'].iloc[-1]}")
        
        # Format the dataset
        output_df = format_single_dataset(df, detected_interval)
        
        # Generate output filename
        base_name = os.path.splitext(filename)[0]
        output_dir = os.path.dirname(input_file)
        output_file = os.path.join(output_dir, f"{base_name}_formatted.csv")
        
        # Save to CSV
        output_df.to_csv(output_file, index=False)
        
        oprint(f"  ✓ Saved to: {output_file}")
        
        return output_file
        
    except Exception as e:
        oprint(f"  ✗ Error processing: {e}")
        import traceback
        traceback.print_exc()
        return None

# ============================================================
# MAIN
# ============================================================

import sys, io

# Ensure stdout can handle Unicode (fixes cp1252 terminals)
if sys.stdout.encoding and sys.stdout.encoding.lower().replace('-', '') != 'utf8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

print("""
============================================================

    █ █▀▀▀▄ █▀▀▀▄   █▀▀▀ █▀▀▀█ █▀▀▀▄ █   █ █▀▀▀█ ▀▀█▀▀ ▀▀█▀▀ █▀▀▀ █▀▀▀▄
    █ █   █ █▄▄▄▀   █▀▀▀ █   █ █▄▄▄▀ █ █ █ █▀▀▀█   █     █   █▀▀▀ █▄▄▄▀
    █ █▄▄▄▀ █   █   █    █▄▄▄█ █   █ █   █ █   █   █     █   █▄▄▄ █   █  v1.0.5

    [ A P   G A S   &   E L E C T R I C ]
============================================================
""")
print("Formats 15-min, 30-min, or hourly interval data to hourly.")
print("Supports: PSEG, ESG, ESG Multi-Meter, BGE, First Energy, COMED, DUQ formats")
print("File types: CSV and Excel (.csv, .xlsx, .xls)")

# Check if a file was dragged onto the exe (passed as argument)
if len(sys.argv) > 1:
    drag_drop_file = sys.argv[1].strip().strip('"').strip("'")
    if os.path.exists(drag_drop_file):
        print(f"\nFile received: {drag_drop_file}")
        result = format_interval_data(drag_drop_file)
        if result:
            print(f"\n{GREEN}" + "="*60)
            print("✓ Formatting complete!")
            print("="*60 + f"{RESET}")
        input("\nPress Enter to exit...")
        sys.exit(0)

while True:
    print("\nEnter the file path to format (or 'quit' to exit):")
    print("(You can also drag and drop a file onto this program)")
    user_input = input("> ").strip()
    
    # Remove PowerShell's & '...' wrapper from drag and drop
    # Handles: & 'path' or & "path" or just path with quotes
    if user_input.startswith("& "):
        user_input = user_input[2:].strip()
    
    # Remove surrounding quotes (single or double)
    if (user_input.startswith("'") and user_input.endswith("'")) or \
       (user_input.startswith('"') and user_input.endswith('"')):
        user_input = user_input[1:-1]
    
    # Also strip any remaining quotes
    user_input = user_input.strip('"').strip("'")
    
    if user_input.lower() in ['quit', 'exit', 'q']:
        print("\nGoodbye!")
        break
    
    if not user_input:
        print("No file path entered. Please try again.")
        continue
    
    if not os.path.exists(user_input):
        print(f"✗ File not found: {user_input}")
        print("  Please check the path and try again.")
        continue
    
    result = format_interval_data(user_input)
    
    if result:
        print(f"\n{GREEN}" + "="*60)
        print("✓ Formatting complete!")
        print("="*60 + f"{RESET}")

input("\nPress Enter to exit...")