# IDR Formatter - Technical Documentation

## Overview

The IDR Formatter is a Python-based command-line application that processes interval energy usage data from multiple utility formats and converts it into a standardized hourly output format.

**Version:** 1.0.1  
**Author:** AP Gas & Electric Texas  
**Language:** Python 3.x

---

## Technology Stack

### Programming Language
- **Python 3.x** - Core application language

### Required Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `pandas` | Latest | Data manipulation, CSV/Excel reading, datetime handling, resampling |
| `openpyxl` | Latest | Excel file (.xlsx) read/write support |
| `colorama` | Latest | Cross-platform colored terminal output (Windows compatible) |

### Standard Library Modules Used
- `os` - File path operations
- `sys` - Command-line arguments, system exit
- `re` - Regular expressions for pattern matching
- `datetime` - Date/time operations and arithmetic
- `traceback` - Error reporting

### Build Tool
- **PyInstaller** - Converts Python script to standalone Windows executable (.exe)

---

## Installation & Build

### Development Setup
```bash
# Install required packages
pip install pandas openpyxl colorama

# Run the script directly
python "IDR File Formatter.py"
```

### Building the Executable
```powershell
# Navigate to project directory
cd "C:\path\to\project"

# Clean previous builds
Remove-Item -Recurse -Force build -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force dist -ErrorAction SilentlyContinue

# Build with PyInstaller
python -m PyInstaller --onefile --hidden-import=openpyxl --hidden-import=colorama "IDR File Formatter.py"

# Optional: Add custom icon
python -m PyInstaller --onefile --hidden-import=openpyxl --hidden-import=colorama --icon=icon.ico "IDR File Formatter.py"
```

### Build Output
- Executable location: `dist/IDR File Formatter.exe`
- Build artifacts: `build/` folder (can be deleted)
- Spec file: `IDR File Formatter.spec` (PyInstaller configuration)

---

## Architecture

### File Structure
```
IDR File Formatter.py
│
├── IMPORTS & INITIALIZATION
│   ├── Package imports (pandas, os, datetime, re, colorama)
│   └── Color definitions and helper functions
│
├── FORMAT DETECTION FUNCTIONS
│   ├── is_comed_format()
│   ├── is_first_energy_format()
│   ├── is_esg_format()
│   ├── is_bge_format()
│   └── (PSEG is default fallback)
│
├── FORMAT READER FUNCTIONS
│   ├── read_comed_format()
│   ├── read_first_energy_format()
│   ├── read_esg_format()
│   └── read_bge_format()
│
├── DATA PROCESSING FUNCTIONS
│   ├── fill_dst_gap_intervals()
│   └── format_single_dataset()
│
├── MAIN PROCESSING FUNCTION
│   └── format_interval_data()
│
└── MAIN EXECUTION BLOCK
    ├── ASCII header display
    ├── Drag-and-drop handling
    └── Interactive input loop
```

---

## Supported Formats - Technical Details

### 1. PSEG Format
**Detection:** Default fallback when no other format matches  
**Structure:** Simple 2-column CSV/Excel with datetime and usage values  
**Header rows skipped:** 3  
**Columns used:** 0 (datetime), 1 (usage)

```python
df = pd.read_csv(input_file, skiprows=3, usecols=[0, 1], names=['datetime', 'usage'])
```

### 2. ESG Format
**Detection:** Excel file with "IDR Quantity" sheet, or CSV with "Report Period Date" and "Interval Ending" columns  
**Structure:** Wide format with one row per day, columns for each interval  
**Special handling:** 
- Combines duplicate date rows (takes first non-null value per interval)
- Ignores DST fall-back columns (columns containing "DS")
- Header row at index 5 for Excel files
- **Multi-section CSV support:** Handles ESG CSV files with multiple sections (Document Info, Transaction Info, etc.) by scanning for the correct header row
- **Measurement Unit filtering:** Filters for rows where `Measurement Unit = 'KH'` (kWh data) when multiple data sets exist (e.g., K1, K3, KH)

**Interval columns:** `Interval Ending 0015`, `Interval Ending 0030`, etc.

### 3. BGE Format (15-min and Hourly)
**Detection:** Columns `RdgDate` or `ReadDate`, plus `EndTime` and `Kwh`  
**Structure:** Long format with one row per interval  
**Variants:**
- **BGE 15-min:** Uses `RdgDate` column, EndTime values like 15, 30, 45, 100, 115...
- **BGE Hourly:** Uses `ReadDate` column, EndTime values like 59, 159, 259...

```python
# BGE 15-min: EndTime represents actual interval end
# EndTime 115 = 01:15 AM

# BGE Hourly: EndTime represents hour ending
# EndTime 159 = Hour 1 (1:00 AM)
```

### 4. First Energy Format
**Detection:** Contains "Customer Identifier" in first column  
**Structure:** Multiple customer sections in one file, each with metadata and interval data  
**Special handling:**
- Parses multiple customers, creates separate output tabs
- Skips customers with "No Interval Data Found"
- Handles variable column counts in CSV (scans for max columns first)
- Ignores QTY columns and DST columns
- **2359 column handling:** Treats the `2359` column as midnight (24:00) - this represents the last 15-minute interval of the day (23:45-00:00)

**Key markers:**
- `Customer Identifier` - Start of customer section
- `Detailed Interval Usage` - Indicates interval data present
- `Reading Date` - Header row for interval data

### 5. COMED Format
**Detection:** Contains "INTERVAL USAGE DATA" header and `KW_INTERVAL_*` columns  
**Structure:** Wide format with date rows, multiple meters to combine  
**Special handling:**
- Sums usage across all meters for same date/interval
- Supports 48 intervals (30-min) or 96 intervals (15-min)
- Finds header row dynamically by searching for column names

**Interval columns:** `KW_INTERVAL_1` through `KW_INTERVAL_48` (or 96)
- Interval 1 = 12:00-12:30 AM
- Interval 48 = 11:30 PM-12:00 AM

---

## Data Processing Pipeline

### Step 1: Format Detection
```
Input File → Check First Energy → Check COMED → Check ESG → Check BGE → Default PSEG
```

Each detection function reads a small portion of the file and looks for characteristic markers.

### Step 2: Data Reading
Each format reader converts the source data into a standardized DataFrame:
```python
DataFrame with columns: ['datetime', 'usage']
- datetime: pandas Timestamp object
- usage: float (kWh value)
```

### Step 3: Data Cleaning
```python
# Sort chronologically
df = df.sort_values('datetime')

# Remove duplicate timestamps (keep first)
df = df.drop_duplicates(subset=['datetime'], keep='first')
```

### Step 4: Interval Detection
```python
time_diff = (df['datetime'].iloc[1] - df['datetime'].iloc[0]).total_seconds() / 60
# Returns: 15, 30, or 60 minutes
```

### Step 5: DST Gap Filling
**Spring Forward (March):** Detects gaps larger than expected interval during 1-3 AM in March.
```python
# Example: 1:45 AM → 3:00 AM (gap of 75 minutes instead of 15)
# Action: Insert missing intervals with averaged value
avg_value = (value_before_gap + value_after_gap) / 2
```

**Fall Back (November):** Extra hour is ignored (DST columns excluded during reading).

### Step 6: Hourly Resampling
```python
df.set_index('datetime', inplace=True)
hourly_df = df.resample('h', closed='right', label='left').sum()
```
- `closed='right'`: Groups intervals by their end time
- `label='left'`: Labels buckets with start hour (0:00, 1:00, etc.)

### Step 7: Data Trimming
Removes any partial data after the last midnight:
```python
# Find last 0:00 timestamp, remove everything after
midnight_mask = (df['datetime'].dt.hour == 0) & (df['datetime'].dt.minute == 0)
last_midnight_idx = df[midnight_mask].index[-1]
df = df.iloc[:last_midnight_idx + 1]
```

### Step 8: Year Segmentation
Data is split into 8,760-hour segments (1 year):
```python
# Sort newest to oldest
# YEAR 1 = most recent 8,760 hours
# YEAR 2 = next oldest 8,760 hours
# etc.
```

### Step 9: Output Generation
Creates DataFrame with columns:
- Column A: Full dataset datetime (oldest to newest)
- Column B: Full dataset usage
- Column C: Blank separator
- Columns D-E: YEAR 1 datetime and usage
- Columns F: Blank separator
- Columns G-H: YEAR 2 datetime and usage
- (continues for additional years)

---

## Output Format Specification

### CSV Output (Most formats)
```
Intv End Date/Time, Usage, , YEAR 1 - Intv End Date/Time, YEAR 1 - Usage, , YEAR 2 - ...
OUTPUT, , , YEAR 1, , , YEAR 2, ...
01/01/2023 00:00, 123.456, , 01/15/2024 00:00, 234.567, , 01/15/2023 00:00, ...
```

### Excel Output (First Energy only)
- Multiple sheets, one per customer
- Sheet names: Customer ID numbers (truncated to 31 characters)
- Same column structure as CSV within each sheet

### Date Format
`MM/DD/YYYY HH:MM` (e.g., `01/15/2024 00:00`)

### Usage Values
- Rounded to 3 decimal places
- Units: kWh (kilowatt-hours)

---

## Error Handling

### File-Level Errors
- File not found → User-friendly message, prompt for retry
- Unsupported file type → Lists supported types
- Parse errors → Caught with traceback, continues to next file

### Format-Specific Errors
- First Energy with no interval data → Skips customer, logs message
- CSV with variable columns → Pre-scans for max column count
- Encoding issues → Tries UTF-8, Latin-1, then default

### Data-Level Errors
- Invalid datetime values → Skipped with `errors='coerce'`
- Missing/NaN values → Excluded from processing
- Duplicate timestamps → First occurrence kept

---

## Performance Considerations

### Memory Usage
- Large files are processed in-memory using pandas DataFrames
- First Energy CSV pre-scan reads file twice (once for column count, once for data)

### Processing Time
Typical processing times (approximate):
- Small file (<1 year): < 1 second
- Medium file (1-2 years): 1-3 seconds
- Large file (3+ years): 3-10 seconds

### Limitations
- Maximum file size limited by available RAM
- Excel files with 1M+ rows may be slow to process

---

## Terminal Interface

### Color Scheme
- **Default (White):** All standard output
- **Green:** Success messages ("✓ Formatting complete!")

### ASCII Header
```
============================================================

    █ █▀▀▀▄ █▀▀▀▄   █▀▀▀ █▀▀▀█ █▀▀▀▄ █   █ █▀▀▀█ ▀▀█▀▀ ▀▀█▀▀ █▀▀▀ █▀▀▀▄
    █ █   █ █▄▄▄▀   █▀▀▀ █   █ █▄▄▄▀ █ █ █ █▀▀▀█   █     █   █▀▀▀ █▄▄▄▀
    █ █▄▄▄▀ █   █   █    █▄▄▄█ █   █ █   █ █   █   █     █   █▄▄▄ █   █  v1.0.5

    [ A P   G A S   &   E L E C T R I C ]
============================================================
```

### Input Methods
1. **Drag-and-drop:** File path passed as `sys.argv[1]`
2. **Manual entry:** Interactive prompt with PowerShell path cleanup

### PowerShell Path Handling
Strips PowerShell's drag-and-drop wrapper:
```python
# Input: & 'C:\path\to\file.csv'
# Cleaned: C:\path\to\file.csv
if user_input.startswith("& "):
    user_input = user_input[2:].strip()
```

---

## Testing Checklist

### Format Testing
- [ ] PSEG CSV file
- [ ] PSEG Excel file
- [ ] ESG Excel file (with IDR Quantity sheet)
- [ ] ESG Excel file with multiple Measurement Units (K1, K3, KH)
- [ ] ESG CSV file (simple format)
- [ ] ESG CSV file (multi-section format with Document Info, Transaction Info, etc.)
- [ ] BGE 15-min CSV file
- [ ] BGE 15-min Excel file
- [ ] BGE Hourly CSV file
- [ ] BGE Hourly Excel file
- [ ] First Energy Excel file (multiple customers)
- [ ] First Energy CSV file (multiple customers)
- [ ] First Energy file with 2359 column (last interval of day)
- [ ] COMED CSV file (multiple meters)
- [ ] COMED Excel file

### Interval Testing
- [ ] 15-minute interval data
- [ ] 30-minute interval data
- [ ] Hourly interval data

### DST Testing
- [ ] March DST gap (spring forward)
- [ ] November DST (fall back) - should be ignored

### Edge Cases
- [ ] File with only 1 customer (First Energy)
- [ ] File with customers having no interval data
- [ ] File with partial year of data
- [ ] File with 3+ years of data
- [ ] Empty/corrupted file handling

---

## Version History

### v1.0.1 (Current)
- **First Energy 2359 fix:** Last interval column (2359) now correctly treated as midnight instead of being skipped
- **ESG multi-section CSV support:** Properly handles ESG CSV files with multiple sections (Document Info, Transaction Info, Organizations Info, etc.)
- **ESG Measurement Unit filtering:** Filters for 'KH' (kWh) data when multiple data sets exist (K1, K3, KH, etc.)
- Improved error handling for non-data rows in ESG files

### v1.0
- Initial release
- Support for 6 utility formats: PSEG, ESG, BGE 15-min, BGE Hourly, First Energy, COMED
- CSV and Excel input support
- Automatic interval detection (15/30/60 min)
- DST handling (March gap fill, November ignore)
- Multi-customer support (First Energy)
- Multi-meter aggregation (COMED)
- Colored terminal output
- ASCII art branding

---

## Future Enhancements (Potential)

- GUI interface using tkinter
- Batch processing of multiple files
- Configuration file for custom settings
- Additional utility format support
- Data validation reporting
- Export to additional formats (JSON, database)

---

*IDR Formatter v1.0.1 - Technical Documentation*  
*AP Gas & Electric Texas*