# IDR Formatter User Guide

## What is the IDR Formatter?

The IDR Formatter is a tool that converts energy usage data files into a standardized format. It takes interval data (energy readings taken every 15, 30, or 60 minutes) and converts it into hourly data that's organized and easy to work with.

---

## How It Works - Flowchart

```
                                ┌─────────────────────┐
                                │   📁 RECEIVE FILE   │
                                └──────────┬──────────┘
                                           │
                                           ▼
                                ┌─────────────────────┐
                                │  What file type?    │
                                └──────────┬──────────┘
                                           │
                      ┌────────────────────┼────────────────────┐
                      │                    │                    │
                      ▼                    ▼                    ▼
               ┌───────────┐        ┌───────────┐        ┌───────────┐
               │   .csv    │        │   .xlsx   │        │   Other   │
               └─────┬─────┘        └─────┬─────┘        └─────┬─────┘
                     │                    │                    │
                     └─────────┬──────────┘                    ▼
                               │                        ┌─────────────┐
                               ▼                        │ ❌ STOP     │
                    ┌─────────────────────┐             │ Unsupported │
                    │   DETECT FORMAT     │             └─────────────┘
                    └──────────┬──────────┘
                               │
         ┌──────┬──────┬─────┼─────┬──────┬──────┐
         │      │      │     │     │      │      │
         ▼      ▼      ▼     ▼     ▼      ▼      ▼
    ┌───────┐┌─────┐┌─────┐┌───┐┌─────┐┌─────┐┌─────┐
    │ First ││COMED││ DUQ ││ESG││ ESG ││ BGE ││PSEG │
    │Energy ││     ││     ││MM ││     ││     ││     │
    └───┬───┘└──┬──┘└──┬──┘└─┬─┘└──┬──┘└──┬──┘└──┬──┘
        │       │      │     │     │      │      │
        │       ▼      │     ▼     │      │      │
        │ ┌──────────┐ │┌────────┐ │      │      │
        │ │ Combine  │ ││Combine │ │      │      │
        │ │ multiple │ ││multiple│ │      │      │
        │ │ meters   │ ││meters  │ │      │      │
        │ └────┬─────┘ │└───┬────┘ │      │      │
        │      │       │    │      │      │      │
         ▼           └─────────┴─────────┴───────────┘
    ┌─────────────┐                    │
    │ Multiple    │                    │
    │ customers?  │                    │
    └──────┬──────┘                    │
           │                           │
     ┌─────┴─────┐                     │
     │           │                     │
     ▼           ▼                     │
   ┌───┐      ┌───┐                    │
   │Yes│      │No │                    │
   └─┬─┘      └─┬─┘                    │
     │          │                      │
     ▼          │                      │
┌──────────┐    │                      │
│ Process  │    │                      │
│ each one │    │                      │
│ separate │    │                      │
└────┬─────┘    │                      │
     │          │                      │
     └─────┬────┘                      │
           │                           │
           └───────────────────────────┘
                               │
                               ▼
                    ┌─────────────────────┐
                    │ 📊 READ INTERVAL    │
                    │      DATA           │
                    └──────────┬──────────┘
                               │
                               ▼
                    ┌─────────────────────┐
                    │ What interval size? │
                    └──────────┬──────────┘
                               │
              ┌────────────────┼────────────────┐
              │                │                │
              ▼                ▼                ▼
        ┌──────────┐    ┌──────────┐    ┌──────────┐
        │ 15-min   │    │ 30-min   │    │ Hourly   │
        │ 96/day   │    │ 48/day   │    │ 24/day   │
        └────┬─────┘    └────┬─────┘    └────┬─────┘
             │               │               │
             └───────────────┼───────────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │ CONVERT TO HOURLY   │
                  └──────────┬──────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │ Check for DST gaps? │
                  └──────────┬──────────┘
                             │
            ┌────────────────┼────────────────┐
            │                │                │
            ▼                ▼                ▼
     ┌────────────┐   ┌────────────┐   ┌────────────┐
     │ March gap  │   │ November   │   │  No DST    │
     │ (missing   │   │ (extra     │   │  issue     │
     │  hour)     │   │  hour)     │   │            │
     └─────┬──────┘   └─────┬──────┘   └─────┬──────┘
           │                │                │
           ▼                ▼                │
     ┌───────────┐   ┌───────────┐          │
     │ Fill with │   │  Ignore   │          │
     │  average  │   │  extra    │          │
     └─────┬─────┘   └─────┬─────┘          │
           │               │                │
           └───────────────┴────────────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │ Trim to last        │
                  │ midnight            │
                  └──────────┬──────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │ ORGANIZE BY YEAR    │
                  │                     │
                  │ Year 1 = Most recent│
                  │ Year 2 = Previous   │
                  │ Year 3, 4... etc    │
                  └──────────┬──────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │ 💾 SAVE FORMATTED   │
                  │       FILE          │
                  └──────────┬──────────┘
                             │
                             ▼
                  ┌─────────────────────┐
                  │     ✅ DONE!        │
                  └─────────────────────┘
```

---

## What Does It Do?

When you receive energy usage data from different utilities, each one sends it in their own format. This tool:

1. **Reads** your energy data file (CSV or Excel)
2. **Recognizes** which utility format it's in
3. **Converts** the data to hourly readings
4. **Organizes** the data by year
5. **Saves** a new formatted file

---

## Supported Utility Formats

The formatter automatically detects and processes files from these sources:

| Utility | What to Look For |
|---------|------------------|
| **PSEG** | Simple 2-column files with date/time and usage |
| **ESG** | Files with "IDR Quantity" sheet or "Interval Ending" columns |
| **ESG Multi-Meter** | ESG files with multiple meters (auto-summed together) |
| **BGE** | Files with columns like "RdgDate" or "ReadDate" and "Kwh" |
| **First Energy** | Files with "Customer Identifier" sections (can have multiple customers) |
| **COMED** | Files with "INTERVAL USAGE DATA" header and multiple meters |
| **DUQ** | Duquesne Light files with "Customer Identity" header and hourly interval data |

---

## How to Use the Formatter

### Method 1: Drag and Drop
1. Find your energy data file
2. Drag the file onto the IDR Formatter icon
3. The formatter will process it automatically
4. Press Enter when finished

### Method 2: Enter File Path
1. Double-click the IDR Formatter to open it
2. Copy the file path of your energy data file
3. Paste it into the formatter and press Enter
4. The formatter will process it
5. Type 'quit' when you're done, or process another file

---

## Where Does the Formatted File Go?

The formatted file is saved in the **same folder** as your original file, with `_formatted` added to the name.

**Examples:**
- Original: `EnergyData.csv` → Formatted: `EnergyData_formatted.csv`
- Original: `CustomerReport.xlsx` → Formatted: `CustomerReport_formatted.xlsx`

**Special Case - First Energy:**
If your file has multiple customers, the output will be an Excel file with a separate tab for each customer, named by their customer ID number.

---

## Understanding the Output

The formatted file contains your energy data organized into columns:

| Column | Description |
|--------|-------------|
| **Intv End Date/Time** | The complete dataset, oldest to newest |
| **Usage** | Energy usage in kWh for each hour |
| **YEAR 1** | The most recent 12 months of data |
| **YEAR 2** | The previous 12 months (if available) |
| **YEAR 3, 4...** | Additional years as needed |

Each "YEAR" section contains up to 8,760 hours (one full year).

---

## Daylight Saving Time (DST)

The formatter automatically handles Daylight Saving Time:

- **Spring Forward (March):** When clocks skip an hour, the formatter fills in the missing hour using an average of the surrounding values
- **Fall Back (November):** The extra hour is ignored to keep all days at 24 hours

This ensures your data always has consistent 24-hour days.

### DUQ Partial Day Filling

DUQ (Duquesne Light) files sometimes have days where data stops mid-day due to VEE data issues. The formatter automatically detects these partial days and fills the missing hours using data from the same day of the week, one week prior. This keeps every day at a uniform 24 hours.

---

## Troubleshooting

### "File not found"
- Make sure the file path is correct
- Check that the file isn't open in another program
- Try dragging and dropping the file instead

### "Unsupported file type"
- The formatter only works with `.csv`, `.xlsx`, and `.xls` files
- If your file is a different format, try saving it as one of these types first

### "No interval data found"
- The file may not contain actual usage data
- Some customer sections in multi-customer files may not have interval data (this is normal - those customers are skipped)

### File processes but looks wrong
- Check that you're using the correct utility's file format
- Make sure the original file hasn't been modified or corrupted

---

## Quick Reference

**Supported File Types:**
- CSV (.csv)
- Excel (.xlsx, .xls)

**Input Intervals:**
- 15-minute data ✓
- 30-minute data ✓
- Hourly data ✓

**Output Format:**
- Always hourly
- Organized by year
- Dates formatted as MM/DD/YYYY HH:MM

---

## Need Help?

If you encounter issues not covered in this guide, check that:
1. Your file is one of the supported formats
2. The file isn't corrupted or incomplete
3. The file contains actual interval usage data

For additional assistance, contact your system administrator or the tool developer.

---

*IDR Formatter v1.1.0 - AP Gas & Electric*
