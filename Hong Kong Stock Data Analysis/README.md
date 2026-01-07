# Hong Kong Stock Data Analysis

This project is a VBA-powered Excel tool for analyzing Hong Kong stock data.  
It calculates key **SMA values** (5, 10, 20, 50, 100, 200) and detects **Golden Cross** events.

---

### Features

- Automatically calculates SMA (Simple Moving Average) for each stock
- Detects **Golden Cross** events using SMA-50 and SMA-200
- Dynamic user interaction with buttons:
  - **Choose a Company**
  - **Golden Cross**
  - **Update All SMA & Golden Cross**
  - **Back to Cover Page**
- Well-organized UI with a cover page and main display sheet

---

### File Structure

- `HK_stockdata.xlsm`: Main Excel file with all VBA modules included
- Includes:
  - `Main`, `Cover`, and `Code` sheets
  - Individual stock worksheets (e.g., `00008.HK`, `00003.HK`)

---

### Golden Cross Logic
A **Golden Cross** is detected when:

> **SMA-50 crosses above SMA-200** from below  
> (i.e., it transitions from SMA-50 < SMA-200 → SMA-50 ≥ SMA-200)

---

### How to Use

1. Open `HK_stockdata.xlsm` in Excel
2. Enable Macros when prompted
3. Click **Run** on the Cover page to enter the main interface
4. Use the buttons on the right-hand side to interact

---

