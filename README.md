# Stock Valuation Dashboard

Automate your fundamental analysis with this Python-powered Google Sheets dashboard. Fetch real-time data from Yahoo Finance for both Thai (SET) and International stocks, calculate advanced valuation metrics (DCF, Graham Number, PEG), and keep your data organized with automatic formatting.

## Features

- **Real-time Data:** Automatically fetches Price, Market Cap, P/E, EPS, Dividend Yield, and ROE.
- **Advanced Valuations:** 
  - **PEG Ratio:** Price/Earnings to Growth.
  - **Graham Number:** Conservative "fair price" based on Benjamin Graham's formula.
  - **Relative P/E Valuation:** Quick fair value estimate based on industry standard multiples.
  - **Simplified DCF:** 5-year Discounted Cash Flow projection.
- **Smart Logic for Thai Stocks:** Includes fallbacks for missing historical data on SET stocks by pulling growth rates directly from Yahoo Finance.
- **Dynamic Formatting:** 
  - Automatically **freezes Column A** (Stock Name) and **Row 1** (Headers).
  - Highlights negative numbers in **Red** text.
  - **Right-aligns** all financial data for professional readability.
- **Flexible Structure:** Move or delete columns as you like—the script finds data by header names, not fixed positions.

## Google Cloud Console Setup

To connect this script to your Google Sheet, you must first set up a Service Account and enable the necessary APIs:

1.  **Create a Project:** Go to the [Google Cloud Console](https://console.cloud.google.com/) and create a new project.
2.  **Enable APIs:** In the "APIs & Services" dashboard, click "Enable APIs and Services" and search for/enable:
    - **Google Sheets API**
    - **Google Drive API**
3.  **Create Service Account:**
    - Go to "IAM & Admin" > "Service Accounts".
    - Click "Create Service Account", give it a name, and click "Create and Continue".
    - You can skip the optional role selection and click "Done".
4.  **Generate Credentials File:**
    - Click on your new Service Account from the list.
    - Go to the **Keys** tab, click "Add Key" > "Create new key".
    - Select **JSON** and click "Create". A file will download to your computer.
    - Rename this file to `credentials.json` and place it in the same folder as your Python script.
5.  **Share your Google Sheet:**
    - Open your `credentials.json` and find the `"client_email"` address.
    - Open your Google Sheet, click the "Share" button, and paste that email address to give it "Editor" access.

## Getting Started

### 1. Prerequisites
- Python 3.x installed.

### 2. Setup
1.  **Create your Google Sheet:**
    - Go to [Google Sheets](https://sheets.new) and create a new blank spreadsheet.
    - Click **File > Import > Upload** and upload the `Final Stock Dashboad - Final.csv` file from this repository.
    - Choose **"Replace current sheet"** to set up the dashboard template with the correct headers.
    - Copy the **URL** of your new Google Sheet.
2.  **Configure the Script:**
    - Open `stock_updater.py` in your code editor.
    - Locate the `SHEET_URL` variable and paste your copied URL:
      ```python
      SHEET_URL = "https://docs.google.com/spreadsheets/d/your_id_here/edit"
      ```
3.  **Install Dependencies:**
    - Open your terminal or command prompt and run:
      ```bash
      pip install gspread requests
      ```
4.  **Connect to Google Cloud:**
    - Follow the **Google Cloud Console Setup** guide below to generate your `credentials.json` file.
    - Place the `credentials.json` file in the same folder as the script.
    - **Crucial:** Share your Google Sheet with the Service Account email found in your `credentials.json`.

### 3. Usage
1.  Ensure you have at least one column named **"Yahoo Ticker"** and your stock tickers are filled in (e.g., `AAPL` or `CPALL.BK`).
2.  Run the script:
    ```bash
    python stock_updater.py
    ```

## Dashboard Columns
| Column | Description |
| --- | --- |
| **Yahoo Ticker** | The unique ID for Yahoo Finance (e.g., AAPL, PTT.BK) |
| **Earning Yield %** | (EPS / Price) * 100. Shows the "interest rate" of the company's profit. |
| **Dividend %** | The annual cash payout yield to shareholders. |
| **Growth Rate** | % change in EPS from the previous period. |
| **PEG Ratio** | P/E divided by Growth. A PEG < 1.0 often indicates an undervalued stock. |
| **Relative PE Val** | Fair value estimate using a P/E multiple of 20. |
| **DCF** | Present value of future cash flows (5-year projection). |

## Troubleshooting
- **Missing PEG:** Ensure you have a value in "Previous EPS" or that Yahoo Finance has growth data for that ticker.
- **Unauthorized Error:** Double-check that you have shared your Google Sheet with the Service Account email.

---
*Note: This tool is for educational purposes only. Always perform your own due diligence before making investment decisions.*
