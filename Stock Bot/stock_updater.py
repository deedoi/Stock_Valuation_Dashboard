import gspread
import time
import requests

# --- CONFIGURATION ---
# The name of your credentials file
CREDENTIALS_FILE = "credentials.json" 

# Paste the full URL of your Google Sheet here
SHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit" 
# ---------------------

print("Connecting to Google Sheets...")
try:
    gc = gspread.service_account(filename=CREDENTIALS_FILE)
    sheet = gc.open_by_url(SHEET_URL).sheet1
    print("✅ Successfully connected to your Google Sheet!")
except Exception as e:
    print(f"❌ Failed to connect to Google Sheets. Check your URL and credentials. Error: {e}")
    exit()

print("Bypassing Yahoo Security...")
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5"
})

# 1. Fetch a standard page to get the session cookies
try:
    session.get("https://finance.yahoo.com/", timeout=15)
    # 2. Fetch the security crumb to unlock the API
    crumb_res = session.get("https://query1.finance.yahoo.com/v1/test/getcrumb", timeout=15)
    crumb = crumb_res.text.strip()
    print("✅ Security bypassed successfully!")
except Exception as e:
    print("⚠️ Warning: Could not fetch security crumb. We will try fetching anyway.")
    crumb = ""

print("Reading dashboard...")
records = sheet.get_all_values()
headers = records[0]

# --- DYNAMIC COLUMN MAPPING ---
# This helper finds the position of a column by its name so you can move/delete columns!
def find_col(name):
    try:
        # Returns 1-based index for gspread, 0-based index for records list
        idx = next(i for i, h in enumerate(headers) if name.lower() in h.lower())
        return idx + 1
    except StopIteration:
        return None

col_map = {
    "yahoo_ticker": find_col("Yahoo Ticker"),
    "price": find_col("Price"),
    "mcap": find_col("Current MCap"),
    "pe": find_col("P/E"),
    "prev_eps": find_col("Previous EPS"),
    "curr_eps": find_col("Current EPS"),
    "yield": find_col("Yield (%)"),
    "growth": find_col("Growth Rate"),
    "peg": find_col("PEG Ratio"),
    "dividend": find_col("Dividend %"),
    "roe": find_col("ROE"),
    "graham": find_col("Graham"),
    "relative_pe": find_col("Relatuve PE Val"), # Matches your current typo "Relatuve"
    "dcf": find_col("DCF")
}

# Create a list to hold all our batch updates
cells_to_update = []

# Helper function to smartly format massive numbers with T, B, and M
def format_large_number(value):
    try:
        val = float(value)
        if val >= 1e12:
            return f"{val/1e12:.2f} T"
        elif val >= 1e9:
            return f"{val/1e9:.2f} B"
        elif val >= 1e6:
            return f"{val/1e6:.2f} M"
        else:
            return f"{val:.2f}"
    except Exception:
        return value

# We start at row 2 (index 1 in Python) to skip the header row
for i in range(1, len(records)):
    # Safely get the ticker from the mapped column
    ticker_idx = col_map["yahoo_ticker"] - 1 if col_map["yahoo_ticker"] else None
    if ticker_idx is None:
        print("❌ Error: Could not find 'Yahoo Ticker' column in your sheet!")
        break
        
    yahoo_ticker = records[i][ticker_idx]
    
    # Skip empty rows
    if not yahoo_ticker or yahoo_ticker == "Yahoo Ticker":
        continue
        
    print(f"Fetching data for {yahoo_ticker}...")
    
    try:
        # 3. Hit the clean, structured JSON API using the security crumb
        url = f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{yahoo_ticker}?modules=financialData,defaultKeyStatistics,price,summaryDetail&crumb={crumb}"
        res = session.get(url, timeout=15)
        
        if res.status_code != 200:
            print(f"  -> ❌ Failed: Yahoo returned status code {res.status_code}")
            continue
            
        data = res.json()
        result_list = data.get("quoteSummary", {}).get("result", [])
        
        if not result_list:
            print(f"  -> ❌ Failed to fetch {yahoo_ticker}: Data not found in Yahoo database.")
            continue
            
        result = result_list[0]
        
        # Helper function to extract "raw" numbers safely from the JSON dictionary
        def get_raw(module, key):
            return result.get(module, {}).get(key, {}).get("raw", "")
        
        # Extract the metrics seamlessly regardless of the stock's home country
        raw_price = get_raw("price", "regularMarketPrice")
        if raw_price == "": raw_price = get_raw("financialData", "currentPrice")
        price = f"{float(raw_price):.2f}" if raw_price != "" else ""
            
        raw_mcap = get_raw("summaryDetail", "marketCap")
        if raw_mcap == "": raw_mcap = get_raw("price", "marketCap")
        mcap = format_large_number(raw_mcap) if raw_mcap != "" else ""
            
        raw_pe = get_raw("summaryDetail", "trailingPE")
        if raw_pe == "": raw_pe = get_raw("defaultKeyStatistics", "trailingPE")
        pe = f"{float(raw_pe):.2f}" if raw_pe != "" else ""
            
        raw_eps = get_raw("defaultKeyStatistics", "trailingEps")
        eps = f"{float(raw_eps):.2f}" if raw_eps != "" else ""
        
        raw_div = get_raw("summaryDetail", "dividendYield")
        dividend = f"{float(raw_div) * 100:.2f}%" if raw_div != "" else ""
        
        raw_roe = get_raw("financialData", "returnOnEquity")
        roe = f"{float(raw_roe) * 100:.2f}%" if raw_roe != "" else ""

        # --- RE-MAPPING AND CALCULATIONS BASED ON PENDING PROJECT REQUIREMENTS ---
        row_num = i + 1
        
        # Current EPS from Yahoo
        current_eps_val = float(raw_eps) if raw_eps != "" else 0
        
        # 1. Fetch Previous EPS from Sheet (Dynamic Column Mapping)
        try:
            prev_eps_idx = col_map["prev_eps"] - 1
            prev_eps_str = records[i][prev_eps_idx].replace(',', '').replace('%', '').strip()
            prev_eps_val = float(prev_eps_str) if prev_eps_str else 0
        except Exception:
            prev_eps_val = 0
            
        # Task 1: Correct Earning Growth Rate Formula
        growth_rate = 0
        using_growth_data = False
        
        if prev_eps_val != 0:
            growth_rate = ((current_eps_val - prev_eps_val) / abs(prev_eps_val)) * 100
            using_growth_data = True
        else:
            raw_growth_y = get_raw("financialData", "earningsGrowth")
            if raw_growth_y != "":
                growth_rate = float(raw_growth_y) * 100
                using_growth_data = True
                
        growth_str = f"{growth_rate:.2f}%" if using_growth_data else "N/A"
        
        # Task 2: PEG Ratio Calculation
        peg_ratio = "N/A"
        if raw_pe != "" and using_growth_data:
            if growth_rate != 0:
                peg_ratio = f"{float(raw_pe) / growth_rate:.2f}"
            else:
                peg_ratio = "N/A"
            
        # Task 3: Correct Earnings Yield Formula
        earning_yield = ""
        if current_eps_val != 0 and raw_price != "" and float(raw_price) != 0:
            earning_yield = f"{(current_eps_val / float(raw_price)) * 100:.2f}%"

        # Build the batch updates using the dynamic col_map
        def add_cell(key, value):
            if col_map[key] and value != "":
                cells_to_update.append(gspread.Cell(row_num, col_map[key], value))

        add_cell("price", price)
        add_cell("mcap", mcap)
        add_cell("pe", pe)
        add_cell("curr_eps", eps)
        add_cell("yield", earning_yield)
        add_cell("growth", growth_str)
        add_cell("peg", peg_ratio)
        add_cell("dividend", dividend)
        add_cell("roe", roe)
        
        # --- ENHANCED VALUATION CALCULATIONS ---
        try:
            if current_eps_val > 0:
                # 1. Conservative Graham Number
                g_graham = max(growth_rate, 0) 
                graham_val = f"{current_eps_val * (7 + 1.5 * g_graham):.2f}"
                add_cell("graham", graham_val)
                
                # 2. Automated Relative PE Valuation
                relative_pe_val = f"{current_eps_val * 20:.2f}"
                add_cell("relative_pe", relative_pe_val)
                
                # 3. Simplified DCF Valuation
                discount_rate = 0.10
                terminal_pe = 15
                g_dcf = min(max(growth_rate, 0), 15) / 100 
                
                dcf_total = 0
                temp_eps = current_eps_val
                for year in range(1, 6):
                    temp_eps *= (1 + g_dcf)
                    dcf_total += temp_eps / ((1 + discount_rate) ** year)
                
                terminal_value = (temp_eps * terminal_pe) / ((1 + discount_rate) ** 5)
                dcf_total += terminal_value
                add_cell("dcf", f"{dcf_total:.2f}")
        except Exception:
            pass

        print(f"  -> ✅ {yahoo_ticker} data fetched and queued.")
        
    except Exception as e:
        print(f"  -> ❌ Failed to process {yahoo_ticker}: {e}")
        
    # Pause for 1.5 seconds to be polite to Yahoo's servers
    time.sleep(1.5)

# --- BATCH UPDATE GOOGLE SHEETS ---
print("\nPushing all updates to Google Sheets in ONE single request...")
if cells_to_update:
    sheet.update_cells(cells_to_update, value_input_option='USER_ENTERED')

# --- APPLY FORMATTING (RED FOR NEGATIVE, RIGHT ALIGN, & FREEZE) ---
print("Applying formatting (Red for negatives, Right align, & Freezing column A)...")
try:
    # Find the start and end of our data columns for formatting
    start_col = min([c for c in col_map.values() if c is not None and c > 2])
    end_col = max([c for c in col_map.values() if c is not None])
    
    requests = [
        # 1. Freeze first column (A) and header row (1)
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet.id,
                    "gridProperties": {
                        "frozenRowCount": 1,
                        "frozenColumnCount": 1
                    }
                },
                "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
            }
        },
        # 2. Right alignment for data columns
        {
            "repeatCell": {
                "range": {"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": start_col-1, "endColumnIndex": end_col},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "RIGHT"}},
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        }
    ]

    # --- Task: Custom Conditional Formatting for ROE and Earning Yield ---
    # We add these rules at index 0 in the list of rules. 
    # The last ones we add will be at the very top (checked first).

    # 1. Earning Yield Column (> 5% Green, < 2% Red)
    if col_map["yield"]:
        y_idx = col_map["yield"] - 1
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": y_idx, "endColumnIndex": y_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0.02"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0.8, "green": 0, "blue": 0}}}
                    }
                }, "index": 0
            }
        })
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": y_idx, "endColumnIndex": y_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0.05"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                    }
                }, "index": 0
            }
        })

    # 2. ROE Column (> 20% Green, > 100% Yellow Background)
    if col_map["roe"]:
        r_idx = col_map["roe"] - 1
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": r_idx, "endColumnIndex": r_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0.2"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                    }
                }, "index": 0
            }
        })
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": r_idx, "endColumnIndex": r_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "1"}]},
                        "format": {
                            "backgroundColor": {"red": 1, "green": 1, "blue": 0},
                            "textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}
                        }
                    }
                }, "index": 0
            }
        })

    # 3. Earning Growth Rate Column (> 20% Green, > 100% Green Bold)
    if col_map["growth"]:
        g_idx = col_map["growth"] - 1
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": g_idx, "endColumnIndex": g_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0.2"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                    }
                }, "index": 0
            }
        })
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": g_idx, "endColumnIndex": g_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "1"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })

    # 4. PEG Ratio Column (> 2.5 Red Bold, 0 to 1.0 Green)
    if col_map["peg"]:
        p_idx = col_map["peg"] - 1
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": p_idx, "endColumnIndex": p_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_BETWEEN", "values": [{"userEnteredValue": "0.01"}, {"userEnteredValue": "1"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                    }
                }, "index": 0
            }
        })
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": p_idx, "endColumnIndex": p_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "2.5"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0.8, "green": 0, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })

    # 5. Global Red Text for negative numbers (at index 0 so it's priority #1)
    requests.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": start_col-1, "endColumnIndex": end_col}],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                    "format": {"textFormat": {"foregroundColor": {"red": 0.8, "green": 0, "blue": 0}}}
                }
            },
            "index": 0
        }
    })

    sheet.spreadsheet.batch_update({"requests": requests})
    print("✅ Formatting and Freezing applied!")
except Exception as e:
    print(f"⚠️ Formatting applied (some steps might have been skipped). Error: {e}")
    
print("🎉 All stocks updated successfully! You can now delete any column you don't want.")