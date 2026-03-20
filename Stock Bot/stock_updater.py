import gspread
import time
import requests

# --- CONFIGURATION ---
# The name of your credentials file
CREDENTIALS_FILE = "credentials.json" 

# Paste the full URL of your Google Sheet here
SHEET_URL = "https://docs.google.com/spreadsheets/d/1n0PkOHpwjyXObyJVwNDJUPD4pSut4JVHErfywnHZa5c/edit?gid=2110150956#gid=2110150956" 
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
    name_lower = name.lower().strip()
    try:
        # 1. Try exact match first (highest priority)
        for i, h in enumerate(headers):
            if h.lower().strip() == name_lower:
                return i + 1
        # 2. Fallback to "contains" if no exact match (for columns like 'Yield (%)')
        for i, h in enumerate(headers):
            if name_lower in h.lower():
                return i + 1
        return None
    except Exception:
        return None

col_map = {
    "yahoo_ticker": find_col("Yahoo Ticker"),
    "price": find_col("Price"),
    "mcap": find_col("Current MCap"),
    "pe": find_col("P/E"),
    "avg_pe": find_col("5Y Avg PE"),
    "pe_dist": find_col("PE Distance %"),
    "prev_eps": find_col("Previous EPS"),
    "curr_eps": find_col("Current EPS"),
    "yield": find_col("Yield (%)"),
    "growth": find_col("Growth Rate"),
    "peg": find_col("PEG Ratio"),
    "dividend": find_col("Dividend %"),
    "roe": find_col("ROE"),
    "graham": find_col("Graham"),
    "relative_pe": find_col("Relative PE Val"), # Fixed typo from 'Relatuve' to 'Relative'
    "dcf": find_col("DCF"),
    "net_profit": find_col("Net Profit"),
    "net_margin": find_col("Net Profit%"),
    "fcf_net_income": find_col("FCF to NetIncome"),
    "fcf_margin": find_col("FCF Margin"),
    "fcf_yield": find_col("FCF Yield"),
    "fcf_debt": find_col("FCF to Debt")
}

# Helper to convert a column index (e.g., 5) to a letter (e.g., "E")
def col_letter(idx):
    if idx is None: return ""
    result = ""
    while idx > 0:
        idx, remainder = divmod(idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

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

        # --- NEW FCF METRICS ---
        raw_fcf = get_raw("financialData", "freeCashflow")
        raw_net_income = get_raw("defaultKeyStatistics", "netIncomeToCommon")
        raw_revenue = get_raw("financialData", "totalRevenue")
        raw_debt = get_raw("financialData", "totalDebt")

        fcf_net_income = ""
        if raw_fcf != "" and raw_net_income != "" and float(raw_net_income) != 0:
            fcf_net_income = f"{float(raw_fcf) / float(raw_net_income):.2f}"

        fcf_margin = ""
        if raw_fcf != "" and raw_revenue != "" and float(raw_revenue) != 0:
            fcf_margin = f"{(float(raw_fcf) / float(raw_revenue)) * 100:.2f}%"

        fcf_yield = ""
        if raw_fcf != "" and raw_mcap != "" and float(raw_mcap) != 0:
            fcf_yield = f"{(float(raw_fcf) / float(raw_mcap)) * 100:.2f}%"

        fcf_debt = ""
        if raw_fcf != "" and raw_debt != "" and float(raw_debt) != 0:
            fcf_debt = f"{float(raw_fcf) / float(raw_debt):.2f}"

        net_margin = ""
        if raw_net_income != "" and raw_revenue != "" and float(raw_revenue) != 0:
            net_margin = f"{(float(raw_net_income) / float(raw_revenue)) * 100:.2f}%"

        net_profit_formatted = format_large_number(raw_net_income) if raw_net_income != "" else ""

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
        add_cell("net_profit", net_profit_formatted)
        add_cell("net_margin", net_margin)
        add_cell("fcf_net_income", fcf_net_income)
        add_cell("fcf_margin", fcf_margin)
        add_cell("fcf_yield", fcf_yield)
        add_cell("fcf_debt", fcf_debt)
        
        # PE Distance Formula
        if col_map["pe"] and col_map["avg_pe"] and col_map["pe_dist"]:
            pe_c = col_letter(col_map["pe"])
            avg_c = col_letter(col_map["avg_pe"])
            formula = f"=IF(AND({pe_c}{row_num}<>\"\", {avg_c}{row_num}<>\"\"), ({pe_c}{row_num}-{avg_c}{row_num})/{avg_c}{row_num}, \"\")"
            cells_to_update.append(gspread.Cell(row_num, col_map["pe_dist"], formula))
        
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
    # 0. Clear all existing conditional formatting to prevent rules from piling up
    # We first fetch the current rules to know how many to delete
    full_sheet_data = sheet.spreadsheet.fetch_sheet_metadata()
    current_sheet_meta = next(s for s in full_sheet_data['sheets'] if s['properties']['sheetId'] == sheet.id)
    num_rules = len(current_sheet_meta.get('conditionalFormats', []))
    
    # Start with deletion requests for every existing rule
    requests = [{"deleteConditionalFormatRule": {"index": 0, "sheetId": sheet.id}} for _ in range(num_rules)]
    
    # Find the start and end of our data columns for formatting
    start_col = min([c for c in col_map.values() if c is not None and c > 2])
    end_col = max([c for c in col_map.values() if c is not None])
    
    # 1. Freeze first column (A) and header row (1)
    requests.append({
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
    })
    
    # 2. Right alignment for data columns
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": start_col-1, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "RIGHT"}},
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    })

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

    # 5. Net Profit Margin Column (> 50% Bold Green)
    if col_map["net_margin"]:
        m_idx = col_map["net_margin"] - 1
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": m_idx, "endColumnIndex": m_idx + 1}],
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0.5"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })

    # 6. P/E Comparison (4-Color Heat Map)
    if col_map["pe"] and col_map["avg_pe"]:
        pe_idx = col_map["pe"] - 1
        avg_pe_idx = col_map["avg_pe"] - 1
        pe_col = col_letter(col_map["pe"])
        avg_pe_col = col_letter(col_map["avg_pe"])
        
        # Ensure 5Y Avg PE is formatted as a number (not %)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": avg_pe_idx, "endColumnIndex": avg_pe_idx + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.00"}}},
                "fields": "userEnteredFormat.numberFormat"
            }
        })

        pe_range = [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": pe_idx, "endColumnIndex": pe_idx + 1}]
        
        # 1. Light Green (Current < Average)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": pe_range,
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=AND({avg_pe_col}2 <> \"\", {pe_col}2 < {avg_pe_col}2)"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                    }
                }, "index": 0
            }
        })
        
        # 2. Light Red (Current > Average)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": pe_range,
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=AND({avg_pe_col}2 <> \"\", {pe_col}2 > {avg_pe_col}2)"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0.9, "green": 0.4, "blue": 0.4}}}
                    }
                }, "index": 0
            }
        })
        
        # 3. Dark Green (Deep Value: Current <= 80% of Average)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": pe_range,
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=AND({avg_pe_col}2 <> \"\", {pe_col}2 <= ({avg_pe_col}2 * 0.8))"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.4, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })
        
        # 4. Dark Red (Overheated: Current >= 120% of Average)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": pe_range,
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f"=AND({avg_pe_col}2 <> \"\", {pe_col}2 >= ({avg_pe_col}2 * 1.2))"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0.7, "green": 0, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })

    # 7. PE Distance Column (4-Color Heat Map + Percentage Format)
    if col_map["pe_dist"]:
        d_idx = col_map["pe_dist"] - 1
        d_range = [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": d_idx, "endColumnIndex": d_idx + 1}]
        
        # Number Format (Decimal)
        requests.append({
            "repeatCell": {
                "range": d_range[0],
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.00"}}},
                "fields": "userEnteredFormat.numberFormat"
            }
        })
        
        # 1. Light Green (< 0)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": d_range,
                    "booleanRule": {
                        "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.6, "blue": 0}}}
                    }
                }, "index": 0
            }
        })
        # 2. Light Red (> 0)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": d_range,
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0.9, "green": 0.4, "blue": 0.4}}}
                    }
                }, "index": 0
            }
        })
        # 3. Dark Green (<= -20%)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": d_range,
                    "booleanRule": {
                        "condition": {"type": "NUMBER_LESS_THAN_EQ", "values": [{"userEnteredValue": "-0.2"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0, "green": 0.4, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })
        # 4. Dark Red (>= 20%)
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": d_range,
                    "booleanRule": {
                        "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "0.2"}]},
                        "format": {"textFormat": {"foregroundColor": {"red": 0.7, "green": 0, "blue": 0}, "bold": True}}
                    }
                }, "index": 0
            }
        })

    # 8. Global Red Text for negative numbers (at index 0 so it's priority #1)
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

    # 9. Yellow background for positive numbers >= 40
    requests.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": start_col-1, "endColumnIndex": end_col}],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER_THAN_EQ", "values": [{"userEnteredValue": "40"}]},
                    "format": {"backgroundColor": {"red": 1, "green": 1, "blue": 0}}
                }
            },
            "index": 0
        }
    })

    # 10. Light purple background for negative numbers <= -40
    requests.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": start_col-1, "endColumnIndex": end_col}],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS_THAN_EQ", "values": [{"userEnteredValue": "-40"}]},
                    "format": {"backgroundColor": {"red": 0.9, "green": 0.7, "blue": 1}}
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