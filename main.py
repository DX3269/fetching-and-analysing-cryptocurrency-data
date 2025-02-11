import requests
import xlwings as xw
from datetime import datetime
import time
import os


# Fetching cryptocurrency data from CoinGecko
def fetch_cryptocurrency_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print("Error fetching data:", e)
        return []


# Creating and updating Excel file with the help of xlwings
def update_excel(data, workbook_path):
    try:

        if os.path.exists(workbook_path):
            wb = xw.Book(workbook_path)
        else:
            wb = xw.Book()
            wb.save(workbook_path)

        sheet = wb.sheets.active
        sheet.clear()

        headers = ["Rank", "Name", "Symbol", "Price (USD)",
                   "Market Cap (USD)", "24h Volume (USD)",
                   "24h Change (%)", "Timestamp"]
        sheet.range("A1").value = headers

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for idx, coin in enumerate(data, start=2):
            row = [
                idx - 1,
                coin.get("name"),
                coin.get("symbol").upper(),
                coin.get("current_price"),
                coin.get("market_cap"),
                coin.get("total_volume"),
                coin.get("price_change_percentage_24h"),
                current_time
            ]
            sheet.range(f"A{idx}").value = row

        perform_analysis(sheet, data, len(data) + 4)

        sheet.autofit()

        print(f"Excel sheet updated successfully at {current_time}")

    except Exception as e:
        print("Error updating Excel:", e)


# Performing analysis which will be displayed at the end of Excel sheet
def perform_analysis(sheet, data, start_row):

    price_list = [coin.get("current_price", 0) for coin in data]
    market_cap_list = [(coin["name"], coin["market_cap"]) for coin in data]
    price_change_list = [(coin["name"], coin.get("price_change_percentage_24h", 0)) for coin in data]

    top_5_by_market_cap = sorted(market_cap_list, key=lambda x: x[1], reverse=True)[:5]
    top_5_text = ", ".join([f"{name} (${market_cap:,})" for name, market_cap in top_5_by_market_cap])

    average_price = sum(price_list) / len(price_list) if price_list else 0

    highest_change = max(price_change_list, key=lambda x: x[1])
    lowest_change = min(price_change_list, key=lambda x: x[1])

    analysis_headers = ["Analysis", "Result"]
    analysis_data = [
        ["Top 5 by Market Cap", top_5_text],
        ["Average Price (Top 50)", f"${average_price:,.2f}"],
        ["Highest 24h Price Change", f"{highest_change[0]}: {highest_change[1]:.2f}%"],
        ["Lowest 24h Price Change", f"{lowest_change[0]}: {lowest_change[1]:.2f}%"]
    ]

    sheet.range(f"A{start_row}").value = analysis_headers
    sheet.range(f"A{start_row + 1}").value = analysis_data


def job():
    print("Fetching data and updating Excel...")
    data = fetch_cryptocurrency_data()
    if data:
        update_excel(data, "live_crypto_data.xlsx")


print("Updating Excel every 5 minutes... (Press Ctrl+C to stop)")

job()

while True:
    time.sleep(300)   # For every 5 minutes data will be refreshed and updated
    job()

