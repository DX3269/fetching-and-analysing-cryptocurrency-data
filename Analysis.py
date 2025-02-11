import requests
from datetime import datetime
from fpdf import FPDF


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


# Analysing the fetched data
def perform_analysis(data):
    insights = {}

    top_5_by_market_cap = sorted(data, key=lambda x: x.get("market_cap", 0), reverse=True)[:5]
    insights["Top 5 by Market Cap"] = [(coin["name"], coin["market_cap"]) for coin in top_5_by_market_cap]

    prices = [coin.get("current_price", 0) for coin in data]
    insights["Average Price"] = sum(prices) / len(prices) if prices else 0

    price_changes = [(coin["name"], coin.get("price_change_percentage_24h", 0)) for coin in data]
    insights["Highest 24h Price Change"] = max(price_changes, key=lambda x: x[1], default=("N/A", 0))
    insights["Lowest 24h Price Change"] = min(price_changes, key=lambda x: x[1], default=("N/A", 0))

    return insights


# Generating a PDF report and it will be saved in the same directory in which file is present
def generate_pdf_report(data, insights, filename="crypto_report.pdf"):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Cryptocurrency Market Analysis Report", ln=True, align="C")
    pdf.ln(10)

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"Report Generated On: {current_time}", ln=True)
    pdf.ln(10)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Top 5 Cryptocurrencies by Market Cap:", ln=True)
    pdf.ln(5)
    pdf.set_font("Arial", "", 12)
    for name, market_cap in insights["Top 5 by Market Cap"]:
        pdf.cell(0, 10, f"{name}: ${market_cap:,.2f}", ln=True)
    pdf.ln(10)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Average Price of Top 50 Cryptocurrencies:", ln=True)
    pdf.ln(5)
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, f"${insights['Average Price']:,.2f}", ln=True)
    pdf.ln(10)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Price Change Analysis (24h):", ln=True)
    pdf.ln(5)
    pdf.set_font("Arial", "", 12)
    highest = insights["Highest 24h Price Change"]
    lowest = insights["Lowest 24h Price Change"]
    pdf.cell(0, 10, f"Highest: {highest[0]} with {highest[1]:.2f}% change", ln=True)
    pdf.cell(0, 10, f"Lowest: {lowest[0]} with {lowest[1]:.2f}% change", ln=True)

    # Saving all the data of the PDF
    pdf.output(filename)
    print(f"Report saved as '{filename}'")



def main():
    data = fetch_cryptocurrency_data()
    if not data:
        print("No data available to generate the report.")
        return

    insights = perform_analysis(data)
    generate_pdf_report(data, insights)


if __name__ == "__main__":
    main()
