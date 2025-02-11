# import requests
# import pandas as pd
# import time
# from openpyxl import load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows

# # API URL for fetching top 50 cryptocurrencies from CoinGecko
# API_URL = "https://api.coingecko.com/api/v3/coins/markets"
# PARAMS = {
#     "vs_currency": "usd",
#     "order": "market_cap_desc",
#     "per_page": 50,
#     "page": 1,
#     "sparkline": "false"
# }

# EXCEL_FILE = "crypto_live_data.xlsx"

# # Function to fetch cryptocurrency data
# def fetch_crypto_data():
#     response = requests.get(API_URL, params=PARAMS)
#     if response.status_code == 200:
#         return response.json()
#     else:
#         print("Error fetching data")
#         return []

# # Function to process and update Excel sheet
# def update_excel():
#     while True:
#         data = fetch_crypto_data()
#         if not data:
#             continue
        
#         # Creating DataFrame
#         df = pd.DataFrame(data, columns=[
#             "name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"
#         ])
        
#         # Sorting top 5 by market cap
#         top_5 = df.nlargest(5, "market_cap")
#         avg_price = df["current_price"].mean()
#         highest_change = df["price_change_percentage_24h"].max()
#         lowest_change = df["price_change_percentage_24h"].min()
        
#         # Save data to Excel
#         with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
#             df.to_excel(writer, sheet_name="Live Crypto Data", index=False)
#             top_5.to_excel(writer, sheet_name="Top 5 Market Cap", index=False)
            
#             summary_data = pd.DataFrame({
#                 "Metric": ["Average Price", "Highest 24h Change", "Lowest 24h Change"],
#                 "Value": [avg_price, highest_change, lowest_change]
#             })
#             summary_data.to_excel(writer, sheet_name="Summary", index=False)
        
#         print("Excel updated with latest data...")
#         time.sleep(300)  # Update every 5 minutes

# # Run the function to start updating the Excel file
# update_excel()





import requests
import pandas as pd
import openpyxl
import time
import schedule

EXCEL_FILE = "Live_Crypto_Data.xlsx"

def fetch_crypto_data():
    """Fetch live cryptocurrency data from CoinGecko API."""
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }

    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print("Error fetching data:", response.status_code)
        return []

def analyze_data(data):
    """Analyze top 50 cryptocurrency data."""
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])

    # Top 5 cryptocurrencies by market cap
    top_5_by_market_cap = df.nlargest(5, "market_cap")

    # Average price of top 50 cryptocurrencies
    avg_price = df["current_price"].mean()

    # Highest and lowest 24-hour percentage price change
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax()]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin()]

    analysis = {
        "Top 5 Cryptos by Market Cap": top_5_by_market_cap[["name", "market_cap"]].to_dict(orient="records"),
        "Average Price of Top 50": avg_price,
        "Highest 24h Change": highest_change[["name", "price_change_percentage_24h"]].to_dict(),
        "Lowest 24h Change": lowest_change[["name", "price_change_percentage_24h"]].to_dict()
    }

    return df, analysis

def update_excel(data):
    """Update the live Excel sheet with new data."""
    df, analysis = analyze_data(data)

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Live Crypto Data", index=False)

        # Writing Analysis Data
        analysis_df = pd.DataFrame({
            "Metric": ["Average Price of Top 50", "Highest 24h Change", "Lowest 24h Change"],
            "Value": [
                f"${analysis['Average Price of Top 50']:.2f}",
                f"{analysis['Highest 24h Change']['name']} ({analysis['Highest 24h Change']['price_change_percentage_24h']:.2f}%)",
                f"{analysis['Lowest 24h Change']['name']} ({analysis['Lowest 24h Change']['price_change_percentage_24h']:.2f}%)"
            ]
        })
        analysis_df.to_excel(writer, sheet_name="Analysis", index=False)

    print("Excel file updated successfully.")

def job():
    """Scheduled job to fetch data and update Excel."""
    print("Fetching live cryptocurrency data...")
    data = fetch_crypto_data()
    if data:
        update_excel(data)

# Schedule the job to run every 5 minutes
schedule.every(5).minutes.do(job)

print("Crypto tracker started. Updating data every 5 minutes...")
job()  # Run once before starting the loop

while True:
    schedule.run_pending()
    time.sleep(1)
