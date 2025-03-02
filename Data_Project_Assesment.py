import requests
import pandas as pd
import time
import openpyxl

def fetch_crypto_data():
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
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return []

def process_data(data):
    if not data:
        return None, None, None, None, None
    
    df = pd.DataFrame(data, columns=["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"])
    df.dropna(inplace=True)
    df.sort_values(by="market_cap", ascending=False, inplace=True)
    
    if df.empty:
        return None, None, None, None, None
    
    top_5 = df.head(5)
    avg_price = df["current_price"].mean()
    highest_change = df.loc[df["price_change_percentage_24h"].idxmax()]
    lowest_change = df.loc[df["price_change_percentage_24h"].idxmin()]
    
    return df, top_5, avg_price, highest_change, lowest_change

def update_excel(df, top_5, avg_price, highest_change, lowest_change):
    if df is None:
        print("No data to update in Excel.")
        return
    
    file_name = "Crypto_Live_Data.xlsx"
    with pd.ExcelWriter(file_name, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Live Data", index=False)
        top_5.to_excel(writer, sheet_name="Top 5", index=False)
        
        summary = pd.DataFrame({
            "Metric": ["Average Price", "Highest Change (24h)", "Lowest Change (24h)"],
            "Value": [avg_price, f"{highest_change['name']} ({highest_change['price_change_percentage_24h']}%)", 
                      f"{lowest_change['name']} ({lowest_change['price_change_percentage_24h']}%)"]
        })
        summary.to_excel(writer, sheet_name="Summary", index=False)
    print("Excel updated successfully")

def live_update():
    while True:
        print("Fetching latest cryptocurrency data...")
        data = fetch_crypto_data()
        df, top_5, avg_price, highest_change, lowest_change = process_data(data)
        update_excel(df, top_5, avg_price, highest_change, lowest_change)
        
        print("Waiting for 5 minutes before next update...")
        time.sleep(300)  # Update every 5 minutes

if __name__ == "__main__":
    live_update()
