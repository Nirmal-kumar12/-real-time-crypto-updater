# -real-time-crypto-updater

This project fetches real-time cryptocurrency data for the top 50 cryptocurrencies using the CoinGecko API and updates an Excel sheet every 5 minutes. It performs basic analysis, including:

> Identifying the Top 5 cryptocurrencies by market cap
> Calculating the average price of the top 50 cryptocurrencies
> Finding the highest and lowest 24-hour price change percentages

Tech Stack:

> Python (requests, pandas, openpyxl)
> CoinGecko API (for live crypto data)
> Excel Automation (to store and update data)

📌 Features:

> Fetches live cryptocurrency data every 5 minutes
> Stores data in an auto-updating Excel sheet
> Provides key insights like top movers and average prices
> Handles API errors and missing data gracefully

📂 How to Run:

1. Clone the repository
2. Install dependencies: pip install requests pandas openpyxl
3. Run the script: python script.py
   
🔗 Future Enhancements:

> Adding data visualization (charts/graphs)
> Creating a web dashboard for real-time monitoring
