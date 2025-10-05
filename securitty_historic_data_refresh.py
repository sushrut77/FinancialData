import yfinance as yf
import pandas as pd
from datetime import datetime

# NOTE: The 'auto_adjust' utility import is not necessary since 'auto_adjust' is passed
# as an argument to yf.download, so I removed it for cleaner code.

# --- Configuration ---
# Define the default end date as today
END_DATE = datetime.now().strftime('%Y-%m-%d')
# Define the default start date as 5 years ago
START_DATE = (datetime.now() - pd.DateOffset(years=5)).strftime('%Y-%m-%d')
START_DATE = "1990-01-01"


def read_tickers_from_file(filename="tickers.txt"):
    """
    Reads a list of stock ticker symbols from a text file, ignoring empty lines.

    The file should contain one ticker symbol per line.
    Example content of tickers.txt:
    AAPL
    MSFT

    GOOGL
    """
    print(f"Attempting to read tickers from: {filename}")
    try:
        with open(filename, 'r') as f:
            # Read lines, strip leading/trailing whitespace, and filter out any resulting empty strings
            tickers = [line.strip() for line in f if line.strip()]
        return tickers
    except FileNotFoundError:
        print(f"Error: The file '{filename}' was not found in the current directory.")
        return []


def download_historical_prices(ticker_symbol: str, start_date: str = START_DATE, end_date: str = END_DATE,
                               interval: str = "1d"):
    """
    Downloads historical market data for a given security ticker and saves it as
    a new sheet (tab) in the 'historical_prices.xlsx' Excel file.

    If the sheet for the given ticker already exists, it will be replaced.
    Other sheets in the file will remain untouched.

    Args:
        ticker_symbol (str): The ticker symbol of the security (e.g., 'AAPL', 'GOOGL').
        start_date (str): The start date for the data download (format: 'YYYY-MM-DD').
                          Defaults to 5 years ago.
        end_date (str): The end date for the data download (format: 'YYYY-MM-DD').
                        Defaults to today.
        interval (str): The data interval (e.g., '1d' for daily, '1wk' for weekly, '1mo' for monthly).
                        Defaults to '1d'.
    """
    print(f"--- Starting download for {ticker_symbol} ---")
    print(f"Date Range: {start_date} to {end_date}, Interval: {interval}")

    excel_filename = "historical_prices.xlsx"
    sheet_name = ticker_symbol.upper()  # Use uppercase ticker as sheet name
    historic_security_data = None

    try:
        # Download the data using yfinance
        historic_security_data = yf.download(
            ticker_symbol,
            start=start_date,
            end=end_date,
            interval=interval,
            auto_adjust=True  # Automatically adjust prices for splits and dividends
        )

        if historic_security_data.empty:
            print(f"Error: No data found for ticker '{ticker_symbol}' in the specified range.")
            return

        # Use pandas.ExcelWriter to open the file in append mode ('a') and write to a specific sheet.
        # if_sheet_exists='replace' ensures that if you run the same ticker twice,
        # the sheet is updated with the latest data, but other sheets are preserved.
        with pd.ExcelWriter(
                excel_filename,
                mode='a',
                if_sheet_exists='overlay',
                engine='openpyxl'  # Use openpyxl engine, necessary for append mode
        ) as writer:
            # Write the DataFrame to the specific sheet, including the Date Index
            historic_security_data.to_excel(writer, sheet_name=sheet_name, index=True)

        print(f"\nSuccessfully downloaded data and saved to sheet '{sheet_name}' in: {excel_filename}")
        print("-" * 40)
        print(historic_security_data.head())  # Display the first few rows

    except FileNotFoundError:
        # This handles the initial case where the Excel file might not exist yet.
        # In this case, we write the file normally (which creates it), then the 'a' mode
        # will work for subsequent calls.
        print(f"File {excel_filename} not found. Creating a new file...")
        try:
            # We need to re-download the data if the file wasn't found in the first attempt's scope
            # (though in this structure, 'data' should still be available)
            if historic_security_data is not None:
                historic_security_data.to_excel(excel_filename, sheet_name=sheet_name, index=True)
                print(f"Successfully created file and saved data to sheet '{sheet_name}'.")
        except Exception as e:
            print(f"Error creating file {excel_filename}: {e}")

    except Exception as e:
        print(f"\nAn error occurred while fetching or saving data for {ticker_symbol}: {e}")


# --- Example Usage ---
if __name__ == "__main__":

    # Read the list of tickers from the file
    ticker_list = read_tickers_from_file()

    if not ticker_list:
        print("Skipping download process. Please ensure 'tickers.txt' exists and contains tickers.")
    else:
        print(f"\nStarting download for {len(ticker_list)} tickers...")
        # Iterate over the list and download prices for each ticker
        for ticker in ticker_list:
            # You can customize start_date, end_date, or interval for all tickers here if needed.
            download_historical_prices(ticker_symbol=ticker)

    # 1. Old Example: download_historical_prices(ticker_symbol='AAPL')
    # 2. Old Example: download_historical_prices(ticker_symbol='MSFT', start_date='2019-01-01', end_date='2024-01-31')
