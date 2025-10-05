import pandas as pd
import matplotlib.pyplot as plt

# --- Load Excel Sheet ---
# Replace with your actual file and sheet name
excel_file = 'historical_prices.xlsx'
sheet_name = 'AAPL'  # Example: use 'MSFT', 'GOOGL', etc.

# Read the data
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Ensure 'date' column is in datetime format
df['date'] = pd.to_datetime(df['date'])

# --- Plotting ---
plt.figure(figsize=(12, 6))
plt.plot(df['date'], df['Close'], label='Close Price', color='blue')

# Add labels and title
plt.xlabel('Date')
plt.ylabel('Price')
plt.title(f'Date-wise Closing Price for {sheet_name}')
plt.grid(True)
plt.legend()
plt.tight_layout()

# Show the plot
plt.show()
A