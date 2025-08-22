import pandas as pd
import matplotlib.pyplot as plt

# Read Excel file
file_path = 'Master_Summary.xlsx'
df = pd.read_excel(file_path)

# Ensure Ticker is string
df['Ticker'] = df['Ticker'].astype(str)

# Initial Capital
initial_capital = 100000

# Total combined profit across all stocks
total_profit = df['Profit'].sum()

# Total final capital
total_final_capital = initial_capital + total_profit

# Plotting combined bar
fig, ax = plt.subplots(figsize=(8,6))

ax.bar('All Stocks', initial_capital, label='Initial Capital', color='skyblue')
ax.bar('All Stocks', total_profit, bottom=initial_capital, label='Total Profit', color='orange')

ax.set_ylabel('Amount')
ax.set_title('Initial Capital vs Total Profit for All Stocks')
ax.legend()

plt.tight_layout()
plt.show()

print(f"Initial Capital: ₹{initial_capital}")
print(f"Total Profit from all stocks: ₹{total_profit}")
print(f"Total Capital at Year End: ₹{total_final_capital}")
