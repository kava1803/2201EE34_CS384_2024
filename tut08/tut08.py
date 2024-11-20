import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import mplfinance as mpf
import matplotlib.dates as mdates

# Load the dataset
df = pd.read_csv("./infy_stock.csv")

# Show the first 10 rows
df.head(10)

"""#Missing data iputation(Mean)"""

print(df.isnull().sum(), "\n\n")

# Fill missing values for only numerical columns (excluding 'Date')
df.fillna(df.select_dtypes(include=['float', 'int']).mean(), inplace=True)

print(df.isnull().sum(), "\n\n")

df.info()

"""#Closing price Line Chart using Matplotlib"""

# Format the Date column as datetime if not already done
df['Date'] = pd.to_datetime(df['Date'])

# Plot the Closing Price over Time
plt.figure(figsize=(12, 6))
plt.plot(df['Date'], df['Close'], color='gray', label='Closing Price')

# Formatting the date to show only years on the x-axis
plt.gca().xaxis.set_major_locator(mdates.YearLocator())
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y'))

# Other plot settings
plt.title('Closing Price of Stock Over Time')
plt.xlabel('Date')
plt.ylabel('Closing Price')
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

"""#Candlestick plot of the whole Dataframe using Mplfinance(Too much data to plot)"""

# Candlestick plot using mplfinance
mpf.plot(df.set_index('Date'), type='candle', style='charles',title="Candlestick Chart", ylabel="Price")

"""#Candlestick plot of trailing 100 datapoints(Visually more understandable)"""

# Limiting the data to the last 100 entries
mpf.plot(df.set_index('Date').tail(100), type='candle', style='charles', title="Candlestick Chart (Last 100 Days)", ylabel="Price")

"""#Statistical Calculations"""

# Calculate Daily Return Percentage
df['Daily Return'] = ((df['Close'] - df['Open']) / df['Open']) * 100

# Calculate average and median of daily returns
mean_return = df['Daily Return'].mean()
median_return = df['Daily Return'].median()
std_close = df['Close'].std()

print(f"Mean Daily Return: {mean_return}")
print(f"Median Daily Return: {median_return}")
print(f"Standard Deviation of Closing Prices: {std_close} \n\n")

df.sample(10)

"""# Moving Averages"""

# Calculate the 50-day and 200-day moving averages
df['50-day MA'] = df['Close'].rolling(window=50).mean()
df['200-day MA'] = df['Close'].rolling(window=200).mean()

# Plot the moving averages
plt.figure(figsize=(10, 6))
plt.plot(df['Date'], df['Close'], label='Closing Price', color='blue')
plt.plot(df['Date'], df['50-day MA'], label='50-day Moving Average', color='red')
plt.plot(df['Date'], df['200-day MA'], label='200-day Moving Average', color='green')
plt.title("50-day and 200-day Moving Averages")
plt.xlabel("Time")
plt.ylabel("Price")
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

"""#Volatility"""

df['30-day Volatility'] = df['Close'].rolling(window=30).std()

# Plot the 30-day rolling volatility
plt.figure(figsize=(10, 6))
sns.lineplot(x='Date', y='30-day Volatility', data=df, color='purple')

# Format the x-axis to show only the year
plt.gca().xaxis.set_major_locator(mdates.YearLocator())
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y'))

# Plot settings
plt.title("30-day Rolling Volatility")
plt.xlabel("Time")
plt.ylabel("Volatility (Standard Deviation)")
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

"""#Bullish and Bearish Trends"""

# Identify bullish (1) and bearish (-1) trends
df['Trend'] = 0
df.loc[df['50-day MA'] > df['200-day MA'], 'Trend'] = 1
df.loc[df['50-day MA'] < df['200-day MA'], 'Trend'] = -1

# Plot the trends
plt.figure(figsize=(10, 6))
plt.plot(df['Date'], df['Close'], label='Closing Price', color='blue')
plt.plot(df['Date'], df['50-day MA'], label='50-day MA', color='red')
plt.plot(df['Date'], df['200-day MA'], label='200-day MA', color='green')

# Mark bullish trends in green and bearish trends in red
plt.fill_between(df['Date'], df['Close'], where=(df['Trend'] == 1), color='green', alpha=0.2, label='Bullish Trend')
plt.fill_between(df['Date'], df['Close'], where=(df['Trend'] == -1), color='red', alpha=0.2, label='Bearish Trend')

plt.title("Bullish and Bearish Trends")
plt.xlabel("Time")
plt.ylabel("Price")
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

df.sample(10)

