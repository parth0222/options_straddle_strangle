import alpaca_trade_api as tradeapi
import pandas as pd
from datetime import datetime, timedelta
import pytz
import xlsxwriter
import configparser
import os

# --- 1. Load Configuration ---
config = configparser.ConfigParser()
config.read('config.ini')

try:
    ALPACA_API_KEY = config['alpaca']['api_key']
    ALPACA_SECRET_KEY = config['alpaca']['secret_key']
    ALPACA_BASE_URL = config['alpaca']['base_url']
except KeyError as e:
    print(f"Error: Missing configuration in config.ini: {e}")
    exit()

# --- 2. Data Acquisition (for Underlying Asset) ---
def load_historical_data_alpaca(ticker, start_date, end_date):
    try:
        api = tradeapi.REST(ALPACA_API_KEY, ALPACA_SECRET_KEY, ALPACA_BASE_URL)
        barset = api.get_bars(ticker, "1D", start_date, end_date).df
        if barset.empty:
            print(f"Warning: No historical data found for {ticker} between {start_date} and {end_date} from Alpaca.")
            return None
        return barset[['close']].rename(columns={'close': 'Close'})
    except Exception as e:
        print(f"Error downloading data for {ticker} from Alpaca: {e}")
        return None

# --- 3. Backtest Long Straddle/Strangle Strategy (Simplified) ---
def backtest_long_volatility(historical_data, initial_capital=10000, capital_per_trade=2500, days_to_expiration=30, atm_offset_percentage=0.01, otm_offset_percentage_strangle=0.05, use_strangle=False):
    """
    This is a SIMPLIFIED backtest of a long straddle or strangle strategy.
    It does NOT use actual options data or implied volatility. Premiums are approximated.
    """
    print(f"\n--- Backtesting Long {'Strangle' if use_strangle else 'Straddle'} Strategy ---")

    if historical_data is None or historical_data.empty:
        print("Error: No historical data provided.")
        return [], initial_capital

    trades = []
    capital = initial_capital

    for i in range(len(historical_data) - days_to_expiration):
        entry_date = historical_data.index[i].date()
        entry_price = historical_data['Close'].iloc[i]

        # Approximate ATM strike for Straddle
        atm_strike = round(entry_price / 5) * 5

        call_strike = 0
        put_strike = 0

        if use_strangle:
            call_strike = round(entry_price * (1 + otm_offset_percentage_strangle) / 5) * 5
            put_strike = round(entry_price * (1 - otm_offset_percentage_strangle) / 5) * 5
        else:  # Straddle
            call_strike = atm_strike
            put_strike = atm_strike

        # Approximate premiums (very simplified)
        premium_call_per_share = 0.02 * entry_price * (1 + atm_offset_percentage) if not use_strangle else 0.015 * entry_price * (1 + otm_offset_percentage_strangle)
        premium_put_per_share = 0.02 * entry_price * (1 + atm_offset_percentage) if not use_strangle else 0.015 * entry_price * (1 + otm_offset_percentage_strangle)

        num_contracts = max(1, int(capital_per_trade // ((premium_call_per_share + premium_put_per_share) * 100)))
        total_cost = (premium_call_per_share + premium_put_per_share) * 100 * num_contracts

        expiration_date = historical_data.index[i + days_to_expiration].date()
        price_at_expiration = historical_data['Close'].iloc[i + days_to_expiration]

        call_profit = max(0, price_at_expiration - call_strike) * 100 * num_contracts - premium_call_per_share * 100 * num_contracts
        put_profit = max(0, put_strike - price_at_expiration) * 100 * num_contracts - premium_put_per_share * 100 * num_contracts

        total_profit = call_profit + put_profit - (0 if i == 0 else 0) # No cost on subsequent days in this simple model

        capital += total_profit
        trades.append({
            'entry_date': entry_date.strftime('%Y-%m-%d'),
            'expiration_date': expiration_date.strftime('%Y-%m-%d'),
            'underlying_price_at_entry': entry_price,
            'call_strike': call_strike,
            'put_strike': put_strike,
            'approx_call_premium_paid': premium_call_per_share * 100 * num_contracts,
            'approx_put_premium_paid': premium_put_per_share * 100 * num_contracts,
            'underlying_price_at_expiration': price_at_expiration,
            'call_profit': call_profit,
            'put_profit': put_profit,
            'total_profit': total_profit,
            'capital': capital,
            'strategy': 'Long Strangle' if use_strangle else 'Long Straddle',
            'num_contracts': num_contracts
        })

    return trades, capital

# --- 4. Export Results ---
def export_to_excel(trades, filename="long_volatility_backtest_alpaca.xlsx"):
    if not trades:
        print("No trades to export.")
        return

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    headers = list(trades[0].keys())
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    for row_num, trade in enumerate(trades):
        for col_num, key in enumerate(headers):
            worksheet.write(row_num + 1, col_num, trade[key])

    workbook.close()
    print(f"Backtest results exported to {filename}")

# --- 5. Main Execution ---
if __name__ == "__main__":
    ticker = input("Enter the stock ticker: ").upper()
    start_date_str = input("Enter the start date (YYYY-MM-DD): ")
    end_date_str = input("Enter the end date (YYYY-MM-DD): ")
    volatility_strategy = input("Enter 'straddle' or 'strangle': ").lower()

    start_date = start_date_str
    end_date = end_date_str
    print(f"Using backtest period: {start_date} to {end_date}")

    data = load_historical_data_alpaca(ticker, start_date, end_date)

    use_strangle_flag = False
    if volatility_strategy == 'strangle':
        use_strangle_flag = True
    elif volatility_strategy != 'straddle':
        print("Invalid volatility strategy entered. Defaulting to Straddle.")

    if data is not None:
        trades, final_capital = backtest_long_volatility(
            data,
            initial_capital=10000,
            capital_per_trade=2500,
            days_to_expiration=30,
            atm_offset_percentage=0.01, # Small offset for ATM approximation
            otm_offset_percentage_strangle=0.05, # Adjust for how far OTM you want the strangle
            use_strangle=use_strangle_flag
        )

        print(f"\n--- Alpaca Long {'Strangle' if use_strangle_flag else 'Straddle'} Backtesting Results ---")
        print(f"Final Capital: ${final_capital:.2f}")
        print(f"Total Profit: ${final_capital - 10000:.2f}")
        print(f"Number of Trades: {len(trades)}")

        export_to_excel(trades, filename=f"long_{'strangle' if use_strangle_flag else 'straddle'}_backtest_alpaca.xlsx")

        # Optional: Print trade details
        # for trade in trades:
        #     print(trade)