from flask import Flask, request, render_template, send_file
from test_alpaca_backtest_gemini6 import load_historical_data_alpaca, backtest_long_volatility
import io
import pandas as pd

app = Flask(__name__)
last_trades = []  # We'll store trades globally for download

@app.route("/", methods=["GET", "POST"])
def index():
    global last_trades
    result = None

    if request.method == "POST":
        ticker = request.form["ticker"].upper()
        start_date = request.form["start_date"]
        end_date = request.form["end_date"]
        strategy = request.form["strategy"]
        use_strangle = strategy.lower() == "strangle"

        data = load_historical_data_alpaca(ticker, start_date, end_date)
        if data is not None:
            trades, final_capital = backtest_long_volatility(data, use_strangle=use_strangle)
            total_profit = final_capital - 10000
            result = {
                "strategy": strategy.capitalize(),
                "final_capital": f"${final_capital:.2f}",
                "total_profit": f"${total_profit:.2f}",
                "num_trades": len(trades),
                "download_ready": True
            }
            last_trades = trades

    return render_template("index.html", result=result)

@app.route("/download")
def download_excel():
    global last_trades
    if not last_trades:
        return "No data to export", 400

    df = pd.DataFrame(last_trades)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Backtest Results')

    output.seek(0)
    return send_file(output, as_attachment=True, download_name="backtest_results.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    app.run(debug=True)
