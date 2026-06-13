
import pandas as pd
import yfinance as yf
from datetime import time

# =========================
# SETTINGS
# =========================
TICKER = "^NSEI"
GAP_THRESHOLD = 0.22      # percent
RR_RATIO = 2.0            # 2 = 1:2, 3 = 1:3
LOOKBACK_DAYS = 60

# =========================
# DOWNLOAD DATA
# =========================
print("Downloading data...")
df = yf.download(
    TICKER,
    period=f"{LOOKBACK_DAYS}d",
    interval="5m",
    auto_adjust=False,
    progress=False
)

if df.empty:
    raise ValueError("No data downloaded from Yahoo Finance.")

if isinstance(df.columns, pd.MultiIndex):
    df.columns = [c[0] for c in df.columns]

df = df.reset_index()

dt_col = df.columns[0]
df["Datetime"] = pd.to_datetime(df[dt_col])
df["Date"] = df["Datetime"].dt.date
df["Time"] = df["Datetime"].dt.time

trades = []

for day in sorted(df["Date"].unique()):
    day_df = df[df["Date"] == day].copy()

    day_df = day_df.sort_values("Datetime").reset_index(drop=True)

# Use first candle of the trading day
first_candle = day_df.iloc[0]

    range_high = float(first_candle["High"])
    range_low = float(first_candle["Low"])
    day_open = float(first_candle["Open"])

    prev_days = sorted([d for d in df["Date"].unique() if d < day])
    if not prev_days:
        continue

    prev_day = prev_days[-1]
    prev_df = df[df["Date"] == prev_day]

    close_rows = prev_df[prev_df["Time"] <= time(15,30)]
    if len(close_rows) == 0:
        continue

    prev_close = float(close_rows.iloc[-1]["Close"])

    gap_pct = ((day_open - prev_close) / prev_close) * 100

    if abs(gap_pct) < GAP_THRESHOLD:
        continue

    trade_taken = False

    day_df = day_df.sort_values("Datetime").reset_index(drop=True)

    for i in range(len(day_df)-1):

        current = day_df.iloc[i]

        if current["Time"] <= time(9,15):
            continue

        close_price = float(current["Close"])

        direction = None

        if close_price > range_high:
            direction = "LONG"

        elif close_price < range_low:
            direction = "SHORT"

        if direction is None:
            continue

        entry_row = day_df.iloc[i+1]
        entry_price = float(entry_row["Open"])

        if direction == "LONG":
            sl = range_low
            risk = entry_price - sl

            if risk <= 0:
                break

            target = entry_price + risk * RR_RATIO

        else:
            sl = range_high
            risk = sl - entry_price

            if risk <= 0:
                break

            target = entry_price - risk * RR_RATIO

        exit_price = None
        result = "EOD"
        r_multiple = 0

        future = day_df.iloc[i+2:]

        for _, bar in future.iterrows():

            high = float(bar["High"])
            low = float(bar["Low"])

            if direction == "LONG":

                # SL FIRST ASSUMPTION
                if low <= sl:
                    exit_price = sl
                    result = "LOSS"
                    r_multiple = -1
                    break

                if high >= target:
                    exit_price = target
                    result = "WIN"
                    r_multiple = RR_RATIO
                    break

            else:

                if high >= sl:
                    exit_price = sl
                    result = "LOSS"
                    r_multiple = -1
                    break

                if low <= target:
                    exit_price = target
                    result = "WIN"
                    r_multiple = RR_RATIO
                    break

        if exit_price is None:
            eod = day_df[day_df["Time"] == time(15,25)]

if len(eod):
    exit_price = float(eod.iloc[0]["Close"])
    
            else:
                exit_price = float(day_df.iloc[-1]["Close"])

            if direction == "LONG":
                r_multiple = (exit_price - entry_price) / risk
            else:
                r_multiple = (entry_price - exit_price) / risk

        trades.append({
            "Date": day,
            "Gap %": round(gap_pct, 2),
            "Direction": direction,
            "Entry": round(entry_price, 2),
            "SL": round(sl, 2),
            "Target": round(target, 2),
            "Exit": round(exit_price, 2),
            "Result": result,
            "R Multiple": round(r_multiple, 2)
        })

        trade_taken = True
        break

results = pd.DataFrame(trades)

if len(results) == 0:
    print("No trades found.")
else:
    wins = (results["Result"] == "WIN").sum()
    losses = (results["Result"] == "LOSS").sum()

    summary = pd.DataFrame({
        "Metric": [
            "Total Trades",
            "Wins",
            "Losses",
            "Win Rate %",
            "Average R"
        ],
        "Value": [
            len(results),
            wins,
            losses,
            round((wins / len(results)) * 100, 2),
            round(results["R Multiple"].mean(), 2)
        ]
    })

    with pd.ExcelWriter("NIFTY_Gap_Breakout_Backtest.xlsx") as writer:
        results.to_excel(writer, sheet_name="Trades", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print("Saved: NIFTY_Gap_Breakout_Backtest.xlsx")
