import pandas as pd
import yfinance as yf
import time

# ==========================================
# HELPER FUNCTION
# ==========================================

def max_streak(series, value):

    max_count = 0
    current = 0

    for x in series:

        if x == value:
            current += 1
            max_count = max(max_count, current)
        else:
            current = 0

    return max_count


# ==========================================
# READ INPUT FILE
# ==========================================

results = []

trades = pd.read_excel("input.xlsx")

print(f"Loaded {len(trades)} trades")

# ==========================================
# PROCESS EACH STOCK
# ==========================================

for _, row in trades.iterrows():

    stock = str(row["Stock"]).strip().upper()

    signal_date = pd.to_datetime(row["Date"])

    ticker = stock + ".NS"

    print(f"\nProcessing {ticker}")

    try:

        data = yf.download(
            ticker,
            start="2023-01-01",
            end=pd.Timestamp.today().strftime("%Y-%m-%d"),
            auto_adjust=False,
            progress=False
        )

        if data.empty:
            print(f"No data for {ticker}")
            continue

    except Exception as e:

        print(f"Download error for {ticker}: {e}")
        continue

    # Fix yfinance MultiIndex columns
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.droplevel(1)

    data.reset_index(inplace=True)

    data["Date"] = pd.to_datetime(data["Date"])

    # --------------------------------------
    # Find next trading day
    # --------------------------------------

    future = data[data["Date"] > signal_date]

    if len(future) == 0:

        print(f"No future data for {stock}")

        continue

    entry_day = future.iloc[0]

    entry_date = entry_day["Date"]

    entry_open = float(entry_day["Open"])

    # --------------------------------------
    # Target and Stoploss
    # --------------------------------------

    target = entry_open * 1.05

    stoploss = entry_open * 0.975

    # --------------------------------------
    # Check what hits first
    # --------------------------------------

    trade_data = future.copy()

    result = "NO HIT"

    days_taken = None

    exit_date = None

    for i, (_, candle) in enumerate(trade_data.iterrows()):

        high = float(candle["High"])

        low = float(candle["Low"])

        # TARGET FIRST

        if high >= target:

            result = "TARGET"

            days_taken = i

            exit_date = candle["Date"]

            break

        # STOPLOSS

        if low <= stoploss:

            result = "STOPLOSS"

            days_taken = i

            exit_date = candle["Date"]

            break

    # --------------------------------------
    # Trade Return
    # --------------------------------------

    if result == "TARGET":
        trade_return = 5.0

    elif result == "STOPLOSS":
        trade_return = -2.5

    else:
        trade_return = 0.0

    # --------------------------------------
    # Save Result
    # --------------------------------------

    results.append({

        "Stock": stock,

        "Signal Date": signal_date.date(),

        "Entry Date": entry_date.date(),

        "Exit Date": exit_date.date() if exit_date is not None else None,

        "Entry Open": round(entry_open, 2),

        "Target": round(target, 2),

        "Stoploss": round(stoploss, 2),

        "Result": result,

        "Days": days_taken,

        "Return %": trade_return

    })

    print(
        f"{stock} | "
        f"{result} | "
        f"Days={days_taken}"
    )

    time.sleep(1)

# ==========================================
# RESULTS DATAFRAME
# ==========================================

results_df = pd.DataFrame(results)

if results_df.empty:

    print("No results generated.")

    exit()

# ==========================================
# SUMMARY STATISTICS
# ==========================================

total_trades = len(results_df)

winners = len(
    results_df[
        results_df["Result"] == "TARGET"
    ]
)

losers = len(
    results_df[
        results_df["Result"] == "STOPLOSS"
    ]
)

no_hits = len(
    results_df[
        results_df["Result"] == "NO HIT"
    ]
)

win_rate = round(
    winners / total_trades * 100,
    2
)

avg_days_target = round(
    results_df[
        results_df["Result"] == "TARGET"
    ]["Days"].mean(),
    2
)

avg_days_stoploss = round(
    results_df[
        results_df["Result"] == "STOPLOSS"
    ]["Days"].mean(),
    2
)

gross_profit = winners * 5

gross_loss = losers * 2.5

if gross_loss > 0:
    profit_factor = round(
        gross_profit / gross_loss,
        2
    )
else:
    profit_factor = "Infinity"

expectancy = round(
    results_df["Return %"].mean(),
    2
)

max_win_streak = max_streak(
    results_df["Result"],
    "TARGET"
)

max_loss_streak = max_streak(
    results_df["Result"],
    "STOPLOSS"
)

# ==========================================
# MONTHLY STATS
# ==========================================

results_df["Signal Date"] = pd.to_datetime(
    results_df["Signal Date"]
)

results_df["Month"] = (
    results_df["Signal Date"]
    .dt.to_period("M")
)

monthly = (
    results_df
    .groupby("Month")
    .agg(
        Trades=("Result", "count"),
        Winners=("Result",
                 lambda x:
                 (x == "TARGET").sum())
    )
)

monthly["Win Rate %"] = round(
    monthly["Winners"]
    / monthly["Trades"]
    * 100,
    2
)

# ==========================================
# SUMMARY SHEET
# ==========================================

summary = pd.DataFrame({

    "Metric": [

        "Total Trades",

        "Winners",

        "Losers",

        "No Hit",

        "Win Rate %",

        "Average Days Target",

        "Average Days Stoploss",

        "Profit Factor",

        "Expectancy %",

        "Max Win Streak",

        "Max Loss Streak"

    ],

    "Value": [

        total_trades,

        winners,

        losers,

        no_hits,

        win_rate,

        avg_days_target,

        avg_days_stoploss,

        profit_factor,

        expectancy,

        max_win_streak,

        max_loss_streak

    ]

})

# ==========================================
# SAVE TO EXCEL
# ==========================================

with pd.ExcelWriter(
    "results.xlsx",
    engine="openpyxl"
) as writer:

    results_df.to_excel(
        writer,
        sheet_name="Trades",
        index=False
    )

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    monthly.to_excel(
        writer,
        sheet_name="Monthly Stats"
    )

print("\n================================")
print("BACKTEST COMPLETE")
print("results.xlsx created")
print("================================")