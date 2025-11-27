import xlwings as xw
import yfinance as yf
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime


def main():
    wb = xw.Book.caller()
    app = wb.app

    # Get ticker from named range (before deleting sheet!)
    try:
        ticker = wb.names["TICKER"].refers_to_range.value
        if not ticker:
            app.alert(
                "Please enter a ticker symbol in the TICKER range!",
                title="Missing Ticker",
            )
            return
    except:
        app.alert(
            "Named range 'TICKER' not found!\n\nPlease create it first:\n1. Select a cell with ticker symbol\n2. Formulas → Define Name\n3. Name it 'TICKER'",
            title="Setup Required",
        )
        return

    # Delete and recreate dashboard sheet
    dashboard_name = "Stock Dashboard"
    try:
        wb.sheets[dashboard_name].delete()
    except:
        pass

    sheet = wb.sheets.add(dashboard_name, after=wb.sheets[0])
    sheet.activate()

    # Fetch stock data
    stock = yf.Ticker(ticker)
    hist = stock.history(period="1mo")

    if hist.empty:
        app.alert(
            f"No data found for ticker: {ticker.upper()}\n\nPlease check the ticker symbol and try again.",
            title="Invalid Ticker",
        )
        sheet.delete()
        return

    # Get company info
    try:
        info = stock.info
    except:
        info = {"longName": ticker.upper(), "sector": "N/A"}

    # Header section
    sheet.range("A1").value = f"✅ Stock Analysis Dashboard - {ticker.upper()}"
    sheet.range("A1").font.size = 18
    sheet.range("A1").font.bold = True
    sheet.range("A1").font.color = (0, 102, 204)

    # Stock info
    sheet.range("A3").value = "Ticker:"
    sheet.range("B3").value = ticker.upper()
    sheet.range("B3").font.bold = True
    sheet.range("B3").font.size = 14

    sheet.range("A4").value = "Company:"
    sheet.range("B4").value = info.get("longName", "N/A")

    sheet.range("A5").value = "Sector:"
    sheet.range("B5").value = info.get("sector", "N/A")

    sheet.range("A6").value = "Current Price:"
    current_price = hist["Close"].iloc[-1]
    sheet.range("B6").value = current_price
    sheet.range("B6").number_format = "$#,##0.00"
    sheet.range("B6").font.size = 12
    sheet.range("B6").font.bold = True

    # Calculate metrics
    price_change = hist["Close"].iloc[-1] - hist["Close"].iloc[0]
    price_change_pct = (price_change / hist["Close"].iloc[0]) * 100

    sheet.range("A7").value = "30-Day Change:"
    change_text = f"${price_change:.2f} ({price_change_pct:+.2f}%)"
    sheet.range("B7").value = change_text
    sheet.range("B7").font.color = (0, 128, 0) if price_change >= 0 else (255, 0, 0)
    sheet.range("B7").font.bold = True

    sheet.range("A8").value = "30-Day High:"
    sheet.range("B8").value = hist["High"].max()
    sheet.range("B8").number_format = "$#,##0.00"

    sheet.range("A9").value = "30-Day Low:"
    sheet.range("B9").value = hist["Low"].min()
    sheet.range("B9").number_format = "$#,##0.00"

    sheet.range("A10").value = "Avg Volume:"
    sheet.range("B10").value = f"{hist['Volume'].mean():,.0f}"

    # Historical data table
    sheet.range("D3").value = "Historical Data (Last 30 Days)"
    sheet.range("D3").font.bold = True
    sheet.range("D3").font.size = 12

    # Prepare data for Excel
    df_display = hist[["Open", "High", "Low", "Close", "Volume"]].copy()
    df_display.index = df_display.index.strftime("%Y-%m-%d")
    df_display = df_display.round(2)

    # Write to Excel
    sheet.range("D4").value = df_display

    # Format the table
    table_range = sheet.range(f"D4:I{4 + len(df_display)}")
    table_range.api.Borders.Weight = 2

    # Header row formatting
    header_range = sheet.range("D4:I4")
    header_range.color = (0, 102, 204)
    header_range.font.color = (255, 255, 255)
    header_range.font.bold = True

    # Auto-fit columns
    sheet.range("A:I").columns.autofit()

    # Create matplotlib chart
    # Create figure and subplots
    fig, (ax1, ax2) = plt.subplots(
        2, 1, figsize=(10, 8), sharex=True, gridspec_kw={"height_ratios": [3, 1]}
    )
    plt.subplots_adjust(hspace=0.1)

    # Price chart (Top)
    ax1.plot(
        hist.index, hist["Close"], color="#0066cc", label="Close Price", linewidth=2
    )
    ax1.fill_between(hist.index, hist["Close"], alpha=0.3, color="#0066cc")
    ax1.set_title(f"{ticker.upper()} - 30 Day Price Chart")
    ax1.set_ylabel("Price ($)")
    ax1.grid(True, alpha=0.3)

    # Add min/max markers
    max_idx = hist["Close"].idxmax()
    min_idx = hist["Close"].idxmin()
    max_val = hist["Close"].max()
    min_val = hist["Close"].min()

    ax1.scatter(max_idx, max_val, color="green", marker="^", s=100, zorder=5)
    ax1.scatter(min_idx, min_val, color="red", marker="v", s=100, zorder=5)

    # Volume chart (Bottom)
    colors = ["green" if c >= o else "red" for c, o in zip(hist["Close"], hist["Open"])]
    ax2.bar(hist.index, hist["Volume"], color=colors, alpha=0.6)
    ax2.set_ylabel("Volume")
    ax2.set_xlabel("Date")
    ax2.grid(True, alpha=0.3)

    # Format x-axis to show all dates
    ax2.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))

    # Rotate labels and align them
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45, ha="right")

    # Insert chart directly
    sheet.pictures.add(
        fig,
        name="stock_chart",
        update=True,
        left=sheet.range("K4").left,
        top=sheet.range("K4").top,
        width=600,
        height=450,
    )
    plt.close(fig)

    # Add timestamp
    sheet.range("A12").value = (
        f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    sheet.range("A12").font.size = 9
    sheet.range("A12").font.color = (128, 128, 128)

    # Success message
    app.alert(
        f"✅ Dashboard created successfully for {ticker.upper()}!\n\nCurrent Price: ${current_price:.2f}\n30-Day Change: {change_text}",
        title="Success",
    )


if __name__ == "__main__":
    xw.Book("demo_automation.xlsm").set_mock_caller()
    main()
