# hfstock

A simple graphical tool for querying A-share stock information. The GUI is built with Tkinter and retrieves data from the Eastmoney public APIs.

## Features

- **Search by code or name**: enter a stock code or name to get the latest information. If the stock cannot be found through the API but the input looks like a valid code, the code is used directly.
- **Market status handling**: determines whether the market has opened or closed and shows "未开盘" or "未收盘" when appropriate.
- **Display of key fields**: shows current closing price, daily percent change, previous closing price and previous percent change.
- **History dropdown**: keeps a list of all queries performed during the session so you can quickly review past results.
- **Excel export**: save the history to an `.xlsx` file with one click. The export uses `pandas` and `openpyxl`.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

Run the application with Python 3:

```bash
python3 stock.py
```

A window will open where you can enter a stock name or code to query. After multiple queries you can export the results from the bottom left button.
