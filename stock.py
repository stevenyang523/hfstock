import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import datetime
import pandas as pd  # 需要 pip install pandas openpyxl

# --------- 样式设置（INS风格简约） ---------
BG_COLOR = "#f8f8f8"
BTN_COLOR = "#1a73e8"  # 更现代的蓝色
BTN_FG = "#fff"
FONT_MAIN = ("Microsoft YaHei UI", 14)  # 使用微软雅黑
FONT_RESULT = ("Microsoft YaHei UI", 12)
FONT_TITLE = ("Microsoft YaHei UI", 22, "bold")
FONT_AUTHOR = ("Microsoft YaHei UI", 11)  # 稍微加大字号
AUTHOR_COLOR = "#1a73e8"  # 使用与标题相同的蓝色

def get_stock_info(query):
    """
    查询A股股票信息，query可以是股票代码或名称
    返回dict，包含所需字段
    """
    code = None
    name = None
    market = None
    try:
        url = f"https://searchapi.eastmoney.com/api/suggest/get?input={query}&type=14"
        resp = requests.get(url, timeout=5)
        data = resp.json()
        data_table = data.get("QuotationCodeTable") or data.get("Data") or {}
        stock_list = data_table.get("Data") or data_table.get("data") or []
        if stock_list:
            stock = stock_list[0]
            code = stock.get("Code") or stock.get("code")
            name = stock.get("Name") or stock.get("name")
            market = stock.get("Market") or stock.get("market")  # 0=深市, 1=沪市
    except Exception as e:
        code = name = market = None

    # 如果接口查不到，且输入是6位数字，直接用输入
    if (not code or not name or market is None) and query.isdigit() and len(query) == 6:
        code = query
        # 科创板688开头也属于沪市
        if code.startswith("6") or code.startswith("688"):
            market = "1"
        else:
            market = "0"
        name = ""

    if not code or market is None:
        return {"error": "未找到该股票（代码或市场信息缺失）"}

    # 2. 东方财富K线接口获取最近两天行情
    try:
        secid = f"{str(market)}.{code}"
        url2 = (
            f"https://push2his.eastmoney.com/api/qt/stock/kline/get"
            f"?secid={secid}&fields1=f1,f2,f3,f4,f5,f6"
            f"&fields2=f51,f52,f53,f54,f55,f56,f57,f58"
            f"&klt=101&fqt=1&end=20500101&lmt=2"
        )
        resp2 = requests.get(url2, timeout=5)
        j = resp2.json()
        print("DEBUG: url2 =", url2)
        print("DEBUG: kline接口返回：", j)
        klines = j.get("data", {}).get("klines")
        # 用K线接口返回的name字段补全股票名称
        if not name:
            name = j.get("data", {}).get("name", "")
        if not name:
            name = "-"
        if not klines or not isinstance(klines, list) or len(klines) < 1:
            return {"error": "该股票无历史行情，可能为停牌、新股、退市或接口异常"}
        # 过滤掉空或格式不对的K线
        valid_klines = [k for k in klines if k and len(k.split(",")) >= 8]
        if not valid_klines:
            return {"error": "无有效K线数据"}
        def parse_kline(kline):
            arr = kline.split(",")
            return {
                "date": arr[0],
                "open": float(arr[1]),
                "close": float(arr[2]),
                "pct": float(arr[7])  # 涨跌幅百分比
            }
        today_info = parse_kline(valid_klines[-1])
        today_date = today_info["date"]
        today_close = today_info["close"]
        today_pct = today_info["pct"]

        prev_close = "-"
        prev_pct = "-"
        if len(valid_klines) >= 2:
            prev_info = parse_kline(valid_klines[-2])
            prev_close = prev_info["close"]
            prev_pct = prev_info["pct"]

        now_dt = datetime.datetime.now()
        if now_dt.strftime("%Y-%m-%d") != today_date:
            now_price_str = "未开盘"
            pct_today_str = "未开盘"
            closed = True
        elif now_dt.hour < 15:
            now_price_str = "未收盘"
            pct_today_str = "未收盘"
            closed = False
        else:
            now_price_str = f"{today_close:.2f}"
            pct_today_str = f"{today_pct:+.2f}%"
            closed = True
    except Exception as e:
        return {"error": f"行情数据获取失败（{e}）"}

    return {
        "name": name,
        "code": code,
        "now_price": now_price_str,
        "closed": closed,
        "pct_today": pct_today_str,
        "yesterday_close": prev_close,
        "prev_pct": f"{prev_pct:+.2f}%" if isinstance(prev_pct, float) else prev_pct
    }

class StockApp:
    def __init__(self, root):
        self.root = root
        self.root.title("华福计算机A股信息查询")
        self.root.configure(bg=BG_COLOR)
        self.history = []
        self.create_widgets()

    def create_widgets(self):
        # 主容器
        main_container = tk.Frame(self.root, bg=BG_COLOR)
        main_container.pack(fill="both", expand=True, padx=40, pady=20)

        # 标题
        title_frame = tk.Frame(main_container, bg=BG_COLOR)
        title_frame.pack(fill="x", pady=(0, 20))
        tk.Label(title_frame, text="华福计算机A股信息查询", font=FONT_TITLE, bg=BG_COLOR, fg="#1a73e8").pack()

        # 查询区
        query_frame = tk.Frame(main_container, bg=BG_COLOR)
        query_frame.pack(fill="x", pady=(0, 20))
        
        # 美化输入框
        entry_style = {"font": FONT_MAIN, "width": 25, "relief": "solid", "bd": 1}
        self.query_var = tk.StringVar()
        entry = tk.Entry(query_frame, textvariable=self.query_var, **entry_style)
        entry.pack(side="left", padx=(0, 10))
        entry.focus()

        # 美化查询按钮
        btn_style = {
            "font": FONT_MAIN,
            "bg": BTN_COLOR,
            "fg": BTN_FG,
            "relief": "flat",
            "cursor": "hand2",
            "padx": 20,
            "pady": 5
        }
        btn = tk.Button(query_frame, text="查询", command=self.query, **btn_style)
        btn.pack(side="left")

        # 结果区
        self.result_frame = tk.Frame(main_container, bg=BG_COLOR)
        self.result_frame.pack(fill="both", expand=True, pady=(0, 20))

        # 结果内容
        self.result_labels = {}
        for i, key in enumerate([
            "股票名称", "股票代码", "当日收盘价", "当日涨跌幅", "前一交易日收盘价", "前一交易日涨跌幅"
        ]):
            lbl = tk.Label(self.result_frame, text=f"{key}：", font=FONT_RESULT, anchor="w", bg=BG_COLOR)
            lbl.grid(row=i, column=0, sticky="w", pady=5)
            val = tk.Label(self.result_frame, text="", font=FONT_RESULT, anchor="w", bg=BG_COLOR, fg="#333")
            val.grid(row=i, column=1, sticky="w", pady=5)
            self.result_labels[key] = val

        # 历史查询下拉框
        history_frame = tk.Frame(main_container, bg=BG_COLOR)
        history_frame.pack(fill="x", pady=(0, 15))
        tk.Label(history_frame, text="历史查询结果：", font=FONT_RESULT, bg=BG_COLOR).pack(side="left")
        self.history_var = tk.StringVar()
        self.history_combo = ttk.Combobox(history_frame, textvariable=self.history_var, state="readonly", width=50, font=FONT_RESULT)
        self.history_combo.pack(side="left", padx=10, fill="x", expand=True)

        # 导出按钮和作者信息放在同一行
        bottom_frame = tk.Frame(main_container, bg=BG_COLOR)
        bottom_frame.pack(fill="x", pady=(0, 15))
        
        # 导出按钮靠左
        export_btn = tk.Button(bottom_frame, text="一键导出Excel", command=self.export_excel, **btn_style)
        export_btn.pack(side="left")
        
        # 作者信息靠右，增加宽度确保文字完整显示
        author_label = tk.Label(
            bottom_frame, 
            text="作者：王鑫旸", 
            font=FONT_AUTHOR, 
            bg=BG_COLOR, 
            fg=AUTHOR_COLOR,
            padx=10,
            width=15,  # 增加固定宽度
            anchor="e"  # 文字右对齐
        )
        author_label.pack(side="right", padx=(0, 10))  # 增加右侧内边距

    def query(self):
        q = self.query_var.get().strip()
        if not q:
            messagebox.showwarning("提示", "请输入股票名称或代码")
            return
        self.show_result({"股票名称": "查询中...", "股票代码": "", "当日收盘价": "", "当日涨跌幅": "",
                          "前一交易日收盘价": "", "前一交易日涨跌幅": ""})
        self.root.after(100, lambda: self._do_query(q))

    def _do_query(self, q):
        info = get_stock_info(q)
        if "error" in info:
            messagebox.showerror("查询失败", info["error"])
            self.show_result({k: "" for k in self.result_labels})
            return

        # 结果填充
        result = {
            "股票名称": info["name"],
            "股票代码": info["code"],
            "当日收盘价": info["now_price"],
            "当日涨跌幅": info["pct_today"],
            "前一交易日收盘价": f"{info['yesterday_close']:.2f}" if isinstance(info['yesterday_close'], float) else info['yesterday_close'],
            "前一交易日涨跌幅": info["prev_pct"]
        }
        self.show_result(result)

        # 保存到历史
        display_str = f"{result['股票名称']}({result['股票代码']}) | 收盘:{result['当日收盘价']} | 涨跌幅:{result['当日涨跌幅']} | 昨收:{result['前一交易日收盘价']} | 昨涨跌幅:{result['前一交易日涨跌幅']}"
        self.history.append(result)
        self.history_combo['values'] = [f"{i+1}. {self.history[i]['股票名称']}({self.history[i]['股票代码']})" for i in range(len(self.history))]
        self.history_combo.current(len(self.history)-1)
        self.history_var.set(display_str)

        if "notice" in info:
            messagebox.showinfo("提示", info["notice"])

    def show_result(self, data):
        for k, v in data.items():
            if k in self.result_labels:
                self.result_labels[k].config(text=v)

    def export_excel(self):
        if not self.history:
            messagebox.showwarning("提示", "没有可导出的查询结果")
            return
        df = pd.DataFrame(self.history)
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"{now_str}_查询.xlsx"
        # 弹出保存对话框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=default_filename,
            title="导出查询结果到Excel"
        )
        if not file_path:
            return
        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("导出成功", f"已导出到 {file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出Excel失败：{e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("500x500")
    app = StockApp(root)
    root.mainloop()
