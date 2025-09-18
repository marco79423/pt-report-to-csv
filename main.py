"""
投資組合績效報告轉 CSV 轉換器
"""

import pandas as pd
from datetime import datetime
from pathlib import Path


def main():
    """
    將 MultiCharts 投資組合績效報告的 Excel 檔轉為 CSV 格式
    """

    # 輸入與輸出檔案路徑
    portfolio_file = "Portfolio Performance Report.xlsx"
    csv_file = "portfolio_trades.csv"
    symbol_file = "symbol.csv"

    # 檢查檔案是否存在
    if not Path(portfolio_file).exists():
        print(f'檔案 {portfolio_file} 不存在')
        return
    print(f'讀取績效報告: {portfolio_file}')

    if not Path(symbol_file).exists():
        print(f'檔案 {symbol_file} 不存在')
        return
    print(f'讀取商品檔: {symbol_file}')

    # 讀取商品一大點價值對應表
    symbol_dict = read_symbol_point_values(symbol_file)
    if not symbol_dict:
        print('商品檔讀取失敗')
        return

    # 讀取交易資料
    df = read_trades_from_excel(portfolio_file)
    if df is None:
        print('績效報告讀取失敗')
        return

    # 處理資料
    print(f"讀取 {len(df)} 筆交易資料")
    processed_data = process_trading_data(df, symbol_dict)
    if processed_data:
        print(f'{len(df)} 筆交易資料處理完畢')
    else:
        print('沒有資料')
        return

    # 輸出 CSV
    save_to_csv(processed_data, csv_file)


def read_symbol_point_values(symbol_csv_path):
    try:
        df = pd.read_csv(symbol_csv_path, encoding='utf-8-sig')
        symbol_dict = dict(zip(df['商品名稱'], df['一大點價值']))
        return symbol_dict
    except Exception as e:
        print(f"讀取商品檔失敗: {e}")
        return {}


def read_trades_from_excel(excel_file_path):
    try:
        df = pd.read_excel(excel_file_path, sheet_name='List of Trades', header=2)
        return df
    except Exception as e:
        print(f'Excel 檔讀取失敗: {e}')
        return None

def process_trading_data(df, symbol_dict):
    """
    依需求處理交易資料：
    - 資料以成對（開倉/平倉）出現
    - 輸出欄位：商品名稱、交易時間、成交價、成交口數、一大點價值
    - EntryLong/ExitShort 視為買進（口數為正）
    - EntryShort/ExitLong 視為賣出（口數為負）
    - 一大點價值根據 symbol.csv 決定
    """
    results = []

    # 兩列為一組（每 2 列一組）進行處理
    for i in range(0, len(df), 2):
        if i + 1 >= len(df):
            break

        row1 = df.iloc[i]
        row2 = df.iloc[i + 1]

        # 根據商品名稱從 symbol.csv 查詢一大點價值
        symbol_name = row1['Symbol Name']
        point_value = symbol_dict.get(symbol_name, 0)  # 如果找不到商品，預設為 0

        # 取得價格資料（仍需要用於輸出）
        price1 = row1['Price']
        price2 = row2['Price']

        # 處理第一筆交易（開倉）
        date_time = combine_datetime(row1['Date'], row1['Time'])
        trade_type = row1['Type']

        # 判斷買賣別並設定口數正負號
        if trade_type in ['EntryLong', 'ExitShort']:
            contracts = abs(row1['Contracts'])  # 買進 = 正
        else:  # EntryShort, ExitLong
            contracts = -abs(row1['Contracts'])  # 賣出 = 負

        results.append({
            '商品名稱': symbol_name,
            '交易時間': date_time,
            '成交價': price1,
            '成交口數': contracts,
            '一大點價值': round(point_value, 2)
        })

        # 處理第二筆交易（平倉）
        # 平倉筆「商品名稱」沿用開倉筆，避免出現空白或不同名稱
        symbol_name2 = symbol_name
        date_time2 = combine_datetime(row2['Date'], row2['Time'])
        trade_type2 = row2['Type']

        # 判斷買賣別並設定口數正負號
        if trade_type2 in ['EntryLong', 'ExitShort']:
            contracts2 = abs(row2['Contracts'])  # 買進 = 正
        else:  # EntryShort, ExitLong
            contracts2 = -abs(row2['Contracts'])  # 賣出 = 負

        results.append({
            '商品名稱': symbol_name2,
            '交易時間': date_time2,
            '成交價': price2,
            '成交口數': contracts2,
            '一大點價值': round(point_value, 2)
        })

    return results


def combine_datetime(date, time):
    """
    合併日期與時間欄位為 yyyy/MM/dd HH:mm:ss 格式
    """
    try:
        if pd.isna(date) or pd.isna(time):
            return ""

        # 若為字串，先轉為 datetime
        if isinstance(date, str):
            date_obj = pd.to_datetime(date).date()
        else:
            date_obj = date.date() if hasattr(date, 'date') else date

        if isinstance(time, str):
            time_obj = pd.to_datetime(time).time()
        else:
            time_obj = time.time() if hasattr(time, 'time') else time

        # 合併日期與時間
        combined = datetime.combine(date_obj, time_obj)
        return combined.strftime('%Y/%m/%d %H:%M:%S')
    except Exception as e:
        print(f'時間整合失敗: {date}, {time}: {e}')
        return ""


def save_to_csv(data, output_file_path):
    """
    將處理後的資料輸出為 CSV 檔
    """
    try:
        df = pd.DataFrame(data)
        df.to_csv(output_file_path, index=False, encoding='utf-8-sig')
        print(f"CSV 儲存成功: {output_file_path}")
    except Exception as e:
        print(f"CSV 儲存失敗: {e}")



if __name__ == "__main__":
    main()