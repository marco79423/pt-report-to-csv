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
    portfolio_files = ['Portfolio Performance Report.xlsx', '投資組合績效報告.xlsx']
    csv_file = "portfolio_trades.csv"
    symbol_file = "symbol.csv"

    # 檢查檔案是否存在
    target_portfolio_file = None
    for portfolio_file in portfolio_files:
        if Path(portfolio_file).exists():
            target_portfolio_file = portfolio_file
            break
    if not target_portfolio_file:
        print(f'績效報告檔不存在')
        return
    print(f'讀取績效報告: {target_portfolio_file}')

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
    df = read_trades_from_excel(target_portfolio_file)
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
        symbol_dict = {}
        for _, row in df.iterrows():
            symbol_name = row['商品名稱']
            point_value = row['一大點價值']
            fee = row.get('手續費', 0)  # 如果沒有手續費欄位，預設為 0
            symbol_dict[symbol_name] = {
                'point_value': point_value,
                'fee': fee
            }
        return symbol_dict
    except Exception as e:
        print(f"讀取商品檔失敗: {e}")
        return {}


def read_trades_from_excel(excel_file_path):
    """
    讀取 Excel 檔案，支援中英文工作表名稱
    """
    # 支援的工作表名稱（中英文）
    sheet_names = ['List of Trades', '交易明細']
    
    for sheet_name in sheet_names:
        try:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=2)
            print(f'成功讀取工作表: {sheet_name}')
            return df
        except Exception:
            continue
    
    print(f'Excel 檔讀取失敗: 找不到支援的工作表 {sheet_names}')
    return None

def normalize_symbol_name(symbol_name):
    """
    正規化符號名稱，移除括號及其內容
    例如: "OSE.NK225M HOT (30 Minutes)" -> "OSE.NK225M HOT"
    """
    if not symbol_name:
        return symbol_name
    
    # 找到第一個左括號的位置
    paren_pos = symbol_name.find('(')
    if paren_pos != -1:
        # 移除括號及其後面的內容，並去除尾部空白
        return symbol_name[:paren_pos].strip()
    
    return symbol_name


def get_column_value(row, column_mappings):
    """
    根據欄位對照表取得欄位值，支援中英文欄位名稱
    """
    for column_name in column_mappings:
        if column_name in row:
            return row[column_name]
    return None


def calculate_fee(fee_config, price, contracts):
    """
    計算手續費
    fee_config: 手續費設定，可以是固定值或百分比字串
    price: 成交價格
    contracts: 成交口數（絕對值）
    """
    if fee_config == 0 or not fee_config:
        return 0
    
    # 如果是百分比（含有 % 符號）
    if isinstance(fee_config, str) and '%' in str(fee_config):
        percentage = float(str(fee_config).replace('%', ''))
        return abs(price * contracts * percentage / 100)
    else:
        # 固定值
        return abs(float(fee_config) * contracts)


def process_trading_data(df, symbol_dict):
    """
    依需求處理交易資料：
    - 資料以成對（開倉/平倉）出現
    - 輸出欄位：商品名稱、交易時間、成交價、成交口數、一大點價值
    - EntryLong/ExitShort 視為買進（口數為正）
    - EntryShort/ExitLong 視為賣出（口數為負）
    - 一大點價值根據 symbol.csv 決定
    """
    # 欄位對照表（中英文）
    column_mappings = {
        'symbol_name': ['Symbol Name', '商品名稱'],
        'type': ['Type', '類型'],
        'price': ['Price', '價格'],
        'date': ['Date', '日期'],
        'time': ['Time', '時間'],
        'contracts': ['Contracts', '數量']
    }
    
    # 交易類型對照表（英文對中文）
    type_mappings = {
        'EntryLong': '進入Long',
        'EntryShort': '進入Short',
        'ExitShort': '離開Long',
        'ExitLong': '離開Short'
    }
    
    results = []

    # 兩列為一組（每 2 列一組）進行處理
    for i in range(0, len(df), 2):
        if i + 1 >= len(df):
            break

        row1 = df.iloc[i]
        row2 = df.iloc[i + 1]

        # 根據商品名稱從 symbol.csv 查詢一大點價值和手續費
        symbol_name = normalize_symbol_name(get_column_value(row1, column_mappings['symbol_name']))
        symbol_info = symbol_dict.get(symbol_name, {'point_value': 0, 'fee': 0})
        point_value = symbol_info.get('point_value', 0)
        fee_config = symbol_info.get('fee', 0)

        # 取得價格資料（仍需要用於輸出）
        price1 = get_column_value(row1, column_mappings['price'])
        price2 = get_column_value(row2, column_mappings['price'])

        # 處理第一筆交易（開倉）
        date1 = get_column_value(row1, column_mappings['date'])
        time1 = get_column_value(row1, column_mappings['time'])
        date_time = combine_datetime(date1, time1)
        trade_type = get_column_value(row1, column_mappings['type'])

        # 判斷買賣別並設定口數正負號
        if trade_type in ['EntryLong', 'ExitShort', '進入Long', '離開Long']:
            contracts = abs(get_column_value(row1, column_mappings['contracts']))  # 買進 = 正
        else:  # EntryShort, ExitLong, 進入Short, 離開Short
            contracts = -abs(get_column_value(row1, column_mappings['contracts']))  # 賣出 = 負

        # 計算手續費
        fee = calculate_fee(fee_config, price1, abs(contracts))

        results.append({
            '商品名稱': symbol_name,
            '交易時間': date_time,
            '成交價': price1,
            '成交口數': contracts,
            '手續費': round(fee, 2),
            '一大點價值': round(point_value, 2)
        })

        # 處理第二筆交易（平倉）
        # 平倉筆「商品名稱」沿用開倉筆，避免出現空白或不同名稱
        symbol_name2 = symbol_name
        date2 = get_column_value(row2, column_mappings['date'])
        time2 = get_column_value(row2, column_mappings['time'])
        date_time2 = combine_datetime(date2, time2)
        trade_type2 = get_column_value(row2, column_mappings['type'])

        # 判斷買賣別並設定口數正負號
        if trade_type2 in ['EntryLong', 'ExitShort', '進入Long', '離開Short']:
            contracts2 = abs(get_column_value(row2, column_mappings['contracts']))  # 買進 = 正
        else:  # EntryShort, ExitLong, 進入Short, 離開Long
            contracts2 = -abs(get_column_value(row2, column_mappings['contracts']))  # 賣出 = 負

        # 計算手續費
        fee2 = calculate_fee(fee_config, price2, abs(contracts2))

        results.append({
            '商品名稱': symbol_name2,
            '交易時間': date_time2,
            '成交價': price2,
            '成交口數': contracts2,
            '手續費': round(fee2, 2),
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