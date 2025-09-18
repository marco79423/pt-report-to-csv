#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
投資組合績效報告轉 CSV 轉換器
將 MultiCharts 投資組合績效報告的 Excel 檔轉為 CSV 格式
"""

import pandas as pd
import csv
from datetime import datetime
from pathlib import Path


def read_symbol_point_values(symbol_csv_path):
    """
    讀取 symbol.csv 檔案，建立商品名稱與一大點價值的對應表
    """
    try:
        df = pd.read_csv(symbol_csv_path, encoding='utf-8-sig')
        # 建立字典，以商品名稱為 key，一大點價值為 value
        symbol_dict = dict(zip(df['商品名稱'], df['一大點價值']))
        return symbol_dict
    except Exception as e:
        print(f"Error reading symbol CSV file: {e}")
        return {}


def read_trades_from_excel(excel_file_path):
    """
    讀取 Excel 檔中的「List of Trades」工作表
    表頭從第 3 列開始（索引 2）
    """
    try:
        # 讀取「List of Trades」工作表，從第 3 列（0 起算為 2）作為表頭
        df = pd.read_excel(excel_file_path, sheet_name='List of Trades', header=2)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def calculate_point_value(price1, price2, profit, contracts):
    """
    以價格差與損益計算每一大點的價值
    一大點價值 = 總損益 /（價格差 × 口數）
    """
    try:
        if contracts == 0 or pd.isna(contracts):
            return 0

        if pd.isna(price1) or pd.isna(price2) or pd.isna(profit):
            return 0

        price_diff = abs(float(price2) - float(price1))
        if price_diff == 0:
            return 0

        # 一大點價值 = 總損益 /（價格差 × 口數）
        point_val = abs(float(profit)) / (price_diff * abs(float(contracts)))
        return point_val
    except (ValueError, ZeroDivisionError, TypeError) as e:
        print(f"Error calculating point value: {e}, price1={price1}, price2={price2}, profit={profit}, contracts={contracts}")
        return 0


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
        print(f"Error combining datetime: {e}")
        return ""


def save_to_csv(data, output_file_path):
    """
    將處理後的資料輸出為 CSV 檔
    """
    try:
        df = pd.DataFrame(data)
        df.to_csv(output_file_path, index=False, encoding='utf-8-sig')
        print(f"CSV file saved successfully: {output_file_path}")
    except Exception as e:
        print(f"Error saving CSV file: {e}")


def main():
    """
    主程式：將 Portfolio Performance Report.xlsx 轉為 CSV
    """
    # 輸入與輸出檔案路徑
    excel_file = "Portfolio Performance Report.xlsx"
    csv_file = "portfolio_trades.csv"
    symbol_file = "symbol.csv"

    # 檢查檔案是否存在
    if not Path(excel_file).exists():
        print(f"Error: {excel_file} not found in current directory")
        return
    
    if not Path(symbol_file).exists():
        print(f"Error: {symbol_file} not found in current directory")
        return

    print(f"Reading Excel file: {excel_file}")
    print(f"Reading symbol file: {symbol_file}")

    # 讀取商品一大點價值對應表
    symbol_dict = read_symbol_point_values(symbol_file)
    if not symbol_dict:
        print("Failed to read symbol CSV file or no data found")
        return

    print(f"Loaded {len(symbol_dict)} symbol point values")

    # 從 Excel 讀取交易資料
    df = read_trades_from_excel(excel_file)
    if df is None:
        print("Failed to read Excel file")
        return

    print(f"Found {len(df)} rows of trading data")
    print("Columns:", list(df.columns))

    # 處理資料
    processed_data = process_trading_data(df, symbol_dict)

    if processed_data:
        print(f"Processed {len(processed_data)} trade records")

        # 輸出 CSV
        save_to_csv(processed_data, csv_file)
    else:
        print("No data to process")


if __name__ == "__main__":
    main()