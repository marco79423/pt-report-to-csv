# pt_report_to_csv

![logo](./logo.png)

Multicharts Porfolio report to csv 

結果可以用 [期貨交易績效分析器](https://toolset.marco79423.net/zh-TW/futures-performance) 呈現

## 使用方式

### Python 執行
```bash
python main.py
```

### 打包成執行檔
執行 `build_exe.bat` 即可自動安裝依賴並打包成單一執行檔：

```bash
build_exe.bat
```

打包完成後，執行檔會位於 `dist\pt-report-to-csv.exe`

### 執行檔使用說明
1. 將 "Portfolio Performance Report.xlsx" 放在執行檔同一資料夾
2. 將 "symbol.csv" 放在執行檔同一資料夾  
3. 雙擊執行 pt-report-to-csv.exe
4. 程式會產生 "portfolio_trades.csv" 輸出檔

## 檔案說明
- `main.py` - 主程式
- `main.spec` - PyInstaller 設定檔
- `build_exe.bat` - 自動打包批次檔
- `symbol.csv` - 商品一大點價值對應表
- `Portfolio Performance Report.xlsx` - MultiCharts 績效報告檔案