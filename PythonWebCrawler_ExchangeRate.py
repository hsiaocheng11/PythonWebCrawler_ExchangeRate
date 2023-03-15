import requests
from io import StringIO
import pandas as pd
from prettytable import PrettyTable
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
import matplotlib.pyplot as plt
from openpyxl.utils.dataframe import dataframe_to_rows
import os
print("目前支援的幣別 美金 (USD)、港幣 (HKD)、英鎊 (GBP)、澳幣 (AUD)、加拿大幣 (CAD)、新加坡幣 (SGD)、瑞士法郎 (CHF)、日圓 (JPY)、南非幣 (ZAR)、瑞典幣 (SEK)、紐元 (NZD)、泰幣 (THB)、菲國比索 (PHP)、印尼幣 (IDR)、歐元 (EUR)、韓元 (KRW)、越南盾 (VND)、馬來幣 (MYR)、人民幣 (CNY)")
while True:
    try:
        option = input("請輸入 1.台幣兌換外幣匯率時間範圍查詢工具 2.台幣兌換外幣計算機")
        if (option == "1"):
            currencyId = input("請輸入欲查詢幣別")
            optionHistory = input("1:欲查詢特定年/月(只支援到過去1年) 2:欲查詢最近三個月或半年")
            if (optionHistory == "1"):
                dateHistory = input("請輸入欲查詢時間範圍 (格式2023-02)")
                url = 'https://rate.bot.com.tw/xrt/flcsv/0/' + \
                dateHistory+'/'+currencyId   # 牌告匯率 CSV 網址
            else:
                dateHistory = input("請輸入欲查詢時間範圍 (1:最近半年)/(2:最近三個月))")
                if (dateHistory=="1"):
                    url = 'https://rate.bot.com.tw/xrt/flcsv/0/l6m'+'/'+currencyId   # 牌告匯率 CSV 網址
                else:
                    url = 'https://rate.bot.com.tw/xrt/flcsv/0/L3M'+'/'+currencyId   # 牌告匯率 CSV 網址  

            rate = requests.get(url)   # 爬取網址內容
            rate.encoding = 'utf-8'    # 調整回應訊息編碼為 utf-8，避免編碼不同造成亂碼

            # 將回應的文字資料轉換成檔案物件
            rate_file = StringIO(rate.text)

            # 讀取 CSV 檔案
            date_list = []
            rate_list = []
            for line in rate_file.readlines()[1:]:
                items = line.strip().split(',')
                date_list.append(items[0])
                rate_list.append(float(items[13]))

            # 建立 DataFrame
            df = pd.DataFrame({'Date': date_list, 'Exchange Rate': rate_list})

            # 設定圖表樣式
            plt.style.use('seaborn')

            # 繪製折線圖
            plt.plot(date_list, rate_list)

            # 建立 Excel 檔案
            wb = Workbook()
            ws = wb.active

            # 將 DataFrame 寫入 Excel
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

            # 設定圖表範圍
            xdata = Reference(ws, min_col=1, min_row=2, max_row=len(date_list)+1)
            ydata = Reference(ws, min_col=2, min_row=1, max_row=len(rate_list)+1)

            # 建立圖表物件
            chart = LineChart()
            chart.add_data(ydata, titles_from_data=True)
            chart.set_categories(xdata)

            # 設定圖表寬度和高度
            chart.width = chart.width * 1.5
            chart.height = chart.height * 1.5

            # 將圖表加入 Excel
            ws.add_chart(chart, 'D2')  

            # 儲存 Excel 檔案
            if not os.path.exists('result'):
                os.makedirs('result')

            filename = 'exchange_rate.xlsx'
            filepath = os.path.join('result', filename)

            wb.save(filepath)
            print(f'{filename}已儲存至{filepath}')
            resumeOption = input("1.繼續 2.結束")
            if (resumeOption=="2"):break
        else:
            url = 'https://rate.bot.com.tw/xrt/flcsv/0/day'   # 牌告匯率 CSV 網址
            rate = requests.get(url)   # 爬取網址內容
            rate.encoding = 'utf-8'    # 調整回應訊息編碼為 utf-8，避免編碼不同造成亂碼
            rt = rate.text             # 以文字模式讀取內容
            rts = rt.split('\n')       # 使用「換行」將內容拆分成串列

            table = PrettyTable()      # 建立 PrettyTable 物件
            table.field_names = ["幣別", "匯率"]  # 設定欄位名稱

            for i in rts[1:]:  # 讀取串列的每個項目
                try:                             # 使用 try 避開最後一行的空白行
                    a = i.split(',')             # 每個項目用逗號拆分成子串列
                    table.add_row([a[0], a[12]])  # 取出第一個 ( 0 ) 和第十三個項目 ( 12 )，並加入表格
                except:
                    break

            print(table)                # 印出表格

            # 讓使用者輸入欲轉換的金額和幣別
            amount = input("請輸入欲轉換的金額")
            currency = input("請輸入欲轉換的幣別")

            # 找到使用者輸入的幣別對應的匯率
            for i in rts[1:]:
                try:
                    a = i.split(',')
                    if a[0] == currency:
                        rate = float(a[12])
                        break
                except:
                    break

            # 計算轉換後的金額
            converted_amount = float(amount) * rate

            # 輸出轉換後的金額
            print(f"{amount} {currency} 可轉換為 {converted_amount:.2f} TWD")
            resumeOption = input("1.繼續 2.結束")
            if (resumeOption=="2"):break
    except:break