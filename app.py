from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

# 你的 Excel 檔案名稱
EXCEL_FILE = 'data.xlsx'

@app.route('/', methods=['GET', 'POST'])
def index():
    name = ''
    results = None
    roster_info = None
    expense_summary = None
    message = ''
    custom_heading = ''

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if not name:
            message = '請輸入姓名！'
            return render_template('index.html', name=name, message=message)

        try:
            # 讀取 Excel 的兩個分頁
            df_integrate = pd.read_excel(EXCEL_FILE, sheet_name='整合')
            df_roster = pd.read_excel(EXCEL_FILE, sheet_name='名冊')

            # 篩選整合分頁的資料：找個案姓名
            results = df_integrate[df_integrate['姓名'] == name]

            if results.empty:
                message = f'找不到個案姓名：{name}'
                return render_template('index.html', name=name, message=message)

            # 轉型前處理：將 '-' 替換成 '0'，再轉成數字型態，避免轉型錯誤
            num_cols = ['費用', '看護費', '車資']  # 根據實際欄位名稱調整
            for col in num_cols:
                if col in results.columns:
                    results[col] = results[col].astype(str).replace('-', '0')
                    results[col] = pd.to_numeric(results[col], errors='coerce').fillna(0).astype(int)

            # 篩選名冊分頁的個案資料
            roster_filtered = df_roster[df_roster['姓名'] == name]
            if not roster_filtered.empty:
                # 假設名冊有這些欄位，取第一筆資料
                roster_info = {
                    '月費': roster_filtered.iloc[0].get('月費', 0),
                    '補助款': roster_filtered.iloc[0].get('補助款', 0),
                    '雜費': roster_filtered.iloc[0].get('雜費', 0),
                    '積欠': roster_filtered.iloc[0].get('積欠', 0),
                    '溢收': roster_filtered.iloc[0].get('溢收', 0),
                    '合計': roster_filtered.iloc[0].get('合計', 0),
                }
            else:
                roster_info = None

            # 雜費統計（依類別分類加總）
            expense_summary = {}
            # 預設分類欄位，依照你 Excel 欄位調整
            expense_categories = ['醫療', '看護費', '車資', '耗材', '其他', '農會購物', '退費']

            for cat in expense_categories:
                if cat == '退費':
                    # 沖銷退費加總：費用欄位為負值的和（或者你可用特定判斷）
                    expense_summary[cat] = results.loc[results['類別'] == '沖銷', '費用'].sum()
                else:
                    expense_summary[cat] = results.loc[results['類別'] == cat, '費用'].sum()

            # 雜費小計(不包含退費)
            expense_summary['雜費小計'] = sum(expense_summary[cat] for cat in expense_categories if cat != '退費')
            # 雜費總計 = 小計 - 退費
            expense_summary['雜費總計'] = expense_summary['雜費小計'] + expense_summary.get('退費', 0)

            custom_heading = '本月'  # 可依需求改動

        except Exception as e:
            message = f'發生錯誤：{e}'

    return render_template('index.html',
                           name=name,
                           results=results,
                           roster_info=roster_info,
                           expense_summary=expense_summary,
                           message=message,
                           custom_heading=custom_heading)


if __name__ == '__main__':
    app.run(debug=True)
