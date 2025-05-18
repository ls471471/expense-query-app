from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

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
            # 讀取 Excel 分頁
            df_integrate = pd.read_excel(EXCEL_FILE, sheet_name='整合')
            df_roster = pd.read_excel(EXCEL_FILE, sheet_name='名冊')

            # 篩選整合表中該姓名的資料
            results = df_integrate[df_integrate['姓名'] == name]

            if results.empty:
                message = f'找不到個案姓名：{name}'
                return render_template('index.html', name=name, message=message)

            # 數字欄位處理：將 '-' 轉為 0，轉型成 int
            num_cols = ['費用', '看護費', '車資']
            for col in num_cols:
                if col in results.columns:
                    results[col] = results[col].astype(str).replace('-', '0')
                    results[col] = pd.to_numeric(results[col], errors='coerce').fillna(0).astype(int)

            # 讀取名冊資料
            roster_filtered = df_roster[df_roster['姓名'] == name]
            if not roster_filtered.empty:
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

            # 計算雜費統計
            expense_summary = {}
            expense_categories = ['醫療', '看護費', '車資', '耗材', '其他', '農會購物', '退費']

            for cat in expense_categories:
                if cat == '退費':
                    expense_summary[cat] = results.loc[results['類別'] == '沖銷', '費用'].sum()
                else:
                    expense_summary[cat] = results.loc[results['類別'] == cat, '費用'].sum()

            expense_summary['雜費小計'] = sum(expense_summary[cat] for cat in expense_categories if cat != '退費')
            expense_summary['雜費總計'] = expense_summary['雜費小計'] + expense_summary.get('退費', 0)

            custom_heading = '本月'

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

