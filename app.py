from flask import Flask, render_template, request
import pandas as pd
import os

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
            filtered = df_integrate[df_integrate['姓名'] == name].copy()

            if filtered.empty:
                message = f'找不到個案姓名：{name}'
                return render_template('index.html', name=name, message=message)

            # 數字欄位處理
            num_cols = ['費用', '看護費', '車資']
            for col in num_cols:
                if col in filtered.columns:
                    filtered[col] = (
                        filtered[col]
                        .replace(['-', '－', '–', '', None], 0)
                        .fillna(0)
                        .apply(lambda x: 0 if str(x).strip() in ['-', '－', '–', ''] else x)
                    )
                    filtered[col] = pd.to_numeric(filtered[col], errors='coerce').fillna(0).astype(int)

            # 顯示欄位
            results = filtered[["類別", "日期&項目", "費用", "看護費", "車資"]].copy()

            # 讀取名冊資料，並轉為整數
            roster_filtered = df_roster[df_roster['姓名'] == name]
            if not roster_filtered.empty:
                row = roster_filtered.iloc[0]
                roster_info = {
                    '月費': int(row.get('月費', 0) or 0),
                    '補助款': int(row.get('補助款', 0) or 0),
                    '雜費': int(row.get('雜費', 0) or 0),
                    '積欠': int(row.get('積欠', 0) or 0),
                    '溢收': int(row.get('溢收', 0) or 0),
                    '合計': int(row.get('合計', 0) or 0),
                }

            # 雜費統計
            expense_summary = {}

            medical = filtered[filtered["類別"] == "醫療"]
            expense_summary["醫療"] = int(medical["費用"].fillna(0).sum())
            expense_summary["看護費"] = int(medical["看護費"].fillna(0).sum())
            expense_summary["車資"] = int(medical["車資"].fillna(0).sum())

            consumables = filtered[filtered["類別"] == "耗材"]
            expense_summary["耗材"] = int(consumables["費用"].fillna(0).sum())

            others = filtered[filtered["類別"] == "其他"]
            expense_summary["其他"] = int(others["費用"].fillna(0).sum())

            farmers = filtered[filtered["類別"] == "農會"]
            expense_summary["農會購物"] = int(farmers["費用"].fillna(0).sum())

            refund = filtered[filtered["類別"] == "沖銷"]
            expense_summary["退費"] = int(refund["費用"].fillna(0).sum())

            # 雜費小計與總計
# 雜費小計與總計（包含車資）
expense_summary['雜費小計'] = int(
    expense_summary["醫療"]
    + expense_summary["看護費"]
    + expense_summary["車資"]
    + expense_summary["耗材"]
    + expense_summary["其他"]
    + expense_summary["農會購物"]
)

            expense_summary['雜費總計'] = expense_summary['雜費小計'] + expense_summary.get("退費", 0)

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
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)

