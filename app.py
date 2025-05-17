from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# 載入 Excel 各分頁資料
df_detail = pd.read_excel("data.xlsx", sheet_name="整合")
df_roster = pd.read_excel("data.xlsx", sheet_name="名冊")
df_roster.columns = df_roster.columns.str.strip()  # 移除欄位名稱空白

@app.route("/", methods=["GET", "POST"])
def index():
    results = None
    roster_info = None
    message = None
    custom_heading = ""
    name = ""
    expense_summary = None

    if request.method == "POST":
        name = request.form["name"].strip()
        filtered = df_detail[df_detail["姓名"] == name]

        if not filtered.empty:
            results = filtered[["類別", "日期&項目", "費用", "看護費", "車資"]]
            results = results.fillna("-")

            # 統計雜費：依照你的規則處理
            expense_summary = {}

            # 類別為「醫療」
            medical = filtered[filtered["類別"] == "醫療"]
            expense_summary["醫療"] = medical["費用"].fillna(0).sum()
            expense_summary["看護費"] = medical["看護費"].fillna(0).sum()
            expense_summary["車資"] = medical["車資"].fillna(0).sum()

            # 類別為「耗材」
            consumables = filtered[filtered["類別"] == "耗材"]
            expense_summary["耗材"] = consumables["費用"].fillna(0).sum()

            # 類別為「其他」
            others = filtered[filtered["類別"] == "其他"]
            expense_summary["其他"] = others["費用"].fillna(0).sum()

            # 類別為「農會」
            farmers = filtered[filtered["類別"] == "農會"]
            expense_summary["農會購物"] = farmers["費用"].fillna(0).sum()

            # 雜費小計（六項加總）
            subtotal = sum([
                expense_summary["醫療"],
                expense_summary["看護費"],
                expense_summary["車資"],
                expense_summary["耗材"],
                expense_summary["其他"],
                expense_summary["農會購物"]
            ])
            expense_summary["雜費小計"] = subtotal

            # 類別為「沖銷」的退費
            refund = filtered[filtered["類別"] == "沖銷"]["費用"].fillna(0).sum()
            expense_summary["退費"] = refund

            # 雜費總計
            total_misc = subtotal - refund
            expense_summary["雜費總計"] = total_misc

            # 四捨五入轉為整數
            for key in expense_summary:
                expense_summary[key] = int(round(expense_summary[key]))

        # 抓名冊資料
        filtered_roster = df_roster[df_roster["姓名"] == name]
        if not filtered_roster.empty:
            row = filtered_roster.iloc[0]
            roster_info = {
                "月費": row.get("月費", "-"),
                "補助款": row.get("補助款", "-"),
                "雜費": row.get("雜費", "-"),
                "積欠": row.get("積欠", "-"),
                "溢收": row.get("溢收", "-"),
                "合計": row.get("合計", "-")
            }
            custom_heading = row.get("月份", "")

        if results is None and roster_info is None:
            message = f"查無姓名「{name}」的資料，請確認輸入正確。"

    return render_template("index.html",
                           name=name,
                           results=results,
                           roster_info=roster_info,
                           message=message,
                           expense_summary=expense_summary,
                           custom_heading=custom_heading)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

