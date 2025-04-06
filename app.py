from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# 載入 Excel 各分頁資料
df_detail = pd.read_excel("data.xlsx", sheet_name="整合")
df_roster = pd.read_excel("data.xlsx", sheet_name="名冊")

@app.route("/", methods=["GET", "POST"])
def index():
    results = None
    roster_info = None
    message = None

    if request.method == "POST":
        name = request.form["name"].strip()  # 去除前後空白
        filtered = df_detail[df_detail["姓名"] == name]

        # 處理費用明細資料
        if not filtered.empty:
            results = filtered[["類別", "項目", "費用", "看護費", "車資"]]
            results = results.fillna("-")

        # 處理名冊資料（取第一筆符合者）
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

        # 如果兩邊都找不到
        if results is None and roster_info is None:
            message = f"查無姓名「{name}」的資料，請確認輸入正確。"

    # 除錯用：Log 顯示目前變數內容
    print("🟡 results =", results)
    print("🟢 roster_info =", roster_info)

    return render_template("index.html", results=results, roster_info=roster_info, message=message)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
