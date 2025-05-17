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
    summary = None
    subtotal = 0
    refund = 0
    total_misc = 0
    custom_heading = ""
    name = ""
    expense_summary = None

    if request.method == "POST":
        name = request.form["name"].strip()  # 去除前後空白
        filtered = df_detail[df_detail["姓名"] == name]

        if not filtered.empty:
            results = filtered[["類別", "日期&項目", "費用", "看護費", "車資"]]
            results = results.fillna("-")

            # 統計各類別金額
            categories = ["醫療費", "看護費", "車資", "耗材", "其他", "農會購物"]
            summary = {}
            for cat in categories:
                if cat == "看護費":
                    summary[cat] = filtered["看護費"].fillna(0).sum()
                elif cat == "車資":
                    summary[cat] = filtered["車資"].fillna(0).sum()
                else:
                    summary[cat] = filtered.loc[filtered["類別"] == cat, "費用"].fillna(0).sum()

            subtotal = sum(summary.values())
            refund = filtered.loc[filtered["類別"] == "沖銷(退費)", "費用"].fillna(0).sum()
            total_misc = subtotal - refund

            # 將統計結果包裝成 expense_summary 傳到前端
            expense_summary = {
                "醫療費": int(summary.get("醫療費", 0)),
                "看護費": int(summary.get("看護費", 0)),
                "車資": int(summary.get("車資", 0)),
                "耗材": int(summary.get("耗材", 0)),
                "其他": int(summary.get("其他", 0)),
                "農會購物": int(summary.get("農會購物", 0)),
                "雜費小計": int(subtotal),
                "退費": int(refund),
                "雜費總計": int(total_misc)
            }

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
