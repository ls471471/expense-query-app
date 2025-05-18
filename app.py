from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# 載入 Excel
df_detail = pd.read_excel("data.xlsx", sheet_name="整合")
df_roster = pd.read_excel("data.xlsx", sheet_name="名冊")

df_detail.columns = df_detail.columns.str.strip()
df_roster.columns = df_roster.columns.str.strip()
df_detail["姓名"] = df_detail["姓名"].astype(str).str.strip()
df_roster["姓名"] = df_roster["姓名"].astype(str).str.strip()

def safe_int_or_dash(x):
    try:
        return int(float(x))
    except (ValueError, TypeError):
        return "-"

@app.template_filter("safe_int")
def safe_int(value):
    try:
        return int(value)
    except (ValueError, TypeError):
        return "-"

@app.route("/", methods=["GET", "POST"])
def index():
    results = None
    roster_info = None
    message = None
    custom_heading = ""
    name = ""
    expense_summary = None

    if request.method == "POST":
        name = str(request.form["name"]).strip()
        filtered = df_detail[df_detail["姓名"] == name]

        if not filtered.empty:
            num_cols = ["費用", "看護費", "車資"]
            filtered[num_cols] = filtered[num_cols].fillna(0)
            results = filtered[["類別", "日期&項目"] + num_cols].copy()
            results[num_cols] = results[num_cols].astype(int)

            # 雜費統計
            expense_summary = {}
            expense_summary["醫療"] = int(filtered[filtered["類別"] == "醫療"]["費用"].fillna(0).sum())
            expense_summary["看護費"] = int(filtered[filtered["類別"] == "醫療"]["看護費"].fillna(0).sum())
            expense_summary["車資"] = int(filtered[filtered["類別"] == "醫療"]["車資"].fillna(0).sum())
            expense_summary["耗材"] = int(filtered[filtered["類別"] == "耗材"]["費用"].fillna(0).sum())
            expense_summary["其他"] = int(filtered[filtered["類別"] == "其他"]["費用"].fillna(0).sum())
            expense_summary["農會購物"] = int(filtered[filtered["類別"] == "農會"]["費用"].fillna(0).sum())

            subtotal = sum(expense_summary[key] for key in ["醫療", "看護費", "車資", "耗材", "其他", "農會購物"])
            expense_summary["雜費小計"] = subtotal

            refund = abs(filtered[filtered["類別"] == "沖銷"]["費用"].fillna(0).sum())
            expense_summary["退費"] = int(refund)

            expense_summary["雜費總計"] = int(subtotal - refund)

        filtered_roster = df_roster[df_roster["姓名"] == name]
        if not filtered_roster.empty:
            row = filtered_roster.iloc[0]
            roster_info = {
                "月費": safe_int_or_dash(row.get("月費")),
                "補助款": safe_int_or_dash(row.get("補助款")),
                "雜費": safe_int_or_dash(row.get("雜費")),
                "積欠": safe_int_or_dash(row.get("積欠")),
                "溢收": safe_int_or_dash(row.get("溢收")),
                "合計": safe_int_or_dash(row.get("合計")),
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
    app.run(host="0.0.0.0", port=port, debug=True)  # 加 debug=True 方便除錯


