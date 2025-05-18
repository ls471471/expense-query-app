from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

# 載入 Excel 各分頁資料
df_detail = pd.read_excel("data.xlsx", sheet_name="整合")
df_roster = pd.read_excel("data.xlsx", sheet_name="名冊")

# 移除欄位名稱空白，統一姓名欄位格式為字串（防止數字型別）
df_detail.columns = df_detail.columns.str.strip()
df_roster.columns = df_roster.columns.str.strip()
df_detail["姓名"] = df_detail["姓名"].astype(str).str.strip()
df_roster["姓名"] = df_roster["姓名"].astype(str).str.strip()

# 安全轉換函數（無法轉換時回傳 "-"）
def safe_int_or_dash(x):
    try:
        return int(float(x))
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
            for col in num_cols:
                results[col] = results[col].apply(safe_int_or_dash)

            # 統計雜費金額（只計算可以加總的數字）
            expense_summary = {}

            medical = filtered[filtered["類別"] == "醫療"]
            expense_summary["醫療"] = medical["費用"].fillna(0).sum()
            expense_summary["看護費"] = medical["看護費"].fillna(0).sum()
            expense_summary["車資"] = medical["車資"].fillna(0).sum()

            consumables = filtered[filtered["類別"] == "耗材"]
            expense_summary["耗材"] = consumables["費用"].fillna(0).sum()

            others = filtered[filtered["類別"] == "其他"]
            expense_summary["其他"] = others["費用"].fillna(0).sum()

            farmers = filtered[filtered["類別"] == "農會"]
            expense_summary["農會購物"] = farmers["費用"].fillna(0).sum()

            subtotal = sum([
                expense_summary["醫療"],
                expense_summary["看護費"],
                expense_summary["車資"],
                expense_summary["耗材"],
                expense_summary["其他"],
                expense_summary["農會購物"]
            ])
            expense_summary["雜費小計"] = subtotal

            refund_df = filtered[filtered["類別"] == "沖銷"]
            refund = abs(refund_df["費用"].fillna(0).sum())
            expense_summary["退費"] = refund

            total_misc = subtotal - refund
            expense_summary["雜費總計"] = total_misc

            for key in expense_summary:
                expense_summary[key] = int(round(expense_summary[key]))

        # 查詢名冊資料
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
    app.run(host="0.0.0.0", port=port)

