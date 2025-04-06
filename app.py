from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)
EXCEL_PATH = "data.xlsx"

@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    name = ""
    if request.method == "POST":
        name = request.form.get("name")
        if name:
            df = pd.read_excel(EXCEL_PATH, sheet_name="整合")
            filtered = df[df["姓名"] == name]
            if not filtered.empty:
                # 將 NaN 轉成空字串，避免前端出現 nan
                filtered = filtered.fillna("")
                results = filtered[["類別", "日期&項目", "費用", "看護費", "車資"]].values.tolist()
    return render_template("index.html", results=results, name=name)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
