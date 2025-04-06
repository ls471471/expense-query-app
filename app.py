from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

EXCEL_PATH = "data.xlsx"  # 確保這個檔案已經推上 GitHub 並存在根目錄

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
                results = filtered[["類別", "日期&項目", "費用", "看護費", "車資"]].values.tolist()
    return render_template("index.html", results=results, name=name)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))  # Render 會給你一個環境變數叫 PORT
    app.run(host="0.0.0.0", port=port)
