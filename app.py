from flask import Flask, render_template, request
import pandas as pd

app = Flask(__name__)

EXCEL_PATH = "data.xlsx"

@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    if request.method == "POST":
        name = request.form.get("name")
        if name:
            df = pd.read_excel(EXCEL_PATH, sheet_name="整合")
            filtered = df[df["姓名"] == name]
            if not filtered.empty:
                results = filtered[["日期", "項目", "數量", "單價", "單位", "金額", "備註"]].values.tolist()
    return render_template("index.html", results=results)

if __name__ == "__main__":
    app.run()

