from flask import Flask, request, send_file, render_template_string
import fitz  # PyMuPDF
import json
import re
import os
import xlsxwriter
from io import BytesIO

app = Flask(__name__)

# HTML form to upload PDF
HTML_FORM = '''
<!DOCTYPE html>
<html>
<head>
    <title>PDF to Excel Converter</title>
</head>
<body>
    <h2>Upload PDF Statement</h2>
    <form method="POST" action="/convert" enctype="multipart/form-data">
        <input type="file" name="pdf_file" accept=".pdf" required><br><br>
        <button type="submit">Convert and Download Excel</button>
    </form>
</body>
</html>
'''

@app.route("/")
def index():
    return render_template_string(HTML_FORM)

@app.route("/convert", methods=["POST"])
def convert_pdf_to_excel():
    uploaded_file = request.files["pdf_file"]
    if not uploaded_file:
        return "No file uploaded"

    pdf_bytes = uploaded_file.read()
    doc = fitz.open("pdf", pdf_bytes)

    lines = []
    for page in doc:
        lines.extend(page.get_text().split("\n"))

    transactions = []
    i = 0
    date_pattern = re.compile(r"\d{2}-\d{2}-\d{4}")
    amount_pattern = re.compile(r"^\d{1,3}(?:,\d{3})*(?:\.\d{2})$")
    prev_balance = None

    while i < len(lines):
        line = lines[i].strip()
        if date_pattern.match(line):
            date = line
            i += 1
            particulars = []

            while i < len(lines) and not lines[i].startswith("Chq:") and not amount_pattern.match(lines[i].strip()):
                particulars.append(lines[i].strip())
                i += 1

            if i < len(lines) and lines[i].startswith("Chq:"):
                i += 1

            amounts = []
            while i < len(lines) and len(amounts) < 2:
                if amount_pattern.match(lines[i].strip()):
                    amounts.append(lines[i].strip())
                i += 1

            deposit, withdrawal, balance = "", "", ""
            if len(amounts) == 2:
                balance = amounts[1]
                b = float(balance.replace(",", ""))
                a = float(amounts[0].replace(",", ""))
                if prev_balance is not None:
                    if b > prev_balance:
                        deposit = amounts[0]
                    else:
                        withdrawal = amounts[0]
                else:
                    deposit = amounts[0]
                prev_balance = b
            elif len(amounts) == 1:
                balance = amounts[0]
                prev_balance = float(balance.replace(",", ""))

            transactions.append({
                "Date": date,
                "Particulars": " ".join(particulars),
                "Deposit": deposit,
                "Withdrawal": withdrawal,
                "Balance": balance
            })
        else:
            i += 1

    # Create Excel in memory
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Statement")

    headers = ["Date", "Particulars", "Deposit", "Withdrawal", "Balance"]
    for col, head in enumerate(headers):
        worksheet.write(0, col, head)

    for row, txn in enumerate(transactions, start=1):
        worksheet.write(row, 0, txn["Date"])
        worksheet.write(row, 1, txn["Particulars"])
        worksheet.write(row, 2, txn["Deposit"])
        worksheet.write(row, 3, txn["Withdrawal"])
        worksheet.write(row, 4, txn["Balance"])

    workbook.close()
    output.seek(0)

   
    return send_file(output, as_attachment=True, download_name="converted_statement.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=10000)
