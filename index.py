from flask import Flask, request, send_file, render_template_string
import fitz  # PyMuPDF
import json
import re
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from io import BytesIO

app = Flask(__name__)
HTML_FORM = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>PDF to Excel Converter</title>
    <style>
        body {
            background: #f0f4f8;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
        }
        .container {
            max-width: 450px;
            background: white;
            margin: 80px auto;
            padding: 30px 40px;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.1);
        }
        h2 {
            margin-top: 0;
            color: #2c3e50;
            text-align: center;
        }
        input[type="file"] {
            width: 100%;
            padding: 12px;
            margin: 18px 0;
            border: 2px dashed #ccc;
            border-radius: 8px;
            background-color: #f9f9f9;
            cursor: pointer;
            transition: border-color 0.3s;
        }
        input[type="file"]:hover {
            border-color: #3498db;
        }
        button {
            width: 100%;
            padding: 12px;
            background-color: #3498db;
            color: white;
            font-size: 16px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            transition: background 0.3s;
        }
        button:hover {
            background-color: #2980b9;
        }
        .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 13px;
            color: #777;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>📄 PDF to Excel Converter</h2>
        <form method="POST" action="/convert" enctype="multipart/form-data">
            <input type="file" name="pdf_file" accept=".pdf" required>
            <button type="submit">Convert and Download</button>
        </form>
        <div class="footer">Built with ❤️ for your bank statements</div>
    </div>
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
                "date": date,
                "particulars": " ".join(particulars),
                "deposit": deposit,
                "withdrawal": withdrawal,
                "balance": balance
            })
        else:
            i += 1

    # Write to Excel using openpyxl
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bank Transactions"

    headers = ["Date", "Particulars", "Deposit", "Withdrawal", "Balance"]
    ws.append(headers)

    for tx in transactions:
        ws.append([
            tx["date"],
            tx["particulars"],
            tx["deposit"],
            tx["withdrawal"],
            tx["balance"]
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_len + 4

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="converted_statement.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True, port=10000)
