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
      margin: 0;
      padding: 0;
      background: url('https://images.unsplash.com/photo-1531746790731-6c087fecd65a?auto=format&fit=crop&w=1600&q=80') no-repeat center center fixed;
      background-size: cover;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
    }
    .container {
      background: rgba(255, 255, 255, 0.95);
      padding: 40px;
      border-radius: 15px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.2);
      width: 100%;
      max-width: 450px;
      text-align: center;
    }
    h2 {
      color: #2c3e50;
      margin-bottom: 20px;
    }
    input[type="file"] {
      width: 100%;
      padding: 14px;
      margin: 20px 0;
      border: 2px dashed #ccc;
      border-radius: 10px;
      background-color: #f8f8f8;
      cursor: pointer;
    }
    input[type="file"]:hover {
      border-color: #3498db;
    }
    button {
      width: 100%;
      padding: 14px;
      background-color: #3498db;
      color: white;
      font-size: 16px;
      border: none;
      border-radius: 10px;
      cursor: pointer;
      transition: background 0.3s ease;
    }
    button:hover {
      background-color: #2980b9;
    }
    .footer {
      margin-top: 20px;
      font-size: 13px;
      color: #555;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📄 PDF to Excel Converter</h2>
    <form method="POST" action="/convert" enctype="multipart/form-data">
      <input type="file" name="pdf_file" accept=".pdf" required>
      <button type="submit">Convert & Download</button>
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
                amt_line = lines[i].strip()
                if amount_pattern.match(amt_line):
                    amounts.append(amt_line)
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

    # Create Excel
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

    # Auto-adjust column width
    for col_idx, col_cells in enumerate(ws.columns, start=1):
        max_length = 0
        for cell in col_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 4

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
    app.run(debug=True, host="0.0.0.0", port=10000)
