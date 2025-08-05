from flask import Flask, request, send_file, render_template_string
from lxml import etree
from openpyxl import Workbook
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

html_form = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Convert XML to XLSX</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      font-size: 2em;
      padding: 40px;
      background-color: #f9f9f9;
    }
    h1 {
      margin-bottom: 30px;
    }
    label {
      font-weight: bold;
      color: #004d00;
    }
    input[type="file"] {
      display: block;
      margin-bottom: 20px;
      font-size: 2em;
      font-weight: bold;
      color: #004d00;
      padding: 10px;
    }
    input[type="submit"] {
      display: block;
      margin-top: 10px;
      background-color: transparent;
      border: 4px solid #cc0066;
      color: #cc0066;
      font-weight: bold;
      padding: 20px 40px;
      border-radius: 16px;
      cursor: pointer;
      transition: background-color 0.2s ease;
    }
    input[type="submit"]:hover {
      background-color: #ffe6f0;
    }
  </style>
</head>
<body>
  <h1>Chọn file XML cần convert</h1>
  <form method="post" enctype="multipart/form-data">
    <label for="xmlfile">File nguồn:</label><br>
    <input type="file" name="xmlfile" id="xmlfile"><br>
    <input type="submit" value="Convert">
  </form>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def upload_convert():
    if request.method == "POST":
        file = request.files["xmlfile"]
        tree = etree.parse(file)
        root = tree.getroot()

        # Tạo file Excel (xlsx)
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # Namespace
        namespace = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
        rows = root.findall(".//ss:Row", namespaces=namespace)

        for row_idx, row in enumerate(rows, start=1):
            col_idx = 1
            for cell in row.findall(".//ss:Cell", namespaces=namespace):
                data = cell.find(".//ss:Data", namespaces=namespace)
                if data is not None:
                    ws.cell(row=row_idx, column=col_idx, value=data.text)
                col_idx += 1

        # Ghi ra bộ nhớ
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(output, download_name="converted.xlsx", as_attachment=True)

    return render_template_string(html_form)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
