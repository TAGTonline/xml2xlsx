from flask import Flask, request, send_file, render_template_string
from lxml import etree
import xlwt
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

html_form = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Convert XML to XLS</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      font-size: 2em; /* Font to gấp đôi */
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
      font-size: 2em;         /* Tăng kích thước */
      font-weight: bold;      /* Chữ in đậm */
      color: #004d00;         /* Màu xanh đậm */
      padding: 10px;
    }

    input[type="submit"] {
      display: block;
      margin-top: 10px;
      background-color: transparent;
      border: 4px solid #cc0066; /* Hồng đậm */
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

        # Tạo workbook Excel
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Sheet1")

        # Tìm tất cả rows (dựa vào namespace SpreadsheetML)
        namespace = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
        rows = root.findall(".//ss:Row", namespaces=namespace)

        for row_idx, row in enumerate(rows):
            col_idx = 0
            for cell in row.findall(".//ss:Cell", namespaces=namespace):
                data = cell.find(".//ss:Data", namespaces=namespace)
                if data is not None:
                    ws.write(row_idx, col_idx, data.text)
                col_idx += 1

        # Lưu ra bộ nhớ
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(output, download_name="converted.xls", as_attachment=True)

    return render_template_string(html_form)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
