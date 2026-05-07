from flask import Flask, render_template_string, request, send_file
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Inches, RGBColor, Pt
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Rapport intervention</title>

    <style>
        body{
            font-family: Arial;
            background:#f4f6f9;
            padding:40px;
        }

        .box{
            background:white;
            padding:30px;
            border-radius:12px;
            max-width:500px;
            margin:auto;
            box-shadow:0 0 10px rgba(0,0,0,0.1);
        }

        input{
            width:100%;
            padding:10px;
            margin-top:5px;
            margin-bottom:20px;
        }

        button{
            background:#0066cc;
            color:white;
            border:none;
            padding:12px 20px;
            border-radius:8px;
            cursor:pointer;
        }

        button:hover{
            background:#004999;
        }
    </style>
</head>

<body>

<div class="box">

<h2>Rapport d'intervention</h2>

<form method="POST" enctype="multipart/form-data">

  Client :
  <input name="client">

  Technicien :
  <input name="technicien">

  Intervention :
  <input name="intervention">

  Photo :
  <input type="file" name="photo" accept="image/*">

  <button type="submit">
      Générer rapport
  </button>

</form>

</div>

</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        client = request.form.get("client", "")
        technicien = request.form.get("technicien", "")
        intervention = request.form.get("intervention", "")

        photo = request.files.get("photo")

        # =========================
        # DOCUMENT WORD
        # =========================

        doc = Document()

        # TITRE
        title = doc.add_heading(level=1)

        run = title.add_run("Rapport d'intervention")
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 102, 204)

        doc.add_paragraph("")

        # TABLEAU
        table = doc.add_table(rows=3, cols=2)
        table.style = "Table Grid"

        table.cell(0,0).text = "Client"
        table.cell(0,1).text = client

        table.cell(1,0).text = "Technicien"
        table.cell(1,1).text = technicien

        table.cell(2,0).text = "Intervention"
        table.cell(2,1).text = intervention

        doc.add_paragraph("")

        # PHOTO
        if photo and photo.filename:

            filename = secure_filename(photo.filename)

            photo_path = os.path.join(
                UPLOAD_FOLDER,
                filename
            )

            photo.save(photo_path)

            p = doc.add_paragraph()
            r = p.add_run("Photo intervention")
            r.bold = True

            doc.add_picture(photo_path, width=Inches(4))

        # NOM FICHIER
        fichier = "rapport.docx"

        doc.save(fichier)

        # TÉLÉCHARGEMENT AUTO
        return send_file(
            fichier,
            as_attachment=True
        )

    return render_template_string(HTML)


if __name__ == "__main__":

    port = int(os.environ.get("PORT", 10000))

    app.run(
        host="0.0.0.0",
        port=port
    )
