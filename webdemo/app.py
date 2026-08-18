import os
import sys
import time
import traceback

from flask import Flask, request, render_template_string, send_file, redirect, url_for

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "scripts"))
import reconciliar_alas as core  # noqa: E402

UPLOAD_DIR = os.path.join(BASE, "uploads")
OUT_DIR = os.path.join(BASE, "salida")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

app = Flask(__name__)

FORM_HTML = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Conciliación cobros Alas — demo</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 640px; margin: 60px auto; color: #222; }
    h1 { font-size: 22px; }
    .box { border: 1px solid #ddd; border-radius: 10px; padding: 24px; }
    input[type=text] { width: 100%; padding: 8px; margin-bottom: 16px; box-sizing: border-box; }
    input[type=file] { margin-bottom: 16px; }
    button { background: #4472C4; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 15px; cursor: pointer; }
    .error { color: #b00020; white-space: pre-wrap; }
    .note { color: #666; font-size: 13px; margin-top: 20px; }
  </style>
</head>
<body>
  <h1>Conciliación mensual de cobros Alas — demo de prueba</h1>
  <div class="box">
    <form action="/upload" method="post" enctype="multipart/form-data">
      <label>Mes (para nombrar los archivos, ej: Julio 2026)</label>
      <input type="text" name="mes" value="Julio 2026" required>
      <label>Archivo de Flapp (export directo del sistema, o ya cruzado con "Cobrado x Alas")</label>
      <input type="file" name="archivo" accept=".csv,.xlsx,.xls" required>
      <label>Archivo crudo de Alas (opcional -- solo si subiste arriba el export de Flapp SIN cruzar)</label>
      <input type="file" name="archivo_alas" accept=".csv,.xlsx,.xls">
      <button type="submit">Conciliar</button>
    </form>
  </div>
  {% if error %}<p class="error">{{ error }}</p>{% endif %}
  <p class="note">Demo temporal corriendo en el sandbox de Claude — solo para probar que la lógica
  funciona con tu archivo real. No uses esto para el proceso mensual definitivo.<br>
  Si subes solo el primer archivo, debe venir ya cruzado por finanzas (con columna "Cobrado x Alas").
  Si subes los dos, el cruce se hace automáticamente por shipmentId, sin que finanzas tenga que
  hacerlo a mano.</p>
</body>
</html>
"""

RESULT_HTML = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Resultado — Conciliación Alas</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 720px; margin: 60px auto; color: #222; }
    h1 { font-size: 22px; }
    .box { border: 1px solid #ddd; border-radius: 10px; padding: 24px; margin-bottom: 20px; }
    table { border-collapse: collapse; width: 100%; margin: 16px 0; }
    td, th { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 14px; }
    th { background: #4472C4; color: white; }
    a.button { display: inline-block; background: #4472C4; color: white; text-decoration: none;
               padding: 10px 20px; border-radius: 6px; margin-right: 10px; }
    pre { white-space: pre-wrap; background: #f7f7f7; padding: 16px; border-radius: 8px; font-size: 13px; }
    a.back { display:block; margin-top: 20px; }
  </style>
</head>
<body>
  <h1>Conciliación de {{ mes }} — lista</h1>
  <div class="box">
    <table>
      <tr><th>Categoría</th><th>Cantidad</th></tr>
      {% for cat, n in counts.items() %}
      <tr><td>{{ cat }}</td><td>{{ n }}</td></tr>
      {% endfor %}
    </table>
    <a class="button" href="{{ url_for('download', slug=slug, kind='xlsx') }}">Descargar Excel</a>
    <a class="button" href="{{ url_for('download', slug=slug, kind='md') }}">Descargar borrador de email</a>
  </div>
  <div class="box">
    <h2 style="font-size:16px;">Vista previa del borrador de email</h2>
    <pre>{{ email_preview }}</pre>
  </div>
  <a class="back" href="/">&larr; Probar con otro archivo</a>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(FORM_HTML, error=None)


@app.post("/upload")
def upload():
    mes = request.form.get("mes", "").strip() or "Sin_Mes"
    f = request.files.get("archivo")
    f_alas = request.files.get("archivo_alas")
    if not f or f.filename == "":
        return render_template_string(FORM_HTML, error="No se recibió ningún archivo.")

    ext = os.path.splitext(f.filename)[1].lower() or ".csv"
    slug = "".join(c if c.isalnum() else "_" for c in mes) + f"_{int(time.time())}"
    in_path = os.path.join(UPLOAD_DIR, slug + ext)
    f.save(in_path)

    alas_path = None
    if f_alas and f_alas.filename:
        ext_alas = os.path.splitext(f_alas.filename)[1].lower() or ".csv"
        alas_path = os.path.join(UPLOAD_DIR, slug + "_alas" + ext_alas)
        f_alas.save(alas_path)

    out_dir = os.path.join(OUT_DIR, slug)
    os.makedirs(out_dir, exist_ok=True)

    try:
        if alas_path:
            resultados = core.procesar_dos_archivos(in_path, alas_path)
        else:
            resultados = core.procesar(in_path)
        xlsx_path = os.path.join(out_dir, "Conciliacion.xlsx")
        df, disputas, revisar, desactualizadas, sin_comuna = core.construir_excel(resultados, mes, xlsx_path)
        email_txt = core.construir_email(mes, disputas, revisar, desactualizadas, sin_comuna)
        md_path = os.path.join(out_dir, "Borrador_Email.md")
        with open(md_path, "w", encoding="utf-8") as fh:
            fh.write(email_txt)
    except SystemExit as e:
        return render_template_string(FORM_HTML, error=str(e))
    except Exception:
        return render_template_string(FORM_HTML, error="Error procesando el archivo:\n" + traceback.format_exc())

    from collections import Counter
    counts = Counter(r["categoria"] for r in resultados)

    return render_template_string(
        RESULT_HTML,
        mes=mes,
        slug=slug,
        counts=counts,
        email_preview=email_txt,
    )


@app.get("/download/<slug>/<kind>")
def download(slug, kind):
    out_dir = os.path.join(OUT_DIR, slug)
    if kind == "xlsx":
        return send_file(os.path.join(out_dir, "Conciliacion.xlsx"), as_attachment=True,
                          download_name=f"Conciliacion_Alas_{slug}.xlsx")
    elif kind == "md":
        return send_file(os.path.join(out_dir, "Borrador_Email.md"), as_attachment=True,
                          download_name=f"Borrador_Email_Alas_{slug}.md")
    return "no encontrado", 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
