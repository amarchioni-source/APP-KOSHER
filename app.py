import hmac
import io
import os
import secrets
import time
import zipfile

import pandas as pd
from flask import Flask, Response, abort, render_template, request, send_file

from report_builder import (DESTINO_DEFECTO, cargar_barcodes, generar_reportes,
                            parsear_barcodes_extra)

BASE = os.path.dirname(os.path.abspath(__file__))
MODELO = os.path.join(BASE, "modelo", "modelo.xlsx")
BARCODES = os.path.join(BASE, "barcodes.csv")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024

_CACHE = {}          # token -> {"ts": float, "files": [(nombre, bytes)]}   (se usa 1 solo worker)
_TTL = 30 * 60


def _limpiar():
    now = time.time()
    for t in [t for t, v in _CACHE.items() if now - v["ts"] > _TTL]:
        _CACHE.pop(t, None)


@app.before_request
def _auth():
    """Si se define APP_PASSWORD, se pide contraseña (cualquier usuario)."""
    pwd = os.environ.get("APP_PASSWORD")
    if not pwd or request.path == "/healthz":
        return None
    a = request.authorization
    if a and a.password and hmac.compare_digest(a.password, pwd):
        return None
    return Response("Acceso restringido", 401, {"WWW-Authenticate": 'Basic realm="Reportes de produccion"'})


@app.get("/healthz")
def healthz():
    return "ok"


@app.get("/")
def index():
    return render_template("index.html", destino=DESTINO_DEFECTO, resultados=None)


@app.post("/generar")
def generar():
    archivos = [f for f in request.files.getlist("archivos") if f and f.filename]
    destino = (request.form.get("destino") or DESTINO_DEFECTO).strip()
    if not archivos:
        return render_template("index.html", destino=destino, error="Seleccione al menos un archivo .xlsx."), 400
    barcodes = cargar_barcodes(BARCODES)
    barcodes.update(parsear_barcodes_extra(request.form.get("barcodes_extra")))

    resultados, avisos, vistos = [], [], set()
    for f in archivos:
        try:
            df = pd.read_excel(f.stream)
            reps, av = generar_reportes(df, barcodes, MODELO, destino)
        except Exception as e:                                   # mensaje claro al usuario
            avisos.append(f"{f.filename}: {e}")
            continue
        avisos += [f"{f.filename}: {a}" for a in av]
        for r in reps:
            if r["archivo"] in vistos:
                avisos.append(f"{f.filename}: la fecha {r['fecha']} ya estaba en otro archivo; se omitió.")
                continue
            vistos.add(r["archivo"])
            resultados.append(r)

    token = None
    if resultados:
        _limpiar()
        token = secrets.token_urlsafe(12)
        _CACHE[token] = {"ts": time.time(), "files": [(r["archivo"], r["bytes"]) for r in resultados]}
    return render_template("index.html", destino=destino, resultados=resultados, avisos=avisos, token=token)


@app.get("/descargar/<token>/<int:i>")
def descargar(token, i):
    item = _CACHE.get(token)
    if not item or i >= len(item["files"]):
        abort(404)
    nombre, data = item["files"][i]
    return send_file(io.BytesIO(data), as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/descargar-zip/<token>")
def descargar_zip(token):
    item = _CACHE.get(token)
    if not item:
        abort(404)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, data in item["files"]:
            z.writestr(nombre, data)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="reportes_produccion.zip", mimetype="application/zip")


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
