import hmac
import io
import os
import secrets
import time
import zipfile

import pandas as pd
from flask import Flask, Response, abort, render_template, request, send_file
from werkzeug.exceptions import HTTPException

import report_israel as RI
from report_builder import (DESTINO_DEFECTO, cargar_barcodes, generar_reportes,
                            parsear_barcodes_extra)

BASE = os.path.dirname(os.path.abspath(__file__))


def _primero(*rutas):
    """Primera ruta que exista (acepta archivos en carpetas o sueltos en la raíz)."""
    for r in rutas:
        if os.path.exists(r):
            return r
    return rutas[0]


MODELO = _primero(os.path.join(BASE, "modelo", "modelo.xlsx"), os.path.join(BASE, "modelo.xlsx"))
MODELO_ISRAEL = _primero(os.path.join(BASE, "modelo", "modelo_israel.xlsx"), os.path.join(BASE, "modelo_israel.xlsx"))
BARCODES = os.path.join(BASE, "barcodes.csv")
INDICES_ISRAEL = os.path.join(BASE, "indices_israel.csv")
_TPL = os.path.join(BASE, "templates")
if not os.path.exists(os.path.join(_TPL, "index.html")):
    _TPL = BASE

app = Flask(__name__, template_folder=_TPL)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024

_CACHE = {}          # token -> {"ts": float, "files": [(nombre, bytes)]}   (se usa 1 solo worker)
_TTL = 30 * 60


# ------------------------------------------------------------------ utilidades
def _limpiar():
    now = time.time()
    for t in [t for t, v in _CACHE.items() if now - v["ts"] > _TTL]:
        _CACHE.pop(t, None)


def _n(x, dec=0):
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _plural(n, uno, varios):
    return uno if n == 1 else varios


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


@app.errorhandler(Exception)
def _error(e):
    """Muestra el error real en pantalla (la app ya está protegida por contraseña)."""
    if isinstance(e, HTTPException):
        return e
    app.logger.exception("Error no controlado")
    return Response(f"Error en la aplicación: {type(e).__name__}: {e}", 500, {"Content-Type": "text/plain; charset=utf-8"})


@app.get("/healthz")
def healthz():
    return "ok"


@app.get("/")
def index():
    return render_template("index.html", destino=DESTINO_DEFECTO, tipo="usa", resultados=None)


# ------------------------------------------------------------------ tarjetas de resultado
def _card_usa(r, idx):
    alarmas = []
    if r["sin_barcode"]:
        alarmas.append({"titulo": "Productos sin barcode (incluidos al final de la lista)",
                        "lineas": [{"cod": s["Codigo"], "texto": s["desc"],
                                    "extra": f"{s['cajas']:.0f} {_plural(s['cajas'], 'caja', 'cajas')}"}
                                   for s in r["sin_barcode"]]})
    return {"tipo": "usa", "fecha": r["fecha"], "idx": idx, "alarmas": alarmas, "notas": [], "pct": r["pct_spiced"],
            "etiqueta": f"Falta barcode ({len(r['sin_barcode'])})",
            "kpis": [("Productos", str(r["productos"])), ("Cajas", _n(r["cajas"])), ("Kg bruto", _n(r["bruto"], 3)),
                     ("Kg neto", _n(r["neto"], 3)), ("% Spiced", _n(r["pct_spiced"], 2) + " %")]}


def _card_israel(r, idx):
    alarmas, notas = [], []
    if r["faltan_indice"]:
        alarmas.append({"titulo": "Códigos sin índice NETO o barcode (celdas vacías en el reporte)",
                        "lineas": [{"cod": f["Codigo"], "texto": f["desc"],
                                    "extra": f"{f['cartones']} {_plural(f['cartones'], 'cartón', 'cartones')}"}
                                   for f in r["faltan_indice"]]})
    if r["sin_venc"]:
        alarmas.append({"titulo": "Cartones sin fecha de vencimiento (celda vacía)",
                        "lineas": [{"cod": "", "texto": f"{r['sin_venc']} {_plural(r['sin_venc'], 'cartón', 'cartones')} sin BEST BEFORE DATE",
                                    "extra": ""}],
                        "ayuda": RI.PLAZOS})
    if r["inferidos"]:
        notas.append("Índice y barcode inferidos del par GF AA (patrón observado, no confirmado): " +
                     ", ".join(f"{f['Codigo']} ({f['cartones']})" for f in r["inferidos"]) + ".")
    if r["sin_codbar"]:
        notas.append(f"{r['sin_codbar']} cartones sin CodBarS (celda de código vacía).")
    if r["sin_codbar_104"]:
        notas.append(f"{r['sin_codbar_104']} cartones en estado 104 sin CodBarS: se incluyen con la celda de código vacía.")
    return {"tipo": "israel", "fecha": r["fecha"], "idx": idx, "alarmas": alarmas, "notas": notas, "pct": None,
            "etiqueta": "Datos pendientes",
            "kpis": [("Productos", str(r["productos"])), ("Cartones", _n(r["cartones"])), ("Unidades", _n(r["unidades"])),
                     ("Kg bruto", _n(r["bruto"], 3)), ("Kg neto", _n(r["neto"], 3))]}


def _control_lineas(c):
    out = []
    if c["excl_desc"]:
        tot = int(sum(e["cartones"] for e in c["excl_desc"]))
        out.append(f"Líneas excluidas por descripción: {tot} cartones en {len(c['excl_desc'])} "
                   f"{_plural(len(c['excl_desc']), 'descripción', 'descripciones')}.")
    if c["s_cartones"]:
        out.append(f"Cartones en estado S (con salida) excluidos: {int(c['s_cartones'])}.")
    if c["estados_otros"]:
        out.append("Estados no contemplados, excluidos: " +
                   ", ".join(f"{e['estado'] or '(vacío)'} ({int(e['cartones'])})" for e in c["estados_otros"]) + ".")
    if c["dias_sin_archivo"]:
        out.append("Días hábiles sin producción dentro del período: " + ", ".join(c["dias_sin_archivo"]) + ".")
    return out


# ------------------------------------------------------------------ generación
@app.post("/generar")
def generar():
    tipo = request.form.get("tipo", "usa")
    tipo = tipo if tipo in ("usa", "israel", "ambos") else "usa"
    destino = (request.form.get("destino") or DESTINO_DEFECTO).strip()
    ctx = {"destino": destino, "tipo": tipo}
    archivos = [f for f in request.files.getlist("archivos") if f and f.filename]
    if not archivos:
        return render_template("index.html", error="Seleccione al menos un archivo .xlsx.", **ctx), 400

    avisos, dfs, hashes = [], [], {}
    for f in archivos:                       # lectura + archivos de contenido idéntico: se usa uno solo
        try:
            df = pd.read_excel(f.stream)
        except Exception as e:
            avisos.append(f"{f.filename}: no se pudo leer el archivo ({e}).")
            continue
        h = int(pd.util.hash_pandas_object(df.astype(str), index=False).sum())
        if h in hashes:
            avisos.append(f"{f.filename} tiene el mismo contenido que {hashes[h]}; se usó uno solo.")
            continue
        hashes[h] = f.filename
        dfs.append((f.filename, df))

    files, cards = [], []
    pend_usa, pend_israel, control_lineas, excl_desc, control_idx = [], [], [], [], None

    if tipo in ("usa", "ambos"):
        barcodes = cargar_barcodes(BARCODES)
        barcodes.update(parsear_barcodes_extra(request.form.get("barcodes_extra")))
        vistos = set()
        for nombre, df in dfs:
            try:
                reps, av = generar_reportes(df, barcodes, MODELO, destino)
            except Exception as e:
                avisos.append(f"USA · {nombre}: {e}")
                continue
            avisos += [f"USA · {nombre}: {a}" for a in av]
            for r in reps:
                if r["archivo"] in vistos:
                    avisos.append(f"USA · {nombre}: la fecha {r['fecha']} ya estaba en otro archivo; se omitió.")
                    continue
                vistos.add(r["archivo"])
                files.append((r["archivo"], r["bytes"]))
                cards.append(_card_usa(r, len(files) - 1))
                for s in r["sin_barcode"]:
                    if s["Codigo"] not in [p[0] for p in pend_usa]:
                        pend_usa.append((s["Codigo"], s["desc"]))

    if tipo in ("israel", "ambos"):
        try:
            indices = RI.cargar_indices(INDICES_ISRAEL)
            indices.update(RI.parsear_indices_extra(request.form.get("indices_extra")))
            normal = []
            for nombre, df in dfs:
                try:
                    nd, av = RI.normalizar_stock(df)
                    avisos += [f"Israel · {nombre}: {a}" for a in av]
                    normal.append(nd)
                except Exception as e:
                    avisos.append(f"Israel · {nombre}: {e}")
            if normal:
                reps, control, av = RI.generar_israel(pd.concat(normal, ignore_index=True), indices, MODELO_ISRAEL)
                avisos += [f"Israel: {a}" for a in av]
                for r in reps:
                    files.append((r["archivo"], r["bytes"]))
                    cards.append(_card_israel(r, len(files) - 1))
                    for f in r["faltan_indice"]:
                        if f["Codigo"] not in [p[0] for p in pend_israel]:
                            pend_israel.append((f["Codigo"], f["desc"]))
                files.append(("Control producción Israel.xlsx", RI.crear_control(reps, control)))
                control_idx = len(files) - 1
                control_lineas = _control_lineas(control)
                excl_desc = [{"desc": e["desc"], "cartones": int(e["cartones"])} for e in control["excl_desc"]]
        except Exception as e:
            avisos.append(f"Israel: {e}")

    token = None
    if files:
        _limpiar()
        token = secrets.token_urlsafe(12)
        _CACHE[token] = {"ts": time.time(), "files": files}
    ctx.update({
        "resultados": cards, "avisos": avisos, "token": token, "nfiles": len(files), "control_idx": control_idx,
        "control_lineas": control_lineas, "excl_desc": excl_desc,
        "pend_usa": "\n".join(f"{c}," for c, d in pend_usa),
        "pend_israel": "\n".join(f"{c},,,{d}," for c, d in pend_israel),
        "n_pend_usa": len(pend_usa), "n_pend_israel": len(pend_israel),
    })
    return render_template("index.html", **ctx)


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
