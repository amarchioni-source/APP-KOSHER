"""Genera el 'Production report - DEVESA' (formato USA Kosher) a partir de la bajada de stock.

Toma el modelo de Excel tal cual (modelo/modelo.xlsx) y reemplaza únicamente los datos de la hoja
visible, editando el XML directamente para conservar estilos, anchos, hojas ocultas y fórmulas.
"""
import copy
import csv
import datetime as dt
import io
import re
import zipfile

import pandas as pd
from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
q = lambda t: "{%s}%s" % (NS, t)

DESTINO_DEFECTO = "KOSHER USA"
SHEET_XML = "xl/worksheets/sheet2.xml"      # hoja visible del modelo
FILAS_MODELO = 34                            # filas de datos que admite el modelo sin desplazar (3..36)
FACTOR_ESCALA = 100000

REQUERIDAS = ["Codigo", "Destino", "MercDescripcion", "Unidad", "Peso", "PesoBruto", "Fecha", "FechaFaena"]


# ------------------------------------------------------------------ lectura de datos
def cargar_barcodes(path):
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            cod, bc = (row.get("codigo") or "").strip().upper(), (row.get("barcode") or "").strip()
            if cod and bc:
                out[cod] = bc
    return out


def parsear_barcodes_extra(texto):
    out = {}
    for linea in (texto or "").splitlines():
        partes = [p.strip() for p in re.split(r"[,;\t]", linea) if p.strip()]
        if len(partes) >= 2:
            out[partes[0].upper()] = partes[1]
    return out


def normalizar(df):
    """Unifica los nombres de columna de las distintas bajadas."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    ren = {}
    if "cajas" in df.columns:
        ren["cajas"] = "Unidad"
    if "kg" in df.columns:
        ren["kg"] = "Peso"
    if "kg bruto" in df.columns:
        ren["kg bruto"] = "PesoBruto"
    df = df.rename(columns=ren)
    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError("Al archivo le faltan las columnas: " + ", ".join(faltan))
    df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.normalize()
    df["FechaFaena"] = pd.to_datetime(df["FechaFaena"]).dt.normalize()
    for c in ["Unidad", "Peso", "PesoBruto"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Codigo"] = df["Codigo"].astype(str).str.strip().str.upper()
    return df


# ------------------------------------------------------------------ armado de cada día
def _serial(d):
    return (pd.Timestamp(d).to_pydatetime() - dt.datetime(1899, 12, 30)).days


def agrupar(df_dia, barcodes):
    g = df_dia.groupby("Codigo").agg(
        desc=("MercDescripcion", lambda s: s.dropna().astype(str).iloc[0]),
        faena=("FechaFaena", "max"),
        cajas=("Unidad", "sum"),
        bruto=("PesoBruto", "sum"),
        neto=("Peso", "sum"),
    ).reset_index()
    g["barcode"] = g["Codigo"].map(barcodes)
    g["_orden"] = g["barcode"].fillna("~")            # sin barcode, al final
    g = g.sort_values(["_orden", "Codigo"]).reset_index(drop=True)
    g["cajas"] = g["cajas"].round(0)
    g["bruto"] = g["bruto"].round(3)
    g["neto"] = g["neto"].round(3)
    return g


def _workbook_bytes(g, dia, modelo_path):
    n = len(g)
    k = max(0, n - FILAS_MODELO)                      # filas extra a insertar
    z = zipfile.ZipFile(modelo_path)
    infos = z.infolist()
    files = {i.filename: z.read(i.filename) for i in infos}

    sst = etree.fromstring(files["xl/sharedStrings.xml"])
    idx = {}
    for i, si in enumerate(sst.findall(q("si"))):
        t = si.find(q("t"))
        if t is not None and len(si) == 1:
            idx.setdefault(t.text, i)
    cnt = [len(sst.findall(q("si")))]

    def sid(text):
        if text in idx:
            return idx[text]
        si = etree.SubElement(sst, q("si"))
        t = etree.SubElement(si, q("t"))
        t.text = text
        if text != text.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        idx[text] = cnt[0]
        cnt[0] += 1
        return idx[text]

    sh = etree.fromstring(files[SHEET_XML])
    sd = sh.find(q("sheetData"))

    if k:                                              # desplazar el bloque de totales
        for r in sd.findall(q("row")):
            rn = int(r.get("r"))
            if rn >= 37:
                r.set("r", str(rn + k))
                for c in r.findall(q("c")):
                    c.set("r", re.match(r"[A-Z]+", c.get("r")).group(0) + str(rn + k))
    rows = {int(r.get("r")): r for r in sd.findall(q("row"))}

    if n > 33:                                         # filas 36.. con el estilo de datos de la fila 35
        r35 = rows[35]
        for rn in range(36, 3 + n):
            if rn == 36:
                src = {re.match(r"[A-Z]+", c.get("r")).group(0): c for c in r35.findall(q("c"))}
                for c in rows[36].findall(q("c")):
                    s_ = src[re.match(r"[A-Z]+", c.get("r")).group(0)].get("s")
                    c.attrib.pop("s", None)
                    if s_ is not None:
                        c.set("s", s_)
                rows[36].set("ht", r35.get("ht"))
            else:
                new = copy.deepcopy(r35)
                new.set("r", str(rn))
                for c in new.findall(q("c")):
                    c.set("r", re.match(r"[A-Z]+", c.get("r")).group(0) + str(rn))
                pos = [i for i, r in enumerate(sd.findall(q("row"))) if int(r.get("r")) > rn][0]
                sd.insert(pos, new)
                rows[rn] = new

    def setcell(r, col, val, kind):
        ref = f"{col}{r}"
        c = [c for c in rows[r].findall(q("c")) if c.get("r") == ref][0]
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop("t", None)
        if val is None:
            return
        v = etree.SubElement(c, q("v"))
        if kind == "s":
            c.set("t", "s")
            v.text = str(sid(val))
        else:
            v.text = repr(val) if isinstance(val, float) else str(val)

    S = _serial(dia)
    for i in range(3, 3 + max(33, n)):
        j = i - 3
        if j < n:
            x = g.iloc[j]
            setcell(i, "A", x["barcode"] if isinstance(x["barcode"], str) else None, "s")
            setcell(i, "B", "AZUL", "s")
            setcell(i, "C", S, "n")
            setcell(i, "D", x["desc"], "s")
            setcell(i, "E", x["Codigo"], "s")
            setcell(i, "F", _serial(x["faena"]), "n")
            setcell(i, "G", int(x["cajas"]), "n")
            setcell(i, "H", float(x["bruto"]), "n")
            setcell(i, "I", float(x["neto"]), "n")
        else:
            for col in "ABCDEFGHI":
                setcell(i, col, None, "n")

    # fórmulas (mismas del modelo, con el rango ajustado) y valores en caché
    T, E, SP, NSP, P = 37 + k, 36 + k, 40 + k, 41 + k, 44 + k
    sp = g[g["desc"].str.upper().str.contains("SPICED")]
    V = {}
    for col, c in zip("GHI", ["cajas", "bruto", "neto"]):
        tot = int(g[c].sum()) if col == "G" else round(float(g[c].sum()), 3)
        spv = int(sp[c].sum()) if col == "G" else round(float(sp[c].sum()), 3)
        V[f"{col}{T}"] = (f"SUM({col}3:{col}{E})", tot)
        V[f"{col}{SP}"] = (f'+SUMIF(D3:D{E},"*SPICED*",{col}3:{col}{E})', spv)
        V[f"{col}{NSP}"] = (f"+{col}{T}-{col}{SP}", tot - spv if col == "G" else round(tot - spv, 3))
    V[f"G{P}"] = (f"G{SP}/G{T}", V[f"G{SP}"][1] / V[f"G{T}"][1] if V[f"G{T}"][1] else 0)
    for r in rows.values():
        for c in r.findall(q("c")):
            if c.get("r") in V:
                f, v = V[c.get("r")]
                c.find(q("f")).text = f
                c.find(q("v")).text = repr(v)

    sh.find(q("sheetViews")).find(q("sheetView")).find(q("pane")).set("topLeftCell", "A3")
    if k:
        sh.find(q("dimension")).set("ref", f"A1:J{80 + k}")
        sh.find(q("autoFilter")).set("ref", f"A2:I{57 + k}")
        files["xl/workbook.xml"] = files["xl/workbook.xml"].decode("utf8").replace(
            "$A$2:$I$57", f"$A$2:$I${57 + k}").encode("utf8")
    sst.set("uniqueCount", str(cnt[0]))
    sst.set("count", str(int(sst.get("count")) + n * 3))
    files[SHEET_XML] = etree.tostring(sh, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["xl/sharedStrings.xml"] = etree.tostring(sst, xml_declaration=True, encoding="UTF-8", standalone=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in infos:
            zo.writestr(i.filename, files[i.filename])
    return buf.getvalue(), V


def generar_reportes(df, barcodes, modelo_path, destino=DESTINO_DEFECTO):
    """Devuelve (reportes, avisos). Un reporte por cada fecha con producción del destino."""
    df = normalizar(df)
    avisos = []
    kdf = df[df["Destino"].astype(str).str.strip().str.upper() == destino.strip().upper()].copy()
    if kdf.empty:
        destinos = ", ".join(sorted(df["Destino"].dropna().astype(str).unique())[:15])
        raise ValueError(f"No hay registros con destino «{destino}». Destinos en el archivo: {destinos}")

    # Algunas bajadas traen los pesos multiplicados por 100.000: se detecta por kg/caja
    cajas, neto = kdf["Unidad"].sum(), kdf["Peso"].sum()
    if cajas > 0 and neto / cajas > 1000:
        kdf["Peso"] /= FACTOR_ESCALA
        kdf["PesoBruto"] /= FACTOR_ESCALA
        avisos.append(f"Los pesos venían multiplicados por {FACTOR_ESCALA:,}; se dividieron automáticamente."
                      .replace(",", "."))

    reportes = []
    for dia in sorted(kdf["Fecha"].unique()):
        dia = pd.Timestamp(dia)
        g = agrupar(kdf[kdf["Fecha"] == dia], barcodes)
        data, V = _workbook_bytes(g, dia, modelo_path)
        k = max(0, len(g) - FILAS_MODELO)
        sin_bc = g[g["barcode"].isna()][["Codigo", "desc", "cajas"]].to_dict("records")
        reportes.append({
            "fecha": dia.strftime("%d/%m/%Y"),
            "archivo": f"Production_report_-_DEVESA_{dia:%d-%m}_-{dia:%Y}_USA_-.xlsx",
            "bytes": data,
            "productos": len(g),
            "cajas": V[f"G{37 + k}"][1],
            "bruto": V[f"H{37 + k}"][1],
            "neto": V[f"I{37 + k}"][1],
            "pct_spiced": V[f"G{44 + k}"][1] * 100,
            "sin_barcode": sin_bc,
        })
    return reportes, avisos
