"""Convierte el Manual de Funciones (Word) en datos-manual.js para la app.

Uso:  py -I herramientas/manual-a-datos.py "ruta/al/MANUAL DE FUNCIONES.docx"
Genera datos-manual.js (window.HUB_MANUAL) con cargos y diccionario de competencias.
Lo que se edite luego en la pantalla de Configuración manda sobre este archivo.
"""
import json, re, sys, unicodedata
from pathlib import Path
import docx

def limpio(t):
    return re.sub(r"\s+", " ", t.replace("\xa0", " ")).strip()

def slug(t):
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", t).strip("-")

def filas(tabla):
    out = []
    for r in tabla.rows:
        celdas = []
        for c in r.cells:
            x = c.text.strip()
            if not celdas or celdas[-1] != x:
                celdas.append(x)
        out.append(celdas)
    return out

def main(ruta):
    d = docx.Document(ruta)
    competencias = {}
    cargos, actual = [], None
    for t in d.tables:
        f = filas(t)
        if not f:
            continue
        cab = limpio(f[0][0]).upper()
        if cab.startswith("COMPETENCIAS NIVEL"):
            nivel = "estrategico" if "ESTRAT" in cab else "operativo"
            for fila in f[2:]:
                if len(fila) >= 2:
                    competencias.setdefault(limpio(fila[0]), {"descripcion": limpio(fila[1]), "niveles": []})["niveles"].append(nivel)
            continue
        claves = {limpio(x[0]).rstrip(":").lower(): x for x in f if x}
        if "denominación del cargo" in claves:
            fila = claves["denominación del cargo"]
            jefe = claves.get("cargo jefe inmediato") or claves.get("reporta a") or ["", ""]
            sup = claves.get("cargos que supervisa", ["", ""])
            actual = {
                "id": slug(fila[1]), "nombre": limpio(fila[1]), "codigo": limpio(fila[2]) if len(fila) > 2 else "",
                "area": limpio(claves.get("unidad / área", ["", ""])[1]),
                "jefeInmediato": limpio(jefe[1].replace("\n", " / ")),
                "supervisa": limpio(sup[1].replace("\n", " / ")),
                "mision": "", "competencias": [],
                "nivel": "estrategico" if len(cargos) < 11 else "operativo",
            }
            cargos.append(actual)
            continue
        if actual and cab.startswith("MISIÓN DEL CARGO") and len(f) > 1:
            actual["mision"] = limpio(" ".join(x[0] for x in f[1:]))
        if actual and ("competencias" in claves):
            actual["competencias"] = [limpio(x) for x in claves["competencias"][1].split("\n") if limpio(x)]
    datos = {"fuente": Path(ruta).name, "competencias": competencias, "cargos": cargos}
    js = "// Generado por herramientas/manual-a-datos.py desde el Manual de Funciones (GTH-MFC-01).\n"
    js += "window.HUB_MANUAL = " + json.dumps(datos, ensure_ascii=False, indent=1) + ";\n"
    salida = Path(__file__).resolve().parent.parent / "datos-manual.js"
    salida.write_text(js, encoding="utf-8")
    print(f"{len(cargos)} cargos, {len(competencias)} competencias -> {salida}")
    for c in cargos:
        print(f"  [{c['nivel'][:4]}] {c['nombre']} <- {c['jefeInmediato']} | {len(c['competencias'])} comp")

if __name__ == "__main__":
    main(sys.argv[1])
