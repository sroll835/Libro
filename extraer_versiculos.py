"""Extrae todos los versículos de Biblia.pdf (Reina Valera 1909) a versiculos.json.

Uso:  python extraer_versiculos.py
"""
import json
import re
from pathlib import Path

from pypdf import PdfReader

AQUI = Path(__file__).resolve().parent
PDF = AQUI.parent / "Biblia.pdf"
SALIDA = AQUI / "versiculos.json"
DESFASE_PAGINAS = 3  # la página 1 del índice es la página 4 (índice 0) del PDF... +3

# Número de capítulos canónico por libro, para validar la extracción.
CAPITULOS = [50, 40, 27, 36, 34, 24, 21, 4, 31, 24, 22, 25, 29, 36, 10, 13, 10, 42, 150, 31,
             12, 8, 66, 52, 5, 48, 12, 14, 3, 9, 1, 4, 7, 3, 3, 3, 2, 14, 4, 28, 16, 24, 21, 28,
             16, 16, 13, 6, 6, 4, 4, 5, 3, 6, 4, 3, 1, 13, 5, 5, 3, 5, 1, 1, 1, 22]


def leer_indice(reader):
    libros = []
    for i in (2, 3):
        for linea in reader.pages[i].extract_text().splitlines():
            m = re.match(r"^(.+?)\s*(?:\.\s*)+\s*(\d+)\s*$", linea.strip())
            if m:
                libros.append((m.group(1).strip(), int(m.group(2))))
    return libros


def limpiar(texto):
    texto = re.sub(r"\s+", " ", texto)
    texto = re.sub(r"\s+([,;.:?!])", r"\1", texto)
    texto = re.sub(r"([,;:])(?=[^\s\d])", r"\1 ", texto)
    return texto.strip()


def dividir_libro(nombre, texto):
    """Recorre el texto buscando, en orden, el siguiente versículo o el siguiente capítulo."""
    capitulos = {}
    cap, ver = 1, 1
    m = re.match(r"\s*1\s", texto)
    if not m:
        raise ValueError(f"{nombre}: no empieza con el versículo 1")
    pos, inicio = m.end(), m.end()

    def cerrar(fin):
        capitulos.setdefault(cap, []).append(limpiar(texto[inicio:fin]))

    while True:
        mv = re.compile(rf"(?<![\w\d]){ver + 1}\s").search(texto, pos)
        mc = re.compile(rf"(?<![\w\d]){cap + 1}\s+1\s").search(texto, pos)
        # "17 1 Y ABRAHAM..." puede ser versículo 17 o capítulo 17: los capítulos
        # empiezan con una palabra en mayúsculas.
        es_capitulo = mc and mv and mc.start() == mv.start() and re.match(
            r"(?:\S+\s+){0,2}?[A-ZÁÉÍÓÚÑ]{2,}", texto[mc.end():mc.end() + 60])
        if mv and not es_capitulo and (not mc or mv.start() <= mc.start()):
            cerrar(mv.start())
            ver += 1
            pos = inicio = mv.end()
        elif mc:
            cerrar(mc.start())
            cap, ver = cap + 1, 1
            pos = inicio = mc.end()
        else:
            cerrar(len(texto))
            return capitulos


def main():
    reader = PdfReader(str(PDF))
    libros = leer_indice(reader)
    assert len(libros) == 66, f"Se esperaban 66 libros en el índice, hay {len(libros)}"

    # Todas las líneas del cuerpo, sin la cabecera de cada página
    lineas = []  # (indice_pagina, texto)
    for p in range(4, len(reader.pages)):
        filas = reader.pages[p].extract_text().splitlines()
        for fila in filas[1:]:
            lineas.append((p, fila.strip()))

    # Línea donde aparece el título de cada libro
    posiciones = []
    for nombre, pagina in libros:
        pdf_idx = pagina + DESFASE_PAGINAS
        # Algunos títulos vienen sin el número (p. ej. "Crónicas" para "1 Crónicas")
        titulos = {nombre, re.sub(r"^\d\s+", "", nombre)}
        idx = next(i for i, (p, t) in enumerate(lineas) if p == pdf_idx and t in titulos)
        posiciones.append(idx)
    posiciones.append(len(lineas))

    salida, problemas = [], []
    for n, (nombre, _) in enumerate(libros):
        texto = " ".join(t for _, t in lineas[posiciones[n] + 1: posiciones[n + 1]])
        caps = dividir_libro(nombre, texto)
        if len(caps) != CAPITULOS[n]:
            problemas.append(f"{nombre}: {len(caps)} capítulos (se esperaban {CAPITULOS[n]})")
        for c, versos in caps.items():
            for v, t in enumerate(versos, 1):
                if not t:  # el PDF deja algunos versículos vacíos (texto movido al siguiente)
                    continue
                salida.append({"libro": nombre, "orden_libro": n + 1, "capitulo": c, "versiculo": v, "texto": t})

    SALIDA.write_text(json.dumps(salida, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"{len(salida)} versículos guardados en {SALIDA.name}")
    for p in problemas:
        print("AVISO:", p)


if __name__ == "__main__":
    main()
