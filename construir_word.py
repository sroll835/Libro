"""Maqueta los versículos decodificados en libros de Word elegantes (un archivo por libro o volumen).

Uso:
  python construir_word.py                       Todos los libros que tengan versículos decodificados
  python construir_word.py --libro "Juan"        Solo un libro
  python construir_word.py --por-volumen 150     Partir en volúmenes de N versículos (por defecto 0 = un solo archivo por libro)
  python construir_word.py --resultados muestra  Usa otra carpeta de resultados
"""
import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

AQUI = Path(__file__).resolve().parent
VERSICULOS = AQUI / "versiculos.json"
SALIDA = AQUI / "libros"

FUENTE = "Garamond"
BURDEOS = RGBColor(0x6B, 0x1E, 0x2B)
DORADO = RGBColor(0xA8, 0x87, 0x4A)
TINTA = RGBColor(0x2B, 0x2B, 0x2B)
AZUL = RGBColor(0x1F, 0x2F, 0x4A)
GRIS = RGBColor(0x6E, 0x66, 0x5C)
HEX_DORADO, HEX_MARFIL = "A8874A", "F7F2E7"

EMOJIS = re.compile("[\U0001F000-\U0001FAFF☀-➿️‍]")
NIVELES = [
    ("I", "Superficie", "Literal y narrativo"),
    ("II", "Lingüístico", "Texto original y forma literaria"),
    ("III", "Histórico-cultural", "El mundo detrás del texto"),
    ("IV", "Simbólico-teológico", "Profundidad doctrinal"),
    ("V", "Tradición e interpretación", "Historia del texto en la historia"),
    ("VI", "Cábala simbólica", "El interior del interior"),
    ("VII", "Dimensión mística", "El nivel supremo"),
]


# ------------------------------------------------------------ utilidades XML

# Word exige que los hijos de pPr / rPr / settings aparezcan en el orden del esquema OOXML
ORDEN_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
ORDEN_RPR = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
             "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
ORDEN_SETTINGS = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
                  "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
                  "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts", "embedSystemFonts",
                  "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges",
                  "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "gutterAtTop",
                  "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState",
                  "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
                  "stylePaneSortMethod", "documentType", "mailMerge", "revisionView", "trackRevisions",
                  "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
                  "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
                  "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope",
                  "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders",
                  "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets",
                  "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
                  "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
                  "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
                  "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
                  "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
                  "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml",
                  "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml",
                  "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags",
                  "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr",
                  "endnotePr", "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang",
                  "clrSchemeMapping", "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures",
                  "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
                  "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"]


def insertar(padre, elemento, orden):
    """Coloca `elemento` en `padre` respetando el orden del esquema (reemplaza si ya existía)."""
    nombre = elemento.tag.split("}")[1]
    for viejo in padre.findall(qn(f"w:{nombre}")):
        padre.remove(viejo)
    siguientes = {qn(f"w:{t}") for t in orden[orden.index(nombre) + 1:]}
    for i, hijo in enumerate(padre):
        if hijo.tag in siguientes:
            padre.insert(i, elemento)
            return
    padre.append(elemento)


def borde(parrafo, lados=("top", "bottom"), color=HEX_DORADO, grosor=6, espacio=6):
    bdr = OxmlElement("w:pBdr")
    for lado in ("top", "left", "bottom", "right"):  # orden exigido dentro de pBdr
        if lado in lados:
            e = OxmlElement(f"w:{lado}")
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(grosor))
            e.set(qn("w:space"), str(espacio))
            e.set(qn("w:color"), color)
            bdr.append(e)
    insertar(parrafo._p.get_or_add_pPr(), bdr, ORDEN_PPR)


def sombreado(parrafo, color=HEX_MARFIL):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    insertar(parrafo._p.get_or_add_pPr(), shd, ORDEN_PPR)


def espaciado_letras(run, puntos):
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:val"), str(int(puntos * 20)))
    insertar(run._r.get_or_add_rPr(), sp, ORDEN_RPR)


def campo(parrafo, instruccion, marcador=""):
    """Inserta un campo de Word (PAGE, TOC...)."""
    def fld(tipo):
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), tipo)
        r = OxmlElement("w:r")
        r.append(f)
        return r
    parrafo._p.append(fld("begin"))
    r = OxmlElement("w:r")
    t = OxmlElement("w:instrText")
    t.set(qn("xml:space"), "preserve")
    t.text = f" {instruccion} "
    r.append(t)
    parrafo._p.append(r)
    parrafo._p.append(fld("separate"))
    run = parrafo.add_run(marcador)
    run.italic = True
    parrafo._p.append(fld("end"))


def fuente_completa(estilo_o_run, nombre=FUENTE):
    rpr = estilo_o_run.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        insertar(rpr, fonts, ORDEN_RPR)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        fonts.set(qn(a), nombre)
    fonts.set(qn("w:cs"), "Times New Roman")  # hebreo y griego


# ------------------------------------------------------------ estilos

def preparar_documento():
    doc = Document()
    settings = doc.settings.element
    for etiqueta in ("w:mirrorMargins", "w:updateFields"):
        e = OxmlElement(etiqueta)
        if etiqueta == "w:updateFields":  # Word rellena el índice al abrir el archivo
            e.set(qn("w:val"), "true")
        insertar(settings, e, ORDEN_SETTINGS)

    s = doc.sections[0]
    s.page_width, s.page_height = Cm(17), Cm(24)
    s.top_margin, s.bottom_margin = Cm(2.2), Cm(2.2)
    s.left_margin, s.right_margin = Cm(2.3), Cm(1.8)  # interior / exterior (márgenes simétricos)
    s.header_distance = s.footer_distance = Cm(1.1)

    est = doc.styles
    normal = est["Normal"]
    normal.font.name, normal.font.size, normal.font.color.rgb = FUENTE, Pt(11.5), TINTA
    fuente_completa(normal)
    pf = normal.paragraph_format
    pf.space_after, pf.line_spacing = Pt(6), 1.18
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control = True

    def titulo(nombre, tam, color, versalitas=True, negrita=False, cursiva=False, antes=0, despues=6):
        st = est[nombre]
        st.font.name, st.font.size, st.font.color.rgb = FUENTE, Pt(tam), color
        st.font.bold, st.font.italic, st.font.small_caps = negrita, cursiva, versalitas
        fuente_completa(st)
        st.paragraph_format.space_before, st.paragraph_format.space_after = Pt(antes), Pt(despues)
        st.paragraph_format.keep_with_next = True
        return st

    h1 = titulo("Heading 1", 30, BURDEOS, antes=150, despues=4)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.page_break_before = True
    h2 = titulo("Heading 2", 17, BURDEOS, antes=4, despues=10)
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    titulo("Heading 3", 14.5, BURDEOS, antes=22, despues=8)
    titulo("Heading 4", 12.5, AZUL, versalitas=False, negrita=True, antes=12, despues=4)
    titulo("Heading 5", 11.5, AZUL, versalitas=False, negrita=True, cursiva=True, antes=8, despues=2)

    for nombre in ("List Bullet", "List Bullet 2"):
        est[nombre].font.name = FUENTE
        est[nombre].paragraph_format.space_after = Pt(3)

    vers = est.add_style("Versiculo", WD_STYLE_TYPE.PARAGRAPH)
    vers.base_style = normal
    vers.font.size, vers.font.italic, vers.font.color.rgb = Pt(14), True, TINTA
    vf = vers.paragraph_format
    vf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    vf.left_indent = vf.right_indent = Cm(0.8)
    vf.space_before, vf.space_after, vf.line_spacing = Pt(6), Pt(18), 1.3
    return doc


# ------------------------------------------------------------ piezas del libro

def ornamento(doc, texto="✦", tam=12, antes=0, despues=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(antes), Pt(despues)
    r = p.add_run(texto)
    r.font.color.rgb, r.font.size = DORADO, Pt(tam)
    return p


def linea_centrada(doc, texto, tam, color=TINTA, versalitas=False, cursiva=False, letras=0, antes=0, despues=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(antes), Pt(despues)
    r = p.add_run(texto)
    r.font.size, r.font.color.rgb, r.font.small_caps, r.italic = Pt(tam), color, versalitas, cursiva
    if letras:
        espaciado_letras(r, letras)
    return p


def salto(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def portada(doc, libro, subtitulo):
    linea_centrada(doc, "LA BIBLIA", 20, DORADO, letras=8, antes=110)
    linea_centrada(doc, "Decodificada", 44, BURDEOS, versalitas=True, letras=2, despues=4)
    ornamento(doc, "─────  ✦  ─────", 12, despues=10)
    linea_centrada(doc, "Los siete niveles del Gran Decodificador Sagrado", 13, GRIS, cursiva=True, despues=70)
    linea_centrada(doc, libro, 32, TINTA, versalitas=True, letras=3)
    if subtitulo:
        linea_centrada(doc, subtitulo, 13, GRIS, cursiva=True, antes=6)
    linea_centrada(doc, "Texto bíblico: Reina Valera 1909 · Dominio público", 9.5, GRIS, antes=150)
    salto(doc)


def pagina_niveles(doc):
    linea_centrada(doc, "Los Siete Niveles", 22, BURDEOS, versalitas=True, letras=1.5, antes=40, despues=2)
    ornamento(doc, "✦", 11, despues=18)
    for num, nombre, desc in NIVELES:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run(f"{num}   ")
        r.font.color.rgb, r.font.size = DORADO, Pt(13)
        r = p.add_run(nombre)
        r.font.small_caps, r.font.size, r.font.color.rgb = True, Pt(15), TINTA
        p.add_run("\n")
        r = p.add_run(desc)
        r.italic, r.font.size, r.font.color.rgb = True, Pt(11), GRIS
    linea_centrada(doc, "Cada versículo se recorre desde la letra hasta el silencio: "
                        "de lo que el texto dice a lo que el texto revela.",
                   11, GRIS, cursiva=True, antes=24)
    salto(doc)


def indice(doc):
    linea_centrada(doc, "Índice", 22, BURDEOS, versalitas=True, letras=1.5, antes=20, despues=2)
    ornamento(doc, "✦", 11, despues=14)
    p = doc.add_paragraph()
    campo(p, 'TOC \\o "1-2" \\h \\z \\u',
          "Haz clic derecho aquí y elige «Actualizar campo» para generar el índice.")


def encabezado_y_pie(doc, libro):
    s = doc.sections[0]
    s.different_first_page_header_footer = True
    p = s.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(libro)
    r.font.small_caps, r.font.size, r.font.color.rgb = True, Pt(9.5), DORADO
    espaciado_letras(r, 1.5)
    p = s.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.style.font.size = Pt(9.5)
    campo(p, "PAGE", "1")


def apertura_capitulo(doc, libro, cap, versos):
    doc.add_heading(f"Capítulo {cap}", level=1)
    ornamento(doc, "─────  ✦  ─────", 11, despues=26)
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.35
    for v in versos:
        r = p.add_run(f"{v['versiculo']}")
        r.font.superscript, r.font.color.rgb, r.bold = True, DORADO, True
        p.add_run(f" {v['texto']} ")
    linea_centrada(doc, f"{libro} {cap}", 9.5, GRIS, versalitas=True, letras=1, antes=10)


# ------------------------------------------------------------ Markdown -> Word

INLINE = re.compile(r"(\*\*\*.+?\*\*\*|\*\*.+?\*\*|(?<![\w*])\*(?!\s).+?(?<!\s)\*(?![\w*])|`.+?`)")


def texto_enriquecido(parrafo, texto, color=None, cursiva_base=False):
    for trozo in INLINE.split(texto):
        if not trozo:
            continue
        negrita = cursiva = False
        if trozo.startswith("***") and trozo.endswith("***") and len(trozo) > 6:
            trozo, negrita, cursiva = trozo[3:-3], True, True
        elif trozo.startswith("**") and trozo.endswith("**") and len(trozo) > 4:
            trozo, negrita = trozo[2:-2], True
        elif trozo.startswith("*") and trozo.endswith("*") and len(trozo) > 2:
            trozo, cursiva = trozo[1:-1], True
        elif trozo.startswith("`") and trozo.endswith("`"):
            trozo = trozo[1:-1]
        r = parrafo.add_run(trozo)
        r.bold = negrita or None
        r.italic = (cursiva != cursiva_base) or None
        if color:
            r.font.color.rgb = color


def limpiar_titulo(t):
    return EMOJIS.sub("", t).strip(" #—-").strip()


def oraculo(doc, lineas):
    ornamento(doc, "❦", 14, antes=6, despues=4)
    for i, linea in enumerate(lineas):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.left_indent = pf.right_indent = Cm(1)
        pf.space_after, pf.line_spacing = Pt(0), 1.4
        sombreado(p)
        lados = ["left", "right"] + (["top"] if i == 0 else []) + (["bottom"] if i == len(lineas) - 1 else [])
        borde(p, lados, espacio=10)
        texto_enriquecido(p, linea, color=BURDEOS, cursiva_base=True)
        for r in p.runs:
            r.font.size = Pt(13)
    ornamento(doc, "❦", 14, antes=6, despues=8)


def markdown(doc, md):
    lineas = md.replace("\r", "").split("\n")
    i, parrafo = 0, []

    def volcar():
        if parrafo:
            p = doc.add_paragraph()
            texto_enriquecido(p, " ".join(parrafo))
            parrafo.clear()

    while i < len(lineas):
        linea = lineas[i]
        s = linea.strip()
        m_tit = re.match(r"^(#{1,6})\s+(.*)$", s)
        m_vin = re.match(r"^(\s*)[-*•]\s+(.*)$", linea)
        m_num = re.match(r"^(\s*)(\d+)[.)]\s+(.*)$", linea)
        if not s:
            volcar()
        elif m_tit:
            volcar()
            nivel = len(m_tit.group(1))
            titulo = limpiar_titulo(m_tit.group(2))
            if nivel <= 2:  # "## NIVEL N — ..."
                p = doc.add_heading(titulo.title() if titulo.isupper() else titulo, level=3)
                borde(p, ("bottom",), grosor=4, espacio=4)
            else:
                p = doc.add_heading(titulo, level=min(nivel + 1, 5))
                titulo_limpio = INLINE.sub(lambda m: m.group(0).strip("*`"), titulo)
                p.runs[0].text = titulo_limpio
        elif s.startswith(">"):
            volcar()
            citas = []
            while i < len(lineas) and lineas[i].strip().startswith(">"):
                contenido = lineas[i].strip().lstrip(">").strip()
                if contenido:
                    citas.append(contenido)
                i += 1
            if citas:
                oraculo(doc, citas)
            continue
        elif re.match(r"^(-{3,}|\*{3,}|_{3,})$", s):
            volcar()
            ornamento(doc, "·  ✦  ·", 10, antes=6, despues=6)
        elif m_vin:
            volcar()
            estilo = "List Bullet 2" if len(m_vin.group(1)) >= 2 else "List Bullet"
            texto_enriquecido(doc.add_paragraph(style=estilo), m_vin.group(2))
        elif m_num:
            volcar()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Cm(0.8), Cm(-0.5)
            r = p.add_run(f"{m_num.group(2)}. ")
            r.font.color.rgb, r.bold = BURDEOS, True
            texto_enriquecido(p, m_num.group(3))
        elif s.lower().startswith(("**nivel de certeza", "nivel de certeza", "*nivel de certeza")):
            volcar()
            p = doc.add_paragraph()
            sombreado(p)
            borde(p, ("left",), grosor=18, espacio=8)
            p.paragraph_format.left_indent = Cm(0.3)
            p.paragraph_format.space_before = Pt(8)
            texto_enriquecido(p, s, color=GRIS)
            for r in p.runs:
                r.font.size = Pt(10.5)
        else:
            parrafo.append(s)
        i += 1
    volcar()


def pagina_versiculo(doc, v, md):
    ornamento(doc, "✦", 11, antes=0, despues=2).paragraph_format.page_break_before = True
    h = doc.add_heading(f"{v['libro']} {v['capitulo']}:{v['versiculo']}", level=2)
    espaciado_letras(h.runs[0], 1.5)
    p = doc.add_paragraph(style="Versiculo")
    borde(p, ("top", "bottom"), espacio=10)
    p.add_run(f"«{v['texto']}»")
    markdown(doc, md)


# ------------------------------------------------------------ volúmenes

def volumenes(capitulos, limite):
    if not limite or limite <= 0:  # sin límite: todo el libro en un solo archivo
        yield list(capitulos)
        return
    actual, cuenta = [], 0
    for cap, versos in capitulos.items():
        if actual and cuenta + len(versos) > limite:
            yield actual
            actual, cuenta = [], 0
        actual.append(cap)
        cuenta += len(versos)
    if actual:
        yield actual


def construir(libro, orden, capitulos, resultados, limite):
    carpeta = resultados / f"{orden:02d} {libro}"
    hechos = {(c, v["versiculo"]): (carpeta / f"{c:03d}_{v['versiculo']:03d}.md")
              for c, vs in capitulos.items() for v in vs}
    hechos = {k: r for k, r in hechos.items() if r.exists()}
    if not hechos:
        return
    con_datos = {c: vs for c, vs in capitulos.items() if any((c, v["versiculo"]) in hechos for v in vs)}
    grupos = list(volumenes(con_datos, limite))
    SALIDA.mkdir(exist_ok=True)
    for n, caps in enumerate(grupos, 1):
        varios = len(grupos) > 1
        rango = f"Capítulo {caps[0]}" if len(caps) == 1 else f"Capítulos {caps[0]} – {caps[-1]}"
        subtitulo = f"Volumen {n} · {rango}" if varios else rango
        doc = preparar_documento()
        encabezado_y_pie(doc, libro)
        portada(doc, libro, subtitulo)
        pagina_niveles(doc)
        indice(doc)
        total = 0
        for c in caps:
            apertura_capitulo(doc, libro, c, capitulos[c])
            for v in capitulos[c]:
                ruta = hechos.get((c, v["versiculo"]))
                if ruta:
                    pagina_versiculo(doc, v, ruta.read_text(encoding="utf-8"))
                    total += 1
        tramo = f"cap {caps[0]}" if len(caps) == 1 else f"cap {caps[0]}-{caps[-1]}"
        sufijo = f" - Vol {n:02d} ({tramo})" if varios else ""
        # Al cambiar el reparto en volúmenes, los archivos anteriores de este libro quedan obsoletos
        if n == 1:
            for viejo in SALIDA.glob(f"Biblia Decodificada - {orden:02d} {libro}*.docx"):
                viejo.unlink()
        nombre = SALIDA / f"Biblia Decodificada - {orden:02d} {libro}{sufijo}.docx"
        doc.save(str(nombre))
        print(f"{nombre.name}: {total} versículos")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--libro")
    p.add_argument("--por-volumen", type=int, default=0)
    p.add_argument("--resultados", default="resultados")
    args = p.parse_args()

    libros = {}
    for v in json.loads(VERSICULOS.read_text(encoding="utf-8")):
        libros.setdefault((v["orden_libro"], v["libro"]), {}).setdefault(v["capitulo"], []).append(v)
    for (orden, libro), caps in libros.items():
        if args.libro and libro.lower() != args.libro.lower():
            continue
        construir(libro, orden, caps, AQUI / args.resultados, args.por_volumen)


if __name__ == "__main__":
    main()
