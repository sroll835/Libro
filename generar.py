"""Pasa cada versículo por el prompt "Biblia decodificada" usando la API de Claude.

Comandos:
  python generar.py estimar  [--libro "Juan"]          Estima el costo antes de gastar nada
  python generar.py prueba   "Juan 3:16"               Decodifica UN versículo ahora mismo (para probar)
  python generar.py enviar   [--libro "Juan"] [--max 500]   Envía versículos pendientes en lotes (Batch API, -50%)
  python generar.py descargar                          Descarga los lotes terminados
  python generar.py estado                             Muestra el avance
  python generar.py revisar  [--libro "Génesis"] [--detalle]  Lista los versículos incompletos o recortados

Un versículo cuenta como "hecho" solo si su archivo pasa el control de calidad (ver problemas()).
Los incompletos se vuelven a enviar con "enviar"; la versión vieja se guarda en resultados_respaldo.

Opciones comunes: --modelo claude-opus-5-5 | claude-sonnet-5-5 | claude-haiku-5-5
                  --esfuerzo low | medium | high | xhigh | max
                  --min-palabras 1000   (mínimo de palabras para aceptar un versículo)

Necesita la variable de entorno ANTHROPIC_API_KEY con tu clave de https://console.anthropic.com
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

import anthropic
from anthropic.types.message_create_params import MessageCreateParamsNonStreaming
from anthropic.types.messages.batch_create_params import Request
from docx import Document

AQUI = Path(__file__).resolve().parent
PROMPT_DOCX = AQUI.parent / "prompt biblia decodificada.docx"
VERSICULOS = AQUI / "versiculos.json"
RESULTADOS = AQUI / "resultados"
RESPALDO = AQUI / "resultados_respaldo"
LOTES = AQUI / "lotes.json"
MIN_PALABRAS = 1000  # Gn 1:1-1:5 (los buenos) tienen 1.190-2.030; los recortados, ~590

# Secciones que deben estar en todos los versículos (las que usan los resultados buenos)
SECCIONES = [
    "Traducción propia", "Lo que oye un lector sin formación", "Contexto narrativo", "Género y estructura",
    "Texto original", "Palabras clave", "Simbolismo central", "Canon completo", "Patrística",
    "Preguntas para el estudio", "El alma del texto", "Las cuatro capas del Pardés", "Las Sefirot",
    "Gematría", "Aplicación ética-espiritual", "7.1 — Vías Apofática y Catafática",
    "7.2 — Corrientes Místicas", "7.3 — Kenosis y Theosis", "7.4 — Pneumatología",
    "7.5 — Cristología Profunda", "7.6 — Escatología y Apocalíptica", "7.7 — Teología de la Belleza",
    "7.8 — Oráculo Sagrado",
]
# Frases con las que el modelo confiesa que recortó
FRASES_RECORTE = re.compile(
    r"se tratan con brevedad|se tratan brevemente|verso de lista|se omite\s*[.;]|"
    r"no hay una gematría tradicional relevante", re.I)

MODELO = "claude-opus-5-5"
ESFUERZO = "high"
MAX_TOKENS = 32000
POR_LOTE = 5000
# Precios por millón de tokens (entrada, salida). El Batch API cobra la mitad.
PRECIOS = {
    "claude-opus-5-5": (4.00, 20.00),
    "claude-sonnet-5-5": (2.00, 10.00),
    "claude-haiku-5-5": (0.10, 0.50),
}
TOKENS_SALIDA_ESTIMADOS = 12000  # por versículo, incluido el razonamiento interno
TOKENS_ENTRADA_ESTIMADOS = 3500

FORMATO = """

FORMATO DE SALIDA (obligatorio: tu respuesta se maqueta automáticamente en un libro)
- Escribe solo en Markdown y en español. Sin saludo, sin preámbulo y sin despedida.
- No repitas el texto del versículo al inicio: ya va impreso encima de tu respuesta.
- Cada nivel empieza exactamente con: "## NIVEL N — TÍTULO DEL NIVEL" (sin emojis).
- Subapartados con "### " y, si hace falta, "#### ". Listas con "- ".
- Termina cada nivel con una línea: "**Nivel de certeza:** Alta | Media | Baja — justificación breve".
- Escribe hebreo, arameo o griego en su alfabeto original seguido de la transliteración en *cursiva*.
- El Oráculo Sagrado (7.8) va bajo "### 7.8 — Oráculo Sagrado" y cada una de sus líneas empieza con "> ".
- No uses tablas, bloques de código ni emojis.
- Recibes versículos vecinos solo como contexto: decodifica únicamente el versículo indicado.

PROFUNDIDAD (obligatorio)
- Cada versículo se publica como una entrada completa del libro: desarrolla los siete niveles con la
  misma profundidad aunque sea una genealogía, una lista de nombres o un versículo breve. Extensión
  orientativa: entre 1.200 y 2.000 palabras.
- No resumas, no abrevies y no anuncies que omites o tratas algo «con brevedad». La regla de explicar
  con brevedad vale solo para un nivel ENTERO que de verdad no aplique, y aun así justifícalo en un párrafo.
  Si falta un dato tradicional (por ejemplo, no hay gematría clásica), ofrece el análisis más cercano.
- Usa siempre estos subapartados «### »: Traducción propia; Lo que oye un lector sin formación;
  Contexto narrativo; Género y estructura (Nivel 1). Texto original; Palabras clave; Figuras literarias;
  Intertextualidad (Nivel 2). Contexto histórico; Contexto cultural; Contexto religioso comparado (Nivel 3).
  Simbolismo central; Temas teológicos; Canon completo (Nivel 4). Patrística; Rabínica y targúmica;
  Medieval y Reforma; Moderna y contemporánea; Preguntas para el estudio (Nivel 5). El alma del texto;
  Las cuatro capas del Pardés; Las Sefirot; Gematría; Aplicación ética-espiritual (Nivel 6). 7.1 a 7.8 (Nivel 7).
- Nivel 2: texto hebreo (o griego) con vocales, más la Septuaginta y la Vulgata con su transliteración
  o traducción, y al menos tres palabras clave explicadas («#### »), cada una en un párrafo.
- Nivel 7: cada subapartado de 7.1 a 7.7 tiene al menos un párrafo completo; el Oráculo, de 5 a 9 líneas.
"""


def cargar_prompt():
    texto = "\n".join(p.text for p in Document(str(PROMPT_DOCX)).paragraphs)
    # El pasaje va en el mensaje del usuario, no en el prompt de sistema
    texto = texto.replace("[Pegar aquí el texto bíblico]", "(se indica en el mensaje del usuario)")
    return texto.strip() + FORMATO


def cargar_versiculos():
    return json.loads(VERSICULOS.read_text(encoding="utf-8"))


def id_de(v):
    return f"L{v['orden_libro']:02d}_C{v['capitulo']:03d}_V{v['versiculo']:03d}"


def ruta_resultado(v):
    return RESULTADOS / f"{v['orden_libro']:02d} {v['libro']}" / f"{v['capitulo']:03d}_{v['versiculo']:03d}.md"


def mensaje_usuario(v, todos, indice):
    vecinos = [x for x in todos[max(0, indice - 3): indice + 4]
               if x["libro"] == v["libro"] and x is not v]
    contexto = "\n".join(f"{x['capitulo']}:{x['versiculo']} {x['texto']}" for x in vecinos)
    return (f"PASAJE A DECODIFICAR\n{v['libro']} {v['capitulo']}:{v['versiculo']} (Reina Valera 1909)\n"
            f"«{v['texto']}»\n\nContexto inmediato (no decodificar):\n{contexto}")


def parametros(v, todos, indice, sistema, args):
    return {
        "model": args.modelo,
        "max_tokens": MAX_TOKENS,
        "thinking": {"type": "adaptive"},
        "output_config": {"effort": args.esfuerzo},
        "system": [{"type": "text", "text": sistema, "cache_control": {"type": "ephemeral"}}],
        "messages": [{"role": "user", "content": mensaje_usuario(v, todos, indice)}],
    }


def filtrar(todos, libro):
    if not libro:
        return list(enumerate(todos))
    return [(i, v) for i, v in enumerate(todos) if v["libro"].lower() == libro.lower()]


def leer_lotes():
    return json.loads(LOTES.read_text(encoding="utf-8")) if LOTES.exists() else []


def guardar_lotes(lotes):
    LOTES.write_text(json.dumps(lotes, ensure_ascii=False, indent=1), encoding="utf-8")


def texto_de(mensaje):
    return "\n".join(b.text for b in mensaje.content if b.type == "text").strip()


def problemas(texto, min_palabras=MIN_PALABRAS):
    """Devuelve la lista de defectos de un resultado; vacía si está completo."""
    p = []
    niveles = set(re.findall(r"^## NIVEL (\d)", texto, re.M))
    if niveles != set("1234567"):
        p.append("faltan niveles: " + ",".join(sorted(set("1234567") - niveles)))
    certezas = texto.count("**Nivel de certeza:**")
    if certezas < 7:
        p.append(f"{certezas}/7 líneas de certeza")
    palabras = len(texto.split())
    if palabras < min_palabras:
        p.append(f"solo {palabras} palabras")
    titulos = set(re.findall(r"^### (.+?)\s*$", texto, re.M))
    faltan = [s for s in SECCIONES if not any(t.startswith(s) for t in titulos)]
    if faltan:
        p.append("faltan secciones: " + "; ".join(faltan))
    oraculo = texto.split("7.8 — Oráculo Sagrado", 1)[-1] if "7.8 — Oráculo Sagrado" in texto else ""
    lineas = len(re.findall(r"^> ", oraculo, re.M))
    if lineas < 5:
        p.append(f"oráculo de {lineas} líneas")
    recorte = FRASES_RECORTE.search(texto)
    if recorte:
        p.append(f"dice «{recorte.group(0).strip()}»")
    return p


def hecho(v, min_palabras=MIN_PALABRAS):
    ruta = ruta_resultado(v)
    return ruta.exists() and not problemas(ruta.read_text(encoding="utf-8"), min_palabras)


def guardar(v, texto):
    ruta = ruta_resultado(v)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    if ruta.exists():  # nunca perder una versión anterior
        viejo = RESPALDO / ruta.relative_to(RESULTADOS)
        viejo.parent.mkdir(parents=True, exist_ok=True)
        viejo.write_text(ruta.read_text(encoding="utf-8"), encoding="utf-8")
    ruta.write_text(texto, encoding="utf-8")


# ----------------------------------------------------------------- comandos

def cmd_estimar(args):
    todos = cargar_versiculos()
    pendientes = [v for _, v in filtrar(todos, args.libro) if not hecho(v, args.min_palabras)]
    n = len(pendientes)
    print(f"Versículos pendientes: {n:,}".replace(",", "."))
    print(f"Supuesto: ~{TOKENS_SALIDA_ESTIMADOS} tokens de salida por versículo "
          f"(usa 'prueba' para medir el real)\n")
    for modelo, (pe, ps) in PRECIOS.items():
        costo = n * (TOKENS_ENTRADA_ESTIMADOS * pe + TOKENS_SALIDA_ESTIMADOS * ps) / 1e6 / 2
        print(f"  {modelo:<20} con Batch API: ~US$ {costo:,.0f}".replace(",", "."))


def cmd_prueba(args):
    m = re.match(r"^(.+?)\s+(\d+):(\d+)$", args.referencia.strip())
    if not m:
        sys.exit('Formato: "Libro capítulo:versículo", por ejemplo "Juan 3:16"')
    todos = cargar_versiculos()
    clave = (m.group(1).lower(), int(m.group(2)), int(m.group(3)))
    indice = next((i for i, v in enumerate(todos)
                   if (v["libro"].lower(), v["capitulo"], v["versiculo"]) == clave), None)
    if indice is None:
        sys.exit("No encontré ese versículo.")
    v = todos[indice]
    print(f"Decodificando {v['libro']} {v['capitulo']}:{v['versiculo']} con {args.modelo}... (tarda unos minutos)")
    cliente = anthropic.Anthropic()
    with cliente.messages.stream(**parametros(v, todos, indice, cargar_prompt(), args)) as stream:
        mensaje = stream.get_final_message()
    if mensaje.stop_reason == "refusal":
        sys.exit("El modelo rechazó la petición.")
    texto = texto_de(mensaje)
    guardar(v, texto)
    u = mensaje.usage
    pe, ps = PRECIOS.get(args.modelo, (0, 0))
    costo = (u.input_tokens * pe + u.output_tokens * ps) / 1e6
    print(f"Guardado en {ruta_resultado(v)}")
    print(f"Tokens: entrada {u.input_tokens:,}, salida {u.output_tokens:,}. "
          f"Costo de esta prueba: ~US$ {costo:.2f} (con Batch API sería la mitad).")
    if mensaje.stop_reason == "max_tokens":
        print("AVISO: la respuesta se cortó por longitud.")
    defectos = problemas(texto, args.min_palabras)
    print(f"Palabras: {len(texto.split()):,}. Control de calidad: "
          + ("OK" if not defectos else "INCOMPLETO — " + " | ".join(defectos)))


def cmd_enviar(args):
    todos = cargar_versiculos()
    lotes = leer_lotes()
    en_curso = {cid for l in lotes if not l.get("descargado") for cid in l["ids"]}
    pendientes = [(i, v) for i, v in filtrar(todos, args.libro)
                  if id_de(v) not in en_curso and not hecho(v, args.min_palabras)]
    if args.max:
        pendientes = pendientes[:args.max]
    if not pendientes:
        print("No hay versículos pendientes de enviar.")
        return
    sistema = cargar_prompt()
    cliente = anthropic.Anthropic()
    for inicio in range(0, len(pendientes), POR_LOTE):
        grupo = pendientes[inicio:inicio + POR_LOTE]
        lote = cliente.messages.batches.create(requests=[
            Request(custom_id=id_de(v),
                    params=MessageCreateParamsNonStreaming(**parametros(v, todos, i, sistema, args)))
            for i, v in grupo
        ])
        lotes.append({"id": lote.id, "modelo": args.modelo, "ids": [id_de(v) for _, v in grupo],
                      "creado": time.strftime("%Y-%m-%d %H:%M")})
        guardar_lotes(lotes)
        print(f"Lote {lote.id} enviado con {len(grupo)} versículos.")
    print("Los lotes suelen tardar entre 1 y 24 horas. Luego ejecuta: python generar.py descargar")


def cmd_descargar(args):
    todos = {id_de(v): v for v in cargar_versiculos()}
    lotes = leer_lotes()
    cliente = anthropic.Anthropic()
    for l in lotes:
        if l.get("descargado"):
            continue
        lote = cliente.messages.batches.retrieve(l["id"])
        if lote.processing_status != "ended":
            c = lote.request_counts
            print(f"{l['id']}: en proceso ({c.succeeded} listos, {c.processing} procesando)")
            continue
        ok, fallos, incompletos = 0, [], []
        for r in cliente.messages.batches.results(l["id"]):
            v = todos[r.custom_id]
            ref = f"{v['libro']} {v['capitulo']}:{v['versiculo']}"
            if r.result.type != "succeeded":
                fallos.append(f"{ref} ({r.result.type})")
                continue
            m = r.result.message
            if m.stop_reason == "refusal":
                fallos.append(f"{ref} (rechazo)")
                continue
            texto = texto_de(m)
            defectos = problemas(texto, args.min_palabras)
            if m.stop_reason == "max_tokens":
                defectos.insert(0, "cortado por longitud")
            ruta = ruta_resultado(v)
            # Un resultado defectuoso no reemplaza a uno que ya estaba bien
            if defectos and ruta.exists() and not problemas(ruta.read_text(encoding="utf-8"), args.min_palabras):
                incompletos.append(f"{ref} (se conserva el anterior: {defectos[0]})")
                continue
            guardar(v, texto)
            if defectos:
                incompletos.append(f"{ref} ({defectos[0]})")
            else:
                ok += 1
        l["descargado"] = True
        guardar_lotes(lotes)
        print(f"{l['id']}: {ok} completos, {len(incompletos)} incompletos, {len(fallos)} fallidos.")
        for titulo, lista in (("Fallidos", fallos), ("Incompletos", incompletos)):
            if lista:
                print(f"  {titulo} (se reenviarán con 'enviar'):", ", ".join(lista[:20]),
                      "..." if len(lista) > 20 else "")


def cmd_estado(args):
    todos = cargar_versiculos()
    por_libro = {}
    for v in todos:
        t = por_libro.setdefault(v["libro"], [0, 0, 0])  # completos, incompletos, total
        t[2] += 1
        if ruta_resultado(v).exists():
            t[0 if hecho(v, args.min_palabras) else 1] += 1
    hechos = sum(t[0] for t in por_libro.values())
    malos = sum(t[1] for t in por_libro.values())
    print(f"Completos: {hechos:,} de {len(todos):,} ({100 * hechos / len(todos):.2f} %)")
    print(f"Incompletos (hay que regenerar): {malos:,}")
    abiertos = [l for l in leer_lotes() if not l.get("descargado")]
    print(f"Lotes sin descargar: {len(abiertos)}")
    for libro, (h, m, t) in por_libro.items():
        if h or m:
            print(f"  {libro}: {h}/{t} completos, {m} incompletos")


def cmd_revisar(args):
    todos = cargar_versiculos()
    malos = []
    for _, v in filtrar(todos, args.libro):
        ruta = ruta_resultado(v)
        if ruta.exists():
            defectos = problemas(ruta.read_text(encoding="utf-8"), args.min_palabras)
            if defectos:
                malos.append((v, defectos))
    for v, defectos in malos if args.detalle else []:
        print(f"{v['libro']} {v['capitulo']}:{v['versiculo']}: " + " | ".join(defectos))
    print(f"\n{len(malos):,} versículos incompletos. Se regenerarán con: python generar.py enviar"
          + (f' --libro "{args.libro}"' if args.libro else "")
          + " (la versión vieja queda en resultados_respaldo\\)")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="comando", required=True)
    for nombre in ("estimar", "prueba", "enviar", "descargar", "estado", "revisar"):
        s = sub.add_parser(nombre)
        s.add_argument("--modelo", default=MODELO)
        s.add_argument("--esfuerzo", default=ESFUERZO, choices=["low", "medium", "high", "xhigh", "max"])
        s.add_argument("--min-palabras", type=int, default=MIN_PALABRAS)
        if nombre in ("estimar", "enviar", "revisar"):
            s.add_argument("--libro")
        if nombre == "revisar":
            s.add_argument("--detalle", action="store_true")
        if nombre == "enviar":
            s.add_argument("--max", type=int)
        if nombre == "prueba":
            s.add_argument("referencia")
    args = p.parse_args()
    globals()[f"cmd_{args.comando}"](args)


if __name__ == "__main__":
    main()
