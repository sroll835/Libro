"""Herramientas para generar con Claude Code (sin API key).

  python control.py prompt                  Escribe prompt_completo.md (prompt + reglas de formato y profundidad)
  python control.py pasaje 36 27            Muestra el pasaje a decodificar con su contexto y la ruta de salida
  python control.py check 36 27             Control de calidad del archivo de ese versículo (OK o lista de defectos)
  python control.py pendientes [N]          Próximos N versículos de Génesis sin resultado completo
"""
import sys
from pathlib import Path

import generar as g

LIBRO = "Génesis"
AQUI = Path(__file__).resolve().parent


def buscar(cap, ver):
    todos = g.cargar_versiculos()
    for i, v in enumerate(todos):
        if v["libro"] == LIBRO and v["capitulo"] == cap and v["versiculo"] == ver:
            return todos, i, v
    sys.exit(f"No existe {LIBRO} {cap}:{ver}")


def main():
    cmd = sys.argv[1]
    if cmd == "prompt":
        (AQUI / "prompt_completo.md").write_text(g.cargar_prompt(), encoding="utf-8")
        print("prompt_completo.md escrito")
    elif cmd == "pasaje":
        todos, i, v = buscar(int(sys.argv[2]), int(sys.argv[3]))
        print(g.mensaje_usuario(v, todos, i))
        print(f"\nARCHIVO DE SALIDA: {g.ruta_resultado(v)}")
    elif cmd == "check":
        _, _, v = buscar(int(sys.argv[2]), int(sys.argv[3]))
        ruta = g.ruta_resultado(v)
        if not ruta.exists():
            print("NO EXISTE " + str(ruta))
            return
        texto = ruta.read_text(encoding="utf-8")
        defectos = g.problemas(texto)
        print(f"{len(texto.split())} palabras. " + ("OK" if not defectos else "DEFECTOS: " + " | ".join(defectos)))
    elif cmd == "pendientes":
        n = int(sys.argv[2]) if len(sys.argv) > 2 else 20
        faltan = [f"{v['capitulo']}:{v['versiculo']}" for v in g.cargar_versiculos()
                  if v["libro"] == LIBRO and not g.hecho(v)]
        print(f"{len(faltan)} pendientes. Próximos: " + " ".join(faltan[:n]))


if __name__ == "__main__":
    main()
