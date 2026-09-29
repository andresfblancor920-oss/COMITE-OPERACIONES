"""Genera el PDF del Procedimiento de Información de Entrega y Despacho KAIR.

Uso:
    pip install playwright
    python scripts/generar_pdf_procedimiento_entrega.py [ruta_chromium]

El contenido y el diseño viven en
documentos/procedimiento_entrega_despacho/fuente/procedimiento.html;
se edita ese archivo y se vuelve a correr este script.
"""
import pathlib
import sys

from playwright.sync_api import sync_playwright

RAIZ = pathlib.Path(__file__).resolve().parent.parent / "documentos" / "procedimiento_entrega_despacho"
FUENTE = RAIZ / "fuente" / "procedimiento.html"
SALIDA = RAIZ / "Procedimiento_Entrega_Despacho_KAIR.pdf"


def main():
    opciones = {"executable_path": sys.argv[1]} if len(sys.argv) > 1 else {}
    with sync_playwright() as p:
        navegador = p.chromium.launch(**opciones)
        pagina = navegador.new_page()
        pagina.goto(FUENTE.as_uri())
        pagina.wait_for_load_state("networkidle")
        pagina.evaluate("document.fonts.ready")
        pagina.pdf(path=str(SALIDA), width="8.5in", height="11in", print_background=True)
        navegador.close()
    print(f"PDF generado: {SALIDA}")


if __name__ == "__main__":
    main()
