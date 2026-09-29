# Política de Devolución de Productos — KAIR by Kingspan

Documento formal para clientes y equipo comercial de KAIR, construido sobre la
misma estructura y condiciones de la *Política de Devolución de Producto*
de Kingspan Paneles Aislados S.A.S. (Versión 1 · 04-09-2024).

| Archivo | Uso |
|---|---|
| `Politica_Devolucion_KAIR.pdf` | Documento final para enviar/publicar (4 páginas, tamaño carta) |
| `politica_devolucion_kair.html` | Fuente editable (textos, colores, versión y fecha) |
| `generar_pdf.js` | Regenera el PDF desde el HTML |
| `assets/` | Imágenes KAIR (portada, render exterior, planta, interior) |
| `fuentes/` | Tipografías Quicksand y Montserrat embebidas |

Estructura: 1. Condiciones generales · 2. Procedimiento de devolución ·
3. Procesamiento de devoluciones · 4. Envíos incorrectos · 5. Contacto ·
Anexo: guía rápida (infografía de la ruta de devolución).

Correo oficial de devoluciones: **devolucioneskair@kingspan.com**

## Regenerar el PDF

```bash
NODE_PATH=$(npm root -g) node politicas/kair_devoluciones/generar_pdf.js
```

Requiere Node.js con Playwright (Chromium). Para una nueva versión, actualice
el texto "Versión 1 - 29-09-2026" en los pies de página del HTML.
