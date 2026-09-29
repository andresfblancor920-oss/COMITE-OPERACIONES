# Política de Devolución de Productos — KAIR by Kingspan

Documento formal para clientes y equipo comercial de KAIR, construido sobre la
misma estructura y condiciones de la *Política de Devolución de Producto*
de Kingspan Paneles Aislados S.A.S. (Versión 1 · 04-09-2024).

| Archivo | Uso |
|---|---|
| `Politica_Devolucion_KAIR.pdf` | Documento final para enviar/publicar (8 páginas, tamaño carta) |
| `Plantilla_Packing_List_KAIR.xlsx` | Lista detallada (Packing List) editable + hoja de verificación remisión vs. lista detallada, con ejemplo del módulo muestra |
| `crear_packing_list_excel.py` | Regenera la plantilla Excel del Packing List (la sobrescribe) |
| `politica_devolucion_kair.html` | Fuente editable (textos, colores, versión y fecha) |
| `generar_pdf.js` | Regenera el PDF desde el HTML |
| `assets/` | Imágenes KAIR (portada, render exterior, planta, interior) |
| `fuentes/` | Tipografías Quicksand y Montserrat embebidas |

Estructura: 1. Condiciones generales · 2. Procedimiento de devolución ·
3. Procesamiento de devoluciones · 4. Envíos incorrectos ·
5. Entrega y documentación de cada despacho (remisión, lista detallada /
Packing List, información de transporte, registro fotográfico, verificación
en la entrega y procedimiento paso a paso con el cruce remisión vs. lista
detallada) · 6. Contacto ·
Anexo A: guía rápida de devolución · Anexo B: ejemplo visual de remisión ·
Anexo C: ejemplo visual de lista detallada (Packing List).

Los ejemplos de los anexos B y C corresponden al módulo habitacional muestra
(remisión N° 22390); los datos personales (contacto, teléfono, conductor,
nombres y firma) se omitieron en `assets/ejemplo_*.png`.

Correo oficial de devoluciones: **devolucioneskair@kingspan.com**

## Regenerar el PDF

```bash
NODE_PATH=$(npm root -g) node politicas/kair_devoluciones/generar_pdf.js
```

Requiere Node.js con Playwright (Chromium). La plantilla Excel se regenera con
`python politicas/kair_devoluciones/crear_packing_list_excel.py` (requiere `openpyxl`). Para una nueva versión, actualice
el texto "Versión 1 - 29-09-2026" en los pies de página del HTML.
