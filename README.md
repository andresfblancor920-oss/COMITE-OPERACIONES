# Comité de Operaciones — Plantilla semanal → PowerPoint automático

Este proyecto entrega un **flujo repetible**: cada semana se actualiza un
único archivo Excel con la información operativa, y con un solo comando se
genera automáticamente el PowerPoint listo para presentar ante el comité
operacional y la dirección — mismo diseño, mismos colores, mismo orden,
todas las semanas.

```
plantilla/Plantilla_Semanal_Comite_Operaciones.xlsx   ← se llena cada semana
scripts/generar_ppt_comite.py                         ← genera el .pptx
scripts/crear_plantilla_excel.py                      ← (re)genera la plantilla Excel
ejemplos/Ejemplo_Comite_Operaciones_Semana.pptx        ← muestra de referencia (con los datos de ejemplo)
salidas/                                               ← aquí quedan los .pptx generados cada semana (no se versiona)
```

## Flujo semanal

1. Abra `plantilla/Plantilla_Semanal_Comite_Operaciones.xlsx`.
2. Actualice **solo las celdas resaltadas en amarillo** de cada hoja
   (Portada, Resumen Ejecutivo, KPIs, Proyectos, Incidentes, Riesgos,
   Plan de Acción, Próximos Pasos). Las columnas de variación, % de
   cumplimiento, estado (semáforo) y nivel de riesgo se calculan solas
   con fórmulas.
3. Guarde el archivo (con Excel o LibreOffice, para que las fórmulas
   queden calculadas).
4. Ejecute:

   ```bash
   pip install -r requirements.txt        # una sola vez
   python scripts/generar_ppt_comite.py plantilla/Plantilla_Semanal_Comite_Operaciones.xlsx
   ```

5. El script crea automáticamente `salidas/Comite_Operaciones_<fecha>.pptx`
   con las diapositivas: Portada, Resumen Ejecutivo, KPIs (tabla y
   gráfico), Proyectos, Incidentes, Riesgos, Plan de Acción y Próximos
   Pasos — con semáforos de color (verde/amarillo/rojo) calculados a
   partir de los datos.

No hay que tocar el PowerPoint manualmente: si cambia el Excel y se
vuelve a correr el script, las diapositivas se regeneran solas. El diseño
visual vive en `scripts/generar_ppt_comite.py`, así que un cambio de
marca/colores/logo se hace una sola vez ahí y aplica a todas las semanas
futuras.

### Personalizar la plantilla

`scripts/crear_plantilla_excel.py` es el generador de la plantilla Excel.
Si se necesita agregar una columna, una hoja nueva o cambiar un KPI por
defecto, se edita ese script y se vuelve a correr (esto crea una plantilla
en blanco nueva; no lo corra sobre un archivo ya lleno de datos de la
semana, porque lo sobrescribe).

---

## Otros reportes de análisis recomendados para el comité

La plantilla actual cubre el "pulso semanal" (KPIs, proyectos, incidentes,
riesgos, compromisos). Para una presentación más completa ante dirección,
conviene complementarla — no necesariamente cada semana, pero sí con una
cadencia definida — con:

| Reporte | Frecuencia sugerida | Qué aporta |
|---|---|---|
| **Dashboard de tendencias históricas** (mismos KPIs pero en serie de 8–12 semanas, no solo semana vs. semana anterior) | Mensual | Distingue una variación puntual de una tendencia real; evita decisiones reactivas por un solo mal dato |
| **Análisis de causa raíz (RCA)** de incidentes mayores (severidad alta) | Por evento | Dirección no solo necesita saber "qué pasó" sino "por qué" y "qué garantiza que no se repita" |
| **Reporte de capacidad y utilización de recursos** (infraestructura, equipos, personas) | Mensual | Anticipa cuellos de botella antes de que se conviertan en incidentes |
| **SLA/OLA scorecard** con contrapartes internas y proveedores externos | Mensual/trimestral | Consolida cumplimiento contractual, no solo operativo |
| **Costos operativos vs. presupuesto** (OPEX real vs. plan, por área) | Mensual | Conecta la operación con el impacto financiero, lenguaje que dirección valora especialmente |
| **Satisfacción del cliente interno/externo** (CSAT/NPS) | Mensual | Mide el resultado percibido, no solo el cumplimiento técnico |
| **Reporte de cumplimiento normativo/seguridad de la información** | Trimestral | Cubre un riesgo reputacional y regulatorio que rara vez aparece en el pulso semanal |
| **Benchmarking vs. metas anuales / OKRs** | Trimestral | Sitúa el desempeño semanal dentro del objetivo del año, útil para revisiones de dirección |
| **Análisis de rotación y clima del equipo operativo** | Trimestral | El desempeño operacional depende de la estabilidad del equipo; suele omitirse y es un riesgo silencioso |
| **Forecast / proyección a 2–4 semanas** (no solo lo ya ocurrido) | Semanal o quincenal | Cambia el comité de "mirar el retrovisor" a anticipar decisiones |

## Información adicional que valdría la pena comunicar

- **Semáforo consolidado de un solo vistazo** (una diapositiva "resumen
  ejecutivo en 1 minuto" con 4–5 métricas clave y su tendencia ↑↓, antes
  de entrar al detalle) — ya incluida en la plantilla, pero conviene
  mantenerla siempre como la primera diapositiva de contenido.
- **Comparación contra la meta anual/trimestral**, no solo contra la
  semana anterior, para que dirección entienda si se va bien o mal
  respecto al objetivo del período, no solo respecto al último dato.
- **Impacto en el negocio**, traduciendo métricas técnicas a lenguaje de
  negocio (p. ej. "4 horas de indisponibilidad ≈ X en ingresos no
  procesados" o "SLA incumplido en proveedor clave ≈ riesgo de penalidad
  contractual").
- **Estado de decisiones/compromisos abiertos de comités anteriores**,
  para dar seguimiento y evitar que el plan de acción se "reinicie" cada
  semana (la hoja "Plan de Acción" ya soporta esto si se arrastran las
  filas pendientes semana a semana).
- **Riesgos emergentes** (aquellos que aún no se materializan pero se
  ven venir: vencimiento de contratos, fin de soporte de un sistema,
  dependencia de una sola persona clave, etc.), separados de los
  incidentes ya ocurridos.
- **Lecciones aprendidas** de incidentes cerrados, para mostrar que la
  operación mejora de forma continua y no solo reacciona.
- **Indicadores de seguridad laboral/operacional**, si aplica al tipo de
  operación (accidentabilidad, cumplimiento de protocolos), un tema que
  comités de dirección suelen exigir de forma explícita.

Estas adiciones se pueden incorporar como hojas nuevas en el Excel (o
columnas adicionales en las existentes) y diapositivas nuevas en el
script, siguiendo el mismo patrón que las secciones ya implementadas.

---

## Procedimiento de Información de Entrega y Despacho KAIR

`documentos/procedimiento_entrega_despacho/Procedimiento_Entrega_Despacho_KAIR.pdf`
formaliza, ante clientes y equipo comercial, la información que acompaña
cada despacho del módulo habitacional KAIR by Kingspan: remisión(es),
lista detallada de empaque (Packing List), información del transporte
por correo electrónico y registro fotográfico. Sigue el formato de la
Política de Devolución de Producto KAIR e incluye como anexos una guía
rápida y los ejemplos de remisión y Packing List.

Para modificarlo, edite `documentos/procedimiento_entrega_despacho/fuente/procedimiento.html`
y regenere el PDF con:

```bash
pip install playwright
python scripts/generar_pdf_procedimiento_entrega.py
```
