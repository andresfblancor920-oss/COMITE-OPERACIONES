# Commercial Review (Kingspan Middle Americas): Data Engine → deck

Flujo mensual del comité comercial con la **nueva estructura de canales y segmentos**:

```
plantilla/Kingspan_Commercial_Data_Engine.xlsx      ← se actualiza cada mes (celdas amarillas)
plantilla/Deck_Commercial_Review_Plantilla.pptx     ← deck actualizado (sirve de base el mes siguiente)
scripts/actualizar_deck_comercial.py                ← regenera los slides desde el Excel
scripts/crear_data_engine_comercial.py              ← (re)crea el Data Engine desde el original + Presupuesto Junta
plantilla/fuentes/                                  ← archivos de origen (Data Engine OCT original y Presupuesto Junta)
```

```bash
python scripts/actualizar_deck_comercial.py plantilla/Kingspan_Commercial_Data_Engine.xlsx plantilla/Deck_Commercial_Review_Plantilla.pptx salidas/Commercial_Review_<mes>.pptx
```

El script es idempotente (se puede correr sobre el deck del mes anterior). Si el Excel no tiene
valores calculados (nunca se abrió en Excel), se recalcula solo con LibreOffice.

### Reglas de la nueva estructura (hoja `22_Canales_Segmentos`)

| Concepto | Regla |
|---|---|
| Ventas Alternativas | Valor del canal actual |
| Industrial + Distribución | Total IMP − Ventas Alternativas, prorrateado con Industrial / (Industrial + Distribución) del canal actual; % manual opcional (fila 16) |
| Construcción / Refrigeración | Total de los 3 canales × % asignado manualmente (B5 general o fila 24 por métrica) |
| Presupuesto | `21_Ppto_Junta`: Presupuesto de Junta mensual × participación del presupuesto **anterior** por canal (selector B5: Anterior / Junta) |

Aplica a ventas, order intake, backlog y presupuesto.

### Slides que se generan desde el Excel

| Slide | Fuente |
|---|---|
| Commercial Overview: ventas YTD por canal y por segmento + **OI por región** ordenado por volumen | `22_Canales_Segmentos`, `19_OI_Regiones` (selector Mes/YTD) |
| Order Intake / Backlog del mes: misma estructura, con etiqueta de m² centrada en cada barra | `02_OI` (I:J), `04_Backlog` (G:H) |
| OI by channel / Backlog by channel: 3 canales + 2 segmentos | `22_Canales_Segmentos` filas 42–63 |
| Lost Business: tarjetas KPI, tabla ejecutiva y valor perdido por motivo | `20_Lost_Business` (tabla con filtros + consulta interactiva) |
| RSM Performance by groups (budget de Junta) + resumen por canales y segmentos | `RSM_Resumen_Canales` T:W, `22_Canales_Segmentos` filas 66–75 |

Los puntos abiertos de datos están en la hoja `Validaciones`.

---

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
