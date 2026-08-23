# Plan de refactor por capas — app.py (O-Leasing)

Diagnóstico: `app.py` tiene **7,348 líneas** en un solo archivo. No cumple
la arquitectura por capas obligatoria (`orange-crew-standards`, sección 4:
`/api`, `/core`, `/models`, `/ui`). Este documento es el plan de cómo
dividirlo, en qué orden y cómo verificar que nada se rompe en cada paso.
No se ejecuta todo de un jalón: es software real, con dinero y contratos
reales adentro — cada fase se separa para poder probarla y hacer rollback
sola si algo sale mal.

## 0. Ya hecho (línea base de este plan)

- `models/db.py` — conexión SQLite, respaldos automáticos, chequeo de
  integridad. Ya sigue el patrón correcto (recibe `db_path` explícito, no
  lee `st.session_state`).
- `models/auth.py` — usuarios, contraseñas, roles. Mismo patrón.

## Estado del refactor (actualizado)

- **Fase 0 — hecha.** `models/db.py`, `models/auth.py`.
- **Fase 1 — hecha y probada.** `reports/excel.py`, `reports/pdf.py`
  extraídas de app.py. Se corrió una prueba real generando un Excel y dos
  PDFs con datos de ejemplo — los tres salieron con bytes válidos.
- **Fase 2 (parcial) — hecha y probada.** `models/configuracion.py`:
  `get_cfg`/`set_cfg`, catálogo de cuentas, codificación de cuentas y
  `cta_sg`. Mismo patrón de inyección (`set_resolver_conexion`) — probado
  con una base de datos SQLite real en memoria, incluyendo generación de
  número de cuenta. `app.py` bajó de 7,348 a 6,975 líneas.
- **Sigue en Fase 2** (no ejecutado todavía): el bloque de pólizas
  (`cta_sg` → `guardar`, ~400 líneas) mezcla cálculo puro con escritura
  a base de datos en funciones vecinas — ver detalle abajo.

**Recomendación:** antes de seguir, prueba a fondo lo que ya se movió
(reports, login, multiempresa, roles) con la app corriendo de verdad.
Si algo se ve raro, es mucho más fácil encontrarlo ahora, con 2 fases
movidas, que después de mover 6 fases de un jalón sin poder probarlas.

## 1. Inventario real de app.py (por qué son 7,348 líneas)

| Bloque | Líneas aprox. | Contenido |
|---|---|---|
| Config, CSS, splash/login | 1 – 700 | `st.set_page_config`, tema visual, video de bienvenida, login |
| Funciones de datos y cálculo | 700 – 4,000 | ~85 funciones: cuentas, catálogo, pólizas, amortización, conciliación CFDI, comparación de analíticas |
| Generación de reportes | 1,736 – 2,072 | Excel (`formatear_hoja_excel`, `excel_con_formato`) y PDF (`pdf_estado_cuenta`, `pdf_poliza`) |
| Menú y encabezado | 4,000 – 4,345 | `GRUPOS`, descripciones, sidebar, barra superior |
| 24 pantallas (`if/elif menu==`) | 4,345 – 7,348 | ~3,000 líneas de interfaz, una por página del menú |

Pantallas más grandes (las que más van a pesar en el refactor):

- Estado de Cuenta — 579 líneas
- Punto de Equilibrio — 424 líneas
- Dashboard & Cartera — 338 líneas
- Gestor de Bajas — 295 líneas
- Conciliación de Facturas (`_render_conciliacion`) — ya es función aparte, 805 líneas

## 2. Módulos objetivo

```
o-leasing/
  app.py                  # SOLO: config, CSS, splash/login, menú, despacho a ui.*
  core/
    amortizacion.py       # calc_amort, calc_res_amort, tablas de amortización
    polizas.py            # generar_poliza_custom, pol_inicial, pol_parcialidad, pol_comision, pol_ajuste_evento
    conciliacion.py        # parse_cfdi, conciliar_factura, detectar_duplicados
    analiticas.py          # parse_analiticas_excel, comparar_analiticas, sugerir_ajuste_poliza
    punto_equilibrio.py    # calcular_punto_equilibrio y afines
    proyecciones.py        # proy_residual, proy_rentas, proy_intereses
  models/
    db.py                  # (ya existe)
    auth.py                # (ya existe)
    contratos.py           # obtener, guardar, eliminar_contrato_completo, renumerar_contrato
    catalogo.py             # cargar/guardar_catalogo, reglas de póliza custom
    configuracion.py        # get_cfg/set_cfg, codificación de cuentas, empresas.json
    anotaciones.py          # agregar/obtener/eliminar_anotacion
    eventos.py               # registrar/obtener eventos especiales
    facturas.py              # guardar_facturas_batch, marcar_factura_cancelada, etc.
  reports/
    excel.py                # formatear_hoja_excel, excel_con_formato
    pdf.py                   # pdf_estado_cuenta, pdf_poliza, _pie_pdf_marca
  ui/
    theme.py                 # C, PAL, CSS, sfig, explain, estado_vacio, titled_chart/table
    dashboard.py              # pantalla Dashboard & Cartera
    estado_cuenta.py           # pantalla Estado de Cuenta
    carga_altas.py               # Carga Masiva y Altas
    editar_eliminar.py
    gestor_bajas.py
    morosidad.py
    eventos_especiales.py
    anotaciones.py
    polizas_contables.py
    intereses_mes.py
    facturacion_intereses.py
    conciliacion.py            # hoy _render_conciliacion, ya extraída como función
    analiticas.py               # hoy _render_analiticas
    amortizacion_tablas.py
    proyeccion.py
    rentabilidad.py
    punto_equilibrio.py
    reportes_cliente.py
    cuentas_macro.py
    contpaqi.py
    respaldo.py
    multiempresa.py
    usuarios_roles.py
```

Cada `ui/<pantalla>.py` expone una función `render()` que recibe lo que
necesite (nunca importa `app` para evitar ciclos). `app.py` termina
viéndose así en el despacho:

```python
if menu == "Dashboard & Cartera":
    ui.dashboard.render()
elif menu == "Estado de Cuenta":
    ui.estado_cuenta.render()
...
```

## 3. Orden de ejecución (fases separadas, cada una verificable sola)

1. **`reports/excel.py` y `reports/pdf.py`** — funciones puras de
   exportación, bajo acoplamiento, riesgo bajo. *(siguiente fase a ejecutar)*
2. **`core/`** — cálculos financieros (amortización, pólizas, punto de
   equilibrio, conciliación, analíticas). Riesgo medio: hay que mover
   también sus pruebas (`test_finanzas.py`) y confirmar que sigan pasando.
3. **`models/`** — el resto de las funciones de acceso a datos que hoy
   viven sueltas en app.py (contratos, catálogo, configuración,
   anotaciones, eventos, facturas).
4. **`ui/theme.py`** — helpers visuales compartidos (`sfig`, `explain`,
   `estado_vacio`, `titled_chart`, `titled_table`) y las constantes de
   color/CSS. Todas las pantallas van a importar de aquí.
5. **`ui/<pantalla>.py` uno por uno** — empezando por las más chicas y
   ya extraídas de forma natural (`Reportes por Cliente`, 16 líneas;
   `Contpaqi`, 23) para probar el patrón, subiendo hasta las grandes
   (`Estado de Cuenta`, `Punto de Equilibrio`). Ya existe el precedente
   correcto: `_render_conciliacion` y `_render_analiticas` YA están como
   funciones aparte — solo hay que moverlas de archivo, no reescribirlas.
6. **`app.py` final** — queda como orquestador: `set_page_config`, CSS,
   splash/login, menú, despacho. Debería bajar de 7,348 a ~500-700 líneas.

## 4. Verificación en cada fase

- `python -m py_compile` de todos los archivos tocados.
- Correr `test_finanzas.py` (ya existe en el proyecto) después de mover
  cualquier función de cálculo.
- Probar a mano las pantallas movidas en esa fase (abrir la app,
  navegar a cada una, confirmar que se ve y calcula igual).
- Un commit por fase (`refactor(reports): extrae excel.py y pdf.py de app.py`),
  nunca un commit gigante con las 6 fases juntas — así, si algo se rompe,
  el rollback es de una fase, no de todo el refactor.

## 5. Riesgos a vigilar

- **Constantes compartidas** (`C`, `PAL`, rutas de logos en `assets/`):
  hoy viven como variables sueltas en app.py y las usan tanto `core/`
  como `reports/` como `ui/`. Deben vivir en un solo lugar
  (`ui/theme.py` para colores, un `config.py` nuevo para rutas de
  `BASE_DIR`/`assets`) para que no haya imports circulares.
- **`st.session_state` y `st.cache_data`**: varias funciones de cálculo
  hoy están decoradas con `@st.cache_data` y dependen de `get_db_path()`
  para invalidarse cuando cambias de empresa — al mover estas funciones
  hay que llevarse el decorador y su firma tal cual, no solo el cuerpo.
- **Import circular entre `ui/*` y `core/*`**: `core/` nunca debe
  importar de `ui/`; si una función de cálculo hoy llama a `st.warning`
  o pinta algo directamente, eso es una señal de que está mezclando
  UI con lógica de negocio y hay que separarla en dos funciones al
  moverla (una calcula, otra pinta).
