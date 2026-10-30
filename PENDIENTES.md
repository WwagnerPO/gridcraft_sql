# Pendientes de nuestra versión (SQL Server + macros VBA)

Lista de trabajo de lo propio. Lo del GridCraft oficial está en `docs/gaps.md` (no se toca).
Marcar `[x]` al terminar y anotar el commit.

Actualizado: 2026-10-30.

## Integración con el oficial

- [ ] **Macros y "Trust" del oficial.** El oficial añadió `Trust` en cada documento
  (`engine/src/lib.rs`: `Local` / `FromFile`, `allows_code()`). Nuestras macros llevan su propio
  "Habilitar contenido" en `ui-egui/src/ext.rs`. Unificarlos: que Habilitar contenido marque el
  documento como de confianza y que `Workbook_Open` respete `allows_code()`.
- [ ] **Colores de los iconos de base de datos.** El oficial pasó los iconos a la paleta del tema
  (`Palette`/`Cat`); `icons/database.rs` sigue con verde y azul fijos. Pasarlos a la paleta para el
  modo oscuro.
- [ ] **Pasar `cargo xtask ci` completo** (incluye `wasm`) en la máquina de desarrollo tras la
  fusión del 2026-10-29. En la nube se pasaron fmt, clippy, layers, assets y los tests de los
  crates afectados.
- [ ] **Instalador de Windows**: lanzar `fork-windows.yml` con una etiqueta nueva
  (`v0.13.0-sql.1`) y probar el `.msi`.

## SQL Server

- [ ] **Consultas en segundo plano**: hoy la ventana espera a que termine la consulta.
- [ ] **Editor de Power Query**: filtrar, combinar y cambiar tipos antes de cargar (hoy solo con
  SQL escrito a mano).
- [ ] Guardar en el archivo la **hora de la última actualización**, en hora local (hoy UTC y no se
  guarda).
- [ ] CLI y MCP: **recordar inicios de sesión** entre ejecuciones (hoy hay que pasar usuario y
  contraseña cada vez).

## Macros VBA

- [ ] **Eventos**: solo funciona `Workbook_Open`. Faltan `Worksheet_Change`,
  `Worksheet_SelectionChange`, `Workbook_BeforeSave`, `Workbook_BeforeClose`, etc.
- [ ] **Archivos de texto**: `Open … For … As #n`, `Print #`, `Write #`, `Input #`, `Line Input`,
  `Close` (hoy: "GridCraft can't run yet").
- [ ] `GoSub` / `Return` y `AddressOf`.
- [ ] **UserForms** (formularios de VBA).
- [ ] **Editor de VBA**: ver y cambiar el código de las macros desde la app.

## Calidad

- [ ] Auditoría: topes de memoria en fórmulas (`SEQUENCE`, `REPT`, `MAKEARRAY`, columnas enteras).
- [ ] Auditoría: conversiones `as` con números del usuario (`f64 as usize`, negativos).
- [ ] Auditoría: nombres definidos que se refieren a sí mismos y `LAMBDA` recursivo.
- [ ] Parámetros hostiles (tipos erróneos, rangos enormes) en el canal de control y MCP para
  nuestros comandos (`data.*` de bases de datos, `vba.*`).
- [ ] `cargo audit` y `cargo deny` (vulnerabilidades y licencias: nada GPL).

## Descartado

- Iniciar sesión con Microsoft Entra (cuentas de Azure): no se hará (2026-10-30).

## Hecho

- [x] 2026-10-30: **conexiones guardadas en el equipo** como Excel (`.odc` en
  `Documentos/My Data Sources`): Exportar archivo de conexión, Conexiones existentes, abrir una
  pidiendo el login. Comandos `data.exportConnectionFile`, `data.existingConnections`,
  `data.fromConnectionFile`.

- [x] 2026-10-29: fusionado el oficial (1045 commits, v0.13.0) con nuestra rama y todo pasado a
  `main`. Incluye el arreglo del `#CIRC!` falso en cadenas de más de 2000 fórmulas.
