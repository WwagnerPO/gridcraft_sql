# Pendientes de nuestra versión (SQL Server + macros VBA)

Lista de trabajo de lo propio. Lo del GridCraft oficial está en `docs/gaps.md` (no se toca).
Marcar `[x]` al terminar y anotar el commit.

Actualizado: 2026-10-30.

## Integración con el oficial

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

- [ ] **Más eventos**: `Worksheet_SelectionChange`, `Worksheet_Activate`, `Workbook_BeforeSave`,
  `Workbook_BeforeClose`, `Workbook_NewSheet`… (ya funcionan `Workbook_Open`, `Worksheet_Change`
  y `Workbook_SheetChange`).
- [ ] `Application.EnableEvents` (hoy los cambios de una macro nunca disparan eventos).
- [ ] **Archivos de texto**: `Open … For … As #n`, `Print #`, `Write #`, `Input #`, `Line Input`,
  `Close` (hoy: "GridCraft can't run yet").
- [ ] `GoSub` / `Return` y `AddressOf`.
- [ ] **UserForms** (formularios de VBA).
- [ ] **Editor de VBA**: ver y cambiar el código de las macros desde la app.

## Calidad

- [ ] `cargo audit` en cada actualización del oficial. Estado 2026-10-30 (tras subir
  `webpki-root-certs` a 1.0.11):
  - `zip` 3.0.0 (RUSTSEC-2026-0367/0368/0369): fallos de `extract()` a disco; GridCraft solo lee
    en memoria y no llama a `extract`, así que no nos afecta. La dependencia es del oficial
    (`xlsx`, `cli`); se arregla cuando el oficial suba a `zip` ≥ 9.
  - Avisos sin vulnerabilidad: `tokio` (`AsyncFd`, no lo usamos; llega por `tiberius`),
    `rand` 0.7, `proc-macro-error`, `zbus_names`/`zvariant` (sin mantenimiento).
- [ ] `cargo deny` (licencias: nada GPL) — no instalado aún.
- [ ] Los 2 tests de UI que fallan sin fuentes (`interaction::column_autocomplete_folds_case_beyond_ascii`,
  `interaction::language_switch_renders_the_ribbon_and_persists`): pasar con `CRAFT_FONTS_DIR` en la
  máquina de desarrollo para confirmar que son solo del entorno.

## Descartado

- Iniciar sesión con Microsoft Entra (cuentas de Azure): no se hará (2026-10-30).

## Hecho

- [x] 2026-10-30: **eventos VBA** `Worksheet_Change` y `Workbook_SheetChange` (en la app con
  cuadros de mensaje reales; en CLI/MCP al acabar el comando). Bugs encontrados por el camino:
  - Al guardar se perdía el `codeName` de cada hoja (lo que une la hoja con su código VBA en
    Excel) → se lee y se guarda; una hoja copiada con nombre repetido lo deja para Excel.
  - `Workbook_Open` solo se buscaba en `ThisWorkbook`: en libros de Excel en español
    (`EsteLibro`) no se ejecutaba → se busca en el módulo del libro sea cual sea su nombre.

- [x] 2026-10-30: **macros y confianza**: las macros de un libro abierto de archivo no se ejecutan
  por ninguna vía (Alt+F8, botones, CLI, MCP, canal de control) hasta Habilitar contenido o
  `vba.trust` (antes `vba.run`/`vba.click` las ejecutaban sin preguntar). Usa el `Trust` del
  oficial para saber si el libro viene de archivo. Documentado en `docs/macros.md`.

- [x] 2026-10-30: **auditoría de calidad** de lo propio:
  - Macros: anidamiento profundo (paréntesis, `If`, `Not`) cerraba la app → "Expression too
    complex"; recursión con presupuesto de pila → "Out of stack space"; las macros corren en un
    hilo propio de 64 MB (llamadas recursivas profundas funcionan también en Windows).
  - Macros: un fallo interno ya no pierde el libro (el hilo devuelve siempre la sesión).
  - Macros: `Like` con varios `*` se colgaba (exponencial) → lineal.
  - Macros: `Space`/`String`/`&`/`Replace`/`Join` gigantes y `ReDim` enorme o con límites que
    desbordan → error 7 "Out of memory" en vez de agotar la memoria.
  - SQL: `data.queriesConnections` perdía la tabla de origen (clave `table` repetida) → `source`.
  - Fuzz nuevo: todos los comandos `data.*` y `vba.*` con parámetros hostiles, sin panics.
  - Comprobado sin cambios: fórmulas gigantes (`SEQUENCE`, `REPT`, `MUNIT`, `MAKEARRAY`, columnas
    enteras), nombres que se refieren a sí mismos y `LAMBDA` recursivo dan error, no crash.

- [x] 2026-10-30: **conexiones guardadas en el equipo** como Excel (`.odc` en
  `Documentos/My Data Sources`): Exportar archivo de conexión, Conexiones existentes, abrir una
  pidiendo el login. Comandos `data.exportConnectionFile`, `data.existingConnections`,
  `data.fromConnectionFile`.

- [x] 2026-10-29: fusionado el oficial (1045 commits, v0.13.0) con nuestra rama y todo pasado a
  `main`. Incluye el arreglo del `#CIRC!` falso en cadenas de más de 2000 fórmulas.
