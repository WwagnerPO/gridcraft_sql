# Versión propia de GridCraft (SQL Server + macros VBA)

Esta copia es GridCraft oficial (`origin` = `storytold/gridcraft`) más dos funciones propias:

- **Bases de datos**: Get Data › From SQL Server, Navigator, Refresh, escribir en la base
  (Save Changes / Export / Run SQL), inicios de sesión recordados, Power Query de Excel.
- **Macros VBA**: leer `vbaProject.bin`, intérprete VBA, modelo de objetos de Excel, ejecutar
  macros desde la ventana (Alt+F8, Habilitar contenido → `Workbook_Open`).

## Ramas

- `main`: nuestra versión = oficial + nuestros commits. Aquí se trabaja.
- `feat/sql-server-connector`: histórica (hasta 2026-10-29); ya está fusionada en `main`.

Remotos: `origin` = oficial (solo descargar), `mio` = https://github.com/WwagnerPO/gridcraft_sql
(nuestra copia pública; en GitHub Actions sólo corre nuestro workflow del instalador).

## Actualizar desde el oficial

```bash
git tag backup-$(date +%F)              # vuelta atrás: git reset --hard backup-AAAA-MM-DD
git fetch origin
git checkout main && git merge origin/main
cargo xtask ci                          # en esta máquina: WGPU_BACKEND=gl cargo xtask ci
git push mio main
```

`git config rerere.enabled true` (ya activado) recuerda cómo se resolvió cada conflicto.

Si hay conflictos:
- `Cargo.lock`: no resolver a mano → `git checkout --theirs Cargo.lock && cargo update -w`.
- `keytips/keytip_table/generated.rs`: es generado; quedarse con `keytips.json` resuelto y
  regenerarlo.
- Cualquier otro: casi siempre es una línea de enganche (ver tabla); conservar la línea del
  oficial y volver a añadir la nuestra.
- El oficial parte archivos grandes en carpetas (p. ej. `dialogs.rs` → `dialogs/show.rs`). Si un
  bloque nuestro "desaparece" en el conflicto, buscar dónde vive ahora la función y poner allí el
  enganche (`git diff <base> <nuestra-rama> -- archivo` enseña qué líneas son nuestras).
- El oficial endurece clippy (`cognitive_complexity`, `too_many_arguments`, `wildcard_imports`):
  nuestro código tiene que cumplirlo igual.

## Publicar el instalador de Windows (.msi)

`.github/workflows/fork-windows.yml` compila en GitHub (Windows, MSVC, WiX 5) con el empaquetado
del oficial, renombrado a "GridCraft SQL" (se instala junto al GridCraft oficial, sin
sustituirlo), y publica el `.msi` como release. Se lanza subiendo una etiqueta:

```bash
git tag v0.7.0-sql.1 && git push mio v0.7.0-sql.1
```

Los workflows del oficial no se activan con estas etiquetas ni con la release; si alguno llega a
aparecer en la pestaña Actions (p. ej. `CI` al subir `main`), desactivarlo con
`gh workflow disable <nombre> -R WwagnerPO/gridcraft_sql`.

## Dónde está nuestro código

Casi todo vive en archivos **solo nuestros**, que el oficial no tiene y por tanto nunca chocan:

| Área | Archivos propios |
|---|---|
| Intérprete VBA | `crates/vba/` |
| Conector SQL Server | `crates/sqlserver/`, `apps/gridcraft/src/logins.rs`, `apps/gridcraft-cli/src/db.rs` |
| Modelo | `crates/model/src/connection.rs` |
| XLSX | `crates/xlsx/src/{connections,mashup,ovba,cfb_write,form_controls}.rs`, `crates/xlsx/src/cfb/streams.rs` |
| Motor | `crates/engine/src/{datasource,vba_host}.rs`, `crates/engine/src/vba_host/objects.rs`, `crates/engine/src/cmd/{connections,dbwrite,vba}.rs` |
| Interfaz | `crates/ui-egui/src/{ext,get_data,macros,controls}.rs`, `crates/ui-egui/src/{dialogs,ribbon,icons,i18n}/…` (`ext.rs`, `database.rs`) |
| Docs | `docs/sql-server.md`, `docs/macros.md`, `PENDIENTES.md`, este archivo |

En los archivos **compartidos** sólo hay líneas de enganche (cada una llama a código propio):

| Archivo compartido | Enganche |
|---|---|
| `Cargo.toml`, `*/Cargo.toml` | dependencias `gridcraft-vba`, `gridcraft-sqlserver`, `keyring` |
| `xtask/src/layers.rs` | capas de `vba` (L1) y `sqlserver` (L6) |
| `model/src/{lib,workbook}.rs` | `mod connection`, campo `Workbook::connections` |
| `xlsx/src/lib.rs` | `mod` de los módulos propios |
| `xlsx/src/{read,sheet_read,sheet_write,write}.rs`, `passthrough/write.rs` | una llamada a `connections::…` cada uno, campo `Out::connections`, `Ctx::query_tables` |
| `xlsx/src/cfb.rs` | `mod streams;` |
| `xlsx/src/{crypto,testing}.rs` | ayudantes base64 / zip |
| `engine/src/lib.rs` | `mod datasource`, `mod vba_host`, campo `Session::data`, diálogo `sqlServer` en `execute_headless` |
| `engine/src/cmd.rs` | `mod` y `specs()` de `connections`, `dbwrite`, `vba` |
| `engine/src/cmd/extra.rs`, `cmd/extra/data.rs` | Queries & Connections (descripción y `queries`) |
| `engine/src/cmd/insert/table.rs` | renombrar tabla / convertir a rango arrastran su conexión |
| `engine/src/cmd/pivot.rs`, `cmd/pivot/layout.rs` | Refresh All refresca primero las consultas |
| `ui-egui/src/lib.rs` | `mod ext`, `get_data`, `macros`, `controls`, campo `SheetApp::ext` |
| `ui-egui/src/services.rs` | `Services::logins` |
| `ui-egui/src/app.rs`, `app/{commands,files}.rs` | `ext::busy`, `ext::frame`, `ext::run_or_alert`, `ext::ui_action`, `ext::after_open` |
| `ui-egui/src/dialogs.rs`, `dialogs/show.rs` | `mod ext;`, `ext::open`, `ext::width`, `ext::owns`/`ext::show` |
| `ui-egui/src/{ribbon,icons,i18n}.rs`, `ribbon/data.rs` | `mod ext;`/`mod database;` y una línea que les llama |
| `ui-egui/src/{keymap,panes}.rs` | Alt+F8, panel Consultas y conexiones |
| `ui-egui/src/grid/show.rs`, `grid/interact.rs` | `ext::paint_controls` (dibujar botones de formulario) y `ext::click_control` (clic en un botón) |
| `ui-egui/src/tabs.rs` | estado de las macros en la barra de estado (`ext::macro_status`) |
| `ui-egui/src/keytips/*`, `tests/keytips_generated.rs` | keytips que dejan de estar "pendientes" |
| `apps/gridcraft/src/main.rs`, `apps/gridcraft-cli/src/main.rs` | instalar el proveedor SQL Server y el almacén de logins |

**Regla para código nuevo**: escribirlo en archivos propios y tocar los compartidos sólo con
una línea que lo llame. Las traducciones nuevas van en `ui-egui/src/i18n/ext.rs`, no en las
tablas de cada idioma; los docs del oficial (`ROADMAP.md`, `docs/gaps.md`…) no se tocan.
