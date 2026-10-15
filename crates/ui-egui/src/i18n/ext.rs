//! Translations of the labels the database and macro features add. [`Language::tr`] looks here
//! after the language's own table, so upstream translations always win.

use super::Language;

pub(super) fn translate(language: Language, text: &str) -> Option<&'static str> {
    match (language, text) {
        (Language::Es, _) => spanish(text),
        (Language::Ja, "Get\nData") => Some("データの\n取得"),
        (Language::Zh, "Get\nData") => Some("获取\n数据"),
        (Language::Ko, "Get\nData") => Some("데이터\n가져오기"),
        (Language::Ru, "Get\nData") => Some("Получить\nданные"),
        (Language::PtBr, "Get\nData") => Some("Obter\nDados"),
        (Language::De, "Get\nData") => Some("Daten\nabrufen"),
        (Language::Cs, "Get\nData") => Some("Načíst\ndata"),
        (Language::Fr, "Get\nData") => Some("Obtenir des\ndonnées"),
        (Language::Sr, "Get\nData") => Some("Preuzmi\npodatke"),
        _ => None,
    }
}

fn spanish(text: &str) -> Option<&'static str> {
    Some(match text {
        "Get\nData" => "Obtener\ndatos",
        // Get Data › From Database, Queries & Connections
        "From Text/CSV…" => "Desde el archivo de texto/CSV…",
        "From SQL Server Database…" => "Desde una base de datos de SQL Server…",
        "From Table/Range" => "Desde una tabla o rango",
        "Data Source Settings…" => "Configuración de origen de datos…",
        "Get Data from a file or a database" => "Obtener datos de un archivo o una base de datos",
        "Refresh\nAll" => "Actualizar\ntodo",
        "Refresh All" => "Actualizar todo",
        "Connection Properties…" => "Propiedades de conexión…",
        "Connection Properties" => "Propiedades de conexión",
        "Queries & Connections" => "Consultas y conexiones",
        "SQL Server database" => "Base de datos de SQL Server",
        "Data Source Settings" => "Configuración de origen de datos",
        "Query Properties" => "Propiedades de la consulta",
        "Server" => "Servidor",
        "Database (optional)" => "Base de datos (opcional)",
        "Advanced options" => "Opciones avanzadas",
        "SQL statement (optional, requires database)" => "Instrucción SQL (opcional, requiere base de datos)",
        "Encrypt connection" => "Cifrar conexión",
        "Trust the server certificate" => "Confiar en el certificado del servidor",
        "You can trust this server's certificate (only on a network you trust), or ask the server's administrator to renew it." => {
            "Puede confiar en el certificado de este servidor (solo en una red de confianza) o pedir al administrador del servidor que lo renueve."
        }
        "Trust the certificate and connect" => "Confiar en el certificado y conectar",
        "OK" => "Aceptar",
        "Cancel" => "Cancelar",
        "Enter a server name." => "Escriba el nombre de un servidor.",
        "A SQL statement needs a database." => "Una instrucción SQL necesita una base de datos.",
        "Sign in to" => "Iniciar sesión en",
        "Password" => "Contraseña",
        "Show password" => "Mostrar contraseña",
        "Connect" => "Conectar",
        "Database" => "Base de datos",
        "Search" => "Buscar",
        "Select multiple items" => "Seleccionar varios elementos",
        "Choose a database." => "Elija una base de datos.",
        "This database has no tables or views you can read." => "Esta base de datos no tiene tablas ni vistas que pueda leer.",
        "Select a table or view to preview its data." => "Seleccione una tabla o vista para obtener una vista previa de sus datos.",
        "Load" => "Cargar",
        "rows loaded." => "filas cargadas.",
        "No database logins are kept in this session." => "No hay inicios de sesión de bases de datos guardados en esta sesión.",
        "Edit Permissions…" => "Editar permisos…",
        "Clear Permissions" => "Borrar permisos",
        "This workbook has no queries. Use Get Data to load data from a database." => {
            "Este libro no tiene consultas. Use Obtener datos para cargar datos de una base de datos."
        }
        "query" => "consulta",
        "queries" => "consultas",
        "Last refreshed" => "Última actualización",
        "Properties…" => "Propiedades…",
        // Writing to a database
        "Write to a database" => "Escribir en una base de datos",
        "Save Changes to Database…" => "Guardar cambios en la base de datos…",
        "Export to SQL Server…" => "Exportar a SQL Server…",
        "Run SQL…" => "Ejecutar SQL…",
        "Save Changes to Database" => "Guardar cambios en la base de datos",
        "Export to SQL Server" => "Exportar a SQL Server",
        "Run SQL" => "Ejecutar SQL",
        "Another server…" => "Otro servidor…",
        "Rows to add" => "Filas que se añadirán",
        "Rows to change" => "Filas que se modificarán",
        "Rows to delete" => "Filas que se eliminarán",
        "Nothing has changed since the last refresh." => "No ha cambiado nada desde la última actualización.",
        "Everything is sent in one transaction: all of it or nothing. Undo in GridCraft doesn't undo it on the server." => {
            "Todo se envía en una sola transacción: todo o nada. Deshacer en GridCraft no lo deshace en el servidor."
        }
        "rows saved to the database." => "filas guardadas en la base de datos.",
        "Table name (schema.table)" => "Nombre de la tabla (esquema.tabla)",
        "New table" => "Tabla nueva",
        "Add rows" => "Añadir filas",
        "Replace all rows" => "Reemplazar todas las filas",
        "Range (blank: the table or region at the active cell)" => "Rango (vacío: la tabla o región de la celda activa)",
        "rows" => "filas",
        "Every row now in the table is deleted first." => "Primero se eliminan todas las filas que hay ahora en la tabla.",
        "rows exported to" => "filas exportadas a",
        "Preview" => "Vista previa",
        "SQL statement" => "Instrucción SQL",
        "Cells for @P1, @P2… (optional, for example B2, C2)" => "Celdas para @P1, @P2… (opcional, por ejemplo B2, C2)",
        "This runs on the server in one transaction and can't be undone from GridCraft. Run it?" => {
            "Se ejecuta en el servidor en una sola transacción y no se puede deshacer desde GridCraft. ¿Ejecutar?"
        }
        "Run" => "Ejecutar",
        "rows affected." => "filas afectadas.",
        "Save changes…" => "Guardar cambios…",
        "User name" => "Nombre de usuario",
        "SQL Server login" => "Inicio de sesión de SQL Server",
        "Windows (another account)" => "Windows (otra cuenta)",
        "Windows (current user)" => "Windows (usuario actual)",
        "GridCraft signs in as the Windows account you're using; no password is asked." => {
            "GridCraft inicia sesión con la cuenta de Windows que está usando; no se pide contraseña."
        }
        "GridCraft keeps the login for this session and never saves it in the workbook." => {
            "GridCraft guarda el inicio de sesión durante esta sesión y nunca lo guarda en el libro."
        }
        "Remember on this computer" => "Recordar en este equipo",
        "The login couldn't be remembered:" => "No se pudo recordar el inicio de sesión:",
        // Macros
        "Macro" => "Macro",
        "Running macro" => "Ejecutando la macro",
        "Stop macro" => "Detener la macro",
        "Macro finished:" => "Macro terminada:",
        "This workbook has no macros." => "Este libro no tiene macros.",
        "GridCraft scripts" => "Scripts de GridCraft",
        "Yes" => "Sí",
        "No" => "No",
        "Abort" => "Anular",
        "Retry" => "Reintentar",
        "Ignore" => "Omitir",
        "Enable Content" => "Habilitar contenido",
        "Macros enabled" => "Macros habilitadas",
        "Macros are disabled in this workbook. Click Enable Content in the security bar to run the button's macro." => {
            "Las macros están deshabilitadas en este libro. Pulse Habilitar contenido en la barra de seguridad para ejecutar la macro del botón."
        }
        "Macros disabled" => "Macros deshabilitadas",
        "SECURITY WARNING Macros have been disabled. Enable them only if you trust where this workbook came from." => {
            "ADVERTENCIA DE SEGURIDAD Las macros se han deshabilitado. Habilítelas solo si confía en el origen de este libro."
        }
        _ => return None,
    })
}
