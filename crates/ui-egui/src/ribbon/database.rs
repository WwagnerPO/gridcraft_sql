//! The Data tab's database groups: Get Data (files and SQL Server), Queries & Connections and
//! Database (write back). `ribbon.rs` draws them with one call.

use super::controls::{control_keytip, menu_items, sep};
use super::{Icon, SheetApp, Ui, big_button, json, small_button};

/// Get Data, Queries & Connections and Database, in place of the Text/CSV-only Get Data button.
pub(super) fn data_groups(app: &mut SheetApp, ui: &mut Ui) {
    let gd = big_button(ui, Icon::Database, "Get\nData", "Get Data from a file or a database", true);
    control_keytip(app, ui, &gd, "data.getData");
    egui::Popup::menu(&gd).show(|ui| {
        menu_items(
            app,
            ui,
            &[
                ("From Text/CSV…", "dialog:open", json!({})),
                ("From SQL Server Database…", "dialog:sqlServer", json!({})),
                ("From Table/Range", "data.fromTableRange", json!({})),
                ("-", "", json!({})),
                ("Existing Connections…", "dialog:existingConnections", json!({})),
                ("Data Source Settings…", "dialog:dataSourceSettings", json!({})),
            ],
        );
    });
    sep(ui);
    // Queries & Connections
    let ra = big_button(ui, Icon::Refresh, "Refresh\nAll", "Refresh All", true);
    control_keytip(app, ui, &ra, "data.refreshAll");
    egui::Popup::menu(&ra).show(|ui| {
        menu_items(
            app,
            ui,
            &[
                ("Refresh All", "data.refreshAll", json!({})),
                ("Refresh", "data.refresh", json!({})),
                ("-", "", json!({})),
                ("Connection Properties…", "dialog:connectionProperties", json!({})),
                ("Export Connection File…", "dialog:exportConnectionFile", json!({})),
            ],
        );
    });
    ui.vertical(|ui| {
        let __k = small_button(ui, Icon::Database, "Queries & Connections", "Queries & Connections", false);
        control_keytip(app, ui, &__k, "data.queriesConnections");
        if __k.clicked() {
            app.grid.pane = if app.grid.pane.as_deref() == Some("queries") { None } else { Some("queries".into()) };
        }
        let __k = small_button(ui, Icon::Settings, "Properties", "Connection Properties", false);
        control_keytip(app, ui, &__k, "data.connectionProperties");
        if __k.clicked() {
            app.open_dialog("connectionProperties", json!({}));
        }
    });
    let db = big_button(ui, Icon::Database, "Database", "Write to a database", true);
    egui::Popup::menu(&db).show(|ui| {
        menu_items(
            app,
            ui,
            &[
                ("Save Changes to Database…", "dialog:saveToDatabase", json!({})),
                ("Export to SQL Server…", "dialog:exportToSqlServer", json!({})),
                ("Run SQL…", "dialog:sqlExecute", json!({})),
            ],
        );
    });
}
