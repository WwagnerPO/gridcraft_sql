//! Database logins remembered in the system's credential store (Windows Credential Manager,
//! macOS Keychain, the Secret Service on Linux), for Get Data's "Remember on this computer".

use gridcraft_ui_egui::LoginStore;

/// The service name the entries are filed under.
const SERVICE: &str = "GridCraft database login";

pub struct SystemLogins;

fn entry(server: &str) -> Option<keyring::Entry> {
    keyring::Entry::new(SERVICE, &server.trim().to_lowercase()).ok()
}

impl LoginStore for SystemLogins {
    fn load(&self, server: &str) -> Option<(String, String)> {
        let secret = entry(server)?.get_password().ok()?;
        let v: serde_json::Value = serde_json::from_str(&secret).ok()?;
        Some((v.get("user")?.as_str()?.to_string(), v.get("password")?.as_str()?.to_string()))
    }

    fn save(&self, server: &str, user: &str, password: &str) -> Result<(), String> {
        let e = entry(server).ok_or_else(|| "the system's credential store isn't available".to_string())?;
        e.set_password(&serde_json::json!({"user": user, "password": password}).to_string()).map_err(|e| e.to_string())
    }

    fn forget(&self, server: &str) {
        if let Some(e) = entry(server) {
            let _ = e.delete_credential();
        }
    }
}
