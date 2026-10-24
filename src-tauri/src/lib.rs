#[tauri::command]
async fn save_file_dialog(filename: String, bytes: Vec<u8>) -> Result<Option<String>, String> {
    // File selection is explicitly controlled by the user. No unrestricted write command.
    if filename.len() > 255 || bytes.len() > 150 * 1024 * 1024 {
        return Err("Nome ou arquivo muito grande".into());
    }
    tauri::async_runtime::spawn_blocking(move || {
        let Some(path) = rfd::FileDialog::new().set_file_name(&filename).save_file() else {
            return Ok(None);
        };
        std::fs::write(&path, bytes).map_err(|e| e.to_string())?;
        Ok(Some(path.to_string_lossy().to_string()))
    }).await.map_err(|e| e.to_string())?
}

#[tauri::command]
fn minimize_window(window: tauri::WebviewWindow) -> Result<(), String> { window.minimize().map_err(|e| e.to_string()) }
#[tauri::command]
fn toggle_maximize_window(window: tauri::WebviewWindow) -> Result<(), String> {
    if window.is_maximized().map_err(|e| e.to_string())? { window.unmaximize().map_err(|e| e.to_string()) }
    else { window.maximize().map_err(|e| e.to_string()) }
}
#[tauri::command]
fn close_window(window: tauri::WebviewWindow) -> Result<(), String> { window.close().map_err(|e| e.to_string()) }
#[tauri::command]
fn start_dragging_window(window: tauri::WebviewWindow) -> Result<(), String> { window.start_dragging().map_err(|e| e.to_string()) }

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![save_file_dialog, minimize_window, toggle_maximize_window, close_window, start_dragging_window])
        .run(tauri::generate_context!())
        .expect("Falha ao iniciar Clicky");
}