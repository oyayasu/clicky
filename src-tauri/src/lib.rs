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

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![save_file_dialog])
        .run(tauri::generate_context!())
        .expect("Falha ao iniciar Clicky");
}