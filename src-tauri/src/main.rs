// SPDX-License-Identifier: MIT OR Apache-2.0
// graphPaper as a desktop app: the same pages as the site, with the engine
// bundled, in a window of its own. Files are saved through the system's save
// dialog; nothing goes over the network.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .run(tauri::generate_context!())
        .expect("graphPaper failed to start");
}
