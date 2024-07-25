use std::path::PathBuf;

use serde::Deserialize;

#[derive(Debug, Deserialize)]
pub struct FitykSettings {
    pub columns: Vec<String>,
    pub filenames: Vec<String>,
}

#[derive(Debug, Deserialize)]
pub struct FolderOptions {
    pub read: Option<PathBuf>,
    pub write: Option<PathBuf>,
}

#[derive(Debug, Deserialize)]
pub struct AppOptions {
    pub folder: FolderOptions,
    pub create_csv: bool,
}

impl AppOptions {
    pub fn create_dir(&self) -> std::io::Result<()> {
        if let Some(path) = self.folder.write.as_ref() {
            std::fs::create_dir_all(path)?;
        };

        Ok(())
    }
}

#[derive(Debug, Deserialize)]
pub struct Settings {
    pub settings: FitykSettings,
    pub options: AppOptions,
}

pub fn toml_to_settings(toml_string: &str) -> Result<Settings, toml::de::Error> {
    let settings: Settings = toml::from_str(toml_string)?;

    Ok(settings)
}
