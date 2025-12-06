use anyhow::Result;
use log::{debug, error};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;
use windows::Win32::Media::Audio::EndpointFormFactor;
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::UI::Controls::{HOTKEYF_ALT, HOTKEYF_CONTROL, HOTKEYF_SHIFT};
use windows::Win32::UI::Shell::{FOLDERID_RoamingAppData, KNOWN_FOLDER_FLAG, SHGetKnownFolderPath};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AudioDevice {
    pub id: String,
    pub friendly_name: String,
    // Whether this device will be included in the rotation.
    pub selectable: bool,
    #[serde(skip)]
    pub form_factor: EndpointFormFactor,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct HotkeyConfig {
    #[serde(default = "default_hotkey_vk")]
    pub vk: u8,
    #[serde(default = "default_hotkey_mods")]
    pub mods: u8,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct AppConfig {
    #[serde(default = "default_hotkey_config")]
    pub hotkey: HotkeyConfig,
    #[serde(default)]
    pub devices: HashMap<String, bool>,
}

fn default_hotkey_vk() -> u8 {
    b'X'
}

fn default_hotkey_mods() -> u8 {
    (HOTKEYF_SHIFT | HOTKEYF_CONTROL | HOTKEYF_ALT) as u8
}

fn default_hotkey_config() -> HotkeyConfig {
    HotkeyConfig {
        vk: default_hotkey_vk(),
        mods: default_hotkey_mods(),
    }
}

/// Gets the path to the user's roaming AppData directory
fn get_roaming_appdata_path() -> Result<PathBuf> {
    unsafe {
        let path_ptr =
            SHGetKnownFolderPath(&FOLDERID_RoamingAppData, KNOWN_FOLDER_FLAG::default(), None)?;

        let path_str = path_ptr.to_string()?;
        let path = PathBuf::from(path_str);

        // Free the memory allocated by SHGetKnownFolderPath
        CoTaskMemFree(Some(path_ptr.as_ptr() as *const _));

        Ok(path)
    }
}

/// Gets the full path to the AudioSwitch configuration file
pub fn get_config_file_path() -> Result<PathBuf> {
    let mut path = get_roaming_appdata_path()?;
    path.push("PurpleHatstands");
    path.push("SoundSwitcheroo");

    // Create the directory if it doesn't exist
    if !path.exists() {
        fs::create_dir_all(&path)?;
    }

    path.push("device_config.json");
    debug!("Config file path: {}", path.display());
    Ok(path)
}

/// Saves the selectable state of devices to a JSON file in the roaming AppData directory
pub fn save_config(devices: &[AudioDevice], hotkey_vk: u8, hotkey_mods: u8) -> Result<()> {
    let config_path = get_config_file_path()?;

    // Create a map of device_id -> selectable state
    let device_states: HashMap<String, bool> = devices
        .iter()
        .map(|device| (device.id.clone(), device.selectable))
        .collect();

    let config = AppConfig {
        hotkey: HotkeyConfig {
            vk: hotkey_vk,
            mods: hotkey_mods,
        },
        devices: device_states,
    };

    let json_data = serde_json::to_string_pretty(&config)?;
    fs::write(&config_path, json_data)?;

    debug!("Saved config to: {}", config_path.display());
    Ok(())
}

/// Returns true if this is the first run (config file doesn't exist)
pub fn is_first_run() -> Result<bool> {
    let config_path = get_config_file_path()?;
    Ok(!config_path.exists())
}

/// Loads the config from the JSON file in the roaming AppData directory
/// Uses default values for missing fields
pub fn load_config() -> Result<AppConfig> {
    let config_path = get_config_file_path()?;

    if !config_path.exists() {
        debug!("Config file does not exist, using defaults");
        return Ok(AppConfig {
            hotkey: default_hotkey_config(),
            devices: HashMap::new(),
        });
    }

    let json_data = fs::read_to_string(&config_path)?;
    // serde will use default values for any missing fields
    let config: AppConfig = serde_json::from_str(&json_data).unwrap_or_else(|e| {
        error!("Failed to parse config file: {e}, using defaults");
        AppConfig {
            hotkey: default_hotkey_config(),
            devices: HashMap::new(),
        }
    });

    debug!("Loaded config from: {}", config_path.display());
    Ok(config)
}

/// Applies the loaded selectable state to the current devices
pub fn apply_device_selectable_state(
    devices: &mut [AudioDevice],
    saved_states: &HashMap<String, bool>,
) {
    for device in devices.iter_mut() {
        if let Some(&selectable) = saved_states.get(&device.id) {
            device.selectable = selectable;
            debug!(
                "Applied selectable state for device {}: {}",
                device.friendly_name, selectable
            );
        }
    }
}

/// Merges runtime device list with saved device states
/// Current devices get their selectable state from saved_devices
/// Previously saved devices that are no longer plugged in are preserved in the list
pub fn merge_device_states(current_devices: &mut Vec<AudioDevice>, saved_devices: &[AudioDevice]) {
    // Build a map of saved device states
    let saved_map: HashMap<String, bool> = saved_devices
        .iter()
        .map(|d| (d.id.clone(), d.selectable))
        .collect();

    // Apply saved states to current devices
    for device in current_devices.iter_mut() {
        if let Some(&selectable) = saved_map.get(&device.id) {
            device.selectable = selectable;
            debug!(
                "Restored selectable state for device {}: {}",
                device.friendly_name, selectable
            );
        }
    }

    // Add previously saved devices that are no longer available (unplugged)
    // These devices will appear in settings but can't be selected as current device
    for saved_device in saved_devices.iter() {
        if !current_devices.iter().any(|d| d.id == saved_device.id) {
            debug!(
                "Preserving unplugged device in settings: {}",
                saved_device.friendly_name
            );
            current_devices.push(saved_device.clone());
        }
    }
}
