#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use defer::defer;
use log::{debug, error, info};
use serde::{Deserialize, Serialize};
use simple_error::bail;
use std::cell::RefCell;
use std::collections::HashMap;
use std::error::Error;
use std::fs;
use std::path::PathBuf;
use std::ptr::null_mut;
use std::rc::Rc;
use windows::Win32;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::{GetLastError, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM};
use windows::Win32::Graphics::Dwm::{DwmSetWindowAttribute, DWMWA_USE_IMMERSIVE_DARK_MODE};
use windows::Win32::Graphics::Gdi::{
    GetMonitorInfoW, MonitorFromWindow, MONITORINFO, MONITOR_DEFAULTTONEAREST,
};
use windows::Win32::Media::Audio::{
    eConsole, ERole, EndpointFormFactor, Headphones, Headset, IMMDeviceEnumerator,
    IMMNotificationClient, IMMNotificationClient_Impl, MMDeviceEnumerator,
    PKEY_AudioEndpoint_FormFactor, Speakers,
};
use windows::Win32::System::Com::StructuredStorage::PROPVARIANT;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
    STGM_READ,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Variant::{VT_LPWSTR, VT_UI4};
use windows::Win32::UI::Controls::{
    InitCommonControlsEx, HKM_GETHOTKEY, HKM_SETHOTKEY, HOTKEYF_ALT, HOTKEYF_CONTROL,
    HOTKEYF_SHIFT, ICC_HOTKEY_CLASS, ICC_LISTVIEW_CLASSES, INITCOMMONCONTROLSEX,
    LIST_VIEW_ITEM_STATE_FLAGS, LVCF_TEXT, LVCF_WIDTH, LVCOLUMNW, LVIF_PARAM, LVIF_TEXT,
    LVIS_STATEIMAGEMASK, LVITEMW, LVM_DELETEALLITEMS, LVM_GETITEMCOUNT, LVM_GETITEMSTATE,
    LVM_INSERTCOLUMNW, LVM_INSERTITEMW, LVM_SETEXTENDEDLISTVIEWSTYLE, LVM_SETITEMSTATE,
    LVS_EX_CHECKBOXES,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    RegisterHotKey, UnregisterHotKey, MOD_ALT, MOD_CONTROL, MOD_SHIFT,
};
use windows::Win32::UI::Shell::{
    FOLDERID_RoamingAppData, SHGetKnownFolderPath, ShellExecuteW, Shell_NotifyIconW,
    KNOWN_FOLDER_FLAG, NIF_GUID, NIF_ICON, NIF_MESSAGE, NIF_SHOWTIP, NIF_TIP, NIM_ADD, NIM_DELETE,
    NIM_MODIFY, NIM_SETVERSION, NIN_SELECT, NOTIFYICONDATAW, NOTIFYICONDATAW_0,
    NOTIFYICON_VERSION_4,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DialogBoxParamW,
    DispatchMessageW, EndDialog, GetCursorPos, GetDlgItem, GetMessageW, GetWindowLongPtrW,
    GetWindowRect, InsertMenuItemW, LoadIconW, PostMessageW, PostQuitMessage, RegisterClassExW,
    SendDlgItemMessageW, SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos,
    TrackPopupMenuEx, UnregisterClassW, GWLP_USERDATA, HICON, HMENU, ICON_BIG, ICON_SMALL,
    MENUITEMINFOW, MFS_DISABLED, MFT_STRING, MIIM_FTYPE, MIIM_ID, MIIM_STATE, MIIM_STRING, MSG,
    SWP_NOSIZE, SWP_NOZORDER, SW_SHOWNORMAL, TPM_BOTTOMALIGN, TPM_LEFTALIGN, TPM_RIGHTBUTTON,
    WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_DESTROY, WM_HOTKEY,
    WM_INITDIALOG, WM_QUIT, WM_RBUTTONUP, WM_SETICON, WNDCLASSEXW,
};
use windows_core::{BOOL, GUID};
use windows_strings::{w, PCWSTR};

mod policy_config;
mod safe_strings;

use policy_config::IPolicyConfig;
use safe_strings::with_wide_str;

const NOTIFY_ICON_GUID: GUID = GUID::from_u128(0x8fc84650_4bca_4125_b778_10313f9623df);
const IDD_SETTINGS: u32 = 101;
const IDC_HOTKEY: i32 = 1001;
const IDC_DEVICE_LIST: i32 = 1002;
const WM_DEVICE_CHANGE: u32 = WM_APP + 0x100;

#[windows::core::implement(IMMNotificationClient)]
pub struct CustomImmNotificationClient {
    settings_dialog: Rc<RefCell<Option<HWND>>>,
}

#[allow(non_snake_case)]
impl IMMNotificationClient_Impl for CustomImmNotificationClient_Impl {
    fn OnDeviceStateChanged(
        &self,
        pwstrdeviceid: &windows_core::PCWSTR,
        dwnewstate: windows::Win32::Media::Audio::DEVICE_STATE,
    ) -> windows_core::Result<()> {
        let device_id = unsafe { pwstrdeviceid.to_string().unwrap_or_default() };
        let friendly_name =
            get_device_friendly_name(&device_id).unwrap_or_else(|_| "Unknown".to_string());
        debug!(
            "Device state changed: name='{friendly_name}', id={device_id}, new_state={dwnewstate:?}"
        );
        if let Some(hwnd) = *self.settings_dialog.borrow() {
            unsafe {
                let _ = PostMessageW(Some(hwnd), WM_DEVICE_CHANGE, WPARAM(0), LPARAM(0));
            }
        }
        Ok(())
    }

    fn OnDeviceAdded(&self, pwstrdeviceid: &windows_core::PCWSTR) -> windows_core::Result<()> {
        let device_id = unsafe { pwstrdeviceid.to_string().unwrap_or_default() };
        let friendly_name =
            get_device_friendly_name(&device_id).unwrap_or_else(|_| "Unknown".to_string());
        debug!("Device added: name='{friendly_name}', id={device_id}");
        if let Some(hwnd) = *self.settings_dialog.borrow() {
            unsafe {
                let _ = PostMessageW(Some(hwnd), WM_DEVICE_CHANGE, WPARAM(0), LPARAM(0));
            }
        }
        Ok(())
    }

    fn OnDeviceRemoved(&self, pwstrdeviceid: &windows_core::PCWSTR) -> windows_core::Result<()> {
        let device_id = unsafe { pwstrdeviceid.to_string().unwrap_or_default() };
        let friendly_name =
            get_device_friendly_name(&device_id).unwrap_or_else(|_| "Unknown".to_string());
        debug!("Device removed: name='{friendly_name}', id={device_id}");
        if let Some(hwnd) = *self.settings_dialog.borrow() {
            unsafe {
                let _ = PostMessageW(Some(hwnd), WM_DEVICE_CHANGE, WPARAM(0), LPARAM(0));
            }
        }
        Ok(())
    }

    fn OnDefaultDeviceChanged(
        &self,
        flow: windows::Win32::Media::Audio::EDataFlow,
        role: ERole,
        pwstrdefaultdeviceid: &windows_core::PCWSTR,
    ) -> windows_core::Result<()> {
        let device_id = unsafe { pwstrdefaultdeviceid.to_string().unwrap_or_default() };
        let friendly_name =
            get_device_friendly_name(&device_id).unwrap_or_else(|_| "Unknown".to_string());
        debug!(
            "Default device changed: name='{friendly_name}', flow={flow:?}, role={role:?}, id={device_id}"
        );
        Ok(())
    }

    fn OnPropertyValueChanged(
        &self,
        pwstrdeviceid: &windows_core::PCWSTR,
        key: &Win32::Foundation::PROPERTYKEY,
    ) -> windows_core::Result<()> {
        let device_id = unsafe { pwstrdeviceid.to_string().unwrap_or_default() };
        let friendly_name =
            get_device_friendly_name(&device_id).unwrap_or_else(|_| "Unknown".to_string());
        debug!("Property value changed: name='{friendly_name}', id={device_id}, key={key:?}");
        Ok(())
    }
}

/// Sets the default audio endpoint for the specified role using raw COM interface calls
fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        debug!("Attempting to set default endpoint for device: {device_id}, role: {role:?}",);
        let policy_config: IPolicyConfig =
            CoCreateInstance(&policy_config::CLSID_POLICY_CONFIG, None, CLSCTX_ALL)?;

        // Use safe scoped approach for string conversion
        with_wide_str(device_id, |wide_device_id| {
            policy_config.SetDefaultEndpoint(wide_device_id, role)
        })?;
        Ok(())
    }
}

/// Gets the friendly name for a device given its ID
fn get_device_friendly_name(device_id: &str) -> Result<String, Box<dyn Error>> {
    unsafe {
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

        let device = with_wide_str(device_id, |wide_id| device_enumerator.GetDevice(wide_id))?;

        let props = device.OpenPropertyStore(STGM_READ)?;
        let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
        propvariant_to_string(&friendly_name)
    }
}

/// Gets the current default audio endpoint for debugging
fn get_current_default_endpoint(role: ERole) -> Result<String, Box<dyn Error>> {
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

        let endpoint = device_enumerator
            .GetDefaultAudioEndpoint(windows::Win32::Media::Audio::eRender, role)?;

        let device_id = endpoint.GetId()?;
        let device_id_str = device_id.to_string()?;

        Ok(device_id_str)
    }
}

fn string_to_tip(s: &str) -> [u16; 128] {
    let mut ret = [0u16; 128];
    let encoded: Vec<u16> = s.encode_utf16().collect();
    assert!(encoded.len() < ret.len());
    for (i, &c) in encoded.iter().enumerate() {
        ret[i] = c;
    }
    ret[encoded.len()] = 0; // Null-terminate the string
    ret
}

#[derive(Debug)]
struct AdaptiveIcon {
    light: HICON,
    dark: HICON,
}

impl AdaptiveIcon {
    pub fn new(light_icon_name: &str, dark_icon_name: &str) -> Result<Self, Box<dyn Error>> {
        let light_icon = unsafe { load_icon(light_icon_name)? };
        let dark_icon = unsafe { load_icon(dark_icon_name)? };
        Ok(Self {
            light: light_icon,
            dark: dark_icon,
        })
    }

    pub fn icon(&self) -> Result<HICON, Box<dyn Error>> {
        if is_dark_mode()? {
            Ok(self.dark)
        } else {
            Ok(self.light)
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct AudioDevice {
    id: String,
    friendly_name: String,
    // Whether this device will be included in the rotation.
    selectable: bool,
    #[serde(skip)]
    form_factor: EndpointFormFactor,
}

#[derive(Debug, Serialize, Deserialize)]
struct HotkeyConfig {
    #[serde(default = "default_hotkey_vk")]
    vk: u8,
    #[serde(default = "default_hotkey_mods")]
    mods: u8,
}

#[derive(Debug, Serialize, Deserialize)]
struct AppConfig {
    #[serde(default = "default_hotkey_config")]
    hotkey: HotkeyConfig,
    #[serde(default)]
    devices: HashMap<String, bool>,
}

fn default_hotkey_vk() -> u8 {
    0x58 // 'X' key
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

#[derive(Debug)]
struct AudioSwitch {
    window: HWND,
    icon: AdaptiveIcon,
    popup_menu: HMENU,
    available_devices: Vec<AudioDevice>,

    headphones_icon: AdaptiveIcon,
    headset_icon: AdaptiveIcon,
    speaker_icon: AdaptiveIcon,

    hotkey_vk: u8,
    hotkey_mods: u8,

    notification_client: IMMNotificationClient,
    settings_dialog_hwnd: Rc<RefCell<Option<HWND>>>,
}

impl Drop for AudioSwitch {
    fn drop(&mut self) {
        unsafe {
            let _ = DestroyMenu(self.popup_menu);
        }
    }
}

impl AudioSwitch {
    #![allow(non_upper_case_globals)]
    fn icon_for_form_factor(
        &self,
        form_factor: EndpointFormFactor,
    ) -> Result<HICON, Box<dyn Error>> {
        let adaptive_icon = match form_factor {
            Headphones => &self.headphones_icon,
            Headset => &self.headset_icon,
            Speakers => &self.speaker_icon,
            _ => &self.icon, // Default icon for other form factors
        };
        adaptive_icon.icon()
    }

    fn current_icon(&self) -> Result<HICON, Box<dyn Error>> {
        let current_device_id = get_current_default_endpoint(eConsole)?;
        let current_device = self
            .available_devices
            .iter()
            .find(|d| d.id == current_device_id)
            .ok_or_else(|| simple_error::SimpleError::new("Current device not found"))?;
        self.icon_for_form_factor(current_device.form_factor)
    }

    fn show_popup_menu(&self, x: i32, y: i32) -> Result<(), Box<dyn Error>> {
        debug!("Showing popup menu at ({x}, {y})");
        unsafe {
            // Required to ensure the popup menu disappears again when a user clicks elsewhere.
            SetForegroundWindow(self.window).ok()?;
            TrackPopupMenuEx(
                self.popup_menu,
                TPM_LEFTALIGN.0 | TPM_BOTTOMALIGN.0 | TPM_RIGHTBUTTON.0,
                x,
                y,
                self.window,
                None,
            )
            .ok()?;
        }
        Ok(())
    }

    fn menu_selection(&mut self, id: u32) -> Result<(), Box<dyn Error>> {
        debug!("Menu item selected: {id}");
        unsafe {
            match id {
                // Exit item selected.
                POPUP_EXIT_ID => {
                    PostMessageW(
                        Some(self.window),
                        WM_CLOSE,
                        WPARAM::default(),
                        LPARAM::default(),
                    )?;
                }
                POPUP_ABOUT_ID => {
                    ShellExecuteW(
                        None,
                        PCWSTR(null_mut()),
                        w!("https://github.com/hatstand/sound-switcheroo"),
                        PCWSTR(null_mut()),
                        PCWSTR(null_mut()),
                        SW_SHOWNORMAL,
                    );
                }
                POPUP_SETTINGS_ID => {
                    // Unregister hotkey before opening settings dialog
                    // so the hotkey control can use the same key combination
                    let _ = UnregisterHotKey(None, HOTKEY_ID);

                    let dialog_result = show_settings_dialog(
                        self.window,
                        &mut self.available_devices,
                        &mut self.hotkey_vk,
                        &mut self.hotkey_mods,
                        &self.notification_client,
                        &self.settings_dialog_hwnd,
                    );

                    // Always re-register the hotkey after closing the dialog
                    // Hotkey control uses: HOTKEYF_SHIFT, HOTKEYF_CONTROL, HOTKEYF_ALT
                    let mut mods =
                        windows::Win32::UI::Input::KeyboardAndMouse::HOT_KEY_MODIFIERS(0);
                    if self.hotkey_mods & HOTKEYF_SHIFT as u8 != 0 {
                        mods |= MOD_SHIFT;
                    }
                    if self.hotkey_mods & HOTKEYF_CONTROL as u8 != 0 {
                        mods |= MOD_CONTROL;
                    }
                    if self.hotkey_mods & HOTKEYF_ALT as u8 != 0 {
                        mods |= MOD_ALT;
                    }

                    if let Err(e) = RegisterHotKey(None, HOTKEY_ID, mods, self.hotkey_vk as u32) {
                        error!("Failed to register hotkey: {e}");
                    }

                    // Save settings if dialog was accepted
                    if let Ok(true) = dialog_result {
                        if let Err(e) =
                            save_config(&self.available_devices, self.hotkey_vk, self.hotkey_mods)
                        {
                            error!("Failed to save config: {e}");
                        }
                    }
                }
                _ => {
                    debug!("Unknown menu item selected: {id}");
                }
            }
        }
        Ok(())
    }

    fn next_device(&mut self) -> Result<(), Box<dyn Error>> {
        let current_device = get_current_default_endpoint(eConsole)?;
        debug!("Switching to next device from: {current_device}");
        let current_index = self
            .available_devices
            .iter()
            .position(|d| d.id == current_device)
            .unwrap_or(0);
        debug!("Current device index: {current_index}");
        let selectable_devices: Vec<_> = self
            .available_devices
            .iter()
            .enumerate()
            .filter(|(_, d)| d.selectable)
            .collect();
        if selectable_devices.is_empty() {
            debug!("No selectable devices found");
            return Ok(());
        }

        let (_, cand_device) = selectable_devices
            .iter()
            // Either the first selectable device after the current one,
            .find(|(i, _)| *i > current_index)
            // or the first selectable device if none found as a wraparound.
            .or_else(|| selectable_devices.first())
            .ok_or_else(|| simple_error::SimpleError::new("No selectable devices found"))?;
        info!("Switching to device: {:}", cand_device.friendly_name,);
        set_default_endpoint(&cand_device.id, eConsole)?;
        // Update the tooltip to reflect the new current device.
        let tooltip = cand_device.friendly_name.clone();
        unsafe {
            Shell_NotifyIconW(
                NIM_MODIFY,
                &NOTIFYICONDATAW {
                    cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
                    hWnd: self.window,
                    hIcon: self.icon_for_form_factor(cand_device.form_factor)?,
                    guidItem: NOTIFY_ICON_GUID,
                    // Both NIF_TIP & NIF_SHOWTIP are required to actually show the tooltip.
                    uFlags: NIF_ICON | NIF_MESSAGE | NIF_GUID | NIF_TIP | NIF_SHOWTIP,
                    uCallbackMessage: WM_APP + 0x42,
                    szTip: string_to_tip(&tooltip),
                    Anonymous: NOTIFYICONDATAW_0 {
                        uVersion: NOTIFYICON_VERSION_4,
                    },
                    ..Default::default()
                },
            )
            .ok()?;
        }

        Ok(())
    }
}

const POPUP_EXIT_ID: u32 = 1;
const POPUP_SETTINGS_ID: u32 = 2;
const POPUP_ABOUT_ID: u32 = 3;

unsafe fn create_popup_menu() -> Result<HMENU, Box<dyn Error>> {
    unsafe {
        let menu = CreatePopupMenu()?;
        // Add a menu item to exit the application.
        safe_strings::with_wide_str_mut("Exit", |exit_name| -> Result<(), Box<dyn Error>> {
            InsertMenuItemW(
                menu,
                0,
                true,
                &MENUITEMINFOW {
                    cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                    fMask: MIIM_FTYPE | MIIM_ID | MIIM_STRING,
                    fType: MFT_STRING,
                    dwTypeData: exit_name,
                    cch: exit_name.len() as u32 - 1,
                    wID: POPUP_EXIT_ID,
                    ..Default::default()
                },
            )?;
            Ok(())
        })?;
        // Add Settings menu item
        safe_strings::with_wide_str_mut(
            "Settings...",
            |settings_name| -> Result<(), Box<dyn Error>> {
                InsertMenuItemW(
                    menu,
                    0,
                    true,
                    &MENUITEMINFOW {
                        cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                        fMask: MIIM_FTYPE | MIIM_ID | MIIM_STRING,
                        fType: MFT_STRING,
                        dwTypeData: settings_name,
                        cch: settings_name.len() as u32 - 1,
                        wID: POPUP_SETTINGS_ID,
                        ..Default::default()
                    },
                )?;
                Ok(())
            },
        )?;
        // Add a menu item for the about dialog.
        safe_strings::with_wide_str_mut("About", |about_name| -> Result<(), Box<dyn Error>> {
            InsertMenuItemW(
                menu,
                0,
                true,
                &MENUITEMINFOW {
                    cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                    fMask: MIIM_FTYPE | MIIM_ID | MIIM_STRING,
                    fType: MFT_STRING,
                    dwTypeData: about_name,
                    cch: about_name.len() as u32 - 1,
                    wID: POPUP_ABOUT_ID,
                    ..Default::default()
                },
            )?;
            Ok(())
        })?;
        // Add a nice name to the top of the menu.
        safe_strings::with_wide_str_mut(
            "Sound Switcheroo",
            |title_name| -> Result<(), Box<dyn Error>> {
                InsertMenuItemW(
                    menu,
                    0,
                    true,
                    &MENUITEMINFOW {
                        cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                        fMask: MIIM_FTYPE | MIIM_STATE | MIIM_STRING,
                        fType: MFT_STRING,
                        dwTypeData: title_name,
                        cch: title_name.len() as u32 - 1,
                        fState: MFS_DISABLED,
                        ..Default::default()
                    },
                )?;
                Ok(())
            },
        )?;
        Ok(menu)
    }
}

unsafe fn propvariant_to_string(propvar: &PROPVARIANT) -> Result<String, Box<dyn Error>> {
    unsafe {
        match propvar.vt() {
            VT_LPWSTR => Ok(String::from_utf16_lossy(
                propvar.Anonymous.Anonymous.Anonymous.pwszVal.as_wide(),
            )),
            _ => {
                bail!("Unsupported PROPVARIANT type: {:?}", propvar.vt());
            }
        }
    }
}

fn get_available_audio_devices() -> Result<Vec<AudioDevice>, Box<dyn Error>> {
    let mut devices = Vec::new();
    unsafe {
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let endpoints = device_enumerator.EnumAudioEndpoints(
            windows::Win32::Media::Audio::eRender,
            windows::Win32::Media::Audio::DEVICE_STATE_ACTIVE,
        )?;

        for i in 0..endpoints.GetCount()? {
            let endpoint = endpoints.Item(i)?;
            let device_id = endpoint.GetId()?;
            let device_id_str = device_id.to_string()?;
            let props = endpoint.OpenPropertyStore(STGM_READ)?;
            let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
            let form_factor_var = props.GetValue(&PKEY_AudioEndpoint_FormFactor)?;
            let form_factor: EndpointFormFactor = match form_factor_var.vt() {
                VT_UI4 => {
                    EndpointFormFactor(form_factor_var.Anonymous.Anonymous.Anonymous.ulVal as i32)
                }
                _ => {
                    bail!(
                        "Unsupported PROPVARIANT type for form factor: {:?}",
                        form_factor_var,
                    );
                }
            };
            devices.push(AudioDevice {
                id: device_id_str,
                friendly_name: propvariant_to_string(&friendly_name)?,
                selectable: true,
                form_factor,
            });
        }
    }
    Ok(devices)
}

/// Gets the path to the user's roaming AppData directory
fn get_roaming_appdata_path() -> Result<PathBuf, Box<dyn Error>> {
    unsafe {
        let path_ptr =
            SHGetKnownFolderPath(&FOLDERID_RoamingAppData, KNOWN_FOLDER_FLAG::default(), None)?;

        let path_str = path_ptr.to_string()?;
        let path = PathBuf::from(path_str);

        // Free the memory allocated by SHGetKnownFolderPath
        windows::Win32::System::Com::CoTaskMemFree(Some(path_ptr.as_ptr() as *const _));

        Ok(path)
    }
}

/// Gets the full path to the AudioSwitch configuration file
fn get_config_file_path() -> Result<PathBuf, Box<dyn Error>> {
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
fn save_config(
    devices: &[AudioDevice],
    hotkey_vk: u8,
    hotkey_mods: u8,
) -> Result<(), Box<dyn Error>> {
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

/// Loads the config from the JSON file in the roaming AppData directory
/// Uses default values for missing fields
fn load_config() -> Result<AppConfig, Box<dyn Error>> {
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
fn apply_device_selectable_state(
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
fn merge_device_states(current_devices: &mut Vec<AudioDevice>, saved_devices: &[AudioDevice]) {
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

fn is_dark_mode() -> Result<bool, Box<dyn Error>> {
    let theme_key = windows_registry::CURRENT_USER
        .open(r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")?;
    let light_theme = theme_key.get_u32("AppsUseLightTheme")? == 1;
    Ok(!light_theme)
}

unsafe fn load_icon(icon_name: &str) -> Result<HICON, Box<dyn Error>> {
    unsafe {
        let module = GetModuleHandleW(None)?;
        let icon = with_wide_str(icon_name, |wide_icon_name| {
            LoadIconW(Some(module.into()), wide_icon_name)
        })?;
        if icon.is_invalid() {
            bail!("Failed to load icon: {}", icon_name);
        }
        Ok(icon)
    }
}

struct SettingsDialog {
    devices: Vec<AudioDevice>,
    hotkey_vk: u8,
    hotkey_mods: u8,
    dialog_hwnd_arc: Rc<RefCell<Option<HWND>>>,
}

unsafe fn refresh_device_list(hwnd: HWND, settings: &mut SettingsDialog) {
    unsafe {
        let list_hwnd = GetDlgItem(Some(hwnd), IDC_DEVICE_LIST).unwrap();

        // Save current checkbox states by device ID before refresh
        let count = SendMessageW(
            list_hwnd,
            LVM_GETITEMCOUNT,
            Some(WPARAM(0)),
            Some(LPARAM(0)),
        )
        .0;

        let mut device_states: HashMap<String, bool> = HashMap::new();
        for i in 0..count {
            if (i as usize) < settings.devices.len() {
                let state = SendMessageW(
                    list_hwnd,
                    LVM_GETITEMSTATE,
                    Some(WPARAM(i as usize)),
                    Some(LPARAM(LVIS_STATEIMAGEMASK.0 as isize)),
                )
                .0;
                let device = &settings.devices[i as usize];
                device_states.insert(device.id.clone(), (state & 0x2000) != 0);
            }
        }

        // Fetch current devices and apply saved checkbox states
        // During refresh, only show currently available devices (don't preserve unplugged ones)
        if let Ok(mut current_devices) = get_available_audio_devices() {
            // Apply saved checkbox states to current devices
            for device in current_devices.iter_mut() {
                if let Some(&selectable) = device_states.get(&device.id) {
                    device.selectable = selectable;
                }
            }
            settings.devices = current_devices;
        }

        // Clear the list
        SendMessageW(
            list_hwnd,
            LVM_DELETEALLITEMS,
            Some(WPARAM(0)),
            Some(LPARAM(0)),
        );

        // Re-populate with updated devices
        for (idx, device) in settings.devices.iter().enumerate() {
            safe_strings::with_wide_str_mut(&device.friendly_name, |device_name| {
                let lvi = LVITEMW {
                    mask: LVIF_TEXT | LVIF_PARAM,
                    iItem: idx as i32,
                    pszText: windows_core::PWSTR(device_name.0),
                    lParam: LPARAM(idx as isize),
                    ..Default::default()
                };
                SendMessageW(
                    list_hwnd,
                    LVM_INSERTITEMW,
                    Some(WPARAM(0)),
                    Some(LPARAM(&lvi as *const _ as isize)),
                );

                // Set checkbox state
                let state = if device.selectable { 0x2000 } else { 0x1000 };
                let state_lvi = LVITEMW {
                    stateMask: LVIS_STATEIMAGEMASK,
                    state: LIST_VIEW_ITEM_STATE_FLAGS(state),
                    ..Default::default()
                };
                SendMessageW(
                    list_hwnd,
                    LVM_SETITEMSTATE,
                    Some(WPARAM(idx)),
                    Some(LPARAM(&state_lvi as *const _ as isize)),
                );
            });
        }
    }
}

unsafe extern "system" fn settings_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> isize {
    unsafe {
        match msg {
            WM_INITDIALOG => {
                let settings = lparam.0 as *mut SettingsDialog;
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, settings as isize);

                let settings_ref = &mut *settings;

                // Store dialog handle in Rc for notifications
                *settings_ref.dialog_hwnd_arc.borrow_mut() = Some(hwnd);

                // Enable dark mode and set icon based on system theme
                let dark_mode = is_dark_mode().unwrap_or(false);
                if dark_mode {
                    let dark_mode_enabled: u32 = 1;
                    let _ = DwmSetWindowAttribute(
                        hwnd,
                        DWMWA_USE_IMMERSIVE_DARK_MODE,
                        &dark_mode_enabled as *const _ as *const _,
                        std::mem::size_of::<u32>() as u32,
                    );
                }

                // Set dialog icon
                if let Ok(adaptive_icon) =
                    AdaptiveIcon::new("switcheroo_icon", "switcheroo_dark_icon")
                {
                    let icon = if dark_mode {
                        adaptive_icon.dark
                    } else {
                        adaptive_icon.light
                    };
                    SendMessageW(
                        hwnd,
                        WM_SETICON,
                        Some(WPARAM(ICON_BIG as usize)),
                        Some(LPARAM(icon.0 as isize)),
                    );
                    SendMessageW(
                        hwnd,
                        WM_SETICON,
                        Some(WPARAM(ICON_SMALL as usize)),
                        Some(LPARAM(icon.0 as isize)),
                    );
                }

                // Initialize hotkey control
                let hotkey_value =
                    ((settings_ref.hotkey_mods as u16) << 8) | (settings_ref.hotkey_vk as u16);
                SendDlgItemMessageW(
                    hwnd,
                    IDC_HOTKEY,
                    HKM_SETHOTKEY,
                    WPARAM(hotkey_value as usize),
                    LPARAM(0),
                );

                // Initialize device list
                let list_hwnd = GetDlgItem(Some(hwnd), IDC_DEVICE_LIST).unwrap();

                // Enable checkboxes
                SendMessageW(
                    list_hwnd,
                    LVM_SETEXTENDEDLISTVIEWSTYLE,
                    Some(WPARAM(LVS_EX_CHECKBOXES as usize)),
                    Some(LPARAM(LVS_EX_CHECKBOXES as isize)),
                );

                // Add single column that fills the list view width
                safe_strings::with_wide_str_mut("Device Name", |col_text| {
                    let lvc = LVCOLUMNW {
                        mask: LVCF_TEXT | LVCF_WIDTH,
                        cx: 260, // Match the list view width from resources
                        pszText: windows_core::PWSTR(col_text.0),
                        ..Default::default()
                    };
                    SendMessageW(
                        list_hwnd,
                        LVM_INSERTCOLUMNW,
                        Some(WPARAM(0)),
                        Some(LPARAM(&lvc as *const _ as isize)),
                    );
                });

                // Add devices to list
                for (idx, device) in settings_ref.devices.iter().enumerate() {
                    safe_strings::with_wide_str_mut(&device.friendly_name, |device_name| {
                        let lvi = LVITEMW {
                            mask: LVIF_TEXT | LVIF_PARAM,
                            iItem: idx as i32,
                            pszText: windows_core::PWSTR(device_name.0),
                            lParam: LPARAM(idx as isize),
                            ..Default::default()
                        };
                        SendMessageW(
                            list_hwnd,
                            LVM_INSERTITEMW,
                            Some(WPARAM(0)),
                            Some(LPARAM(&lvi as *const _ as isize)),
                        );

                        // Set checkbox state
                        let state = if device.selectable { 0x2000 } else { 0x1000 }; // Checked/unchecked
                        let state_lvi = LVITEMW {
                            stateMask: LVIS_STATEIMAGEMASK,
                            state: LIST_VIEW_ITEM_STATE_FLAGS(state),
                            ..Default::default()
                        };
                        SendMessageW(
                            list_hwnd,
                            LVM_SETITEMSTATE,
                            Some(WPARAM(idx)),
                            Some(LPARAM(&state_lvi as *const _ as isize)),
                        );
                    });
                }

                // Center the dialog on the current monitor
                let mut dialog_rect = RECT::default();
                if GetWindowRect(hwnd, &mut dialog_rect).is_ok() {
                    let monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
                    let mut monitor_info = MONITORINFO {
                        cbSize: std::mem::size_of::<MONITORINFO>() as u32,
                        ..Default::default()
                    };
                    if GetMonitorInfoW(monitor, &mut monitor_info).as_bool() {
                        let work_area = monitor_info.rcWork;
                        let dialog_width = dialog_rect.right - dialog_rect.left;
                        let dialog_height = dialog_rect.bottom - dialog_rect.top;
                        let center_x =
                            work_area.left + (work_area.right - work_area.left - dialog_width) / 2;
                        let center_y =
                            work_area.top + (work_area.bottom - work_area.top - dialog_height) / 2;
                        let _ = SetWindowPos(
                            hwnd,
                            None,
                            center_x,
                            center_y,
                            0,
                            0,
                            SWP_NOSIZE | SWP_NOZORDER,
                        );
                    }
                }

                // Store dialog handle for notifications - we'll retrieve it from SettingsDialog
                1
            }
            WM_DESTROY => {
                // Clear dialog handle when closing
                let settings = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SettingsDialog;
                if !settings.is_null() {
                    // Dialog is closing, no need to track it anymore
                }
                0
            }
            WM_DEVICE_CHANGE => {
                debug!("Device change detected, refreshing device list");
                let settings = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SettingsDialog;
                if !settings.is_null() {
                    let settings_ref = &mut *settings;
                    refresh_device_list(hwnd, settings_ref);
                }
                1
            }
            WM_COMMAND => {
                let cmd = (wparam.0 & 0xFFFF) as u32;
                match cmd {
                    1 => {
                        // IDOK
                        let settings =
                            GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SettingsDialog;
                        let settings_ref = &mut *settings;

                        // Get hotkey value
                        let hotkey = SendDlgItemMessageW(
                            hwnd,
                            IDC_HOTKEY,
                            HKM_GETHOTKEY,
                            WPARAM(0),
                            LPARAM(0),
                        );
                        settings_ref.hotkey_vk = (hotkey.0 & 0xFF) as u8;
                        settings_ref.hotkey_mods = ((hotkey.0 >> 8) & 0xFF) as u8;

                        // Get device selections
                        let list_hwnd = GetDlgItem(Some(hwnd), IDC_DEVICE_LIST).unwrap();
                        let count = SendMessageW(
                            list_hwnd,
                            LVM_GETITEMCOUNT,
                            Some(WPARAM(0)),
                            Some(LPARAM(0)),
                        )
                        .0;

                        for i in 0..count {
                            let state = SendMessageW(
                                list_hwnd,
                                LVM_GETITEMSTATE,
                                Some(WPARAM(i as usize)),
                                Some(LPARAM(LVIS_STATEIMAGEMASK.0 as isize)),
                            )
                            .0;
                            settings_ref.devices[i as usize].selectable = (state & 0x2000) != 0;
                        }

                        EndDialog(hwnd, 1).expect("EndDialog error");
                        1
                    }
                    2 => {
                        // IDCANCEL
                        EndDialog(hwnd, 0).expect("EndDialog failed IDCANCEL");
                        1
                    }
                    _ => 0,
                }
            }
            _ => 0,
        }
    }
}

fn show_settings_dialog(
    parent: HWND,
    devices: &mut Vec<AudioDevice>,
    hotkey_vk: &mut u8,
    hotkey_mods: &mut u8,
    _notification_client: &IMMNotificationClient,
    dialog_hwnd_arc: &Rc<RefCell<Option<HWND>>>,
) -> Result<bool, Box<dyn Error>> {
    unsafe {
        // Initialize common controls
        let icc = INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_LISTVIEW_CLASSES | ICC_HOTKEY_CLASS,
        };
        InitCommonControlsEx(&icc).ok()?;

        // Fetch currently available devices from the system
        let mut current_devices = get_available_audio_devices()?;

        // Merge current devices with saved state from devices parameter
        merge_device_states(&mut current_devices, devices);

        let mut settings = SettingsDialog {
            devices: current_devices,
            hotkey_vk: *hotkey_vk,
            hotkey_mods: *hotkey_mods,
            dialog_hwnd_arc: dialog_hwnd_arc.clone(),
        };

        let module = GetModuleHandleW(None)?;

        // Dialog proc will store handle in Arc
        let result = DialogBoxParamW(
            Some(module.into()),
            PCWSTR(IDD_SETTINGS as *const u16),
            Some(parent),
            Some(settings_dialog_proc),
            LPARAM(&mut settings as *mut _ as isize),
        );

        // Clear dialog handle after dialog closes
        *dialog_hwnd_arc.borrow_mut() = None;

        if result == 1 {
            *devices = settings.devices;
            *hotkey_vk = settings.hotkey_vk;
            *hotkey_mods = settings.hotkey_mods;
            Ok(true)
        } else {
            Ok(false)
        }
    }
}

const HOTKEY_ID: i32 = 1225708739;

fn main() -> Result<(), Box<dyn Error>> {
    env_logger::init();
    info!("Audio Switch Tool");
    unsafe {
        debug!("Dark mode: {}", is_dark_mode()?);
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        defer!({
            CoUninitialize();
        });
        let module = GetModuleHandleW(None)?;
        // Register a window class for the taskbar icon.
        let class = RegisterClassExW(&WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(window_callback),
            hInstance: module.into(),
            lpszClassName: w!("SoundSwitcheroo"),
            ..Default::default()
        });
        defer!({
            // Unregister the class when done using its atom.
            let _ = UnregisterClassW(PCWSTR(class as *const u16), Some(module.into()));
        });

        // Seems this needs to _not_ be a message-only window for ShellExecute to work.
        let window = CreateWindowExW(
            WINDOW_EX_STYLE(0),
            // Not really a string but the class atom.
            PCWSTR(class as *const u16),
            w!("Audio Switch Tool"),
            WINDOW_STYLE(0),
            0,
            0,
            0,
            0,
            None,
            None,
            Some(module.into()),
            None,
        )
        .inspect_err(|err| {
            error!("Failed to create window: {:?} {:?}", err, GetLastError());
        })?;
        let mut devices = get_available_audio_devices()?;
        // Load config
        let config = load_config()?;
        apply_device_selectable_state(&mut devices, &config.devices);
        let current_device_id = get_current_default_endpoint(eConsole)?;
        let current_device = devices
            .iter()
            .find(|d| d.id == current_device_id)
            .ok_or_else(|| simple_error::SimpleError::new("Current device not found"))?;
        let tooltip = current_device.friendly_name.clone();

        debug!("Registering for device notifications");
        let settings_dialog_hwnd_arc = Rc::new(RefCell::new(None));
        let notification_client: IMMNotificationClient = CustomImmNotificationClient {
            settings_dialog: settings_dialog_hwnd_arc.clone(),
        }
        .into();
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        IMMDeviceEnumerator::RegisterEndpointNotificationCallback(
            &device_enumerator,
            &notification_client,
        )?;

        let me = AudioSwitch {
            window,
            icon: AdaptiveIcon::new("audio_icon", "audio_icon")?,
            popup_menu: create_popup_menu()?,
            available_devices: devices,
            headphones_icon: AdaptiveIcon::new("headphones_icon", "headphones_icon_dark")?,
            headset_icon: AdaptiveIcon::new("headset_icon", "headset_icon_dark")?,
            speaker_icon: AdaptiveIcon::new("speaker_icon", "speaker_icon_dark")?,
            hotkey_vk: config.hotkey.vk,
            hotkey_mods: config.hotkey.mods,
            notification_client: notification_client.clone(),
            settings_dialog_hwnd: settings_dialog_hwnd_arc.clone(),
        };
        // Store the AudioSwitch instance in the window's user data.
        SetWindowLongPtrW(window, GWLP_USERDATA, &me as *const _ as isize);
        let notify_icon_data = &mut NOTIFYICONDATAW {
            cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
            hWnd: window,
            hIcon: me.current_icon()?,
            guidItem: NOTIFY_ICON_GUID,
            // Both NIF_TIP & NIF_SHOWTIP are required to actually show the tooltip.
            uFlags: NIF_ICON | NIF_MESSAGE | NIF_GUID | NIF_TIP | NIF_SHOWTIP,
            uCallbackMessage: WM_APP + 0x42,
            szTip: string_to_tip(&tooltip),
            Anonymous: NOTIFYICONDATAW_0 {
                uVersion: NOTIFYICON_VERSION_4,
            },
            ..Default::default()
        };
        Shell_NotifyIconW(NIM_ADD, notify_icon_data).ok()?;
        defer!({
            // Remove the icon when done.
            debug!("Removing taskbar icon");
            let _ = Shell_NotifyIconW(
                NIM_DELETE,
                &NOTIFYICONDATAW {
                    cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
                    uFlags: NIF_GUID,
                    hWnd: window,
                    guidItem: NOTIFY_ICON_GUID,
                    ..Default::default()
                },
            );
        });
        // Enable better callback API.
        Shell_NotifyIconW(NIM_SETVERSION, notify_icon_data).ok()?;

        debug!("Registering global hotkey");
        RegisterHotKey(None, HOTKEY_ID, MOD_CONTROL | MOD_ALT | MOD_SHIFT, 0x58)?;

        // Enter the message loop.
        info!("Running...");
        loop {
            let mut msg = MSG::default();
            match GetMessageW(&mut msg, None, 0, 0) {
                // Returns false on WM_QUIT.
                BOOL(0) => {
                    assert_eq!(msg.message, WM_QUIT);
                    info!("Quitting...");
                    break;
                }
                // Actual error from GetMessageW.
                BOOL(-1) => {
                    error!("Failed to get message: {:?}", GetLastError());
                }
                // Normal window message.
                BOOL(_) => {
                    if msg.message == WM_HOTKEY && msg.wParam.0 as i32 == HOTKEY_ID {
                        debug!("Hotkey pressed, switching to next device");
                        let raw_me = GetWindowLongPtrW(window, GWLP_USERDATA) as *mut AudioSwitch;
                        let me = raw_me.as_mut().unwrap();
                        if let Err(e) = me.next_device() {
                            error!("Failed to switch to next device: {e}");
                        }
                    }
                    DispatchMessageW(&msg);
                }
            }
        }
    };

    Ok(())
}

const TASKBAR_CB_ID: u32 = WM_APP + 0x42;
#[allow(non_snake_case)]
pub fn LOWORD(l: isize) -> isize {
    l & 0xffff
}

#[allow(non_snake_case)]
pub fn HIWORD(l: isize) -> isize {
    (l >> 16) & 0xffff
}

unsafe extern "system" fn window_callback(
    hwnd: windows::Win32::Foundation::HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe {
        let raw_me = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AudioSwitch;
        match msg {
            TASKBAR_CB_ID => match LOWORD(lparam.0) as u32 {
                // Right click on the taskbar icon.
                WM_RBUTTONUP => {
                    let mut cursor_pos = POINT::default();
                    GetCursorPos(&mut cursor_pos).unwrap();
                    match raw_me
                        .as_mut()
                        .unwrap()
                        .show_popup_menu(cursor_pos.x, cursor_pos.y)
                    {
                        Ok(()) => {}
                        Err(e) => error!("Failed to show popup menu: {e:?}"),
                    }
                    LRESULT(0)
                }
                // Left click on the taskbar icon.
                NIN_SELECT => {
                    match raw_me.as_mut().unwrap().next_device() {
                        Ok(()) => {}
                        Err(e) => error!("Failed to show popup menu: {e:?}"),
                    }
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            },
            // Item in popup menu selected.
            WM_COMMAND => {
                let chosen = LOWORD(wparam.0 as isize) as u32;
                let _ = raw_me.as_mut().unwrap().menu_selection(chosen);
                LRESULT(0)
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}
