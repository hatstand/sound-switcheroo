#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use defer::defer;
use log::{debug, error, info};
use std::cell::RefCell;
use std::error::Error;
use std::ptr::null_mut;
use std::rc::Rc;
use windows::Win32::Foundation::{GetLastError, HWND, LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::Media::Audio::{
    EndpointFormFactor, Headphones, Headset, IMMDeviceEnumerator, IMMNotificationClient,
    MMDeviceEnumerator, Speakers, eConsole,
};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{HOTKEYF_ALT, HOTKEYF_CONTROL, HOTKEYF_SHIFT};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    MOD_ALT, MOD_CONTROL, MOD_SHIFT, RegisterHotKey, UnregisterHotKey,
};
use windows::Win32::UI::Shell::{
    NIF_GUID, NIF_ICON, NIF_MESSAGE, NIF_SHOWTIP, NIF_TIP, NIM_ADD, NIM_DELETE, NIM_MODIFY,
    NIM_SETVERSION, NIN_SELECT, NOTIFYICON_VERSION_4, NOTIFYICONDATAW, NOTIFYICONDATAW_0,
    Shell_NotifyIconW, ShellExecuteW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DispatchMessageW, GWLP_USERDATA,
    GetCursorPos, GetMessageW, GetWindowLongPtrW, HICON, HMENU, InsertMenuItemW, MENUITEMINFOW,
    MFS_DISABLED, MFT_STRING, MIIM_FTYPE, MIIM_ID, MIIM_STATE, MIIM_STRING, MSG, PostMessageW,
    PostQuitMessage, RegisterClassExW, SW_SHOWNORMAL, SetForegroundWindow, SetWindowLongPtrW,
    TPM_BOTTOMALIGN, TPM_LEFTALIGN, TPM_RIGHTBUTTON, TrackPopupMenuEx, UnregisterClassW,
    WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_DESTROY, WM_HOTKEY, WM_QUIT,
    WM_RBUTTONUP, WNDCLASSEXW,
};
use windows_core::{BOOL, GUID};
use windows_strings::{PCWSTR, w};

mod audio_device;
mod config;
mod icon;
mod notification_client;
mod policy_config;
mod safe_strings;
mod settings_dialog;

use audio_device::{
    get_available_audio_devices, get_current_default_endpoint, get_device_friendly_name,
    set_default_endpoint,
};
use config::{AudioDevice, apply_device_selectable_state, load_config, save_config};
use icon::{AdaptiveIcon, is_dark_mode};
use notification_client::CustomImmNotificationClient;
use settings_dialog::show_settings_dialog;

const NOTIFY_ICON_GUID: GUID = GUID::from_u128(0x8fc84650_4bca_4125_b778_10313f9623df);
const IDD_SETTINGS: u32 = 101;
const IDC_HOTKEY: i32 = 1001;
const IDC_DEVICE_LIST: i32 = 1002;
pub(crate) const WM_DEVICE_CHANGE: u32 = WM_APP + 0x100;

// ListView checkbox state constants (state image mask values)
// These represent INDEXTOSTATEIMAGEMASK(1) and INDEXTOSTATEIMAGEMASK(2)
const LVIS_UNCHECKED: isize = 0x1000; // Checkbox unchecked
const LVIS_CHECKED: isize = 0x2000; // Checkbox checked

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
                    if let Ok(true) = dialog_result
                        && let Err(e) =
                            save_config(&self.available_devices, self.hotkey_vk, self.hotkey_mods)
                    {
                        error!("Failed to save config: {e}");
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
