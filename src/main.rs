#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use defer::defer;
use log::{debug, error, info};
use simple_error::bail;
use std::error::Error;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::{GetLastError, HWND, LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::Media::Audio::{
    ERole, IMMDeviceEnumerator, MMDeviceEnumerator, eCommunications, eConsole, eMultimedia,
};
use windows::Win32::System::Com::StructuredStorage::PROPVARIANT;
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
    STGM_READ,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Variant::VT_LPWSTR;
use windows::Win32::UI::Shell::{
    NIF_GUID, NIF_ICON, NIF_MESSAGE, NIF_SHOWTIP, NIF_TIP, NIM_ADD, NIM_DELETE, NIM_SETVERSION,
    NOTIFYICON_VERSION_4, NOTIFYICONDATAW, NOTIFYICONDATAW_0, Shell_NotifyIconW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DispatchMessageW, GWLP_USERDATA,
    GetCursorPos, GetMessageW, GetWindowLongPtrW, HMENU, InsertMenuItemW, LoadIconW, MENUITEMINFOW,
    MFT_STRING, MIIM_FTYPE, MIIM_ID, MIIM_STATE, MIIM_STRING, MSG, PostMessageW, PostQuitMessage,
    RegisterClassExW, SetForegroundWindow, SetWindowLongPtrW, TPM_BOTTOMALIGN, TPM_LEFTALIGN,
    TPM_RIGHTBUTTON, TrackPopupMenuEx, UnregisterClassW, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP,
    WM_CLOSE, WM_COMMAND, WM_DESTROY, WM_QUIT, WM_RBUTTONUP, WNDCLASSEXW,
};
use windows::core::PCWSTR;
use windows_core::{BOOL, GUID, PWSTR};

mod policy_config;
use policy_config::IPolicyConfig;

/// Sets the default audio endpoint for the specified role using raw COM interface calls
fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        println!(
            "Debug: Attempting to set default endpoint for device: {device_id}, role: {role:?}",
        );

        // Create the PolicyConfig instance as IUnknown first
        println!("Debug: Creating PolicyConfig COM instance...");
        let policy_config: IPolicyConfig =
            CoCreateInstance(&policy_config::CLSID_POLICY_CONFIG, None, CLSCTX_ALL)?;

        // Convert device_id to wide string
        let wide_device_id = device_id
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>();
        let pcwstr_device_id = PCWSTR::from_raw(wide_device_id.as_ptr());

        policy_config.SetDefaultEndpoint(pcwstr_device_id, role)?;
        Ok(())
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

fn to_pcwstr(s: &str) -> PCWSTR {
    let mut v = s.encode_utf16().collect::<Vec<u16>>();
    v.push(0);
    PCWSTR(v.as_ptr())
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
struct AudioDevice {
    id: String,
    friendly_name: String,
}

#[derive(Debug)]
struct AudioSwitch {
    window: HWND,
    popup_menu: HMENU,
    available_devices: Vec<AudioDevice>,
    current_device: String,
}

impl AudioSwitch {
    fn hello(&self) {
        info!("Hello from AudioSwitch!: {self:?}");
    }
    fn show_popup_menu(&self, x: i32, y: i32) -> Result<(), Box<dyn Error>> {
        debug!("Showing popup menu at ({}, {})", x, y);
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
        debug!("Menu item selected: {}", id);
        unsafe {
            match id {
                POPUP_EXIT_ID => {
                    debug!("Exit selected");
                    PostMessageW(
                        Some(self.window),
                        WM_CLOSE,
                        WPARAM::default(),
                        LPARAM::default(),
                    )?;
                }
                _ => {
                    debug!("Unknown menu item selected: {}", id);
                    return Ok(());
                }
            }
        }
        Ok(())
    }
}

const POPUP_EXIT_ID: u32 = 1;

unsafe fn create_popup_menu(
    devices: &Vec<AudioDevice>,
    current_device: &str,
) -> Result<HMENU, Box<dyn Error>> {
    unsafe {
        let menu = CreatePopupMenu()?;
        debug!("Popup menu created: {:?}", menu);
        // Add a menu item to exit the application.
        InsertMenuItemW(
            menu,
            0,
            true,
            &MENUITEMINFOW {
                cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                fMask: MIIM_FTYPE | MIIM_ID | MIIM_STRING,
                fType: MFT_STRING,
                dwTypeData: PWSTR("Exit\0".encode_utf16().collect::<Vec<u16>>().as_mut_ptr()),
                cch: "Exit".len() as u32,
                wID: POPUP_EXIT_ID,
                ..Default::default()
            },
        )?;

        for (i, device) in devices.iter().rev().enumerate() {
            debug!("Adding device to popup menu: {:?}", device);
            InsertMenuItemW(
                menu,
                0,
                true,
                &MENUITEMINFOW {
                    cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
                    fMask: MIIM_FTYPE | MIIM_ID | MIIM_STRING | MIIM_STATE,
                    fType: MFT_STRING,
                    fState: if device.id == current_device {
                        windows::Win32::UI::WindowsAndMessaging::MFS_CHECKED
                    } else {
                        windows::Win32::UI::WindowsAndMessaging::MFS_UNCHECKED
                    },
                    dwTypeData: PWSTR(
                        device
                            .friendly_name
                            .encode_utf16()
                            .chain(Some(0))
                            .collect::<Vec<u16>>()
                            .as_mut_ptr(),
                    ),
                    cch: device.friendly_name.chars().count() as u32,
                    wID: POPUP_EXIT_ID + i as u32, // Unique ID for each device
                    ..Default::default()
                },
            )?;
        }

        Ok(menu)
    }
}

unsafe fn propvariant_to_string(propvar: &PROPVARIANT) -> Result<String, Box<dyn Error>> {
    unsafe {
        match propvar.vt() {
            VT_LPWSTR => {
                return Ok(String::from_utf16_lossy(
                    propvar.Anonymous.Anonymous.Anonymous.pwszVal.as_wide(),
                ));
            }
            _ => {
                bail!("Unsupported PROPVARIANT type: {:?}", propvar.vt());
            }
        };
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
            devices.push(AudioDevice {
                id: device_id_str,
                friendly_name: propvariant_to_string(&friendly_name)?,
            });
        }
    }
    Ok(devices)
}

fn main() -> Result<(), Box<dyn Error>> {
    env_logger::init();
    info!("Audio Switch Tool");
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        defer!({
            CoUninitialize();
        });
        let module = GetModuleHandleW(None)?;
        // Register a window class for the taskbar icon.
        let class_name = to_pcwstr("AudioSwitchTool");
        let class = RegisterClassExW(&WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(window_callback),
            hInstance: module.into(),
            lpszClassName: class_name,
            ..Default::default()
        });
        debug!("Class registered: {:?}", class);
        defer!({
            // Unregister the class when done.
            let _ = UnregisterClassW(class_name, Some(module.into()));
        });

        // Seems this needs to _not_ be a message-only window for ShellExecute to work.
        let window = CreateWindowExW(
            WINDOW_EX_STYLE(0),
            PCWSTR(class as *const u16),
            to_pcwstr("Audio Switch Tool"),
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
        debug!("Window created: {:?}", window);
        let devices = get_available_audio_devices()?;
        let current_device = get_current_default_endpoint(eConsole)?;
        let me = AudioSwitch {
            window,
            popup_menu: create_popup_menu(&devices, &current_device)?,
            available_devices: devices,
            current_device: current_device,
        };
        // Store the AudioSwitch instance in the window's user data.
        SetWindowLongPtrW(window, GWLP_USERDATA, &me as *const _ as isize);
        let icon = LoadIconW(Some(module.into()), to_pcwstr("audio_icon"))?;
        let guid = GUID::new()?;
        let notify_icon_data = &mut NOTIFYICONDATAW {
            cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
            hWnd: window,
            hIcon: icon,
            guidItem: guid,
            // Both NIF_TIP & NIF_SHOWTIP are required to actually show the tooltip.
            uFlags: NIF_ICON | NIF_MESSAGE | NIF_GUID | NIF_TIP | NIF_SHOWTIP,
            uCallbackMessage: WM_APP + 0x42,
            szTip: string_to_tip("Audio Switch Tool"),
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
                    guidItem: guid,
                    ..Default::default()
                },
            );
        });
        // Enable better callback API.
        Shell_NotifyIconW(NIM_SETVERSION, notify_icon_data).ok()?;

        // Enter the message loop.
        info!("Running...");
        loop {
            let mut msg = MSG::default();
            debug!("Waiting for message...");
            match GetMessageW(&mut msg, None, 0, 0) {
                BOOL(0) => {
                    assert_eq!(msg.message, WM_QUIT);
                    info!("Quitting...");
                    break;
                }
                BOOL(-1) => {
                    error!("Failed to get message: {:?}", GetLastError());
                }
                BOOL(_) => {
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
    debug!(
        "Window callback: hwnd={:?}, msg={:#x}, wparam={:#x}, lparam={:#x}",
        hwnd, msg, wparam.0, lparam.0
    );
    unsafe {
        let raw_me = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AudioSwitch;
        if !raw_me.is_null() {
            raw_me.as_mut().unwrap().hello();
        }
        match msg {
            TASKBAR_CB_ID => match LOWORD(lparam.0) as u32 {
                WM_RBUTTONUP => {
                    debug!("Right click received");
                    let mut cursor_pos = POINT::default();
                    GetCursorPos(&mut cursor_pos).unwrap();
                    match raw_me
                        .as_mut()
                        .unwrap()
                        .show_popup_menu(cursor_pos.x, cursor_pos.y)
                    {
                        Ok(()) => debug!("Popup menu shown successfully"),
                        Err(e) => error!("Failed to show popup menu: {:?}", e),
                    }
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            },
            WM_COMMAND => {
                debug!("Menu Command received");
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
