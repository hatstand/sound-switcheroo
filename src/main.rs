#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use defer::defer;
use log::{debug, error, info};
use std::error::Error;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::{GetLastError, LPARAM, LRESULT, WPARAM};
use windows::Win32::Media::Audio::{
    ERole, IMMDeviceEnumerator, MMDeviceEnumerator, eCommunications, eConsole, eMultimedia,
};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
    STGM_READ,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::{
    NIF_GUID, NIF_ICON, NIF_MESSAGE, NIF_SHOWTIP, NIF_TIP, NIM_ADD, NIM_DELETE, NIM_SETVERSION,
    NOTIFYICON_VERSION_4, NOTIFYICONDATAW, NOTIFYICONDATAW_0, Shell_NotifyIconW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, GWLP_USERDATA, GetMessageW,
    GetWindowLongPtrW, LoadIconW, MSG, PostQuitMessage, RegisterClassExW, SetWindowLongPtrW,
    UnregisterClassW, WINDOW_EX_STYLE, WINDOW_STYLE, WM_APP, WM_DESTROY, WM_QUIT, WNDCLASSEXW,
};
use windows::core::PCWSTR;
use windows_core::{BOOL, GUID};

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
struct AudioSwitch {}

impl AudioSwitch {
    fn hello(&self) {
        info!("Hello from AudioSwitch!: {self:?}");
    }
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
        let me = Box::new(AudioSwitch {});
        SetWindowLongPtrW(window, GWLP_USERDATA, Box::into_raw(me) as isize);
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

        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let endpoints = device_enumerator.EnumAudioEndpoints(
            windows::Win32::Media::Audio::eRender,
            windows::Win32::Media::Audio::DEVICE_STATE_ACTIVE,
        )?;

        info!("Available audio devices:");
        for i in 0..endpoints.GetCount()? {
            let endpoint = endpoints.Item(i)?;
            let device_id = endpoint.GetId()?;
            let device_id_str = device_id.to_string()?;
            info!(
                "Device {}: {:?} {:?}",
                i,
                device_id_str,
                endpoint.GetState()?
            );
            let props = endpoint.OpenPropertyStore(STGM_READ)?;
            let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
            info!("  Friendly name: {friendly_name:?}");
            info!("  Device ID (raw): {device_id:?}");
            info!("  Device ID (string): {device_id_str}");

            // Example usage: Set the first device as default for console role
            if i == 0 {
                info!("Setting device as default: {device_id_str}");

                // Show current default before change
                match get_current_default_endpoint(windows::Win32::Media::Audio::eConsole) {
                    Ok(current) => info!("Current default before change: {current}"),
                    Err(e) => error!("Failed to get current default: {e}"),
                }

                match set_default_endpoint(&device_id_str, windows::Win32::Media::Audio::eConsole) {
                    Ok(()) => {
                        debug!("Successfully set default endpoint!");
                        // Check if it actually changed
                        match get_current_default_endpoint(windows::Win32::Media::Audio::eConsole) {
                            Ok(current) => info!("Current default after change: {current}"),
                            Err(e) => error!("Failed to get current default after change: {e}"),
                        }
                    }
                    Err(e) => error!("Failed to set default endpoint: {e}"),
                }
            }
        }

        // Get and display the current default endpoint for comparison
        for role in [eConsole, eMultimedia, eCommunications] {
            match get_current_default_endpoint(role) {
                Ok(device_id_str) => {
                    info!("Current default endpoint for role {role:?}: {device_id_str:?}")
                }
                Err(e) => error!("Failed to get current default endpoint for role {role:?}: {e}"),
            }
        }

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
            WM_DESTROY => {
                PostQuitMessage(0);
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        }
    }
}
