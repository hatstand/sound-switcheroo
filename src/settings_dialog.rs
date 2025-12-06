use anyhow::{Context, Result};
use log::{debug, error};
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;
use windows::Win32::Foundation::{GetLastError, HWND, LPARAM, RECT, WPARAM};
use windows::Win32::Graphics::Dwm::{DWMWA_USE_IMMERSIVE_DARK_MODE, DwmSetWindowAttribute};
use windows::Win32::Graphics::Gdi::{
    GetMonitorInfoW, MONITOR_DEFAULTTONEAREST, MONITORINFO, MonitorFromWindow,
};
use windows::Win32::Media::Audio::IMMNotificationClient;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    HKM_GETHOTKEY, HKM_SETHOTKEY, ICC_HOTKEY_CLASS, ICC_LISTVIEW_CLASSES, INITCOMMONCONTROLSEX,
    InitCommonControlsEx, LIST_VIEW_ITEM_STATE_FLAGS, LVCF_TEXT, LVCF_WIDTH, LVCOLUMNW, LVIF_PARAM,
    LVIF_TEXT, LVIS_STATEIMAGEMASK, LVITEMW, LVM_DELETEALLITEMS, LVM_GETITEMCOUNT,
    LVM_GETITEMSTATE, LVM_INSERTCOLUMNW, LVM_INSERTITEMW, LVM_SETEXTENDEDLISTVIEWSTYLE,
    LVM_SETITEMSTATE, LVS_EX_CHECKBOXES,
};
use windows::Win32::UI::WindowsAndMessaging::{
    DLGPROC, DialogBoxParamW, EndDialog, GWLP_USERDATA, GetDlgItem, GetWindowLongPtrW,
    GetWindowRect, ICON_BIG, ICON_SMALL, SWP_NOSIZE, SWP_NOZORDER, SendDlgItemMessageW,
    SendMessageW, SetWindowLongPtrW, SetWindowPos, WM_COMMAND, WM_DESTROY, WM_INITDIALOG,
    WM_SETICON,
};
use windows_core::PCWSTR;

use crate::LOWORD;
use crate::WM_DEVICE_CHANGE;
use crate::audio_device::get_available_audio_devices;
use crate::config::{AudioDevice, merge_device_states};
use crate::icon::{AdaptiveIcon, is_dark_mode};
use crate::safe_strings;

// Dialog resource IDs
const IDD_SETTINGS: u32 = 101;
const IDC_HOTKEY: i32 = 1001;
const IDC_DEVICE_LIST: i32 = 1002;

// ListView checkbox state constants (state image mask values)
// These represent INDEXTOSTATEIMAGEMASK(1) and INDEXTOSTATEIMAGEMASK(2)
const LVIS_UNCHECKED: isize = 0x1000; // Checkbox unchecked
const LVIS_CHECKED: isize = 0x2000; // Checkbox checked

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DialogResult {
    /// User clicked OK/Save button - changes should be saved
    Accepted,
    /// User clicked Cancel or closed the dialog - changes should be discarded
    Cancelled,
}

// Dialog result codes returned by EndDialog and DialogBoxParamW
// Note: 0 indicates an error from DialogBoxParamW, so we use non-zero values
const DIALOG_RESULT_OK: isize = 1;
const DIALOG_RESULT_CANCEL: isize = 2;

/// Result returned by DialogBoxParamW wrapper
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DialogBoxResult {
    Ok,
    Cancel,
}

/// Wrapper for DialogBoxParamW that properly handles return values
/// Returns Err if DialogBoxParamW returns 0 (error) or -1 (error)
fn dialog_box_param(
    hinstance: Option<windows::Win32::Foundation::HINSTANCE>,
    template_name: PCWSTR,
    hwnd_parent: Option<HWND>,
    dialog_func: DLGPROC,
    init_param: LPARAM,
) -> Result<DialogBoxResult> {
    unsafe {
        let result = DialogBoxParamW(
            hinstance,
            template_name,
            hwnd_parent,
            dialog_func,
            init_param,
        );
        match result {
            DIALOG_RESULT_OK => Ok(DialogBoxResult::Ok),
            DIALOG_RESULT_CANCEL => Ok(DialogBoxResult::Cancel),
            0 | -1 => {
                let error = GetLastError();
                anyhow::bail!("DialogBoxParamW failed with error: {:?}", error)
            }
            other => anyhow::bail!("DialogBoxParamW returned unexpected value: {}", other),
        }
    }
}

pub struct SettingsDialog {
    pub devices: Vec<AudioDevice>,
    pub hotkey_vk: u8,
    pub hotkey_mods: u8,
    pub dialog_hwnd_arc: Rc<RefCell<Option<HWND>>>,
}

unsafe fn refresh_device_list(hwnd: HWND, settings: &mut SettingsDialog) -> Result<()> {
    unsafe {
        let list_hwnd =
            GetDlgItem(Some(hwnd), IDC_DEVICE_LIST).context("Failed to get device list control")?;

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
                device_states.insert(device.id.clone(), (state & LVIS_CHECKED) != 0);
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
                let state = if device.selectable {
                    LVIS_CHECKED
                } else {
                    LVIS_UNCHECKED
                };
                let state_lvi = LVITEMW {
                    stateMask: LVIS_STATEIMAGEMASK,
                    state: LIST_VIEW_ITEM_STATE_FLAGS(state as u32),
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
        Ok(())
    }
}

/// Handler for WM_INITDIALOG message
unsafe fn handle_init_dialog(hwnd: HWND, lparam: LPARAM) -> Result<()> {
    unsafe {
        let settings = lparam.0 as *mut SettingsDialog;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, settings as isize);

        let settings_ref = &mut *settings;

        // Store dialog handle in Rc for notifications
        *settings_ref.dialog_hwnd_arc.borrow_mut() = Some(hwnd);

        // Enable dark mode and set icon based on system theme
        let dark_mode = is_dark_mode()?;
        if dark_mode {
            let dark_mode_enabled: u32 = 1;
            DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                &dark_mode_enabled as *const _ as *const _,
                std::mem::size_of::<u32>() as u32,
            )?;
        }

        // Set dialog icon
        let adaptive_icon = AdaptiveIcon::new("switcheroo_icon", "switcheroo_dark_icon")?;
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
        let list_hwnd =
            GetDlgItem(Some(hwnd), IDC_DEVICE_LIST).context("Failed to get device list control")?;

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
                let state = if device.selectable {
                    LVIS_CHECKED
                } else {
                    LVIS_UNCHECKED
                };
                let state_lvi = LVITEMW {
                    stateMask: LVIS_STATEIMAGEMASK,
                    state: LIST_VIEW_ITEM_STATE_FLAGS(state as u32),
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
                SetWindowPos(
                    hwnd,
                    None,
                    center_x,
                    center_y,
                    0,
                    0,
                    SWP_NOSIZE | SWP_NOZORDER,
                )?;
            }
        }

        Ok(())
    }
}

/// Handler for WM_DEVICE_CHANGE message
unsafe fn handle_device_change(hwnd: HWND) -> Result<()> {
    unsafe {
        debug!("Device change detected, refreshing device list");
        let settings = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SettingsDialog;
        if settings.is_null() {
            anyhow::bail!("Settings pointer is null");
        }
        let settings_ref = &mut *settings;
        refresh_device_list(hwnd, settings_ref)?;
        Ok(())
    }
}

/// Handler for WM_COMMAND with IDOK
unsafe fn handle_ok_command(hwnd: HWND) -> Result<()> {
    unsafe {
        let settings = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SettingsDialog;
        if settings.is_null() {
            anyhow::bail!("Settings pointer is null");
        }
        let settings_ref = &mut *settings;

        // Get hotkey value
        let hotkey = SendDlgItemMessageW(hwnd, IDC_HOTKEY, HKM_GETHOTKEY, WPARAM(0), LPARAM(0));
        settings_ref.hotkey_vk = (hotkey.0 & 0xFF) as u8;
        settings_ref.hotkey_mods = ((hotkey.0 >> 8) & 0xFF) as u8;

        // Get device selections
        let list_hwnd =
            GetDlgItem(Some(hwnd), IDC_DEVICE_LIST).context("Failed to get device list control")?;
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
            settings_ref.devices[i as usize].selectable = (state & LVIS_CHECKED) != 0;
        }

        EndDialog(hwnd, DIALOG_RESULT_OK)?;
        Ok(())
    }
}

/// Handler for WM_COMMAND with IDCANCEL
unsafe fn handle_cancel_command(hwnd: HWND) -> Result<()> {
    unsafe {
        EndDialog(hwnd, DIALOG_RESULT_CANCEL)?;
        Ok(())
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
            // Dialog initialization - called when the dialog is first created.
            // Sets up dark mode, icons, hotkey control, device list, and centers the dialog.
            // Returns 1 to set focus to the control specified by the system.
            WM_INITDIALOG => {
                if let Err(e) = handle_init_dialog(hwnd, lparam) {
                    error!("Error in WM_INITDIALOG: {}", e);
                    return 0;
                }
                1
            }
            // Dialog destruction - called when the dialog is being destroyed.
            // No cleanup needed as dialog handle is already cleared.
            // Returns 0 as per Windows convention for WM_DESTROY.
            WM_DESTROY => 0,
            // Custom message sent by IMMNotificationClient when audio devices change.
            // Refreshes the device list to reflect current system state.
            // Returns 1 to indicate the message was processed.
            WM_DEVICE_CHANGE => {
                if let Err(e) = handle_device_change(hwnd) {
                    error!("Error in WM_DEVICE_CHANGE: {}", e);
                    return 0;
                }
                1
            }
            // Command message sent when controls (buttons) are activated.
            // Handles OK (save settings) and Cancel (discard changes) buttons.
            WM_COMMAND => {
                let cmd = LOWORD!(wparam.0) as u32;
                match cmd {
                    // OK button clicked - save hotkey and device selections, close dialog.
                    // Returns 1 to indicate the command was processed.
                    1 => {
                        // IDOK
                        if let Err(e) = handle_ok_command(hwnd) {
                            error!("Error in IDOK command: {}", e);
                            return 0;
                        }
                        1
                    }
                    // Cancel button clicked - discard changes and close dialog.
                    // Returns 1 to indicate the command was processed.
                    2 => {
                        // IDCANCEL
                        if let Err(e) = handle_cancel_command(hwnd) {
                            error!("Error in IDCANCEL command: {}", e);
                            return 0;
                        }
                        1
                    }
                    // Unknown command ID - return 0 to indicate not processed.
                    _ => 0,
                }
            }
            // Unhandled message - return 0 to let system default processing occur.
            _ => 0,
        }
    }
}

pub fn show_settings_dialog(
    parent: HWND,
    devices: &mut Vec<AudioDevice>,
    hotkey_vk: &mut u8,
    hotkey_mods: &mut u8,
    _notification_client: &IMMNotificationClient,
    dialog_hwnd_arc: &Rc<RefCell<Option<HWND>>>,
) -> Result<DialogResult> {
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
        let result = dialog_box_param(
            Some(module.into()),
            PCWSTR(IDD_SETTINGS as *const u16),
            Some(parent),
            Some(settings_dialog_proc),
            LPARAM(&mut settings as *mut _ as isize),
        )?;

        // Clear dialog handle after dialog closes
        *dialog_hwnd_arc.borrow_mut() = None;

        match result {
            DialogBoxResult::Ok => {
                *devices = settings.devices;
                *hotkey_vk = settings.hotkey_vk;
                *hotkey_mods = settings.hotkey_mods;
                Ok(DialogResult::Accepted)
            }
            DialogBoxResult::Cancel => Ok(DialogResult::Cancelled),
        }
    }
}
