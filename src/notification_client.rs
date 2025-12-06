use log::debug;
use std::cell::RefCell;
use std::rc::Rc;
use windows::Win32;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::Media::Audio::{ERole, IMMNotificationClient, IMMNotificationClient_Impl};
use windows::Win32::UI::WindowsAndMessaging::PostMessageW;

use crate::{WM_DEVICE_CHANGE, get_device_friendly_name};

#[windows::core::implement(IMMNotificationClient)]
pub struct CustomImmNotificationClient {
    pub settings_dialog: Rc<RefCell<Option<HWND>>>,
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
