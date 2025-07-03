use ::std::error::Error;
use defer::defer;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::PROPERTYKEY;
use windows::Win32::Media::Audio::{
    IMMDeviceEnumerator, MMDeviceEnumerator, eCommunications, eConsole, eMultimedia,
};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
    STGM_READ,
};
use windows::core::{GUID, Interface, PCWSTR};

use std::ffi::c_void;
use windows::Devices::Custom::DeviceSharingMode;
use windows::Win32::Media::Audio::{ERole, WAVEFORMATEX};
use windows::Win32::System::Com::StructuredStorage::PROPVARIANT;
use windows::core::HRESULT;

// CLSID for the PolicyConfig class
const CLSID_POLICY_CONFIG: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);

#[repr(C)]
#[doc(hidden)]
pub struct IPolicyConfig_Vtbl {
    pub base__: ::windows::core::IUnknown_Vtbl,
    pub GetMixFormat:
        unsafe extern "system" fn(this: *mut c_void, PCWSTR, *mut *mut WAVEFORMATEX) -> HRESULT,
    pub GetDeviceFormat: unsafe extern "system" fn(
        this: *mut c_void,
        PCWSTR,
        i32,
        *mut *mut WAVEFORMATEX,
    ) -> HRESULT,
    pub ResetDeviceFormat: unsafe extern "system" fn(this: *mut c_void, PCWSTR) -> HRESULT,
    pub SetDeviceFormat: unsafe extern "system" fn(
        this: *mut c_void,
        PCWSTR,
        *mut WAVEFORMATEX,
        *mut WAVEFORMATEX,
    ) -> HRESULT,
    pub GetProcessingPeriod:
        unsafe extern "system" fn(this: *mut c_void, PCWSTR, i32, *mut i64, *mut i64) -> HRESULT,
    pub SetProcessingPeriod:
        unsafe extern "system" fn(this: *mut c_void, PCWSTR, *mut i64) -> HRESULT,
    pub GetShareMode:
        unsafe extern "system" fn(this: *mut c_void, PCWSTR, *mut DeviceSharingMode) -> HRESULT,
    pub SetShareMode:
        unsafe extern "system" fn(this: *mut c_void, PCWSTR, *mut DeviceSharingMode) -> HRESULT,
    pub GetPropertyValue: unsafe extern "system" fn(
        this: *mut c_void,
        PCWSTR,
        i32,
        *const PROPERTYKEY,
        *mut PROPVARIANT,
    ) -> HRESULT,
    pub SetPropertyValue: unsafe extern "system" fn(
        this: *mut c_void,
        PCWSTR,
        i32,
        *const PROPERTYKEY,
        *mut PROPVARIANT,
    ) -> HRESULT,

    // The known interface does not include these methods, but they fix the offset for SetDefaultEndpoint.
    pub MysteryMethod1: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub MysteryMethod2: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    pub MysteryMethod3: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,

    pub SetDefaultEndpoint: unsafe extern "system" fn(this: *mut c_void, PCWSTR, ERole) -> HRESULT,
    pub SetEndpointVisibility: unsafe extern "system" fn(this: *mut c_void, PCWSTR, i32) -> HRESULT,
}

/// Sets the default audio endpoint for the specified role using raw COM interface calls
///
pub fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        println!(
            "Debug: Attempting to set default endpoint for device: {}, role: {:?}",
            device_id, role
        );

        // Create the PolicyConfig instance as IUnknown first
        println!("Debug: Creating PolicyConfig COM instance...");
        let unknown: windows::core::IUnknown =
            match CoCreateInstance(&CLSID_POLICY_CONFIG, None, CLSCTX_ALL) {
                Ok(instance) => {
                    println!("Debug: PolicyConfig instance created successfully");
                    instance
                }
                Err(e) => {
                    println!("Debug: Failed to create PolicyConfig instance: {:?}", e);
                    return Err(format!("Failed to create PolicyConfig instance: {:?}", e).into());
                }
            };

        // Convert device_id to wide string
        let wide_device_id = device_id
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>();
        let pcwstr_device_id = PCWSTR::from_raw(wide_device_id.as_ptr());

        println!(
            "Debug: Device ID converted to wide string, length: {}",
            wide_device_id.len()
        );

        // Call SetDefaultEndpoint directly using the vtable
        let vtable: *const IPolicyConfig_Vtbl =
            std::mem::transmute(*(unknown.as_raw() as *const *const usize));

        let first_real_method = &(*vtable).GetMixFormat;

        let set_default_endpoint_fn = &(*vtable).SetDefaultEndpoint;
        let actual = vtable.add(16);
        println!(
            "Debug: VTable address: {:p}, FnAddr: {:p} ActualAddr: {:p} GetMixFormat: {:p}",
            vtable, set_default_endpoint_fn, actual, first_real_method,
        );

        println!("Debug: Calling SetDefaultEndpoint...");
        let hr = set_default_endpoint_fn(unknown.as_raw(), pcwstr_device_id, role);
        // let hr = set_default_endpoint_fn(unknown.as_raw(), pcwstr_device_id, role);

        println!("Debug: SetDefaultEndpoint returned HRESULT: 0x{:08X}", hr.0);

        if hr.is_ok() {
            println!("Debug: SetDefaultEndpoint succeeded");
        } else {
            println!(
                "Debug: SetDefaultEndpoint failed with HRESULT: 0x{:08X}",
                hr.0
            );
            return Err(format!("SetDefaultEndpoint failed with HRESULT: 0x{:08X}", hr.0).into());
        }

        Ok(())
    }
}

/// Gets the current default audio endpoint for debugging
pub fn get_current_default_endpoint(role: ERole) -> Result<String, Box<dyn Error>> {
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

fn main() -> Result<(), Box<dyn Error>> {
    println!("Audio Switch Tool");
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        defer!({
            CoUninitialize();
        });
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let endpoints = device_enumerator.EnumAudioEndpoints(
            windows::Win32::Media::Audio::eRender,
            windows::Win32::Media::Audio::DEVICE_STATE_ACTIVE,
        )?;

        println!("Available audio devices:");
        for i in 0..endpoints.GetCount()? {
            let endpoint = endpoints.Item(i)?;
            let device_id = endpoint.GetId()?;
            let device_id_str = device_id.to_string()?;
            println!(
                "Device {}: {:?} {:?}",
                i,
                device_id_str,
                endpoint.GetState()?
            );
            let props = endpoint.OpenPropertyStore(STGM_READ)?;
            let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
            println!("  Friendly name: {:?}", friendly_name);
            println!("  Device ID (raw): {:?}", device_id);
            println!("  Device ID (string): {}", device_id_str);

            // Example usage: Set the first device as default for console role
            if i == 0 {
                println!("Setting device as default: {}", device_id_str);

                // Show current default before change
                match get_current_default_endpoint(windows::Win32::Media::Audio::eConsole) {
                    Ok(current) => println!("Current default before change: {}", current),
                    Err(e) => println!("Failed to get current default: {}", e),
                }

                match set_default_endpoint(&device_id_str, windows::Win32::Media::Audio::eConsole) {
                    Ok(()) => {
                        println!("Successfully set default endpoint!");
                        // Check if it actually changed
                        match get_current_default_endpoint(windows::Win32::Media::Audio::eConsole) {
                            Ok(current) => println!("Current default after change: {}", current),
                            Err(e) => println!("Failed to get current default after change: {}", e),
                        }
                    }
                    Err(e) => println!("Failed to set default endpoint: {}", e),
                }
            }
        }

        // Get and display the current default endpoint for comparison
        for role in [eConsole, eMultimedia, eCommunications] {
            match get_current_default_endpoint(role) {
                Ok(device_id_str) => println!(
                    "Current default endpoint for role {:?}: {:?}",
                    role, device_id_str
                ),
                Err(e) => println!(
                    "Failed to get current default endpoint for role {:?}: {}",
                    role, e
                ),
            }
        }
    };

    Ok(())
}
