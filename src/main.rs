use ::std::error::Error;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Media::Audio::{IMMDeviceEnumerator, MMDeviceEnumerator};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, STGM_READ,
};
use windows::core::{GUID, Interface, PCWSTR};

// CLSID for the PolicyConfig class
const CLSID_POLICY_CONFIG: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);
// const CLSID_POLICY_CONFIG_VISTA: GUID = GUID::from_u128(0x294935CE_F637_4E7C_A41B_AB255460B862);

const IUNKNOWN_VTABLE_SIZE: usize = 3; // IUnknown has 3 methods: QueryInterface, AddRef, Release

// Audio endpoint roles
pub const E_ROLE_CONSOLE: u32 = 0;
pub const E_ROLE_MULTIMEDIA: u32 = 1;
pub const E_ROLE_COMMUNICATIONS: u32 = 2;

/// Sets the default audio endpoint for the specified role using raw COM interface calls
///
/// # Arguments
/// * `device_id` - The device ID string (can be obtained from IMMDevice::GetId())
/// * `role` - The audio endpoint role (0 = Console, 1 = Multimedia, 2 = Communications)
///
/// # Returns
/// * `Ok(())` on success
/// * `Err(Box<dyn Error>)` on failure
pub fn set_default_endpoint(device_id: &str, role: u32) -> Result<(), Box<dyn Error>> {
    unsafe {
        println!(
            "Debug: Attempting to set default endpoint for device: {}, role: {}",
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
        // It seems like the offset should really be either 10 or 13 (10 + IUnknown methods).
        // However, 16 actually works, at least on Windows 11.
        let vtable = *(unknown.as_raw() as *const *const usize);
        let set_default_endpoint_fn: extern "system" fn(
            *const std::ffi::c_void,
            PCWSTR,
            u32,
        ) -> windows::core::HRESULT = std::mem::transmute(*(vtable.add(16)));

        println!("Debug: Calling SetDefaultEndpoint...");
        let hr = set_default_endpoint_fn(unknown.as_raw(), pcwstr_device_id, role);

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
pub fn get_current_default_endpoint(role: u32) -> Result<String, Box<dyn Error>> {
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

        let endpoint = device_enumerator.GetDefaultAudioEndpoint(
            windows::Win32::Media::Audio::eRender,
            match role {
                0 => windows::Win32::Media::Audio::eConsole,
                1 => windows::Win32::Media::Audio::eMultimedia,
                2 => windows::Win32::Media::Audio::eCommunications,
                _ => windows::Win32::Media::Audio::eConsole,
            },
        )?;

        let device_id = endpoint.GetId()?;
        let device_id_str = device_id.to_string()?;

        Ok(device_id_str)
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    println!("Audio Switch Tool");
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
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
                match get_current_default_endpoint(E_ROLE_CONSOLE) {
                    Ok(current) => println!("Current default before change: {}", current),
                    Err(e) => println!("Failed to get current default: {}", e),
                }

                match set_default_endpoint(&device_id_str, E_ROLE_CONSOLE) {
                    Ok(()) => {
                        println!("Successfully set default endpoint!");
                        // Check if it actually changed
                        match get_current_default_endpoint(E_ROLE_CONSOLE) {
                            Ok(current) => println!("Current default after change: {}", current),
                            Err(e) => println!("Failed to get current default after change: {}", e),
                        }
                    }
                    Err(e) => println!("Failed to set default endpoint: {}", e),
                }
            }
        }

        // Get and display the current default endpoint for comparison
        for role in 0..3 {
            match get_current_default_endpoint(role) {
                Ok(device_id_str) => println!(
                    "Current default endpoint for role {}: {}",
                    role, device_id_str
                ),
                Err(e) => println!(
                    "Failed to get current default endpoint for role {}: {}",
                    role, e
                ),
            }
        }
    };

    Ok(())
}
