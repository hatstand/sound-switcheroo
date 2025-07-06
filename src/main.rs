use defer::defer;
use std::error::Error;
use windows::core::PCWSTR;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Media::Audio::{
    eCommunications, eConsole, eMultimedia, ERole, IMMDeviceEnumerator, MMDeviceEnumerator,
};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
    STGM_READ,
};

mod policy_config;
use policy_config::IPolicyConfig;

/// Sets the default audio endpoint for the specified role using raw COM interface calls
fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        println!(
            "Debug: Attempting to set default endpoint for device: {}, role: {:?}",
            device_id, role
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
