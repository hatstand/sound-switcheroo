use ::std::error::Error;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Media::Audio::{IMMDeviceEnumerator, MMDeviceEnumerator};
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, STGM_READ,
};
use windows::core::{GUID, PCWSTR, Interface};

// CLSID for the PolicyConfig class
const CLSID_POLICY_CONFIG: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);

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
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        
        // Create the PolicyConfig instance as IUnknown first
        let unknown: windows::core::IUnknown = CoCreateInstance(&CLSID_POLICY_CONFIG, None, CLSCTX_ALL)?;
        
        // Convert device_id to wide string
        let wide_device_id = device_id.encode_utf16().chain(Some(0)).collect::<Vec<u16>>();
        let pcwstr_device_id = PCWSTR::from_raw(wide_device_id.as_ptr());
        
        // Call SetDefaultEndpoint directly using the vtable
        // The SetDefaultEndpoint method is at offset 10 in the vtable (after IUnknown methods and 7 other methods)
        let vtable = *(unknown.as_raw() as *const *const usize);
        let set_default_endpoint_fn: extern "system" fn(*const std::ffi::c_void, PCWSTR, u32) -> windows::core::HRESULT = 
            std::mem::transmute(*(vtable.add(10)));
        
        let hr = set_default_endpoint_fn(unknown.as_raw(), pcwstr_device_id, role);
        hr.ok()?;
        
        Ok(())
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
            println!("Device {}: {:?} {:?}", i, device_id, endpoint.GetState()?);
            let props = endpoint.OpenPropertyStore(STGM_READ)?;
            let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
            println!("  Friendly name: {:?}", friendly_name);
            
            // Example usage: Set the first device as default for console role
            if i == 0 {
                let device_id_str = device_id.to_string()?;
                println!("Setting device as default: {}", device_id_str);
                match set_default_endpoint(&device_id_str, E_ROLE_CONSOLE) {
                    Ok(()) => println!("Successfully set default endpoint!"),
                    Err(e) => println!("Failed to set default endpoint: {}", e),
                }
            }
        }
    };

    Ok(())
}
