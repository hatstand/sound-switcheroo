use ::std::error::Error;
use ::windows_core::GUID;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Media::Audio::{IMMDeviceEnumerator, MMDeviceEnumerator};
use windows::Win32::System::Com::{
    CLSCTX_ALL, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx,
    STGM_READ,
};
use windows_core::{IUnknown, Interface, implement};

const CPOLICY_CONFIG_CLIENT_GUID: GUID = GUID::from_u128(179502765086287242678282948019479129033);

pub struct IPolicyConfig {}

fn main() -> Result<(), Box<dyn Error>> {
    println!("Hello, world!");
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let endpoints = device_enumerator.EnumAudioEndpoints(
            windows::Win32::Media::Audio::eRender,
            windows::Win32::Media::Audio::DEVICE_STATE_ACTIVE,
        )?;
        for i in 0..endpoints.GetCount()? {
            let endpoint = endpoints.Item(i)?;
            let device_id = endpoint.GetId()?;
            println!("Device {}: {:?} {:?}", i, device_id, endpoint.GetState()?);
            let props = endpoint.OpenPropertyStore(STGM_READ)?;
            let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
            println!("Friendly name: {:?}", friendly_name);
        }
    };

    Ok(())
}
