use log::debug;
use simple_error::bail;
use std::error::Error;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Media::Audio::{
    DEVICE_STATE_ACTIVE, ERole, EndpointFormFactor, IMMDeviceEnumerator, MMDeviceEnumerator,
    PKEY_AudioEndpoint_FormFactor, eRender,
};
use windows::Win32::System::Com::StructuredStorage::PROPVARIANT;
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, STGM_READ,
};
use windows::Win32::System::Variant::{VT_LPWSTR, VT_UI4};

use crate::config::AudioDevice;
use crate::policy_config::IPolicyConfig;
use crate::safe_strings::with_wide_str;

/// Sets the default audio endpoint for the specified role using raw COM interface calls
pub fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        debug!("Attempting to set default endpoint for device: {device_id}, role: {role:?}",);
        let policy_config: IPolicyConfig =
            CoCreateInstance(&crate::policy_config::CLSID_POLICY_CONFIG, None, CLSCTX_ALL)?;

        // Use safe scoped approach for string conversion
        with_wide_str(device_id, |wide_device_id| {
            policy_config.SetDefaultEndpoint(wide_device_id, role)
        })?;
        Ok(())
    }
}

/// Gets the friendly name for a device given its ID
pub fn get_device_friendly_name(device_id: &str) -> Result<String, Box<dyn Error>> {
    unsafe {
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

        let device = with_wide_str(device_id, |wide_id| device_enumerator.GetDevice(wide_id))?;

        let props = device.OpenPropertyStore(STGM_READ)?;
        let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
        propvariant_to_string(&friendly_name)
    }
}

/// Gets the current default audio endpoint for debugging
pub fn get_current_default_endpoint(role: ERole) -> Result<String, Box<dyn Error>> {
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()?;
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

        let endpoint = device_enumerator.GetDefaultAudioEndpoint(eRender, role)?;

        let device_id = endpoint.GetId()?;
        let device_id_str = device_id.to_string()?;

        Ok(device_id_str)
    }
}

/// Converts a PROPVARIANT to a String
unsafe fn propvariant_to_string(propvar: &PROPVARIANT) -> Result<String, Box<dyn Error>> {
    unsafe {
        match propvar.vt() {
            VT_LPWSTR => Ok(String::from_utf16_lossy(
                propvar.Anonymous.Anonymous.Anonymous.pwszVal.as_wide(),
            )),
            _ => {
                bail!("Unsupported PROPVARIANT type: {:?}", propvar.vt());
            }
        }
    }
}

/// Enumerates all available audio devices
pub fn get_available_audio_devices() -> Result<Vec<AudioDevice>, Box<dyn Error>> {
    let mut devices = Vec::new();
    unsafe {
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let endpoints = device_enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)?;

        for i in 0..endpoints.GetCount()? {
            let endpoint = endpoints.Item(i)?;
            let device_id = endpoint.GetId()?;
            let device_id_str = device_id.to_string()?;
            let props = endpoint.OpenPropertyStore(STGM_READ)?;
            let friendly_name = props.GetValue(&PKEY_Device_FriendlyName)?;
            let form_factor_var = props.GetValue(&PKEY_AudioEndpoint_FormFactor)?;
            let form_factor: EndpointFormFactor = match form_factor_var.vt() {
                VT_UI4 => {
                    EndpointFormFactor(form_factor_var.Anonymous.Anonymous.Anonymous.ulVal as i32)
                }
                _ => {
                    bail!(
                        "Unsupported PROPVARIANT type for form factor: {:?}",
                        form_factor_var,
                    );
                }
            };
            devices.push(AudioDevice {
                id: device_id_str,
                friendly_name: propvariant_to_string(&friendly_name)?,
                selectable: true,
                form_factor,
            });
        }
    }
    Ok(devices)
}
