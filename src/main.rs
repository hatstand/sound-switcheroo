use ::std::error::Error;
use ::std::ffi::c_void;
use defer::defer;
use windows::Devices::Custom::DeviceSharingMode;
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::PROPERTYKEY;
use windows::Win32::Media::Audio::{ERole, WAVEFORMATEX};
use windows::Win32::Media::Audio::{
    IMMDeviceEnumerator, MMDeviceEnumerator, eCommunications, eConsole, eMultimedia,
};
use windows::Win32::System::Com::StructuredStorage::PROPVARIANT;
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx, CoUninitialize,
    STGM_READ,
};
use windows::core::HRESULT;
use windows::core::{GUID, Interface, PCWSTR};
use windows_core::BOOL;

// See https://github.com/Belphemur/AudioEndPointLibrary/blob/master/DefSound/PolicyConfig.h

// CLSID for the PolicyConfig class
const CLSID_POLICY_CONFIG: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);

// VTable for the undocumented PolicyConfig interface.
#[repr(C)]
#[doc(hidden)]
pub struct IPolicyConfig_Vtbl {
    pub base__: ::windows::core::IUnknown_Vtbl,

    // TODO: Confused.
    // The known interface does not include these methods, but they fix the offset for SetDefaultEndpoint.
    // pub MysteryMethod1: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    // pub MysteryMethod2: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
    // pub MysteryMethod3: unsafe extern "system" fn(this: *mut c_void) -> HRESULT,
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

    pub SetDefaultEndpoint: unsafe extern "system" fn(this: *mut c_void, PCWSTR, ERole) -> HRESULT,
    pub SetEndpointVisibility: unsafe extern "system" fn(this: *mut c_void, PCWSTR, i32) -> HRESULT,
}

windows_core::imp::define_interface!(
    IPolicyConfig,
    IPolicyConfig_Vtbl,
    0xf8679f50_850a_41cf_9c72_430f290290c8
);
windows_core::imp::interface_hierarchy!(IPolicyConfig, windows_core::IUnknown);
impl IPolicyConfig {
    pub unsafe fn GetMixFormat<P0>(
        &self,
        device_name: P0,
    ) -> windows_core::Result<*mut WAVEFORMATEX>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed::<*mut WAVEFORMATEX>();
            (Interface::vtable(self).GetMixFormat)(
                Interface::as_raw(self),
                device_name.param().abi(),
                &mut result__,
            )
            .and_then(|| windows_core::Type::from_abi(result__))
        }
    }

    pub unsafe fn GetDeviceFormat<P0>(
        &self,
        device_name: P0,
        default: impl Into<BOOL>,
    ) -> windows_core::Result<*mut WAVEFORMATEX>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed::<*mut WAVEFORMATEX>();
            (Interface::vtable(self).GetDeviceFormat)(
                Interface::as_raw(self),
                device_name.param().abi(),
                default.into().0,
                &mut result__,
            )
            .and_then(|| windows_core::Type::from_abi(result__))
        }
    }

    pub unsafe fn ResetDeviceFormat<P0>(&self, device_name: P0) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).ResetDeviceFormat)(
                Interface::as_raw(self),
                device_name.param().abi(),
            )
            .ok()
        }
    }

    pub unsafe fn SetDeviceFormat<P0>(
        &self,
        device_name: P0,
        mut endpoint_format: WAVEFORMATEX,
        mut mix_format: WAVEFORMATEX,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).SetDeviceFormat)(
                Interface::as_raw(self),
                device_name.param().abi(),
                &mut endpoint_format,
                &mut mix_format,
            )
            .ok()
        }
    }

    pub unsafe fn GetProcessingPeriod<P0>(
        &self,
        device_name: P0,
        default: impl Into<BOOL>,
        default_period: *mut i64,
        min_period: *mut i64,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        (Interface::vtable(self).GetProcessingPeriod)(
            Interface::as_raw(self),
            device_name.param().abi(),
            default.into().0,
            default_period,
            min_period,
        )
        .ok()
    }

    pub unsafe fn SetProcessingPeriod<P0>(
        &self,
        device_name: P0,
        period: *mut i64,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        (Interface::vtable(self).SetProcessingPeriod)(
            Interface::as_raw(self),
            device_name.param().abi(),
            period,
        )
        .ok()
    }

    pub unsafe fn GetShareMode<P0>(
        &self,
        device_name: P0,
    ) -> windows_core::Result<DeviceSharingMode>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed::<DeviceSharingMode>();
            (Interface::vtable(self).GetShareMode)(
                Interface::as_raw(self),
                device_name.param().abi(),
                &mut result__,
            )
            .and_then(|| windows_core::Type::from_abi(result__))
        }
    }

    pub unsafe fn SetShareMode<P0>(
        &self,
        device_name: P0,
        mut mode: DeviceSharingMode,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).SetShareMode)(
                Interface::as_raw(self),
                device_name.param().abi(),
                &mut mode,
            )
            .ok()
        }
    }

    pub unsafe fn GetPropertyValue<P0>(
        &self,
        device_name: P0,
        bFxStore: impl Into<BOOL>,
        key: *const PROPERTYKEY,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed::<PROPVARIANT>();
            (Interface::vtable(self).GetPropertyValue)(
                Interface::as_raw(self),
                device_name.param().abi(),
                bFxStore.into().0,
                key,
                &mut result__,
            )
            .ok()
        }
    }

    pub unsafe fn SetPropertyValue<P0>(
        &self,
        device_name: P0,
        bFxStore: impl Into<BOOL>,
        key: *const PROPERTYKEY,
        propvar: *mut PROPVARIANT,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).SetPropertyValue)(
                Interface::as_raw(self),
                device_name.param().abi(),
                bFxStore.into().0,
                key,
                propvar,
            )
            .ok()
        }
    }

    pub unsafe fn SetDefaultEndpoint<P0>(
        &self,
        device_name: P0,
        role: ERole,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).SetDefaultEndpoint)(
                Interface::as_raw(self),
                device_name.param().abi(),
                role,
            )
            .ok()
        }
    }

    pub unsafe fn SetEndpointVisibility<P0>(
        &self,
        device_name: P0,
        visible: impl Into<BOOL>,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).SetEndpointVisibility)(
                Interface::as_raw(self),
                device_name.param().abi(),
                visible.into().0,
            )
            .ok()
        }
    }
}

fn to_pcwstr(s: &str) -> PCWSTR {
    let wide: Vec<u16> = s.encode_utf16().chain(Some(0)).collect();
    PCWSTR::from_raw(wide.as_ptr())
}

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
            CoCreateInstance(&CLSID_POLICY_CONFIG, None, CLSCTX_ALL)?;

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
