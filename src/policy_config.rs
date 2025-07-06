#![allow(non_snake_case)]
#![allow(non_camel_case_types)]

use ::std::ffi::c_void;
use windows::Devices::Custom::DeviceSharingMode;
use windows::Win32::Foundation::PROPERTYKEY;
use windows::Win32::Media::Audio::{ERole, WAVEFORMATEX};
use windows::Win32::System::Com::StructuredStorage::PROPVARIANT;
use windows::core::{HRESULT, Interface, PCWSTR};
use windows_core::{BOOL, GUID};

// See https://github.com/Belphemur/AudioEndPointLibrary/blob/master/DefSound/PolicyConfig.h

// CLSID for the PolicyConfig class
pub const CLSID_POLICY_CONFIG: GUID = GUID::from_u128(0x870af99c_171d_4f9e_af0d_e63df40c2bc9);

// VTable for the undocumented PolicyConfig interface.
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
        unsafe {
            (Interface::vtable(self).GetProcessingPeriod)(
                Interface::as_raw(self),
                device_name.param().abi(),
                default.into().0,
                default_period,
                min_period,
            )
            .ok()
        }
    }

    pub unsafe fn SetProcessingPeriod<P0>(
        &self,
        device_name: P0,
        period: *mut i64,
    ) -> windows_core::Result<()>
    where
        P0: windows_core::Param<PCWSTR>,
    {
        unsafe {
            (Interface::vtable(self).SetProcessingPeriod)(
                Interface::as_raw(self),
                device_name.param().abi(),
                period,
            )
            .ok()
        }
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
