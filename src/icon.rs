use anyhow::{Result, bail};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{HICON, LoadIconW};

use crate::safe_strings::with_wide_str;

/// A struct that holds both light and dark mode icons and returns the appropriate one
/// based on the current system theme
#[derive(Debug)]
pub struct AdaptiveIcon {
    pub light: HICON,
    pub dark: HICON,
}

impl AdaptiveIcon {
    pub fn new(light_icon_name: &str, dark_icon_name: &str) -> Result<Self> {
        let light_icon = unsafe { load_icon(light_icon_name)? };
        let dark_icon = unsafe { load_icon(dark_icon_name)? };
        Ok(Self {
            light: light_icon,
            dark: dark_icon,
        })
    }

    pub fn icon(&self) -> Result<HICON> {
        if is_dark_mode()? {
            Ok(self.dark)
        } else {
            Ok(self.light)
        }
    }
}

/// Checks if Windows is currently in dark mode
pub fn is_dark_mode() -> Result<bool> {
    let theme_key = windows_registry::CURRENT_USER
        .open(r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")?;
    let light_theme = theme_key.get_u32("AppsUseLightTheme")? == 1;
    Ok(!light_theme)
}

/// Loads an icon from the current module's resources
unsafe fn load_icon(icon_name: &str) -> Result<HICON> {
    unsafe {
        let module = GetModuleHandleW(None)?;
        let icon = with_wide_str(icon_name, |wide_icon_name| {
            LoadIconW(Some(module.into()), wide_icon_name)
        })?;
        if icon.is_invalid() {
            bail!("Failed to load icon: {}", icon_name);
        }
        Ok(icon)
    }
}
