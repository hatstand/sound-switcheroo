use windows::Win32::UI::Input::KeyboardAndMouse::HOT_KEY_MODIFIERS;

/// Converts hotkey modifiers from HOTKEYF_ format (used by hotkey control)
/// to MOD_ format (used by RegisterHotKey)
pub fn hotkeyf_to_mod(hotkeyf_mods: u8) -> HOT_KEY_MODIFIERS {
    use windows::Win32::UI::Controls::{HOTKEYF_ALT, HOTKEYF_CONTROL, HOTKEYF_SHIFT};
    use windows::Win32::UI::Input::KeyboardAndMouse::{MOD_ALT, MOD_CONTROL, MOD_SHIFT};

    let mut mods = HOT_KEY_MODIFIERS(0);
    if hotkeyf_mods & HOTKEYF_SHIFT as u8 != 0 {
        mods |= MOD_SHIFT;
    }
    if hotkeyf_mods & HOTKEYF_CONTROL as u8 != 0 {
        mods |= MOD_CONTROL;
    }
    if hotkeyf_mods & HOTKEYF_ALT as u8 != 0 {
        mods |= MOD_ALT;
    }
    mods
}

/// Converts a string to a fixed-size UTF-16 array suitable for Windows tooltips
pub fn string_to_tip(s: &str) -> [u16; 128] {
    let mut ret = [0u16; 128];
    let encoded: Vec<u16> = s.encode_utf16().collect();
    assert!(encoded.len() < ret.len());
    for (i, &c) in encoded.iter().enumerate() {
        ret[i] = c;
    }
    ret[encoded.len()] = 0; // Null-terminate the string
    ret
}

/// Helper function for extracting low word from a 32-bit value
#[allow(non_snake_case)]
pub fn LOWORD(l: isize) -> isize {
    l & 0xffff
}
