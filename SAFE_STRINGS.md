# Safe Windows String Handling

## Problem
Converting Rust `String` or `&str` to Windows `PCWSTR` or `PWSTR` is inherently unsafe because:
1. The pointer can outlive the UTF-16 data, leading to use-after-free bugs
2. Windows APIs expect null-terminated wide strings (UTF-16)
3. Raw pointer manipulation is required

## Solution
We implemented a safe scoped approach using helper functions in a dedicated module (`src/safe_strings.rs`):

## Usage Examples

### Before (Unsafe):
```rust
fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        let wide_device_id = device_id
            .encode_utf16()
            .chain(Some(0))
            .collect::<Vec<u16>>();
        let pcwstr_device_id = PCWSTR::from_raw(wide_device_id.as_ptr());
        
        // BUG: pcwstr_device_id can outlive wide_device_id!
        policy_config.SetDefaultEndpoint(pcwstr_device_id, role)?;
    }
}
```

### After (Safe):
```rust
fn set_default_endpoint(device_id: &str, role: ERole) -> Result<(), Box<dyn Error>> {
    unsafe {
        let policy_config: IPolicyConfig = 
            CoCreateInstance(&policy_config::CLSID_POLICY_CONFIG, None, CLSCTX_ALL)?;
        
        // Safe: wide string data is guaranteed to live for the duration of the closure
        with_wide_str(device_id, |wide_device_id| {
            policy_config.SetDefaultEndpoint(wide_device_id, role)
        })?;
        Ok(())
    }
}
```