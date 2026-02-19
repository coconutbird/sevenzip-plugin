//! DLL export generation for 7-Zip plugins.
//!
//! This module provides macros and functions to generate the required
//! DLL exports for a 7-Zip plugin.

use std::ffi::c_void;
use windows::Win32::Foundation::{CLASS_E_CLASSNOTAVAILABLE, E_INVALIDARG, S_OK};
use windows::core::{GUID, HRESULT};

use super::com::{HandlerPropId, IID_IINARCHIVE, IID_IOUTARCHIVE};
use super::propvariant::RawPropVariant;

/// Macro to register a format and generate all required DLL exports.
///
/// # Example
///
/// ```rust,ignore
/// use sevenzip_plugin::prelude::*;
///
/// #[derive(Default)]
/// struct MyFormat { /* ... */ }
///
/// impl ArchiveFormat for MyFormat { /* ... */ }
/// impl ArchiveReader for MyFormat { /* ... */ }
///
/// // Read-only format
/// sevenzip_plugin::register_format!(MyFormat);
///
/// // Format that supports editing (requires ArchiveUpdater impl)
/// sevenzip_plugin::register_format!(MyFormat, updatable);
/// ```
#[macro_export]
macro_rules! register_format {
    // Internal implementation - shared by both variants
    (@impl $format:ty, $out_vtbl_fn:path) => {
        static IN_VTBL: $crate::windows::com::IInArchiveVTable<
            $crate::windows::handler::PluginHandler<$format>,
        > = $crate::windows::handler::create_in_vtable::<$format>();

        static OUT_VTBL: $crate::windows::com::IOutArchiveVTable<
            $crate::windows::handler::PluginHandler<$format>,
        > = $out_vtbl_fn();

        static REGISTERED_FORMAT: $crate::windows::handler::RegisteredFormat<$format> =
            $crate::windows::handler::RegisteredFormat::new(&IN_VTBL, &OUT_VTBL);

        #[unsafe(no_mangle)]
        pub unsafe extern "system" fn CreateObject(
            clsid: *const $crate::windows_crate::core::GUID,
            iid: *const $crate::windows_crate::core::GUID,
            out_object: *mut *mut ::std::ffi::c_void,
        ) -> $crate::windows_crate::core::HRESULT {
            unsafe {
                $crate::windows::exports::create_object::<$format>(
                    clsid,
                    iid,
                    out_object,
                    &REGISTERED_FORMAT,
                )
            }
        }

        #[unsafe(no_mangle)]
        pub unsafe extern "system" fn GetNumberOfFormats(
            num_formats: *mut u32,
        ) -> $crate::windows_crate::core::HRESULT {
            unsafe {
                if num_formats.is_null() {
                    return $crate::windows_crate::Win32::Foundation::E_INVALIDARG;
                }
                *num_formats = 1;
                $crate::windows_crate::Win32::Foundation::S_OK
            }
        }

        #[unsafe(no_mangle)]
        pub unsafe extern "system" fn GetHandlerProperty2(
            format_index: u32,
            prop_id: u32,
            value: *mut ::std::ffi::c_void,
        ) -> $crate::windows_crate::core::HRESULT {
            unsafe {
                $crate::windows::exports::get_handler_property2::<$format>(
                    format_index,
                    prop_id,
                    value,
                )
            }
        }
    };

    // Read-only format (no updatable flag)
    ($format:ty) => {
        $crate::register_format!(@impl $format, $crate::windows::handler::create_out_vtable_stub::<$format>);
    };

    // Updatable format (with updatable flag)
    ($format:ty, updatable) => {
        $crate::register_format!(@impl $format, $crate::windows::handler::create_out_vtable::<$format>);
    };
}

/// Log a message to the debug file (if debug feature is enabled).
/// Uses a macro to ensure format arguments are not evaluated in release builds.
#[cfg(feature = "debug")]
macro_rules! log_debug {
    ($($arg:tt)*) => {{
        use std::io::Write;
        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open("C:\\temp\\7zip-plugin-debug.log")
        {
            let _ = writeln!(file, $($arg)*);
        }
    }};
}

#[cfg(not(feature = "debug"))]
macro_rules! log_debug {
    ($($arg:tt)*) => {};
}

/// Convert 16-byte array to GUID at compile time.
pub const fn guid_from_bytes(bytes: &[u8; 16]) -> GUID {
    GUID {
        data1: u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]),
        data2: u16::from_le_bytes([bytes[4], bytes[5]]),
        data3: u16::from_le_bytes([bytes[6], bytes[7]]),
        data4: [
            bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15],
        ],
    }
}

/// Implementation of CreateObject for a registered format.
///
/// # Safety
/// - `clsid`, `iid`, and `out_object` must be valid pointers if non-null
/// - The caller must ensure proper COM reference counting
pub unsafe fn create_object<T: crate::ArchiveReader>(
    clsid: *const GUID,
    iid: *const GUID,
    out_object: *mut *mut c_void,
    format: &'static super::handler::RegisteredFormat<T>,
) -> HRESULT {
    unsafe {
        log_debug!("CreateObject called");

        if clsid.is_null() || iid.is_null() || out_object.is_null() {
            log_debug!("CreateObject: null pointer");
            return E_INVALIDARG;
        }

        let clsid = &*clsid;
        let iid = &*iid;

        // Check if the CLSID matches our format (computed at compile time per monomorphization)
        const { assert!(std::mem::size_of::<GUID>() == 16) };
        let format_guid = guid_from_bytes(&T::class_id());

        log_debug!(
            "CreateObject: clsid={{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            clsid.data1,
            clsid.data2,
            clsid.data3,
            clsid.data4[0],
            clsid.data4[1],
            clsid.data4[2],
            clsid.data4[3],
            clsid.data4[4],
            clsid.data4[5],
            clsid.data4[6],
            clsid.data4[7]
        );
        log_debug!(
            "CreateObject: format_guid={{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            format_guid.data1,
            format_guid.data2,
            format_guid.data3,
            format_guid.data4[0],
            format_guid.data4[1],
            format_guid.data4[2],
            format_guid.data4[3],
            format_guid.data4[4],
            format_guid.data4[5],
            format_guid.data4[6],
            format_guid.data4[7]
        );

        if *clsid != format_guid {
            log_debug!("CreateObject: CLSID mismatch!");
            *out_object = std::ptr::null_mut();
            return CLASS_E_CLASSNOTAVAILABLE;
        }

        // Create handler
        let handler = format.create_handler();

        // Return appropriate interface
        if *iid == IID_IINARCHIVE {
            *out_object = handler;
            return S_OK;
        }

        if *iid == IID_IOUTARCHIVE && T::supports_write() {
            // Return pointer to out_vtbl field
            let handler_ptr = handler as *mut super::handler::PluginHandler<T>;
            *out_object = &(*handler_ptr).out_vtbl as *const _ as *mut c_void;
            return S_OK;
        }

        // Unknown interface
        *out_object = std::ptr::null_mut();
        CLASS_E_CLASSNOTAVAILABLE
    }
}

/// Implementation of GetHandlerProperty2 for a format.
pub fn get_handler_property2<T: crate::ArchiveFormat>(
    format_index: u32,
    prop_id: u32,
    value: *mut c_void,
) -> HRESULT {
    unsafe {
        if format_index != 0 || value.is_null() {
            return E_INVALIDARG;
        }

        let prop = &mut *(value as *mut RawPropVariant);

        match prop_id {
            x if x == HandlerPropId::Name as u32 => {
                prop.set_bstr(T::name());
            }
            x if x == HandlerPropId::ClassId as u32 => {
                // Return GUID as binary blob
                let guid_bytes = T::class_id();
                prop.set_guid(&guid_bytes);
            }
            x if x == HandlerPropId::Extension as u32 => {
                prop.set_bstr(T::extension());
            }
            x if x == HandlerPropId::Update as u32 => {
                prop.set_bool(T::supports_write());
            }
            x if x == HandlerPropId::Signature as u32 => {
                // Return signature bytes for format auto-detection
                if let Some(sig) = T::signature() {
                    prop.set_bytes(sig);
                } else {
                    prop.set_empty();
                }
            }
            x if x == HandlerPropId::SignatureOffset as u32 => {
                // Signature starts at offset 0 by default
                prop.set_u32(0);
            }
            _ => {
                prop.set_empty();
            }
        }

        S_OK
    }
}
