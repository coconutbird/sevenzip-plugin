//! PROPVARIANT handling for 7-Zip properties.
//!
//! 7-Zip uses 16-byte PROPVARIANT, not the 24-byte version from the windows crate.

use std::ffi::c_void;
use windows::Win32::Foundation::SysFreeString;
use windows::core::BSTR;

use std::time::SystemTime;

/// VT (variant type) constants.
pub const VT_EMPTY: u16 = 0;
pub const VT_UI4: u16 = 19;
pub const VT_UI8: u16 = 21;
pub const VT_BSTR: u16 = 8;
pub const VT_BOOL: u16 = 11;
pub const VT_FILETIME: u16 = 64;

/// Windows FILETIME epoch: January 1, 1601
/// Difference between Unix epoch (1970) and Windows epoch (1601) in 100ns intervals
const FILETIME_UNIX_DIFF: u64 = 116444736000000000;

/// Convert a `SystemTime` to Windows FILETIME format (100ns intervals since 1601-01-01).
pub fn systemtime_to_filetime(time: SystemTime) -> u64 {
    match time.duration_since(std::time::UNIX_EPOCH) {
        Ok(duration) => {
            // Convert to 100ns intervals and add the epoch difference
            let nanos_100 = duration.as_nanos() / 100;
            FILETIME_UNIX_DIFF + nanos_100 as u64
        }
        Err(_) => {
            // Time is before Unix epoch - this is rare but possible
            0
        }
    }
}

/// Raw 16-byte PROPVARIANT matching 7-Zip's expectations.
///
/// The windows crate's PROPVARIANT is 24 bytes which causes crashes with 7-Zip.
#[repr(C)]
pub struct RawPropVariant {
    pub vt: u16,
    pub reserved1: u16,
    pub reserved2: u16,
    pub reserved3: u16,
    pub data: u64,
}

// Verify our struct is the correct size
const _: () = assert!(std::mem::size_of::<RawPropVariant>() == 16);

impl Default for RawPropVariant {
    fn default() -> Self {
        Self {
            vt: VT_EMPTY,
            reserved1: 0,
            reserved2: 0,
            reserved3: 0,
            data: 0,
        }
    }
}

impl RawPropVariant {
    /// Clear any allocated resources and reset to VT_EMPTY.
    ///
    /// This must be called before reusing a PROPVARIANT that may contain
    /// allocated data (like BSTR) to prevent memory leaks.
    ///
    /// # Safety
    /// The caller must ensure this PROPVARIANT owns the data it contains
    /// (i.e., it was set by one of the `set_*` methods, not received from 7-Zip).
    pub unsafe fn clear(&mut self) {
        unsafe {
            if self.vt == VT_BSTR {
                let ptr = self.data as *mut u16;
                if !ptr.is_null() {
                    // Reconstruct the BSTR and free it properly
                    SysFreeString(&BSTR::from_raw(ptr));
                }
            }
            self.vt = VT_EMPTY;
            self.data = 0;
        }
    }

    /// Set to VT_EMPTY (no value).
    ///
    /// Note: This does NOT free any previously allocated data. Call `clear()`
    /// first if this PROPVARIANT may contain allocated resources.
    pub fn set_empty(&mut self) {
        self.vt = VT_EMPTY;
        self.data = 0;
    }

    /// Set a BSTR string value.
    ///
    /// # Safety
    /// - If this PROPVARIANT already contains a BSTR, the caller must call
    ///   `clear()` first to avoid memory leaks.
    /// - The caller must ensure 7-Zip properly frees the allocated BSTR,
    ///   or call `clear()` before dropping this PROPVARIANT.
    pub unsafe fn set_bstr(&mut self, value: &str) {
        unsafe {
            self.clear();
            let bstr = BSTR::from(value);
            self.vt = VT_BSTR;
            self.data = bstr.into_raw() as u64;
        }
    }

    /// Set a u32 value.
    ///
    /// # Safety
    /// If this PROPVARIANT already contains allocated data (like BSTR),
    /// the caller must call `clear()` first to avoid memory leaks.
    pub unsafe fn set_u32(&mut self, value: u32) {
        unsafe {
            self.clear();
            self.vt = VT_UI4;
            self.data = value as u64;
        }
    }

    /// Set a u64 value.
    ///
    /// # Safety
    /// If this PROPVARIANT already contains allocated data (like BSTR),
    /// the caller must call `clear()` first to avoid memory leaks.
    pub unsafe fn set_u64(&mut self, value: u64) {
        unsafe {
            self.clear();
            self.vt = VT_UI8;
            self.data = value;
        }
    }

    /// Set a bool value.
    ///
    /// # Safety
    /// If this PROPVARIANT already contains allocated data (like BSTR),
    /// the caller must call `clear()` first to avoid memory leaks.
    pub unsafe fn set_bool(&mut self, value: bool) {
        unsafe {
            self.clear();
            self.vt = VT_BOOL;
            // VARIANT_TRUE = -1 (0xFFFF), VARIANT_FALSE = 0
            self.data = if value { 0xFFFF } else { 0 };
        }
    }

    /// Set a FILETIME value (100ns intervals since 1601).
    ///
    /// # Safety
    /// If this PROPVARIANT already contains allocated data (like BSTR),
    /// the caller must call `clear()` first to avoid memory leaks.
    pub unsafe fn set_filetime(&mut self, value: u64) {
        unsafe {
            self.clear();
            self.vt = VT_FILETIME;
            self.data = value;
        }
    }

    /// Set a GUID value (as binary blob for ClassId property).
    ///
    /// 7-Zip expects VT_BSTR with raw GUID bytes for ClassId.
    ///
    /// # Safety
    /// - If this PROPVARIANT already contains a BSTR, the caller must call
    ///   `clear()` first to avoid memory leaks.
    /// - The caller must ensure 7-Zip properly frees the allocated BSTR.
    pub unsafe fn set_guid(&mut self, bytes: &[u8; 16]) {
        unsafe {
            self.clear();
            // 7-Zip reads ClassId as VT_BSTR containing raw GUID bytes
            // Allocate a BSTR that contains the 16 bytes
            use windows::Win32::Foundation::SysAllocStringByteLen;

            let bstr = SysAllocStringByteLen(Some(bytes.as_slice()));
            self.vt = VT_BSTR;
            self.data = bstr.into_raw() as u64;
        }
    }

    /// Set a raw byte array value (for signature property).
    ///
    /// 7-Zip expects VT_BSTR containing the raw signature bytes.
    ///
    /// # Safety
    /// - If this PROPVARIANT already contains a BSTR, the caller must call
    ///   `clear()` first to avoid memory leaks.
    /// - The caller must ensure 7-Zip properly frees the allocated BSTR.
    pub unsafe fn set_bytes(&mut self, bytes: &[u8]) {
        unsafe {
            self.clear();
            use windows::Win32::Foundation::SysAllocStringByteLen;

            let bstr = SysAllocStringByteLen(Some(bytes));
            self.vt = VT_BSTR;
            self.data = bstr.into_raw() as u64;
        }
    }

    /// Extract a BSTR string from this PROPVARIANT.
    ///
    /// # Safety
    /// Only call if vt == VT_BSTR.
    pub unsafe fn get_bstr(&self) -> Option<String> {
        unsafe {
            if self.vt != VT_BSTR {
                return None;
            }
            let bstr_ptr = self.data as *const u16;
            if bstr_ptr.is_null() {
                return None;
            }
            // Find null terminator
            let mut len = 0;
            while *bstr_ptr.add(len) != 0 {
                len += 1;
            }
            Some(String::from_utf16_lossy(std::slice::from_raw_parts(
                bstr_ptr, len,
            )))
        }
    }

    /// Extract a u64 value from this PROPVARIANT.
    ///
    /// Returns `Some(value)` if the type is VT_UI8 or VT_UI4, `None` otherwise.
    pub fn get_u64(&self) -> Option<u64> {
        match self.vt {
            VT_UI8 => Some(self.data),
            VT_UI4 => Some(self.data & 0xFFFF_FFFF),
            _ => None,
        }
    }

    /// Extract a u32 value from this PROPVARIANT.
    ///
    /// Returns `Some(value)` if the type is VT_UI4, `None` otherwise.
    pub fn get_u32(&self) -> Option<u32> {
        if self.vt == VT_UI4 {
            Some(self.data as u32)
        } else {
            None
        }
    }

    /// Extract a bool value from this PROPVARIANT.
    ///
    /// Returns `Some(value)` if the type is VT_BOOL, `None` otherwise.
    /// VARIANT_TRUE = -1 (0xFFFF), VARIANT_FALSE = 0
    pub fn get_bool(&self) -> Option<bool> {
        if self.vt == VT_BOOL {
            Some(self.data != 0)
        } else {
            None
        }
    }
}

/// Write a RawPropVariant to a raw pointer (used by 7-Zip callbacks).
///
/// # Safety
/// The destination pointer must be valid and point to at least 16 bytes.
pub unsafe fn write_propvariant(dest: *mut c_void, prop: &RawPropVariant) {
    unsafe {
        std::ptr::copy_nonoverlapping(
            prop as *const RawPropVariant,
            dest as *mut RawPropVariant,
            1,
        );
    }
}
