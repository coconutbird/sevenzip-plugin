//! COM interface definitions and GUIDs for 7-Zip.

use std::ffi::c_void;
use windows::core::{GUID, HRESULT};

/// Standard COM interface GUIDs
pub const IID_IUNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_C000_000000000046);
pub const IID_IINARCHIVE: GUID = GUID::from_u128(0x23170f69_40c1_278a_0000_000600600000);
pub const IID_IOUTARCHIVE: GUID = GUID::from_u128(0x23170f69_40c1_278a_0000_000600a00000);

/// Create a 7-Zip format GUID from the format ID byte.
pub const fn make_format_guid(id: u8) -> GUID {
    GUID::from_u128(0x23170f69_40c1_278a_1000_000110000000_u128 | ((id as u128) << 32))
}

/// IInArchive vtable layout matching 7-Zip SDK.
#[repr(C)]
pub struct IInArchiveVtbl<T> {
    // IUnknown (3 methods)
    pub query_interface:
        unsafe extern "system" fn(*mut T, *const GUID, *mut *mut c_void) -> HRESULT,
    pub add_ref: unsafe extern "system" fn(*mut T) -> u32,
    pub release: unsafe extern "system" fn(*mut T) -> u32,
    // IInArchive methods (13 methods)
    pub open: unsafe extern "system" fn(*mut T, *mut c_void, *const u64, *mut c_void) -> HRESULT,
    pub close: unsafe extern "system" fn(*mut T) -> HRESULT,
    pub get_number_of_items: unsafe extern "system" fn(*mut T, *mut u32) -> HRESULT,
    pub get_property: unsafe extern "system" fn(*mut T, u32, u32, *mut c_void) -> HRESULT,
    pub extract: unsafe extern "system" fn(*mut T, *const u32, u32, i32, *mut c_void) -> HRESULT,
    pub get_archive_property: unsafe extern "system" fn(*mut T, u32, *mut c_void) -> HRESULT,
    pub get_number_of_properties: unsafe extern "system" fn(*mut T, *mut u32) -> HRESULT,
    pub get_property_info:
        unsafe extern "system" fn(*mut T, u32, *mut c_void, *mut u32, *mut u32) -> HRESULT,
    pub get_number_of_archive_properties: unsafe extern "system" fn(*mut T, *mut u32) -> HRESULT,
    pub get_archive_property_info:
        unsafe extern "system" fn(*mut T, u32, *mut c_void, *mut u32, *mut u32) -> HRESULT,
}

/// IOutArchive vtable layout matching 7-Zip SDK.
#[repr(C)]
pub struct IOutArchiveVtbl<T> {
    // IUnknown (3 methods)
    pub query_interface:
        unsafe extern "system" fn(*mut T, *const GUID, *mut *mut c_void) -> HRESULT,
    pub add_ref: unsafe extern "system" fn(*mut T) -> u32,
    pub release: unsafe extern "system" fn(*mut T) -> u32,
    // IOutArchive methods (2 methods)
    pub update_items: unsafe extern "system" fn(*mut T, *mut c_void, u32, *mut c_void) -> HRESULT,
    pub get_file_time_type: unsafe extern "system" fn(*mut T, *mut u32) -> HRESULT,
}

/// Property IDs used by 7-Zip.
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PropId {
    Path = 3,
    IsDir = 6,
    Size = 7,
    PackSize = 8,
    Attrib = 9,
    CTime = 10,
    ATime = 11,
    MTime = 12,
}

/// Archive property IDs.
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArchivePropId {
    PhySize = 4,
}

/// Handler property IDs for GetHandlerProperty2.
#[repr(u32)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HandlerPropId {
    Name = 0,
    ClassId = 1,
    Extension = 2,
    AddExtension = 3,
    Update = 4,
    KeepName = 5,
    Signature = 6,
    MultiSignature = 7,
    SignatureOffset = 8,
    AltStreams = 9,
    NtSecure = 10,
    Flags = 11,
}
