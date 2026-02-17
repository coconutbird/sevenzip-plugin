//! COM interface definitions and GUIDs for 7-Zip.
//!
//! This module uses `cppvtable`'s `#[com_interface]` macro to define
//! 7-Zip's COM interfaces with proper vtable layout and calling conventions.

use std::ffi::c_void;

// Import COM types from cppvtable (with windows-compat, GUID = windows_core::GUID)
pub use cppvtable::IID_IUNKNOWN;
// IUnknown is used by the #[com_interface] macro for base vtable generation
use cppvtable::IUnknown;
pub use cppvtable::com::GUID;
use cppvtable::proc::com_interface;

// Re-export HRESULT from windows crate (cppvtable's HRESULT is a wrapper, windows uses it directly)
pub use windows::core::HRESULT;

/// Create a 7-Zip format GUID from the format ID byte.
pub const fn make_format_guid(id: u8) -> GUID {
    GUID::from_u128(0x23170f69_40c1_278a_1000_000110000000_u128 | ((id as u128) << 32))
}

// =============================================================================
// IInArchive - 7-Zip archive reading interface
// =============================================================================

/// IInArchive interface for reading archives.
///
/// This interface is used by 7-Zip to read archive contents.
/// The type parameter `T` is used for type-safe vtable function pointers.
/// Methods:
///   - Slots 0-2: IUnknown (query_interface, add_ref, release)
///   - Slots 3-15: IInArchive methods (13 methods total)
#[com_interface("23170f69-40c1-278a-0000-000600600000")]
pub trait IInArchive<T> {
    /// Open an archive from a stream.
    fn open(
        &mut self,
        stream: *mut c_void,
        max_check_start_position: *const u64,
        open_callback: *mut c_void,
    ) -> HRESULT;

    /// Close the archive.
    fn close(&mut self) -> HRESULT;

    /// Get the number of items in the archive.
    fn get_number_of_items(&self, num_items: *mut u32) -> HRESULT;

    /// Get a property of an item.
    fn get_property(&self, index: u32, prop_id: u32, value: *mut c_void) -> HRESULT;

    /// Extract items from the archive.
    fn extract(
        &mut self,
        indices: *const u32,
        num_items: u32,
        test_mode: i32,
        extract_callback: *mut c_void,
    ) -> HRESULT;

    /// Get an archive-level property.
    fn get_archive_property(&self, prop_id: u32, value: *mut c_void) -> HRESULT;

    /// Get the number of supported item properties.
    fn get_number_of_properties(&self, num_props: *mut u32) -> HRESULT;

    /// Get information about a property.
    fn get_property_info(
        &self,
        index: u32,
        name: *mut c_void,
        prop_id: *mut u32,
        var_type: *mut u32,
    ) -> HRESULT;

    /// Get the number of supported archive properties.
    fn get_number_of_archive_properties(&self, num_props: *mut u32) -> HRESULT;

    /// Get information about an archive property.
    fn get_archive_property_info(
        &self,
        index: u32,
        name: *mut c_void,
        prop_id: *mut u32,
        var_type: *mut u32,
    ) -> HRESULT;
}

// =============================================================================
// IOutArchive - 7-Zip archive writing interface
// =============================================================================

/// IOutArchive interface for writing archives.
///
/// This interface is used by 7-Zip to create or modify archives.
/// The type parameter `T` is used for type-safe vtable function pointers.
/// Methods:
///   - Slots 0-2: IUnknown (query_interface, add_ref, release)
///   - Slots 3-4: IOutArchive methods (2 methods)
#[com_interface("23170f69-40c1-278a-0000-000600a00000")]
pub trait IOutArchive<T> {
    /// Update items in the archive (add, modify, or delete).
    fn update_items(
        &mut self,
        out_stream: *mut c_void,
        num_items: u32,
        update_callback: *mut c_void,
    ) -> HRESULT;

    /// Get the file time type used by this archive format.
    fn get_file_time_type(&self, time_type: *mut u32) -> HRESULT;
}

// =============================================================================
// ICryptoGetTextPassword - Password callback for reading encrypted archives
// =============================================================================

/// ICryptoGetTextPassword interface for password prompts when reading.
///
/// This interface is implemented by 7-Zip's callback object and queried
/// by the plugin when it needs a password to decrypt an archive.
/// Methods:
///   - Slots 0-2: IUnknown (query_interface, add_ref, release)
///   - Slot 3: CryptoGetTextPassword
#[com_interface("23170f69-40c1-278a-0000-000500100000")]
pub trait ICryptoGetTextPassword<T> {
    /// Get the password for decrypting an archive.
    /// Returns S_OK and sets password, or error if cancelled.
    fn crypto_get_text_password(&self, password: *mut *mut u16) -> HRESULT;
}

// =============================================================================
// ICryptoGetTextPassword2 - Password callback for creating encrypted archives
// =============================================================================

/// ICryptoGetTextPassword2 interface for password prompts when writing.
///
/// This interface is implemented by 7-Zip's callback object and queried
/// by the plugin when it needs a password for creating encrypted archives.
/// Methods:
///   - Slots 0-2: IUnknown (query_interface, add_ref, release)
///   - Slot 3: CryptoGetTextPassword2
#[com_interface("23170f69-40c1-278a-0000-000500110000")]
pub trait ICryptoGetTextPassword2<T> {
    /// Get the password for encrypting an archive.
    /// Returns S_OK, sets password_is_defined (non-zero if password provided),
    /// and sets password if defined.
    fn crypto_get_text_password2(
        &self,
        password_is_defined: *mut i32,
        password: *mut *mut u16,
    ) -> HRESULT;
}

// =============================================================================
// Stream Interfaces
// =============================================================================

/// ISequentialInStream - Sequential read-only stream.
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slot 3: Read
#[com_interface("23170f69-40c1-278a-0000-000300010000")]
pub trait ISequentialInStream<T> {
    /// Read data from stream.
    /// Returns S_OK if some data was read (check processed_size).
    /// Returns S_OK with processed_size=0 at end of stream.
    fn read(&self, data: *mut u8, size: u32, processed_size: *mut u32) -> HRESULT;
}

/// ISequentialOutStream - Sequential write-only stream.
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slot 3: Write
#[com_interface("23170f69-40c1-278a-0000-000300020000")]
pub trait ISequentialOutStream<T> {
    /// Write data to stream.
    /// Returns S_OK if some data was written (check processed_size).
    fn write(&self, data: *const u8, size: u32, processed_size: *mut u32) -> HRESULT;
}

/// IInStream - Seekable input stream (extends ISequentialInStream).
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slot 3: Read (from ISequentialInStream)
///   - Slot 4: Seek
#[com_interface("23170f69-40c1-278a-0000-000300030000")]
pub trait IInStream<T> {
    /// Read data from stream.
    fn read(&self, data: *mut u8, size: u32, processed_size: *mut u32) -> HRESULT;

    /// Seek within the stream.
    /// origin: 0 = SEEK_SET, 1 = SEEK_CUR, 2 = SEEK_END
    fn seek(&self, offset: i64, seek_origin: u32, new_position: *mut u64) -> HRESULT;
}

/// IOutStream - Seekable output stream (extends ISequentialOutStream).
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slot 3: Write (from ISequentialOutStream)
///   - Slot 4: Seek
///   - Slot 5: SetSize
#[com_interface("23170f69-40c1-278a-0000-000300040000")]
pub trait IOutStream<T> {
    /// Write data to stream.
    fn write(&self, data: *const u8, size: u32, processed_size: *mut u32) -> HRESULT;

    /// Seek within the stream.
    fn seek(&self, offset: i64, seek_origin: u32, new_position: *mut u64) -> HRESULT;

    /// Set the stream size.
    fn set_size(&mut self, new_size: u64) -> HRESULT;
}

// =============================================================================
// Progress and Callback Interfaces
// =============================================================================

/// IProgress - Base progress interface.
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slot 3: SetTotal
///   - Slot 4: SetCompleted
#[com_interface("23170f69-40c1-278a-0000-000000050000")]
pub trait IProgress<T> {
    /// Set total amount of work.
    fn set_total(&self, total: u64) -> HRESULT;

    /// Set completed amount of work.
    fn set_completed(&self, complete_value: *const u64) -> HRESULT;
}

/// IArchiveExtractCallback - Callback for extraction operations.
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slots 3-4: IProgress (SetTotal, SetCompleted)
///   - Slot 5: GetStream
///   - Slot 6: PrepareOperation
///   - Slot 7: SetOperationResult
#[com_interface("23170f69-40c1-278a-0000-000600200000")]
pub trait IArchiveExtractCallback<T> {
    /// Set total amount to extract.
    fn set_total(&self, total: u64) -> HRESULT;

    /// Set completed amount.
    fn set_completed(&self, complete_value: *const u64) -> HRESULT;

    /// Get output stream for item at index.
    /// ask_extract_mode: 0 = extract, 1 = test, 2 = skip
    fn get_stream(
        &self,
        index: u32,
        out_stream: *mut *mut c_void,
        ask_extract_mode: i32,
    ) -> HRESULT;

    /// Prepare for operation.
    /// ask_extract_mode: 0 = extract, 1 = test, 2 = skip
    fn prepare_operation(&self, ask_extract_mode: i32) -> HRESULT;

    /// Set operation result.
    fn set_operation_result(&self, result_e_operation_result: i32) -> HRESULT;
}

/// IArchiveUpdateCallback - Callback for update/create operations.
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slots 3-4: IProgress (SetTotal, SetCompleted)
///   - Slot 5: GetUpdateItemInfo
///   - Slot 6: GetProperty
///   - Slot 7: GetStream
///   - Slot 8: SetOperationResult
#[com_interface("23170f69-40c1-278a-0000-000600800000")]
pub trait IArchiveUpdateCallback<T> {
    /// Set total amount to process.
    fn set_total(&self, total: u64) -> HRESULT;

    /// Set completed amount.
    fn set_completed(&self, complete_value: *const u64) -> HRESULT;

    /// Get update info for item.
    fn get_update_item_info(
        &self,
        index: u32,
        new_data: *mut i32,
        new_props: *mut i32,
        index_in_archive: *mut u32,
    ) -> HRESULT;

    /// Get property of item.
    fn get_property(&self, index: u32, prop_id: u32, value: *mut c_void) -> HRESULT;

    /// Get input stream for item.
    fn get_stream(&self, index: u32, in_stream: *mut *mut c_void) -> HRESULT;

    /// Set operation result.
    fn set_operation_result(&self, operation_result: i32) -> HRESULT;
}

/// IArchiveOpenCallback - Callback for archive open operations.
///
/// Methods:
///   - Slots 0-2: IUnknown
///   - Slot 3: SetTotal
///   - Slot 4: SetCompleted
#[com_interface("23170f69-40c1-278a-0000-000600100000")]
pub trait IArchiveOpenCallback<T> {
    /// Set total size (may be called multiple times or not at all).
    fn set_total(&self, files: *const u64, bytes: *const u64) -> HRESULT;

    /// Set completed progress.
    fn set_completed(&self, files: *const u64, bytes: *const u64) -> HRESULT;
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
    Crc = 19,
    Encrypted = 21,
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
