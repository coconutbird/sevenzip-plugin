//! Generic COM handler wrapper that bridges safe traits to 7-Zip interfaces.

use std::ffi::c_void;
use std::io::{Read, Seek, SeekFrom};
use std::marker::PhantomData;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::{
    E_INVALIDARG, E_NOINTERFACE, E_NOTIMPL, E_POINTER, S_FALSE, S_OK,
};
use windows::core::{BSTR, HRESULT};

use crate::traits::{ArchiveReader, ArchiveUpdater};

use super::com::{
    ArchivePropId,
    GUID,
    IArchiveExtractCallback,
    IArchiveUpdateCallback,
    ICryptoGetTextPassword,
    ICryptoGetTextPassword2,
    IID_ICRYPTOGETTEXTPASSWORD,
    IID_ICRYPTOGETTEXTPASSWORD2,
    IID_IINARCHIVE,
    IID_IOUTARCHIVE,
    IID_IUNKNOWN,
    IInArchiveVTable,
    IInStream,
    IOutArchiveVTable,
    // Stream and callback wrapper types
    ISequentialInStream,
    ISequentialOutStream,
    PropId,
};

// Import IUnknownVTable for vtable base field initialization
use cppvtable::IUnknownVTable;

use super::propvariant::RawPropVariant;
use crate::types::{PasswordProvider, PasswordRequester};

// Stream seek origins
const STREAM_SEEK_SET: u32 = 0;
const STREAM_SEEK_CUR: u32 = 1;
const STREAM_SEEK_END: u32 = 2;

/// Read data from ISequentialInStream
pub(crate) unsafe fn read_sequential_stream(stream: *mut c_void) -> std::io::Result<Vec<u8>> {
    unsafe {
        if stream.is_null() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "Null stream pointer",
            ));
        }

        let stream = ISequentialInStream::<c_void>::from_ptr_mut(stream);

        let mut data = Vec::new();
        let mut buffer = [0u8; 65536];

        loop {
            let mut bytes_read: u32 = 0;
            let hr = stream.read(buffer.as_mut_ptr(), buffer.len() as u32, &mut bytes_read);
            if hr.is_err() {
                return Err(std::io::Error::other(format!("Read failed: {:?}", hr)));
            }
            if bytes_read == 0 {
                break;
            }
            data.extend_from_slice(&buffer[..bytes_read as usize]);
        }

        Ok(data)
    }
}

// =============================================================================
// Streaming Input Reader
// =============================================================================

/// Wrapper for IInStream that implements `std::io::Read + Seek`.
///
/// This allows zero-copy streaming reads from 7-Zip's input stream,
/// avoiding the need to buffer the entire archive in memory.
pub struct InStreamReader {
    stream: *mut c_void,
    size: u64,
}

impl InStreamReader {
    /// Create a new InStreamReader from a raw IInStream pointer.
    ///
    /// # Safety
    /// The stream pointer must be valid and point to a valid IInStream COM object.
    pub unsafe fn new(stream: *mut c_void) -> std::io::Result<Self> {
        unsafe {
            if stream.is_null() {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "Null stream pointer",
                ));
            }

            let in_stream = IInStream::<c_void>::from_ptr_mut(stream);

            // Get stream size by seeking to end
            let mut size: u64 = 0;
            let hr = in_stream.seek(0, STREAM_SEEK_END, &mut size);
            if hr.is_err() {
                return Err(std::io::Error::other(format!(
                    "Failed to get stream size: {:?}",
                    hr
                )));
            }

            // Seek back to start
            let mut pos: u64 = 0;
            let hr = in_stream.seek(0, STREAM_SEEK_SET, &mut pos);
            if hr.is_err() {
                return Err(std::io::Error::other(format!(
                    "Failed to seek to start: {:?}",
                    hr
                )));
            }

            Ok(Self { stream, size })
        }
    }

    /// Get the total size of the stream in bytes.
    pub fn size(&self) -> u64 {
        self.size
    }

    /// Get the underlying stream as a typed wrapper.
    #[inline]
    fn as_stream(&mut self) -> &mut IInStream<c_void> {
        unsafe { IInStream::<c_void>::from_ptr_mut(self.stream) }
    }
}

impl Read for InStreamReader {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }

        let chunk_size = buf.len().min(u32::MAX as usize) as u32;
        let mut bytes_read: u32 = 0;

        let hr = unsafe {
            self.as_stream()
                .read(buf.as_mut_ptr(), chunk_size, &mut bytes_read)
        };

        if hr.is_err() {
            return Err(std::io::Error::other(format!(
                "Read failed with HRESULT: {:?}",
                hr
            )));
        }

        Ok(bytes_read as usize)
    }
}

impl Seek for InStreamReader {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        let (offset, origin) = match pos {
            SeekFrom::Start(n) => (n as i64, STREAM_SEEK_SET),
            SeekFrom::Current(n) => (n, STREAM_SEEK_CUR),
            SeekFrom::End(n) => (n, STREAM_SEEK_END),
        };

        let mut new_pos: u64 = 0;
        let hr = unsafe { self.as_stream().seek(offset, origin, &mut new_pos) };

        if hr.is_err() {
            return Err(std::io::Error::other(format!(
                "Seek failed with HRESULT: {:?}",
                hr
            )));
        }

        Ok(new_pos)
    }
}

// =============================================================================
// Generic Plugin Handler
// =============================================================================

/// Generic COM handler that wraps a safe archive implementation.
///
/// This struct is the bridge between 7-Zip's COM interfaces and your safe Rust traits.
#[repr(C)]
pub struct PluginHandler<T: ArchiveReader> {
    /// Pointer to IInArchive vtable - MUST be first field for COM compatibility
    pub in_vtbl: *const IInArchiveVTable<Self>,
    /// Pointer to IOutArchive vtable - for writing support
    pub out_vtbl: *const IOutArchiveVTable<Self>,
    /// Reference count
    ref_count: AtomicU32,
    /// The actual archive implementation (safe Rust)
    pub(crate) inner: T,
    /// Input stream pointer (AddRef'd, must Release on close)
    pub(crate) in_stream: *mut c_void,
    /// Physical archive size
    pub(crate) archive_size: u64,
    /// Is archive open
    pub(crate) is_open: bool,
}

impl<T: ArchiveReader> PluginHandler<T> {
    /// Get the inner implementation.
    pub fn inner(&self) -> &T {
        &self.inner
    }

    /// Get the inner implementation mutably.
    pub fn inner_mut(&mut self) -> &mut T {
        &mut self.inner
    }
}

// =============================================================================
// IUnknown implementation
// =============================================================================

unsafe extern "system" fn query_interface<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    riid: *const GUID,
    ppv_object: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        if ppv_object.is_null() {
            return E_POINTER;
        }

        let riid = &*riid;
        if *riid == IID_IUNKNOWN || *riid == IID_IINARCHIVE {
            *ppv_object = this as *mut c_void;
            add_ref(this);
            return S_OK;
        }

        // Return IOutArchive interface if supported.
        // COM requires returning a pointer to the location containing the vtable pointer.
        // Since PluginHandler::out_vtbl is a `*const IOutArchiveVtbl`, we return the address
        // of that field (`&handler.out_vtbl`). The IOutArchive wrapper functions then use
        // `out_vtbl_to_handler()` to recover the base handler pointer.
        if *riid == IID_IOUTARCHIVE && T::supports_write() {
            let handler = &*this;
            *ppv_object = &handler.out_vtbl as *const _ as *mut c_void;
            add_ref(this);
            return S_OK;
        }

        *ppv_object = std::ptr::null_mut();
        E_NOINTERFACE
    }
}

unsafe extern "system" fn add_ref<T: ArchiveReader>(this: *mut PluginHandler<T>) -> u32 {
    unsafe {
        let handler = &*this;
        handler.ref_count.fetch_add(1, Ordering::SeqCst) + 1
    }
}

unsafe extern "system" fn release<T: ArchiveReader>(this: *mut PluginHandler<T>) -> u32 {
    unsafe {
        // Avoid creating a reference before potential deallocation.
        // Access the atomic directly through the raw pointer.
        let count = (*this).ref_count.fetch_sub(1, Ordering::SeqCst) - 1;
        if count == 0 {
            // Release the input stream before destroying the handler
            let handler = &mut *this;
            if !handler.in_stream.is_null() {
                IInStream::<c_void>::from_ptr_mut(handler.in_stream).release();
                handler.in_stream = std::ptr::null_mut();
            }
            handler.inner.close();
            drop(Box::from_raw(this));
        }
        count
    }
}

// =============================================================================
// IInArchive implementation
// =============================================================================

/// Wrapper around 7-Zip's ICryptoGetTextPassword interface.
///
/// This allows plugin authors to request passwords from the user for encrypted archives.
struct PasswordRequesterWrapper {
    crypto_callback: *mut c_void,
}

impl PasswordRequesterWrapper {
    /// Try to create a password requester from an open callback.
    ///
    /// Returns `None` if the callback doesn't support password requests.
    unsafe fn try_from_callback(open_callback: *mut c_void) -> Option<Self> {
        if open_callback.is_null() {
            return None;
        }

        unsafe {
            // Use the wrapper to QueryInterface for ICryptoGetTextPassword
            let callback = IArchiveExtractCallback::<c_void>::from_ptr_mut(open_callback);
            let mut crypto_ptr: *mut c_void = std::ptr::null_mut();
            let hr = callback.query_interface(&IID_ICRYPTOGETTEXTPASSWORD, &mut crypto_ptr);

            if hr.is_ok() && !crypto_ptr.is_null() {
                Some(Self {
                    crypto_callback: crypto_ptr,
                })
            } else {
                None
            }
        }
    }
}

impl PasswordRequester for PasswordRequesterWrapper {
    fn get_password(&self) -> crate::error::Result<Option<String>> {
        unsafe {
            let crypto = ICryptoGetTextPassword::<c_void>::from_ptr_mut(self.crypto_callback);

            let mut password_ptr: *mut u16 = std::ptr::null_mut();
            let hr = crypto.crypto_get_text_password(&mut password_ptr);

            if hr.is_err() {
                return Ok(None); // User cancelled or error
            }

            if password_ptr.is_null() {
                return Ok(Some(String::new()));
            }

            // Convert BSTR (which is *mut u16) to String
            // BSTR layout: length prefix at ptr-2, null-terminated UTF-16 string
            let bstr = BSTR::from_raw(password_ptr);
            let password = bstr.to_string();
            // Don't drop the BSTR - 7-Zip owns it
            std::mem::forget(bstr);

            Ok(Some(password))
        }
    }
}

impl Drop for PasswordRequesterWrapper {
    fn drop(&mut self) {
        unsafe {
            if !self.crypto_callback.is_null() {
                ICryptoGetTextPassword::<c_void>::from_ptr_mut(self.crypto_callback).release();
            }
        }
    }
}

/// Wrapper around 7-Zip's ICryptoGetTextPassword2 interface.
///
/// This allows plugin authors to get the password for creating encrypted archives.
struct PasswordProviderWrapper {
    crypto_callback: *mut c_void,
}

impl PasswordProviderWrapper {
    /// Try to create a password provider from an update callback.
    ///
    /// Returns `None` if the callback doesn't support password setting.
    unsafe fn try_from_callback(update_callback: *mut c_void) -> Option<Self> {
        if update_callback.is_null() {
            return None;
        }

        unsafe {
            // Use wrapper to QueryInterface for ICryptoGetTextPassword2
            let callback = IArchiveUpdateCallback::<c_void>::from_ptr_mut(update_callback);
            let mut crypto_ptr: *mut c_void = std::ptr::null_mut();
            let hr = callback.query_interface(&IID_ICRYPTOGETTEXTPASSWORD2, &mut crypto_ptr);

            if hr.is_ok() && !crypto_ptr.is_null() {
                Some(Self {
                    crypto_callback: crypto_ptr,
                })
            } else {
                None
            }
        }
    }
}

impl PasswordProvider for PasswordProviderWrapper {
    fn get_password(&self) -> crate::error::Result<Option<String>> {
        unsafe {
            let crypto = ICryptoGetTextPassword2::<c_void>::from_ptr_mut(self.crypto_callback);

            let mut password_is_defined: i32 = 0;
            let mut password_ptr: *mut u16 = std::ptr::null_mut();
            let hr = crypto.crypto_get_text_password2(&mut password_is_defined, &mut password_ptr);

            if hr.is_err() {
                return Ok(None); // Error getting password
            }

            // If password is not defined, user doesn't want encryption
            if password_is_defined == 0 {
                return Ok(None);
            }

            if password_ptr.is_null() {
                return Ok(Some(String::new()));
            }

            // Convert BSTR (which is *mut u16) to String
            let bstr = BSTR::from_raw(password_ptr);
            let password = bstr.to_string();
            // Don't drop the BSTR - 7-Zip owns it
            std::mem::forget(bstr);

            Ok(Some(password))
        }
    }
}

impl Drop for PasswordProviderWrapper {
    fn drop(&mut self) {
        unsafe {
            if !self.crypto_callback.is_null() {
                ICryptoGetTextPassword2::<c_void>::from_ptr_mut(self.crypto_callback).release();
            }
        }
    }
}

unsafe extern "system" fn open<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    stream: *mut c_void,
    _max_check_start_position: *const u64,
    open_callback: *mut c_void,
) -> HRESULT {
    unsafe {
        let handler = &mut *this;

        if stream.is_null() {
            return E_POINTER;
        }

        // Release any existing stream before opening a new one
        // This can happen if 7-Zip reopens the archive after an update
        if !handler.in_stream.is_null() {
            IInStream::<c_void>::from_ptr_mut(handler.in_stream).release();
            handler.in_stream = std::ptr::null_mut();
        }

        // Create streaming reader wrapper
        let mut reader = match InStreamReader::new(stream) {
            Ok(r) => r,
            Err(_e) => {
                #[cfg(debug_assertions)]
                eprintln!("[sevenzip-plugin] Failed to create stream reader: {}", _e);
                return S_FALSE;
            }
        };

        let size = reader.size();

        // Try to get password requester from open callback
        let password_requester = PasswordRequesterWrapper::try_from_callback(open_callback);

        // Call the safe streaming open method with password support
        let open_result = if password_requester.is_some() {
            handler.inner.open_with_password(
                &mut reader,
                size,
                password_requester
                    .as_ref()
                    .map(|p| p as &dyn PasswordRequester),
            )
        } else {
            handler.inner.open_with_password(&mut reader, size, None)
        };

        if let Err(_e) = open_result {
            #[cfg(debug_assertions)]
            eprintln!("[sevenzip-plugin] Failed to open archive: {}", _e);
            return S_FALSE;
        }

        // AddRef the stream to keep it alive while we have it
        IInStream::<c_void>::from_ptr_mut(stream).add_ref();

        handler.in_stream = stream;
        handler.archive_size = size;
        handler.is_open = true;
        S_OK
    }
}

unsafe extern "system" fn close<T: ArchiveReader>(this: *mut PluginHandler<T>) -> HRESULT {
    unsafe {
        let handler = &mut *this;
        handler.inner.close();

        // Release the input stream if we have one
        if !handler.in_stream.is_null() {
            IInStream::<c_void>::from_ptr_mut(handler.in_stream).release();
            handler.in_stream = std::ptr::null_mut();
        }

        handler.is_open = false;
        handler.archive_size = 0;
        S_OK
    }
}

unsafe extern "system" fn get_number_of_items<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    num_items: *mut u32,
) -> HRESULT {
    unsafe {
        if num_items.is_null() {
            return E_INVALIDARG;
        }
        let handler = &*this;
        *num_items = handler.inner.item_count() as u32;
        S_OK
    }
}

unsafe extern "system" fn get_property<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    index: u32,
    prop_id: u32,
    value: *mut c_void,
) -> HRESULT {
    unsafe {
        if value.is_null() {
            return E_INVALIDARG;
        }

        let handler = &*this;
        let Some(item) = handler.inner.get_item(index as usize) else {
            return E_INVALIDARG;
        };

        let prop = &mut *(value as *mut RawPropVariant);

        match prop_id {
            x if x == PropId::Path as u32 => {
                prop.set_bstr(&item.name);
            }
            x if x == PropId::Size as u32 => {
                prop.set_u64(item.size);
            }
            x if x == PropId::PackSize as u32 => {
                if let Some(packed) = item.compressed_size {
                    prop.set_u64(packed);
                } else {
                    prop.set_empty();
                }
            }
            x if x == PropId::IsDir as u32 => {
                prop.set_bool(item.is_dir);
            }
            x if x == PropId::MTime as u32 => {
                if let Some(mtime) = item.modified {
                    prop.set_filetime(super::propvariant::systemtime_to_filetime(mtime));
                } else {
                    prop.set_empty();
                }
            }
            x if x == PropId::CTime as u32 => {
                if let Some(ctime) = item.created {
                    prop.set_filetime(super::propvariant::systemtime_to_filetime(ctime));
                } else {
                    prop.set_empty();
                }
            }
            x if x == PropId::ATime as u32 => {
                if let Some(atime) = item.accessed {
                    prop.set_filetime(super::propvariant::systemtime_to_filetime(atime));
                } else {
                    prop.set_empty();
                }
            }
            x if x == PropId::Attrib as u32 => {
                if let Some(attrs) = item.attributes {
                    prop.set_u32(attrs);
                } else {
                    prop.set_empty();
                }
            }
            x if x == PropId::Crc as u32 => {
                if let Some(crc) = item.crc {
                    prop.set_u32(crc);
                } else {
                    prop.set_empty();
                }
            }
            x if x == PropId::Encrypted as u32 => {
                prop.set_bool(item.encrypted);
            }
            _ => {
                prop.set_empty();
            }
        }

        S_OK
    }
}

// Extract mode and result constants
const NASK_EXTRACT: i32 = 0;
const NRESULT_OK: i32 = 0;
const NRESULT_DATA_ERROR: i32 = 1;

/// Wrapper for ISequentialOutStream that implements `std::io::Write`.
///
/// This allows streaming writes directly to 7-Zip's output stream,
/// avoiding intermediate allocations.
struct SeqOutStreamWriter {
    stream: *mut c_void,
}

impl SeqOutStreamWriter {
    /// Create a new SeqOutStreamWriter from a raw ISequentialOutStream pointer.
    fn new(stream: *mut c_void) -> Self {
        Self { stream }
    }

    /// Get the underlying stream as a typed wrapper.
    #[inline]
    fn as_stream(&mut self) -> &mut ISequentialOutStream<c_void> {
        unsafe { ISequentialOutStream::<c_void>::from_ptr_mut(self.stream) }
    }
}

impl std::io::Write for SeqOutStreamWriter {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }

        let chunk_size = buf.len().min(u32::MAX as usize) as u32;
        let mut written: u32 = 0;

        // Safety: we're calling the COM method with valid pointers
        let hr = unsafe {
            self.as_stream()
                .write(buf.as_ptr(), chunk_size, &mut written)
        };

        if hr.is_err() {
            return Err(std::io::Error::other(format!(
                "Write failed with HRESULT: {:?}",
                hr
            )));
        }

        if written == 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::WriteZero,
                "Write returned zero bytes",
            ));
        }

        Ok(written as usize)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        // ISequentialOutStream has no flush method
        Ok(())
    }
}

unsafe extern "system" fn extract<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    indices: *const u32,
    num_items: u32,
    test_mode: i32,
    extract_callback: *mut c_void,
) -> HRESULT {
    unsafe {
        let handler = &mut *this;

        if extract_callback.is_null() {
            return E_INVALIDARG;
        }

        // Get callback wrapper for type-safe method calls
        let callback = IArchiveExtractCallback::<c_void>::from_ptr_mut(extract_callback);

        // Try to get password requester from extract callback
        // (for formats like ZIP where individual files can be encrypted)
        let password_requester = PasswordRequesterWrapper::try_from_callback(extract_callback);

        // Determine which indices to extract
        let extract_all = num_items == u32::MAX;
        let indices_to_extract: Vec<usize> = if extract_all {
            (0..handler.inner.item_count()).collect()
        } else {
            std::slice::from_raw_parts(indices, num_items as usize)
                .iter()
                .map(|&i| i as usize)
                .collect()
        };

        // Calculate total size
        let total_size: u64 = indices_to_extract
            .iter()
            .filter_map(|&i| handler.inner.get_item(i))
            .map(|item| item.size)
            .sum();

        let _ = callback.set_total(total_size);

        let mut completed: u64 = 0;

        for &index in &indices_to_extract {
            // Get item size before mutable borrow for extract()
            let item_size = match handler.inner.get_item(index) {
                Some(item) => item.size,
                None => continue,
            };

            // Get output stream
            let mut out_stream: *mut c_void = std::ptr::null_mut();
            let hr = callback.get_stream(index as u32, &mut out_stream, NASK_EXTRACT);
            if hr.is_err() {
                return hr;
            }

            // Prepare operation
            let _ = callback.prepare_operation(NASK_EXTRACT);

            // If test mode or no stream, skip extraction
            let result = if test_mode != 0 || out_stream.is_null() {
                NRESULT_OK
            } else {
                // Extract data using streaming trait method with password support
                let mut writer = SeqOutStreamWriter::new(out_stream);

                let extract_result = handler.inner.extract_to_with_password(
                    index,
                    &mut writer,
                    password_requester
                        .as_ref()
                        .map(|p| p as &dyn PasswordRequester),
                );

                match extract_result {
                    Ok(_) => NRESULT_OK,
                    Err(_) => NRESULT_DATA_ERROR,
                }
            };

            // Release output stream
            if !out_stream.is_null() {
                ISequentialOutStream::<c_void>::from_ptr_mut(out_stream).release();
            }

            // Set operation result
            let hr = callback.set_operation_result(result);
            if hr.is_err() {
                return hr;
            }

            // Update progress
            completed += item_size;
            let _ = callback.set_completed(&completed);
        }

        S_OK
    }
}

unsafe extern "system" fn get_archive_property<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    prop_id: u32,
    value: *mut c_void,
) -> HRESULT {
    unsafe {
        if value.is_null() {
            return E_INVALIDARG;
        }

        let handler = &*this;
        let prop = &mut *(value as *mut RawPropVariant);

        match prop_id {
            x if x == ArchivePropId::PhySize as u32 => {
                if let Some(size) = handler.inner.physical_size() {
                    prop.set_u64(size);
                } else {
                    prop.set_u64(handler.archive_size);
                }
            }
            _ => {
                prop.set_empty();
            }
        }

        S_OK
    }
}

unsafe extern "system" fn get_number_of_properties<T: ArchiveReader>(
    _this: *mut PluginHandler<T>,
    num_props: *mut u32,
) -> HRESULT {
    unsafe {
        if !num_props.is_null() {
            *num_props = 10; // Path, Size, PackSize, IsDir, MTime, CTime, ATime, Attrib, CRC, Encrypted
        }
        S_OK
    }
}

unsafe extern "system" fn get_property_info<T: ArchiveReader>(
    _this: *mut PluginHandler<T>,
    index: u32,
    name: *mut c_void,
    prop_id: *mut u32,
    var_type: *mut u32,
) -> HRESULT {
    unsafe {
        use super::propvariant::{VT_BOOL, VT_BSTR, VT_FILETIME, VT_UI4, VT_UI8};

        if name.is_null() || prop_id.is_null() || var_type.is_null() {
            return E_INVALIDARG;
        }

        *(name as *mut BSTR) = BSTR::default();

        match index {
            0 => {
                *prop_id = PropId::Path as u32;
                *var_type = VT_BSTR as u32;
            }
            1 => {
                *prop_id = PropId::Size as u32;
                *var_type = VT_UI8 as u32;
            }
            2 => {
                *prop_id = PropId::PackSize as u32;
                *var_type = VT_UI8 as u32;
            }
            3 => {
                *prop_id = PropId::IsDir as u32;
                *var_type = VT_BOOL as u32;
            }
            4 => {
                *prop_id = PropId::MTime as u32;
                *var_type = VT_FILETIME as u32;
            }
            5 => {
                *prop_id = PropId::CTime as u32;
                *var_type = VT_FILETIME as u32;
            }
            6 => {
                *prop_id = PropId::ATime as u32;
                *var_type = VT_FILETIME as u32;
            }
            7 => {
                *prop_id = PropId::Attrib as u32;
                *var_type = VT_UI4 as u32;
            }
            8 => {
                *prop_id = PropId::Crc as u32;
                *var_type = VT_UI4 as u32;
            }
            9 => {
                *prop_id = PropId::Encrypted as u32;
                *var_type = VT_BOOL as u32;
            }
            _ => return E_INVALIDARG,
        }

        S_OK
    }
}

unsafe extern "system" fn get_number_of_archive_properties<T: ArchiveReader>(
    _this: *mut PluginHandler<T>,
    num_props: *mut u32,
) -> HRESULT {
    unsafe {
        if !num_props.is_null() {
            *num_props = 1;
        }
        S_OK
    }
}

unsafe extern "system" fn get_archive_property_info<T: ArchiveReader>(
    _this: *mut PluginHandler<T>,
    index: u32,
    name: *mut c_void,
    prop_id: *mut u32,
    var_type: *mut u32,
) -> HRESULT {
    unsafe {
        use super::propvariant::VT_UI8;

        if name.is_null() || prop_id.is_null() || var_type.is_null() {
            return E_INVALIDARG;
        }

        *(name as *mut BSTR) = BSTR::default();

        match index {
            0 => {
                *prop_id = 4; // PhySize
                *var_type = VT_UI8 as u32;
            }
            _ => return E_INVALIDARG,
        }

        S_OK
    }
}

// =============================================================================
// IOutArchive implementation
// =============================================================================

const NFILETIME_WINDOWS: u32 = 0;

// IOutArchive wrapper functions that convert out_vtbl pointer to handler base
unsafe fn out_vtbl_to_handler<T: ArchiveReader>(
    out_vtbl_ptr: *mut PluginHandler<T>,
) -> *mut PluginHandler<T> {
    unsafe {
        let offset = std::mem::offset_of!(PluginHandler::<T>, out_vtbl);
        (out_vtbl_ptr as *mut u8).sub(offset) as *mut PluginHandler<T>
    }
}

unsafe extern "system" fn out_query_interface<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    riid: *const GUID,
    ppv_object: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        let base = out_vtbl_to_handler(this);
        query_interface(base, riid, ppv_object)
    }
}

unsafe extern "system" fn out_add_ref<T: ArchiveReader>(this: *mut PluginHandler<T>) -> u32 {
    unsafe {
        let base = out_vtbl_to_handler(this);
        add_ref(base)
    }
}

unsafe extern "system" fn out_release<T: ArchiveReader>(this: *mut PluginHandler<T>) -> u32 {
    unsafe {
        let base = out_vtbl_to_handler(this);
        release(base)
    }
}

unsafe extern "system" fn get_file_time_type<T: ArchiveReader>(
    _this: *mut PluginHandler<T>,
    time_type: *mut u32,
) -> HRESULT {
    unsafe {
        if time_type.is_null() {
            return E_POINTER;
        }
        *time_type = NFILETIME_WINDOWS;
        S_OK
    }
}

unsafe extern "system" fn update_items<T: ArchiveReader + ArchiveUpdater>(
    this: *mut PluginHandler<T>,
    out_stream: *mut c_void,
    num_items: u32,
    update_callback: *mut c_void,
) -> HRESULT {
    unsafe {
        let handler = &mut *out_vtbl_to_handler(this);

        if out_stream.is_null() || update_callback.is_null() {
            return E_INVALIDARG;
        }

        // Inner function that does the actual work - allows us to use ? for early returns
        // while ensuring cleanup always happens in the outer function
        let result = update_items_inner(handler, out_stream, num_items, update_callback);

        // ALWAYS clean up, regardless of success or failure
        handler.inner.close();
        handler.is_open = false;

        if !handler.in_stream.is_null() {
            IInStream::<c_void>::from_ptr_mut(handler.in_stream).release();
            handler.in_stream = std::ptr::null_mut();
        }

        result
    }
}

/// Inner implementation of update_items that can return early.
/// Cleanup is handled by the caller.
unsafe fn update_items_inner<T: ArchiveReader + ArchiveUpdater>(
    handler: &mut PluginHandler<T>,
    out_stream: *mut c_void,
    num_items: u32,
    update_callback: *mut c_void,
) -> HRESULT {
    unsafe {
        // Get callback wrapper for type-safe method calls
        let callback = IArchiveUpdateCallback::<c_void>::from_ptr_mut(update_callback);

        use crate::types::UpdateItem;
        let mut updates = Vec::new();
        let mut total_size: u64 = 0;

        // First pass: gather update info and calculate total size
        for i in 0..num_items {
            let mut new_data: i32 = 0;
            let mut new_props: i32 = 0;
            let mut index_in_archive: u32 = u32::MAX;

            let hr = callback.get_update_item_info(
                i,
                &mut new_data,
                &mut new_props,
                &mut index_in_archive,
            );
            if hr.is_err() {
                return hr;
            }

            if new_data != 0 {
                // New file - get size property for progress tracking
                let mut size_prop = RawPropVariant::default();
                let _ = callback.get_property(
                    i,
                    PropId::Size as u32,
                    &mut size_prop as *mut _ as *mut c_void,
                );
                let file_size = size_prop.get_u64().unwrap_or(0);
                total_size += file_size;
            } else if index_in_archive != u32::MAX {
                // Copying existing item - get its size
                if let Some(item) = handler.inner.get_item(index_in_archive as usize) {
                    total_size += item.size;
                }
            }
        }

        // Report total size to 7-Zip
        let _ = callback.set_total(total_size);

        // Second pass: collect data
        for i in 0..num_items {
            let mut new_data: i32 = 0;
            let mut new_props: i32 = 0;
            let mut index_in_archive: u32 = u32::MAX;

            let hr = callback.get_update_item_info(
                i,
                &mut new_data,
                &mut new_props,
                &mut index_in_archive,
            );
            if hr.is_err() {
                return hr;
            }

            if new_data != 0 {
                // Check if this is a directory - skip directories as most archive
                // formats don't need explicit directory entries (paths contain folders)
                let mut is_dir_prop = RawPropVariant::default();
                let _ = callback.get_property(
                    i,
                    PropId::IsDir as u32,
                    &mut is_dir_prop as *mut _ as *mut c_void,
                );
                let is_dir = is_dir_prop.get_bool().unwrap_or(false);

                if is_dir {
                    // Skip directories - report success and continue
                    let _ = callback.set_operation_result(NRESULT_OK);
                    continue;
                }

                // New file - get name and data
                let mut prop = RawPropVariant::default();
                let hr = callback.get_property(
                    i,
                    PropId::Path as u32,
                    &mut prop as *mut _ as *mut c_void,
                );
                if hr.is_err() {
                    return hr;
                }

                let name = prop.get_bstr().unwrap_or_default();

                // Get input stream
                let mut in_stream: *mut c_void = std::ptr::null_mut();
                let hr = callback.get_stream(i, &mut in_stream);
                if hr.is_err() {
                    return hr;
                }

                let data = if !in_stream.is_null() {
                    let result = read_sequential_stream(in_stream);
                    // Release stream
                    ISequentialInStream::<c_void>::from_ptr_mut(in_stream).release();
                    result.unwrap_or_default()
                } else {
                    Vec::new()
                };

                // Don't report progress here - the plugin will report progress
                // during update_streaming when the data is actually written.

                updates.push(UpdateItem::AddNew { name, data });

                // Report operation result for this item
                let _ = callback.set_operation_result(NRESULT_OK);
            } else if index_in_archive != u32::MAX {
                // Copy existing item - don't report progress here since no actual work
                // is done during collection. Progress will be reported by the plugin
                // during update_streaming when the data is actually processed.
                updates.push(UpdateItem::CopyExisting {
                    index: index_in_archive as usize,
                    new_name: None,
                });

                // Report operation result for this item
                let _ = callback.set_operation_result(NRESULT_OK);
            }
            // else: item is being deleted (new_data == 0 && index_in_archive == MAX)
            // - don't add to updates list, which removes it from the archive
        }

        // Create streaming writer for output
        let mut writer = SeqOutStreamWriter::new(out_stream);

        // Try to get password provider from update callback
        // (for creating encrypted archives)
        let password_provider = PasswordProviderWrapper::try_from_callback(update_callback);

        // Create progress callback that reports to 7-Zip
        // The inner format reports (completed, total) in whatever units make sense for it.
        // We scale this to match total_size (what we told 7-Zip) using the ratio.
        let mut progress_fn = |write_completed: u64, write_total: u64| -> bool {
            let scaled = if write_total > 0 {
                // Scale: (completed / total) * total_size
                let ratio = write_completed as f64 / write_total as f64;
                (ratio * total_size as f64).min(total_size as f64) as u64
            } else {
                0
            };
            let _ = callback.set_completed(&scaled);
            true // continue operation
        };

        // Create reader for existing archive (if we have one)
        if handler.in_stream.is_null() {
            // No existing archive - create empty reader
            let mut empty_reader = std::io::Cursor::new(&[] as &[u8]);
            let result = handler.inner.update_streaming_with_password(
                &mut empty_reader,
                0,
                updates,
                &mut writer,
                Some(&mut progress_fn),
                password_provider
                    .as_ref()
                    .map(|p| p as &dyn PasswordProvider),
            );
            match result {
                Ok(_) => {
                    // Report 100% completion after write phase finishes
                    // Re-get the callback wrapper since progress_fn borrowed it
                    let cb = IArchiveUpdateCallback::<c_void>::from_ptr_mut(update_callback);
                    let _ = cb.set_completed(&total_size);
                    S_OK
                }
                Err(_) => S_FALSE,
            }
        } else {
            // Use existing archive stream
            let mut reader = match InStreamReader::new(handler.in_stream) {
                Ok(r) => r,
                Err(_) => return S_FALSE,
            };
            let size = reader.size();

            let result = handler.inner.update_streaming_with_password(
                &mut reader,
                size,
                updates,
                &mut writer,
                Some(&mut progress_fn),
                password_provider
                    .as_ref()
                    .map(|p| p as &dyn PasswordProvider),
            );
            match result {
                Ok(_) => {
                    // Report 100% completion after write phase finishes
                    // Re-get the callback wrapper since progress_fn borrowed it
                    let cb = IArchiveUpdateCallback::<c_void>::from_ptr_mut(update_callback);
                    let _ = cb.set_completed(&total_size);
                    S_OK
                }
                Err(_) => S_FALSE,
            }
        }
    }
}

// Stub for non-updatable formats
unsafe extern "system" fn update_items_stub<T: ArchiveReader>(
    _this: *mut PluginHandler<T>,
    _out_stream: *mut c_void,
    _num_items: u32,
    _update_callback: *mut c_void,
) -> HRESULT {
    E_NOTIMPL
}

// =============================================================================
// Vtable and Handler Creation
// =============================================================================

/// Creates the static IInArchive vtable for a format type.
///
/// This is used internally by the registration macro.
pub const fn create_in_vtable<T: ArchiveReader>() -> IInArchiveVTable<PluginHandler<T>> {
    IInArchiveVTable {
        base: IUnknownVTable {
            query_interface: query_interface::<T>,
            add_ref: add_ref::<T>,
            release: release::<T>,
        },
        open: open::<T>,
        close: close::<T>,
        get_number_of_items: get_number_of_items::<T>,
        get_property: get_property::<T>,
        extract: extract::<T>,
        get_archive_property: get_archive_property::<T>,
        get_number_of_properties: get_number_of_properties::<T>,
        get_property_info: get_property_info::<T>,
        get_number_of_archive_properties: get_number_of_archive_properties::<T>,
        get_archive_property_info: get_archive_property_info::<T>,
    }
}

/// Creates the static IOutArchive vtable for a format type (stub version).
pub const fn create_out_vtable_stub<T: ArchiveReader>() -> IOutArchiveVTable<PluginHandler<T>> {
    IOutArchiveVTable {
        base: IUnknownVTable {
            query_interface: out_query_interface::<T>,
            add_ref: out_add_ref::<T>,
            release: out_release::<T>,
        },
        update_items: update_items_stub::<T>,
        get_file_time_type: get_file_time_type::<T>,
    }
}

/// Creates the static IOutArchive vtable for a format type that supports updates.
pub const fn create_out_vtable<T: ArchiveReader + ArchiveUpdater>()
-> IOutArchiveVTable<PluginHandler<T>> {
    IOutArchiveVTable {
        base: IUnknownVTable {
            query_interface: out_query_interface::<T>,
            add_ref: out_add_ref::<T>,
            release: out_release::<T>,
        },
        update_items: update_items::<T>,
        get_file_time_type: get_file_time_type::<T>,
    }
}

/// A registered format with vtables and metadata.
pub struct RegisteredFormat<T: ArchiveReader> {
    pub in_vtbl: &'static IInArchiveVTable<PluginHandler<T>>,
    pub out_vtbl: &'static IOutArchiveVTable<PluginHandler<T>>,
    _phantom: PhantomData<T>,
}

// SAFETY: RegisteredFormat only contains references to static vtables which are
// immutable and have 'static lifetime. The PhantomData<T> is just a marker.
// Since T: ArchiveReader requires T: Send + 'static, and the vtables are const,
// it is safe to share RegisteredFormat across threads.
unsafe impl<T: ArchiveReader> Send for RegisteredFormat<T> {}
unsafe impl<T: ArchiveReader> Sync for RegisteredFormat<T> {}

impl<T: ArchiveReader> RegisteredFormat<T> {
    /// Create a new registered format with the given vtables.
    pub const fn new(
        in_vtbl: &'static IInArchiveVTable<PluginHandler<T>>,
        out_vtbl: &'static IOutArchiveVTable<PluginHandler<T>>,
    ) -> Self {
        Self {
            in_vtbl,
            out_vtbl,
            _phantom: PhantomData,
        }
    }

    /// Create a new handler instance.
    pub fn create_handler(&self) -> *mut c_void {
        let handler = Box::new(PluginHandler {
            in_vtbl: self.in_vtbl,
            out_vtbl: self.out_vtbl,
            ref_count: AtomicU32::new(1),
            inner: T::default(),
            in_stream: std::ptr::null_mut(),
            archive_size: 0,
            is_open: false,
        });
        Box::into_raw(handler) as *mut c_void
    }
}
