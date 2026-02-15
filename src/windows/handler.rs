//! Generic COM handler wrapper that bridges safe traits to 7-Zip interfaces.

use std::ffi::c_void;
use std::marker::PhantomData;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::{
    E_INVALIDARG, E_NOINTERFACE, E_NOTIMPL, E_POINTER, S_FALSE, S_OK,
};
use windows::core::{BSTR, GUID, HRESULT};

use crate::traits::{ArchiveReader, ArchiveUpdater};

use super::com::{
    ArchivePropId, IID_IINARCHIVE, IID_IOUTARCHIVE, IID_IUNKNOWN, IInArchiveVtbl, IOutArchiveVtbl,
    PropId,
};
use super::propvariant::RawPropVariant;

// Stream seek origins
const STREAM_SEEK_SET: u32 = 0;
const STREAM_SEEK_END: u32 = 2;

/// Read all data from a 7-Zip IInStream into a Vec<u8>.
pub(crate) unsafe fn read_stream_to_vec(stream: *mut c_void) -> std::io::Result<Vec<u8>> {
    unsafe {
        if stream.is_null() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "Null stream pointer",
            ));
        }

        // IInStream vtable layout:
        // 0: QueryInterface, 1: AddRef, 2: Release (IUnknown)
        // 3: Read (ISequentialInStream)
        // 4: Seek (IInStream)
        type SeekFn = unsafe extern "system" fn(*mut c_void, i64, u32, *mut u64) -> HRESULT;
        type ReadFn = unsafe extern "system" fn(*mut c_void, *mut u8, u32, *mut u32) -> HRESULT;

        let vtable = *(stream as *const *const *const c_void);
        let seek_fn: SeekFn = std::mem::transmute(*vtable.add(4));
        let read_fn: ReadFn = std::mem::transmute(*vtable.add(3));

        // Seek to end to get size
        let mut size: u64 = 0;
        let hr = seek_fn(stream, 0, STREAM_SEEK_END, &mut size);
        if hr.is_err() {
            return Err(std::io::Error::other(format!(
                "Failed to seek to end: {:?}",
                hr
            )));
        }

        // Seek back to start
        let mut pos: u64 = 0;
        let hr = seek_fn(stream, 0, STREAM_SEEK_SET, &mut pos);
        if hr.is_err() {
            return Err(std::io::Error::other(format!(
                "Failed to seek to start: {:?}",
                hr
            )));
        }

        // Read all data
        let mut data = vec![0u8; size as usize];
        let mut total_read: usize = 0;

        while total_read < size as usize {
            let mut bytes_read: u32 = 0;
            let hr = read_fn(
                stream,
                data[total_read..].as_mut_ptr(),
                (size as usize - total_read).min(1024 * 1024) as u32,
                &mut bytes_read,
            );
            if hr.is_err() {
                return Err(std::io::Error::other(format!("Failed to read: {:?}", hr)));
            }
            if bytes_read == 0 {
                break;
            }
            total_read += bytes_read as usize;
        }

        data.truncate(total_read);
        Ok(data)
    }
}

/// Read data from ISequentialInStream
pub(crate) unsafe fn read_sequential_stream(stream: *mut c_void) -> std::io::Result<Vec<u8>> {
    unsafe {
        if stream.is_null() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "Null stream pointer",
            ));
        }

        type ReadFn = unsafe extern "system" fn(*mut c_void, *mut u8, u32, *mut u32) -> HRESULT;

        let vtable = *(stream as *const *const *const c_void);
        let read_fn: ReadFn = std::mem::transmute(*vtable.add(3));

        let mut data = Vec::new();
        let mut buffer = [0u8; 65536];

        loop {
            let mut bytes_read: u32 = 0;
            let hr = read_fn(
                stream,
                buffer.as_mut_ptr(),
                buffer.len() as u32,
                &mut bytes_read,
            );
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

/// Write data to ISequentialOutStream
pub(crate) unsafe fn write_to_stream(stream: *mut c_void, data: &[u8]) -> std::io::Result<()> {
    unsafe {
        if stream.is_null() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "Null stream pointer",
            ));
        }

        type WriteFn = unsafe extern "system" fn(*mut c_void, *const u8, u32, *mut u32) -> HRESULT;

        let vtable = *(stream as *const *const *const c_void);
        let write_fn: WriteFn = std::mem::transmute(*vtable.add(3));

        let mut total_written = 0;
        while total_written < data.len() {
            let mut written: u32 = 0;
            let chunk_size = (data.len() - total_written).min(1024 * 1024) as u32;
            let hr = write_fn(
                stream,
                data[total_written..].as_ptr(),
                chunk_size,
                &mut written,
            );
            if hr.is_err() {
                return Err(std::io::Error::other(format!("Write failed: {:?}", hr)));
            }
            if written == 0 {
                break;
            }
            total_written += written as usize;
        }

        Ok(())
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
    pub in_vtbl: *const IInArchiveVtbl<Self>,
    /// Pointer to IOutArchive vtable - for writing support
    pub out_vtbl: *const IOutArchiveVtbl<Self>,
    /// Reference count
    ref_count: AtomicU32,
    /// The actual archive implementation (safe Rust)
    pub(crate) inner: T,
    /// Raw archive data (needed for editing)
    pub(crate) archive_data: Option<Vec<u8>>,
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
            drop(Box::from_raw(this));
        }
        count
    }
}

// =============================================================================
// IInArchive implementation
// =============================================================================

unsafe extern "system" fn open<T: ArchiveReader>(
    this: *mut PluginHandler<T>,
    stream: *mut c_void,
    _max_check_start_position: *const u64,
    _open_callback: *mut c_void,
) -> HRESULT {
    unsafe {
        let handler = &mut *this;

        // Read stream data
        let data = match read_stream_to_vec(stream) {
            Ok(d) => d,
            Err(_e) => {
                #[cfg(debug_assertions)]
                eprintln!("[sevenzip-plugin] Failed to read stream: {}", _e);
                return S_FALSE;
            }
        };

        handler.archive_size = data.len() as u64;

        // Call the safe open method (before moving data)
        if let Err(_e) = handler.inner.open(&data) {
            #[cfg(debug_assertions)]
            eprintln!("[sevenzip-plugin] Failed to open archive: {}", _e);
            return S_FALSE;
        }

        // Store archive data by moving, not cloning
        handler.archive_data = Some(data);
        handler.is_open = true;
        S_OK
    }
}

unsafe extern "system" fn close<T: ArchiveReader>(this: *mut PluginHandler<T>) -> HRESULT {
    unsafe {
        let handler = &mut *this;
        handler.inner.close();
        handler.archive_data = None;
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

// Callback function types
type SetTotalFn = unsafe extern "system" fn(*mut c_void, u64) -> HRESULT;
type SetCompletedFn = unsafe extern "system" fn(*mut c_void, *const u64) -> HRESULT;
type GetStreamFn = unsafe extern "system" fn(*mut c_void, u32, *mut *mut c_void, i32) -> HRESULT;
type PrepareOperationFn = unsafe extern "system" fn(*mut c_void, i32) -> HRESULT;
type SetOperationResultFn = unsafe extern "system" fn(*mut c_void, i32) -> HRESULT;
type WriteFn = unsafe extern "system" fn(*mut c_void, *const u8, u32, *mut u32) -> HRESULT;
type ReleaseFn = unsafe extern "system" fn(*mut c_void) -> u32;

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

        // Get callback vtable functions
        let cb_vtable = *(extract_callback as *const *const *const c_void);
        let set_total: SetTotalFn = std::mem::transmute(*cb_vtable.add(3));
        let set_completed: SetCompletedFn = std::mem::transmute(*cb_vtable.add(4));
        let get_stream: GetStreamFn = std::mem::transmute(*cb_vtable.add(5));
        let prepare_operation: PrepareOperationFn = std::mem::transmute(*cb_vtable.add(6));
        let set_operation_result: SetOperationResultFn = std::mem::transmute(*cb_vtable.add(7));

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

        let _ = set_total(extract_callback, total_size);

        let mut completed: u64 = 0;

        for &index in &indices_to_extract {
            // Get item size before mutable borrow for extract()
            let item_size = match handler.inner.get_item(index) {
                Some(item) => item.size,
                None => continue,
            };

            // Get output stream
            let mut out_stream: *mut c_void = std::ptr::null_mut();
            let hr = get_stream(
                extract_callback,
                index as u32,
                &mut out_stream,
                NASK_EXTRACT,
            );
            if hr.is_err() {
                return hr;
            }

            // Prepare operation
            let _ = prepare_operation(extract_callback, NASK_EXTRACT);

            // If test mode or no stream, skip extraction
            let result = if test_mode != 0 || out_stream.is_null() {
                NRESULT_OK
            } else {
                // Extract data using safe trait method
                match handler.inner.extract(index) {
                    Ok(data) => {
                        // Write to output stream
                        let stream_vtable = *(out_stream as *const *const *const c_void);
                        let write: WriteFn = std::mem::transmute(*stream_vtable.add(3));

                        let mut total_written = 0;
                        while total_written < data.len() {
                            let mut written: u32 = 0;
                            let chunk = (data.len() - total_written).min(1024 * 1024) as u32;
                            let hr = write(
                                out_stream,
                                data[total_written..].as_ptr(),
                                chunk,
                                &mut written,
                            );
                            if hr.is_err() || written == 0 {
                                break;
                            }
                            total_written += written as usize;
                        }
                        NRESULT_OK
                    }
                    Err(_) => NRESULT_DATA_ERROR,
                }
            };

            // Release output stream
            if !out_stream.is_null() {
                let stream_vtable = *(out_stream as *const *const *const c_void);
                let release_fn: ReleaseFn = std::mem::transmute(*stream_vtable.add(2));
                release_fn(out_stream);
            }

            // Set operation result
            let hr = set_operation_result(extract_callback, result);
            if hr.is_err() {
                return hr;
            }

            // Update progress
            completed += item_size;
            let _ = set_completed(extract_callback, &completed);
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
            *num_props = 4; // Path, Size, PackSize, IsDir
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
        use super::propvariant::{VT_BOOL, VT_BSTR, VT_UI8};

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

// Update callback function types
type GetUpdateItemInfoFn =
    unsafe extern "system" fn(*mut c_void, u32, *mut i32, *mut i32, *mut u32) -> HRESULT;
type GetPropertyFn = unsafe extern "system" fn(*mut c_void, u32, u32, *mut c_void) -> HRESULT;
type GetStreamFnUpdate = unsafe extern "system" fn(*mut c_void, u32, *mut *mut c_void) -> HRESULT;
type SetOperationResultUpdateFn = unsafe extern "system" fn(*mut c_void, i32) -> HRESULT;

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

        // Get callback vtable functions
        // IArchiveUpdateCallback vtable layout:
        // 0: QueryInterface, 1: AddRef, 2: Release (IUnknown)
        // 3: SetTotal, 4: SetCompleted (IProgress)
        // 5: GetUpdateItemInfo, 6: GetProperty, 7: GetStream, 8: SetOperationResult
        let cb_vtable = *(update_callback as *const *const *const c_void);
        let set_total: SetTotalFn = std::mem::transmute(*cb_vtable.add(3));
        let set_completed: SetCompletedFn = std::mem::transmute(*cb_vtable.add(4));
        let get_update_item_info: GetUpdateItemInfoFn = std::mem::transmute(*cb_vtable.add(5));
        let get_property: GetPropertyFn = std::mem::transmute(*cb_vtable.add(6));
        let get_stream: GetStreamFnUpdate = std::mem::transmute(*cb_vtable.add(7));
        let set_operation_result: SetOperationResultUpdateFn =
            std::mem::transmute(*cb_vtable.add(8));

        use crate::types::UpdateItem;
        let mut updates = Vec::new();
        let mut total_size: u64 = 0;

        // First pass: gather update info and calculate total size
        for i in 0..num_items {
            let mut new_data: i32 = 0;
            let mut new_props: i32 = 0;
            let mut index_in_archive: u32 = u32::MAX;

            let hr = get_update_item_info(
                update_callback,
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
                let _ = get_property(
                    update_callback,
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
        let _ = set_total(update_callback, total_size);

        let mut completed: u64 = 0;

        // Second pass: collect data with progress updates
        for i in 0..num_items {
            let mut new_data: i32 = 0;
            let mut new_props: i32 = 0;
            let mut index_in_archive: u32 = u32::MAX;

            let hr = get_update_item_info(
                update_callback,
                i,
                &mut new_data,
                &mut new_props,
                &mut index_in_archive,
            );
            if hr.is_err() {
                return hr;
            }

            if new_data != 0 {
                // New file - get name and data
                let mut prop = RawPropVariant::default();
                let hr = get_property(
                    update_callback,
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
                let hr = get_stream(update_callback, i, &mut in_stream);
                if hr.is_err() {
                    return hr;
                }

                let data = if !in_stream.is_null() {
                    let result = read_sequential_stream(in_stream);
                    // Release stream
                    let stream_vtable = *(in_stream as *const *const *const c_void);
                    let release_fn: ReleaseFn = std::mem::transmute(*stream_vtable.add(2));
                    release_fn(in_stream);
                    result.unwrap_or_default()
                } else {
                    Vec::new()
                };

                // Update progress
                completed += data.len() as u64;
                let _ = set_completed(update_callback, &completed);

                updates.push(UpdateItem::AddNew { name, data });

                // Report operation result for this item
                let _ = set_operation_result(update_callback, NRESULT_OK);
            } else {
                // Copy existing item - update progress if we have item info
                if index_in_archive != u32::MAX
                    && let Some(item) = handler.inner.get_item(index_in_archive as usize)
                {
                    completed += item.size;
                    let _ = set_completed(update_callback, &completed);
                }

                updates.push(UpdateItem::CopyExisting {
                    index: index_in_archive as usize,
                    new_name: None,
                });

                // Report operation result for this item
                let _ = set_operation_result(update_callback, NRESULT_OK);
            }
        }

        // Get existing archive data
        let existing_data = handler.archive_data.as_deref().unwrap_or(&[]);

        // Call the safe update method
        match handler.inner.update(existing_data, updates) {
            Ok(output_data) => {
                // Write to output stream
                if write_to_stream(out_stream, &output_data).is_err() {
                    return S_FALSE;
                }
                S_OK
            }
            Err(_) => S_FALSE,
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
pub const fn create_in_vtable<T: ArchiveReader>() -> IInArchiveVtbl<PluginHandler<T>> {
    IInArchiveVtbl {
        query_interface: query_interface::<T>,
        add_ref: add_ref::<T>,
        release: release::<T>,
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
pub const fn create_out_vtable_stub<T: ArchiveReader>() -> IOutArchiveVtbl<PluginHandler<T>> {
    IOutArchiveVtbl {
        query_interface: out_query_interface::<T>,
        add_ref: out_add_ref::<T>,
        release: out_release::<T>,
        update_items: update_items_stub::<T>,
        get_file_time_type: get_file_time_type::<T>,
    }
}

/// Creates the static IOutArchive vtable for a format type that supports updates.
pub const fn create_out_vtable<T: ArchiveReader + ArchiveUpdater>()
-> IOutArchiveVtbl<PluginHandler<T>> {
    IOutArchiveVtbl {
        query_interface: out_query_interface::<T>,
        add_ref: out_add_ref::<T>,
        release: out_release::<T>,
        update_items: update_items::<T>,
        get_file_time_type: get_file_time_type::<T>,
    }
}

/// A registered format with vtables and metadata.
pub struct RegisteredFormat<T: ArchiveReader> {
    pub in_vtbl: &'static IInArchiveVtbl<PluginHandler<T>>,
    pub out_vtbl: &'static IOutArchiveVtbl<PluginHandler<T>>,
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
        in_vtbl: &'static IInArchiveVtbl<PluginHandler<T>>,
        out_vtbl: &'static IOutArchiveVtbl<PluginHandler<T>>,
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
            archive_data: None,
            archive_size: 0,
            is_open: false,
        });
        Box::into_raw(handler) as *mut c_void
    }
}
