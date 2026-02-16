//! Core types for archive items and properties.

use std::time::SystemTime;

/// Information about a single item (file/directory) in an archive.
#[derive(Debug, Clone, Default)]
pub struct ArchiveItem {
    /// File/directory name (path within archive)
    pub name: String,
    /// Uncompressed size in bytes
    pub size: u64,
    /// Compressed size in bytes (if applicable)
    pub compressed_size: Option<u64>,
    /// Last modification time
    pub modified: Option<SystemTime>,
    /// Creation time
    pub created: Option<SystemTime>,
    /// Last access time
    pub accessed: Option<SystemTime>,
    /// Whether this is a directory
    pub is_dir: bool,
    /// File attributes (Windows-style, optional)
    pub attributes: Option<u32>,
    /// CRC32 checksum (optional)
    pub crc: Option<u32>,
    /// Whether this item is encrypted (shows lock icon in 7-Zip)
    pub encrypted: bool,
}

impl ArchiveItem {
    /// Create a new file item with just name and size.
    pub fn file(name: impl Into<String>, size: u64) -> Self {
        Self {
            name: name.into(),
            size,
            is_dir: false,
            ..Default::default()
        }
    }

    /// Create a new directory item.
    pub fn directory(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            size: 0,
            is_dir: true,
            ..Default::default()
        }
    }

    /// Set the compressed size.
    pub fn with_compressed_size(mut self, size: u64) -> Self {
        self.compressed_size = Some(size);
        self
    }

    /// Set the modification time.
    pub fn with_modified(mut self, time: SystemTime) -> Self {
        self.modified = Some(time);
        self
    }

    /// Set the creation time.
    pub fn with_created(mut self, time: SystemTime) -> Self {
        self.created = Some(time);
        self
    }

    /// Set the last access time.
    pub fn with_accessed(mut self, time: SystemTime) -> Self {
        self.accessed = Some(time);
        self
    }

    /// Set the file attributes (Windows-style).
    pub fn with_attributes(mut self, attrs: u32) -> Self {
        self.attributes = Some(attrs);
        self
    }

    /// Set the CRC32 checksum.
    pub fn with_crc(mut self, crc: u32) -> Self {
        self.crc = Some(crc);
        self
    }

    /// Mark this item as encrypted (shows lock icon in 7-Zip).
    pub fn with_encrypted(mut self, encrypted: bool) -> Self {
        self.encrypted = encrypted;
        self
    }
}

/// Progress callback for archive operations.
///
/// Called during long-running operations (like `update_streaming`) to report progress.
/// - `completed`: Number of bytes processed so far
/// - `total`: Total number of bytes to process
///
/// Return `true` to continue the operation, or `false` to request cancellation.
pub type ProgressCallback<'a> = &'a mut dyn FnMut(u64, u64) -> bool;

/// A trait for requesting passwords from 7-Zip's UI.
///
/// This is passed to archive open/extract methods when the user may need
/// to provide a password for encrypted archives.
pub trait PasswordRequester {
    /// Request a password from the user.
    ///
    /// Returns:
    /// - `Ok(Some(password))` - User provided a password
    /// - `Ok(None)` - No password available (user cancelled or not supported)
    /// - `Err(_)` - Error occurred while getting password
    fn get_password(&self) -> crate::error::Result<Option<String>>;
}

/// A trait for getting the password when creating encrypted archives.
///
/// This is passed to archive update methods when the user has requested
/// encryption for the new archive.
pub trait PasswordProvider {
    /// Get the password for encrypting the new archive.
    ///
    /// Returns:
    /// - `Ok(Some(password))` - User wants encryption with this password
    /// - `Ok(None)` - No encryption requested
    /// - `Err(_)` - Error occurred while getting password
    fn get_password(&self) -> crate::error::Result<Option<String>>;
}

/// Describes an update operation for archive editing.
#[derive(Debug, Clone)]
pub enum UpdateItem {
    /// Copy an existing item from the source archive by index.
    CopyExisting {
        /// Index of the item in the original archive
        index: usize,
        /// New name (if renaming), or None to keep original
        new_name: Option<String>,
    },
    /// Add new data to the archive.
    AddNew {
        /// Name/path for the new item
        name: String,
        /// The data to add
        data: Vec<u8>,
    },
}
