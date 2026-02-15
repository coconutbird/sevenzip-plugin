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
    /// Whether this is a directory
    pub is_dir: bool,
    /// File attributes (Windows-style, optional)
    pub attributes: Option<u32>,
    /// CRC32 checksum (optional)
    pub crc: Option<u32>,
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
