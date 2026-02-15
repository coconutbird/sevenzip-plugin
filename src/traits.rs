//! Safe traits that plugin authors implement.

use crate::error::Result;
use crate::types::{ArchiveItem, UpdateItem};
use std::io::Write;

/// Metadata about an archive format.
///
/// This trait defines the static properties of your archive format.
pub trait ArchiveFormat: Default + Send + 'static {
    /// Human-readable name of the format (e.g., "ERA", "ZIP").
    fn name() -> &'static str;

    /// File extension without the dot (e.g., "era", "zip").
    fn extension() -> &'static str;

    /// Unique class ID (GUID) for this format.
    /// Generate with: `[0x12, 0x34, ..., 0xEF]` (16 bytes)
    fn class_id() -> [u8; 16];

    /// Optional file signature/magic bytes for format detection.
    /// Return `None` if the format cannot be detected by magic bytes
    /// (e.g., encrypted formats).
    fn signature() -> Option<&'static [u8]> {
        None
    }

    /// Whether this format supports creating new archives.
    fn supports_write() -> bool {
        false
    }

    /// Whether this format supports editing existing archives.
    fn supports_update() -> bool {
        false
    }
}

/// Trait for reading archives.
///
/// Implement this to allow 7-Zip to open and extract from your archive format.
pub trait ArchiveReader: ArchiveFormat {
    /// Open and parse the archive from raw bytes.
    ///
    /// This is called when 7-Zip opens an archive file.
    /// Store the parsed data internally for later extraction.
    fn open(&mut self, data: &[u8]) -> Result<()>;

    /// Returns the number of items in the archive.
    fn item_count(&self) -> usize;

    /// Get information about an item by index.
    fn get_item(&self, index: usize) -> Option<&ArchiveItem>;

    /// Extract an item's data by index.
    ///
    /// Returns the uncompressed file contents.
    fn extract(&mut self, index: usize) -> Result<Vec<u8>>;

    /// Extract an item's data directly to a writer (streaming).
    ///
    /// This avoids allocating a `Vec<u8>` for the entire file contents,
    /// which is more memory efficient for large files.
    ///
    /// The default implementation calls `extract()` and writes the result.
    /// Override this for better memory efficiency with large files.
    ///
    /// Returns the number of bytes written.
    fn extract_to(&mut self, index: usize, writer: &mut dyn Write) -> Result<u64> {
        let data = self.extract(index)?;
        let len = data.len() as u64;
        writer
            .write_all(&data)
            .map_err(|e| crate::error::Error::Io(e.to_string()))?;
        Ok(len)
    }

    /// Close the archive and release resources.
    ///
    /// Called when 7-Zip is done with the archive.
    fn close(&mut self) {
        // Default: do nothing (Drop will clean up)
    }

    /// Get the physical size of the archive (optional).
    fn physical_size(&self) -> Option<u64> {
        None
    }
}

/// Trait for writing/updating archives.
///
/// Implement this to allow 7-Zip to create new archives or modify existing ones.
/// Creating a new archive is simply updating from empty data with all `AddNew` items.
pub trait ArchiveUpdater: ArchiveReader {
    /// Update an existing archive.
    ///
    /// - `existing_data`: The raw bytes of the existing archive
    /// - `updates`: List of update operations (copy existing or add new)
    ///
    /// Returns the complete new archive data as bytes.
    fn update(&mut self, existing_data: &[u8], updates: Vec<UpdateItem>) -> Result<Vec<u8>>;

    /// Update an existing archive, writing directly to a writer (streaming).
    ///
    /// This avoids allocating a `Vec<u8>` for the entire output archive,
    /// which is more memory efficient for large archives.
    ///
    /// The default implementation calls `update()` and writes the result.
    /// Override this for better memory efficiency with large archives.
    ///
    /// Returns the number of bytes written.
    fn update_to(
        &mut self,
        existing_data: &[u8],
        updates: Vec<UpdateItem>,
        writer: &mut dyn Write,
    ) -> Result<u64> {
        let data = self.update(existing_data, updates)?;
        let len = data.len() as u64;
        writer
            .write_all(&data)
            .map_err(|e| crate::error::Error::Io(e.to_string()))?;
        Ok(len)
    }
}
