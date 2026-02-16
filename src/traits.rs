//! Safe traits that plugin authors implement.

use crate::error::Result;
use crate::types::{
    ArchiveItem, PasswordProvider, PasswordRequester, ProgressCallback, UpdateItem,
};
use std::io::{Read, Seek, Write};

/// A trait alias for types that implement both `Read` and `Seek`.
///
/// This is used for streaming archive input, allowing plugins to read
/// archive data on-demand without buffering the entire file in memory.
pub trait ReadSeek: Read + Seek {}

// Blanket implementation for all types that implement Read + Seek
impl<T: Read + Seek> ReadSeek for T {}

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
    /// Open and parse the archive from a streaming reader.
    ///
    /// This is called when 7-Zip opens an archive file.
    /// The reader supports both sequential reads and seeking, allowing
    /// you to read only the parts of the archive you need.
    ///
    /// - `reader`: A seekable reader for the archive data
    /// - `size`: Total size of the archive in bytes
    ///
    /// Store any parsed metadata internally for later extraction.
    fn open(&mut self, reader: &mut dyn ReadSeek, size: u64) -> Result<()>;

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

    /// Open and parse an encrypted archive with password support.
    ///
    /// This is called instead of `open()` when 7-Zip provides a password callback.
    /// Override this method to support encrypted archives where the password is
    /// needed to read the archive header/index.
    ///
    /// - `reader`: A seekable reader for the archive data
    /// - `size`: Total size of the archive in bytes
    /// - `password_requester`: A callback to request passwords from the user.
    ///   Call `password_requester.get_password()` when you need to decrypt.
    ///
    /// The default implementation ignores the password callback and calls `open()`.
    fn open_with_password(
        &mut self,
        reader: &mut dyn ReadSeek,
        size: u64,
        _password_requester: Option<&dyn PasswordRequester>,
    ) -> Result<()> {
        self.open(reader, size)
    }

    /// Extract an item's data with password support (streaming).
    ///
    /// This is called during extraction when 7-Zip provides a password callback.
    /// Override this method to support encrypted archives where individual files
    /// may be encrypted (like ZIP with encrypted entries).
    ///
    /// - `index`: The item index to extract
    /// - `writer`: Output stream to write the extracted data to
    /// - `password_requester`: A callback to request passwords from the user.
    ///   Call `password_requester.get_password()` when you need to decrypt.
    ///
    /// The default implementation ignores the password callback and calls `extract_to()`.
    fn extract_to_with_password(
        &mut self,
        index: usize,
        writer: &mut dyn Write,
        _password_requester: Option<&dyn PasswordRequester>,
    ) -> Result<u64> {
        self.extract_to(index, writer)
    }
}

/// Trait for writing/updating archives.
///
/// Implement this to allow 7-Zip to create new archives or modify existing ones.
/// Creating a new archive is simply updating from empty data with all `AddNew` items.
pub trait ArchiveUpdater: ArchiveReader {
    /// Update an existing archive with full streaming I/O.
    ///
    /// - `existing`: A seekable reader for the existing archive (or empty if creating new)
    /// - `existing_size`: Size of the existing archive in bytes (0 if creating new)
    /// - `updates`: List of update operations (copy existing or add new)
    /// - `writer`: Output stream to write the new archive to
    /// - `progress`: Optional callback to report progress during the write phase.
    ///   The callback receives `(completed_bytes, total_bytes)` and should return
    ///   `true` to continue or `false` to request cancellation.
    ///
    /// Returns the number of bytes written to the output.
    ///
    /// This method enables true zero-copy updates by reading from the existing
    /// archive and writing to the output without buffering everything in memory.
    ///
    /// ## Progress Reporting
    ///
    /// Plugin implementers should call the progress callback periodically during
    /// write operations to keep the 7-Zip progress bar updated. This is especially
    /// important for formats that do significant work (compression, encryption, etc.)
    /// in the write phase.
    fn update_streaming(
        &mut self,
        existing: &mut dyn ReadSeek,
        existing_size: u64,
        updates: Vec<UpdateItem>,
        writer: &mut dyn Write,
        progress: Option<ProgressCallback<'_>>,
    ) -> Result<u64>;

    /// Update an existing archive with encryption support.
    ///
    /// This is called instead of `update_streaming()` when the user has requested
    /// encryption for the archive. Override this method to support creating encrypted
    /// archives.
    ///
    /// - `existing`: A seekable reader for the existing archive (or empty if creating new)
    /// - `existing_size`: Size of the existing archive in bytes (0 if creating new)
    /// - `updates`: List of update operations (copy existing or add new)
    /// - `writer`: Output stream to write the new archive to
    /// - `progress`: Optional callback to report progress during the write phase
    /// - `password_provider`: A callback to get the encryption password.
    ///   Call `password_provider.get_password()` to get the password for encryption.
    ///
    /// The default implementation ignores the password and calls `update_streaming()`.
    fn update_streaming_with_password(
        &mut self,
        existing: &mut dyn ReadSeek,
        existing_size: u64,
        updates: Vec<UpdateItem>,
        writer: &mut dyn Write,
        progress: Option<ProgressCallback<'_>>,
        _password_provider: Option<&dyn PasswordProvider>,
    ) -> Result<u64> {
        self.update_streaming(existing, existing_size, updates, writer, progress)
    }
}
