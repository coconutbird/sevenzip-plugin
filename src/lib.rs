//! # sevenzip-plugin
//!
//! A safe Rust framework for creating 7-Zip plugins.
//!
//! This crate hides all the unsafe COM/vtable complexity behind safe Rust traits.
//!
//! ## Example
//!
//! ```rust,ignore
//! use sevenzip_plugin::prelude::*;
//!
//! struct MyFormat {
//!     items: Vec<ArchiveItem>,
//!     data: Vec<Vec<u8>>,
//! }
//!
//! impl ArchiveFormat for MyFormat {
//!     fn name() -> &'static str { "MyFormat" }
//!     fn extension() -> &'static str { "myf" }
//!     fn class_id() -> [u8; 16] { /* unique GUID bytes */ }
//! }
//!
//! impl ArchiveReader for MyFormat {
//!     fn open(&mut self, data: &[u8]) -> Result<()> { /* ... */ }
//!     fn item_count(&self) -> usize { self.items.len() }
//!     fn get_item(&self, index: usize) -> Option<&ArchiveItem> { self.items.get(index) }
//!     fn extract(&mut self, index: usize) -> Result<Vec<u8>> { /* ... */ }
//! }
//!
//! sevenzip_plugin::register_format!(MyFormat);
//! ```

mod error;
mod traits;
mod types;

#[cfg(windows)]
#[doc(hidden)]
pub mod windows;

pub mod prelude {
    //! Re-exports of commonly used types and traits.
    pub use crate::error::*;
    pub use crate::traits::*;
    pub use crate::types::*;
}

pub use prelude::*;

// Re-export windows types for use by the macro (under different name to avoid conflict)
#[cfg(windows)]
#[doc(hidden)]
pub use ::windows as windows_crate;
