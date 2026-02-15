# sevenzip-plugin-rs

A safe Rust framework for creating 7-Zip plugins.

This crate hides all the unsafe COM/vtable complexity behind safe Rust traits, letting you focus on your archive format's logic.

## Features

- **Safe abstractions** - Implement simple Rust traits instead of dealing with COM interfaces
- **Read support** - Open and extract files from your custom archive format
- **Write support** - Create and update archives (optional)
- **Windows only** - 7-Zip plugins are Windows DLLs

## Usage

Add to your `Cargo.toml`:

```toml
[lib]
crate-type = ["cdylib"]

[dependencies]
sevenzip-plugin = "0.1"
```

Implement the required traits and register your format:

```rust
use sevenzip_plugin::prelude::*;

#[derive(Default)]
struct MyFormat {
    items: Vec<ArchiveItem>,
    data: Vec<Vec<u8>>,
}

impl ArchiveFormat for MyFormat {
    fn name() -> &'static str { "MyFormat" }
    fn extension() -> &'static str { "myf" }
    fn class_id() -> [u8; 16] {
        // Unique GUID for your format
        [0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08,
         0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F, 0x10]
    }
}

impl ArchiveReader for MyFormat {
    fn open(&mut self, data: &[u8]) -> Result<()> {
        // Parse your archive format here
        Ok(())
    }

    fn item_count(&self) -> usize {
        self.items.len()
    }

    fn get_item(&self, index: usize) -> Option<&ArchiveItem> {
        self.items.get(index)
    }

    fn extract(&mut self, index: usize) -> Result<Vec<u8>> {
        // Return the decompressed data for the item
        self.data.get(index).cloned().ok_or_else(|| Error::IndexOutOfBounds {
            index,
            count: self.data.len(),
        })
    }
}

// Register the format - generates required DLL exports
sevenzip_plugin::register_format!(MyFormat);
```

## Building

Build as a Windows DLL:

```bash
cargo build --release
```

The output will be in `target/release/your_crate_name.dll`.

## Installation

Copy the built DLL to your 7-Zip installation's `Formats` directory (e.g., `C:\Program Files\7-Zip\Formats\`).

## Traits

- **`ArchiveFormat`** - Define your format's metadata (name, extension, GUID)
- **`ArchiveReader`** - Implement reading and extraction
- **`ArchiveUpdater`** - Implement creating/updating archives (optional)

## License

MIT
