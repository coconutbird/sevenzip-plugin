//! Error types for archive operations.

use std::fmt;

/// Result type for archive operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Errors that can occur during archive operations.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum Error {
    /// The archive format is not recognized or is corrupted.
    InvalidFormat(String),
    /// An I/O error occurred.
    Io(String),
    /// The requested item index is out of bounds.
    IndexOutOfBounds { index: usize, count: usize },
    /// A required feature is not supported.
    NotSupported(String),
    /// Generic error with a message.
    Other(String),
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::InvalidFormat(msg) => write!(f, "Invalid format: {}", msg),
            Error::Io(err) => write!(f, "I/O error: {}", err),
            Error::IndexOutOfBounds { index, count } => {
                write!(f, "Index {} out of bounds (count: {})", index, count)
            }
            Error::NotSupported(msg) => write!(f, "Not supported: {}", msg),
            Error::Other(msg) => write!(f, "{}", msg),
        }
    }
}

impl std::error::Error for Error {}

impl From<std::io::Error> for Error {
    fn from(err: std::io::Error) -> Self {
        Error::Io(err.to_string())
    }
}

impl From<String> for Error {
    fn from(msg: String) -> Self {
        Error::Other(msg)
    }
}

impl From<&str> for Error {
    fn from(msg: &str) -> Self {
        Error::Other(msg.to_string())
    }
}
