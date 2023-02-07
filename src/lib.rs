//! Bindings over iTunes COM API
//!
//! This crate makes it possible to control the local instance of iTunes on a Windows PC.
//!
//! ## Abilities
//!
//! This crate is able to **read** info (about playlists, songs, etc.) from the local iTunes instance.<br/>
//! It is also able to **edit** data (add songs to playlists, change track ratings, etc.) on the local iTunes instance.<br/>
//! It is **not** meant to read or edit "cloud" playlists, or to do anything network-related.
//!
//! ## How to use them
//! TODO: code example to start by creating an istance of ...
//!
//!
//! ## Where to find them?
//! Some resources online (e.g. [here](https://www.joshkunz.com/iTunesControl/main.html) or [there](https://documentation.help/iTunesCOM/main.html)
//! expose the documentation generated from IDL files.<br/>
//! They do not tell in which order they are defined, which is meaningful.
//!
//! Opening iTunes.exe in oleview.exe (File > View TypeLib, then open iTunes.exe) generates (pseudo)IDL files that are suitable to correctly define bindings.
//! That's then a matter of finding-and-replacing IDL patterns with Rust patterns and hope the bindings are eventually correct.
//!
//! ## See also
//! TOD: what about Apple Music ? macOS ?

mod com_interfaces;
mod com_enums;

pub use com_interfaces::*;
pub use com_enums::*;

/// The GUID used to create an instance of [`crate::IiTunes`].
pub const ITUNES_APP_COM_GUID: windows::core::GUID = windows::core::GUID::from_u128(0xDC0C2640_1415_4644_875C_6F4D769839BA);

// These types are part of the public API and must be re-exported so that users can use them in their right version.
/// Re-exported type from windows-rs.
pub use windows::{
    core::{BSTR, HRESULT},
    Win32::Foundation::VARIANT_BOOL,
    Win32::System::Ole::IEnumVARIANT,
};

/// Convenience constant
pub const TRUE: crate::VARIANT_BOOL = crate::VARIANT_BOOL(-1);
/// Convenience constant
pub const FALSE: crate::VARIANT_BOOL = crate::VARIANT_BOOL(0);
