//! Raw bindings over iTunes COM API
//!
//! ## Where to find them?
//! Some resources online (e.g. [here](https://www.joshkunz.com/iTunesControl/main.html) or [there](https://documentation.help/iTunesCOM/main.html))
//! expose the documentation generated from IDL files.<br/>
//! They do not tell in which order they are defined, which is meaningful.
//!
//! Opening iTunes.exe in oleview.exe (File > View TypeLib, then open iTunes.exe) generates (pseudo)IDL files that are suitable to correctly define bindings.
//! That's then a matter of finding-and-replacing IDL patterns with Rust patterns and hope the bindings are eventually correct.

mod com_interfaces;
mod com_enums;

pub use com_interfaces::*;
pub use com_enums::*;

/// The GUID used to create an instance of [`crate::sys::IiTunes`].
pub const ITUNES_APP_COM_GUID: windows::core::GUID = windows::core::GUID::from_u128(0xDC0C2640_1415_4644_875C_6F4D769839BA);

// These types are part of the public API and must be re-exported so that users can use them in their right version.
/// Re-exported type from windows-rs.
pub use windows::{
    core::{BSTR, HRESULT},
    Win32::System::Com::VARIANT,
    Win32::Foundation::VARIANT_BOOL,
    Win32::System::Ole::IEnumVARIANT,
};

/// Convenience constant
pub const TRUE: crate::sys::VARIANT_BOOL = crate::sys::VARIANT_BOOL(-1);
/// Convenience constant
pub const FALSE: crate::sys::VARIANT_BOOL = crate::sys::VARIANT_BOOL(0);
