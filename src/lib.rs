//! Bindings over iTunes COM API for Windows
//!
//! # What is this for?
//!
//! iTunes COM API makes it possible to control the local instance of iTunes.
//!
//! This crate is able to **read** info (about playlists, songs, etc.) from the local iTunes instance.<br/>
//! It is also able to **edit** data (add songs to playlists, change track ratings, etc.) on the local iTunes instance.<br/>
//! It is also able to interact with iTunes state and settings (get info about the currently opened windows, get or set EQs, etc.)<br/>
//! It is **not** meant to read or edit "cloud" playlists, or to do anything network-related.
//!
//! # OS and software compatibility
//!
//! This crate is Windows-only.
//! Currently, only iTunes is supported, as Apple Music on Windows does not (yet?) expose a COM interface.<br/>
//! On macOS, it is possible to control iTunes and Apple Music using Apple Script.
//!
//! # How can this crate be used?
//!
//! ## Raw bindings
//!
//! This crate provides raw bindings over the COM API. See the [`sys`] module.
//!
//! ## Safe bindings
//!
//! In case it is built with the `wrappers` Cargo feature, it also provides safe, Rust-typed wrappers over this API.
//! See the [`wrappers`] module.
//!
//! ## Examples
//!
//! Examples are available in the `examples/` folder. Run them with `cargo run --example ... --all-features`.


pub mod sys;
#[cfg(feature = "wrappers")]
pub mod wrappers;
