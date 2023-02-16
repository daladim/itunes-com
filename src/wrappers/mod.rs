//! Safe wrappers over the COM API.
//!
//! You usually want to start by creating an instance of the `iTunes` interface by [`iTunes::new`], then use its various methods.

#![allow(non_snake_case)]
#![allow(non_camel_case_types)]

pub mod iter;

// We'd rather use the re-exported versions, so that they are available to our users.
use crate::com::*;

use windows::core::BSTR;
use windows::core::HRESULT;
use windows::core::Interface;
use windows::Win32::Media::Multimedia::NS_E_PROPERTY_NOT_FOUND;

use windows::Win32::System::Com::{CoInitializeEx, CoCreateInstance, CLSCTX_ALL, COINIT_MULTITHREADED};
use windows::Win32::System::Com::VARIANT;

type DATE = f64; // This type must be a joke. https://learn.microsoft.com/en-us/cpp/atl-mfc-shared/date-type?view=msvc-170
type LONG = i32;

use widestring::ucstring::U16CString;
use num_traits::FromPrimitive;

mod private {
    //! The only reason for this private module is to have a "private" trait in publicly exported types
    //!
    //! See <https://github.com/rust-lang/rust/issues/34537>
    use super::*;

    pub trait ComObjectWrapper {
        type WrappedType: Interface;

        fn from_com_object(com_object: Self::WrappedType) -> Self;
        fn com_object(&self) -> &Self::WrappedType;
    }
}
use private::ComObjectWrapper;

macro_rules! com_wrapper_struct {
    ($struct_name:ident) => {
        ::paste::paste! {
            com_wrapper_struct!($struct_name as [<IIT $struct_name>]);
        }
    };
    ($struct_name:ident as $com_type:ident) => {
        pub struct $struct_name {
            com_object: crate::com::$com_type,
        }

        impl private::ComObjectWrapper for $struct_name {
            type WrappedType = $com_type;

            fn from_com_object(com_object: crate::com::$com_type) -> Self {
                Self {
                    com_object
                }
            }

            fn com_object(&self) -> &crate::com::$com_type {
                &self.com_object
            }
        }
    }
}

macro_rules! str_to_bstr {
    ($string_name:ident, $bstr_name:ident) => {
        let wide = U16CString::from_str_truncate($string_name);
        let $bstr_name = BSTR::from_wide(wide.as_slice())?;
    }
}

macro_rules! no_args {
    ($vis:vis $func_name:ident) => {
        no_args!($vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $func_name:ident as $inherited_type:ty) => {
        $vis fn $func_name(&self) -> windows::core::Result<()> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result: HRESULT = unsafe{ inherited_obj.$func_name() };
            result.ok()
        }
    };
}

macro_rules! get_bstr {
    ($vis:vis $func_name:ident) => {
        get_bstr!($vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $func_name:ident as $inherited_type:ty) => {
        $vis fn $func_name(&self) -> windows::core::Result<String> {
            let mut bstr = BSTR::default();
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(&mut bstr) };
            result.ok()?;

            let v: Vec<u16> = bstr.as_wide().to_vec();
            Ok(U16CString::from_vec_truncate(v).to_string_lossy())
        }
    }
}

macro_rules! internal_set_bstr {
    ($vis:vis $func_name:ident as $inherited_type:ty, key = $key:ident) => {
        $vis fn $func_name(&self, $key: String) -> windows::core::Result<()> {
            str_to_bstr!($key, bstr);
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(bstr) };
            result.ok()
        }
    };
}

macro_rules! set_bstr {
    ($vis:vis $key:ident) => {
        set_bstr!($vis $key as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $key:ident as $inherited_type:ty) => {
        ::paste::paste! {
            internal_set_bstr!($vis [<set_ $key>] as $inherited_type, key = $key);
        }
    };
    ($vis:vis $key:ident, no_set_prefix) => {
        internal_set_bstr!($vis $key as <Self as ComObjectWrapper>::WrappedType, key = $key);
    };
}

macro_rules! get_long {
    ($vis:vis $func_name:ident) => {
        get_long!($vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $func_name:ident as $inherited_type:ty) => {
        $vis fn $func_name(&self) -> windows::core::Result<LONG> {
            let mut value: LONG = 0;
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(&mut value as *mut LONG) };
            result.ok()?;

            Ok(value)
        }
    };
}


macro_rules! internal_set_long {
    ($vis:vis $func_name:ident as $inherited_type:ty, key = $key:ident) => {
        $vis fn $func_name(&self, $key: LONG) -> windows::core::Result<()> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name($key) };
            result.ok()
        }
    };
}

macro_rules! set_long {
    ($vis:vis $key:ident) => {
        ::paste::paste! {
            set_long!($vis $key as <Self as ComObjectWrapper>::WrappedType);
        }
    };
    ($vis:vis $key:ident as $inherited_type:ty) => {
        ::paste::paste! {
            internal_set_long!($vis [<set_ $key>] as $inherited_type, key = $key);
        }
    };
    ($vis:vis $key:ident, no_set_prefix) => {
        ::paste::paste! {
            internal_set_long!($vis $key as <Self as ComObjectWrapper>::WrappedType, key = $key);
        }
    };
}

macro_rules! set_variant {
    ($vis:vis $func_name:ident ( $arg:ident )) => {
        ::paste::paste! {
            $vis fn [<set_ $func_name>](&self, $arg: &VARIANT) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set_ $func_name>]($arg as *const VARIANT) };
                result.ok()
            }
        }
    };
}

macro_rules! get_f64 {
    ($vis:vis $func_name:ident, $float_name:ty) => {
        get_f64!($vis $func_name, $float_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $func_name:ident, $float_name:ty as $inherited_type:ty) => {
        $vis fn $func_name(&self) -> windows::core::Result<$float_name> {
            let mut value: f64 = 0.0;
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(&mut value) };
            result.ok()?;

            Ok(value)
        }
    };
}

macro_rules! set_f64 {
    ($vis:vis $key:ident, $float_name:ty) => {
        set_f64!($vis $key, $float_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $key:ident, $float_name:ty as $inherited_type:ty) => {
        ::paste::paste! {
            $vis fn [<set _$key>](&self, $key: $float_name) -> windows::core::Result<()> {
                let inherited_obj = self.com_object().cast::<$inherited_type>()?;
                let result = unsafe{ inherited_obj.[<set _$key>]($key) };
                result.ok()
            }
        }
    }
}

macro_rules! get_double {
    ($vis:vis $key:ident) => {
        get_f64!($vis $key, f64);
    };
    ($vis:vis $key:ident as $inherited_type:ty) => {
        get_f64!($vis $key, f64 as $inherited_type);
    }
}

macro_rules! set_double {
    ($vis:vis $key:ident) => {
        set_f64!($vis $key, f64);
    }
}

macro_rules! get_date {
    ($vis:vis $key:ident) => {
        get_f64!($vis $key, DATE);
    };
    ($vis:vis $key:ident as $inherited_type:ty) => {
        get_f64!($vis $key, DATE as $inherited_type);
    };
}

macro_rules! set_date {
    ($vis:vis $key:ident) => {
        set_f64!($vis $key, DATE);
    };
    ($vis:vis $key:ident as $inherited_type:ty) => {
        set_f64!($vis $key, DATE as $inherited_type);
    };
}

macro_rules! get_bool {
    ($vis:vis $func_name:ident) => {
        get_bool!($vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $func_name:ident as $inherited_type:ty) => {
        ::paste::paste! {
            $vis fn [<is _$func_name>](&self) -> windows::core::Result<bool> {
                let mut value = crate::com::FALSE;
                let inherited_obj = self.com_object().cast::<$inherited_type>()?;
                let result = unsafe{ inherited_obj.$func_name(&mut value) };
                result.ok()?;

                Ok(value.as_bool())
            }
        }
    };
}



macro_rules! internal_set_bool {
    ($vis:vis $func_name:ident as $inherited_type:ty, key = $key:ident) => {
        $vis fn $func_name(&self, $key: bool) -> windows::core::Result<()> {
            let variant_bool = match $key {
                true => crate::com::TRUE,
                false => crate::com::FALSE,
            };
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(variant_bool) };
            result.ok()
        }
    };
}

macro_rules! set_bool {
    ($vis:vis $key:ident) => {
        set_bool!($vis $key as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $key:ident as $inherited_type:ty) => {
        ::paste::paste! {
            internal_set_bool!($vis [<set_ $key>] as $inherited_type, key = $key);
        }
    };
    ($vis:vis $key:ident, no_set_prefix) => {
        internal_set_bool!($vis $key as <Self as ComObjectWrapper>::WrappedType, key = $key);
    }
}


macro_rules! get_enum {
    ($vis:vis $fn_name:ident, $enum_type:ty) => {
        get_enum!($vis $fn_name, $enum_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $fn_name:ident, $enum_type:ty as $inherited_type:ty) => {
        $vis fn $fn_name(&self) -> windows::core::Result<$enum_type> {
            let mut value: $enum_type = FromPrimitive::from_i32(0).unwrap();
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$fn_name(&mut value as *mut _) };
            result.ok()?;
            Ok(value)
        }
    };
}

macro_rules! set_enum {
    ($vis:vis $fn_name:ident, $enum_type:ty) => {
        set_enum!($vis $fn_name, $enum_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $fn_name:ident, $enum_type:ty as $inherited_type:ty) => {
        ::paste::paste! {
            $vis fn [<set _$fn_name>](&self, value: $enum_type) -> windows::core::Result<()> {
                let inherited_obj = self.com_object().cast::<$inherited_type>()?;
                let result = unsafe{ inherited_obj.[<set _$fn_name>](value) };
                result.ok()
            }
        }
    };
}

macro_rules! create_wrapped_object {
    ($obj_type:ty, $out_obj:ident) => {
        match $out_obj {
            None => Err(windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            )),
            Some(com_object) => Ok(
                <$obj_type>::from_com_object(com_object)
            )
        }
    };
}

macro_rules! get_object {
    ($vis:vis $fn_name:ident, $obj_type:ty) => {
        get_object!($vis $fn_name, $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $fn_name:ident, $obj_type:ty as $inherited_type:ty) => {
        $vis fn $fn_name(&self) -> windows::core::Result<$obj_type> {
            let mut out_obj = None;
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$fn_name(&mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    };
}

macro_rules! get_object_from_str {
    ($vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty) => {
        get_object_from_str!($vis $fn_name($arg_name) -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty as $inherited_type:ty) => {
        $vis fn $fn_name(&self, $arg_name: &str) -> windows::core::Result<$obj_type> {
            str_to_bstr!($arg_name, bstr);
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;

            let mut out_obj = None;
            let result = unsafe{ inherited_obj.$fn_name(bstr, &mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    };
}

macro_rules! get_object_from_variant {
    ($vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty) => {
        get_object_from_variant!($vis $fn_name($arg_name) -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty as $inherited_type:ty) => {
        $vis fn $fn_name(&self, $arg_name:&VARIANT) -> windows::core::Result<$obj_type> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;

            let mut out_obj = None;
            let result = unsafe{ inherited_obj.$fn_name($arg_name as *const VARIANT, &mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    };
}

macro_rules! get_object_from_long {
    ($vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty) => {
        get_object_from_long!($vis $fn_name($arg_name) -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty as $inherited_type:ty) => {
        $vis fn $fn_name(&self, $arg_name:LONG) -> windows::core::Result<$obj_type> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;

            let mut out_obj = None;
            let result = unsafe{ inherited_obj.$fn_name($arg_name, &mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    };
}

macro_rules! set_object {
    ($vis:vis $fn_name:ident, $obj_type:ty) => {
        ::paste::paste! {
            $vis fn [<set _$fn_name>](&self, data: $obj_type) -> windows::core::Result<()> {
                let object_to_set = data.com_object();
                let result = unsafe{ self.com_object.[<set _$fn_name>](object_to_set as *const _) };
                result.ok()
            }
        }
    }
}

macro_rules! item_by_name {
    ($vis:vis $obj_type:ty) => {
        $vis fn ItemByName(&self, name: String) -> windows::core::Result<$obj_type> {
            str_to_bstr!(name, bstr);

            let mut out_obj = None;
            let result = unsafe{ self.com_object.ItemByName(bstr, &mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    }
}

macro_rules! item_by_persistent_id {
    ($vis:vis $obj_type:ty) => {
        $vis fn ItemByPersistentID(&self, id: u64) -> windows::core::Result<$obj_type> {
            let b = id.to_le_bytes();
            let id_high = i32::from_le_bytes(b[..4].try_into().unwrap());
            let id_low = i32::from_le_bytes(b[4..].try_into().unwrap());

            let mut out_obj = None;
            let result = unsafe{ self.com_object.ItemByPersistentID(id_high, id_low, &mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    }
}


pub trait Iterable {
    type Item;

    // Provided by the COM API
    fn Count(&self) -> windows::core::Result<LONG>;
    // Provided by the COM API
    fn item(&self, index: LONG) -> windows::core::Result<<Self as Iterable>::Item>;
}

macro_rules! iterator {
    ($obj_type:ty, $item_type:ident) => {
        impl $obj_type {
            pub fn iter(&self) -> windows::core::Result<iter::Iterator<$obj_type, $item_type>> {
                iter::Iterator::new(&self)
            }
        }

        impl Iterable for $obj_type {
            type Item = $item_type;

            get_long!(Count);

            /// Returns an $item_type object corresponding to the given index (1-based).
            fn item(&self, index: LONG) -> windows::core::Result<<Self as Iterable>::Item> {
                let mut out_obj = None;
                let result = unsafe{ self.com_object.Item(index, &mut out_obj as *mut _) };
                result.ok()?;

                create_wrapped_object!($item_type, out_obj)
            }

            // /// Returns an IEnumVARIANT object which can enumerate the collection.
            // ///
            // /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
            // /// Feel free to open an issue or a pull request to fix this.
            // pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
            //     todo!()
            // }
        }
    }
}



/// The four IDs that uniquely identify an object
#[derive(Debug, Eq, PartialEq)]
pub struct ObjectIDs {
    pub sourceID: LONG,
    pub playlistID: LONG,
    pub trackID: LONG,
    pub databaseID: LONG,
}

/// Many COM objects inherit from this class, which provides some extra methods
pub trait IITObjectWrapper: private::ComObjectWrapper {
    /// Returns the four IDs that uniquely identify this object.
    fn GetITObjectIDs(&self) -> windows::core::Result<ObjectIDs> {
        let mut sourceID: LONG = 0;
        let mut playlistID: LONG = 0;
        let mut trackID: LONG = 0;
        let mut databaseID: LONG = 0;
        let iitobject = self.com_object().cast::<IITObject>().unwrap();
        let result = unsafe{ iitobject.GetITObjectIDs(
            &mut sourceID as *mut LONG,
            &mut playlistID as *mut LONG,
            &mut trackID as *mut LONG,
            &mut databaseID as *mut LONG,
        ) };
        result.ok()?;

        Ok(ObjectIDs{
            sourceID,
            playlistID,
            trackID,
            databaseID,
        })
    }

    /// The name of the object.
    get_bstr!(Name as IITObject);

    /// The name of the object.
    set_bstr!(Name as IITObject);

    /// The index of the object in internal application order (1-based).
    get_long!(Index as IITObject);

    /// The source ID of the object.
    get_long!(sourceID as IITObject);

    /// The playlist ID of the object.
    get_long!(playlistID as IITObject);

    /// The track ID of the object.
    get_long!(trackID as IITObject);

    /// The track database ID of the object.
    get_long!(TrackDatabaseID as IITObject);
}

/// IITSource Interface
///
/// See the generated [`IITSource_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Source);

impl IITObjectWrapper for Source {}

impl Source {
    /// The source kind.
    get_enum!(pub Kind, ITSourceKind);

    /// The total size of the source, if it has a fixed size.
    get_double!(pub Capacity);

    /// The free space on the source, if it has a fixed size.
    get_double!(pub FreeSpace);

    /// Returns a collection of playlists.
    get_object!(pub Playlists, PlaylistCollection);
}

/// IITPlaylistCollection Interface
///
/// See the generated [`IITPlaylistCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(PlaylistCollection);

impl PlaylistCollection {
    /// Returns an IITPlaylist object with the specified name.
    item_by_name!(pub Playlist);

    /// Returns an IITPlaylist object with the specified persistent ID.
    item_by_persistent_id!(pub Playlist);
}

iterator!(PlaylistCollection, Playlist);


/// Several COM objects inherit from this class, which provides some extra methods
pub trait IITPlaylistWrapper: private::ComObjectWrapper {
    /// Delete this playlist.
    no_args!(Delete as IITPlaylist);

    /// Start playing the first track in this playlist.
    no_args!(PlayFirstTrack as IITPlaylist);

    /// Print this playlist.
    fn Print(&self, showPrintDialog: bool, printKind: ITPlaylistPrintKind, theme: String) -> windows::core::Result<()> {
        let show = if showPrintDialog { TRUE } else { FALSE };
        str_to_bstr!(theme, theme);

        let inherited_obj = self.com_object().cast::<IITPlaylist>()?;
        let result = unsafe{ inherited_obj.Print(show, printKind, theme) };
        result.ok()
    }

    /// Search tracks in this playlist for the specified string.
    fn Search(&self, searchText: String, searchFields: ITPlaylistSearchField) -> windows::core::Result<TrackCollection> {
        str_to_bstr!(searchText, searchText);

        let mut out_obj = None;
        let inherited_obj = self.com_object().cast::<IITPlaylist>()?;
        let result = unsafe{ inherited_obj.Search(searchText, searchFields, &mut out_obj as *mut _) };
        result.ok()?;

        create_wrapped_object!(TrackCollection, out_obj)
    }

    /// The playlist kind.
    get_enum!(Kind, ITPlaylistKind as IITPlaylist);

    /// The source that contains this playlist.
    get_object!(Source, Source as IITPlaylist);

    /// The total length of all songs in the playlist (in seconds).
    get_long!(Duration as IITPlaylist);

    /// True if songs in the playlist are played in random order.
    get_bool!(Shuffle as IITPlaylist);

    /// True if songs in the playlist are played in random order.
    set_bool!(Shuffle as IITPlaylist);

    /// The total size of all songs in the playlist (in bytes).
    get_double!(Size as IITPlaylist);

    /// The playback repeat mode.
    get_enum!(SongRepeat, ITPlaylistRepeatMode as IITPlaylist);

    /// The playback repeat mode.
    set_enum!(SongRepeat, ITPlaylistRepeatMode as IITPlaylist);

    /// The total length of all songs in the playlist (in MM:SS format).
    get_bstr!(Time as IITPlaylist);

    /// True if the playlist is visible in the Source list.
    get_bool!(Visible as IITPlaylist);

    /// Returns a collection of tracks in this playlist.
    get_object!(Tracks, TrackCollection as IITPlaylist);
}

/// IITPlaylist Interface
///
/// See the generated [`IITPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Playlist);

impl IITObjectWrapper for Playlist {}

impl IITPlaylistWrapper for Playlist {}


/// IITTrackCollection Interface
///
/// See the generated [`IITTrackCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(TrackCollection);

impl TrackCollection {
    /// Returns an IITTrack object corresponding to the given index, where the index is defined by the play order of the playlist containing the track collection (1-based).
    get_object_from_long!(pub ItemByPlayOrder(Index) -> Track);

    /// Returns an IITTrack object with the specified name.
    item_by_name!(pub Track);

    /// Returns an IITTrack object with the specified persistent ID.
    item_by_persistent_id!(pub Track);
}

iterator!(TrackCollection, Track);

/// Several COM objects inherit from this class, which provides some extra methods
pub trait IITTrackWrapper: private::ComObjectWrapper {
    /// Delete this track.
    no_args!(Delete as IITTrack);

    /// Start playing this track.
    no_args!(Play as IITTrack);

    /// Add artwork from an image file to this track.
    get_object_from_str!(AddArtworkFromFile(filePath) -> Artwork as IITTrack);

    /// The track kind.
    get_enum!(Kind, ITTrackKind as IITTrack);

    /// The playlist that contains this track.
    get_object!(Playlist, Playlist as IITTrack);

    /// The album containing the track.
    get_bstr!(Album as IITTrack);

    /// The album containing the track.
    set_bstr!(Album as IITTrack);

    /// The artist/source of the track.
    get_bstr!(Artist as IITTrack);

    /// The artist/source of the track.
    set_bstr!(Artist as IITTrack);

    /// The bit rate of the track (in kbps).
    get_long!(BitRate as IITTrack);

    /// The tempo of the track (in beats per minute).
    get_long!(BPM as IITTrack);

    /// The tempo of the track (in beats per minute).
    set_long!(BPM as IITTrack);

    /// Freeform notes about the track.
    get_bstr!(Comment as IITTrack);

    /// Freeform notes about the track.
    set_bstr!(Comment as IITTrack);

    /// True if this track is from a compilation album.
    get_bool!(Compilation as IITTrack);

    /// True if this track is from a compilation album.
    set_bool!(Compilation as IITTrack);

    /// The composer of the track.
    get_bstr!(Composer as IITTrack);

    /// The composer of the track.
    set_bstr!(Composer as IITTrack);

    /// The date the track was added to the playlist.
    get_date!(DateAdded as IITTrack);

    /// The total number of discs in the source album.
    get_long!(DiscCount as IITTrack);

    /// The total number of discs in the source album.
    set_long!(DiscCount as IITTrack);

    /// The index of the disc containing the track on the source album.
    get_long!(DiscNumber as IITTrack);

    /// The index of the disc containing the track on the source album.
    set_long!(DiscNumber as IITTrack);

    /// The length of the track (in seconds).
    get_long!(Duration as IITTrack);

    /// True if the track is checked for playback.
    get_bool!(Enabled as IITTrack);

    /// True if the track is checked for playback.
    set_bool!(Enabled as IITTrack);

    /// The name of the EQ preset of the track.
    get_bstr!(EQ as IITTrack);

    /// The name of the EQ preset of the track.
    set_bstr!(EQ as IITTrack);

    /// The stop time of the track (in seconds).
    set_long!(Finish as IITTrack);

    /// The stop time of the track (in seconds).
    get_long!(Finish as IITTrack);

    /// The music/audio genre (category) of the track.
    get_bstr!(Genre as IITTrack);

    /// The music/audio genre (category) of the track.
    set_bstr!(Genre as IITTrack);

    /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
    get_bstr!(Grouping as IITTrack);

    /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
    set_bstr!(Grouping as IITTrack);

    /// A text description of the track.
    get_bstr!(KindAsString as IITTrack);

    /// The modification date of the content of the track.
    get_date!(ModificationDate as IITTrack);

    /// The number of times the track has been played.
    get_long!(PlayedCount as IITTrack);

    /// The number of times the track has been played.
    set_long!(PlayedCount as IITTrack);

    /// The date and time the track was last played.  A value of zero means no played date.
    get_date!(PlayedDate as IITTrack);

    /// The date and time the track was last played.  A value of zero means no played date.
    set_date!(PlayedDate as IITTrack);

    /// The play order index of the track in the owner playlist (1-based).
    get_long!(PlayOrderIndex as IITTrack);

    /// The rating of the track (0 to 100).
    get_long!(Rating as IITTrack);

    /// The rating of the track (0 to 100).
    set_long!(Rating as IITTrack);

    /// The sample rate of the track (in Hz).
    get_long!(SampleRate as IITTrack);

    /// The size of the track (in bytes).
    get_long!(Size as IITTrack);

    /// The start time of the track (in seconds).
    get_long!(Start as IITTrack);

    /// The start time of the track (in seconds).
    set_long!(Start as IITTrack);

    /// The length of the track (in MM:SS format).
    get_bstr!(Time as IITTrack);

    /// The total number of tracks on the source album.
    get_long!(TrackCount as IITTrack);

    /// The total number of tracks on the source album.
    set_long!(TrackCount as IITTrack);

    /// The index of the track on the source album.
    get_long!(TrackNumber as IITTrack);

    /// The index of the track on the source album.
    set_long!(TrackNumber as IITTrack);

    /// The relative volume adjustment of the track (-100% to 100%).
    get_long!(VolumeAdjustment as IITTrack);

    /// The relative volume adjustment of the track (-100% to 100%).
    set_long!(VolumeAdjustment as IITTrack);

    /// The year the track was recorded/released.
    get_long!(Year as IITTrack);

    /// The year the track was recorded/released.
    set_long!(Year as IITTrack);

    /// Returns a collection of artwork.
    get_object!(Artwork, ArtworkCollection as IITTrack);
}

/// IITTrack Interface
///
/// See the generated [`IITTrack_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Track);

impl IITObjectWrapper for Track {}

impl IITTrackWrapper for Track {}

/// IITArtwork Interface
///
/// See the generated [`IITArtwork_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Artwork);

impl Artwork {
    /// Delete this piece of artwork from the track.
    no_args!(pub Delete);

    /// Replace existing artwork data with new artwork from an image file.
    set_bstr!(pub SetArtworkFromFile, no_set_prefix);

    /// Save artwork data to an image file.
    set_bstr!(pub SaveArtworkToFile, no_set_prefix);

    /// The format of the artwork.
    get_enum!(pub Format, ITArtworkFormat);

    /// True if the artwork was downloaded by iTunes.
    get_bool!(pub IsDownloadedArtwork);

    /// The description for the artwork.
    get_bstr!(pub Description);

    /// The description for the artwork.
    set_bstr!(pub Description);
}

/// IITArtworkCollection Interface
///
/// See the generated [`IITArtworkCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(ArtworkCollection);

impl ArtworkCollection {}

iterator!(ArtworkCollection, Artwork);

/// IITSourceCollection Interface
///
/// See the generated [`IITSourceCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(SourceCollection);

impl SourceCollection {
    /// Returns an IITSource object with the specified name.
    item_by_name!(pub Source);

    /// Returns an IITSource object with the specified persistent ID.
    item_by_persistent_id!(pub Source);
}

iterator!(SourceCollection, Source);

/// IITEncoder Interface
///
/// See the generated [`IITEncoder_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Encoder);

impl Encoder {
    /// The name of the the encoder.
    get_bstr!(pub Name);

    /// The data format created by the encoder.
    get_bstr!(pub Format);
}

/// IITEncoderCollection Interface
///
/// See the generated [`IITEncoderCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(EncoderCollection);

impl EncoderCollection {
    /// Returns an IITEncoder object with the specified name.
    item_by_name!(pub Encoder);
}

iterator!(EncoderCollection, Encoder);

/// IITEQPreset Interface
///
/// See the generated [`IITEQPreset_Impl`] trait for more documentation about each function.
com_wrapper_struct!(EQPreset);

impl EQPreset {
    /// The name of the the EQ preset.
    get_bstr!(pub Name);

    /// True if this EQ preset can be modified.
    get_bool!(pub Modifiable);

    /// The equalizer preamp level (-12.0 db to +12.0 db).
    get_double!(pub Preamp);

    /// The equalizer preamp level (-12.0 db to +12.0 db).
    set_double!(pub Preamp);

    /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
    get_double!(pub Band1);

    /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
    set_double!(pub Band1);

    /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
    get_double!(pub Band2);

    /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
    set_double!(pub Band2);

    /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
    get_double!(pub Band3);

    /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
    set_double!(pub Band3);

    /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
    get_double!(pub Band4);

    /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
    set_double!(pub Band4);

    /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
    get_double!(pub Band5);

    /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
    set_double!(pub Band5);

    /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
    get_double!(pub Band6);

    /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
    set_double!(pub Band6);

    /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
    get_double!(pub Band7);

    /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
    set_double!(pub Band7);

    /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
    get_double!(pub Band8);

    /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
    set_double!(pub Band8);

    /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
    get_double!(pub Band9);

    /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
    set_double!(pub Band9);

    /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
    get_double!(pub Band10);

    /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
    set_double!(pub Band10);

    /// Delete this EQ preset.
    internal_set_bool!(pub Delete as <Self as ComObjectWrapper>::WrappedType, key = updateAllTracks);

    /// Rename this EQ preset.
    pub fn Rename(&self, newName: BSTR, updateAllTracks: bool) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITEQPresetCollection Interface
///
/// See the generated [`IITEQPresetCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(EQPresetCollection);

impl EQPresetCollection {
    /// Returns an IITEQPreset object with the specified name.
    item_by_name!(pub EQPreset);
}

iterator!(EQPresetCollection, EQPreset);

/// IITOperationStatus Interface
///
/// See the generated [`IITOperationStatus_Impl`] trait for more documentation about each function.
com_wrapper_struct!(OperationStatus);

impl OperationStatus {
    /// True if the operation is still in progress.
    get_bool!(pub InProgress);

    /// Returns a collection containing the tracks that were generated by the operation.
    get_object!(pub Tracks, TrackCollection);
}

/// IITConvertOperationStatus Interface
///
/// See the generated [`IITConvertOperationStatus_Impl`] trait for more documentation about each function.
com_wrapper_struct!(ConvertOperationStatus);

impl ConvertOperationStatus {
    /// Returns the current conversion status.
    pub unsafe fn GetConversionStatus(&self, trackName: *mut BSTR, progressValue: *mut LONG, maxProgressValue: *mut LONG)  -> windows::core::Result<()> {
        todo!()
    }
    /// Stops the current conversion operation.
    no_args!(pub StopConversion);

    /// Returns the name of the track currently being converted.
    get_bstr!(pub trackName);

    /// Returns the current progress value for the track being converted.
    get_long!(pub progressValue);

    /// Returns the maximum progress value for the track being converted.
    get_long!(pub maxProgressValue);
}

/// IITLibraryPlaylist Interface
///
/// See the generated [`IITLibraryPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(LibraryPlaylist);

impl IITPlaylistWrapper for LibraryPlaylist {}

impl LibraryPlaylist {
    /// Add the specified file path to the library.
    get_object_from_str!(pub AddFile(filePath) -> OperationStatus);

    /// Add the specified array of file paths to the library. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    get_object_from_variant!(pub AddFiles(filePaths) -> OperationStatus);

    /// Add the specified streaming audio URL to the library.
    get_object_from_str!(pub AddURL(URL) -> URLTrack);

    /// Add the specified track to the library.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    get_object_from_variant!(pub AddTrack(iTrackToAdd) -> Track);
}

/// IITURLTrack Interface
///
/// See the generated [`IITURLTrack_Impl`] trait for more documentation about each function.
com_wrapper_struct!(URLTrack);

impl IITTrackWrapper for URLTrack {}

impl URLTrack {
    /// The URL of the stream represented by this track.
    get_bstr!(pub URL);

    /// The URL of the stream represented by this track.
    set_bstr!(pub URL);

    /// True if this is a podcast track.
    get_bool!(pub Podcast);

    /// Update the podcast feed for this track.
    no_args!(pub UpdatePodcastFeed);

    /// Start downloading the podcast episode that corresponds to this track.
    no_args!(pub DownloadPodcastEpisode);

    /// Category for the track.
    get_bstr!(pub Category);

    /// Category for the track.
    set_bstr!(pub Category);

    /// Description for the track.
    get_bstr!(pub Description);

    /// Description for the track.
    set_bstr!(pub Description);

    /// Long description for the track.
    get_bstr!(pub LongDescription);

    /// Long description for the track.
    set_bstr!(pub LongDescription);

    /// Reveal the track in the main browser window.
    no_args!(pub Reveal);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    get_long!(pub AlbumRating);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    set_long!(pub AlbumRating);

    /// The album rating kind.
    get_enum!(pub AlbumRatingKind, ITRatingKind);

    /// The track rating kind.
    get_enum!(pub ratingKind, ITRatingKind);

    /// Returns a collection of playlists that contain the song that this track represents.
    get_object!(pub Playlists, PlaylistCollection);
}

/// IITUserPlaylist Interface
///
/// See the generated [`IITUserPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(UserPlaylist);

impl IITPlaylistWrapper for UserPlaylist {}

impl UserPlaylist {
    /// Add the specified file path to the user playlist.
    get_object_from_str!(pub AddFile(filePath) -> OperationStatus);

    /// Add the specified array of file paths to the user playlist. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    get_object_from_variant!(pub AddFiles(filePaths) -> OperationStatus);

    /// Add the specified streaming audio URL to the user playlist.
    get_object_from_str!(pub AddURL(URL) -> URLTrack);

    /// Add the specified track to the user playlist.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    get_object_from_variant!(pub AddTrack(iTrackToAdd) -> Track);

    /// True if the user playlist is being shared.
    get_bool!(pub Shared);

    /// True if the user playlist is being shared.
    set_bool!(pub Shared);

    /// True if this is a smart playlist.
    get_bool!(pub Smart);

    /// The playlist special kind.
    get_enum!(pub SpecialKind, ITUserPlaylistSpecialKind);

    /// The parent of this playlist.
    get_object!(pub Parent, UserPlaylist);

    /// Creates a new playlist in a folder playlist.
    get_object_from_str!(pub CreatePlaylist(playlistName) -> Playlist);

    /// Creates a new folder in a folder playlist.
    get_object_from_str!(pub CreateFolder(folderName) -> Playlist);

    /// The parent of this playlist.
    set_variant!(pub Parent(iParentPlayList));

    /// Reveal the user playlist in the main browser window.
    no_args!(pub Reveal);
}

/// IITVisual Interface
///
/// See the generated [`IITVisual_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Visual);

impl Visual {
    /// The name of the the visual plug-in.
    get_bstr!(pub Name);
}

/// IITVisualCollection Interface
///
/// See the generated [`IITVisualCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(VisualCollection);

impl VisualCollection {
    /// Returns an IITVisual object with the specified name.
    item_by_name!(pub Visual);
}

iterator!(VisualCollection, Visual);

/// IITWindow Interface
///
/// See the generated [`IITWindow_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Window);

impl Window {
    /// The title of the window.
    get_bstr!(pub Name);

    /// The window kind.
    get_enum!(pub Kind, ITWindowKind);

    /// True if the window is visible. Note that the main browser window cannot be hidden.
    get_bool!(pub Visible);

    /// True if the window is visible. Note that the main browser window cannot be hidden.
    set_bool!(pub Visible);

    /// True if the window is resizable.
    get_bool!(pub Resizable);

    /// True if the window is minimized.
    get_bool!(pub Minimized);

    /// True if the window is minimized.
    set_bool!(pub Minimized);

    /// True if the window is maximizable.
    get_bool!(pub Maximizable);

    /// True if the window is maximized.
    get_bool!(pub Maximized);

    /// True if the window is maximized.
    set_bool!(pub Maximized);

    /// True if the window is zoomable.
    get_bool!(pub Zoomable);

    /// True if the window is zoomed.
    get_bool!(pub Zoomed);

    /// True if the window is zoomed.
    set_bool!(pub Zoomed);

    /// The screen coordinate of the top edge of the window.
    get_long!(pub Top);

    /// The screen coordinate of the top edge of the window.
    set_long!(pub Top);

    /// The screen coordinate of the left edge of the window.
    get_long!(pub Left);

    /// The screen coordinate of the left edge of the window.
    set_long!(pub Left);

    /// The screen coordinate of the bottom edge of the window.
    get_long!(pub Bottom);

    /// The screen coordinate of the bottom edge of the window.
    set_long!(pub Bottom);

    /// The screen coordinate of the right edge of the window.
    get_long!(pub Right);

    /// The screen coordinate of the right edge of the window.
    set_long!(pub Right);

    /// The width of the window.
    get_long!(pub Width);

    /// The width of the window.
    set_long!(pub Width);

    /// The height of the window.
    get_long!(pub Height);

    /// The height of the window.
    set_long!(pub Height);
}

/// IITBrowserWindow Interface
///
/// See the generated [`IITBrowserWindow_Impl`] trait for more documentation about each function.
com_wrapper_struct!(BrowserWindow);

impl BrowserWindow {
    /// True if window is in MiniPlayer mode.
    get_bool!(pub MiniPlayer);

    /// True if window is in MiniPlayer mode.
    set_bool!(pub MiniPlayer);

    /// Returns a collection containing the currently selected track or tracks.
    get_object!(pub SelectedTracks, TrackCollection);

    /// The currently selected playlist in the Source list.
    get_object!(pub SelectedPlaylist, Playlist);

    /// The currently selected playlist in the Source list.
    set_variant!(pub SelectedPlaylist(iPlaylist));
}

/// IITWindowCollection Interface
///
/// See the generated [`IITWindowCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(WindowCollection);

impl WindowCollection {
    /// Returns an IITWindow object with the specified name.
    item_by_name!(pub Window);
}

iterator!(WindowCollection, Window);

/// IiTunes Interface
///
/// See the generated [`IiTunes_Impl`] trait for more documentation about each function.
com_wrapper_struct!(iTunes as IiTunes);

impl iTunes {
    /// Create a new COM object to communicate with iTunes
    pub fn new() -> windows::core::Result<Self> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;
        }

        Ok(Self {
            com_object: unsafe { CoCreateInstance(&crate::com::ITUNES_APP_COM_GUID, None, CLSCTX_ALL)? },
        })
    }

    /// Reposition to the beginning of the current track or go to the previous track if already at start of current track.
    no_args!(pub BackTrack);

    /// Skip forward in a playing track.
    no_args!(pub FastForward);

    /// Advance to the next track in the current playlist.
    no_args!(pub NextTrack);

    /// Pause playback.
    no_args!(pub Pause);

    /// Play the currently targeted track.
    no_args!(pub Play);

    /// Play the specified file path, adding it to the library if not already present.
    set_bstr!(pub PlayFile, no_set_prefix);

    /// Toggle the playing/paused state of the current track.
    no_args!(pub PlayPause);

    /// Return to the previous track in the current playlist.
    no_args!(pub PreviousTrack);

    /// Disable fast forward/rewind and resume playback, if playing.
    no_args!(pub Resume);

    /// Skip backwards in a playing track.
    no_args!(pub Rewind);

    /// Stop playback.
    no_args!(pub Stop);

    /// Start converting the specified file path.
    get_object_from_str!(pub ConvertFile(filePath) -> OperationStatus);

    /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    get_object_from_variant!(pub ConvertFiles(filePaths) -> OperationStatus);

    /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    get_object_from_variant!(pub ConvertTrack(iTrackToConvert) -> OperationStatus);

    /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
    get_object_from_variant!(pub ConvertTracks(iTracksToConvert) -> OperationStatus);

    /// Returns true if this version of the iTunes type library is compatible with the specified version.
    pub fn CheckVersion(&self, majorVersion: LONG, minorVersion: LONG, isCompatible: *mut VARIANT_BOOL) -> windows::core::Result<()> {
        todo!()
    }
    // TODO: implement this, but IITObject is a trait
    // /// Returns an IITObject corresponding to the specified IDs.
    // pub fn GetITObjectByID(&self, sourceID: LONG, playlistID: LONG, trackID: LONG, databaseID: LONG, iObject: *mut Option<IITObject>) -> windows::core::Result<()> {
    //     todo!()
    // }
    /// Creates a new playlist in the main library.
    get_object_from_str!(pub CreatePlaylist(playlistName) -> Playlist);

    /// Open the specified iTunes Store or streaming audio URL.
    set_bstr!(pub OpenURL, no_set_prefix);

    /// Go to the iTunes Store home page.
    no_args!(pub GotoMusicStoreHomePage);

    /// Update the contents of the iPod.
    no_args!(pub UpdateIPod);

    /// [id(0x60020015)]
    /// (no other documentation provided)
    pub fn Authorize(&self, numElems: LONG, data: *const VARIANT, names: *const BSTR) -> windows::core::Result<()> {
        todo!()
    }
    /// Exits the iTunes application.
    no_args!(pub Quit);

    /// Returns a collection of music sources (music library, CD, device, etc.).
    get_object!(pub Sources, SourceCollection);

    /// Returns a collection of encoders.
    get_object!(pub Encoders, EncoderCollection);

    /// Returns a collection of EQ presets.
    get_object!(pub EQPresets, EQPresetCollection);

    /// Returns a collection of visual plug-ins.
    get_object!(pub Visuals, VisualCollection);

    /// Returns a collection of windows.
    get_object!(pub Windows, WindowCollection);

    /// Returns the sound output volume (0 = minimum, 100 = maximum).
    get_long!(pub SoundVolume);

    /// Returns the sound output volume (0 = minimum, 100 = maximum).
    set_long!(pub SoundVolume);

    /// True if sound output is muted.
    get_bool!(pub Mute);

    /// True if sound output is muted.
    set_bool!(pub Mute);

    /// Returns the current player state.
    get_enum!(pub PlayerState, ITPlayerState);

    /// Returns the player's position within the currently playing track in seconds.
    get_long!(pub PlayerPosition);

    /// Returns the player's position within the currently playing track in seconds.
    set_long!(pub PlayerPosition);

    /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
    get_object!(pub CurrentEncoder, Encoder);

    /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
    set_object!(pub CurrentEncoder, Encoder);

    /// True if visuals are currently being displayed.
    get_bool!(pub VisualsEnabled);

    /// True if visuals are currently being displayed.
    set_bool!(pub VisualsEnabled);

    /// True if the visuals are displayed using the entire screen.
    get_bool!(pub FullScreenVisuals);

    /// True if the visuals are displayed using the entire screen.
    set_bool!(pub FullScreenVisuals);

    /// Returns the size of the displayed visual.
    get_enum!(pub VisualSize, ITVisualSize);

    /// Returns the size of the displayed visual.
    set_enum!(pub VisualSize, ITVisualSize);

    /// Returns the currently selected visual plug-in.
    get_object!(pub CurrentVisual, Visual);

    /// Returns the currently selected visual plug-in.
    set_object!(pub CurrentVisual, Visual);

    /// True if the equalizer is enabled.
    get_bool!(pub EQEnabled);

    /// True if the equalizer is enabled.
    set_bool!(pub EQEnabled);

    /// Returns the currently selected EQ preset.
    get_object!(pub CurrentEQPreset, EQPreset);

    /// Returns the currently selected EQ preset.
    set_object!(pub CurrentEQPreset, EQPreset);

    /// The name of the current song in the playing stream (provided by streaming server).
    get_bstr!(pub CurrentStreamTitle);

    /// The URL of the playing stream or streaming web site (provided by streaming server).
    get_bstr!(pub set_CurrentStreamURL);

    /// Returns the main iTunes browser window.
    get_object!(pub BrowserWindow, BrowserWindow);

    /// Returns the EQ window.
    get_object!(pub EQWindow, Window);

    /// Returns the source that represents the main library.
    get_object!(pub LibrarySource, Source);

    /// Returns the main library playlist in the main library source.
    get_object!(pub LibraryPlaylist, LibraryPlaylist);

    /// Returns the currently targeted track.
    get_object!(pub CurrentTrack, Track);

    /// Returns the playlist containing the currently targeted track.
    get_object!(pub CurrentPlaylist, Playlist);

    /// Returns a collection containing the currently selected track or tracks.
    get_object!(pub SelectedTracks, TrackCollection);

    /// Returns the version of the iTunes application.
    get_bstr!(pub Version);

    /// [id(0x6002003b)]
    /// (no other documentation provided)
    set_long!(pub SetOptions, no_set_prefix);

    /// Start converting the specified file path.
    get_object_from_str!(pub ConvertFile2(filePath) -> ConvertOperationStatus);

    /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    get_object_from_variant!(pub ConvertFiles2(filePaths) -> ConvertOperationStatus);

    /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    get_object_from_variant!(pub ConvertTrack2(iTrackToConvert) -> ConvertOperationStatus);

    /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
    get_object_from_variant!(pub ConvertTracks2(iTracksToConvert) -> ConvertOperationStatus);

    /// True if iTunes will process APPCOMMAND Windows messages.
    get_bool!(pub AppCommandMessageProcessingEnabled);

    /// True if iTunes will process APPCOMMAND Windows messages.
    set_bool!(pub AppCommandMessageProcessingEnabled);

    /// True if iTunes will force itself to be the foreground application when it displays a dialog.
    get_bool!(pub ForceToForegroundOnDialog);

    /// True if iTunes will force itself to be the foreground application when it displays a dialog.
    set_bool!(pub ForceToForegroundOnDialog);

    /// Create a new EQ preset.
    get_object_from_str!(pub CreateEQPreset(eqPresetName) -> EQPreset);

    /// Creates a new playlist in an existing source.
    pub fn CreatePlaylistInSource(&self, playlistName: BSTR, iSource: *const VARIANT, iPlaylist: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// Retrieves the current state of the player buttons.
    pub fn GetPlayerButtonsState(&self, previousEnabled: *mut VARIANT_BOOL, playPauseStopState: *mut ITPlayButtonState, nextEnabled: *mut VARIANT_BOOL) -> windows::core::Result<()> {
        todo!()
    }
    /// Simulate click on a player control button.
    pub fn PlayerButtonClicked(&self, playerButton: ITPlayerButton, playerButtonModifierKeys: LONG) -> windows::core::Result<()> {
        todo!()
    }
    /// True if the Shuffle property is writable for the specified playlist.
    pub fn CanSetShuffle(&self, iPlaylist: *const VARIANT, CanSetShuffle: *mut VARIANT_BOOL) -> windows::core::Result<()> {
        todo!()
    }
    /// True if the SongRepeat property is writable for the specified playlist.
    pub fn CanSetSongRepeat(&self, iPlaylist: *const VARIANT, CanSetSongRepeat: *mut VARIANT_BOOL) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns an IITConvertOperationStatus object if there is currently a conversion in progress.
    get_object!(pub ConvertOperationStatus, ConvertOperationStatus);

    /// Subscribe to the specified podcast feed URL.
    set_bstr!(pub SubscribeToPodcast, no_set_prefix);

    /// Update all podcast feeds.
    no_args!(pub UpdatePodcastFeeds);

    /// Creates a new folder in the main library.
    get_object_from_str!(pub CreateFolder(folderName) -> Playlist);

    /// Creates a new folder in an existing source.
    pub fn CreateFolderInSource(&self, folderName: BSTR, iSource: *const VARIANT, iFolder: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// True if the sound volume control is enabled.
    get_bool!(pub SoundVolumeControlEnabled);

    /// The full path to the current iTunes library XML file.
    get_bstr!(pub LibraryXMLPath);

    /// Returns the high 32 bits of the persistent ID of the specified IITObject.
    pub unsafe fn ITObjectPersistentIDHigh(&self, iObject: *const VARIANT, highID: *mut LONG) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns the low 32 bits of the persistent ID of the specified IITObject.
    pub unsafe fn ITObjectPersistentIDLow(&self, iObject: *const VARIANT, lowID: *mut LONG) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns the high and low 32 bits of the persistent ID of the specified IITObject.
    pub unsafe fn GetITObjectPersistentIDs(&self, iObject: *const VARIANT, highID: *mut LONG, lowID: *mut LONG) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns the player's position within the currently playing track in milliseconds.
    get_long!(pub PlayerPositionMS);

    /// Returns the player's position within the currently playing track in milliseconds.
    set_long!(pub PlayerPositionMS);
}

/// IITAudioCDPlaylist Interface
///
/// See the generated [`IITAudioCDPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(AudioCDPlaylist);

impl IITPlaylistWrapper for AudioCDPlaylist {}

impl AudioCDPlaylist {
    /// The artist of the CD.
    get_bstr!(pub Artist);

    /// True if this CD is a compilation album.
    get_bool!(pub Compilation);

    /// The composer of the CD.
    get_bstr!(pub Composer);

    /// The total number of discs in this CD's album.
    get_long!(pub DiscCount);

    /// The index of the CD disc in the source album.
    get_long!(pub DiscNumber);

    /// The genre of the CD.
    get_bstr!(pub Genre);

    /// The year the album was recorded/released.
    get_long!(pub Year);

    /// Reveal the CD playlist in the main browser window.
    no_args!(pub Reveal);
}

/// IITIPodSource Interface
///
/// See the generated [`IITIPodSource_Impl`] trait for more documentation about each function.
com_wrapper_struct!(IPodSource);

impl IPodSource {
    /// Update the contents of the iPod.
    no_args!(pub UpdateIPod);

    /// Eject the iPod.
    no_args!(pub EjectIPod);

    /// The iPod software version.
    get_bstr!(pub SoftwareVersion);
}

/// IITFileOrCDTrack Interface
///
/// See the generated [`IITFileOrCDTrack_Impl`] trait for more documentation about each function.
com_wrapper_struct!(FileOrCDTrack);

impl IITTrackWrapper for FileOrCDTrack {}

impl FileOrCDTrack {
    /// The full path to the file represented by this track.
    get_bstr!(pub Location);

    /// Update this track's information with the information stored in its file.
    no_args!(pub UpdateInfoFromFile);

    /// True if this is a podcast track.
    get_bool!(pub Podcast);

    /// Update the podcast feed for this track.
    no_args!(pub UpdatePodcastFeed);

    /// True if playback position is remembered.
    get_bool!(pub RememberBookmark);

    /// True if playback position is remembered.
    set_bool!(pub RememberBookmark);

    /// True if track is skipped when shuffling.
    get_bool!(pub ExcludeFromShuffle);

    /// True if track is skipped when shuffling.
    set_bool!(pub ExcludeFromShuffle);

    /// Lyrics for the track.
    get_bstr!(pub Lyrics);

    /// Lyrics for the track.
    set_bstr!(pub Lyrics);

    /// Category for the track.
    get_bstr!(pub Category);

    /// Category for the track.
    set_bstr!(pub Category);

    /// Description for the track.
    get_bstr!(pub Description);

    /// Description for the track.
    set_bstr!(pub Description);

    /// Long description for the track.
    get_bstr!(pub LongDescription);

    /// Long description for the track.
    set_bstr!(pub LongDescription);

    /// The bookmark time of the track (in seconds).
    get_long!(pub BookmarkTime);

    /// The bookmark time of the track (in seconds).
    set_long!(pub BookmarkTime);

    /// The video track kind.
    get_enum!(pub VideoKind, ITVideoKind);

    /// The video track kind.
    set_enum!(pub VideoKind, ITVideoKind);

    /// The number of times the track has been skipped.
    get_long!(pub SkippedCount);

    /// The number of times the track has been skipped.
    set_long!(pub SkippedCount);

    /// The date and time the track was last skipped.  A value of zero means no skipped date.
    get_date!(pub SkippedDate);

    /// The date and time the track was last skipped.  A value of zero means no skipped date.
    set_date!(pub SkippedDate);

    /// True if track is part of a gapless album.
    get_bool!(pub PartOfGaplessAlbum);

    /// True if track is part of a gapless album.
    set_bool!(pub PartOfGaplessAlbum);

    /// The album artist of the track.
    get_bstr!(pub AlbumArtist);

    /// The album artist of the track.
    set_bstr!(pub AlbumArtist);

    /// The show name of the track.
    get_bstr!(pub Show);

    /// The show name of the track.
    set_bstr!(pub Show);

    /// The season number of the track.
    get_long!(pub SeasonNumber);

    /// The season number of the track.
    set_long!(pub SeasonNumber);

    /// The episode ID of the track.
    get_bstr!(pub EpisodeID);

    /// The episode ID of the track.
    set_bstr!(pub EpisodeID);

    /// The episode number of the track.
    get_long!(pub EpisodeNumber);

    /// The episode number of the track.
    set_long!(pub EpisodeNumber);

    /// The high 32-bits of the size of the track (in bytes).
    get_long!(pub Size64High);

    /// The low 32-bits of the size of the track (in bytes).
    get_long!(pub Size64Low);

    /// True if track has not been played.
    get_bool!(pub Unplayed);

    /// True if track has not been played.
    set_bool!(pub Unplayed);

    /// The album used for sorting.
    get_bstr!(pub SortAlbum);

    /// The album used for sorting.
    set_bstr!(pub SortAlbum);

    /// The album artist used for sorting.
    get_bstr!(pub SortAlbumArtist);

    /// The album artist used for sorting.
    set_bstr!(pub SortAlbumArtist);

    /// The artist used for sorting.
    get_bstr!(pub SortArtist);

    /// The artist used for sorting.
    set_bstr!(pub SortArtist);

    /// The composer used for sorting.
    get_bstr!(pub SortComposer);

    /// The composer used for sorting.
    set_bstr!(pub SortComposer);

    /// The track name used for sorting.
    get_bstr!(pub SortName);

    /// The track name used for sorting.
    set_bstr!(pub SortName);

    /// The show name used for sorting.
    get_bstr!(pub SortShow);

    /// The show name used for sorting.
    set_bstr!(pub SortShow);

    /// Reveal the track in the main browser window.
    no_args!(pub Reveal);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    get_long!(pub AlbumRating);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    set_long!(pub AlbumRating);

    /// The album rating kind.
    get_enum!(pub AlbumRatingKind, ITRatingKind);

    /// The track rating kind.
    get_enum!(pub ratingKind, ITRatingKind);

    /// Returns a collection of playlists that contain the song that this track represents.
    get_object!(pub Playlists, PlaylistCollection);

    /// The full path to the file represented by this track.
    set_bstr!(pub Location);

    /// The release date of the track.  A value of zero means no release date.
    get_date!(pub ReleaseDate);
}

/// IITPlaylistWindow Interface
///
/// See the generated [`IITPlaylistWindow_Impl`] trait for more documentation about each function.
com_wrapper_struct!(PlaylistWindow);

impl PlaylistWindow {
    /// Returns a collection containing the currently selected track or tracks.
    get_object!(pub SelectedTracks, TrackCollection);

    /// The playlist displayed in the window.
    get_object!(pub Playlist, Playlist);
}
