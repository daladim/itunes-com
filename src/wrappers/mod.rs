//! Safe wrappers over the COM API.

#![allow(non_snake_case)]
#![allow(non_camel_case_types)]

// We'd rather use the re-exported versions, so that they are available to our users.
use crate::com::*;

use windows::core::BSTR;
use windows::core::HRESULT;
use windows::Win32::Media::Multimedia::NS_E_PROPERTY_NOT_FOUND;

use windows::Win32::System::Com::VARIANT;

type DATE = f64; // This type must be a joke. https://learn.microsoft.com/en-us/cpp/atl-mfc-shared/date-type?view=msvc-170
type LONG = i32;

use widestring::ucstring::U16CString;
use num_traits::FromPrimitive;


macro_rules! com_wrapper_struct {
    ($struct_name:ident) => {
        ::paste::paste! {
            pub struct $struct_name {
                com_object: crate::com::[<IIT $struct_name>],
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
    ($func_name:ident) => {
        pub fn $func_name(&self) -> windows::core::Result<()> {
            let result: HRESULT = unsafe{ self.com_object.$func_name() };
            result.ok()
        }
    }
}

macro_rules! get_bstr {
    ($func_name:ident) => {
        pub fn $func_name(&self) -> windows::core::Result<String> {
            let mut bstr = BSTR::default();
            let result = unsafe{ self.com_object.$func_name(&mut bstr) };
            result.ok()?;

            let v: Vec<u16> = bstr.as_wide().to_vec();
            Ok(U16CString::from_vec_truncate(v).to_string_lossy())
        }
    }
}

macro_rules! set_bstr {
    ($key:ident) => {
        ::paste::paste! {
            pub fn [<set _$key>](&self, $key: String) -> windows::core::Result<()> {
                str_to_bstr!($key, bstr);
                let result = unsafe{ self.com_object.[<set _$key>](bstr) };
                result.ok()
            }
        }
    };
    ($key:ident, no_set_prefix) => {
        pub fn $key(&self, $key: String) -> windows::core::Result<()> {
            str_to_bstr!($key, bstr);
            let result = unsafe{ self.com_object.$key(bstr) };
            result.ok()
        }
    }
}

macro_rules! get_long {
    ($func_name:ident) => {
        pub fn $func_name(&self) -> windows::core::Result<i32> {
            let mut value: i32 = 0;
            let result = unsafe{ self.com_object.$func_name(&mut value as *mut i32) };
            result.ok()?;

            Ok(value)
        }
    }
}

macro_rules! set_long {
    ($key:ident) => {
        ::paste::paste! {
            pub fn [<set _$key>](&self, $key: LONG) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set _$key>]($key) };
                result.ok()
            }
        }
    };
    ($key:ident, no_set_prefix) => {
        ::paste::paste! {
            pub fn [<set _$key>](&self, $key: LONG) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.$key($key) };
                result.ok()
            }
        }
    }
}

macro_rules! get_f64 {
    ($func_name:ident, $float_name:ty) => {
        pub fn $func_name(&self) -> windows::core::Result<$float_name> {
            let mut value: f64 = 0.0;
            let result = unsafe{ self.com_object.$func_name(&mut value) };
            result.ok()?;

            Ok(value)
        }
    }
}

macro_rules! set_f64 {
    ($key:ident, $float_name:ty) => {
        ::paste::paste! {
            pub fn [<set _$key>](&self, $key: $float_name) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set _$key>]($key) };
                result.ok()
            }
        }
    }
}

macro_rules! get_double {
    ($key:ident) => {
        get_f64!($key, f64);
    }
}

macro_rules! set_double {
    ($key:ident) => {
        set_f64!($key, f64);
    }
}

macro_rules! get_date {
    ($key:ident) => {
        get_f64!($key, DATE);
    }
}

macro_rules! set_date {
    ($key:ident) => {
        set_f64!($key, DATE);
    }
}

macro_rules! get_bool {
    ($func_name:ident) => {
        ::paste::paste! {
            pub fn [<is _$func_name>](&self) -> windows::core::Result<bool> {
                let mut value = crate::com::FALSE;
                let result = unsafe{ self.com_object.$func_name(&mut value) };
                result.ok()?;

                Ok(value.as_bool())
            }
        }
    }
}

macro_rules! set_bool {
    ($key:ident) => {
        ::paste::paste! {
            pub fn [<set _$key>](&self, $key: bool) -> windows::core::Result<()> {
                let variant_bool = match $key {
                    true => crate::com::TRUE,
                    false => crate::com::FALSE,
                };
                let result = unsafe{ self.com_object.[<set _$key>](variant_bool) };
                result.ok()
            }
        }
    };
    ($key:ident, $arg_name:ident, no_set_prefix) => {
        ::paste::paste! {
            pub fn [<set _$key>](&self, $arg_name: bool) -> windows::core::Result<()> {
                let variant_bool = match $arg_name {
                    true => crate::com::TRUE,
                    false => crate::com::FALSE,
                };
                let result = unsafe{ self.com_object.$key(variant_bool) };
                result.ok()
            }
        }
    };
}


macro_rules! get_enum {
    ($fn_name:ident, $enum_type:ty) => {
        pub fn $fn_name(&self) -> windows::core::Result<$enum_type> {
            let mut value: $enum_type = FromPrimitive::from_i32(0).unwrap();
            let result = unsafe{ self.com_object.$fn_name(&mut value as *mut _) };
            result.ok()?;
            Ok(value)
        }
    }
}

macro_rules! set_enum {
    ($fn_name:ident, $enum_type:ty) => {
        ::paste::paste! {
            pub fn [<set _$fn_name>](&self, value: $enum_type) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set _$fn_name>](value) };
                result.ok()
            }
        }
    }
}

macro_rules! get_object {
    ($fn_name:ident, $obj_type:ident) => {
        pub fn $fn_name(&self) -> windows::core::Result<$obj_type> {
            let mut out_obj = None;
            let result = unsafe{ self.com_object.$fn_name(&mut out_obj as *mut _) };
            result.ok()?;
            out_obj.ok_or_else(|| windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            ))
        }
    }
}

macro_rules! set_object {
    ($fn_name:ident, $obj_type:ident) => {
        ::paste::paste! {
            pub fn [<set _$fn_name>](&self, data: $obj_type) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set _$fn_name>](&data as *const _) };
                result.ok()
            }
        }
    }
}

macro_rules! item_by_name {
    ($obj_type:ident) => {
        pub fn ItemByName(&self, name: String) -> windows::core::Result<$obj_type> {
            let wide = U16CString::from_str_truncate(name);
            let bstr = BSTR::from_wide(wide.as_slice())?;

            let mut out_obj = None;
            let result = unsafe{ self.com_object.ItemByName(bstr, &mut out_obj as *mut _) };
            result.ok()?;
            out_obj.ok_or_else(|| windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            ))
        }
    }
}

macro_rules! item_by_persistent_id {
    ($obj_type:ty) => {
        pub fn ItemByPersistentID(&self, id: u64) -> windows::core::Result<$obj_type> {
            let b = id.to_le_bytes();
            let id_high = i32::from_le_bytes(b[..4].try_into().unwrap());
            let id_low = i32::from_le_bytes(b[4..].try_into().unwrap());

            let mut out_obj = None;
            let result = unsafe{ self.com_object.ItemByPersistentID(id_high, id_low, &mut out_obj as *mut _) };
            result.ok()?;
            out_obj.ok_or_else(|| windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            ))
        }
    }
}

macro_rules! get_enum {
    ($fn_name:ident, $enum_type:ty) => {
        pub fn $fn_name(&self) -> windows::core::Result<$enum_type> {
            let mut value: $enum_type = FromPrimitive::from_i32(0).unwrap();
            let result = unsafe{ self.com_object.$fn_name(&mut value as *mut _) };
            result.ok()?;
            Ok(value)
        }
    }
}

macro_rules! set_enum {
    ($fn_name:ident, $enum_type:ty) => {
        ::paste::paste! {
            pub fn [<set _$fn_name>](&self, value: $enum_type) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set _$fn_name>](value) };
                result.ok()
            }
        }
    }
}

macro_rules! get_object {
    ($fn_name:ident, $obj_type:ident) => {
        pub fn $fn_name(&self) -> windows::core::Result<$obj_type> {
            let mut out_obj = None;
            let result = unsafe{ self.com_object.$fn_name(&mut out_obj as *mut _) };
            result.ok()?;
            out_obj.ok_or_else(|| windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            ))
        }
    }
}

macro_rules! set_object {
    ($fn_name:ident, $obj_type:ident) => {
        ::paste::paste! {
            pub fn [<set _$fn_name>](&self, data: $obj_type) -> windows::core::Result<()> {
                let result = unsafe{ self.com_object.[<set _$fn_name>](&data as *const _) };
                result.ok()
            }
        }
    }
}



/// IITObject Interface
///
/// See the generated [`IITObject_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Object);

impl Object {
    /// Returns the four IDs that uniquely identify this object.
    pub unsafe fn GetITObjectIDs(&self, sourceID: *mut LONG, playlistID: *mut LONG, trackID: *mut LONG, databaseID: *mut LONG) -> windows::core::Result<()> {
        todo!()
    }
    /// The name of the object.
    get_bstr!(Name);

    /// The name of the object.
    set_bstr!(Name);

    /// The index of the object in internal application order (1-based).
    get_long!(Index);

    /// The source ID of the object.
    get_long!(sourceID);

    /// The playlist ID of the object.
    get_long!(playlistID);

    /// The track ID of the object.
    get_long!(trackID);

    /// The track database ID of the object.
    get_long!(TrackDatabaseID);
}

/// IITSource Interface
///
/// See the generated [`IITSource_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Source);

impl Source {
    /// The source kind.
    get_enum!(Kind, ITSourceKind);

    /// The total size of the source, if it has a fixed size.
    get_double!(Capacity);

    /// The free space on the source, if it has a fixed size.
    get_double!(FreeSpace);

    /// Returns a collection of playlists.
    get_object!(Playlists, IITPlaylistCollection);
}

/// IITPlaylistCollection Interface
///
/// See the generated [`IITPlaylistCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(PlaylistCollection);

impl PlaylistCollection {
    /// Returns the number of playlists in the collection.
    get_long!(Count);

    /// Returns an IITPlaylist object corresponding to the given index (1-based).
    item!(IITPlaylist);

    /// Returns an IITPlaylist object with the specified name.
    item_by_name!(IITPlaylist);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns an IITPlaylist object with the specified persistent ID.
    item_by_persistent_id!(IITPlaylist);
}

/// IITPlaylist Interface
///
/// See the generated [`IITPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Playlist);

impl Playlist {
    /// Delete this playlist.
    no_args!(Delete);

    /// Start playing the first track in this playlist.
    no_args!(PlayFirstTrack);

    /// Print this playlist.
    pub fn Print(&self, showPrintDialog: VARIANT_BOOL, printKind: ITPlaylistPrintKind, theme: BSTR) -> windows::core::Result<()> {
        todo!()
    }
    /// Search tracks in this playlist for the specified string.
    pub fn Search(&self, searchText: BSTR, searchFields: ITPlaylistSearchField, iTrackCollection: *mut Option<IITTrackCollection>) -> windows::core::Result<()> {
        todo!()
    }
    /// The playlist kind.
    get_enum!(Kind, ITPlaylistKind);

    /// The source that contains this playlist.
    get_object!(Source, IITSource);

    /// The total length of all songs in the playlist (in seconds).
    get_long!(Duration);

    /// True if songs in the playlist are played in random order.
    get_bool!(Shuffle);

    /// True if songs in the playlist are played in random order.
    set_bool!(Shuffle);

    /// The total size of all songs in the playlist (in bytes).
    get_double!(Size);

    /// The playback repeat mode.
    get_enum!(SongRepeat, ITPlaylistRepeatMode);

    /// The playback repeat mode.
    set_enum!(SongRepeat, ITPlaylistRepeatMode);

    /// The total length of all songs in the playlist (in MM:SS format).
    get_bstr!(Time);

    /// True if the playlist is visible in the Source list.
    get_bool!(Visible);

    /// Returns a collection of tracks in this playlist.
    get_object!(Tracks, IITTrackCollection);
}

/// IITTrackCollection Interface
///
/// See the generated [`IITTrackCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(TrackCollection);

impl TrackCollection {
    /// Returns the number of tracks in the collection.
    get_long!(Count);

    /// Returns an IITTrack object corresponding to the given fixed index, where the index is independent of the play order (1-based).
    item!(IITTrack);

    /// Returns an IITTrack object corresponding to the given index, where the index is defined by the play order of the playlist containing the track collection (1-based).
    pub fn ItemByPlayOrder(&self, Index: LONG, iTrack: *mut Option<IITTrack>) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns an IITTrack object with the specified name.
    item_by_name!(IITTrack);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns an IITTrack object with the specified persistent ID.
    item_by_persistent_id!(IITTrack);
}

/// IITTrack Interface
///
/// See the generated [`IITTrack_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Track);

impl Track {
    /// Delete this track.
    no_args!(Delete);

    /// Start playing this track.
    no_args!(Play);

    /// Add artwork from an image file to this track.
    pub fn AddArtworkFromFile(&self, filePath: BSTR, iArtwork: *mut Option<IITArtwork>) -> windows::core::Result<()> {
        todo!()
    }
    /// The track kind.
    get_enum!(Kind, ITTrackKind);

    /// The playlist that contains this track.
    get_object!(Playlist, IITPlaylist);

    /// The album containing the track.
    get_bstr!(Album);

    /// The album containing the track.
    set_bstr!(Album);

    /// The artist/source of the track.
    get_bstr!(Artist);

    /// The artist/source of the track.
    set_bstr!(Artist);

    /// The bit rate of the track (in kbps).
    get_long!(BitRate);

    /// The tempo of the track (in beats per minute).
    get_long!(BPM);

    /// The tempo of the track (in beats per minute).
    set_long!(BPM);

    /// Freeform notes about the track.
    get_bstr!(Comment);

    /// Freeform notes about the track.
    set_bstr!(Comment);

    /// True if this track is from a compilation album.
    get_bool!(Compilation);

    /// True if this track is from a compilation album.
    set_bool!(Compilation);

    /// The composer of the track.
    get_bstr!(Composer);

    /// The composer of the track.
    set_bstr!(Composer);

    /// The date the track was added to the playlist.
    get_date!(DateAdded);

    /// The total number of discs in the source album.
    get_long!(DiscCount);

    /// The total number of discs in the source album.
    set_long!(DiscCount);

    /// The index of the disc containing the track on the source album.
    get_long!(DiscNumber);

    /// The index of the disc containing the track on the source album.
    set_long!(DiscNumber);

    /// The length of the track (in seconds).
    get_long!(Duration);

    /// True if the track is checked for playback.
    get_bool!(Enabled);

    /// True if the track is checked for playback.
    set_bool!(Enabled);

    /// The name of the EQ preset of the track.
    get_bstr!(EQ);

    /// The name of the EQ preset of the track.
    set_bstr!(EQ);

    /// The stop time of the track (in seconds).
    set_long!(Finish);

    /// The stop time of the track (in seconds).
    get_long!(Finish);

    /// The music/audio genre (category) of the track.
    get_bstr!(Genre);

    /// The music/audio genre (category) of the track.
    set_bstr!(Genre);

    /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
    get_bstr!(Grouping);

    /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
    set_bstr!(Grouping);

    /// A text description of the track.
    get_bstr!(KindAsString);

    /// The modification date of the content of the track.
    get_date!(ModificationDate);

    /// The number of times the track has been played.
    get_long!(PlayedCount);

    /// The number of times the track has been played.
    set_long!(PlayedCount);

    /// The date and time the track was last played.  A value of zero means no played date.
    get_date!(PlayedDate);

    /// The date and time the track was last played.  A value of zero means no played date.
    set_date!(PlayedDate);

    /// The play order index of the track in the owner playlist (1-based).
    get_long!(PlayOrderIndex);

    /// The rating of the track (0 to 100).
    get_long!(Rating);

    /// The rating of the track (0 to 100).
    set_long!(Rating);

    /// The sample rate of the track (in Hz).
    get_long!(SampleRate);

    /// The size of the track (in bytes).
    get_long!(Size);

    /// The start time of the track (in seconds).
    get_long!(Start);

    /// The start time of the track (in seconds).
    set_long!(Start);

    /// The length of the track (in MM:SS format).
    get_bstr!(Time);

    /// The total number of tracks on the source album.
    get_long!(TrackCount);

    /// The total number of tracks on the source album.
    set_long!(TrackCount);

    /// The index of the track on the source album.
    get_long!(TrackNumber);

    /// The index of the track on the source album.
    set_long!(TrackNumber);

    /// The relative volume adjustment of the track (-100% to 100%).
    get_long!(VolumeAdjustment);

    /// The relative volume adjustment of the track (-100% to 100%).
    set_long!(VolumeAdjustment);

    /// The year the track was recorded/released.
    get_long!(Year);

    /// The year the track was recorded/released.
    set_long!(Year);

    /// Returns a collection of artwork.
    get_object!(Artwork, IITArtworkCollection);
}

/// IITArtwork Interface
///
/// See the generated [`IITArtwork_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Artwork);

impl Artwork {
    /// Delete this piece of artwork from the track.
    no_args!(Delete);

    /// Replace existing artwork data with new artwork from an image file.
    set_bstr!(SetArtworkFromFile, no_set_prefix);

    /// Save artwork data to an image file.
    set_bstr!(SaveArtworkToFile, no_set_prefix);

    /// The format of the artwork.
    get_enum!(Format, ITArtworkFormat);

    /// True if the artwork was downloaded by iTunes.
    get_bool!(IsDownloadedArtwork);

    /// The description for the artwork.
    get_bstr!(Description);

    /// The description for the artwork.
    set_bstr!(Description);
}

/// IITArtworkCollection Interface
///
/// See the generated [`IITArtworkCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(ArtworkCollection);

impl ArtworkCollection {
    /// Returns the number of pieces of artwork in the collection.
    get_long!(Count);

    /// Returns an IITArtwork object corresponding to the given index (1-based).
    item!(IITArtwork);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITSourceCollection Interface
///
/// See the generated [`IITSourceCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(SourceCollection);

impl SourceCollection {
    /// Returns the number of sources in the collection.
    get_long!(Count);

    /// Returns an IITSource object corresponding to the given index (1-based).
    item!(IITSource);

    /// Returns an IITSource object with the specified name.
    item_by_name!(IITSource);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns an IITSource object with the specified persistent ID.
    item_by_persistent_id!(IITSource);
}

/// IITEncoder Interface
///
/// See the generated [`IITEncoder_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Encoder);

impl Encoder {
    /// The name of the the encoder.
    get_bstr!(Name);

    /// The data format created by the encoder.
    get_bstr!(Format);
}

/// IITEncoderCollection Interface
///
/// See the generated [`IITEncoderCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(EncoderCollection);

impl EncoderCollection {
    /// Returns the number of encoders in the collection.
    get_long!(Count);

    /// Returns an IITEncoder object corresponding to the given index (1-based).
    item!(IITEncoder);

    /// Returns an IITEncoder object with the specified name.
    item_by_name!(IITEncoder);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITEQPreset Interface
///
/// See the generated [`IITEQPreset_Impl`] trait for more documentation about each function.
com_wrapper_struct!(EQPreset);

impl EQPreset {
    /// The name of the the EQ preset.
    get_bstr!(Name);

    /// True if this EQ preset can be modified.
    get_bool!(Modifiable);

    /// The equalizer preamp level (-12.0 db to +12.0 db).
    get_double!(Preamp);

    /// The equalizer preamp level (-12.0 db to +12.0 db).
    set_double!(Preamp);

    /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
    get_double!(Band1);

    /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
    set_double!(Band1);

    /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
    get_double!(Band2);

    /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
    set_double!(Band2);

    /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
    get_double!(Band3);

    /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
    set_double!(Band3);

    /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
    get_double!(Band4);

    /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
    set_double!(Band4);

    /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
    get_double!(Band5);

    /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
    set_double!(Band5);

    /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
    get_double!(Band6);

    /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
    set_double!(Band6);

    /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
    get_double!(Band7);

    /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
    set_double!(Band7);

    /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
    get_double!(Band8);

    /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
    set_double!(Band8);

    /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
    get_double!(Band9);

    /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
    set_double!(Band9);

    /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
    get_double!(Band10);

    /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
    set_double!(Band10);

    /// Delete this EQ preset.
    set_bool!(Delete, updateAllTracks, no_set_prefix);

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
    /// Returns the number of EQ presets in the collection.
    get_long!(Count);

    /// Returns an IITEQPreset object corresponding to the given index (1-based).
    item!(IITEQPreset);

    /// Returns an IITEQPreset object with the specified name.
    item_by_name!(IITEQPreset);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITOperationStatus Interface
///
/// See the generated [`IITOperationStatus_Impl`] trait for more documentation about each function.
com_wrapper_struct!(OperationStatus);

impl OperationStatus {
    /// True if the operation is still in progress.
    get_bool!(InProgress);

    /// Returns a collection containing the tracks that were generated by the operation.
    get_object!(Tracks, IITTrackCollection);
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
    no_args!(StopConversion);

    /// Returns the name of the track currently being converted.
    get_bstr!(trackName);

    /// Returns the current progress value for the track being converted.
    get_long!(progressValue);

    /// Returns the maximum progress value for the track being converted.
    get_long!(maxProgressValue);
}

/// IITLibraryPlaylist Interface
///
/// See the generated [`IITLibraryPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(LibraryPlaylist);

impl LibraryPlaylist {
    /// Add the specified file path to the library.
    pub fn AddFile(&self, filePath: BSTR, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Add the specified array of file paths to the library. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub fn AddFiles(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Add the specified streaming audio URL to the library.
    pub fn AddURL(&self, URL: BSTR, iURLTrack: *mut Option<IITURLTrack>) -> windows::core::Result<()> {
        todo!()
    }
    /// Add the specified track to the library.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub fn AddTrack(&self, iTrackToAdd: *const VARIANT, iAddedTrack: *mut Option<IITTrack>) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITURLTrack Interface
///
/// See the generated [`IITURLTrack_Impl`] trait for more documentation about each function.
com_wrapper_struct!(URLTrack);

impl URLTrack {
    /// The URL of the stream represented by this track.
    get_bstr!(URL);

    /// The URL of the stream represented by this track.
    set_bstr!(URL);

    /// True if this is a podcast track.
    get_bool!(Podcast);

    /// Update the podcast feed for this track.
    no_args!(UpdatePodcastFeed);

    /// Start downloading the podcast episode that corresponds to this track.
    no_args!(DownloadPodcastEpisode);

    /// Category for the track.
    get_bstr!(Category);

    /// Category for the track.
    set_bstr!(Category);

    /// Description for the track.
    get_bstr!(Description);

    /// Description for the track.
    set_bstr!(Description);

    /// Long description for the track.
    get_bstr!(LongDescription);

    /// Long description for the track.
    set_bstr!(LongDescription);

    /// Reveal the track in the main browser window.
    no_args!(Reveal);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    get_long!(AlbumRating);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    set_long!(AlbumRating);

    /// The album rating kind.
    get_enum!(AlbumRatingKind, ITRatingKind);

    /// The track rating kind.
    get_enum!(ratingKind, ITRatingKind);

    /// Returns a collection of playlists that contain the song that this track represents.
    get_object!(Playlists, IITPlaylistCollection);
}

/// IITUserPlaylist Interface
///
/// See the generated [`IITUserPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(UserPlaylist);

impl UserPlaylist {
    /// Add the specified file path to the user playlist.
    pub fn AddFile(&self, filePath: BSTR, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Add the specified array of file paths to the user playlist. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub fn AddFiles(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Add the specified streaming audio URL to the user playlist.
    pub fn AddURL(&self, URL: BSTR, iURLTrack: *mut Option<IITURLTrack>) -> windows::core::Result<()> {
        todo!()
    }
    /// Add the specified track to the user playlist.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub fn AddTrack(&self, iTrackToAdd: *const VARIANT, iAddedTrack: *mut Option<IITTrack>) -> windows::core::Result<()> {
        todo!()
    }
    /// True if the user playlist is being shared.
    get_bool!(Shared);

    /// True if the user playlist is being shared.
    set_bool!(Shared);

    /// True if this is a smart playlist.
    get_bool!(Smart);

    /// The playlist special kind.
    get_enum!(SpecialKind, ITUserPlaylistSpecialKind);

    /// The parent of this playlist.
    get_object!(Parent, IITUserPlaylist);

    /// Creates a new playlist in a folder playlist.
    pub fn CreatePlaylist(&self, playlistName: BSTR, iPlaylist: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// Creates a new folder in a folder playlist.
    pub fn CreateFolder(&self, folderName: BSTR, iFolder: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// The parent of this playlist.
    pub fn set_Parent(&self, iParentPlayList: *const VARIANT) -> windows::core::Result<()> {
        todo!()
    }
    /// Reveal the user playlist in the main browser window.
    no_args!(Reveal);
}

/// IITVisual Interface
///
/// See the generated [`IITVisual_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Visual);

impl Visual {
    /// The name of the the visual plug-in.
    get_bstr!(Name);
}

/// IITVisualCollection Interface
///
/// See the generated [`IITVisualCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(VisualCollection);

impl VisualCollection {
    /// Returns the number of visual plug-ins in the collection.
    get_long!(Count);

    /// Returns an IITVisual object corresponding to the given index (1-based).
    item!(IITVisual);

    /// Returns an IITVisual object with the specified name.
    item_by_name!(IITVisual);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITWindow Interface
///
/// See the generated [`IITWindow_Impl`] trait for more documentation about each function.
com_wrapper_struct!(Window);

impl Window {
    /// The title of the window.
    get_bstr!(Name);

    /// The window kind.
    get_enum!(Kind, ITWindowKind);

    /// True if the window is visible. Note that the main browser window cannot be hidden.
    get_bool!(Visible);

    /// True if the window is visible. Note that the main browser window cannot be hidden.
    set_bool!(Visible);

    /// True if the window is resizable.
    get_bool!(Resizable);

    /// True if the window is minimized.
    get_bool!(Minimized);

    /// True if the window is minimized.
    set_bool!(Minimized);

    /// True if the window is maximizable.
    get_bool!(Maximizable);

    /// True if the window is maximized.
    get_bool!(Maximized);

    /// True if the window is maximized.
    set_bool!(Maximized);

    /// True if the window is zoomable.
    get_bool!(Zoomable);

    /// True if the window is zoomed.
    get_bool!(Zoomed);

    /// True if the window is zoomed.
    set_bool!(Zoomed);

    /// The screen coordinate of the top edge of the window.
    get_long!(Top);

    /// The screen coordinate of the top edge of the window.
    set_long!(Top);

    /// The screen coordinate of the left edge of the window.
    get_long!(Left);

    /// The screen coordinate of the left edge of the window.
    set_long!(Left);

    /// The screen coordinate of the bottom edge of the window.
    get_long!(Bottom);

    /// The screen coordinate of the bottom edge of the window.
    set_long!(Bottom);

    /// The screen coordinate of the right edge of the window.
    get_long!(Right);

    /// The screen coordinate of the right edge of the window.
    set_long!(Right);

    /// The width of the window.
    get_long!(Width);

    /// The width of the window.
    set_long!(Width);

    /// The height of the window.
    get_long!(Height);

    /// The height of the window.
    set_long!(Height);
}

/// IITBrowserWindow Interface
///
/// See the generated [`IITBrowserWindow_Impl`] trait for more documentation about each function.
com_wrapper_struct!(BrowserWindow);

impl BrowserWindow {
    /// True if window is in MiniPlayer mode.
    get_bool!(MiniPlayer);

    /// True if window is in MiniPlayer mode.
    set_bool!(MiniPlayer);

    /// Returns a collection containing the currently selected track or tracks.
    get_object!(SelectedTracks, IITTrackCollection);

    /// The currently selected playlist in the Source list.
    get_object!(SelectedPlaylist, IITPlaylist);

    /// The currently selected playlist in the Source list.
    pub fn set_SelectedPlaylist(&self, iPlaylist: *const VARIANT) -> windows::core::Result<()> {
        todo!()
    }
}

/// IITWindowCollection Interface
///
/// See the generated [`IITWindowCollection_Impl`] trait for more documentation about each function.
com_wrapper_struct!(WindowCollection);

impl WindowCollection {
    /// Returns the number of windows in the collection.
    get_long!(Count);

    /// Returns an IITWindow object corresponding to the given index (1-based).
    item!(IITWindow);

    /// Returns an IITWindow object with the specified name.
    item_by_name!(IITWindow);

    /// Returns an IEnumVARIANT object which can enumerate the collection.
    ///
    /// Note: I have not figured out how to use it (calling `.Skip(1)` on the returned `IEnumVARIANT` causes a `STATUS_ACCESS_VIOLATION`).<br/>
    /// Feel free to open an issue or a pull request to fix this.
    pub fn _NewEnum(&self, iEnumerator: *mut Option<IEnumVARIANT>) -> windows::core::Result<()> {
        todo!()
    }
}

/// IiTunes Interface
///
/// See the generated [`IiTunes_Impl`] trait for more documentation about each function.
pub struct iTunes {
    com_object: crate::com::IiTunes,
}

impl iTunes {
    pub fn new() {}

    /// Reposition to the beginning of the current track or go to the previous track if already at start of current track.
    no_args!(BackTrack);

    /// Skip forward in a playing track.
    no_args!(FastForward);

    /// Advance to the next track in the current playlist.
    no_args!(NextTrack);

    /// Pause playback.
    no_args!(Pause);

    /// Play the currently targeted track.
    no_args!(Play);

    /// Play the specified file path, adding it to the library if not already present.
    set_bstr!(PlayFile, no_set_prefix);

    /// Toggle the playing/paused state of the current track.
    no_args!(PlayPause);

    /// Return to the previous track in the current playlist.
    no_args!(PreviousTrack);

    /// Disable fast forward/rewind and resume playback, if playing.
    no_args!(Resume);

    /// Skip backwards in a playing track.
    no_args!(Rewind);

    /// Stop playback.
    no_args!(Stop);

    /// Start converting the specified file path.
    pub fn ConvertFile(&self, filePath: BSTR, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub fn ConvertFiles(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub fn ConvertTrack(&self, iTrackToConvert: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
    pub fn ConvertTracks(&self, iTracksToConvert: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns true if this version of the iTunes type library is compatible with the specified version.
    pub fn CheckVersion(&self, majorVersion: LONG, minorVersion: LONG, isCompatible: *mut VARIANT_BOOL) -> windows::core::Result<()> {
        todo!()
    }
    /// Returns an IITObject corresponding to the specified IDs.
    pub fn GetITObjectByID(&self, sourceID: LONG, playlistID: LONG, trackID: LONG, databaseID: LONG, iObject: *mut Option<IITObject>) -> windows::core::Result<()> {
        todo!()
    }
    /// Creates a new playlist in the main library.
    pub fn CreatePlaylist(&self, playlistName: BSTR, iPlaylist: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// Open the specified iTunes Store or streaming audio URL.
    set_bstr!(OpenURL, no_set_prefix);

    /// Go to the iTunes Store home page.
    no_args!(GotoMusicStoreHomePage);

    /// Update the contents of the iPod.
    no_args!(UpdateIPod);

    /// [id(0x60020015)]
    /// (no other documentation provided)
    pub fn Authorize(&self, numElems: LONG, data: *const VARIANT, names: *const BSTR) -> windows::core::Result<()> {
        todo!()
    }
    /// Exits the iTunes application.
    no_args!(Quit);

    /// Returns a collection of music sources (music library, CD, device, etc.).
    get_object!(Sources, IITSourceCollection);

    /// Returns a collection of encoders.
    get_object!(Encoders, IITEncoderCollection);

    /// Returns a collection of EQ presets.
    get_object!(EQPresets, IITEQPresetCollection);

    /// Returns a collection of visual plug-ins.
    get_object!(Visuals, IITVisualCollection);

    /// Returns a collection of windows.
    get_object!(Windows, IITWindowCollection);

    /// Returns the sound output volume (0 = minimum, 100 = maximum).
    get_long!(SoundVolume);

    /// Returns the sound output volume (0 = minimum, 100 = maximum).
    set_long!(SoundVolume);

    /// True if sound output is muted.
    get_bool!(Mute);

    /// True if sound output is muted.
    set_bool!(Mute);

    /// Returns the current player state.
    get_enum!(PlayerState, ITPlayerState);

    /// Returns the player's position within the currently playing track in seconds.
    get_long!(PlayerPosition);

    /// Returns the player's position within the currently playing track in seconds.
    set_long!(PlayerPosition);

    /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
    get_object!(CurrentEncoder, IITEncoder);

    /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
    set_object!(CurrentEncoder, IITEncoder);

    /// True if visuals are currently being displayed.
    get_bool!(VisualsEnabled);

    /// True if visuals are currently being displayed.
    set_bool!(VisualsEnabled);

    /// True if the visuals are displayed using the entire screen.
    get_bool!(FullScreenVisuals);

    /// True if the visuals are displayed using the entire screen.
    set_bool!(FullScreenVisuals);

    /// Returns the size of the displayed visual.
    get_enum!(VisualSize, ITVisualSize);

    /// Returns the size of the displayed visual.
    set_enum!(VisualSize, ITVisualSize);

    /// Returns the currently selected visual plug-in.
    get_object!(CurrentVisual, IITVisual);

    /// Returns the currently selected visual plug-in.
    set_object!(CurrentVisual, IITVisual);

    /// True if the equalizer is enabled.
    get_bool!(EQEnabled);

    /// True if the equalizer is enabled.
    set_bool!(EQEnabled);

    /// Returns the currently selected EQ preset.
    get_object!(CurrentEQPreset, IITEQPreset);

    /// Returns the currently selected EQ preset.
    set_object!(CurrentEQPreset, IITEQPreset);

    /// The name of the current song in the playing stream (provided by streaming server).
    get_bstr!(CurrentStreamTitle);

    /// The URL of the playing stream or streaming web site (provided by streaming server).
    get_bstr!(set_CurrentStreamURL);

    /// Returns the main iTunes browser window.
    get_object!(BrowserWindow, IITBrowserWindow);

    /// Returns the EQ window.
    get_object!(EQWindow, IITWindow);

    /// Returns the source that represents the main library.
    get_object!(LibrarySource, IITSource);

    /// Returns the main library playlist in the main library source.
    get_object!(LibraryPlaylist, IITLibraryPlaylist);

    /// Returns the currently targeted track.
    get_object!(CurrentTrack, IITTrack);

    /// Returns the playlist containing the currently targeted track.
    get_object!(CurrentPlaylist, IITPlaylist);

    /// Returns a collection containing the currently selected track or tracks.
    get_object!(SelectedTracks, IITTrackCollection);

    /// Returns the version of the iTunes application.
    get_bstr!(Version);

    /// [id(0x6002003b)]
    /// (no other documentation provided)
    set_long!(SetOptions, no_set_prefix);

    /// Start converting the specified file path.
    pub fn ConvertFile2(&self, filePath: BSTR, iStatus: *mut Option<IITConvertOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub fn ConvertFiles2(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITConvertOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub fn ConvertTrack2(&self, iTrackToConvert: *const VARIANT, iStatus: *mut Option<IITConvertOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
    pub fn ConvertTracks2(&self, iTracksToConvert: *const VARIANT, iStatus: *mut Option<IITConvertOperationStatus>) -> windows::core::Result<()> {
        todo!()
    }
    /// True if iTunes will process APPCOMMAND Windows messages.
    get_bool!(AppCommandMessageProcessingEnabled);

    /// True if iTunes will process APPCOMMAND Windows messages.
    set_bool!(AppCommandMessageProcessingEnabled);

    /// True if iTunes will force itself to be the foreground application when it displays a dialog.
    get_bool!(ForceToForegroundOnDialog);

    /// True if iTunes will force itself to be the foreground application when it displays a dialog.
    set_bool!(ForceToForegroundOnDialog);

    /// Create a new EQ preset.
    pub fn CreateEQPreset(&self, eqPresetName: BSTR, iEQPreset: *mut Option<IITEQPreset>) -> windows::core::Result<()> {
        todo!()
    }
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
    get_object!(ConvertOperationStatus, IITConvertOperationStatus);

    /// Subscribe to the specified podcast feed URL.
    set_bstr!(SubscribeToPodcast, no_set_prefix);

    /// Update all podcast feeds.
    no_args!(UpdatePodcastFeeds);

    /// Creates a new folder in the main library.
    pub fn CreateFolder(&self, folderName: BSTR, iFolder: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// Creates a new folder in an existing source.
    pub fn CreateFolderInSource(&self, folderName: BSTR, iSource: *const VARIANT, iFolder: *mut Option<IITPlaylist>) -> windows::core::Result<()> {
        todo!()
    }
    /// True if the sound volume control is enabled.
    get_bool!(SoundVolumeControlEnabled);

    /// The full path to the current iTunes library XML file.
    get_bstr!(LibraryXMLPath);

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
    get_long!(PlayerPositionMS);

    /// Returns the player's position within the currently playing track in milliseconds.
    set_long!(PlayerPositionMS);
}

/// IITAudioCDPlaylist Interface
///
/// See the generated [`IITAudioCDPlaylist_Impl`] trait for more documentation about each function.
com_wrapper_struct!(AudioCDPlaylist);

impl AudioCDPlaylist {
    /// The artist of the CD.
    get_bstr!(Artist);

    /// True if this CD is a compilation album.
    get_bool!(Compilation);

    /// The composer of the CD.
    get_bstr!(Composer);

    /// The total number of discs in this CD's album.
    get_long!(DiscCount);

    /// The index of the CD disc in the source album.
    get_long!(DiscNumber);

    /// The genre of the CD.
    get_bstr!(Genre);

    /// The year the album was recorded/released.
    get_long!(Year);

    /// Reveal the CD playlist in the main browser window.
    no_args!(Reveal);
}

/// IITIPodSource Interface
///
/// See the generated [`IITIPodSource_Impl`] trait for more documentation about each function.
com_wrapper_struct!(IPodSource);

impl IPodSource {
    /// Update the contents of the iPod.
    no_args!(UpdateIPod);

    /// Eject the iPod.
    no_args!(EjectIPod);

    /// The iPod software version.
    get_bstr!(SoftwareVersion);
}

/// IITFileOrCDTrack Interface
///
/// See the generated [`IITFileOrCDTrack_Impl`] trait for more documentation about each function.
com_wrapper_struct!(FileOrCDTrack);

impl FileOrCDTrack {
    /// The full path to the file represented by this track.
    get_bstr!(Location);

    /// Update this track's information with the information stored in its file.
    no_args!(UpdateInfoFromFile);

    /// True if this is a podcast track.
    get_bool!(Podcast);

    /// Update the podcast feed for this track.
    no_args!(UpdatePodcastFeed);

    /// True if playback position is remembered.
    get_bool!(RememberBookmark);

    /// True if playback position is remembered.
    set_bool!(RememberBookmark);

    /// True if track is skipped when shuffling.
    get_bool!(ExcludeFromShuffle);

    /// True if track is skipped when shuffling.
    set_bool!(ExcludeFromShuffle);

    /// Lyrics for the track.
    get_bstr!(Lyrics);

    /// Lyrics for the track.
    set_bstr!(Lyrics);

    /// Category for the track.
    get_bstr!(Category);

    /// Category for the track.
    set_bstr!(Category);

    /// Description for the track.
    get_bstr!(Description);

    /// Description for the track.
    set_bstr!(Description);

    /// Long description for the track.
    get_bstr!(LongDescription);

    /// Long description for the track.
    set_bstr!(LongDescription);

    /// The bookmark time of the track (in seconds).
    get_long!(BookmarkTime);

    /// The bookmark time of the track (in seconds).
    set_long!(BookmarkTime);

    /// The video track kind.
    get_enum!(VideoKind, ITVideoKind);

    /// The video track kind.
    set_enum!(VideoKind, ITVideoKind);

    /// The number of times the track has been skipped.
    get_long!(SkippedCount);

    /// The number of times the track has been skipped.
    set_long!(SkippedCount);

    /// The date and time the track was last skipped.  A value of zero means no skipped date.
    get_date!(SkippedDate);

    /// The date and time the track was last skipped.  A value of zero means no skipped date.
    set_date!(SkippedDate);

    /// True if track is part of a gapless album.
    get_bool!(PartOfGaplessAlbum);

    /// True if track is part of a gapless album.
    set_bool!(PartOfGaplessAlbum);

    /// The album artist of the track.
    get_bstr!(AlbumArtist);

    /// The album artist of the track.
    set_bstr!(AlbumArtist);

    /// The show name of the track.
    get_bstr!(Show);

    /// The show name of the track.
    set_bstr!(Show);

    /// The season number of the track.
    get_long!(SeasonNumber);

    /// The season number of the track.
    set_long!(SeasonNumber);

    /// The episode ID of the track.
    get_bstr!(EpisodeID);

    /// The episode ID of the track.
    set_bstr!(EpisodeID);

    /// The episode number of the track.
    get_long!(EpisodeNumber);

    /// The episode number of the track.
    set_long!(EpisodeNumber);

    /// The high 32-bits of the size of the track (in bytes).
    get_long!(Size64High);

    /// The low 32-bits of the size of the track (in bytes).
    get_long!(Size64Low);

    /// True if track has not been played.
    get_bool!(Unplayed);

    /// True if track has not been played.
    set_bool!(Unplayed);

    /// The album used for sorting.
    get_bstr!(SortAlbum);

    /// The album used for sorting.
    set_bstr!(SortAlbum);

    /// The album artist used for sorting.
    get_bstr!(SortAlbumArtist);

    /// The album artist used for sorting.
    set_bstr!(SortAlbumArtist);

    /// The artist used for sorting.
    get_bstr!(SortArtist);

    /// The artist used for sorting.
    set_bstr!(SortArtist);

    /// The composer used for sorting.
    get_bstr!(SortComposer);

    /// The composer used for sorting.
    set_bstr!(SortComposer);

    /// The track name used for sorting.
    get_bstr!(SortName);

    /// The track name used for sorting.
    set_bstr!(SortName);

    /// The show name used for sorting.
    get_bstr!(SortShow);

    /// The show name used for sorting.
    set_bstr!(SortShow);

    /// Reveal the track in the main browser window.
    no_args!(Reveal);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    get_long!(AlbumRating);

    /// The user or computed rating of the album that this track belongs to (0 to 100).
    set_long!(AlbumRating);

    /// The album rating kind.
    get_enum!(AlbumRatingKind, ITRatingKind);

    /// The track rating kind.
    get_enum!(ratingKind, ITRatingKind);

    /// Returns a collection of playlists that contain the song that this track represents.
    get_object!(Playlists, IITPlaylistCollection);

    /// The full path to the file represented by this track.
    set_bstr!(Location);

    /// The release date of the track.  A value of zero means no release date.
    get_date!(ReleaseDate);
}

/// IITPlaylistWindow Interface
///
/// See the generated [`IITPlaylistWindow_Impl`] trait for more documentation about each function.
com_wrapper_struct!(PlaylistWindow);

impl PlaylistWindow {
    /// Returns a collection containing the currently selected track or tracks.
    get_object!(SelectedTracks, IITTrackCollection);

    /// The playlist displayed in the window.
    get_object!(Playlist, IITPlaylist);
}
