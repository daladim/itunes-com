#![allow(non_snake_case)]

// These types are in the public API.
// We'd rather use the re-exported versions, so that they are available to our users.
use crate::HRESULT;
use crate::BSTR;
use crate::VARIANT_BOOL;

use windows::core::IUnknown;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Com::IDispatch_Impl;
use windows::Win32::System::Com::IDispatch_Vtbl;
use windows::core::GUID;

use windows::Win32::System::Com::VARIANT;

use super::com_enums::*;

type DATE = f64; // This type must be a joke. https://learn.microsoft.com/en-us/cpp/atl-mfc-shared/date-type?view=msvc-170
type LONG = i32;
type DOUBLE = f64;

/// IITObject Interface
///
/// See the generated [`IITObject_Impl`] trait for more documentation about each function.
#[windows::core::interface("9FAB0E27-70D7-4E3A-9965-B0C8B8869BB6")]
pub unsafe trait IITObject : IDispatch {
    /// Returns the four IDs that uniquely identify this object.
    pub unsafe fn GetITObjectIDs(&self, sourceID: *mut LONG, playlistID: *mut LONG, trackID: *mut LONG, databaseID: *mut LONG) -> HRESULT;
    /// The name of the object.
    pub unsafe fn Name(&self, Name: *mut BSTR) -> HRESULT;
    /// The name of the object.
    pub unsafe fn set_Name(&self, Name: BSTR) -> HRESULT;
    /// The index of the object in internal application order (1-based).
    pub unsafe fn Index(&self, Index: *mut LONG) -> HRESULT;
    /// The source ID of the object.
    pub unsafe fn sourceID(&self, sourceID: *mut LONG) -> HRESULT;
    /// The playlist ID of the object.
    pub unsafe fn playlistID(&self, playlistID: *mut LONG) -> HRESULT;
    /// The track ID of the object.
    pub unsafe fn trackID(&self, trackID: *mut LONG) -> HRESULT;
    /// The track database ID of the object.
    pub unsafe fn TrackDatabaseID(&self, databaseID: *mut LONG) -> HRESULT;
}

/// IITSource Interface
///
/// See the generated [`IITSource_Impl`] trait for more documentation about each function.
#[windows::core::interface("AEC1C4D3-AEF1-4255-B892-3E3D13ADFDF9")]
pub unsafe trait IITSource : IITObject {
    /// The source kind.
    pub unsafe fn Kind(&self, Kind: *mut ITSourceKind) -> HRESULT;
    /// The total size of the source, if it has a fixed size.
    pub unsafe fn Capacity(&self, Capacity: *mut DOUBLE) -> HRESULT;
    /// The free space on the source, if it has a fixed size.
    pub unsafe fn FreeSpace(&self, FreeSpace: *mut DOUBLE) -> HRESULT;
    /// Returns a collection of playlists.
    pub unsafe fn Playlists(&self, iPlaylistCollection: *mut Option<IITPlaylistCollection>) -> HRESULT;
}

/// IITPlaylistCollection Interface
///
/// See the generated [`IITPlaylistCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("FF194254-909D-4437-9C50-3AAC2AE6305C")]
pub unsafe trait IITPlaylistCollection : IDispatch {
    /// Returns the number of playlists in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITPlaylist object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// Returns an IITPlaylist object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
    /// Returns an IITPlaylist object with the specified persistent ID.
    pub unsafe fn ItemByPersistentID(&self, highID: LONG, lowID: LONG, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
}

/// IITPlaylist Interface
///
/// See the generated [`IITPlaylist_Impl`] trait for more documentation about each function.
#[windows::core::interface("3D5E072F-2A77-4B17-9E73-E03B77CCCCA9")]
pub unsafe trait IITPlaylist : IITObject {
    /// Delete this playlist.
    pub unsafe fn Delete(&self) -> HRESULT;
    /// Start playing the first track in this playlist.
    pub unsafe fn PlayFirstTrack(&self) -> HRESULT;
    /// Print this playlist.
    pub unsafe fn Print(&self, showPrintDialog: VARIANT_BOOL, printKind: ITPlaylistPrintKind, theme: BSTR) -> HRESULT;
    /// Search tracks in this playlist for the specified string.
    pub unsafe fn Search(&self, searchText: BSTR, searchFields: ITPlaylistSearchField, iTrackCollection: *mut Option<IITTrackCollection>) -> HRESULT;
    /// The playlist kind.
    pub unsafe fn Kind(&self, Kind: *mut ITPlaylistKind) -> HRESULT;
    /// The source that contains this playlist.
    pub unsafe fn Source(&self, iSource: *mut Option<IITSource>) -> HRESULT;
    /// The total length of all songs in the playlist (in seconds).
    pub unsafe fn Duration(&self, Duration: *mut LONG) -> HRESULT;
    /// True if songs in the playlist are played in random order.
    pub unsafe fn Shuffle(&self, isShuffle: *mut VARIANT_BOOL) -> HRESULT;
    /// True if songs in the playlist are played in random order.
    pub unsafe fn set_Shuffle(&self, isShuffle: VARIANT_BOOL) -> HRESULT;
    /// The total size of all songs in the playlist (in bytes).
    pub unsafe fn Size(&self, Size: *mut DOUBLE) -> HRESULT;
    /// The playback repeat mode.
    pub unsafe fn SongRepeat(&self, repeatMode: *mut ITPlaylistRepeatMode) -> HRESULT;
    /// The playback repeat mode.
    pub unsafe fn set_SongRepeat(&self, repeatMode: ITPlaylistRepeatMode) -> HRESULT;
    /// The total length of all songs in the playlist (in MM:SS format).
    pub unsafe fn Time(&self, Time: *mut BSTR) -> HRESULT;
    /// True if the playlist is visible in the Source list.
    pub unsafe fn Visible(&self, isVisible: *mut VARIANT_BOOL) -> HRESULT;
    /// Returns a collection of tracks in this playlist.
    pub unsafe fn Tracks(&self, iTrackCollection: *mut Option<IITTrackCollection>) -> HRESULT;
}

/// IITTrackCollection Interface
///
/// See the generated [`IITTrackCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("755D76F1-6B85-4CE4-8F5F-F88D9743DCD8")]
pub unsafe trait IITTrackCollection : IDispatch {
    /// Returns the number of tracks in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITTrack object corresponding to the given fixed index, where the index is independent of the play order (1-based).
    pub unsafe fn Item(&self, Index: LONG, iTrack: *mut Option<IITTrack>) -> HRESULT;
    /// Returns an IITTrack object corresponding to the given index, where the index is defined by the play order of the playlist containing the track collection (1-based).
    pub unsafe fn ItemByPlayOrder(&self, Index: LONG, iTrack: *mut Option<IITTrack>) -> HRESULT;
    /// Returns an IITTrack object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iTrack: *mut Option<IITTrack>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
    /// Returns an IITTrack object with the specified persistent ID.
    pub unsafe fn ItemByPersistentID(&self, highID: LONG, lowID: LONG, iTrack: *mut Option<IITTrack>) -> HRESULT;
}

/// IITTrack Interface
///
/// See the generated [`IITTrack_Impl`] trait for more documentation about each function.
#[windows::core::interface("4CB0915D-1E54-4727-BAF3-CE6CC9A225A1")]
pub unsafe trait IITTrack : IITObject {
    /// Delete this track.
    pub unsafe fn Delete(&self) -> HRESULT;
    /// Start playing this track.
    pub unsafe fn Play(&self) -> HRESULT;
    /// Add artwork from an image file to this track.
    pub unsafe fn AddArtworkFromFile(&self, filePath: BSTR, iArtwork: *mut Option<IITArtwork>) -> HRESULT;
    /// The track kind.
    pub unsafe fn Kind(&self, Kind: *mut ITTrackKind) -> HRESULT;
    /// The playlist that contains this track.
    pub unsafe fn Playlist(&self, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// The album containing the track.
    pub unsafe fn Album(&self, Album: *mut BSTR) -> HRESULT;
    /// The album containing the track.
    pub unsafe fn set_Album(&self, Album: BSTR) -> HRESULT;
    /// The artist/source of the track.
    pub unsafe fn Artist(&self, Artist: *mut BSTR) -> HRESULT;
    /// The artist/source of the track.
    pub unsafe fn set_Artist(&self, Artist: BSTR) -> HRESULT;
    /// The bit rate of the track (in kbps).
    pub unsafe fn BitRate(&self, BitRate: *mut LONG) -> HRESULT;
    /// The tempo of the track (in beats per minute).
    pub unsafe fn BPM(&self, beatsPerMinute: *mut LONG) -> HRESULT;
    /// The tempo of the track (in beats per minute).
    pub unsafe fn set_BPM(&self, beatsPerMinute: LONG) -> HRESULT;
    /// Freeform notes about the track.
    pub unsafe fn Comment(&self, Comment: *mut BSTR) -> HRESULT;
    /// Freeform notes about the track.
    pub unsafe fn set_Comment(&self, Comment: BSTR) -> HRESULT;
    /// True if this track is from a compilation album.
    pub unsafe fn Compilation(&self, isCompilation: *mut VARIANT_BOOL) -> HRESULT;
    /// True if this track is from a compilation album.
    pub unsafe fn set_Compilation(&self, isCompilation: VARIANT_BOOL) -> HRESULT;
    /// The composer of the track.
    pub unsafe fn Composer(&self, Composer: *mut BSTR) -> HRESULT;
    /// The composer of the track.
    pub unsafe fn set_Composer(&self, Composer: BSTR) -> HRESULT;
    /// The date the track was added to the playlist.
    pub unsafe fn DateAdded(&self, DateAdded: *mut DATE) -> HRESULT;
    /// The total number of discs in the source album.
    pub unsafe fn DiscCount(&self, DiscCount: *mut LONG) -> HRESULT;
    /// The total number of discs in the source album.
    pub unsafe fn set_DiscCount(&self, DiscCount: LONG) -> HRESULT;
    /// The index of the disc containing the track on the source album.
    pub unsafe fn DiscNumber(&self, DiscNumber: *mut LONG) -> HRESULT;
    /// The index of the disc containing the track on the source album.
    pub unsafe fn set_DiscNumber(&self, DiscNumber: LONG) -> HRESULT;
    /// The length of the track (in seconds).
    pub unsafe fn Duration(&self, Duration: *mut LONG) -> HRESULT;
    /// True if the track is checked for playback.
    pub unsafe fn Enabled(&self, isEnabled: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the track is checked for playback.
    pub unsafe fn set_Enabled(&self, isEnabled: VARIANT_BOOL) -> HRESULT;
    /// The name of the EQ preset of the track.
    pub unsafe fn EQ(&self, EQ: *mut BSTR) -> HRESULT;
    /// The name of the EQ preset of the track.
    pub unsafe fn set_EQ(&self, EQ: BSTR) -> HRESULT;
    /// The stop time of the track (in seconds).
    pub unsafe fn set_Finish(&self, Finish: LONG) -> HRESULT;
    /// The stop time of the track (in seconds).
    pub unsafe fn Finish(&self, Finish: *mut LONG) -> HRESULT;
    /// The music/audio genre (category) of the track.
    pub unsafe fn Genre(&self, Genre: *mut BSTR) -> HRESULT;
    /// The music/audio genre (category) of the track.
    pub unsafe fn set_Genre(&self, Genre: BSTR) -> HRESULT;
    /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
    pub unsafe fn Grouping(&self, Grouping: *mut BSTR) -> HRESULT;
    /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
    pub unsafe fn set_Grouping(&self, Grouping: BSTR) -> HRESULT;
    /// A text description of the track.
    pub unsafe fn KindAsString(&self, Kind: *mut BSTR) -> HRESULT;
    /// The modification date of the content of the track.
    pub unsafe fn ModificationDate(&self, dateModified: *mut DATE) -> HRESULT;
    /// The number of times the track has been played.
    pub unsafe fn PlayedCount(&self, PlayedCount: *mut LONG) -> HRESULT;
    /// The number of times the track has been played.
    pub unsafe fn set_PlayedCount(&self, PlayedCount: LONG) -> HRESULT;
    /// The date and time the track was last played.  A value of zero means no played date.
    pub unsafe fn PlayedDate(&self, PlayedDate: *mut DATE) -> HRESULT;
    /// The date and time the track was last played.  A value of zero means no played date.
    pub unsafe fn set_PlayedDate(&self, PlayedDate: DATE) -> HRESULT;
    /// The play order index of the track in the owner playlist (1-based).
    pub unsafe fn PlayOrderIndex(&self, Index: *mut LONG) -> HRESULT;
    /// The rating of the track (0 to 100).
    pub unsafe fn Rating(&self, Rating: *mut LONG) -> HRESULT;
    /// The rating of the track (0 to 100).
    pub unsafe fn set_Rating(&self, Rating: LONG) -> HRESULT;
    /// The sample rate of the track (in Hz).
    pub unsafe fn SampleRate(&self, SampleRate: *mut LONG) -> HRESULT;
    /// The size of the track (in bytes).
    pub unsafe fn Size(&self, Size: *mut LONG) -> HRESULT;
    /// The start time of the track (in seconds).
    pub unsafe fn Start(&self, Start: *mut LONG) -> HRESULT;
    /// The start time of the track (in seconds).
    pub unsafe fn set_Start(&self, Start: LONG) -> HRESULT;
    /// The length of the track (in MM:SS format).
    pub unsafe fn Time(&self, Time: *mut BSTR) -> HRESULT;
    /// The total number of tracks on the source album.
    pub unsafe fn TrackCount(&self, TrackCount: *mut LONG) -> HRESULT;
    /// The total number of tracks on the source album.
    pub unsafe fn set_TrackCount(&self, TrackCount: LONG) -> HRESULT;
    /// The index of the track on the source album.
    pub unsafe fn TrackNumber(&self, TrackNumber: *mut LONG) -> HRESULT;
    /// The index of the track on the source album.
    pub unsafe fn set_TrackNumber(&self, TrackNumber: LONG) -> HRESULT;
    /// The relative volume adjustment of the track (-100% to 100%).
    pub unsafe fn VolumeAdjustment(&self, VolumeAdjustment: *mut LONG) -> HRESULT;
    /// The relative volume adjustment of the track (-100% to 100%).
    pub unsafe fn set_VolumeAdjustment(&self, VolumeAdjustment: LONG) -> HRESULT;
    /// The year the track was recorded/released.
    pub unsafe fn Year(&self, Year: *mut LONG) -> HRESULT;
    /// The year the track was recorded/released.
    pub unsafe fn set_Year(&self, Year: LONG) -> HRESULT;
    /// Returns a collection of artwork.
    pub unsafe fn Artwork(&self, iArtworkCollection: *mut Option<IITArtworkCollection>) -> HRESULT;
}

/// IITArtwork Interface
///
/// See the generated [`IITArtwork_Impl`] trait for more documentation about each function.
#[windows::core::interface("D0A6C1F8-BF3D-4CD8-AC47-FE32BDD17257")]
pub unsafe trait IITArtwork : IDispatch {
    /// Delete this piece of artwork from the track.
    pub unsafe fn Delete(&self) -> HRESULT;
    /// Replace existing artwork data with new artwork from an image file.
    pub unsafe fn SetArtworkFromFile(&self, filePath: BSTR) -> HRESULT;
    /// Save artwork data to an image file.
    pub unsafe fn SaveArtworkToFile(&self, filePath: BSTR) -> HRESULT;
    /// The format of the artwork.
    pub unsafe fn Format(&self, Format: *mut ITArtworkFormat) -> HRESULT;
    /// True if the artwork was downloaded by iTunes.
    pub unsafe fn IsDownloadedArtwork(&self, IsDownloadedArtwork: *mut VARIANT_BOOL) -> HRESULT;
    /// The description for the artwork.
    pub unsafe fn Description(&self, Description: *mut BSTR) -> HRESULT;
    /// The description for the artwork.
    pub unsafe fn set_Description(&self, Description: BSTR) -> HRESULT;
}

/// IITArtworkCollection Interface
///
/// See the generated [`IITArtworkCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("BF2742D7-418C-4858-9AF9-2981B062D23E")]
pub unsafe trait IITArtworkCollection : IDispatch {
    /// Returns the number of pieces of artwork in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITArtwork object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iArtwork: *mut Option<IITArtwork>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
}

/// IITSourceCollection Interface
///
/// See the generated [`IITSourceCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("2FF6CE20-FF87-4183-B0B3-F323D047AF41")]
pub unsafe trait IITSourceCollection : IDispatch {
    /// Returns the number of sources in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITSource object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iSource: *mut Option<IITSource>) -> HRESULT;
    /// Returns an IITSource object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iSource: *mut Option<IITSource>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
    /// Returns an IITSource object with the specified persistent ID.
    pub unsafe fn ItemByPersistentID(&self, highID: LONG, lowID: LONG, iSource: *mut Option<IITSource>) -> HRESULT;
}

/// IITEncoder Interface
///
/// See the generated [`IITEncoder_Impl`] trait for more documentation about each function.
#[windows::core::interface("1CF95A1C-55FE-4F45-A2D3-85AC6C504A73")]
pub unsafe trait IITEncoder : IDispatch {
    /// The name of the the encoder.
    pub unsafe fn Name(&self, Name: *mut BSTR) -> HRESULT;
    /// The data format created by the encoder.
    pub unsafe fn Format(&self, Format: *mut BSTR) -> HRESULT;
}

/// IITEncoderCollection Interface
///
/// See the generated [`IITEncoderCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("8862BCA9-168D-4549-A9D5-ADB35E553BA6")]
pub unsafe trait IITEncoderCollection : IDispatch {
    /// Returns the number of encoders in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITEncoder object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iEncoder: *mut Option<IITEncoder>) -> HRESULT;
    /// Returns an IITEncoder object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iEncoder: *mut Option<IITEncoder>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
}

/// IITEQPreset Interface
///
/// See the generated [`IITEQPreset_Impl`] trait for more documentation about each function.
#[windows::core::interface("5BE75F4F-68FA-4212-ACB7-BE44EA569759")]
pub unsafe trait IITEQPreset : IDispatch {
    /// The name of the the EQ preset.
    pub unsafe fn Name(&self, Name: *mut BSTR) -> HRESULT;
    /// True if this EQ preset can be modified.
    pub unsafe fn Modifiable(&self, isModifiable: *mut VARIANT_BOOL) -> HRESULT;
    /// The equalizer preamp level (-12.0 db to +12.0 db).
    pub unsafe fn Preamp(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer preamp level (-12.0 db to +12.0 db).
    pub unsafe fn set_Preamp(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band1(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band1(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band2(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band2(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band3(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band3(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band4(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band4(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band5(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band5(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band6(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band6(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band7(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band7(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band8(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band8(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band9(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band9(&self, level: DOUBLE) -> HRESULT;
    /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn Band10(&self, level: *mut DOUBLE) -> HRESULT;
    /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
    pub unsafe fn set_Band10(&self, level: DOUBLE) -> HRESULT;
    /// Delete this EQ preset.
    pub unsafe fn Delete(&self, updateAllTracks: VARIANT_BOOL) -> HRESULT;
    /// Rename this EQ preset.
    pub unsafe fn Rename(&self, newName: BSTR, updateAllTracks: VARIANT_BOOL) -> HRESULT;
}

/// IITEQPresetCollection Interface
///
/// See the generated [`IITEQPresetCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("AEF4D111-3331-48DA-B0C2-B468D5D61D08")]
pub unsafe trait IITEQPresetCollection : IDispatch {
    /// Returns the number of EQ presets in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITEQPreset object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iEQPreset: *mut Option<IITEQPreset>) -> HRESULT;
    /// Returns an IITEQPreset object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iEQPreset: *mut Option<IITEQPreset>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
}

/// IITOperationStatus Interface
///
/// See the generated [`IITOperationStatus_Impl`] trait for more documentation about each function.
#[windows::core::interface("206479C9-FE32-4F9B-A18A-475AC939B479")]
pub unsafe trait IITOperationStatus : IDispatch {
    /// True if the operation is still in progress.
    pub unsafe fn InProgress(&self, isInProgress: *mut VARIANT_BOOL) -> HRESULT;
    /// Returns a collection containing the tracks that were generated by the operation.
    pub unsafe fn Tracks(&self, iTrackCollection: *mut Option<IITTrackCollection>) -> HRESULT;
}

/// IITConvertOperationStatus Interface
///
/// See the generated [`IITConvertOperationStatus_Impl`] trait for more documentation about each function.
#[windows::core::interface("7063AAF6-ABA0-493B-B4FC-920A9F105875")]
pub unsafe trait IITConvertOperationStatus : IITOperationStatus {
    /// Returns the current conversion status.
    pub unsafe fn GetConversionStatus(&self, trackName: *mut BSTR, progressValue: *mut LONG, maxProgressValue: *mut LONG) -> HRESULT;
    /// Stops the current conversion operation.
    pub unsafe fn StopConversion(&self) -> HRESULT;
    /// Returns the name of the track currently being converted.
    pub unsafe fn trackName(&self, trackName: *mut BSTR) -> HRESULT;
    /// Returns the current progress value for the track being converted.
    pub unsafe fn progressValue(&self, progressValue: *mut LONG) -> HRESULT;
    /// Returns the maximum progress value for the track being converted.
    pub unsafe fn maxProgressValue(&self, maxProgressValue: *mut LONG) -> HRESULT;
}

/// IITLibraryPlaylist Interface
///
/// See the generated [`IITLibraryPlaylist_Impl`] trait for more documentation about each function.
#[windows::core::interface("53AE1704-491C-4289-94A0-958815675A3D")]
pub unsafe trait IITLibraryPlaylist : IITPlaylist {
    /// Add the specified file path to the library.
    pub unsafe fn AddFile(&self, filePath: BSTR, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Add the specified array of file paths to the library. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub unsafe fn AddFiles(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Add the specified streaming audio URL to the library.
    pub unsafe fn AddURL(&self, URL: BSTR, iURLTrack: *mut Option<IITURLTrack>) -> HRESULT;
    /// Add the specified track to the library.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub unsafe fn AddTrack(&self, iTrackToAdd: *const VARIANT, iAddedTrack: *mut Option<IITTrack>) -> HRESULT;
}

/// IITURLTrack Interface
///
/// See the generated [`IITURLTrack_Impl`] trait for more documentation about each function.
#[windows::core::interface("1116E3B5-29FD-4393-A7BD-454E5E327900")]
pub unsafe trait IITURLTrack : IITTrack {
    /// The URL of the stream represented by this track.
    pub unsafe fn URL(&self, URL: *mut BSTR) -> HRESULT;
    /// The URL of the stream represented by this track.
    pub unsafe fn set_URL(&self, URL: BSTR) -> HRESULT;
    /// True if this is a podcast track.
    pub unsafe fn Podcast(&self, isPodcast: *mut VARIANT_BOOL) -> HRESULT;
    /// Update the podcast feed for this track.
    pub unsafe fn UpdatePodcastFeed(&self) -> HRESULT;
    /// Start downloading the podcast episode that corresponds to this track.
    pub unsafe fn DownloadPodcastEpisode(&self) -> HRESULT;
    /// Category for the track.
    pub unsafe fn Category(&self, Category: *mut BSTR) -> HRESULT;
    /// Category for the track.
    pub unsafe fn set_Category(&self, Category: BSTR) -> HRESULT;
    /// Description for the track.
    pub unsafe fn Description(&self, Description: *mut BSTR) -> HRESULT;
    /// Description for the track.
    pub unsafe fn set_Description(&self, Description: BSTR) -> HRESULT;
    /// Long description for the track.
    pub unsafe fn LongDescription(&self, LongDescription: *mut BSTR) -> HRESULT;
    /// Long description for the track.
    pub unsafe fn set_LongDescription(&self, LongDescription: BSTR) -> HRESULT;
    /// Reveal the track in the main browser window.
    pub unsafe fn Reveal(&self) -> HRESULT;
    /// The user or computed rating of the album that this track belongs to (0 to 100).
    pub unsafe fn AlbumRating(&self, Rating: *mut LONG) -> HRESULT;
    /// The user or computed rating of the album that this track belongs to (0 to 100).
    pub unsafe fn set_AlbumRating(&self, Rating: LONG) -> HRESULT;
    /// The album rating kind.
    pub unsafe fn AlbumRatingKind(&self, ratingKind: *mut ITRatingKind) -> HRESULT;
    /// The track rating kind.
    pub unsafe fn ratingKind(&self, ratingKind: *mut ITRatingKind) -> HRESULT;
    /// Returns a collection of playlists that contain the song that this track represents.
    pub unsafe fn Playlists(&self, iPlaylistCollection: *mut Option<IITPlaylistCollection>) -> HRESULT;
}

/// IITUserPlaylist Interface
///
/// See the generated [`IITUserPlaylist_Impl`] trait for more documentation about each function.
#[windows::core::interface("0A504DED-A0B5-465A-8A94-50E20D7DF692")]
pub unsafe trait IITUserPlaylist : IITPlaylist {
    /// Add the specified file path to the user playlist.
    pub unsafe fn AddFile(&self, filePath: BSTR, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Add the specified array of file paths to the user playlist. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub unsafe fn AddFiles(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Add the specified streaming audio URL to the user playlist.
    pub unsafe fn AddURL(&self, URL: BSTR, iURLTrack: *mut Option<IITURLTrack>) -> HRESULT;
    /// Add the specified track to the user playlist.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub unsafe fn AddTrack(&self, iTrackToAdd: *const VARIANT, iAddedTrack: *mut Option<IITTrack>) -> HRESULT;
    /// True if the user playlist is being shared.
    pub unsafe fn Shared(&self, isShared: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the user playlist is being shared.
    pub unsafe fn set_Shared(&self, isShared: VARIANT_BOOL) -> HRESULT;
    /// True if this is a smart playlist.
    pub unsafe fn Smart(&self, isSmart: *mut VARIANT_BOOL) -> HRESULT;
    /// The playlist special kind.
    pub unsafe fn SpecialKind(&self, SpecialKind: *mut ITUserPlaylistSpecialKind) -> HRESULT;
    /// The parent of this playlist.
    pub unsafe fn Parent(&self, iParentPlayList: *mut Option<IITUserPlaylist>) -> HRESULT;
    /// Creates a new playlist in a folder playlist.
    pub unsafe fn CreatePlaylist(&self, playlistName: BSTR, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// Creates a new folder in a folder playlist.
    pub unsafe fn CreateFolder(&self, folderName: BSTR, iFolder: *mut Option<IITPlaylist>) -> HRESULT;
    /// The parent of this playlist.
    pub unsafe fn set_Parent(&self, iParentPlayList: *const VARIANT) -> HRESULT;
    /// Reveal the user playlist in the main browser window.
    pub unsafe fn Reveal(&self) -> HRESULT;
}

/// IITVisual Interface
///
/// See the generated [`IITVisual_Impl`] trait for more documentation about each function.
#[windows::core::interface("340F3315-ED72-4C09-9ACF-21EB4BDF9931")]
pub unsafe trait IITVisual : IDispatch {
    /// The name of the the visual plug-in.
    pub unsafe fn Name(&self, Name: *mut BSTR) -> HRESULT;
}

/// IITVisualCollection Interface
///
/// See the generated [`IITVisualCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("88A4CCDD-114F-4043-B69B-84D4E6274957")]
pub unsafe trait IITVisualCollection : IDispatch {
    /// Returns the number of visual plug-ins in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITVisual object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iVisual: *mut Option<IITVisual>) -> HRESULT;
    /// Returns an IITVisual object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iVisual: *mut Option<IITVisual>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
}

/// IITWindow Interface
///
/// See the generated [`IITWindow_Impl`] trait for more documentation about each function.
#[windows::core::interface("370D7BE0-3A89-4A42-B902-C75FC138BE09")]
pub unsafe trait IITWindow : IDispatch {
    /// The title of the window.
    pub unsafe fn Name(&self, Name: *mut BSTR) -> HRESULT;
    /// The window kind.
    pub unsafe fn Kind(&self, Kind: *mut ITWindowKind) -> HRESULT;
    /// True if the window is visible. Note that the main browser window cannot be hidden.
    pub unsafe fn Visible(&self, isVisible: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is visible. Note that the main browser window cannot be hidden.
    pub unsafe fn set_Visible(&self, isVisible: VARIANT_BOOL) -> HRESULT;
    /// True if the window is resizable.
    pub unsafe fn Resizable(&self, isResizable: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is minimized.
    pub unsafe fn Minimized(&self, isMinimized: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is minimized.
    pub unsafe fn set_Minimized(&self, isMinimized: VARIANT_BOOL) -> HRESULT;
    /// True if the window is maximizable.
    pub unsafe fn Maximizable(&self, isMaximizable: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is maximized.
    pub unsafe fn Maximized(&self, isMaximized: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is maximized.
    pub unsafe fn set_Maximized(&self, isMaximized: VARIANT_BOOL) -> HRESULT;
    /// True if the window is zoomable.
    pub unsafe fn Zoomable(&self, isZoomable: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is zoomed.
    pub unsafe fn Zoomed(&self, isZoomed: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the window is zoomed.
    pub unsafe fn set_Zoomed(&self, isZoomed: VARIANT_BOOL) -> HRESULT;
    /// The screen coordinate of the top edge of the window.
    pub unsafe fn Top(&self, Top: *mut LONG) -> HRESULT;
    /// The screen coordinate of the top edge of the window.
    pub unsafe fn set_Top(&self, Top: LONG) -> HRESULT;
    /// The screen coordinate of the left edge of the window.
    pub unsafe fn Left(&self, Left: *mut LONG) -> HRESULT;
    /// The screen coordinate of the left edge of the window.
    pub unsafe fn set_Left(&self, Left: LONG) -> HRESULT;
    /// The screen coordinate of the bottom edge of the window.
    pub unsafe fn Bottom(&self, Bottom: *mut LONG) -> HRESULT;
    /// The screen coordinate of the bottom edge of the window.
    pub unsafe fn set_Bottom(&self, Bottom: LONG) -> HRESULT;
    /// The screen coordinate of the right edge of the window.
    pub unsafe fn Right(&self, Right: *mut LONG) -> HRESULT;
    /// The screen coordinate of the right edge of the window.
    pub unsafe fn set_Right(&self, Right: LONG) -> HRESULT;
    /// The width of the window.
    pub unsafe fn Width(&self, Width: *mut LONG) -> HRESULT;
    /// The width of the window.
    pub unsafe fn set_Width(&self, Width: LONG) -> HRESULT;
    /// The height of the window.
    pub unsafe fn Height(&self, Height: *mut LONG) -> HRESULT;
    /// The height of the window.
    pub unsafe fn set_Height(&self, Height: LONG) -> HRESULT;
}

/// IITBrowserWindow Interface
///
/// See the generated [`IITBrowserWindow_Impl`] trait for more documentation about each function.
#[windows::core::interface("C999F455-C4D5-4AA4-8277-F99753699974")]
pub unsafe trait IITBrowserWindow : IITWindow {
    /// True if window is in MiniPlayer mode.
    pub unsafe fn MiniPlayer(&self, isMiniPlayer: *mut VARIANT_BOOL) -> HRESULT;
    /// True if window is in MiniPlayer mode.
    pub unsafe fn set_MiniPlayer(&self, isMiniPlayer: VARIANT_BOOL) -> HRESULT;
    /// Returns a collection containing the currently selected track or tracks.
    pub unsafe fn SelectedTracks(&self, iTrackCollection: *mut Option<IITTrackCollection>) -> HRESULT;
    /// The currently selected playlist in the Source list.
    pub unsafe fn SelectedPlaylist(&self, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// The currently selected playlist in the Source list.
    pub unsafe fn set_SelectedPlaylist(&self, iPlaylist: *const VARIANT) -> HRESULT;
}

/// IITWindowCollection Interface
///
/// See the generated [`IITWindowCollection_Impl`] trait for more documentation about each function.
#[windows::core::interface("3D8DE381-6C0E-481F-A865-E2385F59FA43")]
pub unsafe trait IITWindowCollection : IDispatch {
    /// Returns the number of windows in the collection.
    pub unsafe fn Count(&self, Count: *mut LONG) -> HRESULT;
    /// Returns an IITWindow object corresponding to the given index (1-based).
    pub unsafe fn Item(&self, Index: LONG, iWindow: *mut Option<IITWindow>) -> HRESULT;
    /// Returns an IITWindow object with the specified name.
    pub unsafe fn ItemByName(&self, Name: BSTR, iWindow: *mut Option<IITWindow>) -> HRESULT;
    /// Returns an IEnumVARIANT object which can enumerate the collection.
    pub unsafe fn _NewEnum(&self, iEnumerator: *mut Option<IUnknown>) -> HRESULT;
}

/// IiTunes Interface
///
/// See the generated [`IiTunes_Impl`] trait for more documentation about each function.
#[windows::core::interface("9DD6680B-3EDC-40DB-A771-E6FE4832E34A")]
pub unsafe trait IiTunes : IDispatch {
    /// Reposition to the beginning of the current track or go to the previous track if already at start of current track.
    pub unsafe fn BackTrack(&self) -> HRESULT;
    /// Skip forward in a playing track.
    pub unsafe fn FastForward(&self) -> HRESULT;
    /// Advance to the next track in the current playlist.
    pub unsafe fn NextTrack(&self) -> HRESULT;
    /// Pause playback.
    pub unsafe fn Pause(&self) -> HRESULT;
    /// Play the currently targeted track.
    pub unsafe fn Play(&self) -> HRESULT;
    /// Play the specified file path, adding it to the library if not already present.
    pub unsafe fn PlayFile(&self, filePath: BSTR) -> HRESULT;
    /// Toggle the playing/paused state of the current track.
    pub unsafe fn PlayPause(&self) -> HRESULT;
    /// Return to the previous track in the current playlist.
    pub unsafe fn PreviousTrack(&self) -> HRESULT;
    /// Disable fast forward/rewind and resume playback, if playing.
    pub unsafe fn Resume(&self) -> HRESULT;
    /// Skip backwards in a playing track.
    pub unsafe fn Rewind(&self) -> HRESULT;
    /// Stop playback.
    pub unsafe fn Stop(&self) -> HRESULT;
    /// Start converting the specified file path.
    pub unsafe fn ConvertFile(&self, filePath: BSTR, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub unsafe fn ConvertFiles(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub unsafe fn ConvertTrack(&self, iTrackToConvert: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
    pub unsafe fn ConvertTracks(&self, iTracksToConvert: *const VARIANT, iStatus: *mut Option<IITOperationStatus>) -> HRESULT;
    /// Returns true if this version of the iTunes type library is compatible with the specified version.
    pub unsafe fn CheckVersion(&self, majorVersion: LONG, minorVersion: LONG, isCompatible: *mut VARIANT_BOOL) -> HRESULT;
    /// Returns an IITObject corresponding to the specified IDs.
    pub unsafe fn GetITObjectByID(&self, sourceID: LONG, playlistID: LONG, trackID: LONG, databaseID: LONG, iObject: *mut Option<IITObject>) -> HRESULT;
    /// Creates a new playlist in the main library.
    pub unsafe fn CreatePlaylist(&self, playlistName: BSTR, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// Open the specified iTunes Store or streaming audio URL.
    pub unsafe fn OpenURL(&self, URL: BSTR) -> HRESULT;
    /// Go to the iTunes Store home page.
    pub unsafe fn GotoMusicStoreHomePage(&self) -> HRESULT;
    /// Update the contents of the iPod.
    pub unsafe fn UpdateIPod(&self) -> HRESULT;
    /// [id(0x60020015)]
    /// (no other documentation provided)
    pub unsafe fn Authorize(&self, numElems: LONG, data: *const VARIANT, names: *const BSTR) -> HRESULT;
    /// Exits the iTunes application.
    pub unsafe fn Quit(&self) -> HRESULT;
    /// Returns a collection of music sources (music library, CD, device, etc.).
    pub unsafe fn Sources(&self, iSourceCollection: *mut Option<IITSourceCollection>) -> HRESULT;
    /// Returns a collection of encoders.
    pub unsafe fn Encoders(&self, iEncoderCollection: *mut Option<IITEncoderCollection>) -> HRESULT;
    /// Returns a collection of EQ presets.
    pub unsafe fn EQPresets(&self, iEQPresetCollection: *mut Option<IITEQPresetCollection>) -> HRESULT;
    /// Returns a collection of visual plug-ins.
    pub unsafe fn Visuals(&self, iVisualCollection: *mut Option<IITVisualCollection>) -> HRESULT;
    /// Returns a collection of windows.
    pub unsafe fn Windows(&self, iWindowCollection: *mut Option<IITWindowCollection>) -> HRESULT;
    /// Returns the sound output volume (0 = minimum, 100 = maximum).
    pub unsafe fn SoundVolume(&self, volume: *mut LONG) -> HRESULT;
    /// Returns the sound output volume (0 = minimum, 100 = maximum).
    pub unsafe fn set_SoundVolume(&self, volume: LONG) -> HRESULT;
    /// True if sound output is muted.
    pub unsafe fn Mute(&self, isMuted: *mut VARIANT_BOOL) -> HRESULT;
    /// True if sound output is muted.
    pub unsafe fn set_Mute(&self, isMuted: VARIANT_BOOL) -> HRESULT;
    /// Returns the current player state.
    pub unsafe fn PlayerState(&self, PlayerState: *mut ITPlayerState) -> HRESULT;
    /// Returns the player's position within the currently playing track in seconds.
    pub unsafe fn PlayerPosition(&self, playerPos: *mut LONG) -> HRESULT;
    /// Returns the player's position within the currently playing track in seconds.
    pub unsafe fn set_PlayerPosition(&self, playerPos: LONG) -> HRESULT;
    /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
    pub unsafe fn CurrentEncoder(&self, iEncoder: *mut Option<IITEncoder>) -> HRESULT;
    /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
    pub unsafe fn set_CurrentEncoder(&self, iEncoder: *const IITEncoder) -> HRESULT;
    /// True if visuals are currently being displayed.
    pub unsafe fn VisualsEnabled(&self, isEnabled: *mut VARIANT_BOOL) -> HRESULT;
    /// True if visuals are currently being displayed.
    pub unsafe fn set_VisualsEnabled(&self, isEnabled: VARIANT_BOOL) -> HRESULT;
    /// True if the visuals are displayed using the entire screen.
    pub unsafe fn FullScreenVisuals(&self, isFullScreen: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the visuals are displayed using the entire screen.
    pub unsafe fn set_FullScreenVisuals(&self, isFullScreen: VARIANT_BOOL) -> HRESULT;
    /// Returns the size of the displayed visual.
    pub unsafe fn VisualSize(&self, VisualSize: *mut ITVisualSize) -> HRESULT;
    /// Returns the size of the displayed visual.
    pub unsafe fn set_VisualSize(&self, VisualSize: ITVisualSize) -> HRESULT;
    /// Returns the currently selected visual plug-in.
    pub unsafe fn CurrentVisual(&self, iVisual: *mut Option<IITVisual>) -> HRESULT;
    /// Returns the currently selected visual plug-in.
    pub unsafe fn set_CurrentVisual(&self, iVisual: *const IITVisual) -> HRESULT;
    /// True if the equalizer is enabled.
    pub unsafe fn EQEnabled(&self, isEnabled: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the equalizer is enabled.
    pub unsafe fn set_EQEnabled(&self, isEnabled: VARIANT_BOOL) -> HRESULT;
    /// Returns the currently selected EQ preset.
    pub unsafe fn CurrentEQPreset(&self, iEQPreset: *mut Option<IITEQPreset>) -> HRESULT;
    /// Returns the currently selected EQ preset.
    pub unsafe fn set_CurrentEQPreset(&self, iEQPreset: *const IITEQPreset) -> HRESULT;
    /// The name of the current song in the playing stream (provided by streaming server).
    pub unsafe fn CurrentStreamTitle(&self, streamTitle: *mut BSTR) -> HRESULT;
    /// The URL of the playing stream or streaming web site (provided by streaming server).
    pub unsafe fn set_CurrentStreamURL(&self, streamURL: *mut BSTR) -> HRESULT;
    /// Returns the main iTunes browser window.
    pub unsafe fn BrowserWindow(&self, iBrowserWindow: *mut Option<IITBrowserWindow>) -> HRESULT;
    /// Returns the EQ window.
    pub unsafe fn EQWindow(&self, iEQWindow: *mut Option<IITWindow>) -> HRESULT;
    /// Returns the source that represents the main library.
    pub unsafe fn LibrarySource(&self, iLibrarySource: *mut Option<IITSource>) -> HRESULT;
    /// Returns the main library playlist in the main library source.
    pub unsafe fn LibraryPlaylist(&self, iLibraryPlaylist: *mut Option<IITLibraryPlaylist>) -> HRESULT;
    /// Returns the currently targeted track.
    pub unsafe fn CurrentTrack(&self, iTrack: *mut Option<IITTrack>) -> HRESULT;
    /// Returns the playlist containing the currently targeted track.
    pub unsafe fn CurrentPlaylist(&self, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// Returns a collection containing the currently selected track or tracks.
    pub unsafe fn SelectedTracks(&self, iTrackCollection: *mut Option<IITTrackCollection>) -> HRESULT;
    /// Returns the version of the iTunes application.
    pub unsafe fn Version(&self, Version: *mut BSTR) -> HRESULT;
    /// [id(0x6002003b)]
    /// (no other documentation provided)
    pub unsafe fn SetOptions(&self, options: LONG) -> HRESULT;
    /// Start converting the specified file path.
    pub unsafe fn ConvertFile2(&self, filePath: BSTR, iStatus: *mut Option<IITConvertOperationStatus>) -> HRESULT;
    /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
    pub unsafe fn ConvertFiles2(&self, filePaths: *const VARIANT, iStatus: *mut Option<IITConvertOperationStatus>) -> HRESULT;
    /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
    pub unsafe fn ConvertTrack2(&self, iTrackToConvert: *const VARIANT, iStatus: *mut Option<IITConvertOperationStatus>) -> HRESULT;
    /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
    pub unsafe fn ConvertTracks2(&self, iTracksToConvert: *const VARIANT, iStatus: *mut Option<IITConvertOperationStatus>) -> HRESULT;
    /// True if iTunes will process APPCOMMAND Windows messages.
    pub unsafe fn AppCommandMessageProcessingEnabled(&self, isEnabled: *mut VARIANT_BOOL) -> HRESULT;
    /// True if iTunes will process APPCOMMAND Windows messages.
    pub unsafe fn set_AppCommandMessageProcessingEnabled(&self, isEnabled: VARIANT_BOOL) -> HRESULT;
    /// True if iTunes will force itself to be the foreground application when it displays a dialog.
    pub unsafe fn ForceToForegroundOnDialog(&self, ForceToForegroundOnDialog: *mut VARIANT_BOOL) -> HRESULT;
    /// True if iTunes will force itself to be the foreground application when it displays a dialog.
    pub unsafe fn set_ForceToForegroundOnDialog(&self, ForceToForegroundOnDialog: VARIANT_BOOL) -> HRESULT;
    /// Create a new EQ preset.
    pub unsafe fn CreateEQPreset(&self, eqPresetName: BSTR, iEQPreset: *mut Option<IITEQPreset>) -> HRESULT;
    /// Creates a new playlist in an existing source.
    pub unsafe fn CreatePlaylistInSource(&self, playlistName: BSTR, iSource: *const VARIANT, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
    /// Retrieves the current state of the player buttons.
    pub unsafe fn GetPlayerButtonsState(&self, previousEnabled: *mut VARIANT_BOOL, playPauseStopState: *mut ITPlayButtonState, nextEnabled: *mut VARIANT_BOOL) -> HRESULT;
    /// Simulate click on a player control button.
    pub unsafe fn PlayerButtonClicked(&self, playerButton: ITPlayerButton, playerButtonModifierKeys: LONG) -> HRESULT;
    /// True if the Shuffle property is writable for the specified playlist.
    pub unsafe fn CanSetShuffle(&self, iPlaylist: *const VARIANT, CanSetShuffle: *mut VARIANT_BOOL) -> HRESULT;
    /// True if the SongRepeat property is writable for the specified playlist.
    pub unsafe fn CanSetSongRepeat(&self, iPlaylist: *const VARIANT, CanSetSongRepeat: *mut VARIANT_BOOL) -> HRESULT;
    /// Returns an IITConvertOperationStatus object if there is currently a conversion in progress.
    pub unsafe fn ConvertOperationStatus(&self, iStatus: *mut Option<IITConvertOperationStatus>) -> HRESULT;
    /// Subscribe to the specified podcast feed URL.
    pub unsafe fn SubscribeToPodcast(&self, URL: BSTR) -> HRESULT;
    /// Update all podcast feeds.
    pub unsafe fn UpdatePodcastFeeds(&self) -> HRESULT;
    /// Creates a new folder in the main library.
    pub unsafe fn CreateFolder(&self, folderName: BSTR, iFolder: *mut Option<IITPlaylist>) -> HRESULT;
    /// Creates a new folder in an existing source.
    pub unsafe fn CreateFolderInSource(&self, folderName: BSTR, iSource: *const VARIANT, iFolder: *mut Option<IITPlaylist>) -> HRESULT;
    /// True if the sound volume control is enabled.
    pub unsafe fn SoundVolumeControlEnabled(&self, isEnabled: *mut VARIANT_BOOL) -> HRESULT;
    /// The full path to the current iTunes library XML file.
    pub unsafe fn LibraryXMLPath(&self, filePath: *mut BSTR) -> HRESULT;
    /// Returns the high 32 bits of the persistent ID of the specified IITObject.
    pub unsafe fn ITObjectPersistentIDHigh(&self, iObject: *const VARIANT, highID: *mut LONG) -> HRESULT;
    /// Returns the low 32 bits of the persistent ID of the specified IITObject.
    pub unsafe fn ITObjectPersistentIDLow(&self, iObject: *const VARIANT, lowID: *mut LONG) -> HRESULT;
    /// Returns the high and low 32 bits of the persistent ID of the specified IITObject.
    pub unsafe fn GetITObjectPersistentIDs(&self, iObject: *const VARIANT, highID: *mut LONG, lowID: *mut LONG) -> HRESULT;
    /// Returns the player's position within the currently playing track in milliseconds.
    pub unsafe fn PlayerPositionMS(&self, playerPos: *mut LONG) -> HRESULT;
    /// Returns the player's position within the currently playing track in milliseconds.
    pub unsafe fn set_PlayerPositionMS(&self, playerPos: LONG) -> HRESULT;
}

/// IITAudioCDPlaylist Interface
///
/// See the generated [`IITAudioCDPlaylist_Impl`] trait for more documentation about each function.
#[windows::core::interface("CF496DF3-0FED-4D7D-9BD8-529B6E8A082E")]
pub unsafe trait IITAudioCDPlaylist : IITPlaylist {
    /// The artist of the CD.
    pub unsafe fn Artist(&self, Artist: *mut BSTR) -> HRESULT;
    /// True if this CD is a compilation album.
    pub unsafe fn Compilation(&self, isCompiliation: *mut VARIANT_BOOL) -> HRESULT;
    /// The composer of the CD.
    pub unsafe fn Composer(&self, Composer: *mut BSTR) -> HRESULT;
    /// The total number of discs in this CD's album.
    pub unsafe fn DiscCount(&self, DiscCount: *mut LONG) -> HRESULT;
    /// The index of the CD disc in the source album.
    pub unsafe fn DiscNumber(&self, DiscNumber: *mut LONG) -> HRESULT;
    /// The genre of the CD.
    pub unsafe fn Genre(&self, Genre: *mut BSTR) -> HRESULT;
    /// The year the album was recorded/released.
    pub unsafe fn Year(&self, Year: *mut LONG) -> HRESULT;
    /// Reveal the CD playlist in the main browser window.
    pub unsafe fn Reveal(&self) -> HRESULT;
}

/// IITIPodSource Interface
///
/// See the generated [`IITIPodSource_Impl`] trait for more documentation about each function.
#[windows::core::interface("CF4D8ACE-1720-4FB9-B0AE-9877249E89B0")]
pub unsafe trait IITIPodSource : IITSource {
    /// Update the contents of the iPod.
    pub unsafe fn UpdateIPod(&self) -> HRESULT;
    /// Eject the iPod.
    pub unsafe fn EjectIPod(&self) -> HRESULT;
    /// The iPod software version.
    pub unsafe fn SoftwareVersion(&self, SoftwareVersion: *mut BSTR) -> HRESULT;
}

/// IITFileOrCDTrack Interface
///
/// See the generated [`IITFileOrCDTrack_Impl`] trait for more documentation about each function.
#[windows::core::interface("00D7FE99-7868-4CC7-AD9E-ACFD70D09566")]
pub unsafe trait IITFileOrCDTrack : IITTrack {
    /// The full path to the file represented by this track.
    pub unsafe fn Location(&self, Location: *mut BSTR) -> HRESULT;
    /// Update this track's information with the information stored in its file.
    pub unsafe fn UpdateInfoFromFile(&self) -> HRESULT;
    /// True if this is a podcast track.
    pub unsafe fn Podcast(&self, isPodcast: *mut VARIANT_BOOL) -> HRESULT;
    /// Update the podcast feed for this track.
    pub unsafe fn UpdatePodcastFeed(&self) -> HRESULT;
    /// True if playback position is remembered.
    pub unsafe fn RememberBookmark(&self, RememberBookmark: *mut VARIANT_BOOL) -> HRESULT;
    /// True if playback position is remembered.
    pub unsafe fn set_RememberBookmark(&self, RememberBookmark: VARIANT_BOOL) -> HRESULT;
    /// True if track is skipped when shuffling.
    pub unsafe fn ExcludeFromShuffle(&self, ExcludeFromShuffle: *mut VARIANT_BOOL) -> HRESULT;
    /// True if track is skipped when shuffling.
    pub unsafe fn set_ExcludeFromShuffle(&self, ExcludeFromShuffle: VARIANT_BOOL) -> HRESULT;
    /// Lyrics for the track.
    pub unsafe fn Lyrics(&self, Lyrics: *mut BSTR) -> HRESULT;
    /// Lyrics for the track.
    pub unsafe fn set_Lyrics(&self, Lyrics: BSTR) -> HRESULT;
    /// Category for the track.
    pub unsafe fn Category(&self, Category: *mut BSTR) -> HRESULT;
    /// Category for the track.
    pub unsafe fn set_Category(&self, Category: BSTR) -> HRESULT;
    /// Description for the track.
    pub unsafe fn Description(&self, Description: *mut BSTR) -> HRESULT;
    /// Description for the track.
    pub unsafe fn set_Description(&self, Description: BSTR) -> HRESULT;
    /// Long description for the track.
    pub unsafe fn LongDescription(&self, LongDescription: *mut BSTR) -> HRESULT;
    /// Long description for the track.
    pub unsafe fn set_LongDescription(&self, LongDescription: BSTR) -> HRESULT;
    /// The bookmark time of the track (in seconds).
    pub unsafe fn BookmarkTime(&self, BookmarkTime: *mut LONG) -> HRESULT;
    /// The bookmark time of the track (in seconds).
    pub unsafe fn set_BookmarkTime(&self, BookmarkTime: LONG) -> HRESULT;
    /// The video track kind.
    pub unsafe fn VideoKind(&self, VideoKind: *mut ITVideoKind) -> HRESULT;
    /// The video track kind.
    pub unsafe fn set_VideoKind(&self, VideoKind: ITVideoKind) -> HRESULT;
    /// The number of times the track has been skipped.
    pub unsafe fn SkippedCount(&self, SkippedCount: *mut LONG) -> HRESULT;
    /// The number of times the track has been skipped.
    pub unsafe fn set_SkippedCount(&self, SkippedCount: LONG) -> HRESULT;
    /// The date and time the track was last skipped.  A value of zero means no skipped date.
    pub unsafe fn SkippedDate(&self, SkippedDate: *mut DATE) -> HRESULT;
    /// The date and time the track was last skipped.  A value of zero means no skipped date.
    pub unsafe fn set_SkippedDate(&self, SkippedDate: DATE) -> HRESULT;
    /// True if track is part of a gapless album.
    pub unsafe fn PartOfGaplessAlbum(&self, PartOfGaplessAlbum: *mut VARIANT_BOOL) -> HRESULT;
    /// True if track is part of a gapless album.
    pub unsafe fn set_PartOfGaplessAlbum(&self, PartOfGaplessAlbum: VARIANT_BOOL) -> HRESULT;
    /// The album artist of the track.
    pub unsafe fn AlbumArtist(&self, AlbumArtist: *mut BSTR) -> HRESULT;
    /// The album artist of the track.
    pub unsafe fn set_AlbumArtist(&self, AlbumArtist: BSTR) -> HRESULT;
    /// The show name of the track.
    pub unsafe fn Show(&self, showName: *mut BSTR) -> HRESULT;
    /// The show name of the track.
    pub unsafe fn set_Show(&self, showName: BSTR) -> HRESULT;
    /// The season number of the track.
    pub unsafe fn SeasonNumber(&self, SeasonNumber: *mut LONG) -> HRESULT;
    /// The season number of the track.
    pub unsafe fn set_SeasonNumber(&self, SeasonNumber: LONG) -> HRESULT;
    /// The episode ID of the track.
    pub unsafe fn EpisodeID(&self, EpisodeID: *mut BSTR) -> HRESULT;
    /// The episode ID of the track.
    pub unsafe fn set_EpisodeID(&self, EpisodeID: BSTR) -> HRESULT;
    /// The episode number of the track.
    pub unsafe fn EpisodeNumber(&self, EpisodeNumber: *mut LONG) -> HRESULT;
    /// The episode number of the track.
    pub unsafe fn set_EpisodeNumber(&self, EpisodeNumber: LONG) -> HRESULT;
    /// The high 32-bits of the size of the track (in bytes).
    pub unsafe fn Size64High(&self, sizeHigh: *mut LONG) -> HRESULT;
    /// The low 32-bits of the size of the track (in bytes).
    pub unsafe fn Size64Low(&self, sizeLow: *mut LONG) -> HRESULT;
    /// True if track has not been played.
    pub unsafe fn Unplayed(&self, isUnplayed: *mut VARIANT_BOOL) -> HRESULT;
    /// True if track has not been played.
    pub unsafe fn set_Unplayed(&self, isUnplayed: VARIANT_BOOL) -> HRESULT;
    /// The album used for sorting.
    pub unsafe fn SortAlbum(&self, Album: *mut BSTR) -> HRESULT;
    /// The album used for sorting.
    pub unsafe fn set_SortAlbum(&self, Album: BSTR) -> HRESULT;
    /// The album artist used for sorting.
    pub unsafe fn SortAlbumArtist(&self, AlbumArtist: *mut BSTR) -> HRESULT;
    /// The album artist used for sorting.
    pub unsafe fn set_SortAlbumArtist(&self, AlbumArtist: BSTR) -> HRESULT;
    /// The artist used for sorting.
    pub unsafe fn SortArtist(&self, Artist: *mut BSTR) -> HRESULT;
    /// The artist used for sorting.
    pub unsafe fn set_SortArtist(&self, Artist: BSTR) -> HRESULT;
    /// The composer used for sorting.
    pub unsafe fn SortComposer(&self, Composer: *mut BSTR) -> HRESULT;
    /// The composer used for sorting.
    pub unsafe fn set_SortComposer(&self, Composer: BSTR) -> HRESULT;
    /// The track name used for sorting.
    pub unsafe fn SortName(&self, Name: *mut BSTR) -> HRESULT;
    /// The track name used for sorting.
    pub unsafe fn set_SortName(&self, Name: BSTR) -> HRESULT;
    /// The show name used for sorting.
    pub unsafe fn SortShow(&self, showName: *mut BSTR) -> HRESULT;
    /// The show name used for sorting.
    pub unsafe fn set_SortShow(&self, showName: BSTR) -> HRESULT;
    /// Reveal the track in the main browser window.
    pub unsafe fn Reveal(&self) -> HRESULT;
    /// The user or computed rating of the album that this track belongs to (0 to 100).
    pub unsafe fn AlbumRating(&self, Rating: *mut LONG) -> HRESULT;
    /// The user or computed rating of the album that this track belongs to (0 to 100).
    pub unsafe fn set_AlbumRating(&self, Rating: LONG) -> HRESULT;
    /// The album rating kind.
    pub unsafe fn AlbumRatingKind(&self, ratingKind: *mut ITRatingKind) -> HRESULT;
    /// The track rating kind.
    pub unsafe fn ratingKind(&self, ratingKind: *mut ITRatingKind) -> HRESULT;
    /// Returns a collection of playlists that contain the song that this track represents.
    pub unsafe fn Playlists(&self, iPlaylistCollection: *mut Option<IITPlaylistCollection>) -> HRESULT;
    /// The full path to the file represented by this track.
    pub unsafe fn set_Location(&self, Location: BSTR) -> HRESULT;
    /// The release date of the track.  A value of zero means no release date.
    pub unsafe fn ReleaseDate(&self, ReleaseDate: *mut DATE) -> HRESULT;
}

/// IITPlaylistWindow Interface
///
/// See the generated [`IITPlaylistWindow_Impl`] trait for more documentation about each function.
#[windows::core::interface("349CBB45-2E5A-4822-8E4A-A75555A186F7")]
pub unsafe trait IITPlaylistWindow : IITWindow {
    /// Returns a collection containing the currently selected track or tracks.
    pub unsafe fn SelectedTracks(&self, iTrackCollection: *mut Option<IITTrackCollection>) -> HRESULT;
    /// The playlist displayed in the window.
    pub unsafe fn Playlist(&self, iPlaylist: *mut Option<IITPlaylist>) -> HRESULT;
}
