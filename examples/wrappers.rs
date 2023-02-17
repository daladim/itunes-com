//! This example shows how to use the safe wrappers
//!
//! It must be built with the `--all-features` Cargo flag

#![allow(non_snake_case)]

use itunes_com::com::ITSourceKind;
use itunes_com::com::ITPlaylistSearchField;
use itunes_com::wrappers::IITObjectWrapper;
use itunes_com::wrappers::IITPlaylistWrapper;


fn main() {
    let mut iTunes = itunes_com::wrappers::iTunes::new().unwrap();

    iTunes.NextTrack().unwrap();
    iTunes.PlayFile(r"C:\My Music\Artist\Album\Title.mp3".into()).unwrap();

    show_playlists(&iTunes).unwrap();
    search_tracks(&iTunes, "beatles").unwrap();
    test_unique_ids(&iTunes).unwrap();
}


fn show_playlists(iTunes: &itunes_com::wrappers::iTunes) -> windows::core::Result<()> {
    let sources = iTunes.Sources()?;
    for source in sources.iter()? {
        let kind = source.Kind()?;
        println!("Source kind: {:?}", kind);
        if kind == ITSourceKind::ITSourceKindLibrary {
            for pl in source.Playlists()?.iter()? {
                let pl_kind = pl.Kind()?;
                let tracks = pl.Tracks()?;
                let track_count = tracks.iter()?.len();
                let first_track = tracks.ItemByPlayOrder(1);
                let first_track_name = first_track.and_then(|t| t.Name()).unwrap_or(String::from("<no track>"));
                println!("  * {}\t{:?}: {} tracks (first one is {})", pl.Name()?, pl_kind, track_count, first_track_name);
            }
        }
    }

    Ok(())
}

fn search_tracks(iTunes: &itunes_com::wrappers::iTunes, search_text: &str) -> windows::core::Result<()> {
    println!("Searching for \"{}\"...", search_text);

    let library_playlist = iTunes.LibraryPlaylist()?;
    let results = library_playlist.Search(search_text.into(), ITPlaylistSearchField::ITPlaylistSearchFieldAll)?;

    for result in results.iter()? {
        let file_location = result.as_file_or_cd_track().map(|foct| foct.Location());

        println!("  * {} at {:?}", result.Name()?, file_location);
    }

    Ok(())
}

fn test_unique_ids(iTunes: &itunes_com::wrappers::iTunes) -> windows::core::Result<()> {
    let library_playlist = iTunes.LibraryPlaylist()?;
    let first_track = library_playlist.Tracks()?.ItemByPlayOrder(1)?;
    println!("First track is {}", first_track.Name()?);

    let ids = first_track.GetITObjectIDs()?;
    println!("  IDs = {:?}", ids);

    let retrieved = iTunes.GetITObjectByID(ids)?;
    println!("  OK: {:?}", retrieved.as_track().map(|t| t.Name()));

    Ok(())
}
