//! Safe wrappers over the COM API.
//!
//! You usually want to start by creating an instance of the `iTunes` interface by [`iTunes::new`], then use its various methods.

#![allow(non_snake_case)]
#![allow(non_camel_case_types)]

pub mod iter;

// We'd rather use the re-exported versions, so that they are available to our users.
use crate::sys::*;

use windows::core::BSTR;
use windows::core::HRESULT;
use windows::core::Interface;
use windows::Win32::Media::Multimedia::NS_E_PROPERTY_NOT_FOUND;

use windows::Win32::System::Com::{CoInitializeEx, CoCreateInstance, CLSCTX_ALL, COINIT_MULTITHREADED};
use windows::Win32::System::Com::{VARIANT_0, VARIANT_0_0, VARIANT_0_0_0};

type DATE = f64; // This type must be a joke. https://learn.microsoft.com/en-us/cpp/atl-mfc-shared/date-type?view=msvc-170
type LONG = i32;

use widestring::ucstring::U16CString;
use num_traits::FromPrimitive;

pub struct Variant<'a, T: 'a> {
    inner: VARIANT,
    lifetime: PhantomData<&'a T>,
}

impl<'a, T> Variant<'a, T> {
    fn new(inner: VARIANT) -> Self {
        Self { inner, lifetime: PhantomData }
    }

    /// Get the wrapped `VARIANT`
    fn as_raw(&self) -> &VARIANT {
        &self.inner
    }
}

pub type PersistentId = u64;


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
    ($(#[$attr:meta])* $struct_name:ident) => {
        ::paste::paste! {
            com_wrapper_struct!($(#[$attr])* $struct_name as [<IIT $struct_name>]);
        }
    };
    ($(#[$attr:meta])* $struct_name:ident as $com_type:ident) => {
        $(#[$attr])*
        pub struct $struct_name {
            com_object: crate::sys::$com_type,
        }

        impl private::ComObjectWrapper for $struct_name {
            type WrappedType = $com_type;

            fn from_com_object(com_object: crate::sys::$com_type) -> Self {
                Self {
                    com_object
                }
            }

            fn com_object(&self) -> &crate::sys::$com_type {
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
    ($(#[$attr:meta])* $vis:vis $func_name:ident) => {
        no_args!($(#[$attr])* $vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $func_name:ident as $inherited_type:ty) => {
        $(#[$attr])*
        $vis fn $func_name(&self) -> windows::core::Result<()> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result: HRESULT = unsafe{ inherited_obj.$func_name() };
            result.ok()
        }
    };
}

macro_rules! get_bstr {
    ($(#[$attr:meta])* $vis:vis $func_name:ident) => {
        get_bstr!($(#[$attr])* $vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $func_name:ident as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $func_name:ident ( $key:ident ) as $inherited_type:ty) => {
        $(#[$attr])*
        $vis fn $func_name(&mut self, $key: &str) -> windows::core::Result<()> {
            str_to_bstr!($key, bstr);
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(bstr) };
            result.ok()
        }
    };
}

macro_rules! set_bstr {
    ($(#[$attr:meta])* $vis:vis $key:ident) => {
        set_bstr!($(#[$attr])* $vis $key as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $key:ident as $inherited_type:ty) => {
        ::paste::paste! {
            internal_set_bstr!($(#[$attr])* $vis [<set_ $key>] ( $key ) as $inherited_type);
        }
    };
    ($(#[$attr:meta])* $vis:vis $key:ident, no_set_prefix) => {
        internal_set_bstr!($(#[$attr])* $vis $key ($key) as <Self as ComObjectWrapper>::WrappedType);
    };
}

macro_rules! get_long {
    ($(#[$attr:meta])* $vis:vis $func_name:ident) => {
        get_long!($(#[$attr])* $vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $func_name:ident as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $func_name:ident ( $key:ident ) as $inherited_type:ty) => {
        $(#[$attr])*
        $vis fn $func_name(&mut self, $key: LONG) -> windows::core::Result<()> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name($key) };
            result.ok()
        }
    };
}

macro_rules! set_long {
    ($(#[$attr:meta])* $vis:vis $key:ident) => {
        ::paste::paste! {
            set_long!($(#[$attr])* $vis $key as <Self as ComObjectWrapper>::WrappedType);
        }
    };
    ($(#[$attr:meta])* $vis:vis $key:ident as $inherited_type:ty) => {
        ::paste::paste! {
            internal_set_long!($(#[$attr])* $vis [<set_ $key>] ($key) as $inherited_type);
        }
    };
    ($(#[$attr:meta])* $vis:vis $key:ident, no_set_prefix) => {
        ::paste::paste! {
            internal_set_long!($(#[$attr])* $vis $key ($key) as <Self as ComObjectWrapper>::WrappedType);
        }
    };
}

macro_rules! set_playlist {
    ($(#[$attr:meta])* $vis:vis $func_name:ident ( $arg:ident )) => {
        ::paste::paste! {
            $(#[$attr])*
            $vis fn [<set_ $func_name>](&mut self, $arg: &Playlist) -> windows::core::Result<()> {
                let vplaylist = $arg.as_variant();
                let result = unsafe{ self.com_object.[<set_ $func_name>](vplaylist.as_raw() as *const VARIANT) };
                result.ok()
            }
        }
    };
}

macro_rules! get_f64 {
    ($(#[$attr:meta])* $vis:vis $func_name:ident, $float_name:ty) => {
        get_f64!($(#[$attr])* $vis $func_name, $float_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $func_name:ident, $float_name:ty as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $key:ident, $float_name:ty) => {
        set_f64!($(#[$attr])* $vis $key, $float_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $key:ident, $float_name:ty as $inherited_type:ty) => {
        ::paste::paste! {
            $(#[$attr])*
            $vis fn [<set _$key>](&mut self, $key: $float_name) -> windows::core::Result<()> {
                let inherited_obj = self.com_object().cast::<$inherited_type>()?;
                let result = unsafe{ inherited_obj.[<set _$key>]($key) };
                result.ok()
            }
        }
    }
}

macro_rules! get_double {
    ($(#[$attr:meta])* $vis:vis $key:ident) => {
        get_f64!($(#[$attr])* $vis $key, f64);
    };
    ($(#[$attr:meta])* $vis:vis $key:ident as $inherited_type:ty) => {
        get_f64!($(#[$attr])* $vis $key, f64 as $inherited_type);
    }
}

macro_rules! set_double {
    ($(#[$attr:meta])*  $vis:vis $key:ident) => {
        set_f64!($(#[$attr])* $vis $key, f64);
    }
}

macro_rules! get_date {
    ($(#[$attr:meta])* $vis:vis $key:ident) => {
        get_f64!($(#[$attr])* $vis $key, DATE);
    };
    ($(#[$attr:meta])* $vis:vis $key:ident as $inherited_type:ty) => {
        get_f64!($(#[$attr])* $vis $key, DATE as $inherited_type);
    };
}

macro_rules! set_date {
    ($(#[$attr:meta])* $vis:vis $key:ident) => {
        set_f64!($(#[$attr])* $vis $key, DATE);
    };
    ($(#[$attr:meta])* $vis:vis $key:ident as $inherited_type:ty) => {
        set_f64!($(#[$attr])* $vis $key, DATE as $inherited_type);
    };
}

macro_rules! get_bool {
    ($(#[$attr:meta])* $vis:vis $func_name:ident) => {
        get_bool!($(#[$attr])* $vis $func_name as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $func_name:ident as $inherited_type:ty) => {
        ::paste::paste! {
            $(#[$attr])*
            $vis fn [<is _$func_name>](&self) -> windows::core::Result<bool> {
                let mut value = crate::sys::FALSE;
                let inherited_obj = self.com_object().cast::<$inherited_type>()?;
                let result = unsafe{ inherited_obj.$func_name(&mut value) };
                result.ok()?;

                Ok(value.as_bool())
            }
        }
    };
}



macro_rules! internal_set_bool {
    ($(#[$attr:meta])* $vis:vis $func_name:ident ( $key:ident ) as $inherited_type:ty) => {
        $(#[$attr])*
        $vis fn $func_name(&mut self, $key: bool) -> windows::core::Result<()> {
            let variant_bool = match $key {
                true => crate::sys::TRUE,
                false => crate::sys::FALSE,
            };
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;
            let result = unsafe{ inherited_obj.$func_name(variant_bool) };
            result.ok()
        }
    };
}

macro_rules! set_bool {
    ($(#[$attr:meta])* $vis:vis $key:ident) => {
        set_bool!($(#[$attr])* $vis $key as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $key:ident as $inherited_type:ty) => {
        ::paste::paste! {
            internal_set_bool!($(#[$attr])* $vis [<set_ $key>] ( $key ) as $inherited_type);
        }
    };
    ($(#[$attr:meta])* $vis:vis $key:ident, no_set_prefix) => {
        internal_set_bool!($(#[$attr])* $vis $key ($key) as <Self as ComObjectWrapper>::WrappedType);
    }
}


macro_rules! get_enum {
    ($(#[$attr:meta])* $vis:vis $fn_name:ident -> $enum_type:ty) => {
        get_enum!($(#[$attr])* $vis $fn_name -> $enum_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $fn_name:ident -> $enum_type:ty as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $fn_name:ident, $enum_type:ty) => {
        set_enum!($(#[$attr])* $vis $fn_name, $enum_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $fn_name:ident, $enum_type:ty as $inherited_type:ty) => {
        ::paste::paste! {
            $(#[$attr])*
            $vis fn [<set _$fn_name>](&mut self, value: $enum_type) -> windows::core::Result<()> {
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
    ($(#[$attr:meta])* $vis:vis $fn_name:ident -> $obj_type:ty) => {
        get_object!($(#[$attr])* $vis $fn_name -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $fn_name:ident -> $obj_type:ty as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty) => {
        get_object_from_str!($(#[$attr])* $vis $fn_name($arg_name) -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty) => {
        get_object_from_variant!($(#[$attr])* $vis $fn_name($arg_name) -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty as $inherited_type:ty) => {
        $(#[$attr])*
        $vis fn $fn_name<T>(&self, $arg_name:&Variant<T>) -> windows::core::Result<$obj_type> {
            let inherited_obj = self.com_object().cast::<$inherited_type>()?;

            let mut out_obj = None;
            let result = unsafe{ inherited_obj.$fn_name($arg_name.as_raw() as *const VARIANT, &mut out_obj as *mut _) };
            result.ok()?;

            create_wrapped_object!($obj_type, out_obj)
        }
    };
}

macro_rules! get_object_from_long {
    ($(#[$attr:meta])* $vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty) => {
        get_object_from_long!($(#[$attr])* $vis $fn_name($arg_name) -> $obj_type as <Self as ComObjectWrapper>::WrappedType);
    };
    ($(#[$attr:meta])* $vis:vis $fn_name:ident ( $arg_name:ident ) -> $obj_type:ty as $inherited_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $fn_name:ident, $obj_type:ty) => {
        ::paste::paste! {
            $(#[$attr])*
            $vis fn [<set _$fn_name>](&mut self, data: $obj_type) -> windows::core::Result<()> {
                let object_to_set = data.com_object();
                let result = unsafe{ self.com_object.[<set _$fn_name>](object_to_set as *const _) };
                result.ok()
            }
        }
    }
}

macro_rules! item_by_name {
    ($(#[$attr:meta])* $vis:vis $obj_type:ty) => {
        $(#[$attr])*
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
    ($(#[$attr:meta])* $vis:vis $obj_type:ty) => {
        $(#[$attr])*
        $vis fn ItemByPersistentID(&self, id: PersistentId) -> windows::core::Result<$obj_type> {
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
    ///
    /// These ID are "runtime" IDs, only valid for this current session. See [here for more info](https://web.archive.org/web/20201030012249/http://www.joshkunz.com/iTunesControl/interfaceIITObject.html)<br/>
    /// Use [`iTunes::GetITObjectByID`] for the reverse operation.
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

    // Not sure whether the content of the Variant will be released when dropped.
    // To be sure, let's assign it a lifetime anyway
    /// Get a COM `VARIANT` pointing to this object
    fn as_variant(&self) -> Variant<Self::WrappedType> {
        let idispatch = self.com_object().cast::<windows::Win32::System::Com::IDispatch>().unwrap(); // unwrapping here is fine, every IITObject inherits from IDispatch

        // See https://microsoft.public.vc.atl.narkive.com/nSoZZbkL/passing-pointers-using-a-variant
        Variant::new(VARIANT{
            Anonymous: VARIANT_0 {
                Anonymous: std::mem::ManuallyDrop::new(VARIANT_0_0 {
                    vt: windows::Win32::System::Com::VT_DISPATCH,
                    Anonymous: VARIANT_0_0_0 {
                        pdispVal: std::mem::ManuallyDrop::new(Some(idispatch)),
                    },
                    ..Default::default()
                })
            }
        })
    }

    get_bstr!(
        /// The name of the object.
        Name as IITObject);

    set_bstr!(
        /// The name of the object.
        Name as IITObject);

    get_long!(
        /// The index of the object in internal application order (1-based).
        Index as IITObject);

    get_long!(
        /// The source ID of the object.
        sourceID as IITObject);

    get_long!(
        /// The playlist ID of the object.
        playlistID as IITObject);

    get_long!(
        /// The track ID of the object.
        trackID as IITObject);

    get_long!(
        /// The track database ID of the object.
        TrackDatabaseID as IITObject);
}

/// Enum of all structs that directly inherit from [`IITObject`]
pub enum PossibleIITObject {
    Source(Source),
    Playlist(Playlist),
    Track(Track),
}

impl PossibleIITObject {
    fn from_com_object(com_object: IITObject) -> windows::core::Result<PossibleIITObject> {
        if let Ok(source) = com_object.cast::<IITSource>() {
            Ok(PossibleIITObject::Source(Source::from_com_object(source)))
        } else if let Ok(playlist) = com_object.cast::<IITPlaylist>() {
            Ok(PossibleIITObject::Playlist(Playlist::from_com_object(playlist)))
        } else if let Ok(track) = com_object.cast::<IITTrack>() {
            Ok(PossibleIITObject::Track(Track::from_com_object(track)))
        } else {
            Err(windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            ))
        }
    }

    pub fn as_source(&self) -> Option<&Source> {
        match self {
            PossibleIITObject::Source(s) => Some(s),
            _ => None
        }
    }

    pub fn as_playlist(&self) -> Option<&Playlist> {
        match self {
            PossibleIITObject::Playlist(p) => Some(p),
            _ => None
        }
    }

    pub fn as_track(&self) -> Option<&Track> {
        match self {
            PossibleIITObject::Track(t) => Some(t),
            _ => None
        }
    }
}


com_wrapper_struct!(
    /// Safe wrapper over a [`IITSource`](crate::sys::IITSource)
    Source);

impl IITObjectWrapper for Source {}

impl Source {
    get_enum!(
        /// The source kind.
        pub Kind -> ITSourceKind);

    get_double!(
        /// The total size of the source, if it has a fixed size.
        pub Capacity);

    get_double!(
        /// The free space on the source, if it has a fixed size.
        pub FreeSpace);

    get_object!(
        /// Returns a collection of playlists.
        pub Playlists -> PlaylistCollection);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITPlaylistCollection`](crate::sys::IITPlaylistCollection)
    PlaylistCollection);

impl PlaylistCollection {
    item_by_name!(
        /// Returns an IITPlaylist object with the specified name.
        pub Playlist);

    item_by_persistent_id!(
        /// Returns an IITPlaylist object with the specified persistent ID.
        pub Playlist);
}

iterator!(PlaylistCollection, Playlist);


/// Several COM objects inherit from this class, which provides some extra methods
pub trait IITPlaylistWrapper: private::ComObjectWrapper {
    no_args!(
        /// Delete this playlist.
        Delete as IITPlaylist);

    no_args!(
        /// Start playing the first track in this playlist.
        PlayFirstTrack as IITPlaylist);

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

    get_enum!(
        /// The playlist kind.
        Kind -> ITPlaylistKind as IITPlaylist);

    get_object!(
        /// The source that contains this playlist.
        Source -> Source as IITPlaylist);

    get_long!(
        /// The total length of all songs in the playlist (in seconds).
        Duration as IITPlaylist);

    get_bool!(
        /// True if songs in the playlist are played in random order.
        Shuffle as IITPlaylist);

    set_bool!(
        /// True if songs in the playlist are played in random order.
        Shuffle as IITPlaylist);

    get_double!(
        /// The total size of all songs in the playlist (in bytes).
        Size as IITPlaylist);

    get_enum!(
        /// The playback repeat mode.
        SongRepeat -> ITPlaylistRepeatMode as IITPlaylist);

    set_enum!(
        /// The playback repeat mode.
        SongRepeat, ITPlaylistRepeatMode as IITPlaylist);

    get_bstr!(
        /// The total length of all songs in the playlist (in MM:SS format).
        Time as IITPlaylist);

    get_bool!(
        /// True if the playlist is visible in the Source list.
        Visible as IITPlaylist);

    get_object!(
        /// Returns a collection of tracks in this playlist.
        Tracks -> TrackCollection as IITPlaylist);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITPlaylist`](crate::sys::IITPlaylist)
    Playlist);

impl IITObjectWrapper for Playlist {}

impl IITPlaylistWrapper for Playlist {}


com_wrapper_struct!(
    /// Safe wrapper over a [`IITTrackCollection`](crate::sys::IITTrackCollection)
    TrackCollection);

impl TrackCollection {
    get_object_from_long!(
        /// Returns an IITTrack object corresponding to the given index, where the index is defined by the play order of the playlist containing the track collection (1-based).
        pub ItemByPlayOrder(Index) -> Track);

    item_by_name!(
        /// Returns an IITTrack object with the specified name.
        pub Track);

    item_by_persistent_id!(
        /// Returns an IITTrack object with the specified persistent ID.
        pub Track);
}

iterator!(TrackCollection, Track);

/// Several COM objects inherit from this class, which provides some extra methods
pub trait IITTrackWrapper: private::ComObjectWrapper {
    no_args!(
        /// Delete this track.
        Delete as IITTrack);

    no_args!(
        /// Start playing this track.
        Play as IITTrack);

    get_object_from_str!(
        /// Add artwork from an image file to this track.
        AddArtworkFromFile(filePath) -> Artwork as IITTrack);

    get_enum!(
        /// The track kind.
        Kind -> ITTrackKind as IITTrack);

    get_object!(
        /// The playlist that contains this track.
        Playlist -> Playlist as IITTrack);

    get_bstr!(
        /// The album containing the track.
        Album as IITTrack);

    set_bstr!(
        /// The album containing the track.
        Album as IITTrack);

    get_bstr!(
        /// The artist/source of the track.
        Artist as IITTrack);

    set_bstr!(
        /// The artist/source of the track.
        Artist as IITTrack);

    get_long!(
        /// The bit rate of the track (in kbps).
        BitRate as IITTrack);

    get_long!(
        /// The tempo of the track (in beats per minute).
        BPM as IITTrack);

    set_long!(
        /// The tempo of the track (in beats per minute).
        BPM as IITTrack);

    get_bstr!(
        /// Freeform notes about the track.
        Comment as IITTrack);

    set_bstr!(
        /// Freeform notes about the track.
        Comment as IITTrack);

    get_bool!(
        /// True if this track is from a compilation album.
        Compilation as IITTrack);

    set_bool!(
        /// True if this track is from a compilation album.
        Compilation as IITTrack);

    get_bstr!(
        /// The composer of the track.
        Composer as IITTrack);

    set_bstr!(
        /// The composer of the track.
        Composer as IITTrack);

    get_date!(
        /// The date the track was added to the playlist.
        DateAdded as IITTrack);

    get_long!(
        /// The total number of discs in the source album.
        DiscCount as IITTrack);

    set_long!(
        /// The total number of discs in the source album.
        DiscCount as IITTrack);

    get_long!(
        /// The index of the disc containing the track on the source album.
        DiscNumber as IITTrack);

    set_long!(
        /// The index of the disc containing the track on the source album.
        DiscNumber as IITTrack);

    get_long!(
        /// The length of the track (in seconds).
        Duration as IITTrack);

    get_bool!(
        /// True if the track is checked for playback.
        Enabled as IITTrack);

    set_bool!(
        /// True if the track is checked for playback.
        Enabled as IITTrack);

    get_bstr!(
        /// The name of the EQ preset of the track.
        EQ as IITTrack);

    set_bstr!(
        /// The name of the EQ preset of the track.
        EQ as IITTrack);

    set_long!(
        /// The stop time of the track (in seconds).
        Finish as IITTrack);

    get_long!(
        /// The stop time of the track (in seconds).
        Finish as IITTrack);

    get_bstr!(
        /// The music/audio genre (category) of the track.
        Genre as IITTrack);

    set_bstr!(
        /// The music/audio genre (category) of the track.
        Genre as IITTrack);

    get_bstr!(
        /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
        Grouping as IITTrack);

    set_bstr!(
        /// The grouping (piece) of the track.  Generally used to denote movements within classical work.
        Grouping as IITTrack);

    get_bstr!(
        /// A text description of the track.
        KindAsString as IITTrack);

    get_date!(
        /// The modification date of the content of the track.
        ModificationDate as IITTrack);

    get_long!(
        /// The number of times the track has been played.
        PlayedCount as IITTrack);

    set_long!(
        /// The number of times the track has been played.
        PlayedCount as IITTrack);

    get_date!(
        /// The date and time the track was last played.  A value of zero means no played date.
        PlayedDate as IITTrack);

    set_date!(
        /// The date and time the track was last played.  A value of zero means no played date.
        PlayedDate as IITTrack);

    get_long!(
        /// The play order index of the track in the owner playlist (1-based).
        PlayOrderIndex as IITTrack);

    get_long!(
        /// The rating of the track (0 to 100).
        Rating as IITTrack);

    set_long!(
        /// The rating of the track (0 to 100).
        Rating as IITTrack);

    get_long!(
        /// The sample rate of the track (in Hz).
        SampleRate as IITTrack);

    get_long!(
        /// The size of the track (in bytes).
        Size as IITTrack);

    get_long!(
        /// The start time of the track (in seconds).
        Start as IITTrack);

    set_long!(
        /// The start time of the track (in seconds).
        Start as IITTrack);

    get_bstr!(
        /// The length of the track (in MM:SS format).
        Time as IITTrack);

    get_long!(
        /// The total number of tracks on the source album.
        TrackCount as IITTrack);

    set_long!(
        /// The total number of tracks on the source album.
        TrackCount as IITTrack);

    get_long!(
        /// The index of the track on the source album.
        TrackNumber as IITTrack);

    set_long!(
        /// The index of the track on the source album.
        TrackNumber as IITTrack);

    get_long!(
        /// The relative volume adjustment of the track (-100% to 100%).
        VolumeAdjustment as IITTrack);

    set_long!(
        /// The relative volume adjustment of the track (-100% to 100%).
        VolumeAdjustment as IITTrack);

    get_long!(
        /// The year the track was recorded/released.
        Year as IITTrack);

    set_long!(
        /// The year the track was recorded/released.
        Year as IITTrack);

    get_object!(
        /// Returns a collection of artwork.
        Artwork -> ArtworkCollection as IITTrack);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITTrack`](crate::sys::IITTrack)
    Track);

impl IITObjectWrapper for Track {}

impl IITTrackWrapper for Track {}

impl Track {
    /// In case the concrete COM object for this track actually is a derived `FileOrCDTrack`, this is a way to retrieve it
    pub fn as_file_or_cd_track(&self) -> windows::core::Result<FileOrCDTrack> {
        let foct = self.com_object.cast::<IITFileOrCDTrack>()?;
        Ok(FileOrCDTrack::from_com_object(foct))
    }
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITArtwork`](crate::sys::IITArtwork)
    Artwork);

impl Artwork {
    no_args!(
        /// Delete this piece of artwork from the track.
        pub Delete);

    set_bstr!(
        /// Replace existing artwork data with new artwork from an image file.
        pub SetArtworkFromFile, no_set_prefix);

    set_bstr!(
        /// Save artwork data to an image file.
        pub SaveArtworkToFile, no_set_prefix);

    get_enum!(
        /// The format of the artwork.
        pub Format -> ITArtworkFormat);

    get_bool!(
        /// True if the artwork was downloaded by iTunes.
        pub IsDownloadedArtwork);

    get_bstr!(
        /// The description for the artwork.
        pub Description);

    set_bstr!(
        /// The description for the artwork.
        pub Description);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITArtworkCollection`](crate::sys::IITArtworkCollection)
    ArtworkCollection);

impl ArtworkCollection {}

iterator!(ArtworkCollection, Artwork);

com_wrapper_struct!(
    /// Safe wrapper over a [`IITSourceCollection`](crate::sys::IITSourceCollection)
    SourceCollection);

impl SourceCollection {
    item_by_name!(
        /// Returns an IITSource object with the specified name.
        pub Source);

    item_by_persistent_id!(
        /// Returns an IITSource object with the specified persistent ID.
        pub Source);
}

iterator!(SourceCollection, Source);

com_wrapper_struct!(
    /// Safe wrapper over a [`IITEncoder`](crate::sys::IITEncoder)
    Encoder);

impl Encoder {
    get_bstr!(
        /// The name of the the encoder.
        pub Name);

    get_bstr!(
        /// The data format created by the encoder.
        pub Format);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITEncoderCollection`](crate::sys::IITEncoderCollection)
    EncoderCollection);

impl EncoderCollection {
    item_by_name!(
        /// Returns an IITEncoder object with the specified name.
        pub Encoder);
}

iterator!(EncoderCollection, Encoder);

com_wrapper_struct!(
    /// Safe wrapper over a [`IITEQPreset`](crate::sys::IITEQPreset)
    EQPreset);

impl EQPreset {
    get_bstr!(
        /// The name of the the EQ preset.
        pub Name);

    get_bool!(
        /// True if this EQ preset can be modified.
        pub Modifiable);

    get_double!(
        /// The equalizer preamp level (-12.0 db to +12.0 db).
        pub Preamp);

    set_double!(
        /// The equalizer preamp level (-12.0 db to +12.0 db).
        pub Preamp);

    get_double!(
        /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
        pub Band1);

    set_double!(
        /// The equalizer 32Hz band level (-12.0 db to +12.0 db).
        pub Band1);

    get_double!(
        /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
        pub Band2);

    set_double!(
        /// The equalizer 64Hz band level (-12.0 db to +12.0 db).
        pub Band2);

    get_double!(
        /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
        pub Band3);

    set_double!(
        /// The equalizer 125Hz band level (-12.0 db to +12.0 db).
        pub Band3);

    get_double!(
        /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
        pub Band4);

    set_double!(
        /// The equalizer 250Hz band level (-12.0 db to +12.0 db).
        pub Band4);

    get_double!(
        /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
        pub Band5);

    set_double!(
        /// The equalizer 500Hz band level (-12.0 db to +12.0 db).
        pub Band5);

    get_double!(
        /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
        pub Band6);

    set_double!(
        /// The equalizer 1KHz band level (-12.0 db to +12.0 db).
        pub Band6);

    get_double!(
        /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
        pub Band7);

    set_double!(
        /// The equalizer 2KHz band level (-12.0 db to +12.0 db).
        pub Band7);

    get_double!(
        /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
        pub Band8);

    set_double!(
        /// The equalizer 4KHz band level (-12.0 db to +12.0 db).
        pub Band8);

    get_double!(
        /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
        pub Band9);

    set_double!(
        /// The equalizer 8KHz band level (-12.0 db to +12.0 db).
        pub Band9);

    get_double!(
        /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
        pub Band10);

    set_double!(
        /// The equalizer 16KHz band level (-12.0 db to +12.0 db).
        pub Band10);

    internal_set_bool!(
        /// Delete this EQ preset.
        pub Delete(updateAllTracks) as <Self as ComObjectWrapper>::WrappedType);

    /// Rename this EQ preset.
    pub fn Rename(&self, newName: String, updateAllTracks: bool) -> windows::core::Result<()> {
        str_to_bstr!(newName, bstr);
        let var_bool = if updateAllTracks { TRUE } else { FALSE };
        let result = unsafe { self.com_object.Rename(bstr, var_bool) };
        result.ok()
    }
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITEQPresetCollection`](crate::sys::IITEQPresetCollection)
    EQPresetCollection);

impl EQPresetCollection {
    item_by_name!(
        /// Returns an IITEQPreset object with the specified name.
        pub EQPreset);
}

iterator!(EQPresetCollection, EQPreset);

com_wrapper_struct!(
    /// Safe wrapper over a [`IITOperationStatus`](crate::sys::IITOperationStatus)
    OperationStatus);

impl OperationStatus {
    get_bool!(
        /// True if the operation is still in progress.
        pub InProgress);

    get_object!(
        /// Returns a collection containing the tracks that were generated by the operation.
        pub Tracks -> TrackCollection);
}

/// The three items of a `ConversionStatus`
#[derive(Debug)]
pub struct ConversionStatus {
    pub trackName: String,
    pub progressValue: LONG,
    pub maxProgressValue: LONG,
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITConvertOperationStatus`](crate::sys::IITConvertOperationStatus)
    ConvertOperationStatus);

impl ConvertOperationStatus {
    /// Returns the current conversion status.
    pub fn GetConversionStatus(&self) -> windows::core::Result<ConversionStatus> {
        let mut bstr = BSTR::default();
        let mut progressValue = 0;
        let mut maxProgressValue = 0;
        let result = unsafe{ self.com_object.GetConversionStatus(&mut bstr, &mut progressValue as *mut LONG, &mut maxProgressValue as *mut LONG) };
        result.ok()?;

        let v: Vec<u16> = bstr.as_wide().to_vec();
        let trackName = U16CString::from_vec_truncate(v).to_string_lossy();

        Ok(ConversionStatus{ trackName, progressValue, maxProgressValue })
    }

    no_args!(
        /// Stops the current conversion operation.
        pub StopConversion);

    get_bstr!(
        /// Returns the name of the track currently being converted.
        pub trackName);

    get_long!(
        /// Returns the current progress value for the track being converted.
        pub progressValue);

    get_long!(
        /// Returns the maximum progress value for the track being converted.
        pub maxProgressValue);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITLibraryPlaylist`](crate::sys::IITLibraryPlaylist)
    LibraryPlaylist);

impl IITObjectWrapper for LibraryPlaylist {}

impl IITPlaylistWrapper for LibraryPlaylist {}

impl LibraryPlaylist {
    get_object_from_str!(
        /// Add the specified file path to the library.
        pub AddFile(filePath) -> OperationStatus);

    get_object_from_variant!(
        /// Add the specified array of file paths to the library. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
        pub AddFiles(filePaths) -> OperationStatus);

    get_object_from_str!(
        /// Add the specified streaming audio URL to the library.
        pub AddURL(URL) -> URLTrack);

    get_object_from_variant!(
        /// Add the specified track to the library.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
        pub AddTrack(iTrackToAdd) -> Track);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITURLTrack`](crate::sys::IITURLTrack)
    URLTrack);

impl IITObjectWrapper for URLTrack {}

impl IITTrackWrapper for URLTrack {}

impl URLTrack {
    get_bstr!(
        /// The URL of the stream represented by this track.
        pub URL);

    set_bstr!(
        /// The URL of the stream represented by this track.
        pub URL);

    get_bool!(
        /// True if this is a podcast track.
        pub Podcast);

    no_args!(
        /// Update the podcast feed for this track.
        pub UpdatePodcastFeed);

    no_args!(
        /// Start downloading the podcast episode that corresponds to this track.
        pub DownloadPodcastEpisode);

    get_bstr!(
        /// Category for the track.
        pub Category);

    set_bstr!(
        /// Category for the track.
        pub Category);

    get_bstr!(
        /// Description for the track.
        pub Description);

    set_bstr!(
        /// Description for the track.
        pub Description);

    get_bstr!(
        /// Long description for the track.
        pub LongDescription);

    set_bstr!(
        /// Long description for the track.
        pub LongDescription);

    no_args!(
        /// Reveal the track in the main browser window.
        pub Reveal);

    get_long!(
        /// The user or computed rating of the album that this track belongs to (0 to 100).
        pub AlbumRating);

    set_long!(
        /// The user or computed rating of the album that this track belongs to (0 to 100).
        pub AlbumRating);

    get_enum!(
        /// The album rating kind.
        pub AlbumRatingKind -> ITRatingKind);

    get_enum!(
        /// The track rating kind.
        pub ratingKind -> ITRatingKind);

    get_object!(
        /// Returns a collection of playlists that contain the song that this track represents.
        pub Playlists -> PlaylistCollection);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITUserPlaylist`](crate::sys::IITUserPlaylist)
    UserPlaylist);

impl IITObjectWrapper for UserPlaylist {}

impl IITPlaylistWrapper for UserPlaylist {}

impl UserPlaylist {
    get_object_from_str!(
        /// Add the specified file path to the user playlist.
        pub AddFile(filePath) -> OperationStatus);

    get_object_from_variant!(
        /// Add the specified array of file paths to the user playlist. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
        pub AddFiles(filePaths) -> OperationStatus);

    get_object_from_str!(
        /// Add the specified streaming audio URL to the user playlist.
        pub AddURL(URL) -> URLTrack);

    get_object_from_variant!(
        /// Add the specified track to the user playlist.  iTrackToAdd is a VARIANT of type VT_DISPATCH that points to an IITTrack.
        pub AddTrack(iTrackToAdd) -> Track);

    get_bool!(
        /// True if the user playlist is being shared.
        pub Shared);

    set_bool!(
        /// True if the user playlist is being shared.
        pub Shared);

    get_bool!(
        /// True if this is a smart playlist.
        pub Smart);

    get_enum!(
        /// The playlist special kind.
        pub SpecialKind -> ITUserPlaylistSpecialKind);

    get_object!(
        /// The parent of this playlist.
        pub Parent -> UserPlaylist);

    get_object_from_str!(
        /// Creates a new playlist in a folder playlist.
        pub CreatePlaylist(playlistName) -> Playlist);

    get_object_from_str!(
        /// Creates a new folder in a folder playlist.
        pub CreateFolder(folderName) -> Playlist);

    set_playlist!(
        /// The parent of this playlist.
        pub Parent(iParentPlayList));

    no_args!(
        /// Reveal the user playlist in the main browser window.
        pub Reveal);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITVisual`](crate::sys::IITVisual)
    Visual);

impl Visual {
    get_bstr!(
        /// The name of the the visual plug-in.
        pub Name);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITVisualCollection`](crate::sys::IITVisualCollection)
    VisualCollection);

impl VisualCollection {
    item_by_name!(
        /// Returns an IITVisual object with the specified name.
        pub Visual);
}

iterator!(VisualCollection, Visual);

com_wrapper_struct!(
    /// Safe wrapper over a [`IITWindow`](crate::sys::IITWindow)
    Window);

impl Window {
    get_bstr!(
        /// The title of the window.
        pub Name);

    get_enum!(
        /// The window kind.
        pub Kind -> ITWindowKind);

    get_bool!(
        /// True if the window is visible. Note that the main browser window cannot be hidden.
        pub Visible);

    set_bool!(
        /// True if the window is visible. Note that the main browser window cannot be hidden.
        pub Visible);

    get_bool!(
        /// True if the window is resizable.
        pub Resizable);

    get_bool!(
        /// True if the window is minimized.
        pub Minimized);

    set_bool!(
        /// True if the window is minimized.
        pub Minimized);

    get_bool!(
        /// True if the window is maximizable.
        pub Maximizable);

    get_bool!(
        /// True if the window is maximized.
        pub Maximized);

    set_bool!(
        /// True if the window is maximized.
        pub Maximized);

    get_bool!(
        /// True if the window is zoomable.
        pub Zoomable);

    get_bool!(
        /// True if the window is zoomed.
        pub Zoomed);

    set_bool!(
        /// True if the window is zoomed.
        pub Zoomed);

    get_long!(
        /// The screen coordinate of the top edge of the window.
        pub Top);

    set_long!(
        /// The screen coordinate of the top edge of the window.
        pub Top);

    get_long!(
        /// The screen coordinate of the left edge of the window.
        pub Left);

    set_long!(
        /// The screen coordinate of the left edge of the window.
        pub Left);

    get_long!(
        /// The screen coordinate of the bottom edge of the window.
        pub Bottom);

    set_long!(
        /// The screen coordinate of the bottom edge of the window.
        pub Bottom);

    get_long!(
        /// The screen coordinate of the right edge of the window.
        pub Right);

    set_long!(
        /// The screen coordinate of the right edge of the window.
        pub Right);

    get_long!(
        /// The width of the window.
        pub Width);

    set_long!(
        /// The width of the window.
        pub Width);

    get_long!(
        /// The height of the window.
        pub Height);

    set_long!(
        /// The height of the window.
        pub Height);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITBrowserWindow`](crate::sys::IITBrowserWindow)
    BrowserWindow);

impl BrowserWindow {
    get_bool!(
        /// True if window is in MiniPlayer mode.
        pub MiniPlayer);

    set_bool!(
        /// True if window is in MiniPlayer mode.
        pub MiniPlayer);

    get_object!(
        /// Returns a collection containing the currently selected track or tracks.
        pub SelectedTracks -> TrackCollection);

    get_object!(
        /// The currently selected playlist in the Source list.
        pub SelectedPlaylist -> Playlist);

    set_playlist!(
        /// The currently selected playlist in the Source list.
        pub SelectedPlaylist(iPlaylist));
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITWindowCollection`](crate::sys::IITWindowCollection)
    WindowCollection);

impl WindowCollection {
    item_by_name!(
        /// Returns an IITWindow object with the specified name.
        pub Window);
}

iterator!(WindowCollection, Window);

/// The three items of a `PlayerButtonState`
#[derive(Debug, Eq, PartialEq)]
pub struct PlayerButtonState {
    pub previousEnabled: bool,
    pub playPauseStopState: ITPlayButtonState,
    pub nextEnabled: bool,
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IiTunes`](crate::sys::IiTunes)
    iTunes as IiTunes);

impl iTunes {
    /// Create a new COM object to communicate with iTunes
    pub fn new() -> windows::core::Result<Self> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;
        }

        Ok(Self {
            com_object: unsafe { CoCreateInstance(&crate::sys::ITUNES_APP_COM_GUID, None, CLSCTX_ALL)? },
        })
    }

    no_args!(
        /// Reposition to the beginning of the current track or go to the previous track if already at start of current track.
        pub BackTrack);

    no_args!(
        /// Skip forward in a playing track.
        pub FastForward);

    no_args!(
        /// Advance to the next track in the current playlist.
        pub NextTrack);

    no_args!(
        /// Pause playback.
        pub Pause);

    no_args!(
        /// Play the currently targeted track.
        pub Play);

    set_bstr!(
        /// Play the specified file path, adding it to the library if not already present.
        pub PlayFile, no_set_prefix);

    no_args!(
        /// Toggle the playing/paused state of the current track.
        pub PlayPause);

    no_args!(
        /// Return to the previous track in the current playlist.
        pub PreviousTrack);

    no_args!(
        /// Disable fast forward/rewind and resume playback, if playing.
        pub Resume);

    no_args!(
        /// Skip backwards in a playing track.
        pub Rewind);

    no_args!(
        /// Stop playback.
        pub Stop);

    get_object_from_str!(
        /// Start converting the specified file path.
        pub ConvertFile(filePath) -> OperationStatus);

    get_object_from_variant!(
        /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
        pub ConvertFiles(filePaths) -> OperationStatus);

    get_object_from_variant!(
        /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
        pub ConvertTrack(iTrackToConvert) -> OperationStatus);

    get_object_from_variant!(
        /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
        pub ConvertTracks(iTracksToConvert) -> OperationStatus);

    /// Returns true if this version of the iTunes type library is compatible with the specified version.
    pub fn CheckVersion(&self, majorVersion: LONG, minorVersion: LONG) -> windows::core::Result<bool> {
        let mut bool_result = FALSE;
        let result = unsafe{ self.com_object.CheckVersion(majorVersion, minorVersion, &mut bool_result) };
        result.ok()?;
        Ok(bool_result.as_bool())
    }

    /// Returns an IITObject corresponding to the specified IDs.
    pub fn GetITObjectByID(&self, ids: ObjectIDs) -> windows::core::Result<PossibleIITObject> {
        let mut out_obj: Option<IITObject> = None;
        let result = unsafe{ self.com_object.GetITObjectByID(
            ids.sourceID,
            ids.playlistID,
            ids.trackID,
            ids.databaseID,
            &mut out_obj as *mut _
        ) };
        result.ok()?;

        match out_obj {
            None => Err(windows::core::Error::new(
                NS_E_PROPERTY_NOT_FOUND, // this is the closest matching HRESULT I could find...
                windows::h!("Item not found").clone(),
            )),
            Some(obj) => {
                PossibleIITObject::from_com_object(obj)
            },
        }
    }

    get_object_from_str!(
        /// Creates a new playlist in the main library.
        pub CreatePlaylist(playlistName) -> Playlist);

    set_bstr!(
        /// Open the specified iTunes Store or streaming audio URL.
        pub OpenURL, no_set_prefix);

    no_args!(
        /// Go to the iTunes Store home page.
        pub GotoMusicStoreHomePage);

    no_args!(
        /// Update the contents of the iPod.
        pub UpdateIPod);

    // /// [id(0x60020015)]
    // /// (no other documentation provided)
    // pub fn Authorize(&self, numElems: LONG, data: *const VARIANT, names: *const BSTR) -> windows::core::Result<()> {
    //     todo!()
    // }

    no_args!(
        /// Exits the iTunes application.
        pub Quit);

    get_object!(
        /// Returns a collection of music sources (music library, CD, device, etc.).
        pub Sources -> SourceCollection);

    get_object!(
        /// Returns a collection of encoders.
        pub Encoders -> EncoderCollection);

    get_object!(
        /// Returns a collection of EQ presets.
        pub EQPresets -> EQPresetCollection);

    get_object!(
        /// Returns a collection of visual plug-ins.
        pub Visuals -> VisualCollection);

    get_object!(
        /// Returns a collection of windows.
        pub Windows -> WindowCollection);

    get_long!(
        /// Returns the sound output volume (0 = minimum, 100 = maximum).
        pub SoundVolume);

    set_long!(
        /// Returns the sound output volume (0 = minimum, 100 = maximum).
        pub SoundVolume);

    get_bool!(
        /// True if sound output is muted.
        pub Mute);

    set_bool!(
        /// True if sound output is muted.
        pub Mute);

    get_enum!(
        /// Returns the current player state.
        pub PlayerState -> ITPlayerState);

    get_long!(
        /// Returns the player's position within the currently playing track in seconds.
        pub PlayerPosition);

    set_long!(
        /// Returns the player's position within the currently playing track in seconds.
        pub PlayerPosition);

    get_object!(
        /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
        pub CurrentEncoder -> Encoder);

    set_object!(
        /// Returns the currently selected encoder (AAC, MP3, AIFF, WAV, etc.).
        pub CurrentEncoder, Encoder);

    get_bool!(
        /// True if visuals are currently being displayed.
        pub VisualsEnabled);

    set_bool!(
        /// True if visuals are currently being displayed.
        pub VisualsEnabled);

    get_bool!(
        /// True if the visuals are displayed using the entire screen.
        pub FullScreenVisuals);

    set_bool!(
        /// True if the visuals are displayed using the entire screen.
        pub FullScreenVisuals);

    get_enum!(
        /// Returns the size of the displayed visual.
        pub VisualSize -> ITVisualSize);

    set_enum!(
        /// Returns the size of the displayed visual.
        pub VisualSize, ITVisualSize);

    get_object!(
        /// Returns the currently selected visual plug-in.
        pub CurrentVisual -> Visual);

    set_object!(
        /// Returns the currently selected visual plug-in.
        pub CurrentVisual, Visual);

    get_bool!(
        /// True if the equalizer is enabled.
        pub EQEnabled);

    set_bool!(
        /// True if the equalizer is enabled.
        pub EQEnabled);

    get_object!(
        /// Returns the currently selected EQ preset.
        pub CurrentEQPreset -> EQPreset);

    set_object!(
        /// Returns the currently selected EQ preset.
        pub CurrentEQPreset, EQPreset);

    get_bstr!(
        /// The name of the current song in the playing stream (provided by streaming server).
        pub CurrentStreamTitle);

    get_bstr!(
        /// The URL of the playing stream or streaming web site (provided by streaming server).
        pub set_CurrentStreamURL);

    get_object!(
        /// Returns the main iTunes browser window.
        pub BrowserWindow -> BrowserWindow);

    get_object!(
        /// Returns the EQ window.
        pub EQWindow -> Window);

    get_object!(
        /// Returns the source that represents the main library.
        pub LibrarySource -> Source);

    get_object!(
        /// Returns the main library playlist in the main library source.
        pub LibraryPlaylist -> LibraryPlaylist);

    get_object!(
        /// Returns the currently targeted track.
        pub CurrentTrack -> Track);

    get_object!(
        /// Returns the playlist containing the currently targeted track.
        pub CurrentPlaylist -> Playlist);

    get_object!(
        /// Returns a collection containing the currently selected track or tracks.
        pub SelectedTracks -> TrackCollection);

    get_bstr!(
        /// Returns the version of the iTunes application.
        pub Version);

    set_long!(
        /// [id(0x6002003b)]
        /// (no other documentation provided)
        pub SetOptions, no_set_prefix);

    get_object_from_str!(
        /// Start converting the specified file path.
        pub ConvertFile2(filePath) -> ConvertOperationStatus);

    get_object_from_variant!(
        /// Start converting the specified array of file paths. filePaths can be of type VT_ARRAY|VT_VARIANT, where each entry is a VT_BSTR, or VT_ARRAY|VT_BSTR.  You can also pass a JScript Array object.
        pub ConvertFiles2(filePaths) -> ConvertOperationStatus);

    get_object_from_variant!(
        /// Start converting the specified track.  iTrackToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrack.
        pub ConvertTrack2(iTrackToConvert) -> ConvertOperationStatus);

    get_object_from_variant!(
        /// Start converting the specified tracks.  iTracksToConvert is a VARIANT of type VT_DISPATCH that points to an IITTrackCollection.
        pub ConvertTracks2(iTracksToConvert) -> ConvertOperationStatus);

    get_bool!(
        /// True if iTunes will process APPCOMMAND Windows messages.
        pub AppCommandMessageProcessingEnabled);

    set_bool!(
        /// True if iTunes will process APPCOMMAND Windows messages.
        pub AppCommandMessageProcessingEnabled);

    get_bool!(
        /// True if iTunes will force itself to be the foreground application when it displays a dialog.
        pub ForceToForegroundOnDialog);

    set_bool!(
        /// True if iTunes will force itself to be the foreground application when it displays a dialog.
        pub ForceToForegroundOnDialog);

    get_object_from_str!(
        /// Create a new EQ preset.
        pub CreateEQPreset(eqPresetName) -> EQPreset);

    /// Creates a new playlist in an existing source.
    pub fn CreatePlaylistInSource(&self, playlistName: &str, source: &Source) -> windows::core::Result<Playlist> {
        str_to_bstr!(playlistName, bstr);
        let vsource = source.as_variant();
        let mut out_playlist = None;
        let result = unsafe{ self.com_object.CreatePlaylistInSource(bstr, vsource.as_raw() as *const VARIANT, &mut out_playlist as *mut _) };
        result.ok()?;

        create_wrapped_object!(Playlist, out_playlist)
    }

    /// Retrieves the current state of the player buttons.
    pub fn GetPlayerButtonsState(&self) -> windows::core::Result<PlayerButtonState> {
        let mut previousEnabled = FALSE;
        let mut playPauseStopState = ITPlayButtonState::ITPlayButtonStatePlayDisabled;
        let mut nextEnabled = FALSE;
        let result = unsafe{ self.com_object.GetPlayerButtonsState(&mut previousEnabled, &mut playPauseStopState, &mut nextEnabled) };
        result.ok()?;
        Ok(PlayerButtonState{
            previousEnabled: previousEnabled.as_bool(),
            playPauseStopState,
            nextEnabled: nextEnabled.as_bool(),
        })
    }

    /// Simulate click on a player control button.
    pub fn PlayerButtonClicked(&self, playerButton: ITPlayerButton, playerButtonModifierKeys: LONG) -> windows::core::Result<()> {
        let result = unsafe{ self.com_object.PlayerButtonClicked(playerButton, playerButtonModifierKeys) };
        result.ok()
    }

    /// True if the Shuffle property is writable for the specified playlist.
    pub fn CanSetShuffle(&self, iPlaylist: &Playlist) -> windows::core::Result<bool> {
        let vplaylist = iPlaylist.as_variant();
        let mut out_bool = FALSE;
        let result = unsafe{ self.com_object.CanSetShuffle(vplaylist.as_raw() as *const VARIANT, &mut out_bool) };
        result.ok()?;
        Ok(out_bool.as_bool())
    }

    /// True if the SongRepeat property is writable for the specified playlist.
    pub fn CanSetSongRepeat(&self, iPlaylist: &Playlist) -> windows::core::Result<bool> {
        let vplaylist = iPlaylist.as_variant();
        let mut out_bool = FALSE;
        let result = unsafe{ self.com_object.CanSetSongRepeat(vplaylist.as_raw() as *const VARIANT, &mut out_bool) };
        result.ok()?;
        Ok(out_bool.as_bool())
    }

    get_object!(
        /// Returns an IITConvertOperationStatus object if there is currently a conversion in progress.
        pub ConvertOperationStatus -> ConvertOperationStatus);

    set_bstr!(
        /// Subscribe to the specified podcast feed URL.
        pub SubscribeToPodcast, no_set_prefix);

    no_args!(
        /// Update all podcast feeds.
        pub UpdatePodcastFeeds);

    get_object_from_str!(
        /// Creates a new folder in the main library.
        pub CreateFolder(folderName) -> Playlist);

    /// Creates a new folder in an existing source.
    pub fn CreateFolderInSource(&self, folderName: &str, iSource: &Source) -> windows::core::Result<Playlist> {
        str_to_bstr!(folderName, bstr);
        let vsource = iSource.as_variant();
        let mut out_playlist = None;
        let result = unsafe{ self.com_object.CreateFolderInSource(bstr, vsource.as_raw() as *const VARIANT, &mut out_playlist as *mut _) };
        result.ok()?;

        create_wrapped_object!(Playlist, out_playlist)
    }

    get_bool!(
        /// True if the sound volume control is enabled.
        pub SoundVolumeControlEnabled);

    get_bstr!(
        /// The full path to the current iTunes library XML file.
        pub LibraryXMLPath);

    /// Returns the persistent ID of the specified IITObject.
    pub fn GetITObjectPersistentID<T>(&self, iObject: &Variant<T>) -> windows::core::Result<PersistentId> {
        let mut highID: LONG = 0;
        let mut lowID: LONG = 0;
        let result = unsafe{ self.com_object.GetITObjectPersistentIDs(iObject.as_raw() as *const VARIANT, &mut highID, &mut lowID) };
        result.ok()?;

        let bytes = [highID.to_le_bytes(), lowID.to_le_bytes()].concat();
        Ok(PersistentId::from_le_bytes(bytes.try_into().unwrap()))  // cannot panic, the slice has the correct size
    }

    get_long!(
        /// Returns the player's position within the currently playing track in milliseconds.
        pub PlayerPositionMS);

    set_long!(
        /// Returns the player's position within the currently playing track in milliseconds.
        pub PlayerPositionMS);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITAudioCDPlaylist`](crate::sys::IITAudioCDPlaylist)
    AudioCDPlaylist);

impl IITObjectWrapper for AudioCDPlaylist {}

impl IITPlaylistWrapper for AudioCDPlaylist {}

impl AudioCDPlaylist {
    get_bstr!(
        /// The artist of the CD.
        pub Artist);

    get_bool!(
        /// True if this CD is a compilation album.
        pub Compilation);

    get_bstr!(
        /// The composer of the CD.
        pub Composer);

    get_long!(
        /// The total number of discs in this CD's album.
        pub DiscCount);

    get_long!(
        /// The index of the CD disc in the source album.
        pub DiscNumber);

    get_bstr!(
        /// The genre of the CD.
        pub Genre);

    get_long!(
        /// The year the album was recorded/released.
        pub Year);

    no_args!(
        /// Reveal the CD playlist in the main browser window.
        pub Reveal);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITIPodSource`](crate::sys::IITIPodSource)
    IPodSource);

impl IITObjectWrapper for IPodSource {}

impl IPodSource {
    no_args!(
        /// Update the contents of the iPod.
        pub UpdateIPod);

    no_args!(
        /// Eject the iPod.
        pub EjectIPod);

    get_bstr!(
        /// The iPod software version.
        pub SoftwareVersion);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITFileOrCDTrack`](crate::sys::IITFileOrCDTrack)
    FileOrCDTrack);

impl IITObjectWrapper for FileOrCDTrack {}

impl IITTrackWrapper for FileOrCDTrack {}

impl FileOrCDTrack {
    get_bstr!(
        /// The full path to the file represented by this track.
        pub Location);

    no_args!(
        /// Update this track's information with the information stored in its file.
        pub UpdateInfoFromFile);

    get_bool!(
        /// True if this is a podcast track.
        pub Podcast);

    no_args!(
        /// Update the podcast feed for this track.
        pub UpdatePodcastFeed);

    get_bool!(
        /// True if playback position is remembered.
        pub RememberBookmark);

    set_bool!(
        /// True if playback position is remembered.
        pub RememberBookmark);

    get_bool!(
        /// True if track is skipped when shuffling.
        pub ExcludeFromShuffle);

    set_bool!(
        /// True if track is skipped when shuffling.
        pub ExcludeFromShuffle);

    get_bstr!(
        /// Lyrics for the track.
        pub Lyrics);

    set_bstr!(
        /// Lyrics for the track.
        pub Lyrics);

    get_bstr!(
        /// Category for the track.
        pub Category);

    set_bstr!(
        /// Category for the track.
        pub Category);

    get_bstr!(
        /// Description for the track.
        pub Description);

    set_bstr!(
        /// Description for the track.
        pub Description);

    get_bstr!(
        /// Long description for the track.
        pub LongDescription);

    set_bstr!(
        /// Long description for the track.
        pub LongDescription);

    get_long!(
        /// The bookmark time of the track (in seconds).
        pub BookmarkTime);

    set_long!(
        /// The bookmark time of the track (in seconds).
        pub BookmarkTime);

    get_enum!(
        /// The video track kind.
        pub VideoKind -> ITVideoKind);

    set_enum!(
        /// The video track kind.
        pub VideoKind, ITVideoKind);

    get_long!(
        /// The number of times the track has been skipped.
        pub SkippedCount);

    set_long!(
        /// The number of times the track has been skipped.
        pub SkippedCount);

    get_date!(
        /// The date and time the track was last skipped.  A value of zero means no skipped date.
        pub SkippedDate);

    set_date!(
        /// The date and time the track was last skipped.  A value of zero means no skipped date.
        pub SkippedDate);

    get_bool!(
        /// True if track is part of a gapless album.
        pub PartOfGaplessAlbum);

    set_bool!(
        /// True if track is part of a gapless album.
        pub PartOfGaplessAlbum);

    get_bstr!(
        /// The album artist of the track.
        pub AlbumArtist);

    set_bstr!(
        /// The album artist of the track.
        pub AlbumArtist);

    get_bstr!(
        /// The show name of the track.
        pub Show);

    set_bstr!(
        /// The show name of the track.
        pub Show);

    get_long!(
        /// The season number of the track.
        pub SeasonNumber);

    set_long!(
        /// The season number of the track.
        pub SeasonNumber);

    get_bstr!(
        /// The episode ID of the track.
        pub EpisodeID);

    set_bstr!(
        /// The episode ID of the track.
        pub EpisodeID);

    get_long!(
        /// The episode number of the track.
        pub EpisodeNumber);

    set_long!(
        /// The episode number of the track.
        pub EpisodeNumber);

    /// The size of the track (in bytes)
    pub fn Size(&self) -> windows::core::Result<i64> {
        let mut highSize = 0;
        let mut lowSize = 0;
        let result = unsafe{ self.com_object.Size64High(&mut highSize) };
        result.ok()?;
        let result = unsafe{ self.com_object.Size64Low(&mut lowSize) };
        result.ok()?;

        let bytes = [highSize.to_le_bytes(), lowSize.to_le_bytes()].concat();
        Ok(i64::from_le_bytes(bytes.try_into().unwrap())) // cannot panic, the slice has the correct size
    }

    get_bool!(
        /// True if track has not been played.
        pub Unplayed);

    set_bool!(
        /// True if track has not been played.
        pub Unplayed);

    get_bstr!(
        /// The album used for sorting.
        pub SortAlbum);

    set_bstr!(
        /// The album used for sorting.
        pub SortAlbum);

    get_bstr!(
        /// The album artist used for sorting.
        pub SortAlbumArtist);

    set_bstr!(
        /// The album artist used for sorting.
        pub SortAlbumArtist);

    get_bstr!(
        /// The artist used for sorting.
        pub SortArtist);

    set_bstr!(
        /// The artist used for sorting.
        pub SortArtist);

    get_bstr!(
        /// The composer used for sorting.
        pub SortComposer);

    set_bstr!(
        /// The composer used for sorting.
        pub SortComposer);

    get_bstr!(
        /// The track name used for sorting.
        pub SortName);

    set_bstr!(
        /// The track name used for sorting.
        pub SortName);

    get_bstr!(
        /// The show name used for sorting.
        pub SortShow);

    set_bstr!(
        /// The show name used for sorting.
        pub SortShow);

    no_args!(
        /// Reveal the track in the main browser window.
        pub Reveal);

    get_long!(
        /// The user or computed rating of the album that this track belongs to (0 to 100).
        pub AlbumRating);

    set_long!(
        /// The user or computed rating of the album that this track belongs to (0 to 100).
        pub AlbumRating);

    get_enum!(
        /// The album rating kind.
        pub AlbumRatingKind -> ITRatingKind);

    get_enum!(
        /// The track rating kind.
        pub ratingKind -> ITRatingKind);

    get_object!(
        /// Returns a collection of playlists that contain the song that this track represents.
        pub Playlists -> PlaylistCollection);

    set_bstr!(
        /// The full path to the file represented by this track.
        pub Location);

    get_date!(
        /// The release date of the track.  A value of zero means no release date.
        pub ReleaseDate);
}

com_wrapper_struct!(
    /// Safe wrapper over a [`IITPlaylistWindow`](crate::sys::IITPlaylistWindow)
    PlaylistWindow);

impl PlaylistWindow {
    get_object!(
        /// Returns a collection containing the currently selected track or tracks.
        pub SelectedTracks -> TrackCollection);

    get_object!(
        /// The playlist displayed in the window.
        pub Playlist -> Playlist);
}
