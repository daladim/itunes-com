use std::marker::PhantomData;

use windows::Win32::System::Com::VARIANT;


pub type PersistentId = u64;


/// A wrapper around a COM VARIANT type
pub struct Variant<'a, T: 'a> {
    inner: VARIANT,
    lifetime: PhantomData<&'a T>,
}

impl<'a, T> Variant<'a, T> {
    pub(crate) fn new(inner: VARIANT) -> Self {
        Self { inner, lifetime: PhantomData }
    }

    /// Get the wrapped `VARIANT`
    pub fn as_raw(&self) -> &VARIANT {
        &self.inner
    }
}

/// The rating of a track (one to five stars)
pub enum Rating {
    /// No rating
    None,
    /// One-star rating
    One,
    /// Two-star rating
    Two,
    /// Three-star rating
    Three,
    /// Four-star rating
    Four,
    /// Five-star rating
    Five,
}

impl Rating {
    /// Create a `Rating` instance.
    ///
    /// All invalid values are mapped to `Rating::None`
    pub fn from_stars(stars: Option<u8>) -> Self {
        match stars {
            Some(1) => Self::One,
            Some(2) => Self::Two,
            Some(3) => Self::Three,
            Some(4) => Self::Four,
            Some(5) => Self::Five,
            _ => Self::None,
        }
    }

    /// Get the count of stars of this rating
    pub fn stars(&self) -> Option<u8> {
        match self {
            Rating::None => None,
            Rating::One => Some(1),
            Rating::Two => Some(2),
            Rating::Three => Some(3),
            Rating::Four => Some(4),
            Rating::Five => Some(5),
        }
    }
}

// Conversion with LONG because that's what iTunes uses.
impl std::convert::From<super::LONG> for Rating {
    fn from(long: super::LONG) -> Rating {
        match long {
            0..=19 => Rating::None,
            20..=39 => Rating::One,
            40..=59 => Rating::Two,
            60..=79 => Rating::Three,
            80..=99 => Rating::Four,
            100.. => Rating::Five,
            // Not supposed to happen
            _ => Rating::None,
        }
    }
}

impl std::convert::From<Rating> for super::LONG {
    fn from(rating: Rating) -> super::LONG {
        match rating {
            Rating::None => 0,
            Rating::One => 20,
            Rating::Two => 40,
            Rating::Three => 60,
            Rating::Four => 80,
            Rating::Five => 100,
        }
    }
}
