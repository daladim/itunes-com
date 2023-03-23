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
