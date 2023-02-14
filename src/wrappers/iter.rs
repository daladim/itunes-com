//! Wrapper over COM iteration ability
//!
//! I could not manage to make the COM API `_NewEnum` work, so I'm implementing iterators myself instead

use std::marker::PhantomData;

use crate::wrappers::Iterable;
use super::LONG;

pub struct Iterator<'a, Obj, Item> {
    data: &'a Obj,
    count: LONG,
    current: LONG,
    items: PhantomData<Item>,
}

impl<'a, Obj, Items> Iterator<'a, Obj, Items>
where Obj: Iterable + Iterable<Item = Items>
{
    pub(crate) fn new(data: &'a Obj) -> windows::core::Result<Self> {
        let count = data.Count()?;

        Ok(Self {
            data,
            count,
            current: 0,
            items: PhantomData,
        })
    }
}

impl<'a, Obj, Items> std::iter::Iterator for Iterator<'a, Obj, Items>
where Obj: Iterable + Iterable<Item = Items>
{
    type Item = Items;

    fn next(&mut self) -> Option<Self::Item> {
        // COM iterators (or at least iTunes iterators) are 1-based.
        // Let's increment the index _before_ we access it.
        self.current += 1;
        self.data.item(self.current).ok()
    }
}

impl<'a, Obj, Items> std::iter::ExactSizeIterator for Iterator<'a, Obj, Items>
where Obj: Iterable + Iterable<Item = Items>
{
    fn len(&self) -> usize {
        self.count as usize
    }
}

impl<'a, Obj, Items> std::iter::DoubleEndedIterator for Iterator<'a, Obj, Items>
where Obj: Iterable + Iterable<Item = Items>
{
    fn next_back(&mut self) -> Option<Self::Item> {
        self.current -= 1;
        self.data.item(self.current).ok()
    }
}

