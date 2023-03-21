#![allow(non_snake_case)]

use std::sync::Arc;

use criterion::{criterion_group, criterion_main, Criterion};
use windows::Win32::System::Com::{CoInitializeEx, CoCreateInstance, CLSCTX_ALL, COINIT_MULTITHREADED};


pub fn criterion_benchmark(c: &mut Criterion) {
    // On my machine:
    //
    // cloning a COM instance  time:   [19.372 ns 19.612 ns 19.908 ns]
    //                         change: [+0.0486% +2.2397% +4.0036%] (p = 0.03 < 0.05)
    //                         Change within noise threshold.
    // Found 8 outliers among 100 measurements (8.00%)
    //   4 (4.00%) high mild
    //   4 (4.00%) high severe
    //
    // cloning a Rust Arc      time:   [10.872 ns 10.882 ns 10.897 ns]
    //                         change: [-0.3523% -0.0297% +0.3490%] (p = 0.89 > 0.05)
    //                         No change in performance detected.
    // Found 14 outliers among 100 measurements (14.00%)
    //   6 (6.00%) high mild
    //   8 (8.00%) high severe

    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
    }

    let iTunes_com: itunes_com::sys::IiTunes = unsafe { CoCreateInstance(&itunes_com::sys::ITUNES_APP_COM_GUID, None, CLSCTX_ALL).unwrap() };

    c.bench_function("cloning a COM instance", |b| b.iter(|| {
        iTunes_com.clone()
    }));

    let arc = Arc::new(iTunes_com);
    c.bench_function("cloning a Rust Arc", |b| b.iter(|| {
        Arc::clone(&arc)
    }));
}

criterion_group!(benches, criterion_benchmark);
criterion_main!(benches);
