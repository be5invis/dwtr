#![allow(non_snake_case)]

use windows::core::{implement, interface, IUnknown, IUnknownVtbl, HRESULT};
use windows::Win32::Foundation::S_OK;

/// My interface
#[interface("f2496799-9fb3-4933-96c4-46c7ab425974")]
pub(crate) unsafe trait ISvgColor: IUnknown {
    pub(crate) unsafe fn GetColor(
        &self,
        r: *mut f64,
        g: *mut f64,
        b: *mut f64,
        a: *mut f64,
    ) -> HRESULT;
}

#[implement(ISvgColor)]
pub(crate) struct SvgColorImpl {
    color: csscolorparser::Color,
}
impl SvgColorImpl {
    pub(crate) fn new(color: csscolorparser::Color) -> Self {
        Self { color }
    }
}
impl ISvgColor_Impl for SvgColorImpl {
    unsafe fn GetColor(&self, pr: *mut f64, pg: *mut f64, pb: *mut f64, pa: *mut f64) -> HRESULT {
        let (r, g, b, a) = self.color.rgba();
        *pr = r;
        *pg = g;
        *pb = b;
        *pa = a;
        S_OK
    }
}
