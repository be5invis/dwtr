use core::ffi::c_void;
use core::fmt::Write;
use minidom::Element;
use std::{
    cell::RefCell,
    collections::{hash_map::Entry, HashMap},
};
use windows::{
    core::{AsImpl, IUnknown, Interface, Result},
    Win32::Foundation::BOOL,
    Win32::Graphics::{Direct2D::Common::*, DirectWrite::*},
};

use crate::svg_color::ISvgColor;

const SVG_NS: &'static str = "http://www.w3.org/2000/svg";

struct SvgGlyph {
    path_id: usize,
    offset_x: f32,
    offset_y: f32,
    co_scalar: f32,
}

impl SvgGlyph {
    fn as_element(&self) -> Element {
        Element::builder("use", SVG_NS)
            .attr("href", format!("#path{}", self.path_id))
            .attr(
                "transform",
                format!(
                    "translate({} {}) scale({})",
                    self.offset_x, self.offset_y, self.co_scalar
                ),
            )
            .build()
    }
}

struct SvgRun {
    offset_x: f32,
    offset_y: f32,
    rotate_angle: f32,
    color: Option<String>,
    glyphs: Vec<SvgGlyph>,
}
impl SvgRun {
    fn as_element(&self) -> Element {
        Element::builder("g", SVG_NS)
            .attr(
                "transform",
                format!(
                    "translate({} {}) rotate({})",
                    self.offset_x, self.offset_y, self.rotate_angle
                ),
            )
            .attr("fill", self.color.clone().unwrap_or(String::from("black")))
            .append_all(self.glyphs.iter().map(|g| g.as_element()))
            .build()
    }
}

struct SvgDataStorage {
    last_path_id: usize,
    path_defs: HashMap<String, usize>,
    runs: Vec<SvgRun>,
}

impl SvgDataStorage {
    fn new() -> Self {
        Self {
            last_path_id: 0,
            path_defs: HashMap::new(),
            runs: Vec::new(),
        }
    }
    fn add_path_def(&mut self, str: String) -> usize {
        match self.path_defs.entry(str) {
            Entry::Occupied(o) => *o.get(),
            Entry::Vacant(v) => {
                self.last_path_id += 1;
                v.insert(self.last_path_id);
                self.last_path_id
            }
        }
    }
}

#[windows::core::implement(IDWriteTextRenderer1)]
pub(crate) struct SvgTextRenderer {
    canvas_width: f32,
    canvas_height: f32,
    offset_x: RefCell<f32>,
    offset_y: RefCell<f32>,
    store: RefCell<SvgDataStorage>,
}

impl SvgTextRenderer {
    pub(crate) fn new(canvas_width: f32, canvas_height: f32) -> Self {
        Self {
            canvas_width,
            canvas_height,
            offset_x: RefCell::new(0.0),
            offset_y: RefCell::new(0.0),
            store: RefCell::new(SvgDataStorage::new()),
        }
    }

    pub(crate) fn set_offset(&self, x: f32, y: f32) {
        self.offset_x.replace(x);
        self.offset_y.replace(y);
    }

    pub(crate) fn into_xml(&self) -> Element {
        let store = self.store.borrow();

        let defs = Element::builder("defs", SVG_NS)
            .append_all(store.path_defs.iter().map(|(path, id)| {
                Element::builder("path", SVG_NS)
                    .attr("id", format!("path{}", id))
                    .attr("d", path)
                    .build()
            }))
            .build();

        let glyphs = Element::builder("g", SVG_NS)
            .append_all(store.runs.iter().map(|g| g.as_element()))
            .build();

        Element::builder("svg", SVG_NS)
            .attr(
                "viewBox",
                format!("0 0 {} {}", self.canvas_width, self.canvas_height),
            )
            .append(defs)
            .append(glyphs)
            .build()
    }

    fn get_color_from_brush(brush: &Option<IUnknown>) -> Option<String> {
        match brush {
            Some(brush) => match brush.cast::<ISvgColor>() {
                Ok(color) => {
                    let mut sink = csscolorparser::Color::default();
                    unsafe {
                        color
                            .GetColor(&mut sink.r, &mut sink.g, &mut sink.b, &mut sink.a)
                            .unwrap()
                    };
                    Some(sink.to_hex_string())
                }
                _ => None,
            },
            _ => None,
        }
    }
    fn add_path_def(&self, str: String) -> usize {
        let mut store = self.store.borrow_mut();
        store.add_path_def(str)
    }
    fn push_run(&self, run: SvgRun) {
        self.store.borrow_mut().runs.push(run);
    }
}

#[allow(non_snake_case)]
impl IDWritePixelSnapping_Impl for SvgTextRenderer {
    fn IsPixelSnappingDisabled(&self, _client_drawing_context: *const c_void) -> Result<BOOL> {
        Ok(false.into())
    }
    fn GetCurrentTransform(&self, _client_drawing_context: *const c_void) -> Result<DWRITE_MATRIX> {
        Ok(DWRITE_MATRIX {
            m11: 1.0,
            m12: 1.0,
            m21: 1.0,
            m22: 1.0,
            dx: 0.0,
            dy: 0.0,
        })
    }
    fn GetPixelsPerDip(&self, _client_drawing_context: *const c_void) -> Result<f32> {
        Ok(1.0)
    }
}

#[allow(non_snake_case)]
impl IDWriteTextRenderer_Impl for SvgTextRenderer {
    fn DrawGlyphRun(
        &self,
        client_drawing_context: *const c_void,
        baseline_origin_x: f32,
        baseline_origin_y: f32,
        measuring_mode: DWRITE_MEASURING_MODE,
        glyph_run: *const DWRITE_GLYPH_RUN,
        glyph_run_description: *const DWRITE_GLYPH_RUN_DESCRIPTION,
        client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        self.DrawGlyphRun2(
            client_drawing_context,
            baseline_origin_x,
            baseline_origin_y,
            DWRITE_GLYPH_ORIENTATION_ANGLE_0_DEGREES,
            measuring_mode,
            glyph_run,
            glyph_run_description,
            client_drawing_effect,
        )
    }

    fn DrawInlineObject(
        &self,
        _client_drawing_context: *const c_void,
        _origin_x: f32,
        _origin_y: f32,
        _inline_object: &Option<IDWriteInlineObject>,
        _is_sideways: BOOL,
        _is_right_to_left: BOOL,
        _client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawUnderline(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _underline: *const DWRITE_UNDERLINE,
        _client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawStrikethrough(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _strike_through: *const DWRITE_STRIKETHROUGH,
        _client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        Ok(())
    }
}

#[allow(non_snake_case)]
impl IDWriteTextRenderer1_Impl for SvgTextRenderer {
    fn DrawGlyphRun2(
        &self,
        _client_drawing_context: *const c_void,
        baseline_origin_x: f32,
        baseline_origin_y: f32,
        orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _measuring_mode: DWRITE_MEASURING_MODE,
        glyph_run: *const DWRITE_GLYPH_RUN,
        _glyph_run_description: *const DWRITE_GLYPH_RUN_DESCRIPTION,
        client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        if let Some(font_face) = unsafe { (*glyph_run).fontFace.clone() } {
            let mut metrics = DWRITE_FONT_METRICS::default();
            unsafe { font_face.GetMetrics(&mut metrics) }

            let glyph_count = unsafe { (*glyph_run).glyphCount };
            let color = Self::get_color_from_brush(client_drawing_effect);

            let mut run = SvgRun {
                offset_x: baseline_origin_x + *self.offset_x.borrow(),
                offset_y: baseline_origin_y + *self.offset_y.borrow(),
                rotate_angle: dw_angle_to_angle(&orientation_angle, unsafe {
                    (*glyph_run).isSideways.as_bool()
                }),
                color,
                glyphs: Vec::new(),
            };

            let scalar = (metrics.designUnitsPerEm as f32) / unsafe { (*glyph_run).fontEmSize };
            let co_scalar = 1.0 / scalar;

            let geometry_sink: ID2D1SimplifiedGeometrySink = SvgGeometrySink::new(scalar).into();
            let geometry_sink_impl = geometry_sink.as_impl();

            let mut offset_x = 0.0;
            let offset_y = 0.0;

            for i in 0..glyph_count {
                unsafe {
                    let p_glyph_offset = (*glyph_run).glyphOffsets;
                    let p_glyph_offset = if p_glyph_offset.is_null() {
                        p_glyph_offset
                    } else {
                        p_glyph_offset.offset(i as isize)
                    };

                    font_face.GetGlyphRunOutline(
                        (*glyph_run).fontEmSize,
                        (*glyph_run).glyphIndices.offset(i as isize),
                        (*glyph_run).glyphAdvances.offset(i as isize),
                        p_glyph_offset,
                        1,
                        (*glyph_run).isSideways,
                        (*glyph_run).bidiLevel % 2 == 1,
                        geometry_sink.clone(),
                    )?;
                }

                let path_id = self.add_path_def(geometry_sink_impl.reset());
                run.glyphs.push(SvgGlyph {
                    path_id,
                    offset_x,
                    offset_y,
                    co_scalar,
                });

                unsafe {
                    let direction = if (*glyph_run).bidiLevel % 2 == 1 {
                        -1.0
                    } else {
                        1.0
                    };
                    offset_x += direction * *((*glyph_run).glyphAdvances.offset(i as isize));
                }
            }

            self.push_run(run);
        }

        Ok(())
    }

    fn DrawInlineObject2(
        &self,
        _client_drawing_context: *const c_void,
        _origin_x: f32,
        _origin_y: f32,
        _orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _inline_object: &Option<IDWriteInlineObject>,
        _is_sideways: BOOL,
        _is_right_to_left: BOOL,
        _client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawUnderline2(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _underline: *const DWRITE_UNDERLINE,
        _client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawStrikethrough2(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _strike_through: *const DWRITE_STRIKETHROUGH,
        _client_drawing_effect: &Option<IUnknown>,
    ) -> Result<()> {
        Ok(())
    }
}

/// Geometry sink that constructs a <path>
#[windows::core::implement(ID2D1SimplifiedGeometrySink)]
pub(crate) struct SvgGeometrySink {
    scalar: f32,
    body: RefCell<String>,
}

impl SvgGeometrySink {
    fn new(scalar: f32) -> Self {
        Self {
            scalar,
            body: RefCell::new(String::new()),
        }
    }

    fn reset(&self) -> String {
        self.body.replace(String::new())
    }
}

#[allow(non_snake_case)]
impl ID2D1SimplifiedGeometrySink_Impl for SvgGeometrySink {
    fn SetFillMode(&self, _fill_mode: D2D1_FILL_MODE) {}
    fn SetSegmentFlags(&self, _flags: D2D1_PATH_SEGMENT) {}
    fn BeginFigure(&self, start_point: &D2D_POINT_2F, _figure_begin: D2D1_FIGURE_BEGIN) {
        write!(
            self.body.borrow_mut(),
            "M {} {} ",
            start_point.x * self.scalar,
            start_point.y * self.scalar
        )
        .unwrap()
    }
    fn AddLines(&self, points: *const D2D_POINT_2F, points_count: u32) {
        let mut sink = self.body.borrow_mut();
        for i in 0..points_count {
            unsafe {
                let point = points.offset(i as isize);
                write!(
                    sink,
                    "L {} {} ",
                    (*point).x * self.scalar,
                    (*point).y * self.scalar
                )
                .unwrap()
            }
        }
    }
    fn AddBeziers(&self, beziers: *const D2D1_BEZIER_SEGMENT, beziers_count: u32) {
        let mut sink = self.body.borrow_mut();
        for i in 0..beziers_count {
            unsafe {
                let curve = beziers.offset(i as isize);
                write!(
                    sink,
                    "C {} {} {} {} {} {} ",
                    (*curve).point1.x * self.scalar,
                    (*curve).point1.y * self.scalar,
                    (*curve).point2.x * self.scalar,
                    (*curve).point2.y * self.scalar,
                    (*curve).point3.x * self.scalar,
                    (*curve).point3.y * self.scalar
                )
                .unwrap()
            }
        }
    }
    fn EndFigure(&self, figure_end: D2D1_FIGURE_END) {
        if figure_end == D2D1_FIGURE_END_CLOSED {
            write!(self.body.borrow_mut(), "Z ").unwrap();
        }
    }
    fn Close(&self) -> Result<()> {
        Ok(())
    }
}

fn dw_angle_to_angle(angle: &DWRITE_GLYPH_ORIENTATION_ANGLE, is_sideways: bool) -> f32 {
    let mut quarters = match angle {
        &DWRITE_GLYPH_ORIENTATION_ANGLE_0_DEGREES => 0,
        &DWRITE_GLYPH_ORIENTATION_ANGLE_90_DEGREES => 1,
        &DWRITE_GLYPH_ORIENTATION_ANGLE_180_DEGREES => 2,
        &DWRITE_GLYPH_ORIENTATION_ANGLE_270_DEGREES => 3,
        _ => unreachable!(),
    };
    if is_sideways {
        quarters = (1 + quarters) % 4
    }
    90.0 * (quarters as f32)
}
