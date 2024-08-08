use core::ffi::c_void;
use core::fmt::Write;
use indexmap::{map::Entry, IndexMap};
use std::{cell::RefCell, rc::Rc};
use svg::{node::element, Document, Node};
use windows::{
    core::{AsImpl, IUnknown, Interface, Result},
    Win32::Foundation::BOOL,
    Win32::Graphics::{Direct2D::Common::*, DirectWrite::*},
};

use crate::{escape::escape_str, svg_color::ISvgColor};

struct SvgGlyph {
    path_id: usize,
    offset_x: f32,
    offset_y: f32,
}

impl SvgGlyph {
    fn as_element(&self) -> element::Use {
        element::Use::new()
            .set("href", format!("#path{}", self.path_id))
            .set(
                "transform",
                format!("translate({} {})", self.offset_x, self.offset_y,),
            )
    }
}

struct SvgRun {
    offset_x: f32,
    offset_y: f32,
    rotate_angle: f32,
    upm: f32,
    scalar: f32,
    color: Option<String>,
    source_text: String,
    glyphs: Vec<SvgGlyph>,
    copyable: bool,
}
impl SvgRun {
    fn as_element(&self) -> element::Group {
        let mut g = element::Group::new()
            .set(
                "transform",
                format!(
                    "translate({} {}) rotate({}) scale({})",
                    self.offset_x,
                    self.offset_y,
                    self.rotate_angle,
                    1.0 / self.scalar
                ),
            )
            .set("fill", self.color.clone().unwrap_or(String::from("black")))
            .set("data-source-text", escape_str(&self.source_text));

        if self.copyable {
            let mut text_element = element::Text::new()
                .set("x", self.offset_x)
                .set("y", self.offset_y)
                .set("font-size", self.upm)
                .set("fill", "transparent");
            text_element.append(svg::node::Text::new(escape_str(&self.source_text)));
            g.append(text_element)
        }
        for glyph in &self.glyphs {
            g.append(glyph.as_element());
        }
        g
    }
}

pub(crate) struct SvgFrame {
    runs: Vec<SvgRun>,
    copyable: bool,
    frame_title: Option<String>,
    frame_desc: Option<String>,
}

impl SvgFrame {
    pub(crate) fn new() -> Self {
        Self {
            runs: Vec::new(),
            copyable: false,
            frame_desc: None,
            frame_title: None,
        }
    }

    fn as_element(&self) -> element::Group {
        let mut g = element::Group::new();
        if let Some(title) = &self.frame_title {
            g.append(element::Title::new().add(svg::node::Text::new(escape_str(title))));
        }
        if let Some(desc) = &self.frame_desc {
            g.append(element::Description::new().add(svg::node::Text::new(escape_str(desc))));
        }
        for run in &self.runs {
            g.append(run.as_element());
        }
        g
    }
}

pub(crate) struct SharedStore {
    last_path_id: usize,
    path_defs: IndexMap<String, usize>,
}

impl SharedStore {
    pub(crate) fn new() -> Self {
        Self {
            last_path_id: 0,
            path_defs: IndexMap::new(),
        }
    }

    pub(crate) fn add_path_def(&mut self, str: String) -> usize {
        if str.is_empty() {
            return 0;
        }
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

pub(crate) struct SvgDocumentRenderer {
    canvas_width: f32,
    canvas_height: f32,
    shared_store: Rc<RefCell<SharedStore>>,
    frames: Vec<Rc<RefCell<SvgFrame>>>,
}

impl SvgDocumentRenderer {
    pub(crate) fn new(canvas_width: f32, canvas_height: f32) -> Self {
        Self {
            canvas_width,
            canvas_height,
            shared_store: Rc::new(RefCell::new(SharedStore::new())),
            frames: Vec::new(),
        }
    }

    pub(crate) fn create_frame_renderer(
        &mut self,
        offset_x: f32,
        offset_y: f32,
    ) -> SvgFrameRenderer {
        let frame_store = Rc::new(RefCell::new(SvgFrame::new()));
        let frame_renderer = SvgFrameRenderer::new(
            self.shared_store.clone(),
            frame_store.clone(),
            offset_x,
            offset_y,
        );
        self.frames.push(frame_store);
        frame_renderer
    }

    pub(crate) fn into_xml(&self) -> Document {
        let store = self.shared_store.borrow();

        let mut defs = element::Definitions::new();
        for (path_d, id) in &store.path_defs {
            let path = element::Path::new()
                .set("id", format!("path{}", id))
                .set("d", path_d.clone());
            defs.append(path);
        }

        let mut svg = Document::new()
            .set(
                "viewBox",
                format!("0 0 {} {}", self.canvas_width, self.canvas_height),
            )
            .set("width", self.canvas_width)
            .set("height", self.canvas_height)
            .add(defs);

        for frame in &self.frames {
            svg.append(frame.borrow().as_element());
        }

        svg
    }
}

#[windows::core::implement(IDWriteTextRenderer1)]
pub(crate) struct SvgFrameRenderer {
    shared_store: Rc<RefCell<SharedStore>>,
    frame_store: Rc<RefCell<SvgFrame>>,
    // frame properties
    offset_x: f32,
    offset_y: f32,
}

impl SvgFrameRenderer {
    pub(crate) fn new(
        shared_store: Rc<RefCell<SharedStore>>,
        frame_store: Rc<RefCell<SvgFrame>>,
        offset_x: f32,
        offset_y: f32,
    ) -> Self {
        Self {
            shared_store,
            frame_store,
            offset_x,
            offset_y,
        }
    }

    pub(crate) fn set_title(&self, title: Option<String>) {
        self.frame_store.borrow_mut().frame_title = title;
    }
    pub(crate) fn set_desc(&self, desc: Option<String>) {
        self.frame_store.borrow_mut().frame_desc = desc;
    }
    pub(crate) fn set_copyable(&self, copyable: bool) {
        self.frame_store.borrow_mut().copyable = copyable;
    }

    fn get_color_from_brush(&self, brush: Option<&IUnknown>) -> Option<String> {
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
        self.shared_store.borrow_mut().add_path_def(str)
    }
    fn push_run(&self, run: SvgRun) {
        self.frame_store.borrow_mut().runs.push(run);
    }
}

#[allow(non_snake_case)]
impl IDWritePixelSnapping_Impl for SvgFrameRenderer_Impl {
    fn IsPixelSnappingDisabled(&self, _client_drawing_context: *const c_void) -> Result<BOOL> {
        Ok(false.into())
    }
    fn GetCurrentTransform(
        &self,
        _client_drawing_context: *const core::ffi::c_void,
        transform: *mut DWRITE_MATRIX,
    ) -> windows::core::Result<()> {
        unsafe {
            *transform = DWRITE_MATRIX {
                m11: 1.0,
                m12: 1.0,
                m21: 1.0,
                m22: 1.0,
                dx: 0.0,
                dy: 0.0,
            };
        }
        Ok(())
    }
    fn GetPixelsPerDip(&self, _client_drawing_context: *const c_void) -> Result<f32> {
        Ok(1.0)
    }
}

#[allow(non_snake_case)]
impl IDWriteTextRenderer_Impl for SvgFrameRenderer_Impl {
    fn DrawGlyphRun(
        &self,
        client_drawing_context: *const c_void,
        baseline_origin_x: f32,
        baseline_origin_y: f32,
        measuring_mode: DWRITE_MEASURING_MODE,
        glyph_run: *const DWRITE_GLYPH_RUN,
        glyph_run_description: *const DWRITE_GLYPH_RUN_DESCRIPTION,
        client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        IDWriteTextRenderer1_Impl::DrawGlyphRun(
            self,
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
        _inline_object: Option<&IDWriteInlineObject>,
        _is_sideways: BOOL,
        _is_right_to_left: BOOL,
        _client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawUnderline(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _underline: *const DWRITE_UNDERLINE,
        _client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawStrikethrough(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _strike_through: *const DWRITE_STRIKETHROUGH,
        _client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        Ok(())
    }
}

#[allow(non_snake_case)]
impl IDWriteTextRenderer1_Impl for SvgFrameRenderer_Impl {
    fn DrawGlyphRun(
        &self,
        _client_drawing_context: *const c_void,
        baseline_origin_x: f32,
        baseline_origin_y: f32,
        orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _measuring_mode: DWRITE_MEASURING_MODE,
        glyph_run: *const DWRITE_GLYPH_RUN,
        glyph_run_description: *const DWRITE_GLYPH_RUN_DESCRIPTION,
        client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        if let Some(font_face) = unsafe { (*glyph_run).fontFace.as_ref() } {
            let mut metrics = DWRITE_FONT_METRICS::default();
            unsafe { font_face.GetMetrics(&mut metrics) }

            let glyph_count = unsafe { (*glyph_run).glyphCount };
            let color = self.get_color_from_brush(client_drawing_effect);

            let scalar = (metrics.designUnitsPerEm as f32) / unsafe { (*glyph_run).fontEmSize };

            let mut run = SvgRun {
                offset_x: baseline_origin_x + self.offset_x,
                offset_y: baseline_origin_y + self.offset_y,
                rotate_angle: dw_angle_to_angle(&orientation_angle, unsafe {
                    (*glyph_run).isSideways.as_bool()
                }),
                upm: metrics.designUnitsPerEm as f32,
                scalar,
                color,
                source_text: unsafe {
                    String::from_utf16_lossy(std::slice::from_raw_parts(
                        (*glyph_run_description).string.0,
                        (*glyph_run_description).stringLength as usize,
                    ))
                },
                glyphs: Vec::new(),
                copyable: self.frame_store.borrow().copyable,
            };

            let geometry_sink: ID2D1SimplifiedGeometrySink = SvgGeometrySink::new(scalar).into();
            let geometry_sink_impl = unsafe { geometry_sink.as_impl() };

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
                        Some((*glyph_run).glyphAdvances.offset(i as isize)),
                        Some(p_glyph_offset),
                        1,
                        (*glyph_run).isSideways,
                        (*glyph_run).bidiLevel % 2 == 1,
                        &geometry_sink,
                    )?;
                }

                let path_id = self.add_path_def(geometry_sink_impl.reset());
                if path_id > 0 {
                    run.glyphs.push(SvgGlyph {
                        path_id,
                        offset_x: geometry_sink_impl.process_coord(offset_x),
                        offset_y: geometry_sink_impl.process_coord(offset_y),
                    });
                }

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

    fn DrawInlineObject(
        &self,
        _client_drawing_context: *const c_void,
        _origin_x: f32,
        _origin_y: f32,
        _orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _inline_object: Option<&IDWriteInlineObject>,
        _is_sideways: BOOL,
        _is_right_to_left: BOOL,
        _client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawUnderline(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _underline: *const DWRITE_UNDERLINE,
        _client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        Ok(())
    }

    fn DrawStrikethrough(
        &self,
        _client_drawing_context: *const c_void,
        _baseline_origin_x: f32,
        _baseline_origin_y: f32,
        _orientation_angle: DWRITE_GLYPH_ORIENTATION_ANGLE,
        _strike_through: *const DWRITE_STRIKETHROUGH,
        _client_drawing_effect: Option<&IUnknown>,
    ) -> Result<()> {
        Ok(())
    }
}

/// Geometry sink that constructs a <path>
#[windows::core::implement(ID2D1SimplifiedGeometrySink)]
pub(crate) struct SvgGeometrySink {
    scalar: f32,
    body: RefCell<String>,
    last_x: RefCell<f32>,
    last_y: RefCell<f32>,
}

const COORD_RESOLUTION: f32 = 0x100 as f32;

impl SvgGeometrySink {
    fn new(scalar: f32) -> Self {
        Self {
            scalar,
            body: RefCell::new(String::new()),
            last_x: RefCell::new(0.0),
            last_y: RefCell::new(0.0),
        }
    }

    fn reset(&self) -> String {
        self.body.replace(String::new())
    }

    fn process_coord(&self, f: f32) -> f32 {
        (f * self.scalar * COORD_RESOLUTION).round() / COORD_RESOLUTION
    }

    fn set_last_point(&self, x: f32, y: f32) {
        self.last_x.replace(x);
        self.last_y.replace(y);
    }
}

#[allow(non_snake_case)]
impl ID2D1SimplifiedGeometrySink_Impl for SvgGeometrySink_Impl {
    fn SetFillMode(&self, _fill_mode: D2D1_FILL_MODE) {}
    fn SetSegmentFlags(&self, _flags: D2D1_PATH_SEGMENT) {}
    fn BeginFigure(&self, start_point: &D2D_POINT_2F, _figure_begin: D2D1_FIGURE_BEGIN) {
        let cx_orig = start_point.x;
        let cy_orig = start_point.y;
        {
            let cx = self.process_coord(cx_orig);
            let cy = self.process_coord(cy_orig);
            write!(self.body.borrow_mut(), "M {} {} ", cx, cy).unwrap();
        }
        self.set_last_point(cx_orig, cy_orig);
    }
    fn AddLines(&self, points: *const D2D_POINT_2F, points_count: u32) {
        let mut sink = self.body.borrow_mut();
        for i in 0..points_count {
            unsafe {
                let point = points.offset(i as isize);
                let cx_orig = (*point).x;
                let cy_orig = (*point).y;
                {
                    let cx = self.process_coord(cx_orig);
                    let cy = self.process_coord(cy_orig);
                    write!(sink, "L {} {} ", cx, cy).unwrap();
                }
                self.set_last_point(cx_orig, cy_orig);
            }
        }
    }
    fn AddBeziers(&self, beziers: *const D2D1_BEZIER_SEGMENT, beziers_count: u32) {
        let mut sink = self.body.borrow_mut();
        for i in 0..beziers_count {
            unsafe {
                let curve = beziers.offset(i as isize);

                let x0_orig = self.last_x.borrow().clone();
                let y0_orig = self.last_y.borrow().clone();
                let x1_orig = (*curve).point1.x;
                let y1_orig = (*curve).point1.y;
                let x2_orig = (*curve).point2.x;
                let y2_orig = (*curve).point2.y;
                let x3_orig = (*curve).point3.x;
                let y3_orig = (*curve).point3.y;

                let xm1 = x0_orig + (x1_orig - x0_orig) * 1.5;
                let ym1 = y0_orig + (y1_orig - y0_orig) * 1.5;
                let xm2 = x3_orig + (x2_orig - x3_orig) * 1.5;
                let ym2 = y3_orig + (y2_orig - y3_orig) * 1.5;
                {
                    let x1 = self.process_coord(x1_orig);
                    let y1 = self.process_coord(y1_orig);
                    let x2 = self.process_coord(x2_orig);
                    let y2 = self.process_coord(y2_orig);
                    let x3 = self.process_coord(x3_orig);
                    let y3 = self.process_coord(y3_orig);

                    if COORD_RESOLUTION * (xm2 - xm1).abs() < 1.0
                        && COORD_RESOLUTION * (ym2 - ym1).abs() < 1.0
                    {
                        let xm = self.process_coord((xm1 + xm2) / 2.0);
                        let ym = self.process_coord((ym1 + ym2) / 2.0);
                        write!(sink, "Q {} {} {} {} ", xm, ym, x3, y3).unwrap();
                    } else {
                        write!(sink, "C {} {} {} {} {} {} ", x1, y1, x2, y2, x3, y3).unwrap();
                    }
                }
                self.set_last_point(x3_orig, y3_orig);
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
