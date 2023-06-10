use windows::{
    core::{ComInterface, IUnknown, Result, HSTRING, PCWSTR},
    Win32::Graphics::DirectWrite::*,
};

use crate::{
    document::{
        string_to_tag, DocumentContent, DocumentFrame, FontVariationValue, HAlign, TextStyle,
        VAlign,
    },
    svg_color::{ISvgColor, SvgColorImpl},
};

#[derive(Debug)]
pub(crate) struct DocumentAnalyzer {
    // Text encoded in UTF-16
    text: Vec<u16>,
    style_runs: Vec<StyleRun>,
    style_stack: Vec<TextStyle>,
}

impl DocumentAnalyzer {
    pub(crate) fn new() -> Self {
        Self {
            text: Vec::new(),
            style_runs: Vec::new(),
            style_stack: Vec::new(),
        }
    }
    fn sync_style_run_length(&mut self) {
        if let Some(style_run) = self.style_runs.last_mut() {
            style_run.wch_end = self.text.len();
        }
    }
    fn start_style_run(&mut self, style: TextStyle) {
        self.style_runs.push(StyleRun {
            wch_start: self.text.len(),
            wch_end: self.text.len(),
            style,
        })
    }
    pub(crate) fn analyze(&mut self, dc: &DocumentContent) {
        match dc {
            DocumentContent::Text(tx) => {
                let ws: Vec<u16> = tx.encode_utf16().collect();
                self.text.extend(ws);
                self.sync_style_run_length()
            }
            DocumentContent::Style(s) => {
                if let Some(style) = self.style_stack.last_mut() {
                    style.merge(s);
                };
                self.start_style_run(self.style_stack.last().cloned().unwrap_or_default());
            }
            DocumentContent::Embed(sub) => {
                self.sync_style_run_length();
                let last_style = self.style_stack.last().cloned().unwrap_or_default();
                {
                    let inner_style = last_style.clone();
                    self.style_stack.push(inner_style.clone());
                    self.start_style_run(inner_style);
                    for item in sub.iter() {
                        self.analyze(item)
                    }
                    self.style_stack.pop();
                }
                self.start_style_run(last_style);
            }
        }
    }

    pub(crate) fn create_text_layout(
        &self,
        factory: IDWriteFactory,
        format: IDWriteTextFormat,
        canvas_width: f32,
        canvas_height: f32,
        frame: &DocumentFrame,
    ) -> Result<IDWriteTextLayout1> {
        let layout = unsafe {
            factory.CreateTextLayout(
                &self.text,
                &format,
                frame.right.unwrap_or(canvas_width) - frame.left.unwrap_or(0.0),
                frame.bottom.unwrap_or(canvas_height) - frame.top.unwrap_or(0.0),
            )?
        };
        let layout: IDWriteTextLayout1 = layout.cast()?;

        // Set alignment
        unsafe { layout.SetTextAlignment(frame.text_align.clone().into())? };
        // Set direction
        let (read_dir, flow_dir) = frame.writing_mode.clone().into();
        unsafe { layout.SetReadingDirection(read_dir)? }
        unsafe { layout.SetFlowDirection(flow_dir)? }
        unsafe {
            layout.SetLineSpacing(
                DWRITE_LINE_SPACING_METHOD_PROPORTIONAL,
                frame.line_height,
                frame.line_height * frame.baseline_offset,
            )?
        }

        // Set text styles
        for style_run in self.style_runs.iter() {
            if style_run.wch_end <= style_run.wch_start {
                continue;
            }
            let style = &style_run.style;
            let range = DWRITE_TEXT_RANGE {
                startPosition: style_run.wch_start as u32,
                length: (style_run.wch_end - style_run.wch_start) as u32,
            };
            // Apply styles to the layout
            if let Some(family_name) = &style.font_family {
                unsafe {
                    layout.SetFontFamilyName(
                        PCWSTR(HSTRING::from(family_name).as_ptr()),
                        range.clone(),
                    )?
                }
            }
            if let Some(font_weight) = &style.font_weight {
                unsafe { layout.SetFontWeight(DWRITE_FONT_WEIGHT(*font_weight), range.clone())? }
            }
            if let Some(font_width) = &style.font_width {
                unsafe { layout.SetFontStretch(DWRITE_FONT_STRETCH(*font_width), range.clone())? }
            }
            if let Some(font_style) = &style.font_style {
                unsafe { layout.SetFontStyle(font_style.clone().into(), range.clone())? }
            }
            if let Some(font_size) = &style.font_size {
                unsafe { layout.SetFontSize(*font_size, range.clone())? }
            }
            if let Some(color) = &style.color {
                let c = csscolorparser::parse(color).unwrap_or_default();
                let brush: ISvgColor = SvgColorImpl::new(c).into();
                let brush: IUnknown = brush.cast()?;
                unsafe { layout.SetDrawingEffect(&brush, range.clone())? }
            }
            if let Some(lang) = &style.lang {
                unsafe {
                    layout.SetLocaleName(PCWSTR(HSTRING::from(lang).as_ptr()), range.clone())?
                }
            }
            if !style.font_feature_settings.is_empty() {
                let typography = unsafe { factory.CreateTypography()? };
                for (feature, parameter) in style.font_feature_settings.iter() {
                    let feature = DWRITE_FONT_FEATURE {
                        nameTag: DWRITE_FONT_FEATURE_TAG(string_to_tag(feature)),
                        parameter: *parameter,
                    };
                    unsafe { typography.AddFontFeature(feature)? }
                }
                unsafe { layout.SetTypography(&typography, range.clone())? }
            }
            if !style.font_variation_settings.is_empty() {
                let mut axis_values: Vec<DWRITE_FONT_AXIS_VALUE> = Vec::new();
                for (axis, value) in style.font_variation_settings.iter() {
                    match &value {
                        FontVariationValue::Set(x) => {
                            axis_values.push(DWRITE_FONT_AXIS_VALUE {
                                axisTag: DWRITE_FONT_AXIS_TAG(string_to_tag(axis)),
                                value: *x,
                            });
                        }
                        _ => {}
                    }
                }
                if let Ok(layout4) = layout.cast::<IDWriteTextLayout4>() {
                    unsafe { layout4.SetFontAxisValues(&axis_values, range.clone())? }
                }
            }
        }
        Ok(layout)
    }

    pub(crate) fn compute_layout_offset(
        canvas_width: f32,
        canvas_height: f32,
        frame: &DocumentFrame,
        metrics: &DWRITE_TEXT_METRICS,
    ) -> (f32, f32) {
        let factor_h = match frame.horizontal_align {
            HAlign::Left => 0.0,
            HAlign::Center => 0.5,
            HAlign::Right => 1.0,
        };
        let factor_v = match frame.vertical_align {
            VAlign::Top => 0.0,
            VAlign::Center => 0.5,
            VAlign::Bottom => 1.0,
        };
        let frame_left = frame.left.unwrap_or(0.0);
        let frame_top = frame.top.unwrap_or(0.0);
        let frame_right = frame.right.unwrap_or(canvas_width);
        let frame_bottom = frame.bottom.unwrap_or(canvas_height);

        let tm_cx = metrics.left + factor_h * metrics.width;
        let tm_cy = metrics.top + factor_v * metrics.height;

        let frame_cx = frame_left + factor_h * (frame_right - frame_left);
        let frame_cy = frame_top + factor_v * (frame_bottom - frame_top);

        (frame_cx - tm_cx, frame_cy - tm_cy)
    }
}

#[derive(Debug)]
struct StyleRun {
    wch_start: usize,
    wch_end: usize,
    style: TextStyle,
}
