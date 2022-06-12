use std::ffi::OsString;

use windows::{
    core::Result,
    Win32::Graphics::DirectWrite::{
        IDWriteFactory, IDWriteTextFormat, IDWriteTextLayout, DWRITE_FONT_STRETCH,
        DWRITE_FONT_WEIGHT, DWRITE_TEXT_RANGE,
    },
};

use crate::document::{DocumentBody, DocumentContent, Style};

#[derive(Debug)]
pub(crate) struct DocumentAnalyzer {
    // Text encoded in UTF-16
    text: Vec<u16>,
    style_runs: Vec<StyleRun>,
    style_stack: Vec<Style>,
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
    fn start_style_run(&mut self, style: Style) {
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
                let mut last_style = self.style_stack.last().cloned().unwrap_or_default();
                last_style.merge(s);
                self.start_style_run(last_style);
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
        db: &DocumentBody,
    ) -> Result<IDWriteTextLayout> {
        let layout = unsafe {
            factory.CreateTextLayout(
                &self.text,
                format,
                db.right.unwrap_or(canvas_width) - db.left.unwrap_or(0.0),
                db.bottom.unwrap_or(canvas_height) - db.bottom.unwrap_or(0.0),
            )?
        };
        for style_run in self.style_runs.iter() {
            if style_run.wch_end <= style_run.wch_start {
                continue;
            }
            let style = &style_run.style;
            let range = DWRITE_TEXT_RANGE {
                startPosition: style_run.wch_start as u32,
                length: (style_run.wch_end - style_run.wch_start) as u32,
            };
            if let Some(family_name) = &style.font_family {
                let family_name = OsString::from(family_name);
                unsafe { layout.SetFontFamilyName(family_name, range.clone())? }
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
        }
        Ok(layout)
    }
}

#[derive(Debug)]
struct StyleRun {
    wch_start: usize,
    wch_end: usize,
    style: Style,
}
