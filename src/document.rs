use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use windows::Win32::Graphics::DirectWrite::*;

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Document {
    #[serde(default = "default_width")]
    pub(crate) width: f32,
    #[serde(default = "default_height")]
    pub(crate) height: f32,
    #[serde(default)]
    pub(crate) font_files: Vec<String>,
    #[serde(default)]
    pub(crate) frames: Vec<DocumentFrame>,
}

const fn default_width() -> f32 {
    1024.0
}
const fn default_height() -> f32 {
    1024.0
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) struct DocumentFrame {
    pub(crate) left: Option<f32>,
    pub(crate) top: Option<f32>,
    pub(crate) right: Option<f32>,
    pub(crate) bottom: Option<f32>,

    // Accessibility
    pub(crate) title: Option<String>,
    pub(crate) desc: Option<String>,

    #[serde(default)]
    pub(crate) text_align: TextAlign,
    #[serde(default)]
    pub(crate) writing_mode: WritingMode,
    #[serde(default)]
    pub(crate) horizontal_align: HAlign,
    #[serde(default)]
    pub(crate) vertical_align: VAlign,
    #[serde(default = "default_line_height")]
    pub(crate) line_height: f32,
    #[serde(default = "default_baseline")]
    pub(crate) baseline_offset: f32,

    pub(crate) contents: DocumentContent,
}

const fn default_line_height() -> f32 {
    1.50
}
const fn default_baseline() -> f32 {
    0.80
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub(crate) enum DocumentContent {
    Text(String),
    Style(TextStyle),
    Embed(Box<Vec<DocumentContent>>),
}

// Block-level styles

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum WritingMode {
    LrTb,
    LrBt,
    RlTb,
    RlBt,
    TbLr,
    TbRl,
    BtLr,
    BtRl,
}

impl Default for WritingMode {
    fn default() -> Self {
        Self::LrTb
    }
}

impl Into<(DWRITE_READING_DIRECTION, DWRITE_FLOW_DIRECTION)> for WritingMode {
    fn into(self) -> (DWRITE_READING_DIRECTION, DWRITE_FLOW_DIRECTION) {
        match self {
            Self::LrTb => (
                DWRITE_READING_DIRECTION_LEFT_TO_RIGHT,
                DWRITE_FLOW_DIRECTION_TOP_TO_BOTTOM,
            ),
            Self::LrBt => (
                DWRITE_READING_DIRECTION_LEFT_TO_RIGHT,
                DWRITE_FLOW_DIRECTION_BOTTOM_TO_TOP,
            ),
            Self::RlTb => (
                DWRITE_READING_DIRECTION_RIGHT_TO_LEFT,
                DWRITE_FLOW_DIRECTION_TOP_TO_BOTTOM,
            ),
            Self::RlBt => (
                DWRITE_READING_DIRECTION_RIGHT_TO_LEFT,
                DWRITE_FLOW_DIRECTION_BOTTOM_TO_TOP,
            ),
            Self::TbLr => (
                DWRITE_READING_DIRECTION_TOP_TO_BOTTOM,
                DWRITE_FLOW_DIRECTION_LEFT_TO_RIGHT,
            ),
            Self::TbRl => (
                DWRITE_READING_DIRECTION_TOP_TO_BOTTOM,
                DWRITE_FLOW_DIRECTION_RIGHT_TO_LEFT,
            ),
            Self::BtLr => (
                DWRITE_READING_DIRECTION_BOTTOM_TO_TOP,
                DWRITE_FLOW_DIRECTION_LEFT_TO_RIGHT,
            ),
            Self::BtRl => (
                DWRITE_READING_DIRECTION_BOTTOM_TO_TOP,
                DWRITE_FLOW_DIRECTION_RIGHT_TO_LEFT,
            ),
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum TextAlign {
    #[serde(alias = "leading")]
    Left,
    Center,
    #[serde(alias = "trailing")]
    Right,
    Justify,
}

impl Default for TextAlign {
    fn default() -> Self {
        Self::Left
    }
}

impl Into<DWRITE_TEXT_ALIGNMENT> for TextAlign {
    fn into(self) -> DWRITE_TEXT_ALIGNMENT {
        match self {
            TextAlign::Left => DWRITE_TEXT_ALIGNMENT_LEADING,
            TextAlign::Right => DWRITE_TEXT_ALIGNMENT_TRAILING,
            TextAlign::Center => DWRITE_TEXT_ALIGNMENT_CENTER,
            TextAlign::Justify => DWRITE_TEXT_ALIGNMENT_JUSTIFIED,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum HAlign {
    Left,
    Center,
    Right,
}

impl Default for HAlign {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum VAlign {
    Top,
    Center,
    Bottom,
}

impl Default for VAlign {
    fn default() -> Self {
        Self::Top
    }
}

// Run-level styles

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) struct TextStyle {
    #[serde(default)]
    pub(crate) font_family: Option<String>,
    #[serde(default)]
    pub(crate) font_weight: Option<i32>,
    #[serde(default)]
    pub(crate) font_width: Option<i32>,
    #[serde(default)]
    pub(crate) font_style: Option<FontStyle>,
    #[serde(default)]
    pub(crate) font_size: Option<f32>,
    #[serde(default)]
    pub(crate) color: Option<String>,
    #[serde(default)]
    pub(crate) lang: Option<String>,
    #[serde(default)]
    pub(crate) font_feature_settings: BTreeMap<String, u32>,
    #[serde(default)]
    pub(crate) font_variation_settings: BTreeMap<String, FontVariationValue>,
}

impl TextStyle {
    pub(crate) fn merge(&mut self, other: &TextStyle) {
        self.font_family = other.font_family.clone().or(self.font_family.clone());
        self.font_weight = other.font_weight.clone().or(self.font_weight.clone());
        self.font_width = other.font_width.clone().or(self.font_width.clone());
        self.font_style = other.font_style.clone().or(self.font_style.clone());
        self.font_size = other.font_size.clone().or(self.font_size.clone());
        self.color = other.color.clone().or(self.color.clone());
        self.lang = other.lang.clone().or(self.lang.clone());
        self.font_feature_settings
            .extend(other.font_feature_settings.clone());
        self.font_variation_settings
            .extend(other.font_variation_settings.clone());
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) enum FontStyle {
    Normal,
    Oblique,
    Italic,
}

impl Into<DWRITE_FONT_STYLE> for FontStyle {
    fn into(self) -> DWRITE_FONT_STYLE {
        match self {
            FontStyle::Normal => DWRITE_FONT_STYLE_NORMAL,
            FontStyle::Oblique => DWRITE_FONT_STYLE_OBLIQUE,
            FontStyle::Italic => DWRITE_FONT_STYLE_ITALIC,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub(crate) enum FontVariationValue {
    Default,
    Set(f32),
}

/// Convert a string to DW tag. Note that DW uses little endian.
pub(crate) fn string_to_tag(tag_str: &str) -> u32 {
    let mut len: usize = 0;
    let mut result: u32 = 0;
    for ch in tag_str.chars() {
        let code = ch as u32;
        result = (result >> 8) | ((code & 0xff) << 24);
        len += 1;
    }
    while len < 4 {
        result = (result >> 8) | 0x20000000;
        len += 1;
    }
    result
}
