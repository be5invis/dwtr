use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use windows::Win32::Graphics::DirectWrite::{
    DWRITE_FONT_STYLE, DWRITE_FONT_STYLE_ITALIC, DWRITE_FONT_STYLE_NORMAL,
    DWRITE_FONT_STYLE_OBLIQUE,
};

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
    pub(crate) body: Vec<DocumentBody>,
}

const fn default_width() -> f32 {
    1024.0
}
const fn default_height() -> f32 {
    1024.0
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct DocumentBody {
    pub(crate) left: Option<f32>,
    pub(crate) top: Option<f32>,
    pub(crate) right: Option<f32>,
    pub(crate) bottom: Option<f32>,
    pub(crate) contents: DocumentContent,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub(crate) enum DocumentContent {
    Text(String),
    Style(Style),
    Embed(Box<Vec<DocumentContent>>),
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub(crate) struct Style {
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
    pub(crate) font_feature_settings: BTreeMap<String, u32>,
    #[serde(default)]
    pub(crate) font_variation_settings: BTreeMap<String, FontVariationValue>,
}

impl Style {
    pub(crate) fn merge(&mut self, other: &Style) {
        self.font_family = other.font_family.clone().or(self.font_family.clone());
        self.font_weight = other.font_weight.clone().or(self.font_weight.clone());
        self.font_width = other.font_width.clone().or(self.font_width.clone());
        self.font_style = other.font_style.clone().or(self.font_style.clone());
        self.font_size = other.font_size.clone().or(self.font_size.clone());
        self.color = other.color.clone().or(self.color.clone());
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
