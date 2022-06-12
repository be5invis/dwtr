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
#[serde(rename_all = "camelCase")]
pub(crate) struct Style {
    #[serde(default)]
    pub(crate) font_family: Option<String>,
    #[serde(default)]
    pub(crate) font_weight: Option<i32>,
    #[serde(default)]
    pub(crate) font_width: Option<i32>,
    #[serde(default)]
    pub(crate) font_style: Option<FontStyle>,
}

impl Style {
    pub(crate) fn merge(&mut self, other: &Style) {
        self.font_family = other.font_family.clone().or(self.font_family.clone());
        self.font_weight = other.font_weight.clone().or(self.font_weight.clone());
        self.font_width = other.font_width.clone().or(self.font_width.clone());
        self.font_style = other.font_style.clone().or(self.font_style.clone());
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
pub(crate) enum FontStyle {
    Upright,
    Italic,
    Oblique,
}

impl Into<DWRITE_FONT_STYLE> for FontStyle {
    fn into(self) -> DWRITE_FONT_STYLE {
        match self {
            FontStyle::Upright => DWRITE_FONT_STYLE_NORMAL,
            FontStyle::Italic => DWRITE_FONT_STYLE_ITALIC,
            FontStyle::Oblique => DWRITE_FONT_STYLE_OBLIQUE,
        }
    }
}
