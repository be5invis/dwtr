use std::ffi::OsString;

use crate::document::Document;
use glob::glob;
use windows::core::{ComInterface, Result, HSTRING};
use windows::Win32::Foundation::BOOL;
use windows::Win32::Graphics::DirectWrite::*;

pub(crate) fn load_font_collection(
    factory: IDWriteFactory,
    document: &Document,
) -> Result<IDWriteFontCollection1> {
    let factory3: IDWriteFactory3 = factory.cast()?;
    unsafe {
        let fsb = factory3.CreateFontSetBuilder()?;

        for pattern in document.font_files.iter() {
            for entry in glob(pattern).unwrap() {
                if let Ok(font_path) = entry {
                    let font_path = HSTRING::from(OsString::from(font_path));
                    let font_file = factory3.CreateFontFileReference(&font_path, None)?;

                    // Analyzer font file, get face count
                    let mut is_supported: BOOL = BOOL::from(false);
                    let mut font_file_type = DWRITE_FONT_FILE_TYPE::default();
                    let mut font_face_type = DWRITE_FONT_FACE_TYPE::default();
                    let mut num_of_faces: u32 = 0;
                    font_file.Analyze(
                        &mut is_supported,
                        &mut font_file_type,
                        Some(&mut font_face_type),
                        &mut num_of_faces,
                    )?;

                    if is_supported.as_bool() {
                        for i in 0..num_of_faces {
                            let font_face_ref = factory3.CreateFontFaceReference(
                                &font_file,
                                i,
                                DWRITE_FONT_SIMULATIONS_NONE,
                            )?;
                            fsb.AddFontFaceReference2(&font_face_ref)?;
                        }
                    }
                }
            }
        }

        let fs = fsb.CreateFontSet()?;
        factory3.CreateFontCollectionFromFontSet(&fs)
    }
}
