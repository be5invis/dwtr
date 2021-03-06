use core::ptr::null;
use document::Document;
use std::fs::File;
use std::io::BufReader;
use std::path::PathBuf;
use structopt::StructOpt;
use svg_text_render::SvgTextRenderer;
use windows::{
    core::{AsImpl, Interface, Result},
    Win32::Graphics::DirectWrite::*,
};

use crate::{document_analyzer::DocumentAnalyzer, font_loader::load_font_collection};

mod document;
mod document_analyzer;
mod font_loader;
mod svg_color;
mod svg_text_render;

#[derive(Debug, StructOpt)]
#[structopt(name = "dwtr", about = "Text rendering utility (DWrite)")]
struct Opt {
    /// Input file
    #[structopt(parse(from_os_str))]
    input: PathBuf,

    /// Output file, stdout if not present
    #[structopt(short, long, parse(from_os_str))]
    output: Option<PathBuf>,
}

fn main() -> Result<()> {
    let opt = Opt::from_args();
    let file = File::open(opt.input.as_path()).unwrap();
    let reader = BufReader::new(file);
    let document: Document = serde_json::from_reader(reader).unwrap();

    let factory = get_factory()?;
    let font_collection = load_font_collection(factory.cast()?, &document)?;

    let format = unsafe {
        factory.CreateTextFormat(
            "Calibri",
            font_collection,
            DWRITE_FONT_WEIGHT(400),
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            24.0,
            "en-us",
        )?
    };

    let renderer: IDWriteTextRenderer1 =
        SvgTextRenderer::new(document.width, document.height).into();

    for frame in document.frames.iter() {
        let mut analyzer = DocumentAnalyzer::new();
        analyzer.analyze(&frame.contents);

        let text_layout = analyzer.create_text_layout(
            factory.clone(),
            format.clone(),
            document.width,
            document.height,
            frame,
        )?;
        {
            let metrics = unsafe { text_layout.GetMetrics()? };
            let (offset_x, offset_y) = DocumentAnalyzer::compute_layout_offset(
                document.width,
                document.height,
                frame,
                &metrics,
            );
            renderer.as_impl().set_offset(offset_x, offset_y);
        }
        unsafe { text_layout.Draw(null(), renderer.clone(), 0.0, 0.0)? }
    }

    let mut out_stream: Box<dyn std::io::Write> = match opt.output {
        Some(output) => Box::new(std::fs::File::create(output.as_path()).unwrap()),
        None => Box::new(std::io::stdout()),
    };
    renderer
        .as_impl()
        .into_xml()
        .write_to(&mut out_stream)
        .unwrap();

    Ok(())
}

fn get_factory() -> Result<IDWriteFactory> {
    unsafe {
        let factory_raw = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, &IDWriteFactory7::IID)?;
        let factory: IDWriteFactory = factory_raw.cast()?;
        Ok(factory)
    }
}
