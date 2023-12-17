use clap::Parser;
use document::Document;
use std::fs::File;
use std::io::BufReader;
use std::path::PathBuf;
use svg_text_render::SvgDocumentRenderer;
use windows::{core::ComInterface, w, Win32::Graphics::DirectWrite::*};

use crate::{
    document_analyzer::DocumentAnalyzer, error::Result, font_loader::load_font_collection,
};

mod document;
mod document_analyzer;
mod error;
mod escape;
mod font_loader;
mod svg_color;
mod svg_text_render;

#[derive(Debug, clap::StructOpt)]
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
    let file = File::open(opt.input.as_path())?;
    let reader = BufReader::new(file);
    let document: Document = serde_json::from_reader(reader)?;

    let factory = get_factory()?;
    let font_collection = load_font_collection(factory.cast()?, &document)?;

    let format = unsafe {
        factory.CreateTextFormat(
            w!("Calibri"),
            &font_collection,
            DWRITE_FONT_WEIGHT(400),
            DWRITE_FONT_STYLE_NORMAL,
            DWRITE_FONT_STRETCH_NORMAL,
            24.0,
            w!("en-us"),
        )?
    };

    let mut document_renderer = SvgDocumentRenderer::new(document.width, document.height);

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

        let mut metrics = DWRITE_TEXT_METRICS::default();
        unsafe { text_layout.GetMetrics(&mut metrics)? };
        let (offset_x, offset_y) = DocumentAnalyzer::compute_layout_offset(
            document.width,
            document.height,
            frame,
            &metrics,
        );

        let frame_renderer = document_renderer.create_frame_renderer(offset_x, offset_y);
        frame_renderer.set_title(frame.title.clone());
        frame_renderer.set_desc(frame.desc.clone());

        let fr1: IDWriteTextRenderer1 = frame_renderer.into();
        unsafe { text_layout.Draw(None, &fr1, 0.0, 0.0)? }
    }

    let mut out_stream: Box<dyn std::io::Write> = match opt.output {
        Some(output) => Box::new(std::fs::File::create(output.as_path()).unwrap()),
        None => Box::new(std::io::stdout()),
    };

    write!(
        out_stream,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n"
    )?;

    let svg = document_renderer.into_xml();
    svg::write(out_stream, &svg)?;

    Ok(())
}

fn get_factory() -> Result<IDWriteFactory> {
    unsafe {
        let factory_raw = DWriteCreateFactory::<IDWriteFactory7>(DWRITE_FACTORY_TYPE_SHARED)?;
        let factory: IDWriteFactory = factory_raw.cast()?;
        Ok(factory)
    }
}
