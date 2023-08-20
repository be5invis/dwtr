enum Segment {
    Literal(char),
    Escape(&'static str),
}

impl Segment {
    fn from_char(c: char) -> Self {
        match c {
            '<' => Self::Escape("&lt;"),
            '>' => Self::Escape("&gt;"),
            '\'' => Self::Escape("&apos;"),
            '"' => Self::Escape("&quot;"),
            '&' => Self::Escape("&amp;"),
            '\n' => Self::Escape("&#xA;"),
            '\r' => Self::Escape("&#xD;"),
            _ => Self::Literal(c),
        }
    }
}

pub fn escape_str(s: &str) -> String {
    let mut output: String = String::new();
    for ch in s.chars() {
        let seg = Segment::from_char(ch);
        match seg {
            Segment::Literal(ch) => output.push(ch),
            Segment::Escape(esc) => output.push_str(esc),
        }
    }
    output
}
