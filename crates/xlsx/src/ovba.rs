//! VBA projects (`vbaProject.bin`): the module source code inside, per Microsoft's public
//! [MS-OVBA] specification. The project is a compound file: a `PROJECT` stream naming the
//! modules and their kinds, and a `VBA` storage with the `dir` stream (compressed records: code
//! page, modules, where each module's compressed source starts) and one stream per module.
//!
//! Only the source is read; the compiled p-code (`_VBA_PROJECT`, the head of each module stream)
//! is version-specific and ignored. GridCraft runs the source itself.

/// What a module is.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ModuleKind {
    /// A standard module (`Module1`).
    Standard,
    /// A class module.
    Class,
    /// The code behind a workbook or worksheet (`ThisWorkbook`, `Sheet1`).
    Document,
    /// The code behind a UserForm.
    Form,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VbaModule {
    /// The module's name in the editor.
    pub name: String,
    /// The stream in the `VBA` storage holding it.
    pub stream: String,
    pub kind: ModuleKind,
    /// The source, `Attribute` lines included, with `\r\n` line ends.
    pub source: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct VbaProject {
    pub name: String,
    pub code_page: u16,
    pub modules: Vec<VbaModule>,
}

/// Largest decompressed text accepted (a module's source, the dir stream).
const MAX_TEXT: usize = 64 << 20;

// ---------------------------------------------------------------- compression ([MS-OVBA] 2.4.1)

/// Decompresses a CompressedContainer.
pub fn decompress(data: &[u8]) -> Result<Vec<u8>, String> {
    let bad = |m: &str| format!("damaged VBA project: {m}");
    if data.first() != Some(&1) {
        return Err(bad("compressed data has no signature"));
    }
    let mut out: Vec<u8> = Vec::with_capacity(data.len() * 2);
    let mut pos = 1usize;
    while pos < data.len() {
        let Some(h) = data.get(pos..pos + 2).map(|b| u16::from_le_bytes([b[0], b[1]])) else { break };
        pos += 2;
        let size = usize::from(h & 0x0FFF) + 3;
        let compressed = h & 0x8000 != 0;
        let end = (pos + size - 2).min(data.len());
        let chunk_start = out.len();
        if !compressed {
            let raw = data.get(pos..(pos + 4096).min(data.len())).ok_or_else(|| bad("raw chunk out of range"))?;
            out.extend_from_slice(raw);
            pos += raw.len();
            continue;
        }
        while pos < end {
            let flags = data[pos];
            pos += 1;
            for bit in 0..8 {
                if pos >= end {
                    break;
                }
                if flags & (1 << bit) == 0 {
                    out.push(data[pos]);
                    pos += 1;
                } else {
                    let Some(t) = data.get(pos..pos + 2).map(|b| u16::from_le_bytes([b[0], b[1]])) else {
                        return Err(bad("truncated copy token"));
                    };
                    pos += 2;
                    let (offset, length) = copy_token(t, out.len() - chunk_start);
                    if offset > out.len() - chunk_start || offset == 0 {
                        return Err(bad("copy token reaches before its chunk"));
                    }
                    let from = out.len() - offset;
                    for k in 0..length {
                        let b = out[from + k];
                        out.push(b);
                    }
                }
                if out.len() > MAX_TEXT {
                    return Err(bad("decompressed data is too large"));
                }
            }
        }
        pos = end;
    }
    Ok(out)
}

/// The bit split of a copy token given how much of the chunk is decompressed so far.
fn bit_count(done: usize) -> u32 {
    let mut bits = 4;
    while (1usize << bits) < done {
        bits += 1;
    }
    bits.min(12)
}

fn copy_token(t: u16, done: usize) -> (usize, usize) {
    let bits = bit_count(done);
    let length_mask = 0xFFFFu16 >> bits;
    let length = usize::from(t & length_mask) + 3;
    let offset = usize::from((t & !length_mask) >> (16 - bits)) + 1;
    (offset, length)
}

/// Compresses `data` into a CompressedContainer (greedy matching within each 4 KB chunk).
pub fn compress(data: &[u8]) -> Vec<u8> {
    let mut out = vec![1u8];
    for chunk in data.chunks(4096) {
        let mut body = vec![];
        let mut i = 0usize;
        while i < chunk.len() {
            let flag_at = body.len();
            body.push(0u8);
            for bit in 0..8 {
                if i >= chunk.len() {
                    break;
                }
                let bits = bit_count(i);
                let max_len = (0xFFFFusize >> bits) + 3;
                let max_off = 1usize << (16 - bits);
                let (mut best_len, mut best_off) = (0usize, 0usize);
                let lo = i.saturating_sub(max_off);
                for start in (lo..i).rev() {
                    let mut l = 0;
                    while l < max_len && i + l < chunk.len() && chunk[start + l] == chunk[i + l] {
                        l += 1;
                    }
                    if l > best_len {
                        best_len = l;
                        best_off = i - start;
                        if l == max_len {
                            break;
                        }
                    }
                }
                if best_len >= 3 {
                    let token = (((best_off - 1) as u16) << (16 - bits)) | ((best_len - 3) as u16);
                    body.extend_from_slice(&token.to_le_bytes());
                    body[flag_at] |= 1 << bit;
                    i += best_len;
                } else {
                    body.push(chunk[i]);
                    i += 1;
                }
            }
        }
        let header = 0xB000u16 | ((body.len() + 2 - 3) as u16 & 0x0FFF);
        out.extend_from_slice(&header.to_le_bytes());
        out.extend_from_slice(&body);
    }
    out
}

// ---------------------------------------------------------------- text

/// Bytes in the project's code page as text. Windows-1252 and UTF-8 exactly; other code pages
/// read their ASCII part and keep other bytes as Latin-1.
pub fn decode(bytes: &[u8], code_page: u16) -> String {
    const CP1252: [char; 32] = [
        '€', '\u{81}', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u{8d}', 'Ž', '\u{8f}', '\u{90}', '‘', '’', '“', '”', '•', '–', '—',
        '˜', '™', 'š', '›', 'œ', '\u{9d}', 'ž', 'Ÿ',
    ];
    if code_page == 65001 {
        return String::from_utf8_lossy(bytes).into_owned();
    }
    bytes
        .iter()
        .map(|&b| match b {
            0x80..=0x9F if code_page == 1252 => CP1252[usize::from(b - 0x80)],
            b => char::from(b),
        })
        .collect()
}

/// Text as bytes in the project's code page (characters it can't hold become `?`).
pub fn encode(text: &str, code_page: u16) -> Vec<u8> {
    if code_page == 65001 {
        return text.as_bytes().to_vec();
    }
    text.chars()
        .map(|c| {
            let u = c as u32;
            if u < 0x80 || (0xA0..=0xFF).contains(&u) {
                return u as u8;
            }
            if code_page == 1252
                && let Some(i) = (0x80u8..=0x9F).position(|b| decode(&[b], 1252).starts_with(c))
            {
                return 0x80 + i as u8;
            }
            b'?'
        })
        .collect()
}

// ---------------------------------------------------------------- reading

/// The records of a dir stream: (id, data).
fn records(dir: &[u8]) -> Vec<(u16, &[u8])> {
    let mut out = vec![];
    let mut pos = 0usize;
    while let Some(h) = dir.get(pos..pos + 6) {
        let id = u16::from_le_bytes([h[0], h[1]]);
        let mut size = u32::from_le_bytes([h[2], h[3], h[4], h[5]]) as usize;
        // PROJECTVERSION declares 4 bytes but holds 6.
        if id == 0x0009 {
            size = 6;
        }
        let Some(data) = dir.get(pos + 6..pos + 6 + size) else { break };
        out.push((id, data));
        pos += 6 + size;
        if id == 0x0010 {
            break;
        }
    }
    out
}

/// The modules of a VBA project, with their source.
pub fn read_project(bin: &[u8]) -> Result<VbaProject, String> {
    let cf = crate::cfb::Compound::open(bin)?;
    let dir = cf.stream(&["VBA", "dir"])?.ok_or("damaged VBA project: no dir stream")?;
    let dir = decompress(&dir)?;
    let mut p = VbaProject { code_page: 1252, ..Default::default() };
    struct Pending {
        name: String,
        unicode_name: Option<String>,
        stream: String,
        unicode_stream: Option<String>,
        offset: usize,
        procedural: bool,
    }
    let mut cur: Option<Pending> = None;
    let mut done: Vec<Pending> = vec![];
    let utf16 = |d: &[u8]| String::from_utf16_lossy(&d.as_chunks::<2>().0.iter().map(|c| u16::from_le_bytes(*c)).collect::<Vec<_>>());
    for (id, d) in records(&dir) {
        match id {
            0x0003 if d.len() >= 2 => p.code_page = u16::from_le_bytes([d[0], d[1]]),
            0x0004 => p.name = decode(d, p.code_page),
            0x0019 => {
                if let Some(m) = cur.take() {
                    done.push(m);
                }
                cur = Some(Pending {
                    name: decode(d, p.code_page),
                    unicode_name: None,
                    stream: String::new(),
                    unicode_stream: None,
                    offset: 0,
                    procedural: true,
                });
            }
            0x0047 => {
                if let Some(m) = cur.as_mut() {
                    m.unicode_name = Some(utf16(d));
                }
            }
            0x001A => {
                if let Some(m) = cur.as_mut() {
                    m.stream = decode(d, p.code_page);
                }
            }
            0x0032 => {
                if let Some(m) = cur.as_mut() {
                    m.unicode_stream = Some(utf16(d));
                }
            }
            0x0031 if d.len() >= 4 => {
                if let Some(m) = cur.as_mut() {
                    m.offset = u32::from_le_bytes([d[0], d[1], d[2], d[3]]) as usize;
                }
            }
            0x0022 => {
                if let Some(m) = cur.as_mut() {
                    m.procedural = false;
                }
            }
            0x002B => {
                if let Some(m) = cur.take() {
                    done.push(m);
                }
            }
            _ => {}
        }
    }
    if let Some(m) = cur.take() {
        done.push(m);
    }
    // The PROJECT stream tells a document module from a class and a form.
    let project_text = cf.stream(&["PROJECT"])?.map(|b| decode(&b, p.code_page)).unwrap_or_default();
    let kind_of = |name: &str, procedural: bool| {
        if procedural {
            return ModuleKind::Standard;
        }
        for line in project_text.lines() {
            let (k, v) = line.split_once('=').unwrap_or(("", ""));
            let v = v.split('/').next().unwrap_or("").trim_matches('"');
            if v.eq_ignore_ascii_case(name) {
                match k.trim() {
                    "Document" => return ModuleKind::Document,
                    "BaseClass" => return ModuleKind::Form,
                    "Class" => return ModuleKind::Class,
                    _ => {}
                }
            }
        }
        ModuleKind::Class
    };
    for m in done {
        let name = m.unicode_name.unwrap_or(m.name);
        let stream = m.unicode_stream.unwrap_or(m.stream);
        let Some(raw) = cf.stream(&["VBA", &stream])? else { continue };
        let source = match raw.get(m.offset..) {
            Some(c) if !c.is_empty() => decode(&decompress(c)?, p.code_page),
            _ => String::new(),
        };
        let kind = kind_of(&name, m.procedural);
        p.modules.push(VbaModule { name, stream, kind, source });
    }
    Ok(p)
}

/// Builders for tests (other crates' too).
#[cfg(any(test, feature = "test-support"))]
pub mod tests_support {
    use super::{compress, encode};

    fn rec(id: u16, data: &[u8]) -> Vec<u8> {
        let mut v = id.to_le_bytes().to_vec();
        v.extend((data.len() as u32).to_le_bytes());
        v.extend_from_slice(data);
        v
    }

    /// A small VBA project as Excel lays it out: one standard module and one document module.
    pub fn sample_project(modules: &[(&str, bool, &str)]) -> Vec<u8> {
        let mut dir = vec![];
        dir.extend(rec(0x0001, &1u32.to_le_bytes()));
        dir.extend(rec(0x0003, &1252u16.to_le_bytes()));
        dir.extend(rec(0x0004, b"VBAProject"));
        // PROJECTVERSION: declared size 4, six bytes of data.
        dir.extend(0x0009u16.to_le_bytes());
        dir.extend(4u32.to_le_bytes());
        dir.extend([1, 0, 0, 0, 2, 0]);
        dir.extend(rec(0x000F, &(modules.len() as u16).to_le_bytes()));
        dir.extend(rec(0x0013, &0xFFFFu16.to_le_bytes()));
        let mut streams: Vec<(String, Vec<u8>)> = vec![];
        let mut project = String::from("ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n");
        for (name, standard, source) in modules {
            dir.extend(rec(0x0019, name.as_bytes()));
            dir.extend(rec(0x001A, name.as_bytes()));
            let unicode: Vec<u8> = name.encode_utf16().flat_map(u16::to_le_bytes).collect();
            dir.extend(rec(0x0032, &unicode));
            // A p-code head the reader must skip.
            let head = vec![0xAAu8; 37];
            dir.extend(rec(0x0031, &(head.len() as u32).to_le_bytes()));
            dir.extend(rec(if *standard { 0x0021 } else { 0x0022 }, &[]));
            dir.extend(rec(0x002B, &[]));
            let mut stream = head;
            stream.extend(compress(&encode(source, 1252)));
            streams.push((name.to_string(), stream));
            project.push_str(&if *standard { format!("Module={name}\r\n") } else { format!("Document={name}/&H00000000\r\n") });
        }
        dir.extend(rec(0x0010, &[]));
        let dir = compress(&dir);
        let mut list = vec![
            crate::cfb_write::Stream { path: vec!["PROJECT"], data: project.as_bytes() },
            crate::cfb_write::Stream { path: vec!["VBA", "dir"], data: &dir },
        ];
        for (n, b) in &streams {
            list.push(crate::cfb_write::Stream { path: vec!["VBA", n.as_str()], data: b });
        }
        crate::cfb_write::write(&list).unwrap_or_default()
    }
}

#[cfg(test)]
mod tests {
    use super::tests_support::sample_project;
    use super::*;

    #[test]
    fn compression_round_trips() {
        for text in [
            "".to_string(),
            "a".to_string(),
            "abcabcabcabcabcabc".to_string(),
            "Attribute VB_Name = \"Module1\"\r\nSub Hola()\r\n    MsgBox \"Hola\"\r\nEnd Sub\r\n".repeat(200),
            (0..20_000u32).map(|i| char::from(b'a' + (i * 7 % 26) as u8)).collect(),
        ] {
            let c = compress(text.as_bytes());
            assert_eq!(decompress(&c).unwrap(), text.as_bytes(), "len {}", text.len());
        }
        assert!(decompress(&[]).is_err());
        assert!(decompress(&[1, 0xFF, 0xBF, 0x01, 0x05, 0x00]).is_err(), "a copy token before the chunk");
    }

    /// The worked example of [MS-OVBA] 3.2: "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa".
    #[test]
    fn spec_example() {
        let compressed = [
            0x01, 0x2F, 0xB0, 0x00, 0x23, 0x61, 0x61, 0x61, 0x62, 0x63, 0x64, 0x65, 0x82, 0x66, 0x00, 0x70, 0x61, 0x67, 0x68, 0x69, 0x6A, 0x01, 0x38,
            0x08, 0x61, 0x6B, 0x6C, 0x00, 0x30, 0x6D, 0x6E, 0x6F, 0x70, 0x06, 0x71, 0x02, 0x70, 0x04, 0x10, 0x72, 0x73, 0x74, 0x75, 0x76, 0x10, 0x77,
            0x78, 0x79, 0x7A, 0x00, 0x3C,
        ];
        assert_eq!(String::from_utf8(decompress(&compressed).unwrap()).unwrap(), "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa");
    }

    #[test]
    fn reads_modules() {
        let bin = sample_project(&[
            ("Module1", true, "Attribute VB_Name = \"Module1\"\r\nSub Saludo()\r\n    MsgBox \"Año €\"\r\nEnd Sub\r\n"),
            ("ThisWorkbook", false, "Attribute VB_Name = \"ThisWorkbook\"\r\n"),
        ]);
        let p = read_project(&bin).unwrap();
        assert_eq!(p.code_page, 1252);
        assert_eq!(p.modules.len(), 2);
        assert_eq!(p.modules[0].kind, ModuleKind::Standard);
        assert!(p.modules[0].source.contains("MsgBox \"Año €\""));
        assert_eq!(p.modules[1].kind, ModuleKind::Document);
        assert!(read_project(b"not a project").is_err());
    }

    #[test]
    fn code_pages() {
        assert_eq!(decode(&[0x80, 0xF1], 1252), "€ñ");
        assert_eq!(encode("€ñ✓", 1252), vec![0x80, 0xF1, b'?']);
        assert_eq!(decode("ñ".as_bytes(), 65001), "ñ");
    }
}
