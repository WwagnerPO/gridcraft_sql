//! Form-control buttons (Developer › Insert › Button) of a sheet, read from what the file keeps of
//! them: the legacy VML drawing (position, caption, assigned macro) and the sheet's `<controls>`
//! element (the name `Application.Caller` reports, e.g. "Button 1"). The controls stay
//! pass-through content; this reads them, and rewrites a caption in place when a macro changes it.

use gridcraft_model::Preserved;

/// A button and the macro it runs.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormButton {
    /// The VML shape number (`_x0000_s1025` → 1025), the control's `shapeId`.
    pub shape_id: u32,
    /// "Button 1": the name `Application.Caller` and `Shapes("…")` use.
    pub name: String,
    pub caption: String,
    /// The assigned macro without the workbook prefix (`[0]!Main` → `Main`); empty for none.
    pub macro_name: String,
    /// `x:Anchor`: left column, left offset (px), top row, top offset, right column, right
    /// offset, bottom row, bottom offset. Zero-based cells.
    pub anchor: [u32; 8],
}

/// The buttons of a sheet's kept content.
pub fn buttons(p: &Preserved) -> Vec<FormButton> {
    let Some(vml) = p.vml.as_ref() else { return vec![] };
    let text = String::from_utf8_lossy(&vml.data);
    let names = control_names(p);
    let mut out = vec![];
    for shape in shapes(&text) {
        if !shape.contains("ObjectType=\"Button\"") {
            continue;
        }
        let Some(id) = attr(shape, "id").and_then(|i| i.rsplit('s').next().and_then(|n| n.parse::<u32>().ok())) else { continue };
        let anchor: Vec<u32> = tag(shape, "x:Anchor").map(|a| a.split(',').filter_map(|n| n.trim().parse().ok()).collect()).unwrap_or_default();
        let Ok(anchor) = <[u32; 8]>::try_from(anchor) else { continue };
        let macro_name = tag(shape, "x:FmlaMacro").map(|m| strip_book(&unescape(m.trim()))).unwrap_or_default();
        let caption = tag(shape, "v:textbox").map(caption_text).unwrap_or_default();
        let name = names.iter().find(|(s, _)| *s == id).map(|(_, n)| n.clone()).unwrap_or_else(|| format!("Button {}", out.len() + 1));
        out.push(FormButton { shape_id: id, name, caption, macro_name, anchor });
    }
    out
}

/// Replaces a button's caption in the kept VML. False when there is no such button.
pub fn set_caption(p: &mut Preserved, shape_id: u32, caption: &str) -> bool {
    let Some(vml) = p.vml.as_mut() else { return false };
    let text = String::from_utf8_lossy(&vml.data).into_owned();
    let needle = format!("_x0000_s{shape_id}\"");
    let Some(at) = text.find(&needle) else { return false };
    let shape_end = text[at..].find("</v:shape>").map(|i| at + i).unwrap_or(text.len());
    let Some(open) = text[at..shape_end].find("<v:textbox").map(|i| at + i) else { return false };
    let Some(body) = text[open..].find('>').map(|i| open + i + 1) else { return false };
    let Some(close) = text[body..].find("</v:textbox>").map(|i| body + i) else { return false };
    let lines: Vec<String> = caption.split('\n').map(|l| escape(l.trim_end_matches('\r'))).collect();
    let inner = format!("<div style='text-align:center'><font face=\"Calibri\" size=\"220\" color=\"#000000\">{}</font></div>", lines.join("<br>"));
    let mut out = String::with_capacity(text.len() + inner.len());
    out.push_str(&text[..body]);
    out.push_str(&inner);
    out.push_str(&text[close..]);
    vml.data = out.into_bytes().into();
    true
}

/// `<v:shape …>…</v:shape>` blocks.
fn shapes(text: &str) -> Vec<&str> {
    let mut out = vec![];
    let mut rest = text;
    while let Some(start) = rest.find("<v:shape ") {
        let Some(len) = rest[start..].find("</v:shape>") else { break };
        out.push(&rest[start..start + len]);
        rest = &rest[start + len + 10..];
    }
    out
}

fn attr<'a>(el: &'a str, name: &str) -> Option<&'a str> {
    let key = format!(" {name}=\"");
    let i = el.find(&key)? + key.len();
    let j = el[i..].find('"')? + i;
    Some(&el[i..j])
}

fn tag<'a>(el: &'a str, name: &str) -> Option<&'a str> {
    let open = el.find(&format!("<{name}"))?;
    let body = el[open..].find('>')? + open + 1;
    let close = el[body..].find(&format!("</{name}>"))? + body;
    Some(&el[body..close])
}

/// The caption: the textbox's text, line breaks kept.
fn caption_text(html: &str) -> String {
    let with_breaks = html.replace("<br>", "\n").replace("<br/>", "\n").replace("</div><div", "</div>\n<div");
    let mut out = String::new();
    let mut in_tag = false;
    for c in with_breaks.chars() {
        match c {
            '<' => in_tag = true,
            '>' => in_tag = false,
            c if !in_tag => out.push(c),
            _ => {}
        }
    }
    unescape(out.trim()).lines().map(str::trim).collect::<Vec<_>>().join("\n")
}

/// `[0]!Main`, `'Book.xlsm'!Main` → `Main`.
fn strip_book(m: &str) -> String {
    m.rsplit_once('!').map(|(_, n)| n).unwrap_or(m).to_string()
}

/// (shapeId, name) of the sheet's `<control>` elements.
fn control_names(p: &Preserved) -> Vec<(u32, String)> {
    let mut out = vec![];
    for e in p.elements.iter().filter(|e| e.slot == "controls") {
        let mut rest = e.xml.as_str();
        while let Some(i) = rest.find("<control ") {
            let el = &rest[i..rest[i..].find('>').map(|j| i + j).unwrap_or(rest.len())];
            if let (Some(id), Some(name)) = (attr(el, "shapeId").and_then(|s| s.parse().ok()), attr(el, "name")) {
                out.push((id, unescape(name)));
            }
            rest = &rest[i + 9..];
        }
    }
    out
}

fn unescape(s: &str) -> String {
    s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", "\"").replace("&apos;", "'").replace("&#10;", "\n").replace("&amp;", "&")
}

fn escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}

/// Kept content holding the given buttons (shape id, name, caption, macro, anchor), as Excel
/// writes them. For tests here and in other crates.
#[cfg(any(test, feature = "test-support"))]
pub fn sample_preserved(buttons: &[(u32, &str, &str, &str, [u32; 8])]) -> Preserved {
    use gridcraft_model::{PreservedPart, PreservedXml};
    let mut vml = String::from(
        r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">"#,
    );
    let mut controls = String::from("<controls>");
    for (id, name, caption, mac, a) in buttons {
        let anchor = a.iter().map(|n| n.to_string()).collect::<Vec<_>>().join(", ");
        vml.push_str(&format!(
            "<v:shape id=\"_x0000_s{id}\" type=\"#_x0000_t201\"><v:textbox><div style='text-align:center'><font face=\"Calibri\" size=\"220\" color=\"#000000\">{}</font></div></v:textbox><x:ClientData ObjectType=\"Button\"><x:Anchor>{anchor}</x:Anchor><x:FmlaMacro>[0]!{mac}</x:FmlaMacro></x:ClientData></v:shape>",
            escape(caption)
        ));
        controls.push_str(&format!("<control shapeId=\"{id}\" r:id=\"rId{id}\" name=\"{}\"/>", escape(name)));
    }
    vml.push_str("</xml>");
    controls.push_str("</controls>");
    Preserved {
        vml: Some(PreservedPart {
            name: "xl/drawings/vmlDrawing1.vml".into(),
            content_type: String::new(),
            data: vml.into_bytes().into(),
            rels: vec![],
        }),
        elements: vec![PreservedXml { slot: "controls".into(), xml: controls }],
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use gridcraft_model::{PreservedPart, PreservedXml};

    /// A sheet's kept content with two buttons, as Excel writes them.
    pub(crate) fn sample() -> Preserved {
        let vml = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <v:shape id="_x0000_s1025" type="#_x0000_t201" style='position:absolute'>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f"><div style='text-align:center'><font face="Calibri" size="220" color="#000000">Approve
 &amp; sign</font></div></v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>
    1, 5, 2, 3, 3, 8, 4, 10</x:Anchor>
   <x:PrintObject>False</x:PrintObject>
   <x:FmlaMacro>[0]!Approve</x:FmlaMacro>
  </x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" type="#_x0000_t201"><v:textbox><div>Second</div></v:textbox>
  <x:ClientData ObjectType="Button"><x:Anchor>0, 0, 10, 0, 1, 0, 11, 0</x:Anchor></x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1027" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Anchor>0,0,0,0,1,0,1,0</x:Anchor></x:ClientData></v:shape>
</xml>"##;
        Preserved {
            vml: Some(PreservedPart { name: "xl/drawings/vmlDrawing1.vml".into(), content_type: String::new(), data: vml.as_bytes().to_vec().into(), rels: vec![] }),
            elements: vec![PreservedXml { slot: "controls".into(), xml: r#"<controls><mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="1025" r:id="rId3" name="Button 1"/></mc:Choice></mc:AlternateContent><control shapeId="1026" r:id="rId4" name="Button 7"/></controls>"#.into() }],
            ..Default::default()
        }
    }

    #[test]
    fn reads_buttons() {
        let b = buttons(&sample());
        assert_eq!(b.len(), 2, "the note isn't a button");
        assert_eq!(b[0].name, "Button 1");
        assert_eq!(b[0].caption, "Approve\n& sign");
        assert_eq!(b[0].macro_name, "Approve");
        assert_eq!(b[0].anchor, [1, 5, 2, 3, 3, 8, 4, 10]);
        assert_eq!((b[1].name.as_str(), b[1].macro_name.as_str()), ("Button 7", ""));
        assert!(buttons(&Preserved::default()).is_empty());
    }

    #[test]
    fn rewrites_a_caption() {
        let mut p = sample();
        assert!(set_caption(&mut p, 1025, "Done <ok>\nline 2"));
        let b = buttons(&p);
        assert_eq!(b[0].caption, "Done <ok>\nline 2");
        assert_eq!(b[0].macro_name, "Approve", "the rest of the shape is untouched");
        assert_eq!(b[1].caption, "Second");
        assert!(!set_caption(&mut p, 9999, "x"));
    }
}
