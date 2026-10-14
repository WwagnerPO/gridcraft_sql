//! Writing compound files ([MS-CFB], version 3: 512-byte sectors), for VBA projects
//! (`vbaProject.bin`). Streams under 4,096 bytes go in the mini stream, as the format requires.
//! Implemented from Microsoft's public [MS-CFB] specification.

const SECTOR: usize = 512;
const MINI: usize = 64;
const CUTOFF: usize = 4096;
const FREE: u32 = 0xFFFF_FFFF;
const END: u32 = 0xFFFF_FFFE;
const FAT_SECT: u32 = 0xFFFF_FFFD;
const NO_STREAM: u32 = 0xFFFF_FFFF;

/// One entry to write: a path of storage names ending in the stream's name (`["VBA", "dir"]`),
/// and its bytes. Storages are created as the paths need them.
pub struct Stream<'a> {
    pub path: Vec<&'a str>,
    pub data: &'a [u8],
}

struct Entry {
    name: String,
    storage: bool,
    parent: usize,
    data: Vec<u8>,
    start: u32,
    children: Vec<usize>,
}

/// CFB's sibling order: shorter names first, then by upper-cased UTF-16 code units.
fn order(a: &str, b: &str) -> std::cmp::Ordering {
    let ua: Vec<u16> = a.to_uppercase().encode_utf16().collect();
    let ub: Vec<u16> = b.to_uppercase().encode_utf16().collect();
    ua.len().cmp(&ub.len()).then(ua.cmp(&ub))
}

/// A balanced tree over `sorted` (left, right sibling links); returns the subtree's root.
fn tree(sorted: &[usize], links: &mut [(u32, u32)]) -> u32 {
    if sorted.is_empty() {
        return NO_STREAM;
    }
    let mid = sorted.len() / 2;
    let (left, rest) = sorted.split_at(mid);
    let Some((&me, right)) = rest.split_first() else { return NO_STREAM };
    let l = tree(left, links);
    let r = tree(right, links);
    if let Some(x) = links.get_mut(me) {
        *x = (l, r);
    }
    me as u32
}

/// The compound file holding `streams`. Errors when names are invalid or the file would need
/// more than the header's 109 FAT sectors (about 7 MB).
pub fn write(streams: &[Stream<'_>]) -> Result<Vec<u8>, String> {
    let mut entries = vec![Entry { name: "Root Entry".into(), storage: true, parent: 0, data: vec![], start: END, children: vec![] }];
    for s in streams {
        let Some((last, dirs)) = s.path.split_last() else { return Err("empty stream path".into()) };
        let mut at = 0usize;
        for d in dirs {
            let found = entries
                .get(at)
                .and_then(|e| e.children.iter().copied().find(|&c| entries.get(c).is_some_and(|x| x.storage && x.name.eq_ignore_ascii_case(d))));
            at = match found {
                Some(c) => c,
                None => {
                    let i = entries.len();
                    entries.push(Entry { name: (*d).to_string(), storage: true, parent: at, data: vec![], start: END, children: vec![] });
                    if let Some(p) = entries.get_mut(at) {
                        p.children.push(i);
                    }
                    i
                }
            };
        }
        let i = entries.len();
        entries.push(Entry { name: (*last).to_string(), storage: false, parent: at, data: s.data.to_vec(), start: END, children: vec![] });
        if let Some(p) = entries.get_mut(at) {
            p.children.push(i);
        }
    }
    if entries.iter().any(|e| e.name.encode_utf16().count() > 31 || e.name.is_empty() || e.name.contains(['/', '\\', ':', '!'])) {
        return Err("invalid compound file entry name".into());
    }
    // Mini stream: small streams in 64-byte sectors.
    let mut mini = Vec::new();
    let mut mini_fat: Vec<u32> = vec![];
    for e in entries.iter_mut().filter(|e| !e.storage && !e.data.is_empty() && e.data.len() < CUTOFF) {
        let first = mini_fat.len() as u32;
        let n = e.data.len().div_ceil(MINI);
        for k in 0..n {
            mini_fat.push(if k + 1 == n { END } else { first + k as u32 + 1 });
        }
        mini.extend_from_slice(&e.data);
        mini.resize(mini_fat.len() * MINI, 0);
        e.start = first;
    }
    let big: Vec<usize> = (0..entries.len()).filter(|&i| entries.get(i).is_some_and(|e| !e.storage && e.data.len() >= CUTOFF)).collect();
    let dir_sectors = (entries.len() * 128).div_ceil(SECTOR);
    let minifat_sectors = (mini_fat.len() * 4).div_ceil(SECTOR);
    let ministream_sectors = mini.len().div_ceil(SECTOR);
    let big_sectors: usize = big.iter().filter_map(|&i| entries.get(i)).map(|e| e.data.len().div_ceil(SECTOR)).sum();
    let content = dir_sectors + minifat_sectors + ministream_sectors + big_sectors;
    let mut fat_sectors = 1;
    while (fat_sectors + content) > fat_sectors * (SECTOR / 4) {
        fat_sectors += 1;
    }
    if fat_sectors > 109 {
        return Err("the VBA project is too large to write".into());
    }
    let total = fat_sectors + content;
    let mut fat = vec![FREE; fat_sectors * (SECTOR / 4)];
    let mut next = 0u32;
    fn chain(fat: &mut [u32], next: &mut u32, n: usize) -> u32 {
        let first = *next;
        for k in 0..n {
            let id = *next as usize;
            if let Some(x) = fat.get_mut(id) {
                *x = if k + 1 == n { END } else { *next + 1 };
            }
            *next += 1;
        }
        if n == 0 { END } else { first }
    }
    for _ in 0..fat_sectors {
        if let Some(x) = fat.get_mut(next as usize) {
            *x = FAT_SECT;
        }
        next += 1;
    }
    let dir_start = chain(&mut fat, &mut next, dir_sectors);
    let minifat_start = chain(&mut fat, &mut next, minifat_sectors);
    let ministream_start = chain(&mut fat, &mut next, ministream_sectors);
    for &i in &big {
        let n = entries.get(i).map_or(0, |e| e.data.len().div_ceil(SECTOR));
        let start = chain(&mut fat, &mut next, n);
        if let Some(e) = entries.get_mut(i) {
            e.start = start;
        }
    }
    if let Some(root) = entries.first_mut() {
        root.start = if mini.is_empty() { END } else { ministream_start };
    }
    // Directory tree: each storage's children sorted, as a balanced tree.
    let mut links = vec![(NO_STREAM, NO_STREAM); entries.len()];
    let mut child = vec![NO_STREAM; entries.len()];
    for i in 0..entries.len() {
        let mut kids = entries.get(i).map(|e| e.children.clone()).unwrap_or_default();
        kids.sort_by(|&a, &b| order(entries.get(a).map_or("", |e| e.name.as_str()), entries.get(b).map_or("", |e| e.name.as_str())));
        if let Some(c) = child.get_mut(i) {
            *c = tree(&kids, &mut links);
        }
    }
    let mut out = vec![0u8; SECTOR * (total + 1)];
    let put = |out: &mut Vec<u8>, at: usize, b: &[u8]| {
        if let Some(s) = out.get_mut(at..at + b.len()) {
            s.copy_from_slice(b);
        }
    };
    // Header.
    put(&mut out, 0, &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    put(&mut out, 24, &0x003Eu16.to_le_bytes());
    put(&mut out, 26, &3u16.to_le_bytes());
    put(&mut out, 28, &0xFFFEu16.to_le_bytes());
    put(&mut out, 30, &9u16.to_le_bytes());
    put(&mut out, 32, &6u16.to_le_bytes());
    put(&mut out, 44, &(fat_sectors as u32).to_le_bytes());
    put(&mut out, 48, &dir_start.to_le_bytes());
    put(&mut out, 56, &(CUTOFF as u32).to_le_bytes());
    put(&mut out, 60, &(if minifat_sectors == 0 { END } else { minifat_start }).to_le_bytes());
    put(&mut out, 64, &(minifat_sectors as u32).to_le_bytes());
    put(&mut out, 68, &END.to_le_bytes());
    for k in 0..109usize {
        let v = if k < fat_sectors { k as u32 } else { FREE };
        put(&mut out, 76 + k * 4, &v.to_le_bytes());
    }
    let sector_at = |id: u32| SECTOR * (id as usize + 1);
    for (k, v) in fat.iter().enumerate() {
        put(&mut out, SECTOR + k * 4, &v.to_le_bytes());
    }
    // Directory entries.
    for (i, e) in entries.iter().enumerate() {
        let at = sector_at(dir_start) + i * 128;
        let name: Vec<u16> = e.name.encode_utf16().collect();
        for (k, u) in name.iter().enumerate() {
            put(&mut out, at + k * 2, &u.to_le_bytes());
        }
        put(&mut out, at + 64, &(((name.len() + 1) * 2) as u16).to_le_bytes());
        out[at + 66] = if i == 0 {
            5
        } else if e.storage {
            1
        } else {
            2
        };
        out[at + 67] = 1; // black
        let (l, r) = links.get(i).copied().unwrap_or((NO_STREAM, NO_STREAM));
        put(&mut out, at + 68, &l.to_le_bytes());
        put(&mut out, at + 72, &r.to_le_bytes());
        put(&mut out, at + 76, &child.get(i).copied().unwrap_or(NO_STREAM).to_le_bytes());
        let (start, size) = if i == 0 {
            (e.start, mini.len())
        } else if e.storage {
            (0, 0)
        } else {
            (if e.data.is_empty() { END } else { e.start }, e.data.len())
        };
        put(&mut out, at + 116, &start.to_le_bytes());
        put(&mut out, at + 120, &(size as u32).to_le_bytes());
        let _ = e.parent;
    }
    // Unused directory slots are empty entries with no siblings.
    for i in entries.len()..dir_sectors * (SECTOR / 128) {
        let at = sector_at(dir_start) + i * 128;
        put(&mut out, at + 68, &NO_STREAM.to_le_bytes());
        put(&mut out, at + 72, &NO_STREAM.to_le_bytes());
        put(&mut out, at + 76, &NO_STREAM.to_le_bytes());
    }
    if minifat_sectors > 0 {
        let mut b: Vec<u8> = mini_fat.iter().flat_map(|v| v.to_le_bytes()).collect();
        b.resize(minifat_sectors * SECTOR, 0xFF);
        put(&mut out, sector_at(minifat_start), &b);
    }
    if ministream_sectors > 0 {
        put(&mut out, sector_at(ministream_start), &mini);
    }
    for &i in &big {
        if let Some(e) = entries.get(i) {
            put(&mut out, sector_at(e.start), &e.data);
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips_through_the_reader() {
        let big: Vec<u8> = (0..10_000u32).map(|i| (i % 251) as u8).collect();
        let small = b"hello".to_vec();
        let file = write(&[
            Stream { path: vec!["PROJECT"], data: &small },
            Stream { path: vec!["VBA", "dir"], data: &big },
            Stream { path: vec!["VBA", "Module1"], data: b"x" },
            Stream { path: vec!["VBA", "_VBA_PROJECT"], data: &[] },
        ])
        .unwrap();
        let cf = crate::cfb::Compound::open(&file).unwrap();
        assert_eq!(cf.stream(&["PROJECT"]).unwrap().unwrap(), small);
        assert_eq!(cf.stream(&["vba", "DIR"]).unwrap().unwrap(), big);
        assert_eq!(cf.stream(&["VBA", "Module1"]).unwrap().unwrap(), b"x");
        assert_eq!(cf.stream(&["VBA", "_VBA_PROJECT"]).unwrap().unwrap(), b"");
        let mut names = cf.stream_names(&["VBA"]).unwrap();
        names.sort();
        assert_eq!(names, ["Module1", "_VBA_PROJECT", "dir"]);
        assert!(write(&[Stream { path: vec!["a/b"], data: b"" }]).is_err());
    }
}
