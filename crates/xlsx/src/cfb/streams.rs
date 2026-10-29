//! Reading any stream of a compound file by path (`["VBA", "dir"]`), not only root streams: the
//! VBA project (MS-OVBA) keeps its modules in a storage.

use super::{Compound, DirEntry, MAX_STREAM, NO_STREAM, STREAM_OBJECT};

impl Compound<'_> {
    /// The entry called `name` (case-insensitive, stream or storage) among the children of
    /// `parent` (a directory index).
    fn child_entry(&self, parent: usize, name: &str) -> Result<Option<usize>, String> {
        let Some(p) = self.entries.get(parent) else { return Ok(None) };
        let mut stack = vec![p.child];
        let mut seen = 0usize;
        while let Some(i) = stack.pop() {
            if i == NO_STREAM {
                continue;
            }
            seen += 1;
            if seen > self.entries.len() {
                return Err("damaged compound file: directory tree loops".into());
            }
            let Some(e) = self.entries.get(i as usize) else { continue };
            if e.name.eq_ignore_ascii_case(name) {
                return Ok(Some(i as usize));
            }
            stack.push(e.left);
            stack.push(e.right);
        }
        Ok(None)
    }

    /// The names of the streams directly in the storage at `path` (`["VBA"]`).
    #[cfg(test)]
    pub(crate) fn stream_names(&self, path: &[&str]) -> Result<Vec<String>, String> {
        let mut at = 0usize;
        for part in path {
            match self.child_entry(at, part)? {
                Some(i) if self.entries.get(i).is_some_and(|e| e.kind == 1) => at = i,
                _ => return Ok(vec![]),
            }
        }
        let Some(p) = self.entries.get(at) else { return Ok(vec![]) };
        let mut out = vec![];
        let mut stack = vec![p.child];
        let mut seen = 0usize;
        while let Some(i) = stack.pop() {
            if i == NO_STREAM {
                continue;
            }
            seen += 1;
            if seen > self.entries.len() {
                return Err("damaged compound file: directory tree loops".into());
            }
            let Some(e) = self.entries.get(i as usize) else { continue };
            if e.kind == STREAM_OBJECT {
                out.push(e.name.clone());
            }
            stack.push(e.left);
            stack.push(e.right);
        }
        Ok(out)
    }

    /// The stream at `path` (`["VBA", "dir"]`), storages walked by name (case-insensitive).
    pub(crate) fn stream(&self, path: &[&str]) -> Result<Option<Vec<u8>>, String> {
        let Some((last, dirs)) = path.split_last() else { return Ok(None) };
        let mut at = 0usize;
        for part in dirs {
            match self.child_entry(at, part)? {
                Some(i) if self.entries.get(i).is_some_and(|e| e.kind == 1) => at = i,
                _ => return Ok(None),
            }
        }
        let Some(i) = self.child_entry(at, last)? else { return Ok(None) };
        let Some(e) = self.entries.get(i).filter(|e| e.kind == STREAM_OBJECT) else { return Ok(None) };
        self.read_entry(e).map(Some)
    }

    fn read_entry(&self, e: &DirEntry) -> Result<Vec<u8>, String> {
        let Some(root) = self.entries.first() else { return Err("damaged compound file: no root".into()) };
        if e.size > MAX_STREAM {
            return Err("a stream is larger than 1 GiB".into());
        }
        let size = usize::try_from(e.size).map_err(|_| "stream too large".to_string())?;
        let mut data = if e.size < self.mini_cutoff {
            let container = self.read_chain(root.start, root.size)?;
            self.read_mini_chain(&container, e.start, size)?
        } else {
            self.read_chain(e.start, e.size)?
        };
        if data.len() < size {
            return Err("damaged compound file: a stream is shorter than its declared size (the file may be truncated)".into());
        }
        data.truncate(size);
        Ok(data)
    }
}
