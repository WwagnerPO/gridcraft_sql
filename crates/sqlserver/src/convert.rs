//! Platform-neutral pieces of the connector: reading the server name and turning SQL Server's
//! date and time encodings into worksheet serial numbers.

/// Where to connect: `host`, `host,port`, `host\instance` (`tcp:` prefix and `(local)` allowed).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ServerAddress {
    pub host: String,
    pub port: Option<u16>,
    pub instance: Option<String>,
}

pub fn parse_server(text: &str) -> Result<ServerAddress, String> {
    let t = text.trim();
    let t = t.strip_prefix("tcp:").unwrap_or(t).trim();
    let (rest, port) = match t.rsplit_once(',') {
        Some((h, p)) => {
            let p = p.trim().parse::<u16>().map_err(|_| format!("`{}` is not a port number.", p.trim()))?;
            (h.trim(), Some(p))
        }
        None => (t, None),
    };
    let (host, instance) = match rest.split_once('\\') {
        Some((h, i)) if !i.trim().is_empty() => (h.trim(), Some(i.trim().to_string())),
        Some((h, _)) => (h.trim(), None),
        None => (rest, None),
    };
    let host = match host {
        "" => return Err("Enter a server name.".into()),
        "." | "(local)" | "(localdb)" => "localhost",
        h => h,
    };
    if host.chars().any(|c| c.is_whitespace() || c.is_control()) {
        return Err(format!("`{host}` is not a server name."));
    }
    Ok(ServerAddress { host: host.to_string(), port, instance })
}

/// Days from 0001-01-01 (SQL Server's `date` epoch) to 1900-01-01.
const DAYS_TO_1900: i64 = 693_595;

/// The serial number of a date given as days since 1900-01-01 plus a fraction of a day, in the
/// 1900 date system (which counts the fictitious 29 February 1900). `None` before 1900.
pub fn serial(days_since_1900: i64, fraction: f64) -> Option<f64> {
    if days_since_1900 < 0 {
        return None;
    }
    let mut d = days_since_1900 + 1;
    if d >= 60 {
        d += 1;
    }
    if d > 2_958_465 {
        return None;
    }
    Some(d as f64 + fraction.clamp(0.0, 1.0))
}

/// `datetime`: days since 1900-01-01 and 1/300 s ticks.
pub fn datetime(days: i32, ticks: u32) -> Option<f64> {
    serial(days as i64, ticks as f64 / 300.0 / 86_400.0)
}

/// `smalldatetime`: days since 1900-01-01 and minutes.
pub fn small_datetime(days: u16, minutes: u16) -> Option<f64> {
    serial(days as i64, minutes as f64 / 1_440.0)
}

/// `date`: days since 0001-01-01.
pub fn date(days: u32) -> Option<f64> {
    serial(days as i64 - DAYS_TO_1900, 0.0)
}

/// `time(scale)`: `increments` of 10^-scale seconds since midnight, as a fraction of a day.
pub fn time(increments: u64, scale: u8) -> f64 {
    let secs = increments as f64 / 10f64.powi(scale.min(7) as i32);
    (secs / 86_400.0).clamp(0.0, 1.0)
}

/// `datetime2` (`date` + `time`), shifted by `offset_minutes` (`datetimeoffset` stores UTC).
pub fn datetime2(days: u32, increments: u64, scale: u8, offset_minutes: i16) -> Option<f64> {
    let local = (days as i64 - DAYS_TO_1900) as f64 + time(increments, scale) + offset_minutes as f64 / 1_440.0;
    let whole = local.floor();
    serial(whole as i64, local - whole)
}

/// `yyyy-mm-dd` of a date before 1900, which has no serial number: days since 0001-01-01.
pub fn iso_date(days: i64) -> String {
    // Days since 1970-01-01, then the civil-from-days algorithm (proleptic Gregorian).
    let z = days - 719_162 + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = yoe + era * 400 + i64::from(m <= 2);
    format!("{y:04}-{m:02}-{d:02}")
}

/// `[name]` with `]` doubled: a safe identifier in a statement.
pub fn quote_identifier(name: &str) -> String {
    format!("[{}]", name.replace(']', "]]"))
}

/// Hex text of binary data (`0x…`), cut so the cell's text limit holds.
pub fn hex(bytes: &[u8], max_chars: usize) -> String {
    let n = bytes.len().min(max_chars.saturating_sub(2) / 2);
    let mut s = String::with_capacity(2 + n * 2);
    s.push_str("0x");
    for b in bytes.iter().take(n) {
        s.push_str(&format!("{b:02X}"));
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn server_names() {
        let a = |host: &str, port: Option<u16>, instance: Option<&str>| ServerAddress { host: host.into(), port, instance: instance.map(Into::into) };
        assert_eq!(parse_server("db.example.com"), Ok(a("db.example.com", None, None)));
        assert_eq!(parse_server(" tcp:db,1444 "), Ok(a("db", Some(1444), None)));
        assert_eq!(parse_server("HOST\\SQLEXPRESS"), Ok(a("HOST", None, Some("SQLEXPRESS"))));
        assert_eq!(parse_server("(local)"), Ok(a("localhost", None, None)));
        assert_eq!(parse_server("."), Ok(a("localhost", None, None)));
        assert!(parse_server("").is_err());
        assert!(parse_server("db,port").is_err());
        assert!(parse_server("db,99999").is_err());
        assert!(parse_server("bad host").is_err());
    }

    #[test]
    fn dates() {
        // 1900-01-01 is serial 1; 1900-03-01 is 61 (after the fictitious 29 February).
        assert_eq!(datetime(0, 0), Some(1.0));
        assert_eq!(datetime(58, 0), Some(59.0));
        assert_eq!(datetime(59, 0), Some(61.0));
        // 2023-03-15 12:00 → 45000.5.
        assert_eq!(datetime(44_998, 300 * 43_200), Some(45_000.5));
        assert_eq!(small_datetime(44_998, 720), Some(45_000.5));
        assert_eq!(date(DAYS_TO_1900 as u32 + 44_998), Some(45_000.0));
        assert_eq!(date(0), None);
        assert_eq!(time(432_000_000_000, 7), 0.5);
        assert_eq!(time(6, 0), 6.0 / 86_400.0);
        // 2023-03-15 23:00 UTC at +02:00 is 2023-03-16 01:00.
        let v = datetime2(DAYS_TO_1900 as u32 + 44_998, 23 * 3_600, 0, 120).unwrap_or(0.0);
        assert!((v - (45_001.0 + 1.0 / 24.0)).abs() < 1e-9, "{v}");
        assert_eq!(iso_date(0), "0001-01-01");
        assert_eq!(iso_date(DAYS_TO_1900 - 1), "1899-12-31");
        assert_eq!(iso_date(DAYS_TO_1900 + 59), "1900-03-01");
    }

    #[test]
    fn identifiers_and_hex() {
        assert_eq!(quote_identifier("Order Details"), "[Order Details]");
        assert_eq!(quote_identifier("a]; DROP TABLE x;--"), "[a]]; DROP TABLE x;--]");
        assert_eq!(hex(&[0xde, 0xad], 100), "0xDEAD");
        assert_eq!(hex(&[1, 2, 3, 4], 6), "0x0102");
    }
}
