use encoding_rs::GB18030;
use std::{
    fs::{self, OpenOptions},
    io::{BufWriter, Write},
    path::Path,
};

/// To Encode script file from UTF-8 to GB18030 if path information in script include Chinese characters
/// determine if sample string contains Chinese charaters
///
/// ```rust
/// assert!(need_encode("测试"));
/// assert!(need_encodee("test")); // will panic
/// ````
pub fn need_encode(sample: &str) -> bool {
    if sample.chars().any(|c| c > '\u{7F}') {
        true
    } else {
        false
    }
}
/// encode source file into GB18030
///
/// ```rust
/// let source = Path::new(r"D:\Studies\source.txt");
/// let dest = Path::new(r"D:\Studies\dest.txt");
/// encode(source, dest).unwrap();
/// ```
pub fn encode(source: &Path, dest: &Path) -> anyhow::Result<()> {
    let source = fs::read_to_string(source).unwrap();
    let mut output_file = BufWriter::new(
        OpenOptions::new()
            .write(true)
            .create(true)
            .truncate(true)
            .open(dest)
            .unwrap(),
    );
    let (result, _, _) = GB18030.new_encoder().encoding().encode(&source);
    output_file.write_all(&result)?;
    Ok(())
}

#[cfg(test)]
mod encoding_test {
    use super::*;
    #[test]
    fn need_encode_test() {
        let sample =
            r"E:\Users\yuqi01.chen\Desktop\misc\output\测试\f-14-02-01-01-irrc-pfs-km-fas.rtf";
        assert!(need_encode(sample));
        let sample =
            r"E:\Users\yuqi01.chen\Desktop\misc\output\test\f-14-02-01-01-irrc-pfs-km-fas.rtf";
        assert!(!need_encode(sample));
    }

    #[test]
    fn encode_test() {
        let source = Path::new(r"D:\projects\rusty\mobius_kit\.utils\appdata\void_probe\test.vbs");
        let dest = Path::new(r"D:\projects\rusty\mobius_kit\.utils\appdata\void_probe\test.vbs");
        encode(source, dest).expect("failed to encode");
    }
}
