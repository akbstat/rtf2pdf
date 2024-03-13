use anyhow::{anyhow, Result};
use encoding::{encode, need_encode};
use nanoid::nanoid;
use serde::{Deserialize, Serialize};
use std::{
    fs,
    os::windows::process::CommandExt,
    path::{Path, PathBuf},
    process::Command,
};
use vbs::vbs_render;

mod encoding;
mod vbs;

/// convert rtf to pdf
///
/// @args: list of source rtf file and output pdf file
///
/// # Example
/// ```
/// fn main() {
///     let source =
///         Path::new(r"D:\Studies\ak112\303\stats\CSR\product\output\t-14-01-01-01-disp-scr.rtf");
///     let dest =
///         Path::new(r"D:\Studies\ak112\303\stats\CSR\product\output\t-14-01-01-01-disp-scr.pdf");
///     let script_path = Path::new(r"D:\projects\rusty\mobius_kit\.utils\appdata\void_probe");
///     rtf2pdf(vec![(source, dest)], script_path).unwrap();
/// }
/// ```
pub fn rtf2pdf(args: Vec<(PathBuf, PathBuf)>, script_path: &Path) -> Result<()> {
    if args.is_empty() {
        return Ok(());
    }

    let need_encode = need_encode(&args.get(0).unwrap().0.to_string_lossy().to_string());

    let filename = format!(".task-{}.vbs", nanoid!(10));

    let task_script = match script_path.exists() {
        true => PathBuf::from(script_path.join(filename)),
        false => {
            let parent = args.get(0).unwrap().0.to_owned();
            let parent = if parent.is_file() {
                if let Some(p) = parent.parent() {
                    PathBuf::from(p)
                } else {
                    parent
                }
            } else {
                parent
            };
            let tasklist = parent.join(filename);
            tasklist
        }
    };
    let script = vbs_render(args)?;
    fs::write(task_script.as_path(), script)?;
    if need_encode {
        encode(task_script.as_path(), task_script.as_path())?;
    }
    let mut cmd = Command::new("cmd");
    cmd.creation_flags(0x08000000);
    cmd.arg("/C").arg(task_script.to_string_lossy().to_string());
    let result = cmd.output().unwrap();
    if !result.status.success() {
        return Err(anyhow!("{:?}", String::from_utf8(result.stderr)));
    }
    fs::remove_file(task_script)?;
    Ok(())
}

#[derive(Debug, Serialize, Deserialize)]
struct Assign {
    pub source: String,
    pub dest: String,
}

#[cfg(test)]
mod tests {
    use std::path::Path;

    use super::*;

    #[test]
    fn rtf2pdf_test() {
        let script_path = Path::new(r"D:\projects\rusty\mobius_kit\.utils\appdata\void_probe");
        let source = Path::new(
            r"D:\Studies\ak112\303\stats\CSR\product\output\测试\l-16-02-04-08-01-antu-ex-ss.rtf",
        );
        let dest = Path::new(
            r"D:\Studies\ak112\303\stats\CSR\product\output\测试\l-16-02-04-08-01-antu-ex-ss.pdf",
        );
        assert!(matches!(
            rtf2pdf(
                vec![(PathBuf::from(source), PathBuf::from(dest))],
                script_path
            ),
            Ok(())
        ));
    }
}
