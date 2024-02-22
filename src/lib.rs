use anyhow::{anyhow, Result};
use nanoid::nanoid;
use serde::{Deserialize, Serialize};
use std::{fs, os::windows::process::CommandExt, path::PathBuf, process::Command};
use vbs::vbs_render;

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
///     rtf2pdf(vec![(source, dest)]).unwrap();
/// }
/// ```
pub fn rtf2pdf(args: Vec<(PathBuf, PathBuf)>) -> Result<()> {
    if args.is_empty() {
        return Ok(());
    }

    let task_script = {
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
        let tasklist = parent.join(format!(".task-{}.vbs", nanoid!(10)));
        tasklist
    };

    fs::write(task_script.as_path(), vbs_render(args)?)?;
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
        let source = Path::new(
            r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0001.rtf",
        );
        let dest = Path::new(
            r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0001.pdf",
        );
        assert!(matches!(
            rtf2pdf(vec![(PathBuf::from(source), PathBuf::from(dest))]),
            Ok(())
        ));
    }
}
