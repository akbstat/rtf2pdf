use std::path::PathBuf;

use tera::{Context, Tera};

const TEMPLATE: &str = "VBS";
const VBS_TEMPLATE: &str = r#"
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.DisplayAlerts = wdAlertsNone
{% for task in tasks %}
Set objDoc{{loop.index}} = objWord.Documents.Open("{{task.0}}")

objDoc{{loop.index}}.ExportAsFixedFormat "{{task.1}}", 17
objDoc{{loop.index}}.Close False
WScript.Sleep 100
{% endfor %}
objWord.Quit
"#;

pub fn vbs_render(tasks: Vec<(PathBuf, PathBuf)>) -> anyhow::Result<String> {
    let mut template = Tera::default();
    template.add_raw_template(TEMPLATE, VBS_TEMPLATE)?;
    let mut ctx = Context::new();
    ctx.insert("tasks", &tasks);
    Ok(template.render(TEMPLATE, &ctx)?)
}

#[cfg(test)]
mod test_vbs {
    use std::str::FromStr;

    use super::*;
    #[test]
    fn vbs_test() {
        let tasks = vec![
            (
                PathBuf::from_str(
                    r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0001.rtf",
                ).unwrap(),
                PathBuf::from_str(
                    r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0001.pdf",
                ).unwrap(),
            ),
            (
                PathBuf::from_str(
                    r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0002.rtf",
                ).unwrap(),
                PathBuf::from_str(
                    r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0002.pdf",
                ).unwrap(),
            ),
            (
                PathBuf::from_str(
                    r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0003.rtf",
                ).unwrap(),
                PathBuf::from_str(
                    r"D:\Studies\ak112\303\stats\CSR\product\output\rtf_divided\l-16-02-08-05-ecg-ss_part_0003.pdf",
                ).unwrap(),
            ),
        ];
        vbs_render(tasks).unwrap();
    }
}

// Set objDoc3 = objWord.Documents.Open("E:\Users\yuqi01.chen\Desktop\misc\output\ASCO\.temp\f-14-02-01-05-irrc-pfs-for-fas_part_0001.rtf")
// objDoc3.ExportAsFixedFormat "E:\Users\yuqi01.chen\Desktop\misc\output\ASCO\.temp\f-14-02-01-05-irrc-pfs-for-fas_part_0001.pdf", 17
// objDoc3.Close False
// WScript.Sleep 100

// Set objDoc4 = objWord.Documents.Open("E:\Users\yuqi01.chen\Desktop\misc\output\ASCO\.temp\t-14-01-03-01-dm-fas_part_0001.rtf")
// objDoc4.ExportAsFixedFormat "E:\Users\yuqi01.chen\Desktop\misc\output\ASCO\.temp\t-14-01-03-01-dm-fas_part_0001.pdf", 17
// objDoc4.Close False
// WScript.Sleep 100
