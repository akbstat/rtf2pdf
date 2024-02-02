use std::path::PathBuf;

use tera::{Context, Tera};

const TEMPLATE: &str = "VBS";
const VBS_TEMPLATE: &str = r#"
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
{% for task in tasks %}
Set objDoc = objWord.Documents.Open("{{task.0}}")
objDoc.ExportAsFixedFormat "{{task.1}}", 17
objDoc.Close False
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
