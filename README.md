# RTF2PDF
A wrapper to call word application to save rtf files as pdf

# Usage
```rust
fn main() {
    let bin = Path::new(r"D:\projects\rusty\mobius_kit\.bin\rtf2pdf.exe");
    let source =
        Path::new(r"D:\Studies\ak112\303\stats\CSR\product\output\t-14-01-01-01-disp-scr.rtf");
    let dest =
        Path::new(r"D:\Studies\ak112\303\stats\CSR\product\output\t-14-01-01-01-disp-scr.pdf");
    assert!(matches!(rtf2pdf(bin, vec![(source, dest)]), Ok(())));
}
```