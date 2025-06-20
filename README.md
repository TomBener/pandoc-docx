# Custom DOCX Template for Pandoc

Inspired by [pandoc-docx-tools](https://github.com/rnwst/pandoc-docx-tools). Based on the [default DOCX template](https://github.com/jgm/pandoc/tree/main/data/docx) of [Pandoc 3.7.0.2](https://github.com/jgm/pandoc/releases/tag/3.7.0.2).

## Usage

```bash
# Make script executable
chmod +x pandoc-docx.sh

# Generate reference.docx from `reference` folder
./pandoc-docx.sh zip

# Edit reference.docx in Microsoft Word

# Unzip reference.docx to `reference` folder
./pandoc-docx.sh unzip
```

Use the `reference.docx` file as the template:

```bash
./pandoc-docx.sh zip && pandoc -o output.docx --reference-doc=reference.docx input.md --number-sections
```
