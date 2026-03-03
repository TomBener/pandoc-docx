# Custom DOCX Template for Pandoc

Inspired by [pandoc-docx-tools](https://github.com/rnwst/pandoc-docx-tools). Based on the [default DOCX template](https://github.com/jgm/pandoc/tree/main/data/docx) of [Pandoc 3.9](https://github.com/jgm/pandoc/releases/tag/3.9).

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
./pandoc-docx.sh zip && pandoc -o test.docx --reference-doc=reference.docx test.md --number-sections
```

## Custom OpenXML Template

`custom.openxml` is a Pandoc `--template` for controlling the document body structure. Based on the [official default](https://github.com/jgm/pandoc/blob/main/data/templates/default.openxml), with `affiliation` and `email` metadata fields added after `author`.
